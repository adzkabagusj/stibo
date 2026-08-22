"""
AWS Lambda entry point for the Ecomm Asset Resizing Pipeline.

S3 Trigger
───────────
Fires automatically when a product XML is uploaded under:
  s3://map-stibo-outbound-dev/step-export/ecomm-assets/

Supported trigger shapes
─────────────────────────
1. S3 Event (automatic)
   {
     "Records": [{
       "s3": {
         "bucket": { "name": "map-stibo-outbound-dev" },
         "object": { "key": "step-export/ecomm-assets/1304550.xml" }
       }
     }]
   }

2. Direct / manual invocation
   {
     "key": "step-export/ecomm-assets/1304550.xml"
   }

3. API Gateway POST
   {
     "body": "{\"key\": \"step-export/ecomm-assets/1304550.xml\"}"
   }

Response shape
───────────────
Success (200):
  {
    "product_id":   "1304550",
    "input_key":    "step-export/ecomm-assets/1304550.xml",
    "zips_created": 2,
    "s3_keys": [
      "transformed/ecomm-assets/800x800-2026-06-09_10.30.00.zip",
      "transformed/ecomm-assets/1200x1200-2026-06-09_10.30.00.zip"
    ]
  }

Error (4xx / 5xx):
  { "error": "...message..." }

Audit logs
───────────
Every invocation writes a structured JSON audit log to:
  s3://{S3_BUCKET}/etl-audit-logs/Ecomm-Assets/{product_id}/{timestamp}.json
"""

import json
import os
import shutil
import tempfile
import time
import urllib.parse
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import boto3

from resize_pipeline import config
from resize_pipeline.pipeline import run_pipeline
from resize_pipeline.s3_helper import download_bytes, upload_zip


# ── Audit log prefix ──────────────────────────────────────────────────────────
S3_AUDIT_PREFIX: str = "etl-audit-logs/ecomm-assets"


# ── Helpers ───────────────────────────────────────────────────────────────────

def _parse_event(event: dict) -> str:
    """
    Extract the S3 object key from any supported event shape.

    The bucket is always config.S3_BUCKET - it is not read from the event.

    Parameters
    ----------
    event : dict
        Raw Lambda event payload.

    Returns
    -------
    str
        S3 object key e.g. "step-export/ecomm-assets/1304550.xml"

    Raises
    ------
    ValueError
        If the event does not match any recognised shape.
    """
    # ── Shape 4: EventBridge "Object Created" notification ────────────────────
    # event["source"] == "aws.s3"
    # event["detail-type"] == "Object Created"
    # event["detail"]["bucket"]["name"] -> bucket
    # event["detail"]["object"]["key"]  -> key (URL-encoded)
    if event.get("source") == "aws.s3" and event.get("detail-type") == "Object Created":
        try:
            key = urllib.parse.unquote_plus(
                event["detail"]["object"]["key"]
            )
            print(
                f"  [PARSE] EventBridge shape detected. "
                f"bucket={event['detail']['bucket']['name']!r}  key={key!r}"
            )
            return key
        except (KeyError, TypeError) as exc:
            raise ValueError(
                f"EventBridge event missing expected detail fields: {exc}. "
                "Expected: event['detail']['bucket']['name'] and "
                "event['detail']['object']['key']."
            ) from exc

    # ── Shape 1: S3 event trigger ─────────────────────────────────────────────
    if "Records" in event:
        key = urllib.parse.unquote_plus(
            event["Records"][0]["s3"]["object"]["key"]
        )
        print(f"  [PARSE] S3 Records shape detected. key={key!r}")
        return key

    # ── Shape 2: Direct / manual invocation ──────────────────────────────────
    if "key" in event:
        print(f"  [PARSE] Direct invocation shape detected. key={event['key']!r}")
        return event["key"]

    # ── Shape 3: API Gateway POST ─────────────────────────────────────────────
    if "body" in event:
        try:
            body = json.loads(event["body"])
        except (json.JSONDecodeError, TypeError) as exc:
            raise ValueError(
                f"Could not parse event body as JSON: {exc}"
            ) from exc
        if "key" in body:
            print(f"  [PARSE] API Gateway shape detected. key={body['key']!r}")
            return body["key"]

    raise ValueError(
        "Unrecognised event shape. Expected one of:\n"
        "  1. EventBridge : { source: 'aws.s3', detail-type: 'Object Created', detail: { bucket: { name }, object: { key } } }\n"
        "  2. S3 event    : { Records: [{ s3: { bucket, object } }] }\n"
        "  3. Direct      : { key: 'step-export/ecomm-assets/filename.xml' }\n"
        "  4. API Gateway : { body: '{\"key\": \"step-export/ecomm-assets/filename.xml\"}' }"
    )

def _success(body: dict) -> dict:
    """Return a structured 200 response."""
    return {
        "statusCode": 200,
        "headers":    {"Content-Type": "application/json"},
        "body":       json.dumps(body),
    }


def _error(status_code: int, message: str) -> dict:
    """Return a structured error response."""
    print(f"[ERROR] {message}")
    return {
        "statusCode": status_code,
        "headers":    {"Content-Type": "application/json"},
        "body":       json.dumps({"error": message}),
    }


def _write_audit_log(audit: dict) -> None:
    """
    Serialise audit dict to JSON and write it to S3 under etl-audit-logs/.

    S3 key pattern:
        etl-audit-logs/Ecomm-Assets/{product_id}/{YYYY-MM-DDTHH-MM-SS}Z.json

    Failures are logged but never re-raised so audit logging never
    interrupts the main pipeline response.

    Parameters
    ----------
    audit : dict
        Structured audit record (see lambda_handler for schema).
    """
    if not config.S3_BUCKET:
        print("[AUDIT] S3_BUCKET not set - audit log skipped.")
        return

    product_id = audit.get("product_id") or "ecomm-assets"
    # Use a filesystem-safe timestamp: colons replaced with hyphens
    ts = audit.get("started_at", datetime.now(timezone.utc).isoformat())
    ts_safe = ts.replace(":", "-").replace(".", "-")

    s3_key = f"{S3_AUDIT_PREFIX}/{product_id}/{ts_safe}Z.json"

    try:
        s3_client = boto3.client("s3")
        s3_client.put_object(
            Bucket=config.S3_BUCKET,
            Key=s3_key,
            Body=json.dumps(audit, indent=2).encode("utf-8"),
            ContentType="application/json",
        )
        print(f"[AUDIT] Log written to s3://{config.S3_BUCKET}/{s3_key}")
    except Exception as exc:
        # Never let audit failure break the pipeline response
        print(f"[AUDIT] WARNING - failed to write audit log: {exc}")


# ── Lambda handler ────────────────────────────────────────────────────────────
def lambda_handler(event: dict, context: Any) -> dict:
    """
    Main Lambda entry point.

    Steps
    ─────
    1. Validate required environment variables.
    2. Parse event to extract the S3 object key.
    3. Download the product XML from S3 to /tmp.
    4. Run the full resize pipeline.
    5. Upload each output ZIP to transformed/ecomm-assets/ in S3.
    6. Write a structured audit log to etl-audit-logs/ in S3.
    7. Return a structured response.

    Parameters
    ----------
    event : dict
        Lambda event payload (see module docstring for supported shapes).
    context : LambdaContext
        Lambda runtime context (not used directly).

    Returns
    -------
    dict
        HTTP-style response with statusCode and JSON body.
    """
    # ── Full raw event log - no truncation so exact payload is visible ────────
    try:
        print(f"[HANDLER] Full raw event:\n{json.dumps(event, indent=2)}")
    except (TypeError, ValueError):
        print(f"[HANDLER] Full raw event (repr):\n{event!r}")

    invoked_at = datetime.now(timezone.utc).isoformat()
    wall_start = time.monotonic()

    # ── Initialise audit record ───────────────────────────────────────────────
    audit: dict = {
        "started_at":    invoked_at,
        "finished_at":   None,
        "duration_s":    None,
        "product_id":    None,
        "input_key":     None,
        "status":        "STARTED",
        "status_code":   None,
        "zips_created":  0,
        "s3_keys":       [],
        "error":         None,
    }

    # ── Validate required environment variables ───────────────────────────────
    missing_env = [
        var for var in (
            "STIBO_BASE_URL",
            "STIBO_TOKEN_URL",
            "STIBO_CLIENT_ID",
            "STIBO_CLIENT_SECRET",
        )
        if not os.environ.get(var)
    ]
    if missing_env:
        msg = (
            f"Missing required environment variable(s): {', '.join(missing_env)}. "
            "Set them in the Lambda configuration."
        )
        audit.update(status="ERROR", status_code=500, error=msg)
        _write_audit_log(audit)
        return _error(500, msg)

    # ── Parse event ───────────────────────────────────────────────────────────
    try:
        key = _parse_event(event)
    except ValueError as exc:
        msg = str(exc)
        audit.update(status="ERROR", status_code=400, error=msg)
        _write_audit_log(audit)
        return _error(400, msg)

    product_id = Path(key).stem
    audit["input_key"]  = key
    audit["product_id"] = product_id

    print(f"[HANDLER] Input  : s3://{config.S3_BUCKET}/{key}")

    # ── Work inside a temp directory (auto-cleaned in finally) ────────────────
    tmp_dir = Path(tempfile.mkdtemp(dir="/tmp"))

    try:
        # ── Step 1: Download product XML from S3 ──────────────────────────────
        print("[HANDLER] Downloading XML ...")
        try:
            xml_bytes = download_bytes(key)
        except RuntimeError as exc:
            msg = str(exc)
            audit.update(status="ERROR", status_code=404, error=msg)
            _write_audit_log(audit)
            return _error(404, msg)

        xml_path = tmp_dir / Path(key).name
        xml_path.write_bytes(xml_bytes)
        print(f"[HANDLER] XML written to {xml_path}  ({len(xml_bytes) // 1024} KB)")

        # ── Step 2: Run the pipeline ──────────────────────────────────────────
        output_dir = tmp_dir / "output"
        output_dir.mkdir()

        try:
            zip_paths = run_pipeline(
                xml_path=xml_path,
                output_dir=output_dir,
            )
        except RuntimeError as exc:
            msg = f"Pipeline error: {exc}"
            audit.update(status="ERROR", status_code=500, error=msg)
            _write_audit_log(audit)
            return _error(500, msg)
        except ValueError as exc:
            msg = f"XML parse error: {exc}"
            audit.update(status="ERROR", status_code=400, error=msg)
            _write_audit_log(audit)
            return _error(400, msg)

        if not zip_paths:
            msg = (
                "Pipeline produced no output. "
                "Check that the XML contains valid assets and recognised channel IDs."
            )
            audit.update(status="NO_OUTPUT", status_code=422, error=msg)
            _write_audit_log(audit)
            return _error(422, msg)

        # ── Step 3: Upload ZIPs to S3 ─────────────────────────────────────────
        s3_keys: list[str] = []

        print(f"\n[HANDLER] Uploading {len(zip_paths)} ZIP(s) to S3 ...")

        for zip_path in zip_paths:
            try:
                s3_key = upload_zip(
                    filename=zip_path.name,
                    data=zip_path.read_bytes(),
                )
                s3_keys.append(s3_key)
                print(f"  Uploaded -> s3://{config.S3_BUCKET}/{s3_key}")
            except RuntimeError as exc:
                # Log but continue - partial upload is better than total failure
                print(f"  [ERROR] Failed to upload {zip_path.name}: {exc}")

        if not s3_keys:
            msg = "All ZIP uploads to S3 failed."
            audit.update(status="ERROR", status_code=500, error=msg)
            _write_audit_log(audit)
            return _error(500, msg)

        # ── Step 4: Build success response ────────────────────────────────────
        finished_at = datetime.now(timezone.utc).isoformat()
        duration_s  = round(time.monotonic() - wall_start, 3)

        audit.update(
            finished_at  = finished_at,
            duration_s   = duration_s,
            status       = "SUCCESS",
            status_code  = 200,
            zips_created = len(s3_keys),
            s3_keys      = s3_keys,
            error        = None,
        )
        _write_audit_log(audit)

        response_body = {
            "product_id":   product_id,
            "input_key":    key,
            "zips_created": len(s3_keys),
            "s3_keys":      s3_keys,
        }

        print(f"\n[HANDLER] Done - {len(s3_keys)} ZIP(s) uploaded in {duration_s}s.")
        for k in s3_keys:
            print(f"    s3://{config.S3_BUCKET}/{k}")

        return _success(response_body)

    finally:
        # ── Always clean up /tmp - safe for warm Lambda reuse ─────────────────
        shutil.rmtree(tmp_dir, ignore_errors=True)
        print(f"[HANDLER] Cleaned up {tmp_dir}")
