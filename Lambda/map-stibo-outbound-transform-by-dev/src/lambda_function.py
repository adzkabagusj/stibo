"""
lambda_function.py
==================
Router - Blue Yonder & Musical Box ETL

All pipelines live inside this single Lambda deployment.
The router inspects the incoming event and delegates accordingly:

    EventBridge schedule event
        -> sftp_by_response_poller.poll_sftp()   (SFTP -> S3)
           Drops files into response/inbound/, which fires an S3
           event back into this same Lambda -> Step 2 below.

    S3 key: step-export/blueyonder/**
        -> by_outbound.handle_outbound()          (Step 1: Stibo -> BY)

    S3 key: response/inbound/rsp_musicalbox_*
        -> mb_inbound.handle_mb_response()        (MB: Musical Box -> Stibo)

    S3 key: response/inbound/rsp_buyingplan_*  (or any other inbound file)
        -> by_inbound.handle_by_response()        (Step 2: BY -> Stibo)

EventBridge rule must point its target at this Lambda.
The S3 trigger also points at this same Lambda.
"""

import datetime
import json
import logging
import os
import sys
from pathlib import Path
from urllib.parse import unquote_plus

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
    force=True,
)
log = logging.getLogger("lambda_by_etl")

from by_outbound             import handle_outbound, AuditLogger as OutboundAuditLogger
from by_inbound              import handle_by_response, AuditLogger as InboundAuditLogger
from mb_inbound              import handle_mb_response, AuditLogger as MBInboundAuditLogger
from sftp_by_response_poller import poll_sftp

OUTBOUND_PREFIX    = os.environ.get("PREFIX",          "step-export/blueyonder/")
INBOUND_PREFIX     = os.environ.get("INBOUND_PREFIX",  "response/inbound/")
INBOUND_EXTENSIONS = (".txt", ".csv", ".json")

MB_FILENAME_PREFIX = "musicalbox"

# ─────────────────────────────────────────────────────────────────────────────
# TIMEZONE
# ─────────────────────────────────────────────────────────────────────────────
_TZ_GMT7 = datetime.timezone(datetime.timedelta(hours=7))


def _parse_event(event):
    if "Records" in event:
        rec = event["Records"][0]["s3"]
        return rec["bucket"]["name"], unquote_plus(rec["object"]["key"])
    if "detail" in event:
        detail = event["detail"]
        if "bucket" in detail and "object" in detail:
            return detail["bucket"]["name"], unquote_plus(detail["object"]["key"])
    return None, None


def _is_scheduled_event(event) -> bool:
    source      = event.get("source", "")
    detail_type = event.get("detail-type", "")
    if source in ("aws.events", "aws.scheduler") and "Scheduled" in detail_type:
        return True
    if event.get("trigger") == "sftp-poll":
        return True
    return False


def lambda_handler(event, context):
    log.info("HANDLER STARTED - file=lambda_function.py (router)")
    log.info("Router invoked. source=%s detail-type=%s",
             event.get("source", "s3"),
             event.get("detail-type", "ObjectCreated"))

    if _is_scheduled_event(event):
        log.info("[Router] -> SFTP poller: sftp_by_response_poller.poll_sftp")
        return poll_sftp(event, context)

    src_bucket, s3_key = _parse_event(event)

    if not src_bucket or not s3_key:
        log.error("Could not parse bucket/key from event.")
        return {"statusCode": 400, "error": "Unparseable event"}

    etl_id   = int(datetime.datetime.now(_TZ_GMT7).strftime("%Y%m%d%H%M%S"))
    stem     = Path(s3_key).stem
    filename = Path(s3_key).name

    # ── Shared inbound folder ─────────────────────────────────────────────
    if s3_key.startswith(INBOUND_PREFIX):
        if not s3_key.lower().endswith(INBOUND_EXTENSIONS):
            log.warning("[Router] Skipping unsupported inbound file type: %s", s3_key)
            return {"statusCode": 200, "processed": 0, "reason": "unsupported_extension"}

        # Route by filename
        if MB_FILENAME_PREFIX in filename.lower():
            log.info("[Router] -> MB Inbound: mb_inbound.handle_mb_response (%s)", filename)
            auditor = MBInboundAuditLogger(etl_id=etl_id, raw_event=event, context=context)
            try:
                result = handle_mb_response(src_bucket, s3_key, auditor)
                return {"statusCode": 200, "processed": 1, "step": "mb_inbound", "result": result}
            except Exception as exc:
                log.error("[Router] MB Inbound pipeline failed: %s", exc)
                auditor.end_file("failed")
                return {"statusCode": 500, "step": "mb_inbound", "error": str(exc)}
            finally:
                auditor.flush(stem)

        else:
            log.info("[Router] -> Step 2 (Inbound): by_inbound.handle_by_response (%s)", filename)
            auditor = InboundAuditLogger(etl_id=etl_id, raw_event=event, context=context)
            try:
                result = handle_by_response(src_bucket, s3_key, auditor)
                return {"statusCode": 200, "processed": 1, "step": 2, "result": result}
            except Exception as exc:
                log.error("[Router] Inbound pipeline failed: %s", exc)
                auditor.end_file("failed")
                return {"statusCode": 500, "step": 2, "error": str(exc)}
            finally:
                auditor.flush(stem)

    # ── Outbound ──────────────────────────────────────────────────────────
    if s3_key.startswith(OUTBOUND_PREFIX):
        if not s3_key.lower().endswith(".xml"):
            log.warning("[Router] Skipping non-XML key in outbound prefix: %s", s3_key)
            return {"statusCode": 200, "processed": 0, "reason": "not_xml"}
        log.info("[Router] -> Step 1 (Outbound): by_outbound.handle_outbound")
        auditor = OutboundAuditLogger(etl_id=etl_id, raw_event=event, context=context)
        try:
            result = handle_outbound(src_bucket, s3_key, auditor)
            return result
        except Exception as exc:
            log.error("[Router] Outbound pipeline failed: %s", exc)
            auditor.end_file("failed")
            return {"statusCode": 500, "step": 1, "error": str(exc)}
        finally:
            auditor.flush(stem)

    log.warning(
        "[Router] Key does not match any known prefix. "
        "Outbound=%r  Inbound=%r  Key=%r",
        OUTBOUND_PREFIX, INBOUND_PREFIX, s3_key,
    )
    return {"statusCode": 200, "processed": 0, "reason": "unrecognised_prefix"}