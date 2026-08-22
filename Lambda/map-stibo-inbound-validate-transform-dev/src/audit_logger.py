"""
audit_logger.py
===============
Shared AuditLogger for the Stibo Inbound ETL pipeline.

One instance is created per Lambda invocation (in router.py).
It is passed down through:
    router.py → brand/lambda_function.py → brand/main.py

At the end of every invocation, router.py calls auditor.flush(brand)
which writes ONE JSON file to S3:
    s3://map-stibo-inbound-raw-dev/raw/logs/{brand}/{YYYYMMDD_HHMMSS}.json

This file is NEVER flushed by brand files — only by router.py.
"""

import datetime
import json
import logging
import os

import boto3

log = logging.getLogger("audit_logger")

AUDIT_BUCKET = os.environ["RAW_BUCKET"]
AUDIT_PREFIX = "raw/logs"


def _utcnow_iso() -> str:
    return datetime.datetime.utcnow().isoformat() + "Z"


class AuditLogger:
    """
    Collects structured evidence for one Lambda invocation.
    Covers logs from router.py, brand lambda_function.py, and brand main.py.
    Flushed ONCE by router.py at the very end.
    """

    def __init__(self, event: dict, context):
        self._start_ts  = datetime.datetime.utcnow()
        self.event      = event
        self.context    = context

        # ── Router section ────────────────────────────────────────────────
        self.router: dict = {
            "trigger_type":    None,   # "brand_specific" | "global_fanout"
            "original_key":    None,
            "detected_brand":  None,
            "routing_decision": None,
            "started_at":      _utcnow_iso(),
        }

        # ── Lambda_function section ───────────────────────────────────────
        self.lambda_fn: dict = {
            "files_classified": {},    # {ftype: filename}
            "files_missing":    [],
            "downloads":        [],    # [{ftype, filename, status}]
            "conversions":      [],    # [{filename, from_ext, to_ext, status}]
            "metadata_parsed":  {},    # comp_code, sbu, brand, season, seq
            "diff":             {},    # result from linelist_diff
            "xml_uploads":      [],    # list of s3 keys uploaded
            "status":           None,
            "error":            None,
        }

        # ── Main.py / ETL section ─────────────────────────────────────────
        self.etl: dict = {
            "loaders":           {},   # {loader_name: {status, rows/attrs/etc}}
            "articles_mapped":   0,
            "articles_written":  0,
            "articles_skipped":  0,
            "total_variants":    0,
            "validation_warnings_count": 0,
            "xml_file":          None,
            "xml_size_kb":       None,
            "status":            None,
            "error":             None,
        }

        # ── Per-article evidence ──────────────────────────────────────────
        self.articles: list = []

        # ── Validation warnings (full text) ──────────────────────────────
        self.validation_warnings: list = []

    # ═════════════════════════════════════════════════════════════════════
    # ROUTER methods
    # ═════════════════════════════════════════════════════════════════════

    def set_router_context(self, trigger_type: str, original_key: str,
                           detected_brand: str, routing_decision: str):
        self.router["trigger_type"]     = trigger_type
        self.router["original_key"]     = original_key
        self.router["detected_brand"]   = detected_brand
        self.router["routing_decision"] = routing_decision

    # ═════════════════════════════════════════════════════════════════════
    # LAMBDA_FUNCTION methods
    # ═════════════════════════════════════════════════════════════════════

    def set_files_classified(self, found: dict):
        """found = {ftype: {filename, key, last_modified}}"""
        self.lambda_fn["files_classified"] = {
            k: v["filename"] for k, v in found.items()
        }

    def set_files_missing(self, missing: list):
        self.lambda_fn["files_missing"] = missing

    def record_download(self, ftype: str, filename: str, status: str, error: str = None):
        self.lambda_fn["downloads"].append({
            "ftype":    ftype,
            "filename": filename,
            "status":   status,       # "ok" | "failed"
            "error":    error,
        })

    def record_conversion(self, filename: str, from_ext: str,
                          to_ext: str, status: str, error: str = None):
        self.lambda_fn["conversions"].append({
            "filename": filename,
            "from_ext": from_ext,
            "to_ext":   to_ext,
            "status":   status,       # "ok" | "failed"
            "error":    error,
        })

    def set_metadata_parsed(self, meta: dict):
        self.lambda_fn["metadata_parsed"] = meta

    def set_diff_result(self, result: dict):
        self.lambda_fn["diff"] = result

    def set_xml_uploads(self, uploaded: list):
        self.lambda_fn["xml_uploads"] = uploaded

    def set_lambda_status(self, status: str, error: str = None):
        self.lambda_fn["status"] = status
        self.lambda_fn["error"]  = error

    # ═════════════════════════════════════════════════════════════════════
    # MAIN.PY / ETL methods
    # ═════════════════════════════════════════════════════════════════════

    def record_loader(self, name: str, status: str, **kwargs):
        """
        Record outcome of any loader.
        kwargs can be anything: rows=245, attributes=120, lovs=18, etc.
        """
        self.etl["loaders"][name] = {"status": status, **kwargs}

    def set_etl_counts(self, mapped: int, written: int,
                       skipped: int, variants: int,
                       xml_file: str = None, xml_size_kb: int = None):
        self.etl["articles_mapped"]  = mapped
        self.etl["articles_written"] = written
        self.etl["articles_skipped"] = skipped
        self.etl["total_variants"]   = variants
        if xml_file:
            self.etl["xml_file"]     = xml_file
        if xml_size_kb is not None:
            self.etl["xml_size_kb"]  = xml_size_kb

    def set_etl_status(self, status: str, error: str = None):
        self.etl["status"] = status
        self.etl["error"]  = error

    def record_article(self, article_no: str, status: str,
                       warnings: list = None, null_fields: list = None):
        """
        status: "written" | "skipped" | "failed"
        """
        self.articles.append({
            "article_no":  article_no,
            "status":      status,
            "warnings":    warnings   or [],
            "null_fields": null_fields or [],
            "recorded_at": _utcnow_iso(),
        })

    def add_validation_warnings(self, warns: list):
        self.validation_warnings.extend(warns)
        self.etl["validation_warnings_count"] = len(self.validation_warnings)

    # ═════════════════════════════════════════════════════════════════════
    # FLUSH — called ONLY by router.py
    # ═════════════════════════════════════════════════════════════════════

    def flush(self, brand: str):
        """
        Write the complete audit document to S3.
        Called once by router.py in its finally block.

        S3 path:
            s3://map-stibo-inbound-raw-dev/raw/logs/{brand}/{YYYYMMDD_HHMMSS}.json
        """
        end_ts   = datetime.datetime.utcnow()
        duration = round((end_ts - self._start_ts).total_seconds(), 3)
        dt_str   = self._start_ts.strftime("%Y%m%d_%H%M%S")

        document = {
            "schema_version":   "1.0",
            "lambda_function":  os.environ.get("AWS_LAMBDA_FUNCTION_NAME", "unknown"),
            "aws_request_id":   getattr(self.context, "aws_request_id", "unknown"),
            "brand":            brand,
            "started_at":       self._start_ts.isoformat() + "Z",
            "finished_at":      end_ts.isoformat() + "Z",
            "duration_seconds": duration,

            # ── EVIDENCE: exact raw input ──────────────────────────────
            "input_event": self.event,

            # ── EVIDENCE: router decision ──────────────────────────────
            "router": self.router,

            # ── EVIDENCE: lambda_function outcomes ─────────────────────
            "lambda_function": self.lambda_fn,

            # ── EVIDENCE: ETL / main.py outcomes ──────────────────────
            "etl": self.etl,

            # ── EVIDENCE: per-article detail ──────────────────────────
            "articles": self.articles,

            # ── EVIDENCE: all validation warning texts ─────────────────
            "validation_warnings": self.validation_warnings,

            # ── EVIDENCE: top-level totals ─────────────────────────────
            "totals": {
                "articles_mapped":   self.etl.get("articles_mapped",  0),
                "articles_written":  self.etl.get("articles_written", 0),
                "articles_skipped":  self.etl.get("articles_skipped", 0),
                "total_variants":    self.etl.get("total_variants",   0),
                "validation_warnings": len(self.validation_warnings),
                "xml_uploads":       len(self.lambda_fn.get("xml_uploads", [])),
            },
        }

        key = f"{AUDIT_PREFIX}/{brand}/{dt_str}.json"

        try:
            s3 = boto3.client("s3")
            s3.put_object(
                Bucket      = AUDIT_BUCKET,
                Key         = key,
                Body        = json.dumps(document, indent=2, default=str).encode("utf-8"),
                ContentType = "application/json",
            )
            log.info("✅ Audit log written → s3://%s/%s", AUDIT_BUCKET, key)
        except Exception as exc:
            # Never let audit logging crash the ETL
            log.error("❌ Audit log upload failed: %s", exc)