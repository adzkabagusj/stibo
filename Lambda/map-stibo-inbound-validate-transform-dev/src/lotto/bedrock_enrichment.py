"""
bedrock_enrichment.py — AI attribute enrichment for Lotto "recap sample AI ingestion"
====================================================================================
Called by recap_main.run() (via recap_ai_main) between Pass 1 (map + validate) and
Pass 2 (XML).  For every mapped article it calls the enrichment Lambda in the Bedrock
sandbox account — one request per article — with the recap attributes and the
article's in-cell image.

The Lambda returns 10 LOV attributes (display value + LOV ID) and 10 AI translation
attributes.  Successful results are stored on the article as ``art["ai_attributes"]``
and written by _add_generic_values() after the rule-based attributes, so values the
ETL already set (e.g. AT_SportsCategoryEN) are never overwritten.

Failure never stops the ETL: an article whose enrichment fails is written without AI
attributes and the reason is recorded in the audit log.

Transport (first one configured wins):
    Function URL — HTTPS POST with a shared secret header; needs no IAM permission.
        BEDROCK_ENRICHMENT_URL        https://<id>.lambda-url.us-east-1.on.aws/
        BEDROCK_ENRICHMENT_TOKEN      shared secret (sent as the x-enrichment-token header)
    Lambda invoke — IAM-authenticated; the execution role needs lambda:InvokeFunction.
        BEDROCK_ENRICHMENT_FUNCTION_ARN

Other environment:
    BEDROCK_ENRICHMENT_ENABLED        "false" disables enrichment (kill switch); default "true"
    BEDROCK_ENRICHMENT_MAX_WORKERS    parallel requests, default 5
"""
from __future__ import annotations

import base64
import json
import logging
import os
import time
import urllib.error
import urllib.parse
import urllib.request
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Callable

log = logging.getLogger(__name__)

ENV_URL = "BEDROCK_ENRICHMENT_URL"
ENV_TOKEN = "BEDROCK_ENRICHMENT_TOKEN"
ENV_FUNCTION_ARN = "BEDROCK_ENRICHMENT_FUNCTION_ARN"
ENV_ENABLED = "BEDROCK_ENRICHMENT_ENABLED"
ENV_MAX_WORKERS = "BEDROCK_ENRICHMENT_MAX_WORKERS"

TOKEN_HEADER = "x-enrichment-token"
HTTP_TIMEOUT_S = 150            # the sandbox Lambda times out at 120 s
HTTP_ATTEMPTS = 3
RETRYABLE_HTTP_STATUS = {429, 502, 503, 504}

LOV_ATTRIBUTES = (
    "AT_Color", "AT_Silhouette", "AT_Fastening", "AT_MaterialUpper", "AT_Material",
    "AT_PatternPrint", "AT_Width", "AT_Occasion", "AT_Interest", "AT_SportsCategoryEN",
)
TRANSLATION_ATTRIBUTES = (
    "AT_ColorDescriptionID", "AT_ColorDescriptionTH", "AT_EComProductNameID", "AT_EComProductNameVN",
    "AT_ShortDescriptionID", "AT_ShortDescriptionMY", "AT_ShortDescriptionKH",
    "AT_LongDescriptionID", "AT_LongDescriptionPH", "AT_CareInstructionID",
)


# ── configuration ────────────────────────────────────────────────────────────

def enrichment_config() -> dict:
    """{"enabled", "mode" ("url" | "lambda"), "target", "reason"} from the environment."""
    if os.environ.get(ENV_ENABLED, "true").strip().lower() in ("false", "0", "no", "off"):
        return {"enabled": False, "mode": "", "target": "", "reason": f"{ENV_ENABLED} is false"}

    url = os.environ.get(ENV_URL, "").strip()
    if url:
        if not url.startswith("https://"):
            return {"enabled": False, "mode": "url", "target": url, "reason": f"{ENV_URL} must start with https://"}
        if not os.environ.get(ENV_TOKEN, "").strip():
            return {"enabled": False, "mode": "url", "target": url, "reason": f"{ENV_TOKEN} is not set"}
        return {"enabled": True, "mode": "url", "target": url, "reason": ""}

    function_arn = os.environ.get(ENV_FUNCTION_ARN, "").strip()
    if function_arn:
        return {"enabled": True, "mode": "lambda", "target": function_arn, "reason": ""}

    return {"enabled": False, "mode": "", "target": "",
            "reason": f"set {ENV_URL} + {ENV_TOKEN} (or {ENV_FUNCTION_ARN})"}


def make_invoker(cfg: dict) -> Callable[[dict], dict]:
    if cfg["mode"] == "url":
        token = os.environ[ENV_TOKEN].strip()
        return lambda payload: invoke_url(cfg["target"], token, payload)

    import boto3
    from botocore.config import Config

    arn = cfg["target"]
    client = boto3.client(
        "lambda", region_name=arn.split(":")[3],
        config=Config(read_timeout=HTTP_TIMEOUT_S, connect_timeout=10, retries={"max_attempts": 4, "mode": "adaptive"}),
    )
    return lambda payload: invoke_lambda(client, arn, payload)


# ── transports ───────────────────────────────────────────────────────────────

def invoke_url(url: str, token: str, payload: dict) -> dict:
    """POST to the Function URL; returns the JSON body (retries throttling and network errors)."""
    data = json.dumps(payload).encode("utf-8")
    for attempt in range(1, HTTP_ATTEMPTS + 1):
        request = urllib.request.Request(
            url, data=data, method="POST",
            headers={"content-type": "application/json", TOKEN_HEADER: token},
        )
        try:
            with urllib.request.urlopen(request, timeout=HTTP_TIMEOUT_S) as response:
                return json.loads(response.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            body = e.read().decode("utf-8", errors="replace")
            if e.code in RETRYABLE_HTTP_STATUS and attempt < HTTP_ATTEMPTS:
                time.sleep(2 ** attempt)
                continue
            try:
                parsed = json.loads(body)
                if isinstance(parsed, dict) and parsed.get("status"):
                    return parsed          # 401 / 422 / 500 carry a status body from the handler
            except ValueError:
                pass
            return {"status": "error", "supp_art": payload.get("supp_art"), "error": f"HTTP {e.code}: {body[:300]}"}
        except (urllib.error.URLError, TimeoutError, OSError) as e:
            if attempt < HTTP_ATTEMPTS:
                time.sleep(2 ** attempt)
                continue
            return {"status": "error", "supp_art": payload.get("supp_art"), "error": f"{type(e).__name__}: {e}"}
    return {"status": "error", "supp_art": payload.get("supp_art"), "error": "no attempts made"}


def invoke_lambda(client, function_arn: str, payload: dict) -> dict:
    response = client.invoke(FunctionName=function_arn, InvocationType="RequestResponse",
                             Payload=json.dumps(payload).encode("utf-8"))
    body = response["Payload"].read().decode("utf-8")
    if response.get("FunctionError"):
        return {"status": "error", "supp_art": payload.get("supp_art"), "error": body[:500]}
    return json.loads(body)


# ── enrichment ───────────────────────────────────────────────────────────────

def _fmt_date(value) -> str:
    if hasattr(value, "strftime"):
        return value.strftime("%d-%m-%Y")
    return str(value or "").strip()


def build_attributes(art: dict, brand: str) -> dict:
    """Recap input columns the enrichment model may see (names as in the recap template)."""
    attributes = {
        "Brand": brand.upper(),
        "Season": art.get("season_raw"),
        "Supplier": art.get("supplier"),
        "Division": art.get("division_col"),
        "MD Category": art.get("md_category"),
        "Color": art.get("colour"),
        "Color Code": art.get("colour_code"),
        "Gender": art.get("gender_raw"),
        "Age Group": art.get("age_group"),
        "Size Range": art.get("size_range"),
        "Outsole Material": art.get("outsole_material"),
        "Upper Material": art.get("upper_material"),
        "FOB Currency": art.get("currency"),
        "FOB Price": art.get("fob"),
        "ETA DATE": _fmt_date(art.get("eta_date")),
    }
    return {k: str(v).strip() for k, v in attributes.items() if v not in (None, "")}


def to_ai_attributes(result: dict) -> tuple[dict, list[str]]:
    """Valid attributes from a Lambda result: {attr_id: {"value", "id"}} plus skipped-attribute notes."""
    attributes, skipped = {}, []
    for attr_id in LOV_ATTRIBUTES:
        entry = (result.get("lov_attributes") or {}).get(attr_id) or {}
        if entry.get("value") and entry.get("lov_id"):
            attributes[attr_id] = {"value": entry["value"], "id": str(entry["lov_id"])}
        else:
            skipped.append(f"{attr_id}: no valid LOV value")
    for attr_id in TRANSLATION_ATTRIBUTES:
        text = str((result.get("translation_attributes") or {}).get(attr_id) or "").strip()
        if text:
            attributes[attr_id] = {"value": text, "id": ""}
        else:
            skipped.append(f"{attr_id}: empty")
    return attributes, skipped


def _safe_target(cfg: dict) -> str:
    """Target for logs: the URL host only (never query strings or secrets)."""
    if cfg["mode"] == "url":
        return urllib.parse.urlsplit(cfg["target"]).netloc
    return cfg["target"]


def enrich_articles(mapped_articles: list[dict], images: dict, brand: str = "LOTTO",
                    auditor=None, image_issues: list[dict] | None = None,
                    invoker: Callable[[dict], dict] | None = None) -> dict:
    """Enrich *mapped_articles* in place (``art["ai_attributes"]``) and return a summary.

    images: {excel_row: CellImage} from excel_images.extract_in_cell_images().
    invoker: request function (payload -> result dict); defaults to the configured transport.
    """
    started = time.perf_counter()
    cfg = enrichment_config()
    summary = {"status": "ok", "requested": len(mapped_articles), "succeeded": 0, "failed": 0,
               "skipped_no_image": 0, "mode": cfg["mode"], "target": _safe_target(cfg)}
    warnings: list[str] = []
    for issue in image_issues or []:
        warnings.append(f"[BedrockEnrichment] image {issue['severity']} row {issue['excel_row']} "
                        f"{issue['supp_art']}: {issue['code']} — {issue['message']}")

    if not cfg["enabled"]:
        summary.update(status="disabled", reason=cfg["reason"])
        log.warning("[BedrockEnrichment] disabled — %s; XML is written without AI attributes", cfg["reason"])
        _record(auditor, summary, warnings, started)
        return summary

    invoker = invoker or make_invoker(cfg)
    jobs = []
    for art in mapped_articles:
        image = images.get(art.get("excel_row"))
        if image is None:
            summary["skipped_no_image"] += 1
            warnings.append(f"[BedrockEnrichment] {art.get('article_no')} (row {art.get('excel_row')}): "
                            f"no usable in-cell image — written without AI attributes")
            continue
        jobs.append((art, {
            "supp_art": art.get("article_no"),
            "attributes": build_attributes(art, brand),
            "image_base64": base64.b64encode(image.data).decode("ascii"),
        }))

    max_workers = max(1, int(os.environ.get(ENV_MAX_WORKERS, "5")))
    log.info("[BedrockEnrichment] %s %s for %d article(s), %d worker(s)",
             cfg["mode"], summary["target"], len(jobs), max_workers)
    with ThreadPoolExecutor(max_workers=max_workers) as pool:
        futures = {pool.submit(invoker, payload): art for art, payload in jobs}
        for future in as_completed(futures):
            art = futures[future]
            article_no = art.get("article_no")
            try:
                result = future.result()
            except Exception as e:  # network / permission errors: fall back per article
                result = {"status": "error", "error": f"{type(e).__name__}: {e}"}

            if result.get("status") != "ok":
                summary["failed"] += 1
                detail = result.get("error") or "; ".join(result.get("errors") or []) or result.get("status")
                warnings.append(f"[BedrockEnrichment] {article_no}: {result.get('status')} — {detail}")
                log.warning("[BedrockEnrichment] %s failed: %s", article_no, detail)
                continue

            attributes, skipped = to_ai_attributes(result)
            art["ai_attributes"] = attributes
            summary["succeeded"] += 1
            for note in skipped + [f"validation: {e}" for e in result.get("validation_errors") or []]:
                warnings.append(f"[BedrockEnrichment] {article_no}: {note}")
            log.info("[BedrockEnrichment] %s enriched: %d attribute(s) in %.1fs", article_no, len(attributes),
                     (result.get("latency_ms") or 0) / 1000)

    if summary["failed"] or summary["skipped_no_image"]:
        summary["status"] = "partial" if summary["succeeded"] else "failed"
    _record(auditor, summary, warnings, started)
    return summary


def _record(auditor, summary: dict, warnings: list[str], started: float) -> None:
    summary["duration_s"] = round(time.perf_counter() - started, 1)
    log.info("[BedrockEnrichment] %s", summary)
    if auditor:
        auditor.record_loader("bedrock_enrichment", **summary)
        if warnings:
            auditor.add_validation_warnings(warnings)
