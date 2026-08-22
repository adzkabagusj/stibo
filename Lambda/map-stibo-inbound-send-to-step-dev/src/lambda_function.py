"""
╔══════════════════════════════════════════════════════════════════╗
║          STIBO INBOUND UPLOADER  — v2.3  (AWS Lambda)          ║
║  OIDC Client Credentials → upload XML to Stibo inbound IIEP   ║
╚══════════════════════════════════════════════════════════════════╝

Environment Variables (set in Lambda console or via Secrets Manager):

  OIDC / Auth:
    STIBO_TOKEN_URL      — OIDC token endpoint URL
    STIBO_CLIENT_ID      — OIDC client ID
    STIBO_CLIENT_SECRET  — OIDC client secret  ← store in Secrets Manager
    STIBO_GRANT_TYPE     — (optional) default: client_credentials

  Inbound endpoint URLs — one per IIEP:
    STIBO_INBOUND_URL_ARTICLE_PLANNING   — e.g. …/IIEP_ArticlePlanning/upload-direct
    STIBO_INBOUND_URL_EAN_UPDATE         — e.g. …/IIEP_EANUpdate/upload-direct
    STIBO_INBOUND_URL_ARTICLE_MAINTENANCE— e.g. …/IIEP_ArticleMaintenance/upload-direct

  Tuning (all optional):
    STIBO_MAX_RETRIES    — default: 3
    STIBO_RETRY_BACKOFF  — default: 2 (seconds, doubles each retry)
    STIBO_TOKEN_LEEWAY   — default: 30 (seconds before expiry to refresh)
    STIBO_CONTEXT        — default: Context1
    STIBO_WORKSPACE      — default: Main

Lambda trigger:
  S3 event notification on map-stibo-inbound-processed-dev
  Prefix:  processed/stepxml/
  Suffix:  .xml
  Handler: stibo_uploader.lambda_handler

Lambda Event Payload (two modes):

  Mode A — S3 trigger (automatic when XML lands in processed bucket):
    {
      "Records": [{
        "s3": {
          "bucket": { "name": "map-stibo-inbound-processed-dev" },
          "object": { "key": "processed/stepxml/adidas/0888-SP-Adidas-Line%20List-Multi-SS26-1.xml" }
        }
      }]
    }

  Mode B — direct invocation (endpoint_type is optional — overrides filename routing):
    {
      "xml_s3_bucket":  "map-stibo-inbound-processed-dev",
      "xml_s3_key":     "processed/stepxml/adidas/0888-SP-Adidas-Line List-Multi-SS26-1.xml",
      "blank_ids":      true,
      "endpoint_type":  "ARTICLE_PLANNING"   ← optional override
    }

  Valid endpoint_type values: ARTICLE_PLANNING | EAN_UPDATE | ARTICLE_MAINTENANCE

Required Lambda IAM permissions:
  - s3:GetObject   on map-stibo-inbound-processed-dev
  - s3:PutObject   on map-stibo-inbound-processed-dev  (for bgId receipts)
  - secretsmanager:GetSecretValue  (if using Secrets Manager for client_secret)

Required layers / dependencies:
  - None — uses only Python stdlib (urllib) + boto3 (built into Lambda runtime)

──────────────────────────────────────────────────────────────────
FILENAME → ENDPOINT ROUTING
──────────────────────────────────────────────────────────────────

Routing is determined by case-insensitive substring match against
the filename (not the full S3 key). Rules are evaluated in order;
the first match wins.

  ARTICLE_PLANNING    — Line List, Linelist, Recap Sample, Price List
  EAN_UPDATE          — Backlog, EAN Source, Order Sheet, Packing List, OOR
  ARTICLE_MAINTENANCE — TDD, DTB, Ecommerce File, Ecommerce

Add or adjust patterns in FILENAME_ROUTING_RULES below as new
file types are onboarded — no code changes needed elsewhere.
"""

from __future__ import annotations

import json
import logging
import os
import re
import sys
import time
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode, unquote_plus, quote
from urllib.request import Request, urlopen

import boto3

# ──────────────────────────────────────────────────────────────────
# LOGGING  — Lambda writes to CloudWatch automatically
# ──────────────────────────────────────────────────────────────────
log = logging.getLogger(__name__)
log.setLevel(logging.INFO)

if not log.handlers:
    log.addHandler(logging.StreamHandler(sys.stdout))


# ══════════════════════════════════════════════════════════════════
# ENDPOINT ROUTING
# ══════════════════════════════════════════════════════════════════

# Maps each logical endpoint key to the env-var that holds its URL.
ENDPOINT_ENV_VARS: dict[str, str] = {
    "ARTICLE_PLANNING":    "STIBO_INBOUND_URL_ARTICLE_PLANNING",
    "EAN_UPDATE":          "STIBO_INBOUND_URL_EAN_UPDATE",
    "ARTICLE_MAINTENANCE": "STIBO_INBOUND_URL_ARTICLE_MAINTENANCE",
}

# Ordered list of (pattern, endpoint_key) pairs.
# Patterns are matched case-insensitively against the bare filename.
# First match wins — IMPORTANT: put more-specific / compound patterns
# before the broader single-keyword patterns they overlap with.
#
# Real filename format: 0888-{FQ|SP}-{Brand}-{FileType}-{Mono|Multi}-{Season}-{n}.xlsx
# Routing is matched against the full filename string.
#
# Overlap to watch:
#   "TDD and Price List" contains "Price List" → TDD rule must come first
#   "Packing List"       contains "List"       → already fine (no bare "list" rule)
FILENAME_ROUTING_RULES: list[tuple[str, str]] = [
    # ── Article Maintenance — MUST come before price\s*list ───────
    (r"\btdd\b",                  "ARTICLE_MAINTENANCE"),
    (r"\bdtb\b",                  "ARTICLE_MAINTENANCE"),
    (r"ecommerce|eommerce",       "ARTICLE_MAINTENANCE"),

    # ── Article Planning — MUST come before order\s*form (EAN) ───
    (r"line\s*list",              "ARTICLE_PLANNING"),
    (r"linesheet",                "ARTICLE_PLANNING"),
    (r"recap\s*sample",           "ARTICLE_PLANNING"),
    (r"price\s*list",             "ARTICLE_PLANNING"),
    (r"price\s*master",           "ARTICLE_PLANNING"), 
    (r"product\s*bible",          "ARTICLE_PLANNING"),     # 2XU
    (r"\binline\b",               "ARTICLE_PLANNING"),
    (r"receive.*linelist",        "ARTICLE_PLANNING"),
    (r"1st\s*revise.*linelist",   "ARTICLE_PLANNING"),
    (r"2nd\s*revise.*linelist",   "ARTICLE_PLANNING"),
    (r"ofs[-\s]*fc[-\s]*pt[-\s]*mitra", "ARTICLE_PLANNING"),
    (r"fob.*(order\s*form|orderform)", "ARTICLE_PLANNING"), # FOB Order Form only — requires "fob"
    (r"lotto.*(order\s*form|orderform)|(order\s*form|orderform).*lotto", "ARTICLE_PLANNING"),  # Lotto Order Form

    # ── EAN Update ───────────────────────────────────────────────
    (r"backlog",                  "EAN_UPDATE"),
    (r"ean\s*source",             "EAN_UPDATE"),
    (r"order\s*sheet",            "EAN_UPDATE"),
    (r"packing\s*list",           "EAN_UPDATE"),
    (r"packaging\s*list",         "EAN_UPDATE"),
    (r"\boor\b",                  "EAN_UPDATE"),
    (r"order\s*confirmation",     "EAN_UPDATE"),           # 2XU, Nike, Crocs
    (r"open\s*shipment",          "EAN_UPDATE"),
    (r"shipment\s*confirmation",  "EAN_UPDATE"),           # Crocs
    (r"order\s*form",             "EAN_UPDATE"), 
    (r"\btaf\b",                  "EAN_UPDATE"),           # ONR
    (r"planet\s*sport",           "EAN_UPDATE"),           # ONR
    (r"fob\s*order(?!\s*form|form)", "EAN_UPDATE"),       
]

# Pre-compile for efficiency
_COMPILED_RULES: list[tuple[re.Pattern, str]] = [
    (re.compile(pattern, re.IGNORECASE), endpoint_key)
    for pattern, endpoint_key in FILENAME_ROUTING_RULES
]


def resolve_endpoint(filename: str, override: str | None = None) -> tuple[str, str]:
    """
    Return (endpoint_url, endpoint_key) for the given filename.

    If `override` is provided (e.g. from direct-invocation payload),
    it bypasses filename matching and is used directly.

    Raises ValueError if no rule matches and no override is given.
    Raises EnvironmentError if the resolved env-var is not set.
    """
    if override:
        endpoint_key = override.upper().strip()
        if endpoint_key not in ENDPOINT_ENV_VARS:
            raise ValueError(
                f"Unknown endpoint_type '{override}'. "
                f"Valid values: {', '.join(ENDPOINT_ENV_VARS)}"
            )
        log.info("[Route] Override → %s", endpoint_key)
    else:
        endpoint_key = None
        for pattern, key in _COMPILED_RULES:
            if pattern.search(filename):
                endpoint_key = key
                log.info("[Route] '%s' matched pattern '%s' → %s",
                         filename, pattern.pattern, endpoint_key)
                break

        if endpoint_key is None:
            raise ValueError(
                f"No routing rule matched filename '{filename}'. "
                "Add a pattern to FILENAME_ROUTING_RULES or pass "
                "endpoint_type explicitly in the event payload."
            )

    env_var = ENDPOINT_ENV_VARS[endpoint_key]
    url = os.environ.get(env_var, "").strip()
    if not url:
        raise EnvironmentError(
            f"Endpoint URL env var '{env_var}' is not set. "
            f"Add it to the Lambda environment variables for endpoint '{endpoint_key}'."
        )

    return url, endpoint_key


# ──────────────────────────────────────────────────────────────────
# CONFIG  — sourced exclusively from environment variables
# ──────────────────────────────────────────────────────────────────

def _require_env(key: str) -> str:
    val = os.environ.get(key, "").strip()
    if not val:
        raise EnvironmentError(
            f"Required environment variable '{key}' is not set. "
            "Configure it in the Lambda environment variables console "
            "or via AWS Systems Manager Parameter Store."
        )
    return val


def _optional_env(key: str, default: str) -> str:
    return os.environ.get(key, default).strip() or default


def load_config() -> dict:
    return {
        # Auth — required
        "token_url":       _require_env("STIBO_TOKEN_URL"),
        "client_id":       _require_env("STIBO_CLIENT_ID"),
        "client_secret":   _require_env("STIBO_CLIENT_SECRET"),
        # Auth — optional
        "grant_type":      _optional_env("STIBO_GRANT_TYPE",   "client_credentials"),
        # Stibo query params
        "stibo_context":   _optional_env("STIBO_CONTEXT",      "Context1"),
        "stibo_workspace": _optional_env("STIBO_WORKSPACE",    "Main"),
        # Retry / token tuning
        "max_retries":     int(_optional_env("STIBO_MAX_RETRIES",   "3")),
        "retry_backoff":   int(_optional_env("STIBO_RETRY_BACKOFF", "2")),
        "token_leeway":    int(_optional_env("STIBO_TOKEN_LEEWAY",  "30")),
    }


# ──────────────────────────────────────────────────────────────────
# MODULE-LEVEL TOKEN CACHE
# ──────────────────────────────────────────────────────────────────
_token_cache: dict[str, Any] = {
    "access_token": None,
    "expires_at":   0.0,
}


# ══════════════════════════════════════════════════════════════════
# STEP ID BLANK-OUT
# ══════════════════════════════════════════════════════════════════

_CLH_ID = re.compile(r'(?<=\s)ID="CLH_[^"]*"')


def blank_step_ids(xml_text: str) -> str:
    """
    Blank only Classification IDs with CLH_ prefix
    (season/confirmed/unconfirmed nodes).

    Products in this XML carry no ID attribute — identity is via
    <KeyValue> elements — so Product tags are left completely untouched.
    """
    def _strip_classification(m: re.Match) -> str:
        tag = m.group(0)
        if _CLH_ID.search(tag):
            return _CLH_ID.sub('ID=""', tag)
        return tag

    xml_text = re.sub(r"<Classification\b[^>]*>", _strip_classification, xml_text)
    return xml_text



# ══════════════════════════════════════════════════════════════════
# OIDC TOKEN
# ══════════════════════════════════════════════════════════════════

def get_access_token(cfg: dict) -> str:
    now = time.time()

    if _token_cache["access_token"] and now < _token_cache["expires_at"] - cfg["token_leeway"]:
        log.info("[OIDC] Reusing cached token (%.0fs remaining)",
                 _token_cache["expires_at"] - now)
        return _token_cache["access_token"]

    log.info("[OIDC] Fetching new token from %s", cfg["token_url"])

    body = urlencode({
        "grant_type":    cfg["grant_type"],
        "client_id":     cfg["client_id"],
        "client_secret": cfg["client_secret"],
    }).encode("utf-8")

    req = Request(
        cfg["token_url"],
        data    = body,
        headers = {"Content-Type": "application/x-www-form-urlencoded"},
        method  = "POST",
    )

    try:
        with urlopen(req, timeout=30) as resp:
            payload = json.loads(resp.read().decode("utf-8"))
    except HTTPError as exc:
        body_text = exc.read().decode("utf-8", errors="replace")[:300]
        log.error("[OIDC] Token fetch failed: HTTP %d — %s", exc.code, body_text)
        raise

    token      = payload["access_token"]
    expires_in = int(payload.get("expires_in", 300))

    _token_cache["access_token"] = token
    _token_cache["expires_at"]   = now + expires_in

    log.info("[OIDC] Token obtained — expires in %ds", expires_in)
    return token


# ══════════════════════════════════════════════════════════════════
# S3 HELPERS
# ══════════════════════════════════════════════════════════════════

def read_xml_from_s3(bucket: str, key: str) -> str:
    """Download XML from S3 and return as a UTF-8 string."""
    log.info("[S3] Reading s3://%s/%s", bucket, key)
    obj = boto3.client("s3").get_object(Bucket=bucket, Key=key)
    return obj["Body"].read().decode("utf-8")


def _save_bgid_receipt(bucket: str, key: str, filename: str,
                       bg_id: str, endpoint_key: str) -> None:
    """
    Save a small JSON receipt alongside the XML in S3 after a successful
    Stibo upload, recording the bgId and which endpoint was used.

    Key: processed/stepxml/bgid/{principal}/{filename}.json
    """
    receipt = {
        "filename":      filename,
        "s3_key":        key,
        "bg_id":         bg_id,
        "endpoint":      endpoint_key,
        "uploaded_at":   time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
    }
    parts        = key.strip("/").split("/")
    principal    = parts[2] if len(parts) >= 4 else "unknown"
    receipt_key  = f"processed/stepxml/bgid/{principal}/{filename}.json"
    receipt_body = json.dumps(receipt, indent=2).encode("utf-8")

    boto3.client("s3").put_object(
        Bucket      = bucket,
        Key         = receipt_key,
        Body        = receipt_body,
        ContentType = "application/json",
    )
    log.info("[S3] bgId receipt saved → s3://%s/%s  (bgId=%s, endpoint=%s)",
             bucket, receipt_key, bg_id, endpoint_key)


# ══════════════════════════════════════════════════════════════════
# STIBO POST  — binary body (mirrors Postman "binary" mode)
# ══════════════════════════════════════════════════════════════════

def _post_to_stibo(xml_text: str, filename: str,
                   inbound_url: str, cfg: dict) -> str | bool:
    token = get_access_token(cfg)
    body  = xml_text.encode("utf-8")

    url = (
        f"{inbound_url}?"
        + urlencode(
            {
                "fileName":  filename,
                "context":   cfg["stibo_context"],
                "workspace": cfg["stibo_workspace"],
            },
            quote_via=quote
        )
    )

    req = Request(
        url,
        data    = body,
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type":  "application/octet-stream",
        },
        method  = "POST",
    )

    try:
        with urlopen(req, timeout=120) as resp:
            resp_body = resp.read().decode("utf-8", errors="replace")
            status    = resp.status
    except HTTPError as exc:
        status    = exc.code
        resp_body = exc.read().decode("utf-8", errors="replace")[:500]

    log.info("[Stibo] HTTP %d — %s", status, resp_body[:500])

    if 200 <= status < 300:
        try:
            bg_id = json.loads(resp_body).get("id", "")
        except (json.JSONDecodeError, Exception):
            bg_id = ""
        log.info("[Stibo] ✓ '%s' accepted — bgId: %s", filename, bg_id)
        return bg_id or True

    if 400 <= status < 500 and status != 429:
        log.error("[Stibo] ✗ Client error %d (will not retry): %s", status, resp_body[:300])
        raise HTTPError(url, status, resp_body, {}, None)

    raise URLError(f"HTTP {status}: {resp_body[:200]}")


def upload_xml(xml_text: str, filename: str,
               inbound_url: str, cfg: dict) -> str | bool:
    """Upload with retry and exponential back-off. Returns bgId string on success."""

    if os.environ.get("STIBO_DRY_RUN", "").lower() in ("1", "true", "yes"):
        _ = get_access_token(cfg)   # validate token endpoint is reachable
        log.info("[DryRun] Would POST '%s' (%d bytes) to %s",
                 filename, len(xml_text.encode("utf-8")), inbound_url)
        log.info("[DryRun] XML head: %s", xml_text[:500])
        return "DRY_RUN_BG_ID"

    max_retries = cfg["max_retries"]
    backoff     = cfg["retry_backoff"]

    for attempt in range(1, max_retries + 1):
        try:
            result = _post_to_stibo(xml_text, filename, inbound_url, cfg)
            if result:
                return result
        except HTTPError:
            return False
        except (URLError, OSError) as exc:
            log.warning("[Stibo] Attempt %d/%d failed: %s", attempt, max_retries, exc)

        if attempt < max_retries:
            wait = backoff ** attempt
            log.info("[Stibo] Retrying in %ds …", wait)
            time.sleep(wait)

    log.error("[Stibo] All %d attempts exhausted for '%s'", max_retries, filename)
    return False
# ══════════════════════════════════════════════════════════════════
# LAMBDA HANDLER
# ══════════════════════════════════════════════════════════════════

def lambda_handler(event: dict, context: Any) -> dict:
    log.info("[Lambda] Event: %s", json.dumps(event)[:500])

    # ── Step 1: Parse event ───────────────────────────────────────
    try:
        bucket, key, blank_ids, endpoint_override = _parse_event(event)
    except (KeyError, IndexError, ValueError) as exc:
        log.error("[Lambda] Bad event payload: %s", exc)
        return _response(400, f"Bad event payload: {exc}")

    filename = key.split("/")[-1]

    # ── Step 2: Resolve endpoint URL ─────────────────────────────
    try:
        inbound_url, endpoint_key = resolve_endpoint(filename, endpoint_override)
        log.info("[Lambda] Endpoint: %s → %s", endpoint_key, inbound_url)
    except (ValueError, EnvironmentError) as exc:
        log.error("[Lambda] Endpoint resolution failed: %s", exc)
        return _response(400, str(exc))

    # ── Step 3: Load config ───────────────────────────────────────
    try:
        cfg = load_config()
    except EnvironmentError as exc:
        log.error("[Lambda] Config error: %s", exc)
        return _response(500, str(exc))

    # ── Step 4: Read XML from S3 ──────────────────────────────────
    try:
        xml_text = read_xml_from_s3(bucket, key)
    except Exception as exc:
        log.error("[Lambda] S3 read failed: %s", exc)
        return _response(500, f"S3 read failed: {exc}")

    # ── Step 5: Blank Classification IDs (in-memory only) ─────────
    if blank_ids:
        log.info("[Lambda] Blanking CLH_ Classification IDs")
        xml_text = blank_step_ids(xml_text)

    # Structural sanity check
    log.info("[DEBUG] XML head: %s", xml_text[:1000])
    log.info("[DEBUG] XML tail: %s", xml_text[-300:])
    log.info("[DEBUG] xmlns count: %d", xml_text.count('xmlns='))
    log.info("[DEBUG] ns0: count: %d", xml_text.count('ns0:'))
    log.info("[DEBUG] <Value count: %d", xml_text.count('<Value '))
    log.info("[DEBUG] <ns0:Value count: %d", xml_text.count('<ns0:Value'))

    # ── Step 6: Upload to Stibo ───────────────────────────────────
    success = upload_xml(xml_text, filename, inbound_url, cfg)

    if success:
        bg_id = success if isinstance(success, str) else ""
        if bg_id:
            try:
                _save_bgid_receipt(bucket, key, filename, bg_id, endpoint_key)
            except Exception as exc:
                log.warning("[S3] Could not save bgId receipt: %s", exc)
        return _response(
            200,
            f"Uploaded '{filename}' to Stibo [{endpoint_key}] successfully. bgId={bg_id}"
        )

    return _response(
        502,
        f"Failed to upload '{filename}' to Stibo [{endpoint_key}] "
        f"after {cfg['max_retries']} attempts."
    )


# ──────────────────────────────────────────────────────────────────
# HELPERS
# ──────────────────────────────────────────────────────────────────

_EXPECTED_PREFIX = "processed/stepxml/"


def _parse_event(event: dict) -> tuple[str, str, bool, str | None]:
    blank_ids         = event.get("blank_ids", False)
    endpoint_override = event.get("endpoint_type")   # None if not present

    if "Records" in event:
        s3_rec = event["Records"][0]["s3"]
        key    = unquote_plus(s3_rec["object"]["key"])

        parts = key.strip("/").split("/")
        if not key.startswith(_EXPECTED_PREFIX) or len(parts) < 4:
            raise ValueError(
                f"Unexpected S3 key '{key}' — "
                f"must be under '{_EXPECTED_PREFIX}{{principal}}/'"
            )

        principal = parts[2]
        log.info("[Lambda] Principal detected: %s", principal)
        return s3_rec["bucket"]["name"], key, blank_ids, endpoint_override

    return event["xml_s3_bucket"], event["xml_s3_key"], blank_ids, endpoint_override


def _response(status_code: int, message: str) -> dict:
    return {
        "statusCode": status_code,
        "body":       json.dumps({"message": message}),
        "headers":    {"Content-Type": "application/json"},
    }


# ══════════════════════════════════════════════════════════════════
# LOCAL TEST ENTRY POINT
# ══════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    import argparse

    logging.basicConfig(
        level  = logging.INFO,
        format = "%(asctime)s  %(levelname)-8s  %(message)s",
    )

    ap = argparse.ArgumentParser(description="Local test — simulates Lambda invocation")
    ap.add_argument("bucket",  help="S3 bucket name")
    ap.add_argument("key",     help="S3 object key")
    ap.add_argument("--no-blank-ids",    action="store_true",
                    help="Skip blanking of Classification IDs")
    ap.add_argument("--endpoint-type",   default=None,
                    help="Override endpoint routing: ARTICLE_PLANNING | EAN_UPDATE | ARTICLE_MAINTENANCE")
    a = ap.parse_args()

    result = lambda_handler(
        {
            "xml_s3_bucket":  a.bucket,
            "xml_s3_key":     a.key,
            "blank_ids":      not a.no_blank_ids,
            "endpoint_type":  a.endpoint_type,
        },
        context=None,
    )
    print(json.dumps(result, indent=2))