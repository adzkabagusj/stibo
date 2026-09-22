# router.py

import json
import logging
from urllib.parse import unquote_plus

# ----------------------------------------------------------
# Brand Handlers (latest folder structure)
# ----------------------------------------------------------
import adidas.lambda_function as adidas_handler
import new_balance.lambda_function as new_balance_handler
import smiggle.lambda_function as smiggle_handler
import aldo.lambda_function as aldo_handler
import diadora.lambda_function as diadora_handler
import crocs.lambda_function as crocs_handler
import lotto.lambda_function as lotto_handler
import onr.lambda_function as onr_handler   # ON Running (taf_main, taf_ss26_main, planet_sport_main, linelist_main)
import nike.lambda_function as nike_handler 
import dr_marten.lambda_function as dr_marten_handler  # DR Marten (retail_price_master_main)
import birkenstock.lambda_function as birkenstock_handler  # Birkenstock (OFS_FC_Mitra_ss27)
import staccato.lambda_function as staccato_handler
import pazzion.lambda_function as pazzion_handler
import implus.lambda_function as implus_handler
from audit_logger import AuditLogger   # ← shared logger
from metadata_s3 import ROOT_METADATA_KEYS
import twoxu.lambda_function as twoxu_handler   # 2XU brand (folder named twoxu/ — Python ids can't start with digit)
import clarks.lambda_function as clarks_handler  # Clarks (order_form_accs_main)
import steve_madden.lambda_function as steve_madden_handler  # Steve Madden (footwear_po_main)
import asics.lambda_function as asics_handler  # Asics (footwear_main)
import PTP.lambda_function as ptp_handler
import astec.lambda_function as astec_handler
import airwalk.lambda_function as airwalk_handler
import ellesse.lambda_function as ellesse_handler
import k_swiss.lambda_function as k_swiss_handler
import reebok.lambda_function as reebok_handler
import anta.lambda_function as anta_handler  # Anta (linelist_main — mapping-driven ETL)
import new_era.lambda_function as new_era_handler  # New Era (order_form_acc_main — mapping-driven ETL)
import anta_1.lambda_function as anta_1_handler  # Anta (linelist_main — mapping-driven ETL)

logging.basicConfig(level=logging.INFO)
log = logging.getLogger("router")

# ----------------------------------------------------------
# File types that live globally at raw/metadata/ root
# ----------------------------------------------------------
GLOBAL_TYPES = {"mdd", "attributes"}
ROOT_GLOBAL_KEYS = set(ROOT_METADATA_KEYS.values())


def parse_event(event):
    detail = event.get("detail", {})
    bucket = detail.get("bucket", {}).get("name", "")
    key = detail.get("object", {}).get("key", "")
    return bucket, key


def extract_principal(key: str) -> str | None:
    """
    raw/metadata/adidas/file.xlsx      → "adidas"       (brand-specific file)
    raw/metadata/new-balance/file.xlsx → "new-balance"  (brand-specific file)
    raw/metadata/file.xlsx             → None            (global file)
    """
    parts = key.strip("/").split("/")
    if len(parts) >= 4:
        return parts[2].lower()
    return None  # ← global file sitting directly under raw/metadata/


def _is_global_file(key: str) -> bool:
    """
    Returns True if the uploaded file sits directly under raw/metadata/
    (i.e. it is a global MDD or Attributes file), not inside a brand sub-folder.

    raw/metadata/Master Data Dictionary.xlsx         → True
    raw/metadata/Attributes List_complete_v6_25032026.xlsx  -> True
    raw/metadata/adidas/linelist.xlsx                → False
    raw/metadata/new-balance/linelist.xlsx           → False
    """
    key = key.strip("/")
    if key in ROOT_GLOBAL_KEYS:
        return True

    parts = key.split("/")
    # raw / metadata / filename  = exactly 3 parts  → global
    return len(parts) == 3 and _detect_global_file_type(parts[-1]) in GLOBAL_TYPES


def _detect_global_file_type(filename: str) -> str | None:
    name = filename.lower()
    if "attribute" in name:
        return "attributes"
    if "mdd" in name or "master data" in name:
        return "mdd"
    return None


# ----------------------------------------------------------
# Active Brands
# ----------------------------------------------------------

BRAND_ROUTER = {
    "adidas":adidas_handler.lambda_handler,
    "new-balance": new_balance_handler.lambda_handler,
    "smiggle": smiggle_handler.lambda_handler,
    "aldo": aldo_handler.lambda_handler,
    "diadora": diadora_handler.lambda_handler,
    "crocs":  crocs_handler.lambda_handler,
    "2xu":    twoxu_handler.lambda_handler,   # S3 folder: raw/metadata/2xu/
    "lotto":  lotto_handler.lambda_handler,
    "onr":    onr_handler.lambda_handler,     
    "on-running": onr_handler.lambda_handler,
    "nike":   nike_handler.lambda_handler,
    "dr-martens": dr_marten_handler.lambda_handler,  # S3 folder: raw/metadata/dr-marten/
    "staccato": staccato_handler.lambda_handler,
    "birkenstock": birkenstock_handler.lambda_handler,  # S3 folder: raw/metadata/birkenstock/
    "pazzion": pazzion_handler.lambda_handler,
    "clarks": clarks_handler.lambda_handler,  # S3 folder: raw/metadata/clarks/
    "steve-madden": steve_madden_handler.lambda_handler,  # S3 folder: raw/metadata/steve_madden/
    "asics": asics_handler.lambda_handler,  # S3 folder: raw/metadata/asics/
    "ptp":     ptp_handler.lambda_handler,
    "implus": implus_handler.lambda_handler,  # S3 folder: raw/metadata/implus/
    "astec": astec_handler.lambda_handler,
    "airwalk": airwalk_handler.lambda_handler,
    "ellesse": ellesse_handler.lambda_handler,
    "k-swiss": k_swiss_handler.lambda_handler,
    "reebok": reebok_handler.lambda_handler,
    "anta": anta_handler.lambda_handler,  # S3 folder: raw/metadata/anta/
    "new-era": new_era_handler.lambda_handler,  # S3 folder: raw/metadata/new-era/
    "anta-1": anta_1_handler.lambda_handler,  # S3 folder: raw/metadata/anta/
}


def lambda_handler(event, context):

    log.info("Event Received: %s", json.dumps(event))

    bucket, key = parse_event(event)
    key = unquote_plus(key)

    if not key.startswith("raw/metadata/") and not _is_global_file(key):
        return {
            "statusCode": 200,
            "message": "Ignored. Outside metadata folder"
        }

    # Global file uploaded (mdd / attributes at raw/metadata/ root)
    # Fan out to ALL active brands so each one re-runs its ETL
    # and picks up the updated global file via _list_principal_files().
    if _is_global_file(key):
        log.info(
            "Global file detected: '%s' — fanning out ETL to all active brands: %s",
            key, list(BRAND_ROUTER.keys())
        )
        results = {}
        for brand, handler in BRAND_ROUTER.items():
            log.info("  Triggering ETL for brand: %s", brand)

            # Each brand gets its OWN auditor so each produces its own log file.
            auditor = AuditLogger(event=event, context=context)
            auditor.set_router_context(
                trigger_type    = "global_fanout",
                original_key    = key,
                detected_brand  = brand,
                routing_decision = f"global file triggered fanout to all brands; routing to '{brand}'",
            )

            # Build a synthetic event that looks like a brand-specific upload.
            # lambda_function._extract_principal() reads parts[2] from the key,
            # so we point it at raw/metadata/{brand}/ with a harmless placeholder
            # filename that won't match any FILE_TYPE_KEYWORDS and will be skipped.
            synthetic_event = {
                **event,   # preserve source-account / region fields
                "detail": {
                    "bucket": {"name": bucket},
                    "object": {"key": f"raw/metadata/{brand}/.__global_trigger__"},
                },
            }
            try:
                results[brand] = handler(synthetic_event, context, auditor)
            except Exception as e:
                log.error("  ETL failed for brand '%s': %s", brand, e)
                auditor.set_lambda_status("error", error=str(e))
                results[brand] = {"statusCode": 500, "error": str(e)}
            finally:
                # Always flush — even if the brand handler crashed
                auditor.flush(brand)

        return {
            "statusCode": 200,
            "trigger":    "global_file",
            "file":       key,
            "fanout":     results,
        }

    # ── Brand-specific file uploaded ────────────────────────────────────────
    principal = extract_principal(key)

    if not principal:
        log.error("Unable to detect brand folder from key: %s", key)
        return {
            "statusCode": 400,
            "error": "Unable to detect brand folder"
        }

    log.info("Detected Brand: %s", principal)

    handler = BRAND_ROUTER.get(principal)

    if not handler:
        log.warning("No processor configured for brand: %s", principal)
        return {
            "statusCode": 400,
            "error": f"No processor configured for brand: {principal}"
        }

    log.info("Routing to handler: %s", handler.__module__)

    # Create ONE auditor for this brand invocation
    auditor = AuditLogger(event=event, context=context)
    auditor.set_router_context(
        trigger_type     = "brand_specific",
        original_key     = key,
        detected_brand   = principal,
        routing_decision = f"brand-specific file upload; routing to '{principal}' handler",
    )

    try:
        result = handler(event, context, auditor)
    except Exception as e:
        log.error("ETL failed for brand '%s': %s", principal, e)
        auditor.set_lambda_status("error", error=str(e))
        result = {"statusCode": 500, "error": str(e)}
    finally:
        # Always flush — even if the brand handler crashed
        auditor.flush(principal)

    return result
