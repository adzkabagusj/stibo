# MAA-Backend

Comprehensive technical documentation for the **Stibo inbound metadata-to-XML ETL backend**.

---

## 1) What this repository does

This codebase implements a multi-brand AWS Lambda pipeline that:

1. Watches for inbound product metadata files uploaded to S3 (`raw/metadata/...`).
2. Routes each upload to the correct brand-specific processor.
3. Validates required source files for that brand and file type.
4. Normalizes input formats (`.xlsx`, `.xlsb`, `.csv` where supported).
5. Runs brand ETL modules to generate STEP XML output files.
6. Uploads XML to a processed S3 bucket (`processed/stepxml/...`).
7. Produces per-invocation structured audit logs in S3 (`raw/logs/...`).

The architecture is intentionally split into:
- A **global router** (`router.py`) for brand dispatch and global file fan-out.
- A **shared audit logger** (`audit_logger.py`) for consistent observability.
- **Brand handlers** (`<brand>/lambda_function.py`) for file classification, pre-processing, ETL dispatch, and XML upload.
- **Brand ETL modules** (`<brand>/*.py`) that implement business transformations.

---

## 2) High-level architecture

### Entry point and control flow

- **Router Lambda entrypoint**: `router.lambda_handler(event, context)`.
- Router accepts EventBridge S3 Object Created events and only processes keys under:
  - `raw/metadata/`

### Two trigger paths

1. **Brand-specific upload** (e.g., `raw/metadata/new-balance/...`)
   - Router extracts principal (brand folder), selects handler from `BRAND_ROUTER`, and invokes exactly one brand handler.

2. **Global support file upload** (e.g., `raw/metadata/Master Data Dictionary.xlsx`)
   - Router treats root-level metadata files as global and fans out synthetic brand events to **all configured brands** so each ETL can refresh against updated global dependencies.

### Output artifacts

- XML: `s3://$PROCESSED_BUCKET/processed/stepxml/{brand}/{file}.xml`
- Audit log JSON: `s3://$RAW_BUCKET/raw/logs/{brand}/{YYYYMMDD_HHMMSS}.json`

---

## 3) Repository layout

```text
.
├── router.py                      # Global dispatcher/fan-out
├── audit_logger.py                # Shared structured audit collector + S3 flush
├── adidas/
│   ├── lambda_function.py         # Adidas handler
│   ├── main.py                    # Core Adidas ETL
│   ├── linelist_diff.py           # Diff helper
│   ├── backlog_main.py            # Backlog ETL path
│   ├── tdd_main.py                # TDD ETL path
│   └── dtb_main.py                # DTB ETL path
├── new_balance/
│   ├── lambda_function.py         # Unified NB handler
│   ├── inline_main_accessories.py
│   ├── inline_main_footwear.py
│   ├── licensed_linelist_main.py
│   ├── licensed_ecommerce_main.py
│   ├── inline_ecommerce_main.py
│   ├── inline_ean_source.py
│   └── licensed_ordersheet_main.py
├── smiggle/
│   ├── lambda_function.py         # Unified Smiggle handler
│   ├── linelist_main.py
│   ├── packing_list_main.py
│   └── ecommerce_main.py
├── diadora/
│   ├── lambda_function.py         # Unified Diadora handler
│   └── main.py
├── aldo/
│   ├── lambda_function.py         # Currently commented legacy draft/reference
│   ├── main.py
│   ├── ecommerce.py
│   └── oor.py
```

---

## 4) Router behavior (`router.py`)

### Brand registry

The router maps known principals to brand handlers:
- `adidas`
- `new-balance`
- `smiggle`
- `aldo`
- `diadora`

through `BRAND_ROUTER`.

### Principal extraction

For keys shaped like `raw/metadata/{principal}/{filename}`, principal is taken from path segment index `2`.

### Global-file detection and fan-out

A file is treated as global if it is directly under `raw/metadata/` (exactly 3 path segments). In that case:
- Router loops through all configured brands.
- Creates one `AuditLogger` per brand invocation.
- Builds synthetic key `raw/metadata/{brand}/.__global_trigger__` so each brand can re-run and pick latest global files.
- Flushes audit regardless of success/failure (`finally`).

### Fault handling

Each brand invocation is isolated:
- Exceptions are captured per brand.
- Invocation-level error is recorded in audit (`set_lambda_status("error")`).
- Router still returns a composite result for fan-out mode.

---

## 5) Shared observability (`audit_logger.py`)

`AuditLogger` captures structured execution evidence across all layers.

### Sections in audit document

1. **router**
   - trigger type, original key, detected brand, routing decision.

2. **lambda_function**
   - file classification
   - missing required types
   - per-file download outcomes
   - conversions (`xlsb/csv -> xlsx`)
   - parsed metadata
   - diffs
   - uploaded XML keys
   - status + error

3. **etl**
   - loader statuses
   - article/variant counts
   - validation warning counts
   - XML filename and size
   - status + error

4. **articles**
   - per-article status (`written/skipped/failed`) and warnings

5. **validation_warnings**
   - full warning strings

### Flush semantics

- Audit flush is performed **only by router** (not brand handlers).
- Flush target: `$RAW_BUCKET/raw/logs/{brand}/{timestamp}.json`.
- Includes runtime metadata (request id, duration, start/end timestamps, raw input event).

---

## 6) Brand handlers and processing logic

## 6.1 Adidas (`adidas/lambda_function.py`)

### Purpose
Unified ingestion handler for Adidas metadata files and ETL kick-off.

### Key points
- Uses isolated tmp root: `/tmp/stibo_workdir_adidas`.
- Supports type detection including: `linelist`, `backlog`, `tdd`, `dtb`, `mdd`, `attributes`, `naming` (+ additional shared/licensed keywords).
- Converts `.xlsb` and `.csv` to `.xlsx` where needed.
- Required type policy currently keyed by brand code (`ADI` requires `linelist`; default requires `linelist`).
- Invokes ETL modules:
  - `adidas.main`
  - `adidas.tdd_main`
  - `adidas.backlog_main`
  - `adidas.dtb_main`
- Uploads generated XML into processed bucket.

### Metadata parsing
Contains filename-driven metadata parsing utilities (e.g., season detection patterns such as `SP26`, `FW2026`).

---

## 6.2 New Balance (`new_balance/lambda_function.py`)

### Purpose
Single handler that dispatches to multiple NB ETL flows based on filename subtype.

### Detection order (important)
1. `linelist_footwear`
2. `linelist_apparel`
3. `ecommerce_licensed`
4. `ecommerce_inline`
5. generic `ecommerce`
6. `ordersheet_licensed`
7. `linelist_licensed`
8. shared support types (`mdd`, `attributes`, `naming`)

### ETL dispatch map
- `linelist_apparel` -> `inline_main_accessories`
- `linelist_footwear` -> `inline_main_footwear`
- `linelist_licensed` -> `licensed_linelist_main`
- `ecommerce_licensed` -> `licensed_ecommerce_main`
- `ecommerce_inline` -> `inline_ecommerce_main`
- `ean_source` -> `inline_ean_source`
- `ordersheet_licensed` -> `licensed_ordersheet_main`

### Required file sets
`REQUIRED_TYPES_BY_TRIGGER` enforces per-trigger dependency sets (e.g., `mdd`, `attributes`, `naming` plus the triggered business file).

---

## 6.3 Smiggle (`smiggle/lambda_function.py`)

### Purpose
Unified handler for Smiggle file families:
- linelist (catalogue/wholesale/MAPI)
- packing list
- ecommerce

### Distinct behavior
- Tmp root isolated to `/tmp/stibo_workdir_smiggle`.
- Recognizes extra linelist keywords (`catalogue`, `catalog`, `mapi`, `wholesale`, `handover`).
- Has trigger-aware required sets:
  - linelist trigger requires linelist + global support files
  - packing trigger requires packing + global support files
  - ecommerce trigger requires ecommerce + global support files

### ETL dispatch map
- `linelist` -> `smiggle.linelist_main`
- `packing` -> `smiggle.packing_list_main`
- `ecommerce` -> `smiggle.ecommerce_main`

---

## 6.4 Diadora (`diadora/lambda_function.py`)

### Purpose
Simpler unified handler focused on inline pricelist processing.

### Current business type
- `pricelist` (keywords: `pricelist`, `price list`, `price_list`)

### Requirements
- For `pricelist` trigger, requires `pricelist + mdd + attributes + naming`.

### ETL dispatch
- `pricelist` -> `diadora.main`

---

## 6.5 Aldo (`aldo/lambda_function.py`)

The file is currently fully commented out and appears to be a legacy/reference draft handler (not executable in present state). It documents intended support for:
- linelist / trenzashop
- OOR report + OOR size run
- MCR
- ecommerce
- global support files

If Aldo routing is active in `router.py`, production behavior depends on whether this file is later restored to executable code.

---

## 7) Input conventions

### S3 key patterns

- Brand-specific input: `raw/metadata/{brand}/{filename}`
- Global support input: `raw/metadata/{filename}`

### Supported inbound file extensions

Across handlers, accepted source formats typically include:
- `.xlsx`
- `.xlsb`
- `.csv`
- `.xlsm` (notably in Aldo draft logic)

### Filename keyword classification

Each handler determines file role by case-insensitive substring matching against local keyword dictionaries. **Detection order matters** where overlapping keywords exist.

---

## 8) Temporary workspace conventions

Most handlers set `LAMBDA_TMP_DIR` before importing ETL modules so module-level base paths resolve correctly at runtime.

Observed tmp roots:
- Adidas: `/tmp/stibo_workdir_adidas`
- Smiggle: `/tmp/stibo_workdir_smiggle`
- New Balance + Diadora: `/tmp/stibo_workdir`

Typical structure:

```text
/tmp/<workdir>/
  input/
    <file_type>/...downloaded files...
  output/
    xml/...generated xml...
```

---

## 9) Environment variables

Common variables:

- `RAW_BUCKET`
  - Source metadata bucket (and audit log sink in current logger).

- `PROCESSED_BUCKET`
  - Destination bucket for generated STEP XML.

Optional/brand-specific:

- `COMP_CODE`
  - Used as fallback in some brand flows (e.g., New Balance/Diadora comments).

- `ALDO_SBU` (in commented Aldo draft)
  - Default SBU fallback for Aldo flow.

- `AWS_LAMBDA_FUNCTION_NAME`
  - Read by `AuditLogger.flush` for audit metadata.

---

## 10) Error handling and resiliency model

- Router ignores non-metadata paths safely (`statusCode: 200`).
- Missing required files generally causes early return (wait-for-complete-upload model) rather than hard crash.
- File conversion and download failures are captured in audit with status/error details.
- Router always flushes audit in `finally`, even when brand handler throws.
- Global fan-out isolates failures by brand and returns per-brand result map.

---

## 11) Operational notes

1. **Keyword changes are behavior changes**
   - Modifying `FILE_TYPE_KEYWORDS` or order can alter ETL dispatch.

2. **Global file updates trigger all brands**
   - Uploading root-level MDD/naming causes full fan-out refresh.

3. **Aldo handler state**
   - Aldo lambda file is commented; verify deployment artifact before enabling Aldo traffic.

4. **Runtime package requirements**
   - Handlers import `boto3`, `pandas`, `openpyxl`, `pyxlsb` and require these in Lambda/package layer.

5. **Potential conflict risk in shared tmp dir**
   - NB and Diadora both use `/tmp/stibo_workdir`; usually safe per invocation container, but separate names can further reduce warm-container cross-flow risk.

---

## 12) End-to-end execution walkthrough

Example: upload `raw/metadata/new-balance/0888-...-Line List (Footwear)-...xlsx`

1. EventBridge emits S3 Object Created event.
2. `router.lambda_handler` receives event.
3. Router sees non-global key and extracts principal `new-balance`.
4. Router creates `AuditLogger` and calls `new_balance.lambda_function.lambda_handler`.
5. NB handler lists principal files + required globals, classifies by filename, validates required set.
6. Handler downloads files into `/tmp/stibo_workdir/input/...`.
7. Handler converts format if needed (`xlsb/csv -> xlsx`).
8. Handler dispatches to ETL module mapped for `linelist_footwear`.
9. ETL writes XML under `/tmp/stibo_workdir/output/xml/`.
10. Handler uploads XML to `processed/stepxml/new-balance/...` and records uploaded keys.
11. Control returns to router.
12. Router flushes complete audit JSON to `raw/logs/new-balance/<timestamp>.json`.

---

## 13) Local development guidance

Because this code is AWS-event and S3 dependent, local tests typically rely on:
- Stubbed/mocked EventBridge event payloads.
- Mocked S3 client behavior (or sandbox test buckets).
- Setting required env vars before import/invocation:
  - `RAW_BUCKET`
  - `PROCESSED_BUCKET`
  - optional brand-specific defaults

Minimum event shape expected by router/handlers:

```json
{
  "detail": {
    "bucket": {"name": "<bucket>"},
    "object": {"key": "raw/metadata/<brand>/<file>.xlsx"}
  }
}
```

---

## 14) Suggested hardening backlog

1. Add a dependency manifest (`requirements.txt` or `pyproject.toml`) if absent.
2. Add unit tests for:
   - key parsing
   - file-type detection priority
   - required-type validation
   - global fan-out routing
3. Add explicit contract tests for each brand with fixture filenames.
4. Add CI lint + static checks.
5. Clarify/revive Aldo executable handler or remove from router until ready.
6. Add idempotency/correlation IDs in logs for cross-system traceability.

---

## 15) Quick reference

- **Router**: `router.py`
- **Audit schema + upload**: `audit_logger.py`
- **Brand handlers**:
  - `adidas/lambda_function.py`
  - `new_balance/lambda_function.py`
  - `smiggle/lambda_function.py`
  - `diadora/lambda_function.py`
  - `aldo/lambda_function.py` (commented reference)

If you are onboarding, start with `router.py`, then `audit_logger.py`, then one brand handler + its ETL module set.
