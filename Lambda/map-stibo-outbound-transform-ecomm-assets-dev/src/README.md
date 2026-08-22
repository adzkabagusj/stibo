# Outbound Ecomm Assets — Resizing Pipeline

AWS Lambda pipeline that downloads product images from STIBO, resizes them
per channel specifications, and delivers ZIP archives to S3.

---

## How It Works

```
S3 upload (product.xml)
  s3://map-stibo-outbound-dev/step-export/ecomm-assets/
        │
        ▼  EventBridge / S3 trigger / manual invoke
  lambda_handler.py
        │
        ▼
  config.py  ──►  downloads Image Requirements.xlsx from S3 (cold-start)
                  parses specs + channel mappings
                  manages OIDC token cache
        │
        ▼
  xml_parser.py  ──►  parses product XML
                       extracts asset IDs + channel IDs
        │
        ▼
  downloader.py  ──►  fetches OIDC bearer token (cached)
                       downloads images from STIBO REST API
        │
        ▼
  image_processor.py  ──►  resizes each image per spec
                             applies margins + background color
                             exports JPEG ≤ 500 KB
        │
        ▼
  zipper.py  ──►  groups output JPEGs by canvas size
                  creates timestamped ZIPs
        │
        ▼
  s3_helper.py  ──►  uploads ZIPs to
                      s3://map-stibo-outbound-dev/transformed/ecomm-assets/
```

---

## Project Structure

```
MAA-Backend/
├── lambda_handler.py        # Lambda entry point
├── requirements.txt         # Python dependencies
└── resize_pipeline/
    ├── __init__.py
    ├── config.py            # Loads specs from S3, manages OIDC token cache
    ├── models.py            # Shared data classes
    ├── xml_parser.py        # Parses STIBO product XML
    ├── downloader.py        # Downloads images from STIBO API
    ├── image_processor.py   # Pillow resize + JPEG export
    ├── zipper.py            # Groups images and creates ZIPs
    ├── pipeline.py          # Orchestrates all steps
    └── s3_helper.py         # S3 read / write helpers
```

---

## S3 Layout

All pipeline operations use a **single bucket**: `map-stibo-outbound-dev`

| Purpose | S3 Path |
|---|---|
| Requirements Excel | `Image Requirements.xlsx` *(root of bucket)* |
| Input product XML | `step-export/ecomm-assets/{filename}.xml` |
| Output ZIPs | `transformed/ecomm-assets/{filename}.zip` |
| Audit logs | `etl-audit-logs/{product_id}/{timestamp}Z.json` |

---

## Requirements File (Excel)

Upload `Image Requirements.xlsx` to the **root** of `map-stibo-outbound-dev`.

The file is a **single sheet**. The parser automatically selects the first sheet
whose name contains `image`, `requirement`, `spec`, or `resize`.
If none match, it falls back to the first sheet in the workbook.

### Sheet Layout

| Row(s) | Column A (label) | Columns D–H (values per spec) |
|---|---|---|
| 1 | *(header — ignored)* | spec names |
| 2–9 | field label | field value per spec |
| 10–13 | *(ignored)* | NW margins: top / bottom / left / right |
| 14–17 | *(ignored)* | AW margins: top / bottom / left / right |
| 18–21 | *(ignored)* | NM + AM margins: top / bottom / left / right *(shared)* |
| 22+ | *(empty)* | channel IDs per spec column — stop at first empty cell |

### Column → Spec Mapping

| Column | Spec Name |
|---|---|
| D | `digital_flagship` |
| E | `converse` |
| F | `footlocker` |
| G | `mapclub` |
| H | `zalora` |

### Rows 2–9: Spec Metadata

Column A label is matched by keyword (case-insensitive):

| Keyword in label | Field | Example value |
|---|---|---|
| `canvas`, `size`, `dimension` | Canvas size | `800 X 800` |
| `hex` | Background color | `#FFFFFF` |
| `dpi`, `resolution` | DPI | `72` |

> Canvas size is parsed from strings like `800 X 800`, `800x800`, or `800 x 800`.
> Background color is read **only** from the row labelled `Color Hex` — the
> `Background` row (which contains text like "White") is ignored.

### Rows 10–21: Margins

| Row | Asset Code | Side |
|---|---|---|
| 10 | NW | top |
| 11 | NW | bottom |
| 12 | NW | left |
| 13 | NW | right |
| 14 | AW | top |
| 15 | AW | bottom |
| 16 | AW | left |
| 17 | AW | right |
| 18 | NM + AM | top |
| 19 | NM + AM | bottom |
| 20 | NM + AM | left |
| 21 | NM + AM | right |

> NM and AM share the same margin rows (18–21).

### Rows 22+: Channel IDs

Each spec column (D–H) lists its channel IDs from row 22 downward.
The parser stops at the first empty cell in each column.

Channel IDs are **normalised** before being stored:
1. Trailing `-NNNN` version suffix stripped — e.g. `ID-Digital Flagship-Aldo-0191` → `ID-Digital Flagship-Aldo`
2. Converted to uppercase — e.g. `ID-Digital Flagship-Aldo` → `ID-DIGITAL FLAGSHIP-ALDO`

The XML parser also uppercases channel IDs before lookup so both sides always match:

```xml
<Value ID="ID-Digital Flagship-Aldo">Digital Flagship</Value>
              ↑
  normalised to ID-DIGITAL FLAGSHIP-ALDO before lookup
```

---

## Asset Codes

The last two characters after the final underscore in an asset name
determine which margin rows are applied:

| Code | Meaning | Margin rows |
|---|---|---|
| `NW` | Non-Apparel / Without Model | 10–13 |
| `AW` | Apparel / Without Model | 14–17 |
| `NM` | Non-Apparel / With Model | 18–21 |
| `AM` | Apparel / With Model | 18–21 |

Example: `ADIHP3837_2_NW` → code is `NW`

> Assets whose name does not end with a recognised code (`NW`, `AW`, `NM`, `AM`)
> are **skipped with a warning** — they do not cause the pipeline to fail.

---

## STIBO Authentication

Authentication uses **OIDC client-credentials flow** — no hardcoded bearer token.

A module-level token cache in `config.py` means the token is reused across warm
Lambda invocations and only refreshed when within `STIBO_TOKEN_LEEWAY` seconds
of expiry (default: 30 s).

```
POST {STIBO_TOKEN_URL}
  grant_type    = client_credentials
  client_id     = {STIBO_CLIENT_ID}
  client_secret = {STIBO_CLIENT_SECRET}

→ { "access_token": "…", "expires_in": 3600 }
```

The token is fetched internally by `downloader.py` — callers do not pass tokens.

---

## Environment Variables

### Required

| Variable | Description |
|---|---|
| `STIBO_BASE_URL` | STIBO host URL e.g. `https://your-stibo-host.com` |
| `STIBO_TOKEN_URL` | OIDC token endpoint URL |
| `STIBO_CLIENT_ID` | OIDC client ID |
| `STIBO_CLIENT_SECRET` | OIDC client secret |

### Optional

| Variable | Default | Description |
|---|---|---|
| `S3_BUCKET` | `map-stibo-outbound-dev` | Single S3 bucket for all operations |
| `STIBO_GRANT_TYPE` | `client_credentials` | OAuth grant type |
| `STIBO_TOKEN_LEEWAY` | `30` | Seconds before expiry to refresh token |
| `STIBO_CONTEXT` | `Context1` | STIBO context query parameter |
| `STIBO_WORKSPACE` | `Approved` | STIBO workspace query parameter |

> If any required variable is missing, the Lambda returns HTTP 500 immediately
> without attempting any processing.

---

## Lambda Settings

| Setting | Value |
|---|---|
| Runtime | Python 3.12 |
| Memory | 1024 MB |
| Timeout | 600 s (10 min) |
| Ephemeral storage (`/tmp`) | 1024 MB |

---

## Supported Trigger Shapes

### 1. EventBridge "Object Created" *(recommended)*

Fires automatically via an EventBridge rule when a product XML is uploaded:

```json
{
  "source": "aws.s3",
  "detail-type": "Object Created",
  "detail": {
    "bucket": { "name": "map-stibo-outbound-dev" },
    "object": { "key": "step-export/ecomm-assets/1304550.xml" }
  }
}
```

### 2. S3 Event Trigger

```json
{
  "Records": [{
    "s3": {
      "bucket": { "name": "map-stibo-outbound-dev" },
      "object": { "key": "step-export/ecomm-assets/1304550.xml" }
    }
  }]
}
```

### 3. Direct / Manual Invocation

```json
{
  "key": "step-export/ecomm-assets/1304550.xml"
}
```

### 4. API Gateway POST

```json
{
  "body": "{\"key\": \"step-export/ecomm-assets/1304550.xml\"}"
}
```

> The bucket is always `S3_BUCKET` from config — it is **not** read from the event payload.

---

## Audit Logs

Every invocation (success **and** failure) writes a structured JSON log to:

```
s3://map-stibo-outbound-dev/etl-audit-logs/{product_id}/{timestamp}Z.json
```

When the product ID cannot be extracted (e.g. event parse failure), the folder
falls back to `ecomm-assets`.

Example log:
```json
{
  "started_at":   "2026-06-10T08:30:00.123456+00:00",
  "finished_at":  "2026-06-10T08:30:12.456789+00:00",
  "duration_s":   12.333,
  "product_id":   "1304550",
  "input_key":    "step-export/ecomm-assets/1304550.xml",
  "status":       "SUCCESS",
  "status_code":  200,
  "zips_created": 2,
  "s3_keys":      ["transformed/ecomm-assets/800x800-2026-06-10_08.30.00.zip"],
  "error":        null
}
```

| `status` value | Meaning |
|---|---|
| `STARTED` | Invocation initialised (should never appear in final log) |
| `SUCCESS` | Pipeline completed, ZIPs uploaded |
| `NO_OUTPUT` | Pipeline ran but produced no ZIPs |
| `ERROR` | Any failure — see `error` field for details |

> Audit log failures are **never re-raised** — a broken audit write will not
> affect the pipeline response.

---

## Output

All output files from a single invocation share the same timestamp.
**One ZIP is created per unique canvas size** — if two specs both produce
800×800 images, they go into the same ZIP.

```
s3://map-stibo-outbound-dev/transformed/ecomm-assets/
├── 800x800-2026-06-09_10.30.00.zip
├── 1000x1000-2026-06-09_10.30.00.zip
└── 1200x1200-2026-06-09_10.30.00.zip
```

Each ZIP contains flat JPEG files (no subdirectory nesting) named:
```
{asset_name}_{spec_name}.jpg
e.g. ADIHP3837_2_NW_digital_flagship.jpg
```

### Success Response
```json
{
  "product_id":   "1304550",
  "input_key":    "step-export/ecomm-assets/1304550.xml",
  "zips_created": 2,
  "s3_keys": [
    "transformed/ecomm-assets/800x800-2026-06-09_10.30.00.zip",
    "transformed/ecomm-assets/1200x1200-2026-06-09_10.30.00.zip"
  ]
}
```

---

## Error Handling

| Situation | Behaviour |
|---|---|
| Required env var missing | HTTP 500 immediately — no processing attempted |
| `S3_BUCKET` not set at cold-start | `RESIZE_SPECS` stays empty — pipeline aborts on first invocation |
| `Image Requirements.xlsx` missing from S3 | `RuntimeError` at cold-start — Lambda fails immediately |
| Excel sheet missing canvas size / DPI / color | `ValueError` — pipeline aborts, no output |
| OIDC token request fails | `RuntimeError` — asset download fails, pipeline aborts |
| STIBO returns non-200 for an asset | Asset skipped with `[ERROR]` log — other assets continue |
| Asset name has no recognised code suffix | Asset skipped with `[WARN]` log — pipeline continues |
| Channel ID not found in `CHANNEL_TO_SPEC` | Asset skipped with `[WARN]` log — pipeline continues |
| All asset downloads fail | Pipeline aborts, returns HTTP 500 |
| No recognised channel IDs in XML | Pipeline aborts, returns HTTP 422 |
| JPEG cannot be compressed to ≤ 500 KB | `RuntimeError` for that job — other jobs continue |
| ZIP upload to S3 fails | That ZIP skipped with `[ERROR]` — other ZIPs still uploaded |
| All ZIP uploads fail | Returns HTTP 500 |
| Audit log write fails | Warning logged — pipeline response unaffected |

---

## Deploy

```bash
# 1. Zip the deployment package
zip -r deployment.zip lambda_handler.py resize_pipeline/ -x "**/__pycache__/*"

# 2. Upload to Lambda via AWS CLI
aws lambda update-function-code \
  --function-name ecomm-resize-pipeline \
  --zip-file fileb://deployment.zip \
  --region ap-southeast-1

# 3. Set environment variables
aws lambda update-function-configuration \
  --function-name ecomm-resize-pipeline \
  --environment Variables="{
    S3_BUCKET=map-stibo-outbound-dev,
    STIBO_BASE_URL=https://your-stibo-host.com,
    STIBO_TOKEN_URL=https://your-oidc-provider.com/token,
    STIBO_CLIENT_ID=your-client-id,
    STIBO_CLIENT_SECRET=your-client-secret
  }" \
  --region ap-southeast-1

# 4. Test manually
aws lambda invoke \
  --function-name ecomm-resize-pipeline \
  --payload '{"key":"step-export/ecomm-assets/1304550.xml"}' \
  --cli-binary-format raw-in-base64-out \
  response.json && cat response.json
```

---

## Dependencies

| Package | Purpose |
|---|---|
| `Pillow` | Image resizing + JPEG export |
| `lxml` | Parse product XML |
| `boto3` | S3 read / write |
| `openpyxl` | Parse requirements Excel |

> HTTP requests (OIDC token fetch + image download) use Python's built-in
> `urllib.request` — no `requests` library required.