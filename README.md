# Mitra Active Asia (MAA) Portal

A serverless AWS data pipeline and web interface for onboarding footwear article data into Stibo Systems STEP Master Data Management (MDM).

The MAA Portal allows brand users across various footwear lines to upload article creation datasets (inline and licensed files). The backend validates and transforms brand-specific spreadsheets into standardized payloads required by Stibo STEP.

---

## Architecture & Data Flow

```
[Brand User]
     │
     ▼ (Uploads Inline/Licensed Data)
[S3 Web Portal / Bucket]
     │
     ▼ (S3 Event Trigger)
[Amazon EventBridge]
     │
     ▼ (Routes Event Rule)
[AWS Lambda: Validation & Transformation]
     │ (Applies Brand Specs & template-map.json)
     ▼
[Stibo Systems STEP MDM]

```

1. **Upload:** Users access the portal (`index.html`) hosted on S3 and upload inline/licensed article files.
2. **Trigger:** S3 bucket events trigger Amazon EventBridge rule definitions.
3. **Process:** EventBridge invokes AWS Lambda functions to validate input structure against brand-specific schemas.
4. **Ingest:** Lambda formats and dispatches the payload to the Stibo STEP inbound API/endpoint.


---

## Key Components

### 1. Template Registry (`S3/`)

* **`template-map.json`**: Acts as the central routing index. It maps incoming brand identifiers to their respective validation schemas and template requirements.
* **`Brand Input files & Naming convention.xlsx`**: The master specification sheet. Any change to brand file formats or naming rules originates here before updating JSON configs.

### 2. Event Orchestration (`EventBridge/`)

* Contains infrastructure-as-code and JSON definitions for EventBridge rules.
* Filters S3 upload events and routes payloads to target Lambda handlers based on file paths and prefixes.

### 3. Processing Engine (`Lambda/`)

* **`map-stibo-inbound-validate-transform-dev/src`**: Houses brand-specific validation, column mapping, and formatting business logic.
* Normalizes disparate brand inputs into Stibo STEP XML/JSON inbound specifications.

---

## Data Engineer Maintenance Guide

### Adding or Updating a Brand Pipeline

1. **Review Specification:** Consult `Brand Input files & Naming convention.xlsx` for column schemas and allowed values.
2. **Update Template Mapping:** Add or modify the brand entry in `S3/maa-web-portal-files/templates/template-map.json`.
3. **Implement Transformation Logic:** Update or add brand parsers under `Lambda/map-stibo-inbound-validate-transform-dev/src/` to handle schema transformations for the target brand. Also ensures other Lambdas are setup necessarily (if needed)
4. **Verify Routing:** Ensure the EventBridge rule in `EventBridge/` correctly captures the brand upload prefix.