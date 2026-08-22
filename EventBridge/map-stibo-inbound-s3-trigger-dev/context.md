# Service Context: map-stibo-inbound-s3-trigger-dev

## Environment & Account Info
- **AWS Account ID:** 590184096617
- **Region:** ap-southeast-3 (Jakarta)
- **Environment:** dev
- **Event Bus:** default (Standard)
- **Rule ARN:** arn:aws:events:ap-southeast-3:590184096617:rule/map-stibo-inbound-s3-trigger-dev

---

## Trigger & Routing Logic
1. **Source Event:** AWS S3 `Object Created` event.
2. **Watched Bucket:** `map-stibo-inbound-raw-dev`
3. **Prefix Filter:** `raw/metadata/` (Only objects created inside this folder/prefix will trigger the rule).
4. **Target Destination:** AWS Lambda function `map-stibo-inbound-validate-transform-dev`
   - **Target ARN:** `arn:aws:lambda:ap-southeast-3:590184096617:function:map-stibo-inbound-validate-transform-dev`
   - **Input Type:** Matched event (full, unmodified EventBridge JSON envelope passed directly to Lambda `event` parameter).
   - **Dead Letter Queue (DLQ):** None configured.

---

## Expected Event Data Path for Lambda Handler
When writing or debugging the Lambda handler (`map-stibo-inbound-validate-transform-dev`), access the S3 metadata using:

- **Bucket Name:** `event["detail"]["bucket"]["name"]` -> `"map-stibo-inbound-raw-dev"`
- **Object Key / Path:** `event["detail"]["object"]["key"]` -> e.g., `"raw/metadata/example_file.json"`
- **File Size:** `event["detail"]["object"]["size"]`
- **Event Time:** `event["time"]`