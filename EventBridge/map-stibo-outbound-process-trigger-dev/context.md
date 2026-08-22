# Service Context: map-stibo-outbound-process-trigger-dev

## Environment & Account Info
- **AWS Account ID:** 590184096617
- **Region:** ap-southeast-3 (Jakarta)
- **Environment:** dev
- **Event Bus:** default (Standard)
- **Rule ARN:** arn:aws:events:ap-southeast-3:590184096617:rule/map-stibo-outbound-process-trigger-dev

---

## Architecture Pattern: Fan-Out (1 to 5)
A single S3 upload triggers **5 parallel downstream Lambda functions** simultaneously with the exact same event envelope.

1. **Source Event:** AWS S3 `Object Created` event.
2. **Watched Bucket:** `map-stibo-outbound-dev`
3. **Prefix Filter:** `step-export/` (Only objects created inside `step-export/` trigger the rule).
4. **Target Destinations (5 Lambdas):**
   - `map-stibo-outbound-transform-pos-dev`
   - `map-stibo-outbound-transform-by-dev`
   - `map-stibo-outbound-transform-sql-dev`
   - `map-stibo-outbound-transform-sap-dev`
   - `map-stibo-outbound-transform-gtech-dev`
5. **Input Type:** Matched event (full, unmodified EventBridge JSON payload sent to all 5 functions).
6. **Dead Letter Queue (DLQ):** None configured across any target.

---

## Expected Event Data Path for Lambda Handlers
All 5 Lambda functions can parse the incoming payload using the exact same schema:

- **Bucket Name:** `event["detail"]["bucket"]["name"]` -> `"map-stibo-outbound-dev"`
- **Object Key / Path:** `event["detail"]["object"]["key"]` -> e.g., `"step-export/stibo_outbound_data.xml"`
- **File Size:** `event["detail"]["object"]["size"]`
- **Event Timestamp:** `event["time"]`