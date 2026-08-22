# Service Context: map-MAA-outbound-process-trigger-dev

## Environment & Account Info
- **AWS Account ID:** 590184096617
- **Region:** ap-southeast-3 (Jakarta)
- **Environment:** dev
- **Event Bus:** default (Standard)
- **Rule ARN:** arn:aws:events:ap-southeast-3:590184096617:rule/map-MAA-outbound-process-trigger-dev

---

## Trigger & Routing Logic
1. **Source Event:** AWS S3 `Object Created` event emitted via Amazon EventBridge.
2. **Watched Bucket:** `map-MAA-outbound-dev`
3. **Behavior:** Any new file created/uploaded to `map-MAA-outbound-dev` matches this rule.
4. **Target Destination:** AWS Lambda function `map-MAA-outbound-transform-by-dev`
   - **Target ARN:** `arn:aws:lambda:ap-southeast-3:590184096617:function:map-MAA-outbound-transform-by-dev`
   - **Input Type:** Matched event (full, unmodified EventBridge JSON envelope passed directly to Lambda `event` parameter).
   - **Dead Letter Queue (DLQ):** None configured.

---

## Expected Event Data Path for Lambda Handler
When writing or debugging the Lambda handler (`map-MAA-outbound-transform-by-dev`), access the S3 metadata from the incoming event payload using:

- **Bucket Name:** `event["detail"]["bucket"]["name"]` -> `"map-MAA-outbound-dev"`
- **Object Key / File Name:** `event["detail"]["object"]["key"]`
- **File Size:** `event["detail"]["object"]["size"]`
- **Event Time:** `event["time"]`