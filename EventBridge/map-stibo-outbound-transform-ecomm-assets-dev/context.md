# Service Context: map-stibo-outbound-transform-ecomm-assets-dev

## Environment & Account Info
- **AWS Account ID:** 590184096617
- **Region:** ap-southeast-3 (Jakarta)
- **Environment:** dev
- **Event Bus:** default (Standard)
- **Rule ARN:** arn:aws:events:ap-southeast-3:590184096617:rule/map-stibo-outbound-transform-ecomm-assets-dev

---

## Trigger & Routing Logic
1. **Source Event:** AWS S3 `Object Created` event.
2. **Watched Bucket:** `map-stibo-outbound-dev`
3. **Prefix Filter:** `step-export/ecomm-assets/` (Only objects created inside the `step-export/ecomm-assets/` folder will trigger this rule).
4. **Target Destination:** AWS Lambda function `map-stibo-outbound-transform-ecomm-assets-dev`
   - **Target ARN:** `arn:aws:lambda:ap-southeast-3:590184096617:function:map-stibo-outbound-transform-ecomm-assets-dev`
   - **Input Type:** Matched event (full, unmodified EventBridge JSON envelope passed directly to Lambda `event` parameter).
   - **Dead Letter Queue (DLQ):** None configured.

---

## Expected Event Data Path for Lambda Handler
When writing or debugging the Lambda handler (`map-stibo-outbound-transform-ecomm-assets-dev`), access the S3 metadata using:

- **Bucket Name:** `event["detail"]["bucket"]["name"]` -> `"map-stibo-outbound-dev"`
- **Object Key / Path:** `event["detail"]["object"]["key"]` -> e.g., `"step-export/ecomm-assets/asset_image_highres.jpg"`
- **File Size:** `event["detail"]["object"]["size"]`
- **Event Time:** `event["time"]`