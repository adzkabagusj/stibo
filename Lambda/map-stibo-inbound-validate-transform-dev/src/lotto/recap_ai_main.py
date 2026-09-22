"""
recap_ai_main.py — LOTTO Licensed "recap sample AI ingestion"
=============================================================
Same workbook layout and mapping as recap_main (1st-ingestion recap with the
product picture placed *in* the Image cell), plus Bedrock attribute enrichment:
10 LOV attributes and 10 AI translation attributes per article, produced by the
lotto-bedrock-enrichment Lambda in the Bedrock sandbox account.

File type : "recap_ai" — file name contains "AI-Ingestion"
            (e.g. RECAP_SAMPLE_DEVELOPMENT-LOT-SS27-Footwear_AI-Ingestion.xlsx)
Input     : input/recap_ai/  (kept apart from input/recap/ so a regular recap
            downloaded alongside is never processed with enrichment)
Output    : the same STEP XML as recap_main, named after the workbook.

See lotto/bedrock_enrichment.py and INTEGRATION_BEDROCK.md for configuration.
"""
import lotto.recap_main as recap_main

RECAP_AI_DIR = recap_main.INPUT_DIR / "recap_ai"
RECAP_AI_DIR.mkdir(parents=True, exist_ok=True)


def run(args, auditor=None):
    return recap_main.run(args, auditor=auditor, recap_dir=RECAP_AI_DIR, ai_enrichment=True)
