# Adds UAT test row 8 to the (uat) APP ACC Preline workbook.
# Run ONLY after closing the workbook in Excel (it is file-locked otherwise).
# Idempotent: skips if AT71C320_BLK is already present.
# Run from the repo root:  python scratch/add_uat_row8.py
from copy import copy
from datetime import datetime
from pathlib import Path
import sys

import openpyxl

REPO = Path(__file__).resolve().parent.parent
TARGET = REPO / "ContextForAI" / "New Balance" / "Inline" / \
    "(uat) S227 APP ACC Preline Linelist - SEA.xlsm"

try:
    with open(TARGET, "r+b"):
        pass
except PermissionError:
    sys.exit(f"'{TARGET.name}' is open in Excel — close it and re-run.")

wb = openpyxl.load_workbook(TARGET, keep_vba=True)
ws = wb["APPACC"]
header_row, template_row = 4, 5
hdr = {str(ws.cell(row=header_row, column=c).value).replace("\n", " ").strip(): c
       for c in range(1, ws.max_column + 1) if ws.cell(row=header_row, column=c).value}

existing = {str(ws.cell(row=r, column=hdr["Item Number"]).value).strip()
            for r in range(header_row + 1, ws.max_row + 1)}
if "AT71C320_BLK" in existing:
    print("Row AT71C320_BLK already present — nothing to do.")
    sys.exit(0)

tr = ws.max_row + 1
for c in range(1, ws.max_column + 1):
    ws.cell(row=tr, column=c)._style = copy(ws.cell(row=template_row, column=c)._style)
    ws.cell(row=tr, column=c).value = None
for c in range(1, ws.max_column + 1):
    name = str(ws.cell(row=header_row, column=c).value or "")
    if "what 3 words" in name.lower():
        ws.cell(row=tr, column=c).value = None

vals = {
    "Item - CarryOver/New": "New",
    "Product Number": "AT71C320",
    "Item Number": "AT71C320_BLK",
    "Date Canceled": 0,
    "APAC DA/DROP": "APAC DA",
    "SMS": "Yes",
    "GBU": "Lifestyle",
    "LPA Category": "Total Lifestyle",
    "Product Display Name": "Sport Cotton Jersey Tee",
    "Category": "Total Lifestyle",
    "Line Plan Business": "Lifestyle",
    "Product Line": "Apparel",
    "Silhouette": "Tee",
    "Detailed Silhouette": "Tee",
    "Product Collection": "Sport Style",
    "Gender Duty - Non-Footwear": "Womens",
    "Color Code - Non-Footwear": "BLACK (BLK)",
    "Size Profile": "Womens",
    "Mandatory Global Brand Event": "No",
    "Global Item Intro Date": datetime(2027, 7, 1),
    "SEA Intro Date": datetime(2027, 7, 1),
    "Intro Date - APAC": datetime(2027, 7, 1),
    "Global Item Phaseout Date": datetime(2027, 8, 28),
    "Phaseout Date - APAC": datetime(2027, 8, 28),
    "Product Fit": "Standard",
    "Price Segment": "Good",
    "Global Retail Price (MSRP)": 25,
    "Global Wholesale Price": 12.5,
    "Global Size Offering": "XS,S,M,L,XL",
    "Target Principal Material": "Cotton Jersey",
}
for k, v in vals.items():
    ws.cell(row=tr, column=hdr[k]).value = v
wb.save(TARGET)
wb.close()
print(f"Row {tr} (AT71C320_BLK, Apparel, Target Principal Material='Cotton Jersey') written.")
