# Pazzion — Stibo Mapping Audit
*Template columns on S3 vs. what the Lambda reads and maps to Stibo*

> **Legend**
> - ✅ **Mapped** — Lambda reads this column and maps it to a Stibo attribute
> - ⚠️ **Partial** — Present in template but handled implicitly or via sub-row extraction logic, not direct column access
> - ❌ **NOT Mapped** — Column NEVER read by the Lambda; data is silently dropped

---

## Architecture Note: Dynamic Mapping System

Pazzion uses a **fundamentally different architecture** from New Balance.

Instead of hardcoded `row.get("Column Name")` calls, the Order Form ETL loads a
**"NEW - Brand mapping files Template.xlsx"** file (the `BrandMappingLoader`), reads the
`"Pazzion(Inline)"` sheet, and dynamically iterates over attribute-to-field-name mappings at runtime.

```
Brand Mapping Excel → Col B (AT_xxx) → Col J (Field Name in Brand File)
→ _get_field(row_data, field_name) → LOV lookup or direct text → XML
```

This means the mapping is **not visible in the Python code alone**. The lambda reads:
- `"NEW - Brand mapping files Template.xlsx"` → `Pazzion(Inline)` sheet for the attribute-to-column map
- `MDD.xlsx` for LOV lookups
- The Order Form `.xlsx` itself for row data

Additionally, certain special columns are extracted from **sub-row "Material Description" blocks**
(Heel Height, Available Sizes, Upper material, Division, Category) via `_post_process_colors()`,
because Pazzion's order form uses a stacked/sparse layout rather than flat column headers.

---

## 1. Order Form (`IDM - Spring Summer'26 Order Form.xlsx`)
**Sheets:** `SS'26`, `Styles Selected in Wisma`, `Additional Styles Selected`, `Exclusive Styles`
**Lambda:** `order_form.py` → `build_product_xml()` + `_add_product_values()`

### Sheet Structure (all sheets use same column layout)

| Template Column | Status | Stibo Attribute | How it's read |
|---|---|---|---|
| NO. | ❌ | — | Not read (row index only) |
| PHOTOS | ❌ | — | Not read (image column) |
| *(empty col between PHOTOS and ARTICLE NO.)* | ❌ | — | Skipped (unnamed/merged) |
| ARTICLE NO. | ✅ | KEY_InboundArticle, AT_PrincipalStyleCode | `row_data.get("Article No.") or row_data.get("ARTICLE NO.")` |
| COLOURS | ✅ | AT_PrincipalColorName | Collected from sub-rows via `_post_process_colors()` → `_color` field |
| UNIT PRICE (SGD) | ✅ | AT_FOB | Via brand mapping → field name `"UNIT PRICE\n(SGD)"` or `"Unit Price \n(SGD)"` (case-insensitive, newline-normalized) |
| RETAIL PRICE (SGD) | ✅ | AT_OriginalPrice, AT_CurrentPrice | Via brand mapping → field name `"RETAIL PRICE\n(SGD)"` |
| MATERIAL DESCRIPTION *(label column)* | ⚠️ | Multiple (see below) | Label column used to identify what the next cell contains in sub-rows |
| *(value column — unnamed after MATERIAL DESCRIPTION)* | ⚠️ | Multiple (see below) | Value column, unnamed in header, captured via `_post_process_colors()` |
| PRODUCT HIGHLIGHT | ✅ | AT_ProductHighlight or AT_LongDescriptionEN | Via brand mapping → field name `"PRODUCT HIGHLIGHT"` |
| QUANTITY *(and per-size quantity columns)* | ❌ | — | Not read (ordering quantity, not a Stibo attribute) |
| TOTAL (QTY) | ❌ | — | Not read |
| TOTAL PRICE (SGD) | ❌ | — | Not read |

### Material Description Sub-Row Fields (extracted via `_post_process_colors`)

These are NOT separate columns — they appear as **row label + value pairs** under the MATERIAL DESCRIPTION column:

| Sub-row Label | Status | Stibo Attribute | Notes |
|---|---|---|---|
| `Upper` / `Material` (for bags) | ✅ | AT_Material | Raw value → Material Mapping file → LOV ID |
| `Heel Height` | ✅ | AT_HeelHeight | Raw value → Heel Height Mapping file → LOV ID |
| `Available Sizes` | ✅ | AT_PrincipalSize | Sent as-is (text) |
| `Division` | ✅ | AT_PrincipalMerchandiseHierarchyL1 | Determines PPH parent (Footwear vs Bag) |
| `Category` | ✅ | AT_PrincipalMerchandiseHierarchyL2 | Sent as-is (text) |

### System-Level Attributes (set by code, not from template columns)

These attributes are written with hardcoded or filename-derived values — they do NOT depend on any template column:

| Stibo Attribute | Value Source |
|---|---|
| AT_Brand | Brand code from filename (`PZZ`) |
| AT_CompanyCode | From RNA/source mapping lookup |
| AT_SBU | From RNA/source mapping lookup |
| AT_SAPAge | Hardcoded `"AD"` (Adult) |
| AT_Gender | Hardcoded `"F"` (Female) |
| AT_BYGender | Hardcoded `"F"` (Female) |
| AT_BYAge | Hardcoded `"ADULT"` |
| AT_CountryOrigin | Hardcoded `"CN"` (China) |
| AT_PricingDistributionChannel | Hardcoded `"01"` |
| AT_InboundGenericCode | Computed: `BrandCode + ArticleCode + Color` |
| AT_Season | From filename (e.g. `SS26`) |
| AT_SeasonYear | Derived from season (e.g. `2026`) |
| AT_Country | From filename country code |
| AT_BrandType | From RNA lookup |
| AT_BrandGroup | From RNA lookup |
| AT_BrandStatus | From RNA lookup |
| AT_CountrySize | Hardcoded `"EU"` (footwear only, not sent for bags) |
| AT_RetailPriceCurrency | Derived from country code |

### ❌ Unmapped Template Columns (Order Form)

| Column | Reason |
|---|---|
| NO. | Row index only — no Stibo attribute |
| PHOTOS | Image column — no Stibo attribute via inbound XML |
| QUANTITY (and all per-size qty sub-columns) | Order quantities, not PIM data |
| TOTAL (QTY) | Computed total, not PIM data |
| TOTAL PRICE (SGD) | Computed total, not PIM data |

> [!NOTE]
> The empty/unnamed columns between headers (merged cells in the Excel) are automatically dropped
> by the loader via `df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed:")]`.

### ⚠️ Important: Columns Depend on Brand Mapping File

The exact list of order form columns that map to Stibo attributes is **defined at runtime** by the
`"NEW - Brand mapping files Template.xlsx"` → `Pazzion(Inline)` sheet.
The brand mapping file's `Col J` (Field Name in Brand File) determines which column names are read.

The known mapped columns (from code analysis) are:
- `ARTICLE NO.` / `Article No.` / `Article` → `AT_PrincipalStyleCode`
- `COLOURS` / `Colours` → `AT_PrincipalColorName`
- `UNIT PRICE\n(SGD)` / `Unit Price \n(SGD)` → `AT_FOB`
- `RETAIL PRICE\n(SGD)` / `RETAIL PRICE\n(SGD)` → `AT_OriginalPrice` / `AT_CurrentPrice`
- `PRODUCT HIGHLIGHT` / `Product Highlight` → ecomm description / bullets
- `_heel_height` (sub-row extracted) → `AT_HeelHeight`
- `_available_sizes` (sub-row extracted) → `AT_PrincipalSize`
- `_material_upper` (sub-row extracted) → `AT_Material`
- `_division` (sub-row extracted) → `AT_PrincipalMerchandiseHierarchyL1`
- `_category` (sub-row extracted) → `AT_PrincipalMerchandiseHierarchyL2`

---

## 2. Barcode File (`CI1167 barcode.xlsx`)
**Sheet:** `Sheet`
**Lambda:** `barcode.py` → `group_barcode_rows()`

| Template Column | Status | Stibo Attribute | Notes |
|---|---|---|---|
| Prod Code | ✅ | KEY_InboundArticle (generic key), AT_PrincipalStyleCode | `row.get("Prod Code")` via `loader.find_column()` |
| Colour | ✅ | AT_PrincipalColorName (part of generic key) | `row.get("Colour")` via `loader.find_column()` |
| Size | ✅ | AT_Size / AT_SizeCode (via MDD LOV lookup) | `row.get("Size")` via `loader.find_column()` |
| POS Code | ✅ | AT_PrincipalBarcode / DC_Barcode | `row.get("POS Code")` via `loader.find_column()` |
| SKU Code | ❌ | — | Not explicitly read by mapper |
| Category | ❌ | — | Not read by barcode.py |
| Heel Height | ❌ | — | Not read by barcode.py (only read in order_form.py via sub-row extraction) |
| Retail Price | ❌ | — | Not read by barcode.py |
| Upper Material | ❌ | — | Not read by barcode.py |
| Inner Lining | ❌ | — | Not read by barcode.py |
| Insole Material | ❌ | — | Not read by barcode.py |
| Product Name | ❌ | — | Not read by barcode.py |
| Product Description | ❌ | — | Not read by barcode.py |

> [!NOTE]
> The barcode file's purpose is **exclusively** to provide EAN/barcode data (AT_PrincipalBarcode).
> It is NOT a product data source — all product attributes come from the Order Form.
> Therefore the unmapped columns (SKU Code, Category, Heel Height, Retail Price, Upper Material, etc.)
> are expected to be unused here. They exist in the template to give Pazzion a single file for
> barcode scanning/ordering; only the 4 key columns are needed for the Stibo inbound XML.

**❌ Summary — Barcode file: 9 columns NOT mapped** (by design — barcode file is for EAN data only)
`SKU Code`, `Category`, `Heel Height`, `Retail Price`, `Upper Material`, `Inner Lining`, `Insole Material`, `Product Name`, `Product Description`

---

## Summary Table

| Template | Total Cols | ✅ Mapped | ❌ NOT Mapped | Notes |
|---|---|---|---|---|
| Order Form (all sheets) | ~13 named + sub-rows | 8 (named) + 5 (sub-row) | 5 | Qty/total cols intentionally dropped |
| Barcode | 13 | 4 | 9 | By design — EAN-only file |

---

## ✅ All Critical Stibo Inbound Fields Are Mapped

The following core Stibo attributes are correctly populated:

**From Order Form:**
- `AT_PrincipalStyleCode` — Article No.
- `AT_PrincipalColorName` — Colours (from sub-rows)
- `AT_FOB` — Unit Price (SGD)
- `AT_OriginalPrice` / `AT_CurrentPrice` — Retail Price (SGD)
- `AT_Material` — Upper/Material sub-row → Material Mapping → LOV
- `AT_HeelHeight` — Heel Height sub-row → Heel Height Mapping → LOV
- `AT_PrincipalSize` — Available Sizes sub-row
- `AT_PrincipalMerchandiseHierarchyL1` — Division sub-row
- `AT_PrincipalMerchandiseHierarchyL2` — Category sub-row
- `AT_Season`, `AT_SeasonYear` — From filename
- `AT_Country`, `AT_CountryOrigin`, `AT_CountrySize` — From filename / hardcoded
- `AT_Brand`, `AT_BrandGroup`, `AT_BrandType`, `AT_BrandStatus` — From RNA + MDD
- `AT_Gender`, `AT_SAPAge`, `AT_BYGender`, `AT_BYAge` — Hardcoded defaults (F / AD / ADULT)
- `AT_CompanyCode`, `AT_SBU` — From RNA lookup
- `AT_RetailPriceCurrency` — Derived from country

**From Barcode:**
- `AT_PrincipalBarcode` / `DC_Barcode` — POS Code

---

## ⚠️ Potential Issues / Things to Verify

> [!WARNING]
> **AT_Gender and AT_SAPAge are hardcoded** to `F` (Female) and `AD` (Adult) respectively.
> These do not come from the template — Pazzion does not provide gender data in their order form.
> If the brand ships unisex or kids products, these hardcoded defaults will be incorrect.

> [!WARNING]
> **AT_CountryOrigin is hardcoded** to `CN` (China) for all Pazzion articles.
> The order form does not have a COO column. Verify with the brand that all products are made in China,
> or a COO column needs to be added to the template.

> [!NOTE]
> **The actual column-to-attribute mapping depends on the Brand Mapping file** (`Pazzion(Inline)` sheet).
> If new attributes need to be mapped from the order form, they should be added there — not in the Python code.

> [!NOTE]
> **The barcode file is a separate, independent inbound.** It enriches the same Generic/Variant
> structure built by the Order Form with EAN barcodes. Both files must be processed in the same
> pipeline run to produce complete Stibo XML.
