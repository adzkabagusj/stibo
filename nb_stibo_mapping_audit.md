# New Balance — Stibo Mapping Audit
*Template columns on S3 vs. what the Lambda reads and maps to Stibo*

> **Legend**
> - ✅ **Mapped** — Lambda reads this column and maps it to a Stibo attribute
> - ⚠️ **Partial** — Present in template but handled via fallback or implicit logic (not by exact column name)
> - ❌ **NOT Mapped** — Column NEVER read by the Lambda; data is silently dropped

---

## 1. Inline — Line List (Footwear)
**Template:** `S3/.../inline/line_list_footwear.xlsx` | Sheet: **All Sourced**
**Lambda:** `inline_main_footwear.py` → `map_sku()`

| Template Column | Status | Stibo Attribute | Notes |
|---|---|---|---|
| Primary Image | ❌ | — | Not read |
| Item Number | ✅ | AT_PrincipalStyleCode | `row.get("Item Number")` |
| Product Number | ✅ | AT_SAPStyleCode | `row.get("Product Number")` |
| Size Profile | ✅ | AT_Gender, AT_SAPAge | `row.get("Size Profile")` |
| Intro Period | ✅ | (metadata) | `row.get("Intro Period")` |
| Product Line | ✅ | AT_SAPProductDivision | `row.get("Product Line")` |
| Line Plan Business | ✅ | AT_SAPProductGroup, AT_SportsCategoryEN | `row.get("Line Plan Business")` |
| Category | ✅ | AT_ProductHierarchyL2 | `row.get("Category")` |
| Segment | ✅ | AT_MerchandiseCategory | `row.get("Segment")` |
| Intro Month / CarryOver | ❌ | — | Not read |
| Phaseout Date - Region | ❌ | — | Not read |
| DfLT Style | ❌ | — | Not read |
| Early Buy Style | ❌ | — | Not read |
| Product Status | ❌ | — | Not used for filtering |
| Include? | ❌ | — | Not used for filtering |
| Intro Date - Region | ❌ | — | `Intro Period` used instead |
| Phraseout Date - Region | ❌ | — | Not read |
| Sizes - Region | ✅ | AT_Size / AT_PrincipalSize | `row.get("Sizes - Region")` |
| Minimum Order Quantity | ❌ | — | Not read |
| Product Name | ✅ | AT_PrincipalStyleDescription | `row.get("Product Name")` |
| Sourcing Configuration | ❌ | — | Not read |
| Factory Name | ✅ | AT_VendorName | `row.get("Factory Name")` |
| Country of Origin | ✅ | AT_CountryOrigin | `row.get("Country of Origin")` |
| NBIL Price | ✅ | AT_FOB | `row.get("NBIL Price")` |
| NBIL Price Date | ❌ | — | Not read |
| NRF Color | ✅ | AT_PrincipalColorCode, AT_Color | `row.get("NRF Color")` |
| Item Code | ❌ | — | Not read |
| Retail Price | ✅ | AT_OriginalPrice, AT_CurrentPrice | `row.get("Retail Price")` |
| Product Sequence ID | ❌ | — | Not read |
| Date Added | ❌ | — | Not read |

**❌ Unmapped (15):** `Primary Image`, `Intro Month / CarryOver`, `Phaseout Date - Region`, `DfLT Style`, `Early Buy Style`, `Product Status`, `Include?`, `Intro Date - Region`, `Phraseout Date - Region`, `Minimum Order Quantity`, `Sourcing Configuration`, `NBIL Price Date`, `Item Code`, `Product Sequence ID`, `Date Added`

---

## 2. Inline — Line List (Apparel & Accessories)
**Template:** `S3/.../inline/line_list_apparel_accessories.xlsx` | Sheet: **Pricelist**
**Lambda:** `inline_main_accessories.py` → `map_sku()`

| Template Column | Status | Stibo Attribute | Notes |
|---|---|---|---|
| Item Number | ✅ | AT_PrincipalStyleCode | `row.get("Item Number")` |
| Product Number | ✅ | AT_SAPStyleCode | `row.get("Product Number")` |
| Color Code | ✅ | AT_PrincipalColorCode, AT_ColorCode | `row.get("Color Code")` |
| GBU | ✅ | (stored in dict) | `row.get("GBU")` |
| Line Plan Business | ✅ | AT_SAPProductGroup, AT_SportsCategoryEN | `row.get("Line Plan Business")` |
| Category | ✅ | AT_ProductHierarchyL2 | `row.get("Category")` |
| Channel Type | ✅ | AT_BYArticleType | `row.get("Channel Type")` |
| LPA Category | ✅ | AT_LPACategory | `row.get("LPA Category")` |
| Merchandise Category | ✅ | AT_MerchandiseCategory | `row.get("Merchandise Category")` |
| SMU Type | ✅ | AT_SMUType | `row.get("SMU Type")` |
| SMU Type Detail | ❌ | — | Not read |
| Size Profile | ✅ | AT_Gender, AT_SAPAge | `row.get("Size Profile")` |
| Global Intro Date | ✅ | AT_LaunchingDate (fallback) | `row.get("Global Intro Date")` |
| Region Intro Date | ✅ | AT_LaunchingDate (preferred) | `row.get("Region Intro Date")` |
| Phraseout Date | ✅ | AT_PhaseOutDate | `row.get("Phraseout Date")` |
| Product Display Name | ✅ | AT_PrincipalStyleDescription | `row.get("Product Display Name")` |
| Item - CarryOver/New | ✅ | AT_NatureOfArticle | `row.get("Item - CarryOver/New")` |
| Color Name | ✅ | AT_PrincipalColorName | `row.get("Color Name")` |
| Color Family | ✅ | AT_PrincipalColorDescription | `row.get("Color Family")` |
| Dev Site | ❌ | — | Not read |
| Product Line | ✅ | AT_SAPProductDivision | `row.get("Product Line")` |
| Collection | ✅ | AT_Collection1 | `row.get("Collection")` |
| Silhouette | ✅ | AT_SAPProductCategory, AT_Silhouette | `row.get("Silhouette")` |
| Sizes | ✅ | AT_Size / AT_PrincipalSize | `row.get("Sizes")` |
| Seasonal Strategic Priority | ❌ | — | Not read |
| Buy | ❌ | — | Not read |
| Early Buy Style List | ❌ | — | Not read |
| Fabric Platform Style | ❌ | — | Not read |
| Earliest X-Factory Date | ❌ | — | Not read |
| Fit Version | ✅ | AT_CountrySize | `row.get("Fit Version")` |
| Technologies | ✅ | AT_TechnologyUsed | `row.get("Technologies")` |
| Primary Fabric Content | ✅ | AT_Content, AT_Material | `row.get("Primary Fabric Content")` |
| Secondary Fabric Content | ❌ | — | Not read |
| Global Retail Price | ✅ | AT_OriginalPrice, AT_CurrentPrice | `row.get("Global Retail Price")` |
| NBIL Price | ✅ | AT_FOB | `row.get("NBIL Price")` |
| Supplier | ✅ | AT_VendorName | `row.get("Supplier")` |
| Production Address Name | ❌ | — | Not read |
| COO | ✅ | AT_CountryOrigin | `row.get("COO")` |
| MPQ Global (Product) | ❌ | — | Not read |
| MPQ Global (Item) | ❌ | — | Not read |
| Market MOQ / Style | ❌ | — | Not read (has newline in col name) |
| Market MOQ / Color | ❌ | — | Not read (has newline in col name) |
| Date Added | ❌ | — | Not read |

**❌ Unmapped (14):** `SMU Type Detail`, `Dev Site`, `Seasonal Strategic Priority`, `Buy`, `Early Buy Style List`, `Fabric Platform Style`, `Earliest X-Factory Date`, `Secondary Fabric Content`, `Production Address Name`, `MPQ Global (Product)`, `MPQ Global (Item)`, `Market MOQ / Style`, `Market MOQ / Color`, `Date Added`

---

## 3. Inline — Ecommerce File
**Template:** `S3/.../inline/ecommerce_file.xlsx` | Sheet: **Item**
**Lambda:** `inline_ecommerce_main.py` → `map_article_nb_ftw()`

> **Note:** Ecommerce columns are already **Stibo-standard field names** (exported from Stibo, re-ingested). `sys_*` fields are internal system columns not required by the inbound XML.

| Template Column | Status | Stibo Attribute | Notes |
|---|---|---|---|
| sys_id | ❌ | — | System field, not needed |
| sys_entitytype | ❌ | — | System field |
| sys_fieldset | ❌ | — | System field |
| sys_segmentID | ❌ | — | System field |
| ProductGender | ✅ | AT_Gender (fallback) | `row.get("ProductGender")` |
| ProductDisplayName | ✅ | AT_PrincipalStyleDescription | `row.get("ProductDisplayName")` |
| ProductModelRootName | ❌ | — | Not read |
| ProductPredecessor | ❌ | — | Not read |
| ProductBrand | ❌ | — | Derived from filename |
| ProductType | ❌ | — | Not read |
| ProductParentBundleMaster | ❌ | — | Not read |
| ProductShoeStyle | ❌ | — | Not read |
| ProductAPACPodValidationErrors | ❌ | — | Not read |
| ProductInspirationDesign | ❌ | — | Not read |
| ProductPrimaryMaterial | ✅ | AT_Content / AT_Material | `row.get("ProductPrimaryMaterial")` |
| ProductMaterialPercentages | ✅ | AT_MaterialPercentages | `row.get("ProductMaterialPercentages")` |
| ProductBullet1_en … ProductBullet10_en | ✅ | AT_TechnologyUsed (joined) | via `_join_bullets()` |
| ProductDTCeCommDescription_en | ✅ | AT_LongDescriptionEN | `row.get("ProductDTCeCommDescription_en")` |
| ProductSizeFitBullet1_en … 7_en | ✅ | AT_SizeFitDescription | via `_join_bullets()` |
| ProductMaterialBullet1_en … 5_en | ✅ | AT_MaterialDescription | via `_join_bullets()` |
| ProductCopywritingEnrichmentStatus | ❌ | — | Not read |
| ProductDTCeCommMaster | ❌ | — | Not read |
| ProductECommProductGender | ❌ | — | Not read |
| ItemStyleNumber | ✅ | KEY_InboundArticle | `row.get("ItemStyleNumber")` |
| ItemParentProductNumber | ✅ | Parent identifier | `row.get("ItemParentProductNumber")` |
| ItemGender | ✅ | AT_Gender, AT_SAPAge | `row.get("ItemGender")` |
| ItemProductManager | ❌ | — | Not read |
| ItemProductOwner | ❌ | — | Not read |
| ItemUniqueIdentifiers | ✅ | Variant EAN fallback | `row.get("ItemUniqueIdentifiers")` |
| ItemWeight | ❌ | — | Not read |
| ItemPLMUniqueIdentifier | ❌ | — | Not read |
| ItemUpcSyndication | ✅ | Variant UPC/EAN/Size | `row.get("ItemUpcSyndication")` (JSON parsed) |
| ItemStockingPolicy | ❌ | — | Not read |
| ItemPimItemType | ❌ | — | Not read |
| ItemAccountsStyleToBeSoldIn | ❌ | — | Not read |
| ItemManualLinksEnabled | ❌ | — | Not read |
| ItemCustomOrderingAngleCode | ❌ | — | Not read |
| ItemEnrichmentStatus | ❌ | — | Not read |
| ItemPodValidationErrors | ❌ | — | Not read |
| ItemMarketPlaceErrors | ❌ | — | Not read |
| ItemStatus | ✅ | Filter only (Inactive dropped) | `df["ItemStatus"]` |
| ItemIntroDate | ✅ | AT_LaunchingDate (fallback) | `row.get("ItemIntroDate")` |
| ItemEMEAInLineIntroDate | ❌ | — | Not read |
| ItemLAInLineIntroDate | ❌ | — | Not read |
| ItemGCInLineIntroDate | ❌ | — | Not read |
| ItemAPACInLineIntroDate | ✅ | AT_LaunchingDate (preferred) | `row.get("ItemAPACInLineIntroDate")` |
| ItemNBJInLineIntroDate | ❌ | — | Not read |
| ItemECommIntrolDate | ❌ | — | Not read |
| ItemSeasons | ✅ | AT_Season (parsed) | `row.get("ItemSeasons")` |
| ItemLatestSeason | ✅ | AT_Season (preferred) | `row.get("ItemLatestSeason")` |
| ItemMerchClass | ❌ | — | Not read |
| ItemSMUType | ❌ | — | Not read |
| ItemPhaseoutDate | ❌ | — | Not read |
| ItemEMEAInLinePhaseoutDate | ❌ | — | Not read |
| ItemLAInLinePhaseoutDate | ❌ | — | Not read |
| ItemGCInLinePhaseoutDate | ❌ | — | Not read |
| ItemAPACInLinePhaseoutDate | ❌ | — | Not read |
| ItemNBJInLinePhaseoutDate | ❌ | — | Not read |
| ItemAssortmentforDTCeComm* (x14) | ❌ | — | Not read — all 14 country assortment flags |
| ItemElastic | ❌ | — | Not read |
| ItemColor | ✅ | AT_PrincipalColorCode, AT_Color | `row.get("ItemColor")` |
| ItemFirstColor_en | ✅ | AT_PrincipalColorName | `row.get("ItemFirstColor_en")` |
| ItemFirstGenericColor | ✅ | AT_PrincipalColorDescription | `row.get("ItemFirstGenericColor")` |
| ItemSecondColor_en | ❌ | — | Not read |
| ItemSecondGenericColor | ❌ | — | Not read |
| ItemThirdColor_en | ❌ | — | Not read |
| ItemThirdGenericColor | ❌ | — | Not read |
| ItemFourthColor_en | ❌ | — | Not read |
| ItemFourthGenericColor | ❌ | — | Not read |
| ItemDTCeCommJNBORetailPrice | ✅ | AT_OriginalPrice (secondary RRP) | `row.get("ItemDTCeCommJNBORetailPrice")` |
| ItemDTCeCommJNBOMarkdownPrice | ❌ | — | Not read |
| ItemDTCeCommNorwayRetailPrice | ❌ | — | Not read |
| ItemDTTeCommPrice | ✅ | AT_CurrentPrice (primary RRP) | `row.get("ItemDTTeCommPrice")` |
| ItemDTCeCommNorwayMarkdownPrice | ❌ | — | Not read |
| ItemDTTWholesalePrice | ❌ | — | Not read |
| ItemDTCeCommJapanRetailPrice | ❌ | — | Not read |
| ItemDTCeCommJapanMarkdownPrice | ❌ | — | Not read |

**❌ Key data fields NOT mapped (not system fields):** `ProductModelRootName`, `ProductPredecessor`, `ProductType`, `ProductShoeStyle`, `ProductInspirationDesign`, `ItemSecondColor_en/Family`, `ItemThirdColor_en/Family`, `ItemFourthColor_en/Family`, `ItemPhaseoutDate`, `ItemSMUType`, `ItemMerchClass`, `ItemWeight`, all region phaseout dates, all 14 DTC assortment columns, all markdown price columns

---

## 4. Inline — EAN Source
**Template:** `S3/.../inline/New_balance_EAN.xlsx` | Sheet: **UPC-Report-D365**
**Lambda:** `inline_ean_source.py` → `map_ean_row()`

| Template Column | Status | Stibo Attribute | Notes |
|---|---|---|---|
| ITEM_NUMBER | ✅ | AT_PrincipalStyleCode (generic key) | `row.get("ITEM_NUMBER")` |
| SKU | ✅ | (variant identity) | `row.get("SKU")` |
| SKU UPC | ⚠️ | — | Not read directly; `UPC#` is used instead |
| UK Size | ✅ | AT_UKSize | `row.get("UK Size")` |
| UPC# | ✅ | AT_PrincipalBarcode / DC_Barcode | `row.get("UPC#")` |
| Width Description | ❌ | — | Not read |
| SKU Size | ✅ | AT_SizeCode + AT_Size (via LOV) | `row.get("SKU Size")` |
| WIDTH | ⚠️ | is_footwear flag only | `"Width" in row` — presence check only |
| SKU Color Code | ✅ | AT_ColorCode, AT_Color, AT_MAAColor | `row.get("SKU Color Code")` |

**❌ Unmapped (2):** `Width Description`, `SKU UPC` (not directly read by name)

---

## 5. Licensed — Line List
**Template:** `S3/.../licensed/line_list.xlsx` | Sheet: **Pricelist**
**Lambda:** `licensed_linelist_main.py` → `map_article_nb_licensed()`

| Template Column | Status | Stibo Attribute | Notes |
|---|---|---|---|
| Item Number | ✅ | AT_PrincipalStyleCode | `row.get("Item Number")` |
| Product Number | ✅ | AT_SAPStyleCode | `row.get("Product Number")` |
| Color Code | ✅ | AT_PrincipalColorCode | `row.get("Color Code")` |
| GBU | ✅ | (stored) | `row.get("GBU")` |
| Line Plan Business | ✅ | AT_SAPProductGroup, AT_SportsCategoryEN | `row.get("Line Plan Business")` |
| Category | ✅ | AT_ProductHierarchyL2 | `row.get("Category")` |
| Channel Type | ⚠️ | Filter only | Used to filter Licensed rows only |
| LPA Category | ✅ | AT_LPACategory | `row.get("LPA Category")` |
| Merchandise Category | ✅ | AT_MerchandiseCategory | `row.get("Merchandise Category")` |
| SMU Type | ✅ | AT_SMUType | `row.get("SMU Type")` |
| SMU Type Detail | ❌ | — | Not read |
| Size Profile | ✅ | AT_Gender, AT_SAPAge | `row.get("Size Profile")` |
| Global Intro Date | ✅ | AT_LaunchingDate (fallback) | `row.get("Global Intro Date")` |
| Region Intro Date | ✅ | AT_LaunchingDate (preferred) | `row.get("Region Intro Date")` |
| Phraseout Date | ✅ | AT_PhaseOutDate | `row.get("Phraseout Date")` |
| Product Display Name | ✅ | AT_PrincipalStyleDescription | `row.get("Product Display Name")` |
| Item - CarryOver/New | ✅ | AT_NatureOfArticle | `row.get("Item - CarryOver/New")` |
| Color Name | ✅ | AT_PrincipalColorName | `row.get("Color Name")` |
| Color Family | ✅ | AT_PrincipalColorDescription | `row.get("Color Family")` |
| Dev Site | ❌ | — | Not read |
| Product Line | ✅ | AT_SAPProductDivision | `row.get("Product Line")` |
| Collection | ✅ | AT_Collection1 | `row.get("Collection")` |
| Silhouette | ✅ | AT_SAPProductCategory | `row.get("Silhouette")` |
| Sizes | ✅ | AT_Size / AT_PrincipalSize | `row.get("Sizes")` |
| Seasonal Strategic Priority | ❌ | — | Not read |
| Buy | ❌ | — | Not read |
| Early Buy Style List | ❌ | — | Not read |
| Fabric Platform Style | ❌ | — | Not read |
| Earliest X-Factory Date | ❌ | — | Not read |
| Fit Version | ✅ | AT_CountrySize | `row.get("Fit Version")` |
| Technologies | ✅ | AT_TechnologyUsed | `row.get("Technologies")` |
| Primary Fabric Content | 🐛 | AT_Content, AT_Material | `row.get("Primary Fabric Content ")` — **TRAILING SPACE BUG** |
| Secondary Fabric Content | ❌ | — | Not read |
| Global Retail Price | ✅ | AT_OriginalPrice | `row.get("Global Retail Price")` |
| NBIL Price | ✅ | AT_FOB | `row.get("NBIL Price")` |
| Supplier | ✅ | AT_VendorName | `row.get("Supplier")` |
| Production Address Name | ❌ | — | Not read |
| COO | ✅ | AT_CountryOrigin | `row.get("COO")` |
| MPQ Global (Product) | ❌ | — | Not read |
| MPQ Global (Item) | ❌ | — | Not read |
| Market MOQ / Style | ❌ | — | Not read |
| Market MOQ / Color | ❌ | — | Not read |
| Date Added | ❌ | — | Not read |

> [!CAUTION]
> **BUG — Line 655 of [`licensed_linelist_main.py`](file:///c:/Users/adzka.juniarta/Adzka%20Codes/3%20-%20AWS%20%26%20Stibo%20Development/aws-stibo/Lambda/map-stibo-inbound-validate-transform-dev/src/new_balance/licensed_linelist_main.py#L655-L655):**
> ```python
> fabric_primary = _s(row.get("Primary Fabric Content "))   # ← trailing space!
> ```
> The template column is `"Primary Fabric Content"` (no trailing space). This key will never match, so `fabric_primary` is always `None` — **fabric content is silently dropped** for all Licensed Line List articles.

**❌ Unmapped (14):** `SMU Type Detail`, `Dev Site`, `Seasonal Strategic Priority`, `Buy`, `Early Buy Style List`, `Fabric Platform Style`, `Earliest X-Factory Date`, `Secondary Fabric Content`, `Production Address Name`, `MPQ Global (Product)`, `MPQ Global (Item)`, `Market MOQ / Style`, `Market MOQ / Color`, `Date Added`

---

## 6. Licensed — Ecommerce File
**Template:** `S3/.../licensed/ecommerce_file.xlsx` | Sheet: **Item**
**Lambda:** `licensed_ecommerce_main.py` → `map_article_nb_ecomm()`

Identical column set to Inline Ecommerce (same Stibo Item export schema).
The mapping is **the same** as Section 3. Same columns are read, same columns are dropped.

---

## 7. Licensed — Order Sheet
**Template:** `S3/.../licensed/New_Balance_Licensed_OrderSheet.xlsx` | Sheet: **UPC-Report-D365**
**Lambda:** `licensed_ordersheet_main.py` → `NBOrderSheetLoader`

| Template Column | Status | Stibo Attribute | Notes |
|---|---|---|---|
| ITEM_NUMBER | ✅ | AT_PrincipalStyleCode | `row.get("ITEM_NUMBER")` |
| SKU | ✅ | Variant key | `row.get("SKU")` |
| SKU UPC | ⚠️ | — | Not read directly by name |
| UK Size | ⚠️ | — | In EXPECTED_COLS but not mapped to AT_ attribute |
| UPC# | ✅ | AT_PrincipalBarcode | `row.get("UPC#")` |
| Width Description | ❌ | — | Not read |
| SKU Size | ✅ | AT_SizeCode / AT_Size | `row.get("SKU Size")` |
| WIDTH | ❌ | — | In file but not read by mapper |
| SKU Color Code | ✅ | AT_ColorCode, AT_Color | `row.get("SKU Color Code")` |

**❌ Unmapped (2):** `Width Description`, `WIDTH`

---

## Summary Table

| Template | Total Cols | ✅ Mapped | ❌ NOT Mapped | 🐛 Bug |
|---|---|---|---|---|
| Inline Footwear (All Sourced) | 30 | 15 | 15 | — |
| Inline Apparel & Accessories | 43 | 29 | 14 | — |
| Inline Ecommerce (Item tab) | ~75 | ~20 | ~55* | — |
| Inline EAN Source | 9 | 6 | 2 | — |
| Licensed Line List | 43 | 28 | 14 | Primary Fabric Content (trailing space) |
| Licensed Ecommerce | ~75 | ~20 | ~55* | — |
| Licensed Order Sheet | 9 | 5 | 2 | — |

*Ecommerce has many Stibo-internal system columns (`sys_*`, validation errors, assortment flags, markdown prices) that are intentionally not ingested via the inbound XML.

---

## Critical Bug to Fix

```python
# File: licensed_linelist_main.py — line 655
# WRONG (trailing space causes silent None return):
fabric_primary = _s(row.get("Primary Fabric Content "))

# CORRECT:
fabric_primary = _s(row.get("Primary Fabric Content"))
```

---

## ✅ All Critical Stibo Inbound Fields ARE Mapped

The following core Stibo attributes are correctly populated across all relevant templates:
- `AT_PrincipalStyleCode`, `AT_SAPStyleCode`
- `AT_PrincipalColorCode`, `AT_Color`, `AT_PrincipalColorName`, `AT_PrincipalColorDescription`
- `AT_Gender`, `AT_SAPAge`, `AT_BYAge`
- `AT_CountryOrigin`
- `AT_FOB`, `AT_OriginalPrice`, `AT_CurrentPrice`
- `AT_VendorName`
- `AT_MerchandiseCategory`
- `AT_Size` / `AT_PrincipalSize`
- `AT_PrincipalBarcode` / `DC_Barcode` (EAN and Order Sheet)
- `AT_Season`, `AT_LaunchingDate`
- `AT_TechnologyUsed`, `AT_Content` / `AT_Material`
- `AT_Collection1`, `AT_Silhouette`, `AT_LPACategory`, `AT_SMUType`
- `AT_BYArticleType`, `AT_NatureOfArticle`
- `AT_SBU`, `AT_CompanyCode`, `AT_Brand`, `AT_BrandGroup`
- `AT_SAPProductGroup`, `AT_SportsCategoryEN`, `AT_SAPProductDivision`
