# DB1 Pipeline — IGR Property Transaction Database

Converts raw IGR (Inspector General of Registration) Excel downloads into a clean,
enriched property transaction database (DB1) for Mumbai, Thane, and Pune.

---

## Pipeline Overview

```
Raw IGR Excel
     │
     ▼
[ Stage 1 ]  stage1_pre_llm.py
  Load → Clean → Categorise → Village Map
     │
     ▼  (Excel saved for LLM)
[ Stage 2 ]  stage2_llm_process.py
  LLM Extraction → Manual Review
     │
     ▼  (Excel corrected and saved by analyst)
[ Stage 3 ]  stage3_post_manual.py
  Post-Manual Processing → RERA Index → Geocode → BHK → DB1 Output
```

---

## Stage 1 — Pre-LLM (`stage1_pre_llm.py`)

**Input:** Raw merged IGR Excel file (e.g. `Andheri_Merged_All.xlsx`)  
**Output:** `Andheri data for llm.xlsx`

| Step | Function | What it does |
|---|---|---|
| Load & Clean | `load_and_clean()` | Reads Excel, lowercases columns, drops fully-null rows and rows missing `propertydescription` |
| Date Fix | `fix_dates()` | Formats `registrationdate` and `dateofexecution`; toggle `REGISTRATION_DATE_MDY` / `EXECUTION_DATE_MDY` in config per file |
| Dedup | `deduplicate()` | Drops exact duplicates across 8 key columns |
| Categorise | `categorise()` | Tags each row as `Sale`, `Lease`, or `Other` using Marathi doc-type sets |
| Village Map | `map_villages()` | Maps Marathi `areaname` → English village name via the village directory |

**Config to set per run:**
```python
IGR_EXCEL              # path to the raw merged IGR file
VILLAGE_DIR            # path to the village directory Excel
OUTPUT_PATH            # where to save the output
REGISTRATION_DATE_MDY  # True = MM/DD/YYYY, False = DD/MM/YYYY
EXECUTION_DATE_MDY     # same, for dateofexecution
```

---

## Stage 2 — LLM Processing + Manual Review (`stage2_llm_process.py`)

**Input:** Stage 1 output Excel  
**Output:** Analyst-corrected Excel (saved back to same path)

### Stage 2.1 — LLM Extraction

Send each row's `property_description` to the LLM to extract:

| Column | Description |
|---|---|
| `project_name` | Housing project / building name |
| `flat_no` | Flat/unit number as written |
| `property_type_raw` | Raw type string (`Flat`, `Shop`, `Office`, etc.) |
| `floor_no` | Floor number (word or digit) |
| `net_carpet_area_sqmt` | Net carpet area in sq. metres |
| `location` | Transaction village if different from `igr_village` |

After LLM processing, set `llm_processed = "Yes"` for each processed row.

### Stage 2.2 — Manual Review

The analyst opens the LLM output Excel and:

1. Corrects any mis-extracted values in the columns above.
2. Sets `manual_processed = "Yes"` for every reviewed row (whether corrected or not).
3. Saves the file before running Stage 3.

> **Important:** `manual_processed` controls two things in Stage 3 — non-RERA index
> assignment and BHK/coordinate processing. Rows left as `"No"` are treated as unreviewed.

---

## Stage 3 — Post-Manual Processing (`stage3_post_manual.py`)

**Input:** Manually reviewed Excel (with `llm_processed` and `manual_processed` columns)  
**Output:** `testing file db1.xlsx` — final DB1-ready DataFrame

| Step | Function(s) | What it does |
|---|---|---|
| 3.1 Rename & reshape | `rename_columns()`, `process_unit_and_floor()`, `standardise_column_names()`, `select_required_columns()` | Standardises column names, builds `unit_no` strings, maps floor words to numbers, converts area units |
| 3.2 Split | `split_by_category()` | Separates Sale / Lease / Other; only Sale goes through full processing |
| 3.3 Clean sales | `clean_sale_data()`, `_infer_floor_from_unit()` | Second dedup pass, numeric carpet area, infers floor from unit number pattern |
| 3.4 RERA index | `assign_rera_index()`, `_get_rera_values()` | Exact merge then fuzzy fallback against RERA Grand; assigns `mNR###` codes to non-RERA projects |
| 3.5 Geocode | `geocode_coordinates()` | Fills project lat/lng from cache Excel; falls back to ArcGIS API and updates cache |
| 3.6 Property type | `classify_property_type()` | Maps raw type → `Flat / Shop / Office / Others` |
| 3.7 BHK | `assign_bhk_carpet_match()`, `assign_bhk_range_fallback()`, `finalise_bhk()` | 4-stage BHK assignment: exact carpet match → closest match → skip-type retry → percentile range fallback |
| 3.8 Buyer location | `add_buyer_location()` | Extracts 6-digit pincode from purchaser name field; joins postal CSV for locality/district/state |
| 3.9 Final assembly | `final_assembly()` | Merges Sale + Lease + Other, adds quarter/year columns, appends village lat/lng, applies final column renames |

**Config to set per run:**
```python
CITY                # "mumbai" | "thane" | "pune"
RERA_GRAND_PATH     # RERA Grand master Excel
RERA_KEYWORDS_PATH  # BHK keyword mapping Excel
COORDINATES_PATH    # project coordinates cache Excel
VILLAGE_DIR_PATH    # village directory with lat/lng
POSTAL_CSV_PATH     # postal pincode CSV
INPUT_PATH          # manually reviewed Stage 2 output
OUTPUT_PATH         # final DB1 output path
BHK_MAX_DIFF        # max sq.mt tolerance for BHK carpet matching (default: 5)
```

---

## Output Schema

Final columns in DB1 output (in order):

| Column | Description |
|---|---|
| `index` | RERA index or non-RERA code (`mNR101`, `tNR101`, etc.) |
| `project_name` | Project name (LLM-extracted, manually verified) |
| `village_mr` | Marathi village name from IGR (`areaname`) |
| `location` | English transaction village |
| `igr_village` | Mapped English village name |
| `year` / `quarter` | Derived from `transaction_date` |
| `city` | Mumbai / Thane / Pune |
| `sro_name` / `sro_code` | Sub-Registrar Office name and numeric code |
| `document_no` | IGR document number |
| `transaction_type` | Sale Deed / Sale Agreement / etc. |
| `agreement_price` | Consideration amount (₹) |
| `market_value` | Bajarbhav / stamp duty market value (₹) |
| `property_description` | Full raw property description |
| `transaction_date` | Registration date |
| `floor_no` | Floor number |
| `unit_no` | Formatted unit string (e.g. `Flat no. 904`) |
| `property_type_raw` | Raw LLM-extracted type |
| `net_carpet_area_sqmt` | Net carpet area (sq. metres) |
| `balcony_sqmt` / `terrace_sqmt` | Ancillary areas |
| `seller_name` / `purchaser_name` | Party names |
| `property_category` | Sale / Lease / Other |
| `internaldocumentnumber` / `micrno` / `bank_type` / `party_code` | IGR internal fields |
| `dateofexecution` / `stampdutypaid` / `registrationfees` | Financial/legal fields |
| `project_lat` / `project_lng` | Project-level coordinates |
| `location_lat` / `location_lng` | Village-level coordinates |
| `property_type` | Flat / Shop / Office / Others |
| `bhk_br` | BHK classification |
| `buyer_pincode` / `locality_of_buyer` / `district` / `statename` | Buyer location |
| `tower` | Tower (manually filled post-pipeline) |
| `gross_carpet_sqft` / `rate_on_gca_sqft` | Filled post-pipeline |
| `is_duplicate` / `primary_sale_or_secondary_sale` | Filled post-pipeline |
| `llm_processed` / `manual_processed` | Audit flags |

---

## External Dependencies

```
pandas, numpy, openpyxl   # data processing
geopy                     # ArcGIS geocoding fallback
joblib                    # parallel BHK matching
static                    # local module: result_dict, word_number_dict
```

---

## Required Reference Files

| File | Used in | Purpose |
|---|---|---|
| `Mumbai IGR Village Directory.xlsx` | Stage 1 | Marathi → English village name mapping |
| `Mumbai Villages Premanual Directory.xlsx` | Stage 3 | Village lat/lng lookup |
| `mumbai RERA GRAND EXCEL VERSION.xlsx` | Stage 3 | RERA project master with BHK/carpet data |
| `RERA_All_Keywords_BHK_Prop_Type.xlsx` | Stage 3 | BHK keyword normalisation mapping |
| `project address and its coordinates.xlsx` | Stage 3 | Geocoding cache (auto-updated by pipeline) |
| `postal_pincode.csv` | Stage 3 | Pincode → locality/district/state lookup |
