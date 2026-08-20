# DB1 Pipeline Documentation

**Purpose:** Convert raw Maharashtra IGR (Inspector General of Registration) transaction
data into a cleaned, enriched, DB1-ready dataset for property valuation.

**Cities covered:** Mumbai, Thane, Pune

**Pipeline shape:** 3 sequential stages, with a manual human-review checkpoint
between Stage 2 and Stage 3.

```
Raw IGR Excel                                                     Final DB1 dataset
      │                                                                    ▲
      ▼                                                                    │
┌─────────────┐   ┌──────────────┐   ┌─────────────┐   ┌──────────┐   ┌─────────────┐
│   STAGE 1   │──▶│  STAGE 2.1   │──▶│  STAGE 2.2  │──▶│  STAGE 3 │──▶│   Output    │
│ Load/Clean/ │   │ LLM Extract  │   │   Manual    │   │  Post-   │   │  Excel      │
│ Categorise/ │   │  (Gemini)    │   │   Review    │   │  Manual  │   │             │
│ Village Map │   │              │   │  (human)    │   │Processing│   │             │
└─────────────┘   └──────────────┘   └─────────────┘   └──────────┘   └─────────────┘
```

---

## Table of Contents

1. [Stage 1 — Load, Clean, Categorise, Village Map](#stage-1--load-clean-categorise-village-map)
2. [Stage 2 — LLM Processing + Manual Review](#stage-2--llm-processing--manual-review)
3. [Stage 3 — Post-Manual Processing](#stage-3--post-manual-processing)
4. [End-to-End Column Lineage](#end-to-end-column-lineage)
5. [Configuration Reference](#configuration-reference)
6. [Known Issues & Operational Notes](#known-issues--operational-notes)
7. [Running the Pipeline](#running-the-pipeline)

---

## Stage 1 — Load, Clean, Categorise, Village Map

**Script:** `db1_pipeline_stage1.py`
**Input:** Raw IGR Excel export (multi-village merged file)
**Output:** Three Excel files split by category — Sale, Lease, Other — ready for
Stage 2 LLM processing.

### What it does

1. **Load & clean** (`load_and_clean`)
   Reads the raw Excel, lowercases column names, restricts to a known
   `KEEP_COLUMNS` set, drops fully-empty rows, drops rows with no
   `propertydescription`, and title-cases the property description text.

2. **Fix dates** (`fix_dates`)
   Normalises `registrationdate` and `dateofexecution` to a consistent string
   format. **Important:** each downloaded IGR file may mix date formats
   (MM/DD/YYYY vs DD/MM/YYYY) — this is controlled per-run via
   `Stage1Config.registration_date_mdy` / `execution_date_mdy`, not
   auto-detected.

3. **Deduplicate** (`deduplicate`)
   Drops duplicate rows on a fixed subset of columns (`DEDUP_SUBSET`):
   `docno`, `docname`, `registrationdate`, `sroname`, `propertydescription`,
   `areaname`, `consideration_amt`, `marketvalue`.

4. **Categorise** (`categorise`)
   Tags each row `Sale`, `Lease`, or `Other` by matching `docname` against
   two fixed Marathi document-type sets (`SALE_DOCTYPE`, `LEASE_DOCTYPE`).
   Anything not in either set falls to `Other`.

5. **Map villages** (`map_villages`)
   Joins against a Village Directory Excel to attach an English village name
   (`igr_village`) for each row's `areaname`.

6. **Parse floor / unit / property type from `propertydescription`**
   This is the most intricate part of Stage 1, done with layered regex over
   Marathi text:
   - `extract_floor_raw` — pulls a ~15-character window around floor
     keywords (मजल्यावर, मजला, माळा, फ्लोअर, etc.)
   - `clean_floor` — normalises that window down to a short floor label,
     with different logic branches depending on which keyword matched
     (माळा नं, मजला/मजल्या, फ्लोअर, लेवल)
   - `extract_unit_raw` — pulls a substring around unit-number keywords
     (सदनिका, फ्लॅट, ऑफिस, शॉप, युनिट, etc.) followed by digits
   - `clean_unit` — reduces that to `<text> <number>`, dropping anything
     after the first comma
   - `assign_property_type` — classifies the property description into one
     of: Apartment, Row_House, Bunglow, Warehouse, Industrial, Office, Shop,
     Commercial, Flat, Unit, Chawl, Room, via keyword lookup
     (`PROPERTY_TYPE_KEYWORDS`), with special-case fallback logic for
     ambiguous "युनिट" (unit) mentions
   - Finally, `unit_clean` is reformatted as `"<Property Type> no. <n>"`
     wherever both a numeric unit and a resolved property type exist.

7. **Rename & export**
   Columns are renamed to business-friendly names (see
   [column lineage](#end-to-end-column-lineage) below), then the DataFrame is
   split by `property_category` and written to three separate Excel files:
   `Sample Sale data for llm.xlsx`, `Sample Lease data for llm.xlsx`,
   `Sample Other data for llm.xlsx`.

### Key config (`Stage1Config`)

| Field | Purpose |
|---|---|
| `igr_excel` | Path to raw input Excel |
| `village_dir` | Path to village directory Excel |
| `output_sale` / `output_lease` / `output_other` | Output paths per category |
| `registration_date_mdy` / `execution_date_mdy` | Per-run date-format switch (MM/DD vs DD/MM) |

---

## Stage 2 — LLM Processing + Manual Review

**Script:** `db1_pipeline_stage2.py`
**Input:** One of Stage 1's category output files (typically the Sale file)
**Output:** LLM-enriched Excel, then a human-reviewed version of the same file

Split into two phases:

### 2.1 — LLM Extraction (automated)

For each row's `Bhumapan` (property description) text, the script extracts:

| Field | Description |
|---|---|
| `project_name_en` | Project/building name(s), translated to English |
| `Block No` | Block number, translated to English (numbers unchanged) |
| `Carpet_Area`, `BuildUp_Area`, `Saleable_Area`, `Terrace_Area`, `Balcony_Area`, `Other_Area` | Raw area strings with units, e.g. `"58.96 sq.m"` |

**Extraction approach — deterministic first, LLM second:**
- Area values are extracted **first via regex** (`find_area_raw`, using
  keyword patterns for Built-Up/बिल्ट अप, Carpet/कार्पेट, Saleable/सेलेबल,
  Terrace/टेरेस, Balcony/बाल्कनी, and a generic "Other" fallback). This
  regex layer runs regardless of what the LLM returns.
- `Block No` is also pre-extracted via regex (`ब्लॉक नं[:\s]+...`).
- The **Gemini LLM** (`gemma-3-27b-it` by default) is prompted only for
  `project_name_en` and to translate the pre-extracted `Block No` — it
  returns a strict JSON object with 8 fixed keys.
- On `quota`/`rate limit` errors, the call retries with exponential-ish
  backoff (`attempt * 30 + random(5,15)` seconds) up to `retries` times
  (default 3).

**Parallel / chunked processing:**
- Input is split into chunks (`chunk_size`, default 50 rows).
- Each chunk is processed with a `ThreadPoolExecutor` (`max_workers`,
  default 1–2) via `parallel_extract`.
- Each completed chunk is written **atomically** (`.tmp.xlsx` → `os.replace`
  to final `.xlsx`) so a crash mid-chunk never leaves a corrupt file.
- **Resumable:** `run_parallel_extraction(resume=True)` skips chunks whose
  output file already exists and picks up from
  `highest_completed_chunk(output_dir) + 1`. `resume=False` clears the
  output directory and restarts from chunk 1.

**Final merge (`merge_all_chunks`):**
- Concatenates all chunk files in natural numeric order.
- Converts every raw area column to a `_sqmt` numeric column via
  `carpet_to_sqmt`, which:
  - Normalises decimal spacing (`"193. 122"` → `"193.122"`)
  - Detects a sq-metre unit via a large Marathi/English regex
    (`SQMT_REGEX`) and returns the number as-is
  - Otherwise detects a sq-foot unit (`FOOT_REGEX`) and divides by
    `10.764` to convert to sqmt
  - Returns `NaN` if neither pattern matches

Output columns added: `project_name_en`, `Block No`, and 6 raw area columns
plus their `_sqmt` equivalents (12 new columns total).

### 2.2 — Manual Review (human)

After LLM processing, a human analyst opens the output Excel and:

1. Corrects any mis-extracted `project_name`, `flat_no`, `property_type_raw`,
   `floor_no`, `net_carpet_area_sqmt`, or `location` values.
2. Sets `manual_processed = "Yes"` for every row reviewed, regardless of
   whether a correction was needed. Rows left `"No"` are treated as
   unreviewed by Stage 3.
3. Saves the corrected file back to the same path before Stage 3 runs.

Stage 3 reads `manual_processed` to:
- Assign non-RERA index codes **only** to manually confirmed rows.
- Skip BHK / coordinate logic for unreviewed rows as a fail-safe.

### Key config (`Stage2Config`)

| Field | Purpose |
|---|---|
| `input_excel_path` / `input_sheet_name` | Stage 1 output to process |
| `temp_csv_path` | Intermediate CSV (Excel is converted to CSV before chunking) |
| `output_chunk_dir` | Where per-chunk `.xlsx` files land |
| `final_output_path` | Merged output for manual review |
| `model_name` | Gemini model (default `gemma-3-27b-it`) |
| `chunk_size`, `max_workers`, `resume`, `retries` | Processing controls |

**Credential handling:** the Gemini API key is read from the
`GOOGLE_API_KEY` environment variable — it must be exported before running
this script. (An earlier version of this script had the key hardcoded in
source; that key should be considered compromised and rotated if it hasn't
been already.)

---

## Stage 3 — Post-Manual Processing

**Script:** `db1_pipeline_stage3.py`
**Input:** The manually reviewed Excel from Stage 2.2 (must have
`llm_processed` and `manual_processed` columns)
**Output:** `result_df` — the final DB1-ready dataset, written to
`sample_file_for_db1.xlsx`

This is the largest stage. It runs as a numbered sequence, 3.0 through 3.9.

### 3.0 — Net carpet area cascade (`derive_net_carpet_area`)

Fills `net_carpet_area_sqmt` from whichever area column is available, in
priority order, with **city-dependent conversion factors**:

| Priority | Source column | Conversion |
|---|---|---|
| 1 | `Carpet_Area_sqmt` | used as-is |
| 2 | `BuildUp_Area_sqmt` | ÷ **1.2** (fixed, same for all cities) |
| 3 | `Saleable_Area_sqmt` | ÷ **city-specific divisor** — Mumbai 1.45, Thane 1.4, Pune 1.35 |
| 4 | `Other_Area_sqmt` | used as-is |

The city-specific Saleable divisor is looked up from
`SALEABLE_TO_CARPET_DIVISOR_BY_CITY` using `config.city`; an unrecognised
city raises a `ValueError` rather than silently defaulting.

### 3.1 — Rename, unit/floor, column standardisation

- `rename_columns` — renames Stage 1/2 output columns to their Stage 3
  working names (e.g. `Bhumapan` → `property_description`, `Agreement
  Price(INR)` → `agreement_price`).
- `process_unit_and_floor` — rebuilds `Unit No` as `"<property_type_raw> no.
  <n>"` using `Int64`-safe extraction (avoids pandas dtype conflicts between
  nullable ints and strings), and maps `Floor No` through `word_number_dict`
  (e.g. `"ninth"` → `9`).
- `standardise_columns` — lowercases and underscores all column names.
- `select_required_columns` — restricts to a fixed whitelist of ~27 columns
  going into Stage 3.2.

### 3.2 — Split & clean sale data

- `split_by_category` — splits into `sale`, `lease`, `other` DataFrames
  (lease/other pass through largely unchanged until final assembly).
- `clean_sale_data` — deduplicates sale rows on a business key (`village`,
  `sro_name`, `document_no`, `transaction_type`, `property_description`,
  `transaction_date`, `agreement_price`), coerces `net_carpet_area_sqmt` to
  numeric and logs how many rows fall outside a `[5, 3000]` sqmt sanity
  range, then calls `_infer_floor_from_unit` to back-fill missing floor
  numbers from 3–4 digit unit numbers (e.g. unit `904` → floor `9`, only
  when the last two digits are `< 50`, to avoid misreading e.g. unit `199`
  as floor `1`).

### 3.3 — RERA index assignment (`assign_rera_index`)

Matches each transaction to a RERA project registration index, in three
passes:

1. **Exact merge** on `(project_name, location)` against `(modified_project_name,
   rera_location)` in the RERA Grand Excel.
2. **Fuzzy fallback** (`_get_rera_values`) for rows still missing an index:
   exact match first, then a list-membership check against
   `rera_location_v1` (a column that can contain a stringified Python list
   of alternate location names).
3. **Non-RERA index assignment**: for rows still missing an index *and*
   `manual_processed == "Yes"`, generates synthetic index codes of the form
   `<city-initial>NR<101, 102, ...>` (e.g. `mNR101`) grouped by
   `(project_name, igr_village)`, so the same unregistered project always
   gets the same synthetic index within a run.

### 3.4 — Geocoding (`geocode_coordinates`)

- Builds an `api_call_input` string (`"<project_name>, <igr_village>,
  <city>"`) and looks it up against a cached coordinates Excel first.
- For any input still missing lat/lng, falls back to live geocoding via
  **ArcGIS** (`geopy.geocoders.ArcGIS`), then **appends the new result back
  into the cache file on disk** so repeat runs don't re-hit the API for the
  same input.

### 3.5 — Property type classification (`classify_property_type`)

Maps the raw `property_type_raw` string down to one of 4 buckets: `Flat`
(includes Apartment, Flat/Shop, Duplex), `Shop` (includes Showroom),
`Office`, or `Others`.

### 3.6 — BHK assignment (4-stage cascade)

The most complex logic in the pipeline. Only applies to rows where
`property_type == 'Flat'`.

- **`combine_columns`** — flattens a nested, sometimes doubly-stringified
  BHK-wise carpet area structure (from the RERA Grand Excel's
  `bhk_wise_ca` field) into `{BHK_KEY: [area_values...]}`. Uses nested
  `eval()` calls with fallbacks because the source data's nesting depth is
  inconsistent.
- **`assign_bhk_carpet_match`** — for each sale row, given its
  `net_carpet_area_sqmt` and the matched RERA project's building-wise
  carpet area table:
  - **Stage 1 (exact):** carpet area matches a listed value exactly.
  - **Stage 2 (closest):** nearest listed value within `bhk_max_diff`
    sqmt (default 5).
  - **Stage 3 (retry without skip-types):** if the match landed on a
    non-informative bucket (`UNDEFINED FLATS`, `SHOP`, `OFFICE`, `OTHERS`),
    retries the closest-match search excluding those buckets.
  - Runs in parallel across rows via `joblib.Parallel` (threading backend),
    with a pre-parsed cache (`parse_cache`) so each unique
    `building_wise_carpet_area` string is only `ast.literal_eval`'d once.
- **`assign_bhk_range_fallback`** — for rows still without a BHK, builds
  project-level BHK carpet-area percentile ranges (10th/90th percentile per
  BHK bucket across the whole RERA dataset) and assigns a BHK by which
  range the row's carpet area falls into. This is a coarser, dataset-wide
  fallback rather than a project-specific match.
- **`finalise_bhk`** — title-cases and cleans the `BHK` column, and
  back-fills any still-missing BHK with the row's `property_type` (e.g. a
  Shop with no BHK match gets `BHK = "SHOP"`).

### 3.7 — Buyer location (`add_buyer_location`)

Extracts a 6-digit pincode from the `purchaser_name` field (regex
`\b\d{6}\b`) and joins against a postal pincode CSV to attach buyer
locality/district/state.

### 3.8 — Final assembly (`final_assembly`)

- Re-concatenates the cleaned sale rows with the (mostly unmodified) lease
  and other rows.
- Derives `quarter` (`"Q<n>-<year>"`) and `year` from `transaction_date`.
- Adds placeholder columns for downstream use: `Tower`,
  `gross_carpet_sqft`, `rate_on_gca_sqft`, `is_duplicate`, `Primary
  Sale_or_Secondary Sale` — all `None` at this stage, populated later
  outside this pipeline.
- Maps `sro_code` from `sro_name` via the per-city `SRO_DICT`.
- Selects and reorders the final column set, standardises column names
  again, joins village-level lat/lng from the Village Directory, and
  renames a few columns for the final schema (`village` → `village_mr`,
  `bajarbhav` → `market_value`, `bhk` → `bhk_br`).
- Writes the result to `config.output_path`.

### Key config (`Stage3Config`)

| Field | Purpose |
|---|---|
| `city` | `"mumbai"` \| `"thane"` \| `"pune"` — drives SRO codes and the Saleable→Carpet divisor |
| `input_path` | Manually reviewed Stage 2 output |
| `rera_grand_path` | RERA Grand Excel (project index, location, BHK carpet areas) |
| `rera_keywords_path` | BHK/property-type keyword normalisation table |
| `coordinates_path` | Cached project lat/lng lookup (read **and written**) |
| `village_dir_path` | Village directory (English names + lat/lng) |
| `postal_csv_path` | Pincode → locality/district/state lookup |
| `output_path` | Final DB1-ready Excel |
| `bhk_max_diff` | Max sqmt difference allowed for BHK "closest match" (default 5) |

---

## End-to-End Column Lineage

Selected columns as they're renamed across the pipeline:

| Raw IGR column | Stage 1 output | Stage 3 working name | Final output name |
|---|---|---|---|
| `registrationdate` | `Transaction Date` | `registrationdate` (renamed back) | `transaction_date` |
| `consideration_amt` | `Agreement Price(INR)` | `agreement_price` | `agreement_price` |
| `marketvalue` | `Bajarbhav` | `bajarbhav` | `market_value` |
| `docname` | `Document Type` | `document_type` | *(dropped — replaced by `transaction_type`)* |
| `docno` | `Document No` | `document_no` | `document_no` |
| `sroname` | `SRO Name` | `sro_name` | `sro_name` |
| `propertydescription` | `Bhumapan` | `property_description` | `property_description` |
| `sellerparty` | `Seller Name` | `seller_name` | `seller_name` |
| `purchaserparty` | `Purchaser Name` | `purchaser_name` | `purchaser_name` |
| `areaname` | *(kept as `areaname`)* | `village` | `village_mr` |
| — | `igr_village` (from village directory) | `igr_village` | `igr_village` |
| — | `Property Type` | `property_type_raw` | `property_type_raw` |
| — | `unit_clean` | `Unit No` → `unit_no` | `unit_no` |
| — | `floor_clean` | `Floor No` → `floor_no` | `floor_no` |
| — (Stage 2) | `project_name_en` | `Modified_Project_Name_1` → `project_name` | `project_name` |
| — (Stage 2) | `Carpet_Area_sqmt` etc. | `net_carpet_area_sqmt` (cascade) | `net_carpet_area_sqmt` |
| — (Stage 3) | — | `BHK` | `bhk_br` |
| — (Stage 3) | — | `index` (RERA or synthetic) | `index` |

---

## Known Issues & Operational Notes

- **Gemini API key exposure (Stage 2):** an earlier version of the Stage 2
  script had a live API key hardcoded in source. It now reads
  `GOOGLE_API_KEY` from the environment and fails fast if unset. **If that
  original key hasn't been rotated yet, do so.**
- **Date format is not auto-detected (Stage 1):** `registration_date_mdy`
  and `execution_date_mdy` must be set correctly per input file, or dates
  will silently parse wrong (e.g. `03/04/2024` as March 4th vs April 3rd).
- **City-dependent conversion factors (Stage 3):** the Saleable→Carpet
  divisor is city-specific (Mumbai 1.45 / Thane 1.4 / Pune 1.35) while
  Built-Up→Carpet is fixed at 1.2. Running Stage 3 with the wrong
  `config.city` will silently apply the wrong divisor for any row that
  falls back to Saleable area.
- **Non-RERA index codes reset per run:** synthetic index codes
  (`<city>NR101`, `102`, ...) are generated fresh from the current run's
  unmatched projects each time Stage 3 runs — they are **not** stable
  across separate runs unless the set of unmatched `(project_name,
  igr_village)` pairs is identical. Re-running Stage 3 on an updated input
  file can reassign different synthetic codes to the same project.
  Something to watch for if these codes are used as a persistent join key
  elsewhere.
- **Live geocoding calls (Stage 3):** any project/village combination not
  already in the coordinates cache triggers a live ArcGIS geocode call at
  runtime. This makes Stage 3's runtime and reliability dependent on
  external API availability for first-time projects. New results are
  cached back to disk, so this cost is only paid once per project+village.
- **BHK matching relies on nested, sometimes double-stringified JSON**
  (`combine_columns`) from the RERA Grand Excel — this is inherently fragile
  to upstream data-shape changes in that source file.
- **Floor inference heuristic (Stage 3.2)** assumes unit numbers follow a
  `<floor><unit-on-floor>` pattern (e.g. `904` → floor 9, unit 04) and
  explicitly skips cases where the last two digits are `≥ 50`, to avoid
  misreading three-digit unit numbers as floor+unit. This is a heuristic,
  not a guaranteed-correct parse.

---

## Running the Pipeline

```bash
# Stage 1 — per-run: set date format flags for the specific IGR export
python db1_pipeline_stage1.py
# → produces Sample Sale/Lease/Other data for llm.xlsx

# Stage 2.1 — requires GOOGLE_API_KEY in environment
export GOOGLE_API_KEY=your-key-here
python db1_pipeline_stage2.py
# → produces "llm processed Sale data for manual.xlsx"

# Stage 2.2 — MANUAL STEP
# Open the Stage 2.1 output, correct extracted fields, set
# manual_processed = "Yes" per reviewed row, save back to the same path.

# Stage 3 — set city in Stage3Config before running
python db1_pipeline_stage3.py
# → produces sample_file_for_db1.xlsx (final DB1-ready dataset)
```

Each stage's `run_stageN()` function accepts a config object
(`Stage1Config` / `Stage2Config` / `Stage3Config`), so paths, city, and
processing parameters can be overridden per run without editing the script,
e.g.:

```python
from db1_pipeline_stage3 import run_stage3, Stage3Config

run_stage3(Stage3Config(city="pune", input_path=Path("pune_reviewed.xlsx")))
```
