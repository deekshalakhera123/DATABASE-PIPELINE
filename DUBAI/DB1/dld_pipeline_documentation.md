# Dubai DLD Monthly Transaction Pipeline Documentation

This document provides a comprehensive technical overview and operational guide for `dld_pipeline.py`, which processes raw monthly transaction data from the Dubai Land Department (DLD) and enriches it using Postgres database metadata.

---

## 1. Pipeline Overview
The pipeline functions as a single, combined ETL (Extract, Transform, Load) script that converts raw monthly transactional data in CSV format into a normalized, enriched Excel spreadsheet ready for insertion into the central database. 

It handles data loading, strict deduplication, field normalization, database project-location enrichment, automatic incremental index (`nr`) assignment for new projects, coordinate fallbacks, and schema alignment.

---

## 2. Configuration Settings
All pipeline settings are configured in the `CONFIG` section of the script.

| Parameter | Type | Default Value | Description |
| :--- | :--- | :--- | :--- |
| `RAW_CSV_PATH` | Path | `D:\Database\Dubai_update\1july_12Aug\documentation\1_july_to_12_aug_raw.csv` | Input raw CSV file containing DLD transaction records. |
| `FINAL_OUTPUT_PATH` | Path | `D:\Database\Dubai_update\1july_12Aug\documentation\1_july_to_12_aug_processed.xlsx` | Output Excel spreadsheet with fully processed and enriched records. |
| `TEST_MERGE_PATH` | Path | `D:\Database\Dubai_update\1july_12Aug\documentation\1_july_to_12_aug_test_merge.xlsx` | Output path for intermediate verification after Step A. |
| `CITY_ID` | Integer | `15` | Database identifier for Dubai. |
| `CITY_NAME` | String | `"Dubai"` | Case-sensitive city name. |
| `db_params` | Dictionary | `host: localhost`, `port: 5432`, `database: nilesh`, ... | Credentials and target details for PostgreSQL database. |

---

## 3. Pipeline Execution Workflow

The script executes sequentially from top to bottom, divided into the following key steps:

### STEP 1 & 1b: Load and Deduplicate
1. The script reads the raw input file from `RAW_CSV_PATH`.
2. All columns are loaded as string types (`dtype=str`) with `utf-8-sig` encoding to handle special characters.
3. Whitespace is stripped from column names.
4. **Deduplication**: Rows are evaluated for exact duplication across **22 columns**:
   - `TRANSACTION_NUMBER`, `INSTANCE_DATE`, `GROUP_EN`, `PROCEDURE_EN`, `IS_OFFPLAN_EN`, `IS_FREE_HOLD_EN`, `USAGE_EN`, `AREA_EN`, `PROP_TYPE_EN`, `PROP_SB_TYPE_EN`, `TRANS_VALUE`, `PROCEDURE_AREA`, `ACTUAL_AREA`, `ROOMS_EN`, `PARKING`, `NEAREST_METRO_EN`, `NEAREST_MALL_EN`, `NEAREST_LANDMARK_EN`, `TOTAL_BUYER`, `TOTAL_SELLER`, `MASTER_PROJECT_EN`, `PROJECT_EN`.
   - The first occurrence is preserved (`keep="first"`), and subsequent duplicate rows are discarded.

### STEP 2: Date Parsing and Fiscal Quarter Computation
1. `INSTANCE_DATE` is parsed using format `%Y-%m-%d %H:%M:%S`. Unparseable dates are safely converted to `NaT` (Not a Time).
2. The `year` is extracted.
3. The fiscal `quarter` is formatted as a string in the pattern `Q{quarter}-{year}` (e.g. `Q3-2026`).

### STEP 3: Property Type Normalization
The raw property sub-type column (`PROP_SB_TYPE_EN`) is renamed to `property_type_raw` and mapped to clean core types:
- **Plot**: Agricultural, Land.
- **Flat**: Flat, Government Housing, Hotel Apartment, Residential, Residential Flats, Studio.
- **Villa**: Residential / Attached Villas, Residential / Residential Villa, Residential / Villas, Stacked Townhouses, Villa.
- **Office**: Office.
- **Shop**: Shop, Shopping Mall, Show Rooms.
- **Others**: Airport, Building, Commercial, Commercial / Offices / Residential, Exhibition Center, General Use, Health Club, Hospital, Hotel, Hotel Rooms, Industrial, Labor Camp, Petrol Station, School, Sports Club, Unit, Warehouse, Workshop.
- Unmapped values default to `"Others"`.

### STEP 4, 4a & 4b: Column Renaming and Field Transformations
- **Column Translation**: Raw DLD columns are renamed to standardized database names (e.g., `TRANSACTION_NUMBER` $\rightarrow$ `document_number`, `TRANS_VALUE` $\rightarrow$ `agreement_price`, `AREA_EN` $\rightarrow$ `location_name`, etc.).
- **Transaction Category**: Normalizes transaction categories:
  - `Sales` $\rightarrow$ `Sale`
  - `Mortgage` $\rightarrow$ `Mortgages`
  - `Gifts` $\rightarrow$ `Gifts`
- **Unit Configuration Mapping**: Maps bedroom configurations (`ROOMS_EN` / `unit_configuration`) into normalized structures:
  - `1 B/R` $\rightarrow$ `1Bhk`
  - `2 B/R` $\rightarrow$ `2Bhk`
  - `3 B/R` $\rightarrow$ `3Bhk`
  - `4 B/R` to `10 B/R` $\rightarrow$ `>3Bhk`
  - `Studio` / `Single Room` $\rightarrow$ `Flat`
  - `Office` $\rightarrow$ `Office`
  - `Shop` $\rightarrow$ `Shop`
  - `Gym`, `Hotel`, `Penthouse`, `Store` $\rightarrow$ `Others`

### STEP 5 & 6: Required Columns and Static Values
1. The script ensures that all 70+ required database columns exist. Any missing target columns are added and initialized with `None`.
2. Static/default values are applied:
   - `city_name` = `"Dubai"`, `state_name` = `"Dubai"`, `country_name` = `"United Arab Emirates"`
   - `is_llm_processed` = `"No"`, `is_manual_processed` = `"No"`
   - `source_accessibility` = `"Easy"`, `source_accessibility_way` = `"Download"`
   - `data_type` = `"Registered Document"`, `data_source` = `"Dld"`

---

## 4. Database Enrichment & Indexing Integration

The script connects to Postgres to resolve references and indexes using transactional mapping tables:

```
[Raw Dataframe]
       │
       ▼
[STEP A: Project & Location Match] ──► Query public.dim_project & dim_location
       │                               Validate matches by city_id = 15
       ▼
[STEP B: Assign NR Indexes] ─────────► Check existing mappings (Cache & DB)
       │                               Generate new 'nrXXXX' if unique
       ▼
[STEP C: Coordinate Fallback] ───────► Query public.dim_location
                                       Fill missing coords by location name
```

### STEP A: Project & Location Matching
1. Connects to Postgres using `SQLAlchemy`.
2. Selects all rows from `public.dim_project` joined with `public.dim_location` filtered by `city_id = 15`.
3. Performs a left merge between the active dataframe and the DB lookup dataframe on normalized project and location keys:
   `_project_key` and `_location_key` (derived via lowercase, trimmed, whitespace-normalized strings).
4. **Index Enrichment Rules**:
   - `index` (storing database index project ID) is only inherited if **both** the project name and the location name are matched on that row.
   - If either project or location names are missing or blank, the row is bypassed, and the index remains blank by design.
5. Fills coordinate fields (`project_latitude`, `project_longitude`, `location_latitude`, `location_longitude`) and IDs (`location_id`) from matched records.
6. The intermediate dataset is written to `TEST_MERGE_PATH` for audit tracking.

### STEP B: Assign NR Indexes (`nr` Style Tracking)
This step handles project matching when no pre-existing match was found in `dim_project` during Step A:
1. **Find Maximum Index ID**: Queries the maximum sequential counter of existing `nr`-style indices (e.g. `nr10340`) in `public.transactions` where the index matching regex pattern `^nr(\d+)` belongs to `CITY_ID = 15`.
2. **Current Mapping Cache**: Queries all unique mapping entries of `(location_name, project_name, city_name) -> internal_index_id` from historical transaction tables.
3. **Assignment Logic**:
   - For each row where `index` is null:
     - Check if the key `(location, project, city)` has already been assigned an index during the *current script run*. If so, reuse it.
     - Check if it exists in the *database mappings*. If so, retrieve and assign it.
     - If it is completely new, generate a new index `nr{next_num}`. Increment `next_num` while verifying it does not collide with existing index lists.
4. **Data Cleaning**: Strips trailing suffixes (e.g. `__dubai`), handles decimal conversions (e.g. `1001.0` $\rightarrow$ `1001`), and maps to clean representations.

### STEP C: Location Coordinate Fallback
If project-level coordinate matching was unsuccessful:
1. Queries `public.dim_location` to load all valid locations and coordinates.
2. For any row missing `location_latitude` or `location_longitude` but having a valid `location_name`, it matches on the normalized name and writes the corresponding coordinate.

---

## 5. Final Output Formatting
1. Reorders columns to match the target database schema structure exactly.
2. Formats all object and string columns in Title Case (cosmetic normalization).
3. Writes the finalized dataframe into the target Excel workbook specified by `FINAL_OUTPUT_PATH`.
