# Dubai Land Department (DLD) Transaction Data Processing Pipeline

An automated, end-to-end Python ETL pipeline designed to clean, standardize, enrich, and format raw Dubai Land Department (DLD) real estate transaction data. The pipeline merges raw monthly CSV exports with existing dimension data in PostgreSQL and assigns standardized project indexes and geolocation coordinates.

---

## Table of Contents

- [Overview](#overview)
- [Pipeline Architecture & Workflow](#pipeline-architecture--workflow)
- [Prerequisites & Dependencies](#prerequisites--dependencies)
- [Database Schema Requirements](#database-schema-requirements)
- [Configuration](#configuration)
- [How to Run](#how-to-run)
- [Generated Outputs](#generated-outputs)
- [Security & Best Practices](#security--best-practices)

---

## Overview

The [`dld_pipeline.py`](file:///d:/Pankaj/DubaiDataProcessing/dld_pipeline.py) script consolidates raw transaction processing and database-backed entity resolution into a single executable script.

### Key Objectives:
- **Deduplication**: Eliminates duplicate transactions based on a 22-column business key.
- **Normalization**: Standardizes property types, unit configurations (e.g., `1 B/R` to `1Bhk`), transaction categories, and temporal attributes (quarter and year).
- **Schema Alignment**: Maps raw DLD field names to an internal target schema (70+ columns) matching downstream data warehouse requirements.
- **Entity Resolution & Spatial Enrichment**: Enriches records with geolocation coordinates (`latitude`, `longitude`) and existing master indexes from PostgreSQL dimension tables.
- **Automated Index Assignment**: Automatically assigns collision-free sequential non-registered indexes (`nrXXXX`) to new projects and locations.
- **Data Hygiene**: Enforces blanking rules for orphan records and applies consistent title-casing across text fields.

---

## Pipeline Architecture & Workflow

The pipeline runs sequentially top-to-bottom across distinct processing phases:

```
[Raw DLD CSV]
       │
       ▼
 [Step 1 & 1b]  Load CSV & Drop Duplicates (22 composite fields)
       │
       ▼
 [Step 2 - 4b]  Parse Dates, Map Property Types & Categorize Units
       │
       ▼
 [Step 5 & 6]   Add Required Columns & Default System Metadata
       │
       ▼
 [Step A]       Enrich from Postgres (dim_project & dim_location)
       │        └── Exports intermediate test_merge file
       ▼
 [Step B]       Resolve / Generate Sequential 'nr' Indexes (transactions table)
       │
       ▼
 [Step C]       Location Coordinate Fallback (dim_location)
       │
       ▼
 [Final]        Reorder Schema (70+ columns), Title-Case & Save Final Excel/CSV
```

### Detailed Execution Steps

| Step | Operation | Description |
| :--- | :--- | :--- |
| **1 & 1b** | **Ingestion & Dedup** | Reads raw CSV as string types (`utf-8-sig`). Removes duplicates using composite transaction identifiers (`TRANSACTION_NUMBER`, `INSTANCE_DATE`, `TRANS_VALUE`, `AREA_EN`, `PROJECT_EN`, etc.). |
| **2** | **Date Parsing** | Converts `INSTANCE_DATE` to datetime; extracts integer `year` and quarter string (e.g., `Q3-2026`). |
| **3** | **Property Type Mapping** | Translates raw sub-types (`PROP_SB_TYPE_EN`) into canonical types: `Flat`, `Villa`, `Plot`, `Office`, `Shop`, or `Others`. |
| **4, 4a, 4b**| **Field Renaming & Classification** | Renames raw DLD columns to warehouse schema. Normalizes categories (`Mortgage` → `Mortgages`, `Sales` → `Sale`) and unit configurations (`1 B/R` → `1Bhk`, `4 B/R` → `>3Bhk`). |
| **5 & 6** | **Schema Setup & Defaults** | Fills missing required fields with blanks. Sets static tags: `city_name="Dubai"`, `country_name="United Arab Emirates"`, `data_source="Dld"`, `data_type="Registered Document"`. |
| **A** | **Postgres Entity Match** | Connects to PostgreSQL to query `dim_project` joined with `dim_location` (filtered by `city_id`). Merges coordinates and `internal_index_id` on normalized `(project_name, location_name)`. Saves intermediate checkpoint file (`TEST_MERGE_PATH`). |
| **B** | **Index Assignment** | Identifies maximum `nr` index from `public.transactions` (e.g., `nr1050`). Assigns existing index from DB/file or allocates next sequential `nr` ID (`nr1051`, etc.) for unique project-location pairs. |
| **C** | **Location Coordinate Fallback** | Queries `dim_location` to populate missing latitude/longitude for rows with a valid `location_name` whose project-level match lacked coordinates. |
| **Final** | **Order, Format & Save** | Enforces exact 70-column target order, applies title-casing to all text fields, and exports the final file (`FINAL_OUTPUT_PATH`). |

---

## Prerequisites & Dependencies

### Python Environment
- Python 3.9 or higher

### Required Libraries
Install the required packages using `pip`:

```bash
pip install pandas numpy psycopg2-binary sqlalchemy openpyxl
```

| Package | Purpose |
| :--- | :--- |
| `pandas` | Primary DataFrame transformations and ETL processing |
| `numpy` | Vectorized operations and NaN handling |
| `psycopg2-binary` / `psycopg2` | PostgreSQL adapter for Python |
| `sqlalchemy` | SQL querying and database connection management |
| `openpyxl` | Engine for reading and writing `.xlsx` workbooks |

---

## Database Schema Requirements

The pipeline expects a PostgreSQL database with the following tables and columns accessible to the configured user:

### 1. `public.dim_project`
- `project_name` (text/varchar)
- `internal_index_id` (text/varchar)
- `project_latitude` (numeric/float)
- `project_longitude` (numeric/float)
- `city_id` (integer)
- `location_id` (integer)

### 2. `public.dim_location`
- `location_id` (integer)
- `location_name` (text/varchar)
- `location_latitude` (numeric/float)
- `location_longitude` (numeric/float)
- `city_id` (integer)

### 3. `public.transactions`
- `internal_index_id` (text/varchar)
- `city_id` (integer)
- `city_name` (text/varchar)
- `project_name` (text/varchar)
- `location_name` (text/varchar)

---

## Configuration

All configurable parameters are located at the top of [`dld_pipeline.py`](file:///d:/Pankaj/DubaiDataProcessing/dld_pipeline.py) under the `# ============================== CONFIG ==============================` block:

```python
# File Paths
RAW_CSV_PATH = r"D:\Pankaj\DubaiDataProcessing\transactions-2026-09-17.csv"
FINAL_OUTPUT_PATH = r"D:\Pankaj\DubaiDataProcessing\transactions-2026-09-17_processed.xlsx"
TEST_MERGE_PATH = r"D:\Pankaj\DubaiDataProcessing\transactions-2026-09-17_test_merge.xlsx"

# Scope Parameters
CITY_ID = 1
CITY_NAME = "Dubai"

# PostgreSQL Connection Settings
db_params = {
    "host": "localhost",
    "port": "5433",
    "database": "dubai",
    "user": "postgres",
    "password": "YOUR_PASSWORD"
}
```

> **Note**: Both `.xlsx` and `.csv` extensions are supported for `FINAL_OUTPUT_PATH` and `TEST_MERGE_PATH`. The script automatically selects `to_excel()` or `to_csv()` based on the file extension.

---

## How to Run

1. **Verify Database Connectivity**: Ensure your local or remote PostgreSQL instance is running and accessible on the specified host and port.
2. **Set File Paths**: Edit `RAW_CSV_PATH` in [`dld_pipeline.py`](file:///d:/Pankaj/DubaiDataProcessing/dld_pipeline.py) to point to your input file.
3. **Execute the Script**: Run via terminal or command prompt:

```bash
python dld_pipeline.py
```

### Console Output
The script prints progress logs for each step, including:
- Ingested row count and duplicate removal count
- Unparseable date counts
- Frequency counts for mapped categories
- Number of rows matched against database projects and locations
- Next available `nr` index assigned
- Remaining unmapped/null coordinates and indexes
- Final confirmation of saved file paths

---

## Generated Outputs

1. **Test Merge File (`TEST_MERGE_PATH`)**:
   - Intermediate checkpoint saved after Step A (PostgreSQL `dim_project` and `dim_location` matching).
   - Useful for auditing initial match rates and database coordinate lookups prior to `nr` assignment.

2. **Processed Output File (`FINAL_OUTPUT_PATH`)**:
   - Production-ready dataset with all missing indexes resolved, fallback coordinates populated, title-casing applied, and standard schema enforced.

---

## Security & Best Practices

- **Avoid Hardcoding Credentials**: For production environments, consider reading database credentials from environment variables or a `.env` file rather than plain text in the script:
  ```python
  import os
  db_params = {
      "host": os.getenv("DB_HOST", "localhost"),
      "port": os.getenv("DB_PORT", "5433"),
      "database": os.getenv("DB_NAME", "dubai"),
      "user": os.getenv("DB_USER", "postgres"),
      "password": os.getenv("DB_PASSWORD", "")
  }
  ```
- **Blank Projects Handling**: The script enforces a data hygiene rule where rows without a valid `project_name` will intentionally have `index`, `project_latitude`, and `project_longitude` cleared to prevent assigning false project linkages.
- **Encoding**: Raw CSV input is loaded with `encoding="utf-8-sig"` to prevent BOM (`\ufeff`) header issues. Output CSVs (if configured) are likewise exported with `utf-8-sig`.
