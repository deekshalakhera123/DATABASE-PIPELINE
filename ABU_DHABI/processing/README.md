# Abu Dhabi Real Estate (ADREC) Transaction Data Processing Pipeline

[![Python Version](https://img.shields.io/badge/python-3.8%2B-blue.svg)](https://www.python.org/)
[![Jupyter Notebook](https://img.shields.io/badge/jupyter-notebook-orange.svg)](https://jupyter.org/)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
[![Data Source](https://img.shields.io/badge/data%20source-ADREC-orange.svg)](https://www.adrec.gov.ae/)

A production-grade, standalone data processing and standardization pipeline for Abu Dhabi Real Estate Centre (ADREC) transaction records. This pipeline cleans raw transaction data, standardizes property and asset class taxonomies, enriches spatial coordinates, assigns sequential Non-RERA project identifiers (`internal_index_id`), and formats the output to conform strictly with enterprise real estate database schemas (DB1 standard).

The pipeline is provided in two formats:
1. **Interactive Jupyter Notebook**: [`adrec_pipeline.ipynb`](adrec_pipeline.ipynb) for step-by-step interactive inspection, visualization, and execution.
2. **Production Python Script**: [`adrec_pipeline.py`](adrec_pipeline.py) for command-line execution and automated scheduled processing.

---

## Table of Contents

- [Overview](#overview)
- [Input Files Specification](#input-files-specification)
- [Pipeline Architecture & Stages](#pipeline-architecture--stages)
- [Repository Structure](#repository-structure)
- [Taxonomy & Standardization Rules](#taxonomy--standardization-rules)
  - [Property Type Mapping](#property-type-mapping)
  - [Asset Class & Project Type Mapping](#asset-class--project-type-mapping)
  - [Unit Configuration Normalization](#unit-configuration-normalization)
  - [Spatial Enrichment & Fallbacks](#spatial-enrichment--fallbacks)
  - [Non-RERA Index Assignment (`internal_index_id`)](#non-rera-index-assignment-internal_index_id)
- [Output Schema (68 Columns)](#output-schema-68-columns)
- [Installation & Setup](#installation--setup)
- [Usage Guide](#usage-guide)
  - [Option A: Running the Jupyter Notebook](#option-a-running-the-jupyter-notebook)
  - [Option B: Running the Python Script via CLI](#option-b-running-the-python-script-via-cli)
  - [Option C: Using as a Python Module](#option-c-using-as-a-python-module)
- [Environment Variables](#environment-variables)
- [License](#license)

---

## Overview

The Abu Dhabi Real Estate (ADREC) Pipeline ingests raw property transaction records, deduplicates and normalizes them, and produces clean, standardized datasets ready for database ingestion and spatial analytics.

### Key Capabilities:
- **Zero Server Bloat**: Fully detached from private cloud servers; operates cleanly on local files.
- **Three Core Inputs**: Easily configurable paths for raw transaction data, master project coordinates, and master location coordinates.
- **Deterministic Deduplication**: Purges duplicate transactions while maintaining data integrity.
- **Robust Date Parsing**: Normalizes heterogeneous date formats into standard `DD-MM-YYYY`, `year`, and fiscal `quarter` (`QX-YYYY`).
- **Master Reference Geocoding**: Matches records against master location coordinates and existing project databases.
- **Non-RERA Index Generation**: Assigns sequential, persistent project indexes (`nr468`, `nr469`, etc.) grouped by `(location_name, project_name, city_name)`.
- **Strict DB Schema Compliance**: Outputs 68 standardized columns, places `internal_index_id` at column 0, purges `project_id`, and formats strings in Title Case.

---

## Input Files Specification

The pipeline requires three input files:

| Input File | Default Path | Description |
|:---|:---|:---|
| **Raw Transaction Data** | `AbuDhabi_RawData.csv` | Export from ADREC containing transaction details (`Sale Application Date`, `Property Sale Price (AED)`, `Property Sold Area (SQM)`, `Property Type`, `Asset Class`, `District`, `Project Name`, etc.). |
| **Master Project Coordinates** | `Abudhabi_project_coordinates.csv` | Master reference table containing known project names, existing Non-RERA indexes (`internal_index_id`), and project latitude/longitude. |
| **Master Location Coordinates** | `Abudhabi_location_coordinates.csv` | Master reference table mapping districts to `location_id`, `location_latitude`, and `location_longitude`. |

---

## Pipeline Architecture & Stages

The core processing workflow executes in 12 consecutive steps (Steps 2 through 13):

```text
Step 2: Load Raw CSV (dtype=str, utf-8-sig encoding)
  │
Step 3: Drop Duplicate Records
  │
Step 4: Parse Dates, Year & Fiscal Quarter
  │
Step 5: Standardize Property Type Taxonomy
  │
Step 6: Classify Project Type & Asset Class
  │
Step 7: Normalize Unit Configurations & Populate Defaults
  │
Step 8: Master Location Coordinates & Location ID Lookup
  │
Step 9: Master Project Coordinates Matching
  │
Step 10: Commercial Project Geocoding (Google Places / ArcGIS / Location Fallback)
  │
Step 11: Non-RERA Index Assignment (internal_index_id starting from nr468)
  │
Step 12: DB Schema Alignment (68 Columns) & Title Case Formatting
  │
Step 13: Export Final Clean Datasets (CSV & Excel)
```

---

## Repository Structure

```text
├── adrec_pipeline.ipynb                 # Interactive Jupyter Notebook (Steps 2 to 13)
├── adrec_pipeline.py                    # Standalone production Python script
├── Abudhabi_location_coordinates.csv    # Master reference of location IDs & coordinates
├── Abudhabi_project_coordinates.csv     # Master reference of project coordinates & NR indexes
├── AbuDhabi_RawData.csv                 # Raw sample ADREC transaction dataset
├── requirements.txt                     # Python package dependencies
├── .gitignore                           # Git ignore rules for outputs and caches
└── README.md                            # Comprehensive pipeline documentation
```

---

## Taxonomy & Standardization Rules

### Property Type Mapping

| Raw Property Type (Source) | Standardized Output (`property_type`) |
|:---|:---|
| `apartment`, `duplex`, `residential complex`, `residential complex (staff accommodation)` | `Flat` |
| `villa`, `townhouse / attached villa`, `palace` | `Villa` |
| `office`, `office complex` | `Office` |
| `retail`, `mall / market / retail center` | `Shop` |
| All plot types (`other residential plot`, `other commercial plot`, `plot for villa`, etc.) | `Plot` |
| `clinic / hospital`, `factory`, `farm`, `hotel / resort`, `mosque`, `penthouse`, `school`, etc. | `Others` |

### Asset Class & Project Type Mapping

| Raw Asset Class (Source) | Standardized Output (`project_type`) |
|:---|:---|
| `residential` | `Residential` |
| `commercial` | `Commercial` |
| `industrial & storage` | `Industrial` |
| `infrastructural` | `Infrastructure` |
| `agricultural` | `Agricultural` |
| `educational` | `Educational` |
| `healthcare` | `Healthcare` |
| `recreational` | `Recreational` |
| `religious` | `Religious` |
| `other` / Unclassified | `Others` |

### Unit Configuration Normalization

| Raw Layout (`Property Layout`) | Normalized Configuration (`normalized_unit_configuration`) |
|:---|:---|
| `1 bed` | `1Bhk` |
| `2 beds` | `2Bhk` |
| `3 beds` | `3Bhk` |
| `4 beds`, `5 beds`, `5+ beds`, `6+ beds` | `>3Bhk` |
| `studio` | `Flat` |
| `anchor store`, `line store` | `Shop` |
| `small (up to 50 sqm)`, `medium (50 to 500 sqm)`, `large (500 to 1000 sqm)`, `maxi (over 1000 sqm)` | `Office` |
| `unclassified` / None | `Others` |

### Spatial Enrichment & Fallbacks

1. **Location Coordinates**: Mapped from `Abudhabi_location_coordinates.csv`. If district is `jabal hafeet`, fallback coordinates `(24.0716702, 55.7985914)` are automatically applied.
2. **Master Project Coordinates**: Matched on composite key `(normalized_location, normalized_project, normalized_city)`.
3. **Commercial Project Geocoding**: Projects not found in master records are searched using Google Places API and ArcGIS geolocators. If coordinates cannot be determined, commercial projects fall back to their parent district's location coordinates.
4. **Private Projects**: Any missing/unnamed project is labeled `'Private'`. Private project coordinates that are not in master records intentionally remain blank (null), ensuring private residential assets are not erroneously assigned commercial project markers.

### Non-RERA Index Assignment (`internal_index_id`)

- Existing projects in `Abudhabi_project_coordinates.csv` maintain their assigned NR identifier (e.g. `nr105`, `nr210`).
- The pipeline scans the master reference for the highest existing NR index (`max_nr = 467`).
- New projects are assigned sequential identifiers starting from `nr468` onward.
- Grouping is based on the composite key: `(location_name, project_name, city_name)`. Both commercial and `'Private'` units within the same district share the identical `internal_index_id`.

---

## Output Schema (68 Columns)

The output dataset strictly adheres to the 68-column schema below:

| # | Column Name | Type / Notes |
|:---:|:---|:---|
| 0 | `internal_index_id` | Primary project index (`nr...`), position 0 |
| 1 | `project_name` | Name of the project (or `'Private'`) |
| 2 | `location_id` | Master location ID |
| 3 | `location_name` | Standard district name |
| 4 | `registered_document_village_name` | Null placeholder |
| 5 | `agreement_price` | Transaction amount in AED |
| 6 | `transaction_date` | Date formatted as `DD-MM-YYYY` |
| 7 | `net_carpet_area_sq_m` | Sold area in square meters |
| 8 | `sub_registrar_office_code` | Null placeholder |
| 9 | `sub_registrar_office_name` | Null placeholder |
| 10 | `document_number` | Null placeholder |
| 11 | `transaction_type` | Null placeholder |
| 12 | `city_id` | Standard city ID (`1`) |
| 13 | `city_name` | `'Abu Dhabi'` |
| 14 | `guideline_value` | Null placeholder |
| 15 | `property_description` | Strictly `None` (no mapping) |
| 16 | `floor_number` | Null placeholder |
| 17 | `unit_number` | Null placeholder |
| 18 | `property_type_raw` | Raw source property type |
| 19 | `property_type` | Standardized property type |
| 20 | `year` | Transaction year (`Int64`) |
| 21 | `balcony_sq_m` | Null placeholder |
| 22 | `terrace_sq_m` | Null placeholder |
| 23 | `seller_name` | Null placeholder |
| 24 | `buyer_name` | Null placeholder |
| 25 | `transaction_category` | `'Sale'` |
| 26 | `internal_document_number` | Null placeholder |
| 27 | `micr_number` | Null placeholder |
| 28 | `bank_type` | Null placeholder |
| 29 | `party_code` | Null placeholder |
| 30 | `date_of_agreement_execution` | Date formatted as `DD-MM-YYYY` |
| 31 | `stamp_duty_paid` | Null placeholder |
| 32 | `registration_fee` | Null placeholder |
| 33 | `project_latitude` | Latitude coordinate of the project |
| 34 | `project_longitude` | Longitude coordinate of the project |
| 35 | `location_latitude` | Latitude coordinate of the district |
| 36 | `location_longitude` | Longitude coordinate of the district |
| 37 | `quarter` | Fiscal quarter string (e.g. `Q1-2024`) |
| 38 | `unit_configuration` | Raw layout string (lowercase) |
| 39 | `property_layout` | Raw layout string |
| 40 | `buyer_pincode` | Null placeholder |
| 41 | `buyer_locality` | Null placeholder |
| 42 | `buyer_district` | Null placeholder |
| 43 | `buyer_state` | Null placeholder |
| 44 | `is_llm_processed` | `'No'` |
| 45 | `is_manual_processed` | `'No'` |
| 46 | `tower_name` | Null placeholder |
| 47 | `is_duplicate` | Null placeholder |
| 48 | `sale_type` | Sale sequence (`Primary` / `Secondary`) |
| 49 | `project_type` | Standard project type |
| 50 | `country_name` | `'UAE'` |
| 51 | `state_name` | `'Abu Dhabi'` |
| 52 | `micro_market` | Strictly `None` (no mapping) |
| 53 | `sub_locality` | Community name from raw data |
| 54 | `pincode` | Null placeholder |
| 55 | `parking_count` | Null placeholder |
| 56 | `facing_direction` | Null placeholder |
| 57 | `view_type` | Null placeholder |
| 58 | `furnishing_status` | Null placeholder |
| 59 | `condition_status` | Null placeholder |
| 60 | `source_accessibility` | `'Download'` |
| 61 | `source_accessibility_way` | `'Easy'` |
| 62 | `sourcing_cost` | Null placeholder |
| 63 | `sourcing_time` | Null placeholder |
| 64 | `data_type` | `'Registered Document'` |
| 65 | `data_source` | `'Adrec'` |
| 66 | `normalized_unit_configuration` | Standard unit configuration (`1Bhk`, etc.) |
| 67 | `project_stage` | Sale application type (`Ready` / `Off-Plan`) |

> **Note**: `project_id` is strictly purged from the final output, and `internal_index_id` is the primary project identifier located at index 0.

---

## Installation & Setup

### 1. Install Dependencies
```bash
pip install -r requirements.txt
```

---

## Usage Guide

### Option A: Running the Jupyter Notebook
Open [`adrec_pipeline.ipynb`](adrec_pipeline.ipynb) in VS Code or Jupyter Lab:
```bash
jupyter lab adrec_pipeline.ipynb
```
In Cell 3, verify or customize the three input file paths:
```python
RAW_DATA_PATH = "AbuDhabi_RawData.csv"
PROJECT_COORDS_PATH = "Abudhabi_project_coordinates.csv"
LOCATION_COORDS_PATH = "Abudhabi_location_coordinates.csv"
```
Then run the cells sequentially to observe step-by-step transformations and outputs.

### Option B: Running the Python Script via CLI
The standalone script [`adrec_pipeline.py`](adrec_pipeline.py) provides a flexible command-line interface:

```bash
# Run with default files
python adrec_pipeline.py

# Run skipping external geocoding (fast execution)
python adrec_pipeline.py --skip-geocoding

# Specify custom file paths
python adrec_pipeline.py \
  --input custom_raw.csv \
  --project-coords custom_projects.csv \
  --location-coords custom_locations.csv \
  --output-dir ./output
```

#### CLI Options:
| Flag | Description | Default |
|:---|:---|:---|
| `-i`, `--input` | Path to raw ADREC CSV | `AbuDhabi_RawData.csv` |
| `-p`, `--project-coords` | Path to master project coordinates CSV | `Abudhabi_project_coordinates.csv` |
| `-l`, `--location-coords` | Path to master location coordinates CSV | `Abudhabi_location_coordinates.csv` |
| `-o`, `--output-dir` | Directory to save processed files | Input file's directory |
| `--skip-geocoding` | Skip Google Places / ArcGIS online geocoding | `False` |
| `--google-api-key` | Google Places API key | `GOOGLE_MAPS_API_KEY` env var |

### Option C: Using as a Python Module

```python
from adrec_pipeline import run_pipeline

df, csv_path, xlsx_path = run_pipeline(
    raw_file="AbuDhabi_RawData.csv",
    proj_file="Abudhabi_project_coordinates.csv",
    loc_file="Abudhabi_location_coordinates.csv",
    output_dir="./output",
    skip_geocoding=True
)

print(f"Processed {len(df):,} records.")
```

---

## Environment Variables

To provide a Google Places API key for geocoding without passing it via CLI arguments:

```bash
# Windows PowerShell
$env:GOOGLE_MAPS_API_KEY = "your-api-key"

# Linux / macOS
export GOOGLE_MAPS_API_KEY="your-api-key"
```

---

## License

This project is licensed under the MIT License.
