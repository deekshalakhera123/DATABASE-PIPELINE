# Karnataka IGR Data Processing Pipeline

An end-to-end processing pipeline for Karnataka Inspector General of Registration (IGR) / Index-II registered document data. The pipeline extracts structured attributes from mixed Kannada and English property schedules using an LLM, standardizes the schema, enriches records with geospatial coordinates, and formats data into a database-ready schema.

---

## 📋 Table of Contents
- [Pipeline Overview](#-pipeline-overview)
- [Workflow Architecture](#-workflow-architecture)
- [Project Structure](#-project-structure)
- [Prerequisites & Environment Setup](#-prerequisites--environment-setup)
- [Stage 1: LLM Information Extraction (`llm.ipynb`)](#-stage-1-llm-information-extraction-llmipynb)
  - [Overview & Objectives](#overview--objectives)
  - [Key Configuration Parameters](#key-configuration-parameters)
  - [Extracted Schema & Attributes](#extracted-schema--attributes)
  - [Rate Limiting & Chunking Strategy](#rate-limiting--chunking-strategy)
  - [How to Run Stage 1](#how-to-run-stage-1)
  - [Outputs Generated](#outputs-generated)
- [Stage 2: Data Normalization & Enrichment (`Karnataka_Data_Processing.ipynb`)](#-stage-2-data-normalization--enrichment-karnataka_data_processingipynb)
  - [Overview & Objectives](#overview--objectives-1)
  - [Schema Mapping & Transformation](#schema-mapping--transformation)
  - [Unit Configuration & Property Classification](#unit-configuration--property-classification)
  - [Financial & Transaction Extraction](#financial--transaction-extraction)
  - [Geocoding Resolution (Location & Project Coordinates)](#geocoding-resolution-location--project-coordinates)
  - [Index Generation (`nr101` Grouping)](#index-generation-nr101-grouping)
  - [How to Run Stage 2](#how-to-run-stage-2)
  - [Final Output](#final-output)
- [Summary of Execution Order](#-summary-of-execution-order)
- [Troubleshooting & Best Practices](#-troubleshooting--best-practices)

---

## 🌟 Pipeline Overview

The pipeline is executed strictly in two consecutive stages:

```
┌─────────────────────────────────────────┐
│     Raw Scraped IGR Excel / CSV         │
│  (Contains Kannada/English Property     │
│   Schedule & Registration Details)      │
└────────────────────┬────────────────────┘
                     │
                     ▼
┌─────────────────────────────────────────┐
│  STAGE 1: LLM Extraction (llm.ipynb)    │
│  • Bedrock Gemma-3-27b-it model         │
│  • Chunked parallel processing (60/chk) │
│  • Kannada-to-English translation       │
│  • Dimensions, units, bounds extraction │
│  • Token accounting & audit ledger      │
└────────────────────┬────────────────────┘
                     │
                     ▼
┌─────────────────────────────────────────┐
│  Intermediate LLM Merged Excel Output   │
│  (e.g., final_merged_scraped_data_raw_  │
│         llm_processed.xlsx)             │
└────────────────────┬────────────────────┘
                     │
                     ▼
┌─────────────────────────────────────────┐
│  STAGE 2: Data Processing & Geocoding   │
│  (Karnataka_Data_Processing.ipynb)      │
│  • DB column schema re-alignment        │
│  • BHK unit configuration mapping       │
│  • Financial values & transaction types │
│  • Google Maps Location Geocoding       │
│  • Google Places Project Geocoding      │
│  • Project-level index clustering       │
└────────────────────┬────────────────────┘
                     │
                     ▼
┌─────────────────────────────────────────┐
│      Final Database-Ready Excel         │
│  (Aligned to DB_SEQUENCE standard)      │
└─────────────────────────────────────────┘
```

---

## 🏗 Workflow Architecture

```mermaid
flowchart TD
    subgraph Stage1["Stage 1: LLM Extraction (llm.ipynb)"]
        A[Raw Scraped Data: .xlsx / .csv] --> B[Detect Property Column: '(ಎ) ಆಸ್ತಿ ವಿವರ']
        B --> C[AdvancedRateLimiter: 28 RPM, 10 Parallel Workers]
        C --> D[Bedrock API: google.gemma-3-27b-it]
        D --> E[JSON Salvage & Regex Pre-fill]
        E --> F[Chunk Files: extracted_chunk_*.xlsx]
        F --> G[Merge Chunks: final_merged_llm_processed.xlsx]
        F --> H[Merge Token Ledgers: token_ledger_all_chunks.xlsx]
    end

    subgraph Stage2["Stage 2: Standardization & Geocoding (Karnataka_Data_Processing.ipynb)"]
        G --> I[Load Merged LLM File]
        I --> J[Apply COLUMN_MAP & Align to DB_SEQUENCE]
        J --> K[Map BHK Units & Classify Property Subtypes]
        K --> L[Regex Parse: Transaction Type, Consideration, Market Value]
        L --> M[Format Execution Date & Derive Quarter/Year]
        M --> N[Google Geocoding API: Location Latitude & Longitude]
        N --> O[Google Places API: Project Latitude & Longitude]
        O --> P[Fill Project Index: nr101, nr102, ...]
        P --> Q[Title Case Formatting & Final Export]
    end
```

---

## 📁 Project Structure

```
d:/Pankaj/Karnataka/
│
├── llm.ipynb                        # Stage 1: LLM extraction & Kannada translation
├── Karnataka_Data_Processing.ipynb  # Stage 2: Data cleaning, geocoding & DB alignment
├── README.md                        # Documentation & pipeline execution guide
│
├── db_columns.py                    # (Required dependency) Defines DB_SEQUENCE
├── property_type_dict.py            # (Required dependency) classify_property_subtype logic
│
├── extracted_results_karnataka/     # (Created by Stage 1) Chunked extraction files
│   ├── extracted_chunk_1.xlsx
│   ├── extracted_chunk_2.xlsx
│   └── ...
└── token_ledger_all_chunks.xlsx     # (Created by Stage 1) Token usage & cost audit
```

> [!NOTE]
> `db_columns.py` and `property_type_dict.py` are referenced by `Karnataka_Data_Processing.ipynb`. Ensure they are placed in the working directory or included in your `PYTHONPATH`.

---

## ⚙️ Prerequisites & Environment Setup

### 1. Python Environment
Python 3.10+ is recommended. Install required libraries:

```bash
pip install pandas openpyxl openai tqdm requests numpy
```

### 2. API Keys & Environment Variables

| Variable | Notebook | Description | Default / Fallback |
| :--- | :--- | :--- | :--- |
| `BEDROCK_API_KEY` | `llm.ipynb` | AWS Bedrock Mantle API Key for OpenAI-compatible endpoint | Pre-configured in notebook or pass as environment variable |
| `BEDROCK_REGION` | `llm.ipynb` | AWS Region hosting the endpoint | `ap-south-1` |
| `BEDROCK_BASE_URL` | `llm.ipynb` | Base URL for Bedrock OpenAI client | `https://bedrock-mantle.ap-south-1.api.aws/v1` |
| `BEDROCK_MODEL` | `llm.ipynb` | Bedrock model identifier | `google.gemma-3-27b-it` |
| `GOOGLE_MAPS_API_KEY`| `Karnataka_Data_Processing.ipynb` | Google Maps API key for Geocoding & Places Find Place API | Pre-configured in notebook or pass as environment variable |

---

## 🚀 Stage 1: LLM Information Extraction (`llm.ipynb`)

### Overview & Objectives
Raw Karnataka IGR data stores critical property attributes inside unstructured Kannada / English text in column `(ಎ) ಆಸ್ತಿ ವಿವರ` (Property Description). 

Stage 1 extracts and translates these attributes into structured JSON objects per row.

### Key Configuration Parameters
Set the file paths in the `__main__` block of `llm.ipynb`:

```python
INPUT_FILE = r"E:\IGR\Karnataka_IGR\Processing\final_merged_scraped_data_raw.xlsx"
OUTPUT_DIR = "extracted_results_karnataka"
FINAL_OUTPUT = r"E:\IGR\Karnataka_IGR\Processing\final_merged_scraped_data_raw_llm_processed.xlsx"
```

- **`INPUT_FILE`**: Path to the raw scraped Excel or CSV file.
- **`OUTPUT_DIR`**: Directory where individual batch chunk Excel files will be stored.
- **`FINAL_OUTPUT`**: Path where the merged output file will be written.
- **`chunk_size`**: Default `60` records per chunk.
- **`resume`**: Default `True` (enables resuming from the last unprocessed chunk).

### Extracted Schema & Attributes
The LLM extracts over 40 structured fields, including:

1. **General / Translation**:
   - `property_details_en`: Complete English translation of the property description.
2. **Location Hierarchy**:
   - `hobli_en`, `hobli_original`, `village_en`, `village_original`
   - `taluka_en`, `taluka_original`, `district_en`, `district_original`
   - `grama_panchayat_en`, `ward_no`, `locality_layout_en`, `full_address_en`
3. **Property Identification**:
   - `property_type`: Normalized category (`Site`, `Flat/Apartment`, `House`, `Agricultural Land`, `Converted Land`, `Commercial`, `Other`)
   - `property_sub_type`, `land_use`, `unit_configuration` (e.g. `2 BHK`)
   - `property_no`, `site_no`, `plot_no`, `flat_no`, `floor_no`
   - `survey_no`, `old_survey_no`, `new_survey_no`, `hissa_no`
   - `khata_no`, `e_khata_no`, `pid_no`, `katha_extract_no`
   - `project_name_en`, `project_name_original`
4. **Dimensions & Area**:
   - `east_west_dimension`, `north_south_dimension`
   - `area_sqft`, `area_sqm`, `area_acres`, `area_guntas`, `area_as_written`
   - `built_up_area_sqft`, `built_up_area_sqm`
   - Calculated fields: `total_area_sqft_calc`, `total_area_sqm_calc`
5. **Boundaries**:
   - `boundary_east_en`, `boundary_west_en`, `boundary_north_en`, `boundary_south_en`

### Rate Limiting & Chunking Strategy
- **`AdvancedRateLimiter`**: Sliding window of **28 Requests Per Minute (RPM)** with a max of **10 parallel worker threads**.
- **Resumability**: If the notebook or connection stops, re-running `llm.ipynb` automatically detects existing `extracted_chunk_*.xlsx` files and skips them, resuming from the next chunk.
- **Regex Prefill & Fallback**: If the LLM misses deterministic patterns (such as 15–20 digit PID / E-Khata numbers or standard `Hobli: ..., Village: ...` prefixes), a regex prefiller populates them automatically.

### How to Run Stage 1
1. Open `llm.ipynb` in your Jupyter environment.
2. Verify `BEDROCK_API_KEY`, `INPUT_FILE`, and `OUTPUT_DIR`.
3. Run all cells in `llm.ipynb`.
4. Monitor the tqdm progress bar for real-time RPM, token metrics, and request logs.
5. On completion, verify that `FINAL_OUTPUT` and `token_ledger_all_chunks.xlsx` are created.

### Outputs Generated
- **`extracted_results_karnataka/extracted_chunk_{i}.xlsx`**: Intermediate batch files.
- **`final_merged_scraped_data_raw_llm_processed.xlsx`**: Full dataset merged with LLM fields.
- **`token_ledger_all_chunks.xlsx`**: Three-sheet Excel workbook (`All_Token_Details`, `Per_Chunk_Summary`, `Grand_Summary`) providing token consumption and audit details.

---

## 🧹 Stage 2: Data Normalization & Enrichment (`Karnataka_Data_Processing.ipynb`)

### Overview & Objectives
Stage 2 takes the output from Stage 1 (`final_merged_scraped_data_raw_llm_processed.xlsx`), aligns it with the canonical database schema (`DB_SEQUENCE`), derives financial and transaction fields, geocodes locations and projects, and assigns unique project indices.

### Schema Mapping & Transformation
`COLUMN_MAP` bridges raw and LLM field names to database columns:

| Source Column | Target DB Column |
| :--- | :--- |
| `Serial No.` | `index` |
| `project_name_en` | `project_name` |
| `Village` | `location_name`, `registered_document_village_name` |
| `(A) Property Details` | `property_description` |
| `Date of Execution` | `date_of_agreement_execution` |
| `Name of Executant / Party Who Executed the Document` | `seller_name` |
| `Name of Claimant / Party in Whose Favour the Document Was Executed` | `buyer_name` |
| `floor_no` | `floor_number` |
| `flat_no` | `unit_number` |
| `property_sub_type` | `property_type_raw` |
| `total_area_sqm_calc` | `net_carpet_area_sq_m` |

### Unit Configuration & Property Classification
- **`map_unit()`**: Normalizes raw unit descriptions into standardized buckets:
  - `<1Bhk`, `1Bhk`, `1.5Bhk`, `2Bhk`, `2.25Bhk`, `2.5Bhk`, `2.75Bhk`, `3Bhk`, `>3Bhk`
  - Commercial units: `Office`, `Shop`, `Flat`, `Others`
- **`classify_property_subtype()`**: Imported from `property_type_dict.py`, maps `property_type_raw` to standardized property categories.

### Financial & Transaction Extraction
Parses registration details from column `(B) Nature...`:
- **`transaction_type`**: Regex extracts the Article Name (e.g. *Absolute Sale*, *Agreement of Sale*, *Gift Deed*, *Mortgage*).
- **`agreement_price`**: Numeric extraction of `Consideration Amount`.
- **`guideline_value`**: Numeric extraction of `Market Value`.
- **Date & Quarter Formatting**: Execution dates are converted to `DD/MM/YYYY`, with `year` (e.g. `2024`) and `quarter` (e.g. `Q1-2024`) extracted.

### Geocoding Resolution (Location & Project Coordinates)

#### 1. Location Coordinates (`populate_location_coordinates`)
- **API**: Google Maps Geocoding API (`https://maps.googleapis.com/maps/api/geocode/json`)
- **Query**: Combines `location_name` with `Taluka`, `District`, and `Karnataka`.
- **Constraint**: `administrative_area:Karnataka|country:IN`. Ensures coordinates are strictly within Karnataka.
- **Fields Updated**: `location_latitude`, `location_longitude`, `google_location_address`, `google_place_id`.

#### 2. Project Coordinates (`populate_project_coordinates`)
- **API**: Google Maps Places API (`https://maps.googleapis.com/maps/api/place/findplacefromtext/json`)
- **Query**: `project_name, location_name, taluka, district, Karnataka, India`.
- **Fields Updated**: `project_latitude`, `project_longitude`, `google_project_name`, `google_address`, `google_place_id`.

> [!TIP]
> Both geocoding routines deduplicate queries prior to making API calls, minimizing unnecessary network requests and API quota usage.

### Index Generation (`fill_index` Grouping)
- Matches existing indices for identical `(location_name, project_name)` combinations.
- Assigns new project groups sequentially with prefix `nr` starting from `101` (e.g., `nr101`, `nr102`, `nr103`, ...).

### How to Run Stage 2
1. Ensure `db_columns.py` and `property_type_dict.py` are in the folder.
2. In Cell 1 of `Karnataka_Data_Processing.ipynb`, set the input file path to your Stage 1 merged output:
   ```python
   df = pd.read_excel(r"E:\IGR\Karnataka_IGR\Processing\final_merged_scraped_data_raw_llm_processed.xlsx")
   ```
3. Set your desired output path in Cell 9:
   ```python
   output_path = r"E:\IGR\Karnataka_IGR\Processing\final_karnataka_processed_db.xlsx"
   df.to_excel(output_path, index=False)
   ```
4. Run all notebook cells sequentially:
   - Cell 0–2: Imports, column mapping, BHK normalization, financial extraction.
   - Cell 3–4: Location coordinate resolution.
   - Cell 5–6: Project coordinate resolution.
   - Cell 7: Index clustering (`fill_index`).
   - Cell 8: Title casing strings.
   - Cell 9: Exporting final dataset.

### Final Output
A fully cleaned Excel file matching the `DB_SEQUENCE` standard, populated with:
- Standardized addresses and English property translations
- Verified latitude/longitude coordinates for both location and project
- Normalized BHK configurations and property classifications
- Validated agreement and guideline values
- Clustered project identifiers

---

## 🔄 Summary of Execution Order

| Step | Notebook / Task | Input | Output |
| :---: | :--- | :--- | :--- |
| **1** | Run `llm.ipynb` | Raw Scraped Excel / CSV | `extracted_chunk_*.xlsx`<br>`final_merged_scraped_data_raw_llm_processed.xlsx`<br>`token_ledger_all_chunks.xlsx` |
| **2** | Verify Stage 1 Output | `final_merged_scraped_data_raw_llm_processed.xlsx` | Check extraction rates & distribution summary |
| **3** | Configure Stage 2 Paths | Edit input path in Cell 1 of `Karnataka_Data_Processing.ipynb` | Points to Stage 1 merged Excel |
| **4** | Run `Karnataka_Data_Processing.ipynb` | Merged LLM Excel | `final_karnataka_processed_db.xlsx` (Standardized DB-ready Excel) |

---

## 🛠 Troubleshooting & Best Practices

1. **Bedrock API Rate Limits / Timeouts**:
   - The rate limiter in `llm.ipynb` is tuned for **28 RPM** with **10 parallel workers**. If you experience HTTP 429 errors from Bedrock, lower `max_requests` to `20` or reduce `max_parallel` to `5`.
2. **Resuming Interrupted Extractions**:
   - If `llm.ipynb` stops mid-way, keep `resume=True`. Re-executing the script will detect all saved chunk files and continue from the first unfinished chunk.
3. **Missing Module Errors (`db_columns` / `property_type_dict`)**:
   - Make sure `db_columns.py` and `property_type_dict.py` are present in the same directory as `Karnataka_Data_Processing.ipynb`.
4. **Google Maps Geocoding Quotas**:
   - Verify that your `GOOGLE_MAPS_API_KEY` has Geocoding API and Places API enabled in the Google Cloud Console.
   - If geocoding fails or is incomplete, check `coordinate_status` in the lookup frames (`NOT_FOUND`, `NOT_IN_KARNATAKA`, or `REQUEST_DENIED`).
5. **Memory Usage**:
   - For very large datasets (>50k rows), chunk processing preserves memory. During Stage 2, ensure adequate system RAM when reading large Excel files, or convert the intermediate format to Parquet.
