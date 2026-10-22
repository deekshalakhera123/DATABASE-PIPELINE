# IGR Processing Pipeline

This folder contains a local web dashboard and data pipeline for preparing
Indian real-estate registration (IGR) transaction data. It standardizes source
Excel files, enriches transactions with project and location data, produces a
database-ready Excel file, and can convert that output to Parquet for upload.

## Features

- Process raw extraction workbooks through the end-to-end pipeline.
- Resume processing from a manually corrected workbook.
- Normalize project names, transaction categories, property types, and area
  measurements.
- Match projects against configured RERA reference data and enrich buyer and
  location information.
- Save final Excel output and optionally convert it to Parquet.
- Monitor pipeline progress and logs in a browser dashboard.

## Requirements

- Windows with Python 3.10 or later.
- PostgreSQL access for database lookups and index assignment.
- Google Drive for desktop, or local folders containing the input, manual
  correction, and output files. The current Drive-folder resolver expects the
  Drive desktop sync to be available on `G:`.
- The repository's other processing assets, including the configured RERA
  workbook and `postal_pincode.csv`.
- For database upload, the separate `final_code.py` upload script and its
  configuration must also be available in the project.

## Setup

Open PowerShell in this `Processing` folder and create a virtual environment:

```powershell
py -3.10 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install pandas numpy psycopg2-binary requests openpyxl xlsxwriter pyarrow rapidfuzz scikit-learn tqdm
```

Before running the pipeline:

1. Configure the PostgreSQL connection used by the pipeline in
   `pipeline_core.py`. The database must contain the tables and columns queried
   by the processing code, including `public.transactions` and
   `public.dim_location`.
2. Check the city configuration in `city_config.py`. Pune and Mumbai have
   configured Drive folders; Thane's Drive folders are not configured.
3. Make sure the required input and reference files are available to the
   machine running the dashboard. The Drive location picker scans folders
   accessible through Drive for desktop.
4. If using database upload, verify the upload script and database settings
   outside this folder as well.

## Run the dashboard

```powershell
python server.py
```

The server prints the dashboard URL when it starts (usually
`http://localhost:8000`). If that port is busy, it tries the next ports up to
`8009`. Press `Ctrl+C` in the terminal to stop the server.

## Pipeline modes

- **Mode 1 — Start:** Select a city and raw input Excel file to start processing.
  The pipeline may pause for manual review; use the generated/corrected workbook
  to continue.
- **Mode 2 — Resume:** Continue from a manually corrected Excel workbook,
  starting at the enrichment and finalization stages.
- **Mode 3 — Parquet:** Convert an existing final processed Excel file to
  Parquet. The database upload step can be launched separately from the
  dashboard when the upload script is present.

Use the dashboard's location pickers to select files from the configured Drive
folders, or provide a local file path. The pipeline accepts Excel input
workbooks (`.xlsx`; some utilities also support `.xls` and `.csv`).

## Main files

| File or folder | Purpose |
| --- | --- |
| `server.py` | Starts the dashboard server and coordinates pipeline runs. |
| `pipeline_core.py` | Shared processing logic and PostgreSQL connection settings. |
| `city_config.py` | City IDs, Drive folders, and RERA workbook configuration. |
| `project_name_Std_and_area_conversion.py` | Project-name normalization and area conversion. |
| `rera_matching.py` | RERA reference matching helpers. |
| `transaction_categorizer.py` | Transaction-type categorization. |
| `db_columns.py` | Output column order for the database schema. |
| `parquet_conersion.py` | Excel/CSV-to-Parquet conversion utilities. |
| `web/` | Dashboard HTML, JavaScript, and CSS. |

## Notes

- The dashboard binds to `localhost`; it is intended for local use.
- Google Drive URLs in the city configuration are not direct API integrations:
  Drive files must be available through the local Drive sync.
- Project-coordinate enrichment through Google Places is optional and disabled
  by default.
- Database upload depends on the external upload script and its own runtime
  configuration; completing the local processing pipeline does not by itself
  guarantee that records have been uploaded.
