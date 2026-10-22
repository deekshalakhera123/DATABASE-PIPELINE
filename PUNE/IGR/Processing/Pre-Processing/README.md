# IGR Preprocessing and LLM Extraction

This folder contains two Jupyter notebooks for preparing Indian Registration
(IGR) transaction workbooks and extracting structured property details from
their descriptions:

- `igr_preprocessing.ipynb` merges, cleans, filters, and optionally splits
  Excel workbooks. It also includes optional helpers for merging processed
  files and filtering by year.
- `LLM.ipynb` sends property descriptions to an Amazon Bedrock Mantle-compatible
  endpoint using the OpenAI client, saves chunk-level results, and combines
  them into a final workbook.

The notebooks are examples with machine-specific Windows paths. Update their
input and output paths, folder names, and date/year filters for each city or
dataset before running them.

## Requirements

- Windows and Python 3.10 or later
- Jupyter Notebook or JupyterLab
- `pandas`, `openpyxl`, `lxml`, `tqdm`, and `openai`
- `xlrd` if using the optional generic Excel merge with binary `.xls` files
- A Bedrock Mantle API key and access to the configured model to run `LLM.ipynb`

Install the packages in the Python environment used by Jupyter:

```powershell
py -m pip install pandas openpyxl lxml tqdm openai xlrd jupyterlab
```

`xlrd` is only needed for reading traditional binary `.xls` workbooks. The main
merge example treats `.xls` files as UTF-16 encoded HTML tables, which is how
some IGR exports are formatted.

## Workbook preparation

Open `igr_preprocessing.ipynb` and update the paths and settings in each cell
you intend to use. The steps below describe the main workflow; some notebook
cells are alternatives or dataset-specific examples, not one uninterrupted
pipeline.

### 1. Merge raw source files

The “New Structure wise” cell scans the configured year folders recursively
for `.xls` files, reads the first HTML table from each file using UTF-16
encoding, and combines the rows into one workbook. It adds `_source_file` and
`_source_folder` columns to retain the origin of each row.

Set `root`, the `folders` list, and `output_file` to match the dataset. The
older merge cell scans a single folder and writes a separate output; use one
merge approach, not both, unless you specifically need both outputs.

The additional merge cell later in the notebook is an optional alternative
that searches recursively for `.xls` and `.xlsx` files and tries Excel parsing
before HTML parsing.

### 2. Clean and deduplicate

The cleaning cell reads the merged workbook, normalizes column names to
lowercase, keeps the configured IGR columns when they exist, and drops rows
that are blank across the kept columns. It removes rows without a
`propertydescription`, trims that text, and converts it to title case.

Duplicates are identified using the available columns from this set:
`docno`, `docname`, `registrationdate`, `sroname`, `propertydescription`,
`areaname`, `consideration_amt`, and `marketvalue`. The first matching row is
kept. Confirm that the input file name matches the file created by the merge
step, then set the desired cleaned output path.

### 3. Apply an optional date or year filter

The date-filter cell parses `registrationdate` using `%d/%m/%Y` and keeps
records on or after `2025-10-30`. Rows with dates that cannot be parsed are
excluded. Change the threshold and date format to match the data you are
processing.

An optional year-filter cell removes rows where `year == 2024`. Its inline
comment currently says 2020, so verify the actual condition before running it.
These are separate, optional filters; use only the filter appropriate to the
dataset.

### 4. Split for external processing (optional)

The split cell divides the filtered workbook into four nearly equal Excel
files, `part_1.xlsx` through `part_4.xlsx`. Any remainder rows are included in
the fourth part. Change the input path and `num_parts` if needed. The parts are
written to the notebook's current working directory.

This manual split is not required by `LLM.ipynb`: that notebook independently
splits its configured input into chunks of 60 rows.

### 5. Merge externally processed parts (optional)

The processed-file merge cell looks in a configured folder for `.xlsx` files
whose names begin with `part_`, sorts them by the numeric suffix, and
concatenates them into one workbook. Put only the intended processed parts in
that folder and set the final output path. The names must follow the
`part_<number>.xlsx` pattern.

## LLM extraction

`LLM.ipynb` extracts property, building, area, address, transaction, and other
details from the `propertydescription` column. Its prompt requests English and
original-language fields for project and location names. The notebook uses
parallel requests with rate limiting; defaults in the code are 28 requests
per minute and up to 10 parallel requests. Adjust them to the quota for your
Bedrock account and model.

### Configure and run

1. Set the API key in PowerShell before starting Jupyter:

   ```powershell
   $env:BEDROCK_API_KEY = "your-api-key"
   ```

   Do not place API keys in the notebook or commit them to source control.

2. Optionally configure `BEDROCK_REGION`, `BEDROCK_BASE_URL`, and
   `BEDROCK_MODEL` as environment variables. The defaults in the notebook are
   region `ap-south-1`, a region-derived Bedrock Mantle base URL, and model
   `google.gemma-3-27b-it`.

3. In the `if __name__ == "__main__"` configuration near the end of the
   notebook, set:
   - `INPUT_FILE` to the cleaned workbook to process
   - `OUTPUT_DIR` to a directory for chunk results
   - `FINAL_OUTPUT` to the final extracted workbook path

4. Run the notebook's code cell. It processes the input in 60-row chunks,
   writes chunk workbooks and token-usage ledgers, merges chunk results into
   the final workbook, and prints a summary of selected extracted fields.

The extraction code can resume from existing chunk files. Review the console
output and chunk workbooks after interrupted or failed runs; do not assume a
saved chunk contains successful extractions just because the file exists.

## Output overview

| Stage | Typical output |
| --- | --- |
| Raw merge | One merged workbook with source tracking columns |
| Cleaning | Deduplicated workbook with the selected columns |
| Date/year filter | Workbook containing only records that meet the configured filter |
| Optional manual split | `part_1.xlsx`, `part_2.xlsx`, etc. |
| LLM extraction | Chunk workbooks under `OUTPUT_DIR`, a final extracted workbook, and token-usage ledger workbooks |
| Optional external-part merge | One workbook combining processed `part_<number>.xlsx` files |

Use distinct output names or folders for each run to avoid mixing chunks or
overwriting source data.

## Troubleshooting

- **No source files found:** Check the root and year-folder paths and confirm
  the source files use the expected `.xls` extension.
- **HTML table parsing fails:** Verify the source export is UTF-16 encoded
  HTML; install `lxml` in the notebook environment if needed.
- **Missing columns or unexpectedly few rows:** Check the source workbook's
  headers, the configured keep/dedup columns, and the date format and cutoff.
- **Bedrock authentication or throttling errors:** Confirm the environment
  variable is available to the Jupyter process and lower concurrency or
  requests per minute to fit the account quota.
- **Resume skips a chunk:** The extraction pipeline uses existing chunk files
  to resume. Inspect or remove only the specific invalid chunk file before
  rerunning that chunk.
