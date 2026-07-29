# DB1 Pipeline

Converts raw Maharashtra IGR (Inspector General of Registration) transaction
data into a cleaned, enriched dataset ready for DB1 / property valuation use.

Covers **Mumbai, Thane, and Pune**.

For a full breakdown of every function, config field, and known caveat, see
[`DB1_Pipeline_Documentation.md`](./DB1_Pipeline_Documentation.md). This
README is just the quick-start.

## Pipeline overview

```
Raw IGR Excel → Stage 1 → Stage 2.1 (LLM) → Stage 2.2 (manual review) → Stage 3 → Final Excel
```

| Stage | Script | What it does |
|---|---|---|
| 1 | `db1_pipeline_stage1.py` | Load, clean, dedupe, categorise (Sale/Lease/Other), map villages, parse floor/unit/property type from descriptions |
| 2.1 | `db1_pipeline_stage2.py` | LLM (Gemini) + regex extraction of project name, block no, and area fields; chunked & resumable |
| 2.2 | — (manual) | Human review of Stage 2.1 output; sets `manual_processed = "Yes"` per row |
| 3 | `db1_pipeline_stage3.py` | RERA index matching, geocoding, property type + BHK classification, buyer pincode lookup, final assembly |

## Requirements

- Python 3.9+
- `pandas`, `numpy`, `openpyxl`
- `google-generativeai` (Stage 2)
- `geopy` (Stage 3 — ArcGIS geocoding)
- `joblib`, `tqdm`

```bash
pip install pandas numpy openpyxl google-generativeai geopy joblib tqdm
```

## Setup

Stage 2 requires a Gemini API key set as an environment variable — **never
hardcode it in source**:

```bash
export GOOGLE_API_KEY=your-key-here
```

Each stage also expects a few input files (village directory, RERA data,
coordinates cache, postal pincode CSV — see the full documentation for
exact paths and formats). Update the relevant `StageNConfig` dataclass at
the top of each script, or pass a config object in directly.

## Running

```bash
# Stage 1 — check registration_date_mdy / execution_date_mdy match this file's date format
python db1_pipeline_stage1.py

# Stage 2.1
python db1_pipeline_stage2.py

# Stage 2.2 — manual step: review "llm processed Sale data for manual.xlsx",
# correct extracted fields, set manual_processed = "Yes", save.

# Stage 3 — set city in Stage3Config first (mumbai / thane / pune)
python db1_pipeline_stage3.py
```

Each stage accepts a config override instead of editing the script directly:

```python
from db1_pipeline_stage3 import run_stage3, Stage3Config
from pathlib import Path

run_stage3(Stage3Config(city="pune", input_path=Path("pune_reviewed.xlsx")))
```

## Output

Stage 3 produces the final DB1-ready Excel (default:
`sample_file_for_db1.xlsx`) with one row per transaction, including project
name, RERA index (or a synthetic non-RERA index), location, carpet area,
BHK, buyer locality, and coordinates.

## Known gotchas

- Date format (MM/DD vs DD/MM) isn't auto-detected in Stage 1 — set it per
  input file.
- Saleable→Carpet area conversion is city-specific (Mumbai 1.45 / Thane 1.4
  / Pune 1.35); Built-Up→Carpet is fixed at 1.2. Wrong `city` in
  `Stage3Config` silently applies the wrong factor.
- Synthetic non-RERA index codes (`<city>NR101`, ...) are regenerated per
  run and aren't guaranteed stable across runs.

See [`DB1_Pipeline_Documentation.md`](./DB1_Pipeline_Documentation.md) for
details on all of the above plus the full column lineage and per-function
breakdown.