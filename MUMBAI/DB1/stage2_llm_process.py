# =============================================================================
# DB1 PIPELINE — STAGE 2: LLM Processing + Manual Instruction Placeholders
# =============================================================================
# Input:  Stage 1 output Excel (e.g. "Sample Sale data for llm.xlsx")
# Output: LLM-enriched DataFrame → passed to Stage 3
#
# This stage is split into two phases:
#   2.1  LLM Processing   — send Bhumapan (property description) to LLM, extract fields
#   2.2  Manual Review    — human corrections applied before Stage 3 begins
# =============================================================================

# =============================================================================
# STAGE 2.1 — LLM Extraction
# =============================================================================
# Send each row's `Bhumapan` text to the LLM. Extract:
#
#   | Column               | Description                                              |
#   |----------------------|----------------------------------------------------------|
#   | project_name         | Name of the housing project / building                   |
#   | flat_no              | Flat / unit number as written in the description         |
#   | property_type_raw    | Raw property type string (e.g. "Flat", "Shop", "Office") |
#   | floor_no             | Floor number (word or digit, e.g. "ninth", "9")          |
#   | net_carpet_area_sqmt | Net carpet area in sq. metres (numeric)                  |
#   | location             | Transaction village if different from igr_village        |
#
# Output columns appended to the DataFrame:
#   - All fields above
#   - llm_processed    → "Yes" / "No"  (flag set by LLM processing script)
#   - manual_processed → "Yes" / "No"  (flag set for manually reviewed data)
# =============================================================================

# =============================================================================
# STAGE 2.2 — Manual Review Instructions
# =============================================================================
# After LLM processing, a human analyst reviews the output Excel and:
#   1. Corrects any mis-extracted project_name, flat_no, property_type_raw,
#      floor_no, net_carpet_area_sqmt, or location values.
#   2. Sets manual_processed = "Yes" for every row that has been reviewed,
#      regardless of whether a correction was needed. Rows left as "No"
#      are treated as unreviewed in Stage 3.
#   3. Saves the corrected file back to the same path before running Stage 3.
#
# Stage 3 reads manual_processed to:
#   - Assign non-RERA indices only to manually confirmed rows.
#   - Skip BHK / coordinate logic for unreviewed rows (fail-safe).
# =============================================================================

import json
import os
import random
import re
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

import numpy as np
import pandas as pd
from dotenv import load_dotenv
from tqdm import tqdm

import google.generativeai as genai

# Load .env from the same directory as this script (or any parent dir).
load_dotenv(dotenv_path=Path(__file__).parent / ".env")


# =============================================================================
# CONFIG
# =============================================================================

@dataclass(frozen=True)
class Stage2Config:
    input_excel_path: str = os.getenv("STAGE2_INPUT_EXCEL", r"Sample Sale data for llm.xlsx")
    input_sheet_name: str = os.getenv("STAGE2_INPUT_SHEET", "Sheet1")
    temp_csv_path: str = os.getenv("STAGE2_TEMP_CSV", "temp_bhumapan_data.csv")

    output_chunk_dir: str = os.getenv("STAGE2_CHUNK_DIR", r"extracted_chunks")
    final_output_path: str = os.getenv("STAGE2_FINAL_OUTPUT", r"llm processed Sale data for manual.xlsx")

    model_name: str = os.getenv("STAGE2_MODEL_NAME", "gemma-3-27b-it")
    chunk_size: int = int(os.getenv("STAGE2_CHUNK_SIZE", "50"))
    max_workers: int = int(os.getenv("STAGE2_MAX_WORKERS", "1"))
    resume: bool = os.getenv("STAGE2_RESUME", "True").strip().lower() != "false"
    retries: int = int(os.getenv("STAGE2_RETRIES", "3"))


DEFAULT_CONFIG = Stage2Config()

# Load API key from environment (populated by .env above).
# Never hardcode API keys in source.
_api_key = os.getenv("GOOGLE_API_KEY")
if not _api_key:
    raise RuntimeError(
        "GOOGLE_API_KEY is not set.\n"
        "Add it to MUMBAI/DB1/.env:\n"
        "  GOOGLE_API_KEY=your-key-here\n"
        "or export it in your shell before running this script."
    )
genai.configure(api_key=_api_key)


def get_model(model_name: str = DEFAULT_CONFIG.model_name):
    return genai.GenerativeModel(model_name)


# =============================================================================
# AREA EXTRACTION — deterministic regex helpers
# =============================================================================

NUM = r'(\d{1,3}(?:,\d{2,3})*(?:\.\d+)?|\d+(?:\.\d+)?)'

UNIT = (
    r'(?:चौ\s*\.?\s*मी|चौमी|चौरस\s*मीटर|वर्ग\s*मीटर|वर्गमीटर|'
    r'sq\.?\s*m|sqm|sq\s*meter|square\s*meter|मी|मीटर|'
    r'चौ\s*\.?\s*फ[ुू][ट्त]?|चौफुट|चौरस\s*फुट|वर्ग\s*फुट|वर्गफुट|'
    r'sq\.?\s*ft|sqft|square\s*feet|फ[ुू][ट्त]?|फु)'
)

KW_BUILDUP = r'(?:बिल्ट\.?\s*अप|बिल्टअप|बांधीव|बिल्ट\s*अप|बिल्ट-अप)'
KW_CARPET = r'(?:कार्पेट|कारपेट|चटई|रेरा कार्पेट)'
KW_SALEABLE = r'(?:saleable|सेलेबल|सेलएबल|सेलेब(?:ल)?|सेलेबल\s*एरिया)'
KW_TERRACE = r'(?:टेरेस|टेरस|terrace|टेरेस\s*एरिया|टेरस\s*एरिया)'
KW_BALCONY = r'(?:बाल्कनी|बालकनी|balcony|बाल्कनी\s*एरिया|बालकनी\s*एरिया)'
KW_OTHER = r'(?:एरिया|क्षेत्र|area|एरिया\s*क्षेत्र)'


def _make_patterns(kw: str):
    """
    Two-direction patterns:
      1) keyword ... value unit
      2) value unit ... keyword
    Allow up to 20 non-digit chars between.
    """
    p1 = re.compile(rf'{kw}[^\d]{{0,20}}{NUM}\s*({UNIT})', re.IGNORECASE)
    p2 = re.compile(rf'{NUM}\s*({UNIT})[^\d]{{0,20}}{kw}', re.IGNORECASE)
    return p1, p2


PAT_BUILDUP = _make_patterns(KW_BUILDUP)
PAT_CARPET = _make_patterns(KW_CARPET)
PAT_SALE = _make_patterns(KW_SALEABLE)
PAT_TERRACE = _make_patterns(KW_TERRACE)
PAT_BALCONY = _make_patterns(KW_BALCONY)
PAT_OTHER = _make_patterns(KW_OTHER)


def find_area_raw(text: str, pats) -> Optional[str]:
    """Return e.g. '58.96 चौ.मी' if found, else None."""
    if not isinstance(text, str):
        return None
    for pat in pats:
        m = pat.search(text)
        if m:
            value = m.group(1)
            unit = m.group(2)
            unit = re.sub(r'\s*\.\s*', '.', unit.strip())
            unit = re.sub(r'\s+', ' ', unit)
            return f"{value} {unit}".strip()
    return None


# =============================================================================
# UNIT CONVERSION — sqft/sqmt normalisation (used in final merge)
# =============================================================================

SQMT_REGEX = (
    r"(चोरस मीटर|चो. मी.|स्क्वे मीटर|स्के. मीटर.|चो.मि.|sq मीटर्स|स्के मी|चो. मी.|"
    r"चौ\.?\s*मी|चौ\.?\s*मीटर|चौरस?\s*मीटर|चौ\.?\s*मि.|स्क्वेअर मीटर|चौ . मी|चौ .मी|"
    r"चो मी|स्क्वेर मीटर|चो. मीटर|स्केव्यर मीटर|चौरस मिटर|चो.मीटर|चौ\.?\sमि|चो?\s*मिटर|"
    r"चो?\s*मीटर|चौ\.?\sमि|चौ?\s*मिटर|चौरस?\s*मी|चौ\.?\s*मि.|स्केवर?\s*मीटर|चौ .?\s*मीटर|"
    r"स्क्वायर?\s*मीटर|sqm|sqmt|square\s*meter)"
)
FOOT_REGEX = (
    r"(फुट|फु|फीट|फू|फिट|फ़ूट|फ़ुट|फ़ु|फ़ूट|फ़ुट|फूट|फ़ूत|स्क.फ्ट|रेरा कारपेट|रेरा कार्पेट|"
    r"चौ फट|चो फी|चौरस फ़ीट)"
)
SQFT_TO_SQMT = 10.764


def normalize_decimal_spacing(text) -> str:
    """Fix cases like '193. 122' → '193.122'."""
    if not isinstance(text, str):
        return text
    return re.sub(r"(\d+)\.\s+(\d+)", r"\1.\2", text)


def carpet_to_sqmt(val) -> float:
    """Convert a raw area string (sqmt or sqft, any script) to a sqmt float."""
    if not isinstance(val, str):
        return np.nan

    val = normalize_decimal_spacing(val.replace(",", ""))

    m_sqmt = re.search(rf"(\d+(?:\.\d+)?)[^\d]*{SQMT_REGEX}", val, flags=re.I)
    if m_sqmt:
        return float(m_sqmt.group(1))

    m_sqft = re.search(rf"(\d+(?:\.\d+)?)[^\d]*{FOOT_REGEX}", val, flags=re.I)
    if m_sqft:
        return float(m_sqft.group(1)) / SQFT_TO_SQMT

    return np.nan


# =============================================================================
# LLM-AIDED EXTRACTION
# =============================================================================

EXTRACTION_KEYS = [
    "project_name_en", "Block No", "Carpet_Area", "BuildUp_Area",
    "Saleable_Area", "Terrace_Area", "Balcony_Area", "Other_Area",
]

_EMPTY_EXTRACTION = {k: None for k in EXTRACTION_KEYS}


def _build_prompt(text: str, block_no: Optional[str]) -> str:
    return f"""
You are a real estate document parser. The given input is in Marathi.

1. Extract all project names from the text, translate them to English,
   and list them as a comma-separated string in the key project_name_en.

2. Take the given Block_No value and translate it to English, keeping numbers unchanged,
   and return it in the key Block No.

3. Extract the Carpet Area, Build-Up Area, Saleable Area, Terrace Area, Balcony Area, and Other Area values from the text.
   If available, return them with their units (e.g., "58.96 sq.m" or "630 sq.ft").
   If not present, return null.

Return ONLY a valid JSON object with exactly these eight keys:
- project_name_en
- Block No
- Carpet_Area
- BuildUp_Area
- Saleable_Area
- Terrace_Area
- Balcony_Area
- Other_Area

The Block_No value is: "{block_no if block_no else ''}"

Text:
\"\"\"{text or ''}\"\"\""""


def _parse_llm_json(raw_text: str) -> dict:
    content = (raw_text or "").strip()
    if content.startswith("```"):
        content = re.sub(r"^```(?:json)?\s*", "", content)
        content = re.sub(r"\s*```$", "", content)
    return json.loads(content)


def extract_using_gemini(text: str, model, retries: int = DEFAULT_CONFIG.retries) -> dict:
    """Extract project/block/area fields from one Bhumapan text via regex + Gemini."""
    block_no_match = re.search(r"ब्लॉक\s*नं[:\s]+([^\n,]+)", text or "")
    block_no = block_no_match.group(1).strip() if block_no_match else None

    # Deterministic area extraction (regex) runs regardless of LLM outcome.
    build_raw = find_area_raw(text, PAT_BUILDUP)
    carpet_raw = find_area_raw(text, PAT_CARPET)
    sale_raw = find_area_raw(text, PAT_SALE)
    terrace_raw = find_area_raw(text, PAT_TERRACE)
    balcony_raw = find_area_raw(text, PAT_BALCONY)

    other_raw = None
    if not any([build_raw, carpet_raw, sale_raw, terrace_raw, balcony_raw]):
        other_raw = find_area_raw(text, PAT_OTHER)

    prompt = _build_prompt(text, block_no)

    project_name_en = None
    block_no_out = block_no

    for attempt in range(1, retries + 1):
        try:
            response = model.generate_content(prompt)
            data = _parse_llm_json(response.text)

            if data.get("project_name_en"):
                names = [n.strip() for n in str(data["project_name_en"]).split(",") if n.strip()]
                project_name_en = ",".join(names) if names else None

            block_no_out = data.get("Block No", block_no)
            break

        except Exception as e:
            err_text = str(e).lower()
            if "quota" in err_text or "rate limit" in err_text:
                wait_time = attempt * 30 + random.randint(5, 15)
                print(f"[Retry {attempt}] Rate limit hit. Waiting {wait_time} seconds...")
                time.sleep(wait_time)
                continue
            print(f"[Error] Gemini failed: {e}")
            break

    return {
        "project_name_en": project_name_en,
        "Block No": block_no_out,
        "Carpet_Area": carpet_raw,
        "BuildUp_Area": build_raw,
        "Saleable_Area": sale_raw,
        "Terrace_Area": terrace_raw,
        "Balcony_Area": balcony_raw,
        "Other_Area": other_raw,
    }


# =============================================================================
# PARALLEL EXTRACT (per chunk)
# =============================================================================

def parallel_extract(df_chunk: pd.DataFrame, model, max_workers: int = 2) -> pd.DataFrame:
    """Run extract_using_gemini across a chunk's rows concurrently, order-preserving."""
    results = [None] * len(df_chunk)

    def process_row(idx: int, row: pd.Series) -> dict:
        text = row.get("Bhumapan", "")
        if not isinstance(text, str):
            text = ""
        return extract_using_gemini(text, model)

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {executor.submit(process_row, idx, row): idx for idx, row in df_chunk.iterrows()}
        for future in tqdm(as_completed(futures), total=len(futures), desc="Processing Chunk"):
            idx = futures[future]
            try:
                results[idx] = future.result()
            except Exception as e:
                print(f"[Row {idx}] {e}")
                results[idx] = dict(_EMPTY_EXTRACTION)

    return pd.DataFrame(results, columns=EXTRACTION_KEYS)


# =============================================================================
# RESET / RESUME HELPERS
# =============================================================================

CHUNK_FILENAME_RE = re.compile(r"extracted_chunk_(\d+)\.xlsx$", re.IGNORECASE)


def reset_output_directory(output_dir: str) -> None:
    """Remove .xlsx chunk files (and stale temp/lock files) to force a clean re-run."""
    p = Path(output_dir)
    if not p.exists():
        return
    for file in p.iterdir():
        if file.is_file() and file.suffix.lower() == ".xlsx":
            file.unlink()
    for stale in list(p.glob("extracted_chunk_*.tmp.xlsx")) + list(p.glob("extracted_chunk_*.xlsx.lock")):
        try:
            stale.unlink()
        except OSError:
            pass
    print(f"Cleared existing chunk files from {output_dir}")


def highest_completed_chunk(output_dir: str) -> int:
    """Return highest chunk index present in output_dir; 0 if none."""
    p = Path(output_dir)
    if not p.exists():
        return 0
    max_idx = 0
    for f in p.iterdir():
        if f.is_file():
            m = CHUNK_FILENAME_RE.match(f.name)
            if m:
                max_idx = max(max_idx, int(m.group(1)))
    return max_idx


# =============================================================================
# CHUNK & RUN (with resume, atomic .tmp.xlsx writes)
# =============================================================================

def run_parallel_extraction(
    input_csv_path: str,
    output_dir: str,
    model,
    chunk_size: int = DEFAULT_CONFIG.chunk_size,
    max_workers: int = DEFAULT_CONFIG.max_workers,
    resume: bool = DEFAULT_CONFIG.resume,
) -> None:
    """
    If resume=True, continue from the next chunk after the highest completed one.
    If resume=False, clears output_dir and starts again from chunk 1.
    """
    Path(output_dir).mkdir(parents=True, exist_ok=True)

    if not resume:
        reset_output_directory(output_dir)
        start_from_chunk = 1
    else:
        start_from_chunk = highest_completed_chunk(output_dir) + 1

    df = pd.read_csv(input_csv_path)
    if "Bhumapan" not in df.columns:
        raise ValueError(f"'Bhumapan' column not found. Available columns: {list(df.columns)}")

    total_chunks = (len(df) + chunk_size - 1) // chunk_size
    if start_from_chunk > total_chunks:
        print("Nothing to do. All chunks already processed.")
        return

    print(f"Resume mode: {resume}. Starting at chunk {start_from_chunk} of {total_chunks}.")

    for i in range(start_from_chunk, total_chunks + 1):
        start, end = (i - 1) * chunk_size, min(i * chunk_size, len(df))
        chunk = df.iloc[start:end].reset_index(drop=True)

        final_path = os.path.join(output_dir, f"extracted_chunk_{i}.xlsx")
        tmp_path = os.path.join(output_dir, f"extracted_chunk_{i}.tmp.xlsx")

        if os.path.exists(final_path):
            print(f"Skipping chunk {i} (already processed)")
            continue

        for stale in (tmp_path, os.path.join(output_dir, f"extracted_chunk_{i}.xlsx.lock")):
            if os.path.exists(stale):
                try:
                    os.remove(stale)
                except OSError:
                    pass

        print(f"\nProcessing chunk {i}/{total_chunks} [{start}:{end}]...")

        result_df = parallel_extract(chunk, model, max_workers=max_workers)
        full_df = pd.concat([chunk, result_df], axis=1)

        # Write safely via .tmp.xlsx, then atomic rename.
        full_df.to_excel(tmp_path, index=False, engine="openpyxl")
        os.replace(tmp_path, final_path)

    print("\nAll chunks processed and saved.")


# =============================================================================
# FINAL MERGE (natural chunk order + area normalisation)
# =============================================================================

AREA_COLUMNS = ["Carpet_Area", "BuildUp_Area", "Saleable_Area", "Terrace_Area", "Balcony_Area", "Other_Area"]


def merge_all_chunks(output_dir: str, final_output_path: str) -> None:
    p = Path(output_dir)
    files = [f for f in p.iterdir() if f.is_file() and CHUNK_FILENAME_RE.match(f.name)]
    if not files:
        raise RuntimeError(f"No .xlsx chunk files found in {output_dir}")

    files_sorted = sorted(files, key=lambda f: int(CHUNK_FILENAME_RE.match(f.name).group(1)))
    final_df = pd.concat([pd.read_excel(f) for f in files_sorted], ignore_index=True)

    for col in AREA_COLUMNS:
        final_df[f"{col}_sqmt"] = final_df[col].apply(carpet_to_sqmt)

    Path(final_output_path).parent.mkdir(parents=True, exist_ok=True)
    final_df.to_excel(final_output_path, index=False)
    print(f"Final merged Excel saved to: {final_output_path}")


# =============================================================================
# MAIN
# =============================================================================

def run_stage2(config: Stage2Config = DEFAULT_CONFIG) -> None:
    model = get_model(config.model_name)

    # Prepare temp CSV from Excel.
    df_src = pd.read_excel(config.input_excel_path, sheet_name=config.input_sheet_name)
    df_src.to_csv(config.temp_csv_path, index=False)

    run_parallel_extraction(
        input_csv_path=config.temp_csv_path,
        output_dir=config.output_chunk_dir,
        model=model,
        chunk_size=config.chunk_size,
        max_workers=config.max_workers,
        resume=config.resume,
    )

    merge_all_chunks(config.output_chunk_dir, config.final_output_path)


if __name__ == "__main__":
    run_stage2()