# =============================================================================
# DB1 PIPELINE — STAGE 2: LLM Processing + Manual Instruction Placeholders
# =============================================================================
# Input:  Stage 1 output Excel (Andheri data for llm.xlsx)
# Output: LLM-enriched DataFrame → passed to Stage 3
#
# This stage is split into two phases:
#   2.1  LLM Processing   — send property_description to LLM, extract fields
#   2.2  Manual Review    — human corrections applied before Stage 3 begins
# =============================================================================


# =============================================================================
# STAGE 2.1 — LLM Extraction
# =============================================================================
# Send each row's `property_description` to the LLM.
# Extract the following fields:
#
#   | Column               | Description                                              |
#   |----------------------|----------------------------------------------------------|
#   | project_name         | Name of the housing project / building                   |
#   | flat_no              | Flat / unit number as written in the description         |
#   | property_type_raw    | Raw property type string (e.g. "Flat", "Shop", "Office") |
#   | floor_no             | Floor number (word or digit, e.g. "ninth", "9")          |
#   | net_carpet_area_sqmt | Net carpet area in sq. metres (numeric)                  |
#   | location             | Transaction village if different from igr_village        |
#                            (i.e. the village mentioned inside property_description,
#                             when it differs from the registered areaname)
#
# Output columns appended to the DataFrame:
#   - All fields above
#   - llm_processed  →  "Yes" / "No"  (flag set by LLM processing script)
#   - manual_processed →  "Yes" / "No" (flag set for manual processed data)
# =============================================================================


# =============================================================================
# STAGE 2.2 — Manual Review Instructions
# =============================================================================
# After LLM processing, a human analyst reviews the output Excel and:
#
#   1. Corrects any mis-extracted project_name, flat_no, property_type_raw,
#      floor_no, net_carpet_area_sqmt, or location values.
#
#   2. Sets  manual_processed = "Yes"  for every row that has been reviewed,
#      regardless of whether a correction was needed.
#      Rows left as "No" will be treated as unreviewed in Stage 3.
#
#   3. Saves the corrected file back to the same path before running Stage 3.
#
# The Stage 3 script reads  manual_processed  to:
#   - Assign non-RERA indices only to manually confirmed rows.
#   - Skip BHK / coordinate logic for unreviewed rows (fail-safe).
# =============================================================================
import os
import re
import json
import time
import random
import glob
from pathlib import Path
from typing import Optional

import pandas as pd
from tqdm import tqdm
from concurrent.futures import ThreadPoolExecutor, as_completed

# Optional: pip installa google-generativeai openpyxl
import google.generativeai as genai

# =========================
# ✅ SETUP GEMINI CLIENT
# =========================
# For security, prefer: genai.configure(api_key=os.getenv("GOOGLE_API_KEY"))
genai.configure(api_key="AIzaSyDKr3tuhYMj3nOKyoUnU4SI4fwATHaQ7lI")  # your original line

# Use a model available in this SDK; keep yours or switch to "gemini-1.5-flash"
model = genai.GenerativeModel("gemma-3-27b-it")  # or "gemini-1.5-flash"

# =========================
# ✅ AREA EXTRACTION HELPERS
# =========================
NUM = r'(\d{1,3}(?:,\d{2,3})*(?:\.\d+)?|\d+(?:\.\d+)?)'

UNIT = (
    r'(?:चौ\s*\.?\s*मी|चौमी|चौरस\s*मीटर|वर्ग\s*मीटर|वर्गमीटर|'
    r'sq\.?\s*m|sqm|sq\s*meter|square\s*meter|मी|मीटर|'
    r'चौ\s*\.?\s*फ[ुू][ट्त]?|चौफुट|चौरस\s*फुट|वर्ग\s*फुट|वर्गफुट|'
    r'sq\.?\s*ft|sqft|square\s*feet|फ[ुू][ट्त]?|फु)'
)

KW_BUILDUP  = r'(?:बिल्ट\.?\s*अप|बिल्टअप|बांधीव|बिल्ट\s*अप|बिल्ट-अप)'
KW_CARPET   = r'(?:कार्पेट|कारपेट|चटई|रेरा कार्पेट)'
KW_SALEABLE = r'(?:saleable|सेलेबल|सेलएबल|सेलेब(?:ल)?|सेलेबल\s*एरिया)'
KW_TERRACE  = r'(?:टेरेस|टेरस|terrace|टेरेस\s*एरिया|टेरस\s*एरिया)'
KW_BALCONY  = r'(?:बाल्कनी|बालकनी|balcony|बाल्कनी\s*एरिया|बालकनी\s*एरिया)'
KW_OTHER    = r'(?:एरिया|क्षेत्र|area|एरिया\s*क्षेत्र)'

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
PAT_CARPET  = _make_patterns(KW_CARPET)
PAT_SALE    = _make_patterns(KW_SALEABLE)
PAT_TERRACE = _make_patterns(KW_TERRACE)
PAT_BALCONY = _make_patterns(KW_BALCONY)
PAT_OTHER   = _make_patterns(KW_OTHER)

def _find_area_raw(text: str, pats) -> Optional[str]:
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

# =========================
# ✅ LLM-AIDED EXTRACTION
# =========================
def extract_using_gemini(text, retries=3):
    # Block No via regex first
    block_no_match = re.search(r"ब्लॉक\s*नं[:\s]+([^\n,]+)", text or "")
    block_no = block_no_match.group(1).strip() if block_no_match else None

    # Deterministic area extraction
    build_raw   = _find_area_raw(text, PAT_BUILDUP)
    carpet_raw  = _find_area_raw(text, PAT_CARPET)
    sale_raw    = _find_area_raw(text, PAT_SALE)
    terrace_raw = _find_area_raw(text, PAT_TERRACE)
    balcony_raw = _find_area_raw(text, PAT_BALCONY)

    other_raw = None
    if not any([build_raw, carpet_raw, sale_raw, terrace_raw, balcony_raw]):
        other_raw = _find_area_raw(text, PAT_OTHER)

    prompt = f"""
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

    project_name_en = None
    block_no_out = block_no

    for attempt in range(1, retries + 1):
        try:
            response = model.generate_content(prompt)
            content = (response.text or "").strip()

            # Strip possible code fences
            if content.startswith("```"):
                content = re.sub(r"^```(?:json)?\s*", "", content)
                content = re.sub(r"\s*```$", "", content)

            data = json.loads(content)

            if data.get("project_name_en"):
                project_names = [n.strip() for n in str(data["project_name_en"]).split(",") if n.strip()]
                project_name_en = ",".join(project_names) if project_names else None

            block_no_out = data.get("Block No", block_no)
            break

        except Exception as e:
            err_text = str(e).lower()
            if "quota" in err_text or "rate limit" in err_text:
                wait_time = attempt * 30 + random.randint(5, 15)
                print(f"[Retry {attempt}] Rate limit hit. Waiting {wait_time} seconds...")
                time.sleep(wait_time)
                continue
            else:
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

# =========================
# ✅ PARALLEL EXTRACT
# =========================
def parallel_extract(df_chunk: pd.DataFrame, max_workers=2) -> pd.DataFrame:
    results = [None] * len(df_chunk)

    def process_indexed_row(idx, row):
        text = row.get("Bhumapan", "")
        if not isinstance(text, str):
            text = ""
        return extract_using_gemini(text)

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {executor.submit(process_indexed_row, idx, row): idx for idx, row in df_chunk.iterrows()}
        for future in tqdm(as_completed(futures), total=len(futures), desc="Processing Chunk"):
            idx = futures[future]
            try:
                result = future.result()
                results[idx] = result
            except Exception as e:
                print(f"[Row {idx}] {e}")
                results[idx] = {
                    "project_name_en": None,
                    "Block No": None,
                    "Carpet_Area": None,
                    "BuildUp_Area": None,
                    "Saleable_Area": None,
                    "Terrace_Area": None,
                    "Balcony_Area": None,
                    "Other_Area": None,
                }

    return pd.DataFrame(results, columns=[
        "project_name_en", "Block No", "Carpet_Area", "BuildUp_Area",
        "Saleable_Area", "Terrace_Area", "Balcony_Area", "Other_Area"
    ])

# =========================
# ✅ RESET / RESUME HELPERS
# =========================
CHUNK_FILENAME_RE = re.compile(r"extracted_chunk_(\d+)\.xlsx$", re.IGNORECASE)

def reset_output_directory(output_dir: str):
    """Remove .xlsx chunk files to force re-run from beginning."""
    p = Path(output_dir)
    if p.exists():
        for file in p.iterdir():
            if file.is_file() and file.suffix.lower() == ".xlsx":
                file.unlink()
        # Also clear any stale temp files
        for stale in p.glob("extracted_chunk_*.tmp.xlsx"):
            try: stale.unlink()
            except: pass
        for stale in p.glob("extracted_chunk_*.xlsx.lock"):
            try: stale.unlink()
            except: pass
        print(f"Cleared existing chunk files from {output_dir}")

def _highest_completed_chunk(output_dir: str) -> int:
    """Return highest chunk index present in output_dir; 0 if none."""
    p = Path(output_dir)
    if not p.exists():
        return 0
    max_idx = 0
    for f in p.iterdir():
        if f.is_file():
            m = CHUNK_FILENAME_RE.match(f.name)
            if m:
                idx = int(m.group(1))
                if idx > max_idx:
                    max_idx = idx
    return max_idx

# =========================
# ✅ CHUNK & RUN (WITH RESUME AND .tmp.xlsx)
# =========================
def run_parallel_extraction(input_csv_path, output_dir, chunk_size=50, max_workers=2, resume=True):
    """
    If resume=True, continue from the next chunk after the highest completed one.
    If resume=False, clears output_dir and starts again from chunk 1.
    """
    Path(output_dir).mkdir(parents=True, exist_ok=True)

    if not resume:
        reset_output_directory(output_dir)
        start_from_chunk = 1
    else:
        start_from_chunk = _highest_completed_chunk(output_dir) + 1

    df = pd.read_csv(input_csv_path)
    if "Bhumapan" not in df.columns:
        raise ValueError(f"'Bhumapan' column not found. Available columns: {list(df.columns)}")

    total_chunks = (len(df) + chunk_size - 1) // chunk_size
    if start_from_chunk > total_chunks:
        print("Nothing to do. All chunks already processed.")
        return

    print(f"Resume mode: {resume}. Starting at chunk {start_from_chunk} of {total_chunks}.")

    for i in range(start_from_chunk, total_chunks + 1):
        start = (i - 1) * chunk_size
        end = min(i * chunk_size, len(df))
        chunk = df.iloc[start:end].reset_index(drop=True)

        final_path = os.path.join(output_dir, f"extracted_chunk_{i}.xlsx")
        tmp_path   = os.path.join(output_dir, f"extracted_chunk_{i}.tmp.xlsx")

        if os.path.exists(final_path):
            print(f"Skipping chunk {i} (already processed)")
            continue

        # Clean stale temp/lock files for this chunk
        for stale in (tmp_path, os.path.join(output_dir, f"extracted_chunk_{i}.xlsx.lock")):
            if os.path.exists(stale):
                try: os.remove(stale)
                except: pass

        print(f"\nProcessing chunk {i}/{total_chunks} [{start}:{end}]...")

        result_df = parallel_extract(chunk, max_workers=max_workers)
        full_df = pd.concat([chunk, result_df], axis=1)

        # Write safely via .tmp.xlsx, then atomic rename
        full_df.to_excel(tmp_path, index=False, engine="openpyxl")
        os.replace(tmp_path, final_path)

    print("\nAll chunks processed and saved.")

# =========================
# ✅ FINAL MERGE (NATURAL ORDER)
# =========================
def merge_all_chunks(output_dir, final_output_path):
    p = Path(output_dir)
    files = [f for f in p.iterdir() if f.is_file() and CHUNK_FILENAME_RE.match(f.name)]
    if not files:
        raise RuntimeError(f"No .xlsx chunk files found in {output_dir}")

    files_sorted = sorted(files, key=lambda f: int(CHUNK_FILENAME_RE.match(f.name).group(1)))

    dfs = [pd.read_excel(f) for f in files_sorted]
    final_df = pd.concat(dfs, ignore_index=True)
    Path(final_output_path).parent.mkdir(parents=True, exist_ok=True)
    final_df.to_excel(final_output_path, index=False)
    print(f"Final merged Excel saved to: {final_output_path}")

# =========================
# ✅ MAIN
# =========================
if __name__ == "__main__":
    # Inputs
    INPUT_EXCEL_PATH = r"Andheri data for llm.xlsx"
    INPUT_SHEET_NAME = "Sheet1"
    TEMP_CSV_PATH    = "temp_bhumapan_data.csv"

    OUTPUT_CHUNK_DIR   = r"extracted_chunks"
    FINAL_OUTPUT_PATH  = r"Batch1_LLM_Complete.xlsx"

    # Prepare temp CSV from Excel
    df_src = pd.read_excel(INPUT_EXCEL_PATH, sheet_name=INPUT_SHEET_NAME)
    df_src.to_csv(TEMP_CSV_PATH, index=False)

    # Resume from last completed chunk (set resume=False to restart)
    run_parallel_extraction(
        input_csv_path=TEMP_CSV_PATH,
        output_dir=OUTPUT_CHUNK_DIR,
        chunk_size=50,
        max_workers=1,
        resume=True
    )

    # Merge whatever chunks exist
    merge_all_chunks(OUTPUT_CHUNK_DIR, FINAL_OUTPUT_PATH)
