#!/usr/bin/env python
# coding: utf-8

# ### Import Required Libraries

# In[1]:


import pandas as pd
import os
import glob
from datetime import datetime
import re
import numpy as np
import ast


# =====================================================
# SINGLE FILE TELANGANA DB1 PIPELINE
# Generated from: Telangana_DB1_Pipeline.ipynb
# Purpose: Process ONE Excel file through the complete DB1 pipeline.
#
# How to run:
#   python telangana_single_file_pipeline.py --input-file "C:\path\input.xlsx" --rera-file "C:\path\telangana_rera_grand_excel.xlsx"
#
# Optional:
#   python telangana_single_file_pipeline.py --input-file "C:\path\input.xlsx" --rera-file "C:\path\rera.xlsx" --output-dir "C:\path\output" --enable-geocoding
# =====================================================

import argparse
from pathlib import Path


def parse_args():
    parser = argparse.ArgumentParser(description="Process one Telangana DB1 Excel file through all pipeline steps.")
    parser.add_argument("--input-file", required=True, help="Path of the single input Excel file to process.")
    parser.add_argument("--rera-file", required=True, help="Path of Telangana RERA master Excel file used for project matching.")
    parser.add_argument("--output-dir", default="telangana_single_file_output", help="Folder where all output files will be saved.")
    parser.add_argument("--enable-geocoding", action="store_true", help="Enable ArcGIS geocoding for missing/derived location lat-long. This may be slow.")
    return parser.parse_args()


ARGS = parse_args()
INPUT_FILE = Path(ARGS.input_file)
RERA_MASTER_FILE = Path(ARGS.rera_file)
OUTPUT_DIR = Path(ARGS.output_dir)
ENABLE_GEOCODING = ARGS.enable_geocoding

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

if not INPUT_FILE.exists():
    raise FileNotFoundError(f"Input file not found: {INPUT_FILE}")

if not RERA_MASTER_FILE.exists():
    raise FileNotFoundError(f"RERA master file not found: {RERA_MASTER_FILE}")

print("=" * 70)
print("TELANGANA DB1 SINGLE FILE PIPELINE")
print("=" * 70)
print(f"Input file      : {INPUT_FILE}")
print(f"RERA master file: {RERA_MASTER_FILE}")
print(f"Output folder   : {OUTPUT_DIR}")
print(f"Geocoding       : {'Enabled' if ENABLE_GEOCODING else 'Skipped'}")
print("=" * 70)

# ### Single File Processing

merged_df = pd.read_excel(INPUT_FILE)
print(f"Single input file loaded successfully. Shape: {merged_df.shape}")

# ## Drop Duplicates

# In[3]:


# === SIMPLE VERSION - DROP DUPLICATES ACROSS ALL COLUMNS ===

# Define columns
cols = [
    'S.No.', 'Description of property', 'Reg.Date Exe.Date Pres.Date',
    'Nature & Mkt.Value Con. Value', 'Name of Parties Executant(EX) & Claimants(CL)',
    'Vol/Pg No CD No Doct No/Year', 'Document No', 'District', 'Sub-Registrar Office'
]

# Original count
orig = len(merged_df)
print(f"Original: {orig:,} records")

# Identify duplicate records before dropping them
duplicate_mask = merged_df.duplicated(subset=cols, keep='first')
duplicate_records = merged_df[duplicate_mask].copy()

# Save duplicate records that are removed in this step
duplicate_records_file = OUTPUT_DIR / "duplicate_records_removed.xlsx"
duplicate_records.to_excel(duplicate_records_file, index=False)
print(f"Duplicate records removed file saved: {duplicate_records_file}")

# Drop duplicates
df_cleaned = merged_df.drop_duplicates(subset=cols, keep='first')

# Result
cleaned = len(df_cleaned)
print(f"Cleaned: {cleaned:,} records")
print(f"Removed: {orig - cleaned:,} duplicate records")
print(f"Unique records: {cleaned/orig*100:.2f}% of original")


# In[4]:


df_cleaned.info()


# ## Drop Rows With Rows with only '-' and Rows with 'W-B: 0-0'

# In[5]:


import pandas as pd

# Create a copy of the cleaned dataframe
df_clean = df_cleaned.copy()

# Convert the column to string type and strip whitespace
descriptions = df_clean['Description of property'].fillna('').astype(str).str.strip()

# Pattern to match:
# 1. Rows containing only "-"
# 2. Rows containing exactly "W-B: 0-0"
# 3. Blank values
# 4. Null values
pattern = r'^\s*$|^\s*-\s*$|^\s*W-B:\s*0-0\s*$|^nan$|^None$'

# Find rows matching the pattern
rows_to_delete = descriptions.str.contains(pattern, regex=True)

# Count rows
deleted_count = rows_to_delete.sum()
total_rows_before = len(df_clean)

# Display rows that will be deleted
print("Rows that will be deleted:")
print("=" * 60)

deleted_rows = df_clean[rows_to_delete].copy()

# Count each type separately
hyphen_mask = descriptions.str.contains(
    r'^\s*-\s*$',
    regex=True
)

wb_mask = descriptions.str.contains(
    r'^\s*W-B:\s*0-0\s*$',
    regex=True
)

blank_mask = descriptions.str.contains(
    r'^\s*$',
    regex=True
)

null_mask = df_clean['Description of property'].isnull()

# Add deletion reason for audit clarity
def get_deletion_reason(value):
    if pd.isna(value):
        return "Null Description of property"
    value = str(value).strip()
    if value == "":
        return "Blank Description of property"
    if re.fullmatch(r"-", value):
        return "Only '-' in Description of property"
    if re.fullmatch(r"W-B:\s*0-0", value, flags=re.IGNORECASE):
        return "Only 'W-B: 0-0' in Description of property"
    if value.lower() in ["nan", "none"]:
        return "Invalid text value in Description of property"
    return "Matched deletion pattern"

if len(deleted_rows) > 0:
    deleted_rows["deletion_reason"] = deleted_rows['Description of property'].apply(get_deletion_reason)

    # Print all deleted rows
    for idx, row in deleted_rows.iterrows():
        print(f"Index {idx}: '{row['Description of property']}' | Reason: {row['deletion_reason']}")
else:
    print("No rows found matching the deletion criteria")

print("\nBreakdown:")
print(f"Rows with only '-': {hyphen_mask.sum()}")
print(f"Rows with 'W-B: 0-0': {wb_mask.sum()}")
print(f"Blank rows: {blank_mask.sum()}")
print(f"Null rows in 'Description of property': {null_mask.sum()}")

# --------------------------------------------------
# SAVE ALL INVALID/DELETED RECORDS TO EXCEL
# This file is created even when there are zero deleted rows.
# --------------------------------------------------
deleted_file_name = OUTPUT_DIR / "deleted_records.xlsx"
deleted_rows.to_excel(deleted_file_name, index=False)

print("\nDeleted records Excel file created successfully.")
print(f"File Name: {deleted_file_name}")

print(f"\nTotal rows to delete: {deleted_count}")
print(f"Total rows before deletion: {total_rows_before}")

# --------------------------------------------------
# ASK FOR CONFIRMATION
# --------------------------------------------------

confirm = input("\nDo you want to proceed with deletion? (yes/no): ")

if confirm.lower() == 'yes':

    # Keep rows that DO NOT match the pattern
    df_clean = df_clean[~rows_to_delete].copy()

    # Reset index after deletion
    df_clean.reset_index(drop=True, inplace=True)

    print("\nDeletion completed successfully.")
    print(f"Rows remaining: {len(df_clean)}")

    # --------------------------------------------------
    # SAVE CLEANED DATAFRAME
    # --------------------------------------------------

    cleaned_file_name = OUTPUT_DIR / "cleaned_dataset.xlsx"

    df_clean.to_excel(cleaned_file_name, index=False)

    print("\nCleaned dataset saved successfully.")
    print(f"File Name: {cleaned_file_name}")

else:
    print("Deletion cancelled.")


# In[6]:


df_clean.shape


# In[7]:


# Party type mappings (Seller types, Buyer types)
SELLER_TYPES = {"EX", "MR", "DR", "RR", "PL", "LR","FP"}  # Added PL and LR as Sellers
BUYER_TYPES = {"CL", "ME", "DE", "RE", "AY", "LE","SP"}   # Added AY and LE as Buyers
ALL_PARTY_TYPES = SELLER_TYPES | BUYER_TYPES  # Union of all types

BOUND_START_RE = re.compile(r"(?i)\bbound\w*\s*:\s*")

def clean_spaces(s: str) -> str:
    """Clean extra spaces but preserve intentional newlines"""
    if not isinstance(s, str):
        return s
    # Split into lines, clean each line, then rejoin
    lines = s.split('\n')
    cleaned_lines = [re.sub(r"\s+", " ", line).strip() for line in lines]
    return "\n".join(cleaned_lines)

def normalize_basic(s: str) -> str:
    s = str(s)
    s = re.sub(r"(?i)VILL\s*/\s*COL", "VILL/COL", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def pick_unit_token(v: str, field_name: str = "") -> str:
    if not v:
        return ""

    # For BUILT field, return the full value as is (including unit)
    if field_name == "BUILT":
        return v.strip()

    # For EXTENT field, keep the original logic
    for p in v.split()[:12]:
        if "SQ" in p.upper():
            return p
    return v

def convert_extent_to_sq_ft(extent_value: str) -> str:
    """Convert extent value from SQ.Yd to SQ.Ft (direct conversion)"""
    if pd.isna(extent_value) or not isinstance(extent_value, str) or not extent_value:
        return ""

    extent_value = extent_value.strip()

    # Check if it's in SQ.Yd format (e.g., "190SQ.Yd" or "190 SQ.Yd" or "190.5SQ.Yd")
    sq_yd_match = re.search(r"([\d.]+)\s*SQ\.?Yd\.?", extent_value, re.IGNORECASE)

    if sq_yd_match:
        try:
            sq_yd = float(sq_yd_match.group(1))
            sq_ft = sq_yd * 9  # Direct conversion: 1 Square Yard = 9 Square Feet
            # Format to 2 decimal places - REMOVED UNIT
            return f"{sq_ft:.2f}"
        except ValueError:
            return ""

    return ""

def convert_built_to_sq_ft(built_value: str) -> str:
    """Convert built-up area value from SQ.Ft to SQ.Ft (no conversion needed, just formatting)"""
    if pd.isna(built_value) or not isinstance(built_value, str) or not built_value:
        return ""

    built_value = built_value.strip()

    # Check if it's in SQ.Ft format - handles "50SQ. FT", "50 SQ.FT", "50SQ.Ft", "50 SQ. FT", etc.
    sq_ft_match = re.search(r"([\d.]+)\s*SQ\.?\s*FT\.?", built_value, re.IGNORECASE)

    if sq_ft_match:
        try:
            sq_ft = float(sq_ft_match.group(1))
            # Format to 2 decimal places - REMOVED UNIT
            return f"{sq_ft:.2f}"
        except ValueError:
            return ""

    return ""

def extract_dates(date_text: str) -> dict:
    """Extract Registration, Execution, and Presentation dates"""
    out = {"Registration Date": "", "Execution Date": "", "Presentation Date": ""}
    if pd.isna(date_text) or not isinstance(date_text, str):
        return out

    # Pattern to match (R) date, (E) date, (P) date
    r_match = re.search(r"\(R\)\s*(\d{1,2}-\d{1,2}-\d{4})", date_text, re.IGNORECASE)
    e_match = re.search(r"\(E\)\s*(\d{1,2}-\d{1,2}-\d{4})", date_text, re.IGNORECASE)
    p_match = re.search(r"\(P\)\s*(\d{1,2}-\d{1,2}-\d{4})", date_text, re.IGNORECASE)

    if r_match:
        out["Registration Date"] = r_match.group(1)
    if e_match:
        out["Execution Date"] = e_match.group(1)
    if p_match:
        out["Presentation Date"] = p_match.group(1)

    return out

def extract_document_info(doc_text: str) -> dict:
    """Extract Document type code, Document Type, Market Value, Consideration Value"""
    out = {
        "Document type code": "", 
        "Document Type": "", 
        "Market Value": "", 
        "Consideration Value": ""
    }
    if pd.isna(doc_text) or not isinstance(doc_text, str):
        return out

    # Extract document type code (first 4 digits)
    code_match = re.search(r"^(\d{4})", doc_text.strip())
    if code_match:
        out["Document type code"] = code_match.group(1)

    # Extract document type (between code and Mkt.Value)
    doc_type_match = re.search(r"^\d{4}\s+(.+?)(?:\s+Mkt\.Value:|$)", doc_text, re.IGNORECASE)
    if doc_type_match:
        out["Document Type"] = clean_spaces(doc_type_match.group(1))

    # Extract Market Value
    mkt_match = re.search(r"Mkt\.Value:\s*(?:Rs\.?)?\s*([0-9,]+)", doc_text, re.IGNORECASE)
    if mkt_match:
        out["Market Value"] = mkt_match.group(1).replace(",", "")

    # Extract Consideration Value
    cons_match = re.search(r"Cons\.Value:\s*(?:Rs\.?)?\s*([0-9,]+)", doc_text, re.IGNORECASE)
    if cons_match:
        out["Consideration Value"] = cons_match.group(1).replace(",", "")

    return out

def extract_parties(parties_text: str) -> dict:
    """Extract Seller and Buyer from party information - supports multiple sellers and buyers"""
    out = {"Seller": "", "Buyer": ""}
    if pd.isna(parties_text) or not isinstance(parties_text, str):
        return out

    # Clean the text first
    parties_text = clean_spaces(parties_text)

    # Remove duplicate content (if the text is repeated)
    text_length = len(parties_text)
    half_length = text_length // 2

    if text_length > 20 and parties_text[:half_length] == parties_text[half_length:]:
        parties_text = parties_text[:half_length]

    sellers = []
    buyers = []
    seen_sellers = set()
    seen_buyers = set()

    # FIRST: Split by numbered entries (1., 2., 3., etc.) - THIS IS CRITICAL
    # This regex splits on spaces followed by a number and dot
    entries = re.split(r'\s+(?=\d+\.)', parties_text)

    # If splitting didn't work well, try alternative split
    if len(entries) <= 1:
        # Find all numbered entries
        entries = re.findall(r'\d+\.[^.]*(?:\([^)]+\)[^.]*)*', parties_text)

    # Process each numbered entry separately
    for entry in entries:
        entry = entry.strip()
        if not entry:
            continue

        # Extract the number and the rest
        number_match = re.match(r'(\d+)\.\s*(.*)', entry)
        if number_match:
            number, content = number_match.groups()
        else:
            content = entry

        # Check for party type in this entry
        found_type = None
        for party_type in ALL_PARTY_TYPES:
            type_pattern = rf'\(({party_type})\)'
            type_match = re.search(type_pattern, content, re.IGNORECASE)
            if type_match:
                found_type = party_type.upper()
                # Remove the party type tag from content
                content = re.sub(type_pattern, '', content, flags=re.IGNORECASE).strip()
                break

        if not found_type:
            continue

        # Clean up the name
        name = content.strip()

        # Remove any trailing number patterns
        name = re.sub(r'\s+\d+\.\s*$', '', name)
        name = re.sub(r'^\s+|\s+$', '', name)

        # Skip if name is empty
        if not name:
            continue

        # Create a normalized version for duplicate checking
        normalized_name = re.sub(r'\s+', '', name.upper())
        normalized_name = re.sub(r'[^\w\s]', '', normalized_name)

        # Add to appropriate list based on party type
        if found_type in SELLER_TYPES:
            if normalized_name and normalized_name not in seen_sellers:
                sellers.append(name)
                seen_sellers.add(normalized_name)
        elif found_type in BUYER_TYPES:
            if normalized_name and normalized_name not in seen_buyers:
                buyers.append(name)
                seen_buyers.add(normalized_name)

    # Join with newlines
    out["Seller"] = "\n".join(sellers) if sellers else ""
    out["Buyer"] = "\n".join(buyers) if buyers else ""

    return out

def segment_fields(text: str) -> dict:
    """Extract property description fields from text"""
    out = {k: "" for k in ["VILL/COL", "W-B", "SURVEY", "PLOT", "HOUSE", "APARTMENT", "BLOCK", "FLAT", "EXTENT", "BUILT", "Boundires"]}
    if pd.isna(text):
        return out

    t = normalize_basic(text)

    # ---- Fix 1: Ensure proper spacing before "Boundires:" ----
    # Add space before "Boundires:" if it's attached to other text
    t = re.sub(r"([^ ])Boundires:", r"\1 Boundires:", t, flags=re.IGNORECASE)
    t = re.sub(r"([^ ])bound\w*:", r"\1 bound:", t, flags=re.IGNORECASE)

    # ---- 1) Boundaries: regex slice (most reliable) ----
    bb = BOUND_START_RE.search(t)
    if bb:
        # Get everything after the boundary marker
        remaining_text = t[bb.end():]
        # Find where the next label might start (to capture complete boundaries)
        next_label_pos = len(remaining_text)

        # Look for any of the other labels that might come after boundaries
        for label in ["VILL/COL:", "W-B:", "SURVEY:", "PLOT:", "HOUSE:", "APARTMENT:", "BLOCK:", "FLAT:", "EXTENT:", "BUILT:"]:
            pos = remaining_text.upper().find(label.upper())
            if 0 < pos < next_label_pos:
                next_label_pos = pos

        out["Boundires"] = clean_spaces(remaining_text[:next_label_pos])
        # Remove boundaries part from text so it doesn't interfere with other parsing
        t_main = t[:bb.start()].strip()
    else:
        # ---- Alternative: Try to find boundaries at the end of string ----
        # Look for boundary pattern at the end (common pattern with [N], [S], etc.)
        bound_pattern_at_end = re.search(r"(?i)(bound\w*\s*:.*?)(?:\[[NSEW]\].*?)+$", t)
        if bound_pattern_at_end:
            out["Boundires"] = clean_spaces(bound_pattern_at_end.group(1).split(":", 1)[1])
            t_main = t[:bound_pattern_at_end.start()].strip()
        else:
            t_main = t

    # ---- NEW: Handle VILL/COL when it's at the beginning without label ----
    # Check if text starts with something that's not a known label (like "VENKATGIRI-1")
    if t_main and not t_main.upper().startswith(("VILL/COL:", "W-B:", "SURVEY:", "PLOT:", "HOUSE:", "APARTMENT:", "BLOCK:", "FLAT:", "EXTENT:", "BUILT:")):
        # Extract the first part until we hit a known label
        first_part_match = re.match(r"^([^:]+?)\s+(?=W-B:|SURVEY:|PLOT:|HOUSE:|APARTMENT:|BLOCK:|FLAT:|EXTENT:|BUILT:|Boundires:)", t_main, re.IGNORECASE)
        if first_part_match:
            out["VILL/COL"] = clean_spaces(first_part_match.group(1))
            # Remove the extracted part from t_main for further processing
            t_main = t_main[len(first_part_match.group(1)):].strip()

    # ---- 2) Parse other fields by simple label splits ----
    # Ensure labels have ":" (only for known labels)
    # Be careful with HOUSE to not match "/HOUSE SITE" in VILL/COL
    for k in ["VILL/COL", "W-B", "SURVEY", "PLOT", "APARTMENT", "BLOCK", "FLAT", "EXTENT", "BUILT"]:
        t_main = re.sub(rf"(?i)\b{k}\b\s*(?!:)", f"{k}:", t_main)

    # Handle HOUSE separately with a more precise pattern
    # Only add colon if HOUSE is at word boundary and not part of VILL/COL
    t_main = re.sub(rf"(?i)(?<!/)\bHOUSE\b\s*(?!:)", "HOUSE:", t_main)

    # Clean up any double colons
    t_main = re.sub(r"::+", ":", t_main)

    def grab(label, s):
        # Improved to better detect field boundaries
        # Special handling for HOUSE to avoid matching in VILL/COL
        if label == "HOUSE":
            # More precise pattern for HOUSE
            m = re.search(rf"(?i)(?<!/)\b{re.escape(label)}\s*:\s*(.*?)(?=\s+(?:VILL/COL|W-B|SURVEY|PLOT|APARTMENT|BLOCK|FLAT|EXTENT|BUILT|Boundires)\s*:|$)", s)
        else:
            m = re.search(rf"(?i)\b{re.escape(label)}\s*:\s*(.*?)(?=\s+(?:VILL/COL|W-B|SURVEY|PLOT|HOUSE|APARTMENT|BLOCK|FLAT|EXTENT|BUILT|Boundires)\s*:|$)", s)

        if m:
            value = m.group(1)
            # Additional cleanup: remove any trailing text that might contain next field's label
            value = re.sub(r'\s+(?:VILL/COL|W-B|SURVEY|PLOT|HOUSE|APARTMENT|BLOCK|FLAT|EXTENT|BUILT|Boundires)\s*:.*$', '', value, flags=re.IGNORECASE)
            return clean_spaces(value)
        return ""

    # Only grab VILL/COL if we didn't already extract it from the beginning
    if not out["VILL/COL"]:
        out["VILL/COL"] = grab("VILL/COL", t_main)

    out["W-B"]      = grab("W-B", t_main)
    out["SURVEY"]   = grab("SURVEY", t_main)
    out["PLOT"]     = grab("PLOT", t_main)
    out["HOUSE"]    = grab("HOUSE", t_main)
    out["APARTMENT"] = grab("APARTMENT", t_main)
    out["BLOCK"]    = grab("BLOCK", t_main)
    out["FLAT"]     = grab("FLAT", t_main)
    out["EXTENT"]   = grab("EXTENT", t_main)
    out["BUILT"]    = grab("BUILT", t_main)

    # ---- 3) Clean EXTENT/BUILT ----
    # Handle the case where boundaries text might have leaked into EXTENT
    if out["EXTENT"] and any(bound_word in out["EXTENT"].upper() for bound_word in ["BOUND", "[N]", "[S]", "[E]", "[W]"]):
        # Split on boundary markers
        for bound_marker in [" bound", " Bound", " BOUND", "[N]", "[S]", "[E]", "[W]"]:
            if bound_marker in out["EXTENT"]:
                out["EXTENT"] = out["EXTENT"].split(bound_marker)[0].strip()
                break

    out["EXTENT"] = pick_unit_token(out["EXTENT"], "EXTENT")
    out["BUILT"]  = pick_unit_token(out["BUILT"], "BUILT")

    # ---- Final check: If boundaries still empty but we see boundary patterns in main text ----
    if not out["Boundires"]:
        # Look for boundary pattern anywhere in original text
        bound_match = re.search(r"(?i)(?:bound\w*\s*:|\b(?:north|south|east|west|n|s|e|w)[\s:]*).*?(?:\[[NSEW]\].*?)+", t)
        if bound_match:
            # Extract just the boundary description part
            bound_text = bound_match.group(0)
            if ":" in bound_text:
                out["Boundires"] = clean_spaces(bound_text.split(":", 1)[1])
            else:
                out["Boundires"] = clean_spaces(bound_text)

    return out

# ===== CLASSIFICATION FUNCTION FOR TRANSACTION TYPE =====
def classify_transaction(doc_type):
    sales_types = [
        "Sale Deed",
        "AGREEMENT OF SALE CUM GPA",
        "Sale Agreement Without Possess",
        "Sale Agreement With Possession",
        "CONVEYANCE FOR CONSIDERATION",
        "Sale deed executed by A.P.Hous",
        "Sale Deeds executed by Courts",
        "Sale Certificate",
        "Assignment deed",
        "RECONVEYANCE DEED EXECUTED BY",
        "Sale of life interest",
        "Sale deed executed by or infav",
        "Sale deed executed by Society",
        "Sale deed in favour of State o"
    ]

    lease_types = [
        "Lease Deed",
        "Lease in favour of State/Centr",
        "Lease(others)",
        "Surrender of Lease",
        "Transfer of Lease"
    ]

    if pd.isna(doc_type):
        return "Others"

    doc_type = str(doc_type).strip()

    if doc_type in sales_types:
        return "Sales"
    elif doc_type in lease_types:
        return "Lease"
    else:
        return "Others"

# ===== ENHANCED FUNCTION FOR PROPERTY TYPE CLASSIFICATION =====
# ===== ENHANCED FUNCTION FOR PROPERTY TYPE CLASSIFICATION =====
# ===== ENHANCED FUNCTION FOR PROPERTY TYPE CLASSIFICATION =====
import re
import pandas as pd


# ===== PROCESS THE DATAFRAME =====
print("Starting processing of df_clean...")

# Make a copy to avoid modifying the original
df_processed = df_clean.copy()

src_col = "Description of property"

# Extract property description fields
print("Extracting property description fields...")
parsed = df_processed[src_col].apply(segment_fields).apply(pd.Series)
df_processed = pd.concat([df_processed, parsed], axis=1)

# Add EXTENT in SqFt column (direct conversion from SQ.Yd to SQ.Ft)
print("Converting EXTENT to SqFt...")
df_processed["EXTENT in SqFt"] = df_processed["EXTENT"].apply(convert_extent_to_sq_ft)

# Add BUILT in SqFt column (no conversion, just formatting)
print("Converting BUILT to SqFt...")
df_processed["BUILT in SqFt"] = df_processed["BUILT"].apply(convert_built_to_sq_ft)

# Extract dates from Reg.Date Exe.Date Pres.Date column
if "Reg.Date Exe.Date Pres.Date" in df_processed.columns:
    print("Extracting dates...")
    dates_parsed = df_processed["Reg.Date Exe.Date Pres.Date"].apply(extract_dates).apply(pd.Series)
    # Insert date columns after the original date column
    date_col_idx = df_processed.columns.get_loc("Reg.Date Exe.Date Pres.Date") + 1
    for i, col in enumerate(["Registration Date", "Execution Date", "Presentation Date"]):
        df_processed.insert(date_col_idx + i, col, dates_parsed[col])

# Extract document info from Nature & Mkt.Value Con. Value column
if "Nature & Mkt.Value Con. Value" in df_processed.columns:
    print("Extracting document information...")
    doc_parsed = df_processed["Nature & Mkt.Value Con. Value"].apply(extract_document_info).apply(pd.Series)
    # Insert document columns after the original document column
    doc_col_idx = df_processed.columns.get_loc("Nature & Mkt.Value Con. Value") + 1
    for i, col in enumerate(["Document type code", "Document Type", "Market Value", "Consideration Value"]):
        df_processed.insert(doc_col_idx + i, col, doc_parsed[col])

# Extract seller and buyer from Name of Parties Executant(EX) & Claimants(CL) column
if "Name of Parties Executant(EX) & Claimants(CL)" in df_processed.columns:
    print("Extracting seller and buyer information...")
    parties_parsed = df_processed["Name of Parties Executant(EX) & Claimants(CL)"].apply(extract_parties).apply(pd.Series)
    # Insert party columns after the original party column
    party_col_idx = df_processed.columns.get_loc("Name of Parties Executant(EX) & Claimants(CL)") + 1
    for i, col in enumerate(["Seller", "Buyer"]):
        df_processed.insert(party_col_idx + i, col, parties_parsed[col])

# ===== ADD TRANSACTION TYPE COLUMN AFTER DOCUMENT TYPE =====
# Check if Document Type column exists (it should from the extraction above)
if "Document Type" in df_processed.columns:
    print("Adding Transaction Type column...")
    # Create transaction type values
    transaction_values = df_processed["Document Type"].apply(classify_transaction)

    # Insert Transaction Type column after Document Type
    doc_type_idx = df_processed.columns.get_loc("Document Type")
    df_processed.insert(doc_type_idx + 1, "Transaction Type", transaction_values)


    # ===== PRINT TRANSACTION TYPE SUMMARY =====
    print("\n" + "="*60)
    print("TRANSACTION TYPE SUMMARY")
    print("="*60)
    transaction_summary = df_processed["Transaction Type"].value_counts()
    sales_count = transaction_summary.get("Sales", 0)
    lease_count = transaction_summary.get("Lease", 0)
    others_count = transaction_summary.get("Others", 0)
    total_count = len(df_processed)

    # Print in the requested format
    print(f"Sales\tLease\tOthers\tTotal")
    print(f"{sales_count}\t{lease_count}\t{others_count}\t{total_count}")
    print("="*60)

# ---- Debug: Show rows where Boundaries is still empty ----
empty_boundaries = df_processed[df_processed["Boundires"] == ""]
if not empty_boundaries.empty:
    print(f"Found {len(empty_boundaries)} rows with empty boundaries:")
    for idx, row in empty_boundaries.head(10).iterrows():
        print(f"\nRow {idx}:")
        print(f"Original text: {row[src_col]}")
        print("-" * 50)

# Debug: Show sample of BUILT and BUILT in SqFt to verify conversion
print("\n--- BUILT Conversion Sample (first 10 rows) ---")
sample_rows = df_processed[["BUILT", "BUILT in SqFt"]].head(10)
print(sample_rows.to_string())

# Debug: Show sample of EXTENT and EXTENT in SqFt to verify conversion
print("\n--- EXTENT Conversion Sample (first 10 rows) ---")
sample_rows = df_processed[["EXTENT", "EXTENT in SqFt"]].head(10)
print(sample_rows.to_string())

# Debug: Show sample of parties extraction with new mappings
print("\n--- Parties Extraction Sample with MR/ME, DR/DE, RR/RE Mappings (first 20 rows) ---")
if "Name of Parties Executant(EX) & Claimants(CL)" in df_processed.columns:
    # Get rows that might contain the new party types
    sample_df = df_processed[["Name of Parties Executant(EX) & Claimants(CL)", "Seller", "Buyer"]].head(20)

    # Also check specifically for rows with MR, ME, DR, DE, RR, RE
    print("\nRows with new party types (MR/ME, DR/DE, RR/RE, PL/AY, LR/LE):")
    for idx, row in df_processed.head(50).iterrows():
        text = str(row["Name of Parties Executant(EX) & Claimants(CL)"])
        if any(x in text for x in ['(MR)', '(ME)', '(DR)', '(DE)', '(RR)', '(RE)', '(PL)', '(AY)', '(LR)', '(LE)', '(FP)', '(SP)']):
            print(f"\nRow {idx}:")
            print(f"Original: {text[:100]}..." if len(text) > 100 else f"Original: {text}")
            print(f"Seller:\n{row['Seller']}")
            print(f"Buyer:\n{row['Buyer']}")
            print("-" * 40)

    print("\nFirst 20 rows sample:")
    print(sample_df.to_string())

print(f"\nProcessing complete! df_processed now has {len(df_processed.columns)} columns and {len(df_processed)} rows.")
print("You can continue working with df_processed for further analysis.")


# ### Proprty type filled with flat Where FLAT Column Contain Digit

# In[8]:


# Check each value in the actual column and put "FLAT" if it's a whole number
df_processed['Property_type'] = df_processed['FLAT'].apply(
    lambda x: 'FLAT' if str(x).strip().isdigit() else ''
)

# View the result
print(df_processed[['FLAT', 'Property_type']])


# In[9]:


flats_only = df_processed[df_processed['Property_type'] == 'FLAT']
print(flats_only)


# ### Property Type Using Regex

# In[11]:


import re
import pandas as pd
from tqdm import tqdm

def extract_core_property_text(description_text):
    """
    Extract only the core property information, removing ALL boundaries data
    """
    if not isinstance(description_text, str):
        return ""

    # Work with original case for now, will convert to uppercase later if needed
    text = description_text

    # =========================
    # STRATEGY 1: Split at boundaries marker and take first part
    # =========================
    boundaries_markers = [
        r'\bBoundires\s*:',
        r'\bBoundaries\s*:',
        r'\bBoundry\s*:',
        r'\bBounds\s*:',
        r'\bBoundaries\s*\[',
        r'\bBoundires\s*\[',
    ]

    core_text = text
    for marker in boundaries_markers:
        parts = re.split(marker, text, maxsplit=1, flags=re.IGNORECASE)
        if len(parts) > 1:
            core_text = parts[0]
            break

    # =========================
    # STRATEGY 2: Remove everything after bracket patterns like [N]:, [S]:, etc.
    # =========================
    # Remove any [N]:, [S]:, [E]:, [W]: patterns and everything after them
    core_text = re.sub(r'\[[NSWE]\]\s*:.*$', '', core_text, flags=re.IGNORECASE)

    # Remove standalone brackets with content
    core_text = re.sub(r'\[[^\]]+\]', '', core_text)

    # =========================
    # STRATEGY 3: Remove lines containing boundary keywords
    # =========================
    boundary_keywords = [
        r'Boundires', r'Boundaries', r'Boundry', r'Bounds',
        r'\[N\]', r'\[S\]', r'\[E\]', r'\[W\]',
        r'Gramapanchayath', r'Gram Panchayat', r'Municipal Office',
        r'Road\s*$', r'Wide Road', r'House Of', r'Land Of',
        r'OFFICE OF THE GRAMPANCHAYATH', r'MUNICIPAL OFFICE'
    ]

    for keyword in boundary_keywords:
        core_text = re.sub(rf'.*{keyword}.*(\n|$)', '', core_text, flags=re.IGNORECASE)

    # =========================
    # STRATEGY 4: Keep only the first part before any boundary indicators
    # =========================
    match = re.search(r'(\[N\]\s*:|\[S\]\s*:|\[E\]\s*:|\[W\]\s*:)', core_text, re.IGNORECASE)
    if match:
        core_text = core_text[:match.start()]

    # Clean up extra spaces and normalize
    core_text = re.sub(r'\s+', ' ', core_text).strip()

    return core_text


def classify_property_type_regex(description_text):
    """
    Classify property type using regex patterns
    STRICTLY avoids ALL boundaries data
    Returns one of: Flat, Shop, Office, Parking, Plot, House, Others
    """
    if not isinstance(description_text, str) or not description_text.strip():
        return "Others"

    # =========================
    # FIRST: Extract only core property text (remove ALL boundaries)
    # =========================
    core_text = extract_core_property_text(description_text)

    # If core text is empty, use original but with boundaries removed
    if not core_text:
        # Fallback: remove everything after "Boundires:" or similar
        boundaries_pattern = r'(?:Boundires|Boundaries|Boundry|Bounds)\s*:.*$'
        core_text = re.sub(boundaries_pattern, '', description_text, flags=re.IGNORECASE | re.DOTALL)
        core_text = re.sub(r'\[[^\]]*\]', '', core_text)
        core_text = re.sub(r'\s+', ' ', core_text).strip()

    # If still empty, return Others
    if not core_text:
        return "Others"

    # Convert to uppercase for case-insensitive matching
    text_upper = core_text.upper()

    # =========================
    # CLASSIFICATION ON CORE TEXT ONLY
    # =========================

    # =========================
    # PATTERN 1: SHOP (MUST COME FIRST - Check for SHOP in flat/apartment names)
    # =========================
    shop_patterns = [
        r'\bFLAT\s*:\s*SHOP\d+\b',                           # FLAT: SHOP408, FLAT: SHOP14
        r'\bFLAT\s*:\s*SHOP[A-Z0-9]+\b',                     # FLAT: SHOPG3A, FLAT: SHOPG-5
        r'\bFLAT\s*:\s*\d+SHOP\b',                           # FLAT: 408SHOP
        r'\bFLAT\s*:\s*SHOP\s*[A-Z0-9-]+\b',                 # FLAT: SHOP G-5
        r'\bAPARTMENT\s*:.*\bFLAT\s*:\s*SHOP\d+\b',          # APARTMENT: X FLAT: SHOP408
        r'\bAPARTMENT\s*:.*\bFLAT\s*:\s*SHOP[A-Z0-9]+\b',    # APARTMENT: X FLAT: SHOPG3A
        r'\bSHOP\s+(?:NO|NUMBER)\s*:?\s*\d+',                # SHOP NO 123
        r'\bSHOP\s*:\s*\d+',                                 # SHOP: 123
        r'\bSHOP\s*:\s*[A-Z0-9-]+',                          # SHOP: identifier
        r'\bCOMMERCIAL\s+SHOP\b',                            # Commercial shop
        r'\bSHOP\s+NO\s*\d+\b',                              # SHOP NO 8
        r'\bSHOP\s*:\s*[A-Z]+\b',                            # SHOP: X
        r'\bRETAIL\s+SHOP\b',                                # Retail shop
        r'\bSHOP\b(?!.*\b(?:HOUSE|FLAT)\b)',                 # Shop without house/flat context
    ]

    for pattern in shop_patterns:
        if re.search(pattern, text_upper):
            return "Shop"

    # =========================
    # PATTERN 2: OFFICE (Enhanced for all office variations)
    # =========================
    office_patterns = [
        r'\bFLAT\s*:\s*OFFICE\b',                            # FLAT: OFFICE
        r'\bFLAT\s*:\s*OFC\s+[IVX]+\b',                     # FLAT: OFC I, OFC II, OFC III
        r'\bFLAT\s*:\s*OFF\s+[0-9A-Z-]+\b',                 # FLAT: OFF 4-C
        r'\bFLAT\s*:\s*\d+/OFF\b',                           # FLAT: 403/OFF
        r'\bFLAT\s*:\s*OFFICE\s+SPACE\b',                    # FLAT: OFFICE SPACE
        r'\bFLAT\s*:\s*OFC\b',                               # FLAT: OFC
        r'\bAPARTMENT\s*:.*\bFLAT\s*:\s*OFFICE\b',           # APARTMENT: X FLAT: OFFICE
        r'\bAPARTMENT\s*:.*\bFLAT\s*:\s*OFC\b',              # APARTMENT: X FLAT: OFC
        r'\bAPARTMENT\s*:.*\bFLAT\s*:\s*OFF\b',              # APARTMENT: X FLAT: OFF
        r'\bOFFICE\s+(?:NO|NUMBER)\s*:?\s*\d+',              # OFFICE NO 123
        r'\bOFFICE\s*:\s*\d+',                               # OFFICE: 123
        r'\bOFFICE\s*:\s*[A-Z0-9-]+',                       # OFFICE: identifier
        r'\bOFFICE\s+SPACE\b',                               # Office space
        r'\bOFFICE\s+SUITE\b',                               # Office suite
        r'\bCOMMERCIAL\s+OFFICE\b',                          # Commercial office
        r'\bCORPORATE\s+OFFICE\b',                           # Corporate office
        r'\bOFFICE\b(?!.*\b(?:HOUSE)\b)',                    # Office not near house
    ]

    for pattern in office_patterns:
        if re.search(pattern, text_upper):
            # Double-check it's not from boundaries
            if not re.search(r'(?:MUNICIPAL|GRAMAPANCHAYATH|GRAM\s+PANCHAYAT|GOVERNMENT)', text_upper):
                return "Office"

    # =========================
    # PATTERN 3: FLAT/APARTMENT (Now after SHOP and OFFICE)
    # =========================
    flat_patterns = [
        r'\bAPARTMENT\s*:.*\bFLAT\s+(?:NO|NUMBER)?\s*:?\s*\d+',  # APARTMENT: X FLAT: 123
        r'\bAPARTMENT\s*:.*\bFLAT\s+\d+',                         # APARTMENT: X FLAT 123
        r'\bAPARTMENT\s*:.*\bFLAT\s*:\s*[A-Z0-9-]+',              # APARTMENT: X FLAT: identifier
        r'\bAPARTMENT\s*:\s*[A-Z0-9\s]+\s+FLAT\s*:\s*\d+',        # APARTMENT: NAME FLAT: number
        r'\bAPARTMENT\s*:\s*[A-Z0-9\s]+\s+FLAT\s+\d+',            # APARTMENT: NAME FLAT number
        r'\bFLAT\s+(?:NO|NUMBER)\s*:?\s*\d+',                      # FLAT NO 123
        r'\bFLAT\s*:\s*\d+',                                      # FLAT: 123
        r'\bFLAT\s*:\s*[A-Z0-9-]+',                               # FLAT: identifier
        r'\bFLAT\s+[A-Z]\d+\b',                                   # FLAT A11
        r'\bFLAT\s+NO\s*\d+\b',                                   # FLAT NO 403
        r'\bAPARTMENT\s*:\s*[A-Z0-9\s]+',                         # APARTMENT: NAME
        r'\bMIG-\d+\b',                                          # MIG housing scheme
        r'\bLIG-\d+\b',                                          # LIG housing scheme
        r'\bSCHEDULE\s+[A-Z]\d+\b',                              # Schedule A11
        r'\bFLAT\b.*\bFLOOR\b',                                  # Flat with floor
        r'\bFLOOR\b.*\bFLAT\b',                                  # Floor with flat
    ]

    for pattern in flat_patterns:
        if re.search(pattern, text_upper):
            return "Flat"

    # =========================
    # PATTERN 4: HOUSE
    # =========================
    house_patterns = [
        r'\bHOUSE\s+(?:NO|NUMBER)\s*:?\s*\d+(?:[-\/\d]*)?',  # HOUSE: 1/88, HOUSE NO 123
        r'\bHOUSE\s*:\s*\d+(?:[-\/\d]*)?',                    # HOUSE: 123
        r'\bHOUSE\s*:\s*[A-Z0-9-]+',                         # HOUSE: identifier
        r'\bHOUSE\s+[A-Z]\d+\b',                              # HOUSE A11
        r'\bRESIDENTIAL\s+HOUSE\b',
        r'\bDETACHED\s+HOUSE\b',
        r'\bHOUSE\b',                                         # Any house mention
    ]

    for pattern in house_patterns:
        if re.search(pattern, text_upper):
            return "House"

    # =========================
    # PATTERN 5: PARKING
    # =========================
    parking_patterns = [
        r'\bPARKING\s+(?:SLOT|SPACE|AREA)\b',
        r'\bPARKING\s+NO\b',
        r'\bDRIVE\s+WAY\b',
        r'\bCAR\s+PARK\b',
        r'\bPARKING\s*:\s*X\b',
        r'\bPARKING\b',
    ]

    for pattern in parking_patterns:
        if re.search(pattern, text_upper):
            return "Parking"

    # =========================
    # PATTERN 6: PLOT/LAND
    # =========================
    plot_patterns = [
        r'\bPLOT\s+(?:NO|NUMBER)\s*:?\s*\d+',
        r'\bPLOT\s*:\s*\d+',
        r'\bPLOT\s*:\s*[A-Z0-9-]+',
        r'\bPLOT\s+[A-Z]\d+\b',
        r'\bSURVEY\s*(?:NO|NUMBER)\s*:?\s*\d+',
        r'\bSY\s+NO\s*\d+\b',
        r'\bGRAMAKANTAM\b',                                  # Village land
        r'\bLAND\b(?!.*\b(?:HOUSE|FLAT)\b)',
        r'\bOPEN\s+PLOT\b',
        r'\bRESIDENTIAL\s+PLOT\b',
        r'\bSITE\s+NO\s*\d+\b',
        r'\bVACANT\s+LAND\b',
        r'\bAGRICULTURAL\s+LAND\b',
    ]

    for pattern in plot_patterns:
        if re.search(pattern, text_upper):
            return "Plot"

    # =========================
    # FALLBACK: Check for area-based classification
    # =========================
    extent_match = re.search(r'\bEXTENT\s*:\s*(\d+(?:\.\d+)?)\s*(SQ\.YDS|SQ\.FT|SQ\.M|SQ\.YARD|SQUARE\s+(?:YARDS|FEET|METERS))', text_upper)
    built_match = re.search(r'\bBUILT\s*:\s*(\d+(?:\.\d+)?)\s*(SQ\.FT|SQ\.M|SQUARE\s+(?:FEET|METERS))', text_upper)

    if extent_match and built_match:
        extent_value = float(extent_match.group(1))
        built_value = float(built_match.group(1))
        extent_unit = extent_match.group(2)

        # Convert to consistent unit
        if 'YDS' in extent_unit or 'YARD' in extent_unit:
            extent_sqft = extent_value * 9
        else:
            extent_sqft = extent_value

        # Check if built area is substantial (likely a building)
        if built_value > 500:
            # Look for shop keywords first
            if re.search(r'\bSHOP\b', text_upper):
                return "Shop"
            # Look for office keywords
            elif re.search(r'\bOFFICE\b', text_upper) or re.search(r'\bOFC\b', text_upper):
                return "Office"
            # Look for residential keywords
            elif re.search(r'\bHOUSE\b', text_upper) or re.search(r'\bRESIDENTIAL\b', text_upper):
                return "House"
            elif re.search(r'\bFLAT\b', text_upper) or re.search(r'\bAPARTMENT\b', text_upper):
                return "Flat"
            else:
                return "House"  # Default to house for residential properties

    # =========================
    # FINAL FALLBACK: Check for property type indicators
    # =========================
    if re.search(r'\bSHOP\b', text_upper):
        return "Shop"
    elif re.search(r'\bOFFICE\b', text_upper) or re.search(r'\bOFC\b', text_upper):
        return "Office"
    elif re.search(r'\bHOUSE\b', text_upper):
        return "House"
    elif re.search(r'\bFLAT\b', text_upper) or re.search(r'\bAPARTMENT\b', text_upper):
        return "Flat"
    elif re.search(r'\bPLOT\b', text_upper) or re.search(r'\bLAND\b', text_upper) or re.search(r'\bSURVEY\b', text_upper):
        return "Plot"
    elif re.search(r'\bPARKING\b', text_upper):
        return "Parking"

    return "Others"


def classify_batch_regex(descriptions, show_progress=True):
    """
    Process a batch of descriptions using regex with strict boundary avoidance
    """
    results = []
    iterator = tqdm(descriptions, desc="Classifying property types") if show_progress else descriptions

    for desc in iterator:
        result = classify_property_type_regex(desc)
        results.append(result)

    return results


# =========================
# MAIN PROCESSING FUNCTION
# =========================
def process_property_types_with_regex(df_processed, property_type_col='Property Type', 
                                      description_col=None, overwrite_existing=False):
    """
    Process DataFrame to add/update property types using regex with strict boundary avoidance

    Parameters:
    - df_processed: DataFrame with property descriptions
    - property_type_col: Name of property type column (default: 'Property Type')
    - description_col: Name of description column (auto-detects if None)
    - overwrite_existing: If True, overwrite all; if False, only fill blank rows

    Returns:
    - DataFrame with updated property types
    """
    # Auto-detect description column
    if description_col is None:
        possible_desc_cols = ['Description of property', 'Description', 'Bhumapan', 
                              'property_description', 'description']
        for col in possible_desc_cols:
            if col in df_processed.columns:
                description_col = col
                break

        if description_col is None:
            for col in df_processed.columns:
                if 'description' in col.lower() or 'property' in col.lower():
                    description_col = col
                    break

    if description_col is None:
        raise ValueError(f"No description column found. Available columns: {list(df_processed.columns)}")

    print(f"Using description column: '{description_col}'")

    # Check if property type column exists
    if property_type_col not in df_processed.columns:
        df_processed[property_type_col] = None
        print(f"Created new column: '{property_type_col}'")
        overwrite_existing = True

    # Determine which rows to process
    if overwrite_existing:
        rows_to_process = len(df_processed)
        print(f"Overwriting all {rows_to_process} rows")
        mask_to_process = [True] * len(df_processed)
    else:
        # Only process blank rows
        blank_mask = df_processed[property_type_col].isna() | (df_processed[property_type_col].astype(str).str.strip() == '')
        rows_to_process = blank_mask.sum()
        print(f"Processing {rows_to_process} blank rows out of {len(df_processed)} total")
        mask_to_process = blank_mask

    if rows_to_process == 0:
        print("No rows to process")
        return df_processed

    # Extract descriptions for rows to process
    descriptions = df_processed.loc[mask_to_process, description_col].fillna('').tolist()

    # Classify using regex with strict boundary avoidance
    print("Classifying property types with strict boundary avoidance...")
    property_types = classify_batch_regex(descriptions)

    # Update DataFrame
    df_processed.loc[mask_to_process, property_type_col] = property_types

    # Show summary
    print("\n" + "="*60)
    print("PROPERTY TYPE DISTRIBUTION:")
    print("="*60)
    value_counts = df_processed[property_type_col].value_counts()
    for prop_type, count in value_counts.items():
        percentage = (count / len(df_processed)) * 100
        print(f"  {prop_type}: {count} ({percentage:.1f}%)")
    print("="*60)

    # Show sample of classifications
    print("\n" + "="*60)
    print("SAMPLE CLASSIFICATIONS (First 10):")
    print("="*60)
    sample_df = df_processed[df_processed[property_type_col].notna()].head(10)
    for idx, row in sample_df.iterrows():
        desc_preview = str(row[description_col])[:100] if pd.notna(row[description_col]) else "N/A"
        print(f"\n{row[property_type_col]}: {desc_preview}...")
    print("="*60)

    return df_processed


# =========================
# USAGE EXAMPLE - READY TO PROCESS YOUR DATAFRAME
# =========================
if __name__ == "__main__":
    # Test cases including the new office examples
    test_cases = [
        {
            "text": "VILL/COL: PUPPALGUDA/PUPPAL GUDA W-B: 0-0 SURVEY: 282/P APARTMENT: EON-HYDERABAD FLAT: OFC I EXTENT: .5SQ.Yds BUILT: 51SQ. FT Boundires: [N]: WASHROOMS [S] STAIRCASE [E]: COMMON PASSAGE [W]: COMMON PASSAGE",
            "expected": "Office"
        },
        {
            "text": "VILL/COL: NARSINGI/COMMERCIAL-3 W-B: 0-0 SURVEY: 158/P APARTMENT: JYOTHI OPTIMA FLAT: OFF 4-C EXTENT: 35.4SQ.Yds BUILT: 3000SQ. FT Boundires: [N]: OPEN TO SKY [S] CORRIDOR [E]: OPEN TO SKY [W]: OFFICE SPACE NO.4-B",
            "expected": "Office"
        },
        {
            "text": "VILL/COL: KOKAPET/COMMERCIAL-3 W-B: 0-0 SURVEY: 107/P 108/P HOUSE: . APARTMENT: LAXMI INFOBAHN, TOWER-5 FLAT: OFFICE EXTENT: 1SQ.Yds BUILT: 2500SQ. FT Boundires: [N]: OPEN TO SKY & PART OF THE TOWER 6 OFFICE SPACE [S] OPEN TO SKY [E]: OPEN TO SKY [W]: OPEN TO SKY & PART OF THE TOWER-6 OFFICE SPACE",
            "expected": "Office"
        },
        {
            "text": "VILL/COL: PUPPALGUDA/RESIDENTIAL-2 W-B: 0-0 SURVEY: 285/P APARTMENT: TOWER-2 FLAT: OFFICE EXTENT: 77.5SQ.Yds BUILT: 3876SQ. FT Boundires: [N]: NEIGHBOURS PROPERTY [S] OPEN TO SKY [E]: NEIGHBOURS PROPERTY [W]: NEIGHBOURS PROPERTY",
            "expected": "Office"
        },
        {
            "text": "VILL/COL: NARSINGI/COMMERCIAL-3 W-B: 0-0 SURVEY: 160 161 PLOT: 10 APARTMENT: RAICHANDANI BUSINESS BAY FLAT: SHOP408 EXTENT: 30SQ.Yds BUILT: 1386SQ. FT Boundires: [N]: SHOP NO.407 [S] CORRIDOR [E]: CORRIDOR [W]: OPEN TO SKY / SETBACK",
            "expected": "Shop"
        },
        {
            "text": "VILL/COL: PUPPALGUDA/RESIDENTIAL-2 W-B: 0-0 SURVEY: 285/P APARTMENT: TOWER-6 FLAT: OFFICE EXTENT: 54SQ.Yds BUILT: 2720SQ. FT Boundires: [N]: NEIGHBOURS PROPERTY [S] OPEN TO SKY [E]: NEIGHBOURS PROPERTY [W]: NEIGHBOURS PROPERTY",
            "expected": "Office"
        },
        {
            "text": "VILL/COL: SOLIPUR/RESIDENTIAL W-B: 0-0 HOUSE: 1/88 EXTENT: 136SQ.Yds BUILT: 1120SQ. FT Boundires: [N]: 12'-0\" WIDE ROAD [S] SOLIPUR MUNICIPAL OFFICE [E]: HOUSE OF ANIMONI NARSIMULU [W]: HOUSE OF S VENKATAIAH",
            "expected": "House"
        },
        {
            "text": "VILL/COL: ANANTAWARAM/ANANTAWARAM W-B: 0-0 SURVEY: 130 131 132 133 134 144 145 PLOT: 6-RAAVIPALLE APARTMENT: RAAVI PALLE 06 EXTENT: 1500SQ.Yds BUILT: 4712SQ. FT BLOCK 1Boundires: [N]: Raavi Palle, Flat No 05 [S] Raavi Palle, Flat No 07 [E]: Drive way [W]: community Farm Area",
            "expected": "Flat"
        },
        {
            "text": "SHOP NO: 123 MAIN ROAD EXTENT: 500SQ.FT BUILT: 500SQ.FT",
            "expected": "Shop"
        },
        {
            "text": "FLAT NO: 403 TOWER B EXTENT: 1200SQ.FT BUILT: 1200SQ.FT",
            "expected": "Flat"
        }
    ]

    print("="*60)
    print("TESTING WITH OFFICE, SHOP, AND FLAT CLASSIFICATIONS:")
    print("="*60)

    for i, test in enumerate(test_cases, 1):
        result = classify_property_type_regex(test["text"])
        status = "✓" if result == test["expected"] else "✗"
        print(f"\nTest {i} {status}:")
        print(f"  Expected: {test['expected']}")
        print(f"  Got: {result}")
        print(f"  Preview: {test['text'][:100]}...")

        # Show what the core text looks like after boundary removal
        core_text = extract_core_property_text(test["text"])
        print(f"  Core text: {core_text[:100]}...")

    print("\n" + "="*60)
    print("PROCESSING YOUR DATAFRAME")
    print("="*60)

    # Process your actual DataFrame
    # Make sure 'df_processed' is already loaded with your data
    df_processed = process_property_types_with_regex(
        df_processed,
        property_type_col='Property Type',
        overwrite_existing=False  # Set to True to overwrite all existing property types
    )

    # Save the updated DataFrame
    output_path = OUTPUT_DIR / "extraction_output.xlsx"
    df_processed.to_excel(output_path, index=False)
    print(f"\n✅ Saved to: {output_path}")

    # Verify results
    blank_after = df_processed['Property Type'].isna().sum()
    print(f"\nRemaining blank property types: {blank_after}")
    if blank_after == 0:
        print("✓ All property types filled!")
    else:
        print(f"⚠️  {blank_after} rows still have no property type")


# In[12]:


import pandas as pd

# Clean column names (important to avoid KeyError)
df_processed.columns = df_processed.columns.str.strip()

# Exact mapping
mapping = {
    "flat": "Flat",
    "office": "Office",
    "shop": "Shop"
}

# Direct transformation (no extra column)
df_processed["Property_Class"] = (
    df_processed["Property Type"]
    .astype(str)
    .str.strip()
    .str.lower()
    .map(mapping)
    .fillna("Others")
)


# In[14]:


df_processed.columns


# ## Rename Column Names 

# In[15]:


# Define mapping from current columns to desired column names
column_mapping = {
    'S.No.': 's.no',
    'Description of property': 'property_description',
    'Registration Date': 'transaction_date',
    'Execution Date': 'date_of_agreement_execution',
    'Presentation Date': 'presentation_date',
    'Nature & Mkt.Value Con. Value': 'nature & mkt.value con.value',
    'Document type code': 'document_type_code',
    'Document Type': 'transaction_type',
    'Transaction Type': 'transaction_category',
    'Market Value': 'guideline_value',
    'Consideration Value': 'agreement_price',
    'Name of Parties Executant(EX) & Claimants(CL)': 'party_info',
    'Seller': 'seller_name',
    'Buyer': 'buyer_name',
    'Vol/Pg No CD No Doct No/Year': 'internal_document_number',
    'Document No': 'document_number',
    'District': 'district',
    'Sub-Registrar Office': 'sub_registrar_office_name',
    'VILL/COL': 'village_name',
    'W-B': 'w-b',
    'SURVEY': 'survey_number',
    'PLOT': 'plot_number',
    'HOUSE': 'house_number',
    'APARTMENT': 'project_name',
    'BLOCK': 'block',
    'FLAT': 'unit_number',
    'EXTENT': 'extent_sq_ft',
    'BUILT': 'built_area_sq_ft',
    'Boundires': 'boundaries',
    'EXTENT in SqFt': 'gross_carpet_area_sq_ft',
    'BUILT in SqFt': 'built_area_sq_ft_alternate',
    'Property Type': 'property_type_raw',
    'Property_Class': 'property_type',
    'Final_Area':'final_area',
    'Rate':'rate'
}

# Create mapping for only the columns that exist in your DataFrame
existing_mapping = {col: column_mapping[col] for col in df_processed.columns if col in column_mapping}

# Rename the columns
df_processed = df_processed.rename(columns=existing_mapping)

# Display the renamed columns
print("Renamed columns:")
print(df_processed.columns.tolist())
print(f"\nTotal columns after rename: {len(df_processed.columns)}")


# In[16]:


def clean_village_name(value):
    if pd.isna(value):
        return value

    value = str(value).strip()

    # Keep only data before "/"
    value = value.split("/")[0].strip()

    return value

# Create cleaned village column
df_processed['village_name_cleaned'] = df_processed['village_name'].apply(clean_village_name)

col_to_move = 'village_name_cleaned'
after_col = 'village_name'

# Remove column temporarily
temp = df_processed.pop(col_to_move)

# Get position of village_name column
position = df_processed.columns.get_loc(after_col)

# Insert cleaned column after village_name
df_processed.insert(position + 1, col_to_move, temp)


# In[18]:


# ====== COLUMN NAME ======
col = 'project_name'   # <-- change if your column name is different
new_col = f'{col}_cleaned'  # Name of the new column

# Abbreviations to keep capital
abbreviations = ['SLG', 'SV', 'SR', 'SS', 'GK']

def clean_name(text):
    if pd.isna(text):
        return text

    text = str(text)

    # Remove quotes
    text = re.sub(r"[\"']", "", text)

    # Remove extra spaces
    text = text.strip()

    # Convert to proper case
    text = text.lower().title()

    # Fix 's
    text = re.sub(r"'S\b", "'s", text)

    # Fix abbreviations
    words = text.split()
    words = [w.upper() if w.upper() in abbreviations else w for w in words]

    return " ".join(words)

# Create new column with cleaned names
df_processed[new_col] = df_processed[col].apply(clean_name)


# In[19]:


df_processed.to_excel(OUTPUT_DIR / "test.xlsx", index=False)


# ## Rate

# In[21]:


# Clean column names
df_processed.columns = df_processed.columns.str.strip()

# Convert required columns to numeric
numeric_cols = [
    "agreement_price",
    "built_area_sq_ft_alternate",
    "gross_carpet_area_sq_ft"
]

for col in numeric_cols:
    df_processed[col] = (
        df_processed[col]
        .astype(str)
        .str.replace(",", "", regex=False)
        .str.strip()
    )
    df_processed[col] = pd.to_numeric(df_processed[col], errors="coerce")


# Clean property type
df_processed["property_type_clean"] = (
    df_processed["property_type"]
    .astype(str)
    .str.strip()
    .str.lower()
)

# Select final area
df_processed["Final_Area"] = np.where(
    df_processed["property_type_clean"].isin(["flat", "office", "shop"]),
    df_processed["built_area_sq_ft_alternate"],
    df_processed["gross_carpet_area_sq_ft"]
)

# Avoid division by zero
df_processed.loc[df_processed["Final_Area"] <= 0, "Final_Area"] = np.nan

# Calculate rate
df_processed["Rate"] = df_processed["agreement_price"] / df_processed["Final_Area"]


# In[22]:


df_processed.to_excel(OUTPUT_DIR / "Telangana_Final_With_Rate.xlsx", index=False)


# In[23]:


df_processed.columns


# ### Project Matching with Rera on the basis of poject name with location and without location for index Assign

# In[27]:


import pandas as pd
import re
from rapidfuzz import process, fuzz

# =========================================================
# File paths and column names
# =========================================================
standard_file_path = RERA_MASTER_FILE
output_path = OUTPUT_DIR / "matched_output.xlsx"

# Raw file columns
raw_proj_col = "project_name_cleaned"
raw_village_col = "village_name_cleaned"

# RERA file columns
std_proj_col = "project_name"
std_location_col = "location_name"

# =========================================================
# Matching thresholds
# =========================================================
THRESHOLD_STRONG = 90
THRESHOLD_GOOD = 85
THRESHOLD_REVIEW = 80

# =========================================================
# Light normalization
# =========================================================
def normalize_light(name):
    if pd.isna(name):
        return ""

    name = str(name).lower().strip()
    name = name.replace("&", " and ")
    name = name.replace("/", " ")
    name = name.replace("-", " ")
    name = name.replace(",", " ")
    name = name.replace(".", " ")
    name = re.sub(r"[^a-z0-9\s]", "", name)

    replacements = {
        "apt": "apartment",
        "apts": "apartments",
        "ph": "phase",
        "phs": "phase",
        "bldg": "building",
        "blk": "block",
        "twr": "tower",
    }

    words = name.split()
    words = [replacements.get(w, w) for w in words]

    roman_map = {
        "i": "1", "ii": "2", "iii": "3", "iv": "4",
        "vi": "6", "vii": "7", "viii": "8", "ix": "9", "x": "10"
    }
    words = [roman_map.get(w, w) for w in words]

    name = " ".join(words)
    name = re.sub(r"\bphase\s*(\d+)\b", r"phase \1", name)
    name = re.sub(r"\s+", " ", name).strip()

    return name

# =========================================================
# Combined normalized string helper
# =========================================================
def get_combined_normalized(row, col1, col2, sep=" | "):
    val1 = normalize_light(row.get(col1, ""))
    val2 = normalize_light(row.get(col2, ""))

    if val1 and val2:
        return val1 + sep + val2
    elif val1:
        return val1
    elif val2:
        return val2
    else:
        return ""

# =========================================================
# Match status from score
# =========================================================
def get_status_from_score(score):
    if score >= THRESHOLD_STRONG:
        return "Strong Match"
    elif score >= THRESHOLD_GOOD:
        return "Good Match"
    elif score >= THRESHOLD_REVIEW:
        return "Review"
    else:
        return "No Reliable Match"

# =========================================================
# Load files
# =========================================================
df_standard = pd.read_excel(standard_file_path)
df_raw = df_processed.copy()

# =========================================================
# Check required columns
# =========================================================
if std_proj_col not in df_standard.columns:
    raise ValueError(f"RERA file missing '{std_proj_col}'")
if std_location_col not in df_standard.columns:
    raise ValueError(f"RERA file missing '{std_location_col}'")
if raw_proj_col not in df_raw.columns:
    raise ValueError(f"Raw file missing '{raw_proj_col}'")
if raw_village_col not in df_raw.columns:
    raise ValueError(f"Raw file missing '{raw_village_col}'")

# =========================================================
# Create normalized columns
# =========================================================
df_standard["norm_combined"] = df_standard.apply(
    lambda row: get_combined_normalized(row, std_proj_col, std_location_col), axis=1
)
df_raw["norm_combined"] = df_raw.apply(
    lambda row: get_combined_normalized(row, raw_proj_col, raw_village_col), axis=1
)

df_standard["norm_project_only"] = df_standard[std_proj_col].apply(normalize_light)
df_raw["norm_project_only"] = df_raw[raw_proj_col].apply(normalize_light)

# Remove blanks
df_standard = df_standard.copy()
df_raw = df_raw.copy()

# =========================================================
# Prepare unique choices for combined match
# =========================================================
df_standard_combined = df_standard[df_standard["norm_combined"].str.strip() != ""].copy()
df_standard_combined["proj_len"] = df_standard_combined[std_proj_col].astype(str).str.len()

df_standard_combined_unique = df_standard_combined.loc[
    df_standard_combined.groupby("norm_combined")["proj_len"].idxmax()
].copy()

df_standard_combined_unique.drop(columns=["proj_len"], inplace=True)

combined_choices = df_standard_combined_unique["norm_combined"].tolist()

# Map normalized combined -> full row index
combined_index_map = dict(
    zip(df_standard_combined_unique["norm_combined"], df_standard_combined_unique.index)
)

# =========================================================
# Prepare unique choices for project-only match
# =========================================================
df_standard_proj = df_standard[df_standard["norm_project_only"].str.strip() != ""].copy()
df_standard_proj["proj_len"] = df_standard_proj[std_proj_col].astype(str).str.len()

df_standard_proj_unique = df_standard_proj.loc[
    df_standard_proj.groupby("norm_project_only")["proj_len"].idxmax()
].copy()

df_standard_proj_unique.drop(columns=["proj_len"], inplace=True)

project_only_choices = df_standard_proj_unique["norm_project_only"].tolist()

# Map normalized project -> full row index
project_only_index_map = dict(
    zip(df_standard_proj_unique["norm_project_only"], df_standard_proj_unique.index)
)

# =========================================================
# Step 1: Combined match function
# project + location  <->  project + location
# =========================================================
def get_best_combined_match(raw_norm):
    if not raw_norm or str(raw_norm).strip() == "":
        return pd.Series([None, 0, "No Match", None])

    match = process.extractOne(
        query=raw_norm,
        choices=combined_choices,
        scorer=fuzz.token_sort_ratio,
        score_cutoff=THRESHOLD_REVIEW - 1
    )

    if match is None:
        return pd.Series([None, 0, "No Match", None])

    best_norm, score, _ = match
    best_idx = combined_index_map[best_norm]
    best_project = df_standard.loc[best_idx, std_proj_col]
    status = get_status_from_score(score)

    return pd.Series([best_project, score, status, best_idx])

# =========================================================
# Step 2: Project-only match function
# project only  <->  project only
# =========================================================
def get_best_project_only_match(raw_proj_norm):
    if not raw_proj_norm or str(raw_proj_norm).strip() == "":
        return pd.Series([None, 0, "No Match", None])

    match = process.extractOne(
        query=raw_proj_norm,
        choices=project_only_choices,
        scorer=fuzz.token_sort_ratio,
        score_cutoff=THRESHOLD_REVIEW - 1
    )

    if match is None:
        return pd.Series([None, 0, "No Match", None])

    best_norm, score, _ = match
    best_idx = project_only_index_map[best_norm]
    best_project = df_standard.loc[best_idx, std_proj_col]
    status = get_status_from_score(score)

    return pd.Series([best_project, score, status, best_idx])

# =========================================================
# Apply Step 1
# =========================================================
df_raw[
    [
        "matched_project_name_combined",
        "match_score_combined",
        "match_status_combined",
        "matched_rera_index_combined"
    ]
] = df_raw["norm_combined"].apply(get_best_combined_match)

# =========================================================
# Apply Step 2
# =========================================================
df_raw[
    [
        "matched_project_name_only",
        "match_score_project_only",
        "match_status_project_only",
        "matched_rera_index_project_only"
    ]
] = df_raw["norm_project_only"].apply(get_best_project_only_match)

# =========================================================
# Convert matched RERA index columns to numeric before merge
# =========================================================
df_raw["matched_rera_index_combined"] = pd.to_numeric(
    df_raw["matched_rera_index_combined"],
    errors="coerce"
)

df_raw["matched_rera_index_project_only"] = pd.to_numeric(
    df_raw["matched_rera_index_project_only"],
    errors="coerce"
)

# =========================================================
# Bring RERA columns for combined match
# =========================================================
combined_rera_df = df_standard.copy()
combined_rera_df.columns = [
    col if col == std_proj_col else f"{col}_combined_rera"
    for col in combined_rera_df.columns
]

df_raw = df_raw.merge(
    combined_rera_df,
    how="left",
    left_on="matched_rera_index_combined",
    right_index=True
)

# =========================================================
# Bring RERA columns for project-only match
# =========================================================
project_only_rera_df = df_standard.copy()
project_only_rera_df.columns = [
    col if col == std_proj_col else f"{col}_project_only_rera"
    for col in project_only_rera_df.columns
]

df_raw = df_raw.merge(
    project_only_rera_df,
    how="left",
    left_on="matched_rera_index_project_only",
    right_index=True,
    suffixes=("", "_dup")
)

# =========================================================
# Optional location match check for combined result
# =========================================================
def compute_location_similarity(raw_village, rera_location):
    if pd.isna(raw_village) or pd.isna(rera_location):
        return pd.Series([0, "No Location/Village Data"])

    raw_village = str(raw_village).strip()
    rera_location = str(rera_location).strip()

    if raw_village == "" or rera_location == "":
        return pd.Series([0, "No Location/Village Data"])

    norm_village = normalize_light(raw_village)
    norm_location = normalize_light(rera_location)

    if norm_village == "" or norm_location == "":
        return pd.Series([0, "No Location/Village Data"])

    score = fuzz.token_sort_ratio(norm_village, norm_location)
    status = get_status_from_score(score)
    return pd.Series([score, status])

combined_location_col = f"{std_location_col}_combined_rera"
project_only_location_col = f"{std_location_col}_project_only_rera"

if combined_location_col in df_raw.columns:
    df_raw[
        ["combined_location_match_score", "combined_location_match_status"]
    ] = df_raw.apply(
        lambda row: compute_location_similarity(
            row.get(raw_village_col, ""),
            row.get(combined_location_col, "")
        ),
        axis=1
    )

if project_only_location_col in df_raw.columns:
    df_raw[
        ["project_only_location_match_score", "project_only_location_match_status"]
    ] = df_raw.apply(
        lambda row: compute_location_similarity(
            row.get(raw_village_col, ""),
            row.get(project_only_location_col, "")
        ),
        axis=1
    )

# =========================================================
# Rearrange columns for easy testing
# =========================================================
preferred_order = [
    raw_proj_col,
    raw_village_col,

    "norm_project_only",
    "norm_combined",

    "matched_project_name_combined",
    "match_score_combined",
    "match_status_combined",
    "matched_rera_index_combined",

    combined_location_col,
    "combined_location_match_score",
    "combined_location_match_status",

    "matched_project_name_only",
    "match_score_project_only",
    "match_status_project_only",
    "matched_rera_index_project_only",

    project_only_location_col,
    "project_only_location_match_score",
    "project_only_location_match_status",
]

# Keep only existing preferred columns
preferred_order = [col for col in preferred_order if col in df_raw.columns]

# Add remaining columns after preferred columns
remaining_cols = [col for col in df_raw.columns if col not in preferred_order]

df_raw = df_raw[preferred_order + remaining_cols]

# =========================================================
# Save output
# =========================================================
df_raw.to_excel(output_path, index=False)

# =========================================================
# Summary
# =========================================================
print("\n=== Matching completed successfully ===")
print("Step 1: Project + Location match completed")
print("Step 2: Project-only match completed")
print(f"\nOutput file: {output_path}")

print("\nCombined match status breakdown:")
print(df_raw["match_status_combined"].value_counts(dropna=False))

print("\nProject-only match status breakdown:")
print(df_raw["match_status_project_only"].value_counts(dropna=False))

combined_avg = df_raw.loc[df_raw["match_score_combined"] > 0, "match_score_combined"].mean()
project_only_avg = df_raw.loc[df_raw["match_score_project_only"] > 0, "match_score_project_only"].mean()

print(
    f"\nAverage combined match score: {combined_avg:.2f}"
    if pd.notna(combined_avg) else
    "\nAverage combined match score: 0"
)
print(
    f"Average project-only match score: {project_only_avg:.2f}"
    if pd.notna(project_only_avg) else
    "Average project-only match score: 0"
)

# =========================================================
# Sample preview
# =========================================================
sample_cols = [
    raw_proj_col,
    raw_village_col,
    "norm_combined",
    "norm_project_only",

    "matched_project_name_combined",
    "match_score_combined",
    "match_status_combined",

    "matched_project_name_only",
    "match_score_project_only",
    "match_status_project_only",
]

if combined_location_col in df_raw.columns:
    sample_cols.extend([
        combined_location_col,
        "combined_location_match_score",
        "combined_location_match_status"
    ])

if project_only_location_col in df_raw.columns:
    sample_cols.extend([
        project_only_location_col,
        "project_only_location_match_score",
        "project_only_location_match_status"
    ])

sample_cols = [col for col in sample_cols if col in df_raw.columns]

print("\nSample output rows:")
print(df_raw[sample_cols].head(20).to_string(index=False))
print(df_raw["matched_rera_index_combined"].dtype)
print(df_standard.index.dtype)

print(df_raw["matched_rera_index_project_only"].dtype)
print(df_standard.index.dtype)


# ### Columns rearrangement and kept only those needed

# In[32]:


import pandas as pd

# Load your file
df = df_raw.copy()
# -------------------------------
# 1. INDEX (new)
# -------------------------------
df["index"] = (
    df["index_combined_rera"]
    .combine_first(df["index_project_only_rera"])
)

# -------------------------------
# 2. PROJECT NAME (new)
# -------------------------------
df["project_name"] = (
    df["registered_project_name_combined_rera"]
    .combine_first(df["registered_project_name_project_only_rera"])
    .combine_first(df["project_name_cleaned"])
)

# -------------------------------
# 3. LOCATION (new)
# -------------------------------
# -------------------------------
# 3. LOCATION (new)
# Rule:
# index present  → RERA location
# index null     → village_name_cleaned
# -------------------------------

# Step 1: prepare RERA location
rera_location = (
    df["location_name_combined_rera"]
    .combine_first(df["location_name_project_only_rera"])
)

# Step 2: default assign RERA location
df["location_name"] = rera_location

# Step 3: override where index is null
df.loc[df["index"].isna(), "location_name"] = df.loc[
    df["index"].isna(),
    "village_name_cleaned"
]

# -------------------------------
# 4. BHK WISE CARPET AREA (new)
# -------------------------------
df["bhk_wise_carpet_area"] = (
    df["bhk_wise_carpet_area_combined_rera"]
    .combine_first(df["bhk_wise_carpet_area_project_only_rera"])
)

# -------------------------------
# 5. PROJECT TYPE (new column: project)
# project_only -> combined
# -------------------------------
df["project_type"] = (
    df["project_type_project_only_rera"]
    .combine_first(df["project_type_combined_rera"])
)

# -------------------------------
# 6. PINCODE (new)
# combined -> project_only
# -------------------------------
df["pincode"] = (
    df["pincode_combined_rera"]
    .combine_first(df["pincode_project_only_rera"])
)

# -------------------------------
# 7. PROJECT LATITUDE (new)
# combined -> project_only
# -------------------------------
df["project_latitude"] = (
    df["project_latitude_combined_rera"]
    .combine_first(df["project_latitude_project_only_rera"])
)

# -------------------------------
# 8. PROJECT LONGITUDE (new)
# combined -> project_only
# -------------------------------
df["project_longitude"] = (
    df["project_longitude_combined_rera"]
    .combine_first(df["project_longitude_project_only_rera"])
)

# -------------------------------
# FINAL COLUMN SELECTION
# -------------------------------
final_columns = [
    "index",
    "s.no",
    "property_description",
    "Reg.Date Exe.Date Pres.Date",
    "transaction_date",
    "date_of_agreement_execution",
    "presentation_date",
    "transaction_category",
    "document_type_code",
    "transaction_type",
    "guideline_value",
    "agreement_price",
    "party_info",
    "seller_name",
    "buyer_name",
    "internal_document_number",
    "document_number",
    "district",
    "sub_registrar_office_name",
    "village_name_cleaned",
    "w-b",
    "survey_number",
    "plot_number",
    "house_number",
    "project_name_cleaned",
    "project_name",
    "block",
    "unit_number",
    "extent_sq_ft",
    "built_area_sq_ft",
    "boundaries",
    "gross_carpet_area_sq_ft",
    "built_area_sq_ft_alternate",
    "property_type_raw",
    "property_type",
    "Final_Area",
    "Rate",
    "location_name",
    "bhk_wise_carpet_area",
    "project_type",
    "project_latitude",
    "project_longitude",
    "pincode"

]

# Keep only required columns
df_final = df[final_columns]

# Save output
df_final.to_excel(OUTPUT_DIR / "output_clean.xlsx", index=False)

print(f"Output file created: {OUTPUT_DIR / 'output_clean.xlsx'}")


# In[33]:


df_final.columns


# ### BHK Assign

# In[239]:


import ast
import re
from pathlib import Path

import pandas as pd
import numpy as np


# =========================
# INPUT / OUTPUT PATH
# =========================

OUTPUT_PATH = OUTPUT_DIR / "output_with_bhk_Assign.xlsx"

BHK_MAX_DIFF = 5


# =========================
# PARSE BHK AREA DATA
# =========================

def extract_bhk_from_key(key):
    """
    Converts dirty BHK keys:

    101.2BHK        -> 2BHK
    1003.3BHK       -> 3BHK
    1072.5BHK       -> 2.5BHK
    3432.5BHK       -> 2.5BHK
    519.4BHK        -> 4BHK
    10092BHK        -> 2BHK
    10143BHK        -> 3BHK
    2.5BHK          -> 2.5BHK

    Flat/unit style keys:
    flatno.1023bhk  -> 3BHK
    flatno.1013bhk  -> 3BHK
    unit4023bhk     -> 3BHK

    General logic:
    - Handles decimal patterns safely (only true numeric formats)
    - Handles fused keys by extracting last digit
    - Avoids misclassification like 1023 → 1BHK (fixed)
    """

    key_clean = str(key).lower().replace(" ", "")

    if not key_clean.endswith("bhk"):
        return None

    number_part = key_clean.replace("bhk", "")

    # Case 1: half BHK like 2.5BHK, 1072.5BHK
    if "." in number_part:
        before_dot, after_dot = number_part.split(".", 1)

        if after_dot == "5" and before_dot[-1] in ["1", "2", "3", "4", "5"]:
            return f"{before_dot[-1]}.5BHK"

        # Only apply decimal logic if it's clean numeric format like 101.2
        if re.match(r'^\d+\.\d+$', number_part):
            if after_dot and after_dot[0] in ["1", "2", "3", "4", "5"]:
                return f"{after_dot[0]}BHK"

    # Case 2: fused / flatno / unit keys → take last digit
    match = re.search(r'(\d+)$', number_part)
    if match:
        bhk_number = match.group(1)[-1]

        if bhk_number in ["1", "2", "3", "4", "5"]:
            return f"{bhk_number}BHK"

    return None


def parse_bhk_area_dict(x):
    if pd.isna(x) or str(x).strip() in ["", "[]", "nan", "None"]:
        return {}

    try:
        outer = ast.literal_eval(str(x))
    except Exception:
        return {}

    result = {}

    for project, value in outer.items():
        try:
            value_list = ast.literal_eval(value)
        except Exception:
            continue

        for item in value_list:
            if not isinstance(item, dict):
                continue

            for key, areas in item.items():
                bhk = extract_bhk_from_key(key)

                if not bhk:
                    continue

                for area in areas:
                    try:
                        result.setdefault(bhk, []).append(round(float(area), 2))
                    except Exception:
                        pass

    return result


# =========================
# ASSIGN BHK FROM FINAL AREA SQMT
# =========================

def assign_bhk_from_final_area(row):
    property_type = str(row.get("property_type", "")).strip().lower()

    if property_type != "flat":
        return row.get("property_type")

    try:
        carpet_sqmt = round(float(row["final_area_sqmt"]), 2)
    except:
        return None

    bhk_dict = parse_bhk_area_dict(row.get("bhk_wise_carpet_area"))

    if not bhk_dict:
        return None

    candidates = []

    for bhk, areas in bhk_dict.items():
        for area in areas:
            diff = abs(area - carpet_sqmt)

            if diff <= BHK_MAX_DIFF:
                candidates.append((bhk, area, diff))

    if not candidates:
        return None

    best_match = min(candidates, key=lambda x: x[2])

    return best_match[0]


def assign_matched_area(row):
    property_type = str(row.get("property_type", "")).strip().lower()

    if property_type != "flat":
        return None

    try:
        carpet_sqmt = round(float(row["final_area_sqmt"]), 2)
    except Exception:
        return None

    bhk_dict = parse_bhk_area_dict(row.get("bhk_wise_carpet_area"))

    best_diff = float("inf")
    matched_area = None

    for bhk, areas in bhk_dict.items():
        for area in areas:
            diff = abs(area - carpet_sqmt)

            if diff < best_diff and diff <= BHK_MAX_DIFF:
                best_diff = diff
                matched_area = area

    return matched_area


def assign_match_diff(row):
    property_type = str(row.get("property_type", "")).strip().lower()

    if property_type != "flat":
        return None

    try:
        carpet_sqmt = round(float(row["final_area_sqmt"]), 2)
    except Exception:
        return None

    bhk_dict = parse_bhk_area_dict(row.get("bhk_wise_carpet_area"))

    best_diff = float("inf")

    for bhk, areas in bhk_dict.items():
        for area in areas:
            diff = abs(area - carpet_sqmt)

            if diff < best_diff and diff <= BHK_MAX_DIFF:
                best_diff = diff

    return round(best_diff, 2) if best_diff != float("inf") else None


# =========================
# FALLBACK LOGIC
# =========================

def build_percentile_ranges(df):
    if "bhk_wise_carpet_area" not in df.columns:
        return {}

    expanded = []

    for _, row in df.iterrows():
        bhk_dict = parse_bhk_area_dict(row.get("bhk_wise_carpet_area"))

        for bhk, areas in bhk_dict.items():
            bhk_upper = bhk.upper()

            if bhk_upper in ["1BHK", "2BHK", "3BHK"]:
                for carpet in areas:
                    try:
                        carpet = float(carpet)

                        if 20 < carpet < 300:
                            expanded.append({
                                "BHK": bhk_upper,
                                "carpet": carpet
                            })

                    except:
                        pass

    df_exp = pd.DataFrame(expanded)

    if df_exp.empty:
        return {}

    grouped = df_exp.groupby("BHK")["carpet"].apply(list)

    p10 = {k: np.percentile(v, 10) for k, v in grouped.items()}
    p90 = {k: np.percentile(v, 90) for k, v in grouped.items()}

    ranges = {}

    if all(k in p10 and k in p90 for k in ["1BHK", "2BHK", "3BHK"]):
        ranges["<1BHK"] = (0, p10["1BHK"])
        ranges["1BHK"] = (p10["1BHK"], (p90["1BHK"] + p10["2BHK"]) / 2)
        ranges["2BHK"] = ((p90["1BHK"] + p10["2BHK"]) / 2, (p90["2BHK"] + p10["3BHK"]) / 2)
        ranges["3BHK"] = ((p90["2BHK"] + p10["3BHK"]) / 2, p90["3BHK"])
        ranges[">3BHK"] = (p90["3BHK"], float("inf"))

    return ranges


def assign_bhk_fallback(row, ranges):
    try:
        carpet = float(row["final_area_sqmt"])
    except:
        return None

    if not ranges:
        return None

    # FIRST check >3BHK only if available
    if ">3BHK" in ranges:
        low, high = ranges[">3BHK"]
        if low <= carpet <= high:
            return ">3BHK"

    # THEN check normal BHK ranges only if available
    for bhk in [">3BHK", "3BHK", "2BHK", "1BHK", "<1BHK"]:
        if bhk in ranges:
            low, high = ranges[bhk]
            if low <= carpet <= high:
                return bhk

    return None


# =========================
# MAIN
# =========================

def main():
    df_final = df[final_columns].copy()

    df_final["final_area_sqmt"] = pd.to_numeric(
        df_final["Final_Area"], 
        errors="coerce"
    ) / 10.764

    df_final["final_area_sqmt"] = df_final["final_area_sqmt"].round(2)

    df_final["assigned_bhk"] = df_final.apply(assign_bhk_from_final_area, axis=1)

    df_final["matched_rera_area_sqmt"] = df_final.apply(assign_matched_area, axis=1)

    df_final["bhk_match_diff_sqmt"] = df_final.apply(assign_match_diff, axis=1)

    ranges = build_percentile_ranges(df_final)
    print("RANGES:", ranges)

    mask = df_final["assigned_bhk"].isna()

    df_final.loc[mask, "assigned_bhk"] = df_final.loc[mask].apply(
        lambda row: assign_bhk_fallback(row, ranges),
        axis=1
    )

    df_final.to_excel(OUTPUT_PATH, index=False)

    print("BHK assignment completed.")
    print(f"Output saved at: {OUTPUT_PATH}")

    return df_final


if __name__ == "__main__":
    df_final = main()


# #### Location Lat/long fill by geocoder

# In[226]:

if ENABLE_GEOCODING:
    from geopy.geocoders import ArcGIS
    import time

    geolocator = ArcGIS(timeout=10)
    output_file = OUTPUT_DIR / "geocoded_location_output.csv"

    # make df_final independent
    df_final = df_final.copy()

    # remove duplicate column labels if any
    df_final = df_final.loc[:, ~df_final.columns.duplicated()].copy()

    # clean location
    if "location" not in df_final.columns and "location_name" in df_final.columns:
        df_final["location"] = df_final["location_name"]

    df_final["location"] = df_final["location"].astype(str).str.strip()

    # ONLY UNIQUE LOCATIONS
    unique_locations = df_final["location"].drop_duplicates().reset_index(drop=True)

    # Cache: location → (lat, lng)
    geo_cache = {}

    for i, location in unique_locations.items():

        address = f"{location}, Telangana, India"

        lat, lng = None, None

        try:
            result = geolocator.geocode(address)

            if result:
                lat = result.latitude
                lng = result.longitude

            geo_cache[location] = (lat, lng)

            # UPDATE ALL ROWS HAVING THIS LOCATION
            df_final.loc[
                df_final["location"] == location,
                ["location_latitude", "location_longitude"]
            ] = [lat, lng]

            # SAVE AFTER EACH RECORD
            df_final.to_csv(output_file, index=False)

            print(f"Done {i + 1}/{len(unique_locations)} | {location} -> {lat}, {lng}")

        except Exception as e:
            print(f"Error: {location} -> {e}")
            geo_cache[location] = (None, None)

        time.sleep(1)

    # Final save (safety)
    df_final.to_csv(output_file, index=False)

    print("Final geocoded file saved:", output_file)
else:
    print("Geocoding skipped. Use --enable-geocoding to run ArcGIS location lat-long fill.")

# ### All Data converted into Tiltle Case

# In[227]:


for col in df_final.select_dtypes(include=["object"]).columns:
    df_final[col] = df_final[col].apply(
        lambda x: x.title() if isinstance(x, str) else x
    )

df_final.to_excel(OUTPUT_DIR / "Updated_title_case_output.xlsx", index=False)


# ### Year and Quarter Extracted from transaction_date

# In[228]:


# Convert transaction_date to datetime
df_final["transaction_date"] = pd.to_datetime(
    df_final["transaction_date"],
    format="%d-%m-%Y",
    errors="coerce"
)

# Create year column
df_final["year"] = df_final["transaction_date"].dt.year

# Create quarter column like Q1-2024, Q2-2024
df_final["quarter"] = (
    "Q" + df_final["transaction_date"].dt.quarter.astype(str)
    + "-" +
    df_final["transaction_date"].dt.year.astype(str)
)

# Optional: convert transaction_date back to dd-mm-yyyy format
df_final["transaction_date"] = df_final["transaction_date"].dt.strftime("%d-%m-%Y")

# Save output
df_final.to_excel(OUTPUT_DIR / "Year_Quarter_output.xlsx", index=False)
print("Data Saved in given Location")

print(df_final[["transaction_date", "year", "quarter"]])


# ## Checklist

# In[231]:


import pandas as pd

# =========================
# INPUT / OUTPUT PATH
# =========================

# INPUT_FILE = r"D:\Nilesh\Checklist\TelanganaDB1\2024_Telangana_DB1_With_Year&Quarter_output_.xlsx"

OUTPUT_CHECKLIST = OUTPUT_DIR / "telangana_checklistDB1.xlsx"

OUTPUT_UPDATED_FILE = OUTPUT_DIR / "Telangana_DB1_With_Year_Quarter_output_Updated.xlsx"

# =========================
# STANDARD COLUMNS
# =========================

standard_columns = [
    "project_id",
    "internal_index_id",
    "project_name",
    "location_name",
    "village_name",
    "year",
    "quarter",
    "city_name",
    "transaction_category_id",
    "sub_registrar_office_code",
    "sub_registrar_office_name",
    "document_number",
    "transaction_type",
    "agreement_price",
    "guideline_value",
    "property_description",
    "transaction_date",
    "floor_number",
    "unit_number",
    "property_type_raw",
    "net_carpet_area_sq_m",
    "balcony_sq_m",
    "terrace_sq_m",
    "seller_name",
    "buyer_name",
    "transaction_category",
    "internal_document_number",
    "micr_number",
    "bank_type",
    "party_code",
    "date_of_agreement_execution",
    "stamp_duty_paid",
    "registration_fee",
    "project_latitude",
    "project_longitude",
    "location_latitude",
    "location_longitude",
    "property_type",
    "unit_configuration",
    "buyer_pincode",
    "buyer_locality",
    "buyer_district",
    "buyer_state",
    "is_llm_processed",
    "is_manual_processed",
    "tower_name",
    "gross_carpet_area_sq_ft",
    "price_per_sq_ft_gross_carpet",
    "is_duplicate",
    "sale_type",
    "project_type",
    "country_name",
    "state_name",
    "micro_market",
    "sub_locality",
    "pincode",
    "parking_count",
    "facing_direction",
    "view_type",
    "furnishing_status",
    "condition_status",
    "source_accessibility",
    "source_accessibility_way",
    "sourcing_cost",
    "sourcing_time",
    "data_type",
    "data_source"
]

# =========================
# LOAD EXCEL
# =========================

df = df_final.copy()

df.columns = (
    df.columns
    .astype(str)
    .str.strip()
    .str.replace("\ufeff", "", regex=False)
)

excel_columns = list(df.columns)
total_rows = len(df)

# =========================
# COLUMN MAPPING
# =========================

column_mapping = {
    "index": "internal_index_id"
}

mapped_excel_columns = [
    column_mapping.get(col, col) for col in excel_columns
]

# =========================
# CREATE STANDARD STATUS
# =========================

status_rows = []

for position, standard_col in enumerate(standard_columns, start=1):

    if standard_col in mapped_excel_columns:
        mapped_position = mapped_excel_columns.index(standard_col)
        original_excel_col = excel_columns[mapped_position]

        if original_excel_col == standard_col:
            status = "Available"
            action_required = "No action required"
        else:
            status = "Mapped"
            action_required = f"Excel column `{original_excel_col}` mapped to `{standard_col}`"

        dtype = str(df[original_excel_col].dtype)
        null_count = int(df[original_excel_col].isna().sum())
        non_null_count = int(df[original_excel_col].notna().sum())
        unique_count = int(df[original_excel_col].nunique(dropna=True))

        null_percentage = round((null_count / total_rows) * 100, 2) if total_rows > 0 else 0
        non_null_percentage = round((non_null_count / total_rows) * 100, 2) if total_rows > 0 else 0

        value_present = "Yes" if non_null_count > 0 else "No"

        sample_value = (
            df[original_excel_col].dropna().iloc[0]
            if not df[original_excel_col].dropna().empty
            else None
        )

    else:
        original_excel_col = None
        status = "Missing"
        action_required = "Missing column will be created in final updated file"

        dtype = None
        null_count = None
        non_null_count = 0
        unique_count = 0
        null_percentage = None
        non_null_percentage = 0
        value_present = "No"
        sample_value = None

    status_rows.append({
        "standard_position": position,
        "standard_column": standard_col,
        "excel_column": original_excel_col,
        "status": status,
        "excel_data_type": dtype,
        "total_rows": total_rows,
        "non_null_count": non_null_count,
        "non_null_percentage": non_null_percentage,
        "null_count": null_count,
        "null_percentage": null_percentage,
        "unique_values_count": unique_count,
        "value_present": value_present,
        "sample_value": sample_value,
        "action_required": action_required
    })

status_df = pd.DataFrame(status_rows)

# =========================
# EXTRA EXCEL COLUMNS
# =========================

extra_columns = [
    col for col in excel_columns
    if column_mapping.get(col, col) not in standard_columns
]

extra_rows = []

for col in extra_columns:
    null_count = int(df[col].isna().sum())
    non_null_count = int(df[col].notna().sum())
    unique_count = int(df[col].nunique(dropna=True))

    extra_rows.append({
        "excel_column": col,
        "mapped_name": column_mapping.get(col, col),
        "status": "Extra column in Excel",
        "excel_data_type": str(df[col].dtype),
        "total_rows": total_rows,
        "non_null_count": non_null_count,
        "null_count": null_count,
        "unique_values_count": unique_count,
        "value_present": "Yes" if non_null_count > 0 else "No",
        "action_required": "Extra column retained in final updated file"
    })

extra_df = pd.DataFrame(extra_rows)

# =========================
# CREATE FINAL UPDATED FILE
# =========================

df_updated = df.copy()

df_updated = df_updated.rename(columns=column_mapping)

for col in standard_columns:
    if col not in df_updated.columns:
        df_updated[col] = None

final_columns = standard_columns + [
    col for col in df_updated.columns if col not in standard_columns
]

df_updated = df_updated[final_columns]

df_updated.to_excel(OUTPUT_UPDATED_FILE, index=False)

# =========================
# SUMMARY
# =========================

summary_df = pd.DataFrame([
    {"metric": "Total Excel Columns", "value": len(excel_columns)},
    {"metric": "Total Standard Columns", "value": len(standard_columns)},
    {"metric": "Available Columns", "value": int((status_df["status"] == "Available").sum())},
    {"metric": "Mapped Columns", "value": int((status_df["status"] == "Mapped").sum())},
    {"metric": "Missing Columns", "value": int((status_df["status"] == "Missing").sum())},
    {"metric": "Columns With Values", "value": int((status_df["value_present"] == "Yes").sum())},
    {"metric": "Columns Without Values", "value": int((status_df["value_present"] == "No").sum())},
    {"metric": "Extra Excel Columns", "value": len(extra_columns)}
])

updated_file_status = pd.DataFrame([
    {"metric": "Input Excel Columns", "value": len(excel_columns)},
    {"metric": "Final Updated Columns", "value": len(df_updated.columns)},
    {"metric": "Final Updated Rows", "value": len(df_updated)},
    {"metric": "Missing Columns Created", "value": int((status_df["status"] == "Missing").sum())},
    {"metric": "Extra Columns Retained", "value": len(extra_columns)}
])

# =========================
# SAVE CHECKLIST FILE
# =========================

with pd.ExcelWriter(OUTPUT_CHECKLIST, engine="openpyxl") as writer:
    summary_df.to_excel(writer, sheet_name="Summary", index=False)
    status_df.to_excel(writer, sheet_name="Standard Structure Status", index=False)
    extra_df.to_excel(writer, sheet_name="Extra Excel Columns", index=False)
    updated_file_status.to_excel(writer, sheet_name="Updated File Status", index=False)

print("Checklist file created successfully:")
print(OUTPUT_CHECKLIST)

print("Final updated file created successfully:")
print(OUTPUT_UPDATED_FILE)


print("\nPipeline completed successfully.")
print(f"Final updated file: {OUTPUT_UPDATED_FILE}")
print(f"Checklist file: {OUTPUT_CHECKLIST}")
