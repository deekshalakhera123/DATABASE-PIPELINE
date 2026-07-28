# =============================================================================
# DB1 PIPELINE — STAGE 1: Load, Clean, Categorise, Village Map
# =============================================================================
# Input:  Raw IGR Excel (multi-village merged file)
# Output: Cleaned, categorised DataFrame → saved to Excel for LLM processing
# =============================================================================

import re
import pandas as pd
import numpy as np
from datetime import datetime
from pathlib import Path

from static import result_dict, word_number_dict


# =============================================================================
# CONFIG — edit this block per run
# =============================================================================

IGR_EXCEL   = Path(r".\sample_data.xlsx")
VILLAGE_DIR = Path(r".\Mumbai IGR Village Directory.xlsx")
OUTPUT_PATH = Path(r"Sample data for llm.xlsx")

# NOTE: registrationdate and dateofexecution have mixed formats per file.
#   True  → parse as MM/DD/YYYY  (American)
#   False → parse as DD/MM/YYYY  (Indian)
REGISTRATION_DATE_MDY = True
EXECUTION_DATE_MDY    = True


# =============================================================================
# CONSTANTS
# =============================================================================

KEEP_COLUMNS = [
    'srocode', 'internaldocumentnumber', 'docno', 'docname',
    'registrationdate', 'sroname', 'micrno', 'bank_type', 'party_code',
    'sellerparty', 'purchaserparty', 'propertydescription', 'areaname',
    'consideration_amt', 'marketvalue', 'dateofexecution',
    'stampdutypaid', 'registrationfees',
]

SALE_DOCTYPE = {
    'करारनामा', 'सेल डीड', 'अभिहस्तांतरणपत्र', 'खरेदीखत', 'विक्री करारनामा',
    'अँग्रीमेंट टू सेल', 'डीड ऑफ अपार्टमेंट', 'ट्रान्सफर डीड', 'साठेखत',
    'सेल सर्टिफिकेट', 'सर्टिफिकेट ऑफ सेल', 'कन्व्हेन्स डीड',
    'अविभाज्य हिश्याची पूर्ण विक्री', 'अँग्रीमेंट टू सेल ऑफ फ्लॅट', 'असाईनमेंट डीड',
    'विक्री प्रमाणपत्र', 'अँग्रीमेंट टू सेल ऑफ शॉप', 'इकरारनामा', 'खुषखरेदीखत',
    'अँग्रीमेंट टू सेल ऑफ ऑफिस', '59-हस्तांतरण', 'अँग्रीमेंट टू असाईनमेंट',
    'अपार्टमेंट डीड', 'हस्तांतरणपत्र', 'विक्रीपत्र', 'असाईनमेंट+B45 डीड',
    'फरोक्तखरेदीखत', 'विक्री दाखला', 'हस्तांतरणपत्', 'उलट खरेदीखत',
    'मुदत खरेदीखत', 'मानीव अभिहस्तांतरण', '59हस्तांतरण',
}

LEASE_DOCTYPE = {
    '36-अ-लिव्ह अॅड लायसन्सेस', 'लिव्ह अँणड लायसन्स', 'भाडेपट्टा', 'लीजडीड',
    'किरायानामा', 'अँग्रीमेंट ऑफ ट्रान्स्फर ऑफ टेनन्सी', 'Leave and Licenses',
    'भाडेपट्ट्याचे प्रत्यार्पण', '36-अ-Leave and Licenses', 'भाडेकरार',
    'अँग्रीमेंट टू लीज', 'भाडेपट्ट्याचे हस्तांतरणपत्र', 'असाईनमेंट ऑफ लीज',
    'ट्रान्सफर ऑफ लीज', '36-अ-Leave And Licenses', 'Leave And Licenses',
}

DEDUP_SUBSET = [
    'docno', 'docname', 'registrationdate', 'sroname',
    'propertydescription', 'areaname', 'consideration_amt', 'marketvalue',
]


# =============================================================================
# FUNCTIONS
# =============================================================================

def format_date(x, fmt: str):
    """Format a pd.Timestamp/datetime to string; leave existing strings as-is."""
    if isinstance(x, (pd.Timestamp, datetime)):
        return x.strftime(fmt)
    return x


def load_and_clean(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path)
    df.columns = df.columns.str.lower()
    df = df.dropna(subset=KEEP_COLUMNS, how='all')
    df = df[[c for c in KEEP_COLUMNS if c in df.columns]]
    df = df.dropna(subset=['propertydescription'])
    df['propertydescription'] = df['propertydescription'].str.title()
    return df


def fix_dates(df: pd.DataFrame) -> pd.DataFrame:
    """
    NOTE: Each downloaded IGR file may have mixed date formats.
    Adjust REGISTRATION_DATE_MDY / EXECUTION_DATE_MDY in config above per file.
    """
    reg_fmt  = "%m/%d/%Y" if REGISTRATION_DATE_MDY else "%d/%m/%Y"
    exec_fmt = "%m/%d/%Y" if EXECUTION_DATE_MDY    else "%d/%m/%Y"
    df['registrationdate'] = df['registrationdate'].apply(lambda x: format_date(x, reg_fmt))
    df['dateofexecution']  = df['dateofexecution'].apply(lambda x: format_date(x, exec_fmt))
    return df


def deduplicate(df: pd.DataFrame) -> pd.DataFrame:
    print(f"Before dedup: {df.shape}")
    df = df.drop_duplicates(subset=DEDUP_SUBSET)
    print(f"After dedup:  {df.shape}")
    return df


def categorise(df: pd.DataFrame) -> pd.DataFrame:
    """Tag each row as Sale / Lease / Other using docname sets (O(1) lookup)."""
    conditions = [
        df['docname'].isin(SALE_DOCTYPE),
        df['docname'].isin(LEASE_DOCTYPE),
    ]
    df['property_category'] = np.select(conditions, ['Sale', 'Lease'], default='Other')
    return df


def map_villages(df: pd.DataFrame, village_dir_path: Path) -> pd.DataFrame:
    """Add English village name via the village directory."""
    vdf = pd.read_excel(village_dir_path)
    vdf['Present'] = vdf['Present'].fillna(vdf['village_Name_Eng'])
    mapping = vdf.set_index('village')['Present'].to_dict()
    df['igr_village'] = df['areaname'].map(mapping)
    return df


# =============================================================================
# MAIN
# =============================================================================

def run_stage1() -> pd.DataFrame:
    df = load_and_clean(IGR_EXCEL)
    df = fix_dates(df)
    df = deduplicate(df)
    df = categorise(df)
    df = map_villages(df, VILLAGE_DIR)

    print(f"\nStage 1 complete. Shape: {df.shape}")
    print(df['property_category'].value_counts())

    df.to_excel(OUTPUT_PATH, index=False)
    print(f"Saved → {OUTPUT_PATH}")
    return df


if __name__ == "__main__":
    run_stage1()
