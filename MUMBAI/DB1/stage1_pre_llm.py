# =============================================================================
# DB1 PIPELINE — STAGE 1: Load, Clean, Categorise, Village Map
# =============================================================================
# Input:  Raw IGR Excel (multi-village merged file)
# Output: Cleaned, categorised DataFrame → saved to Excel for LLM processing
# =============================================================================

import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
from dotenv import load_dotenv

from static import result_dict, word_number_dict  # noqa: F401  (kept for parity with original module)

# Load .env from the same directory as this script.
load_dotenv(dotenv_path=Path(__file__).parent / ".env")


# =============================================================================
# CONFIG
# =============================================================================

@dataclass(frozen=True)
class Stage1Config:
    igr_excel: Path = Path(os.getenv("STAGE1_IGR_EXCEL", r".\sample_data.xlsx"))
    village_dir: Path = Path(os.getenv(
        "STAGE1_VILLAGE_DIR",
        r"D:\AI Agent Projects\DATABASE-PIPELINE\PUNE\DB1\Excels Required for DB1\Mumbai IGR Village Directory.xlsx"
    ))
    output_sale: Path = Path(os.getenv("STAGE1_OUTPUT_SALE", r"Sample Sale data for llm.xlsx"))
    output_lease: Path = Path(os.getenv("STAGE1_OUTPUT_LEASE", r"Sample Lease data for llm.xlsx"))
    output_other: Path = Path(os.getenv("STAGE1_OUTPUT_OTHER", r"Sample Other data for llm.xlsx"))

    # NOTE: registrationdate / dateofexecution formats vary per downloaded file.
    #   True  → parse as MM/DD/YYYY (American)
    #   False → parse as DD/MM/YYYY (Indian)
    registration_date_mdy: bool = os.getenv("STAGE1_REGISTRATION_DATE_MDY", "True").strip().lower() != "false"
    execution_date_mdy: bool = os.getenv("STAGE1_EXECUTION_DATE_MDY", "True").strip().lower() != "false"


DEFAULT_CONFIG = Stage1Config()


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

RENAME_MAP = {
    'registrationdate': 'Transaction Date',
    'consideration_amt': 'Agreement Price(INR)',
    'unit_clean': 'Unit No',
    'floor_clean': 'Floor No',
    'docname': 'Document Type',
    'propertydescription': 'Bhumapan',
    'village_Name_Eng': 'IGR Village',
    'sroname': 'SRO Name',
    'docno': 'Document No',
    'marketvalue': 'Bajarbhav',
    'sellerparty': 'Seller Name',
    'purchaserparty': 'Purchaser Name',
}

# Keyword table for property-type classification (checked in dict order).
PROPERTY_TYPE_KEYWORDS = {
    'Apartment': ['अपार्टमेंट', 'अपार्टमेन्ट'],
    'Row_House': ['रो हाऊस', 'रो हाउस'],
    'Bunglow': ['बंगलो', 'बंगला'],
    'Warehouse': ['गोडाऊन'],
    'Industrial': ['इंडस्ट्रियल'],
    'Office': ['ऑफीस', 'ऑॅफीस', 'ऑफिस'],
    'Shop': ['गाळा', 'गाला', 'गाळे', 'शॉप', 'शॉपिंग', 'दुकाने', 'दुकान'],
    'Commercial': ['कमर्शियल', 'कमर्शिअल'],
    'Flat': ['सदनिका', 'सदनिकेचे', 'सदनिक', 'सदनीका', 'सदनिकाचे',
             'फ्लॅट', 'फ़लॅट', 'फ़लॅटचे', 'फलॅट', 'फ्लेट', 'फ्लैट'],
}

UNIT_NO_WORDS = [
    'सदनिका ', 'अपार्टमेंट ', 'अपार्टमेन्ट ', 'फ्लॅट', 'दुकाम', 'ऑॅफीस',
    'ऑफीस', 'ऑफीस', 'ऑफिस', 'युनिट ', 'शॉप', 'प्रिमायसेस', 'दुकान', 'रूम', 'रुम',
]
UNIT_NO_PATTERN = "({}).*?(?:[0-9]+)".format("|".join(UNIT_NO_WORDS))

FLOOR_PATTERN = re.compile(
    ".{15}(मजल्यावर|मजला|मजल्या|माळा|माला|फ्लोअर|मजाल्या|मजलया)"
)


# =============================================================================
# STEP 1 — LOAD / CLEAN
# =============================================================================

def load_and_clean(path: Path) -> pd.DataFrame:
    """Load the raw IGR excel, keep known columns, drop empty rows."""
    df = pd.read_excel(path)
    df.columns = df.columns.str.lower()
    df = df.dropna(subset=KEEP_COLUMNS, how='all')
    df = df[[c for c in KEEP_COLUMNS if c in df.columns]]
    df = df.dropna(subset=['propertydescription'])
    df['propertydescription'] = df['propertydescription'].str.title()
    return df


def _format_date(x, fmt: str):
    """Format a pd.Timestamp/datetime to string; leave existing strings as-is."""
    if isinstance(x, (pd.Timestamp, datetime)):
        return x.strftime(fmt)
    return x


def fix_dates(df: pd.DataFrame, config: Stage1Config = DEFAULT_CONFIG) -> pd.DataFrame:
    """Normalise registrationdate / dateofexecution to consistent string formats.

    NOTE: each downloaded IGR file may mix date formats — set
    config.registration_date_mdy / config.execution_date_mdy per run.
    """
    reg_fmt = "%m/%d/%Y" if config.registration_date_mdy else "%d/%m/%Y"
    exec_fmt = "%m/%d/%Y" if config.execution_date_mdy else "%d/%m/%Y"
    df['registrationdate'] = df['registrationdate'].apply(lambda x: _format_date(x, reg_fmt))
    df['dateofexecution'] = df['dateofexecution'].apply(lambda x: _format_date(x, exec_fmt))
    return df


def deduplicate(df: pd.DataFrame) -> pd.DataFrame:
    print(f"Before dedup: {df.shape}")
    df = df.drop_duplicates(subset=DEDUP_SUBSET)
    print(f"After dedup:  {df.shape}")
    return df


# =============================================================================
# STEP 2 — CATEGORISE / VILLAGE MAP
# =============================================================================

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
# STEP 3 — PROPERTY DESCRIPTION PARSING (floor / unit / property type)
# =============================================================================

def extract_floor_raw(text) -> str | None:
    """Pull the raw floor-related substring out of a property description."""
    try:
        cleaned = re.sub(r'[\'"]', '', str(text))
        cleaned = re.sub(r'\s+', ' ', cleaned)
        match = FLOOR_PATTERN.search(cleaned)
        floor = match.group(0)
    except Exception:
        return None

    try:
        if "माळा" in floor:
            match = re.search("माळा(.{15})", cleaned)
            floor = match.group(0)
    except Exception:
        return None

    return floor


def extract_unit_raw(text) -> str | None:
    """Pull the raw unit/flat-number substring out of a property description."""
    if pd.isna(text):
        return None
    cleaned = re.sub(r'\s+', ' ', text)
    match = re.search(UNIT_NO_PATTERN, cleaned)
    return match.group(0) if match else None


def clean_floor(text) -> str | None:
    """Normalise the raw floor substring into a short floor label."""
    try:
        cleaned = re.sub(r'\s+', ' ', text)

        if "माळा नं" in cleaned:
            match = re.search(r"माळा नं: ([^\s]+)", cleaned)
            return match.group(1).strip(",")

        if "मजला" in cleaned or "मजल्या" in cleaned:
            trimmed = (
                cleaned.replace(' वा', '').replace(' था', '')
                .replace(' रा', '').replace(' ला', '')
                .replace(',', ' ').strip()
            )
            trimmed = trimmed.split(' ')[-2:]
            return ' '.join(trimmed).strip(",")

        if "फ्लोअर" in cleaned:
            parts = cleaned.split(" ")
            idx = parts.index("फ्लोअर")
            return " ".join(parts[idx - 1:idx + 1])

        if "लेवल" in cleaned:
            parts = [p for p in cleaned.split(",") if "लेवल" in p]
            return parts[0].strip(",")

        return cleaned
    except Exception:
        return None


def clean_unit(text) -> str | None:
    """Normalise the raw unit substring down to 'label + number'."""
    if not text:
        return None
    cleaned = re.sub(r'\s+', ' ', text)
    first_segment = str(cleaned).split(",")[0]
    matches = re.findall(r'\d+', first_segment)
    if not matches:
        return None
    last_number = matches[0]
    return first_segment[:first_segment.rindex(last_number) + len(last_number)].strip()


def assign_property_type(property_string, floor_string) -> str | None:
    """Classify a property description into a property type via keyword rules."""
    if not isinstance(property_string, str):
        return None

    property_string = property_string.strip()
    floor_string = floor_string.strip() if isinstance(floor_string, str) else ""

    for prop_type, keywords in PROPERTY_TYPE_KEYWORDS.items():
        if any(keyword in property_string for keyword in keywords):
            return prop_type

    if any(x in property_string for x in ['युनिट', 'यूनिट', 'युनीट']):
        if any(x in property_string for x in ['गाळा', 'दुकान', 'शॉप']):
            return 'Shop'
        return 'Flat' if floor_string else 'Unit'

    if 'चाळ' in property_string:
        return 'Chawl'

    if any(x in property_string for x in ['रूम', 'रुम']):
        return 'Room'

    return None


def enrich_floor_and_unit(df: pd.DataFrame) -> pd.DataFrame:
    """Derive floor / unit / property-type columns from propertydescription."""
    raw_floor = df['propertydescription'].apply(extract_floor_raw)
    raw_unit = df['propertydescription'].apply(extract_unit_raw)

    df['floor_clean'] = raw_floor.apply(clean_floor)
    df['unit_clean'] = raw_unit.apply(clean_unit)

    df['Property Type'] = df.apply(
        lambda row: assign_property_type(row['propertydescription'], row['floor_clean']),
        axis=1,
    )

    # Reduce unit_clean to its numeric component, then reformat as "<Type> no. <n>".
    df['unit_clean'] = (
        df['unit_clean']
        .str.extract(r'(\d+)', expand=False)
        .astype(float)
        .astype(pd.Int64Dtype(), errors='ignore')
    )
    has_unit_and_type = ~df['unit_clean'].isnull() & ~df['Property Type'].isnull()
    df['unit_clean'] = np.where(
        has_unit_and_type,
        df['Property Type'].astype(str) + ' no. ' + df['unit_clean'].astype(str),
        df['unit_clean'].astype(str),
    )

    return df


# =============================================================================
# STEP 4 — FINALISE / EXPORT
# =============================================================================

def rename_and_finalize(df: pd.DataFrame) -> pd.DataFrame:
    df = df.rename(columns=RENAME_MAP)
    return df


def export_by_category(df: pd.DataFrame, config: Stage1Config = DEFAULT_CONFIG) -> None:
    print(f"\nStage 1 complete. Shape: {df.shape}")
    print(df['property_category'].value_counts())

    outputs = {
        'Sale': config.output_sale,
        'Lease': config.output_lease,
        'Other': config.output_other,
    }
    for category, path in outputs.items():
        subset = df[df['property_category'] == category]
        subset.to_excel(path, index=False)
        print(f"Saved → {path}")


# =============================================================================
# MAIN
# =============================================================================

def run_stage1(config: Stage1Config = DEFAULT_CONFIG) -> pd.DataFrame:
    df = load_and_clean(config.igr_excel)
    df = fix_dates(df, config)
    df = deduplicate(df)
    df = categorise(df)
    df = map_villages(df, config.village_dir)
    df = enrich_floor_and_unit(df)
    df = rename_and_finalize(df)
    export_by_category(df, config)
    return df


if __name__ == "__main__":
    run_stage1()