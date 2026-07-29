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
OUTPUT_PATH_SALE = Path(r"Sample Sale data for llm.xlsx")
OUTPUT_PATH_LEASE = Path(r"Sample Lease data for llm.xlsx")
OUTPUT_PATH_OTHER = Path(r"Sample Other data for llm.xlsx")

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


    def floor(string):
        try:

            string=re.sub(r'[\'"]', '', str(string))
            string = re.sub(r'\s+', ' ', string)
            floor = re.search(".{15}(मजल्यावर|मजला|मजल्या|माळा|माला|फ्लोअर|मजाल्या|मजलया)", string)
            floor = floor.group(0)
        except:

            return None
        try:
            if "माळा" in floor:
                floor = re.search("माळा(.{15})", string)
                floor = floor.group(0)
        except:
            return None
        return floor

    def flat_no(string):

        if pd.isna(string):
            return None


        string = re.sub(r'\s+', ' ', string)
        words_list = ['सदनिका ', 'अपार्टमेंट ', 'अपार्टमेन्ट ','फ्लॅट', 'दुकाम', 'ऑॅफीस','ऑफीस','ऑफीस', 'ऑफिस', 'युनिट ', 'शॉप', 'प्रिमायसेस', 'दुकान','रूम','रुम']
        pattern="({}).*?(?:[0-9]+)".format("|".join(words_list))
        match = re.search(pattern, string)
        if match:

            captured_text = match.group(0)
            return captured_text

    df['UNIT_FLOOR']=df['propertydescription'].apply(floor)
    df['UNIT_NO']=df['propertydescription'].apply(flat_no)



    def floor_clean2(string):

        try:
            string = re.sub(r'\s+', ' ', string)
            if "माळा नं" in string:
                majla_string = re.search(r"माळा नं: ([^\s]+)", string)
                majla_string = majla_string.group(1).strip(",")
                return majla_string

            elif "मजला" in string or "मजल्या" in string:
                majla_string = str(string)
                majla_string = (majla_string.replace(' वा','').replace(' था','').replace(' रा','').replace(' ला','').replace(',',' ').strip())
                majla_string = majla_string.split(' ')[-2:]
                majla_string = ' '.join(majla_string).strip(",")
                return majla_string

            elif "फ्लोअर" in string:
                s = string.split(" ")
                index = s.index("फ्लोअर")
                majila_string=" ".join(s[index-1:index+1])
                return majila_string

            elif "लेवल" in string :
                majla_string = str(string)
                majla_string=majla_string.split(",")
                majla_string=[ele for ele in majla_string if "लेवल" in ele]
                return majla_string[0].strip(",")

            else:
                return string
        except:
            return None

    def flat_clean2(string):
        if string:

            string = re.sub(r'\s+', ' ', string)
            string = str(string).split(",")
            string = "".join(string[0])
            matches = re.findall(r'\d+', string)
            if matches:
                last_number = matches[0]
                output = string[:string.rindex(last_number) + len(last_number)].strip()
                return output
        else:
            pass

    df['floor_clean']=df['UNIT_FLOOR'].apply(floor_clean2)
    df['unit_clean']=df['UNIT_NO'].apply(flat_clean2)



    def assign_property_type(property_string, floor_string):

        if not isinstance(property_string, str):
            return None

        property_string = property_string.strip()
        floor_string = floor_string.strip() if isinstance(floor_string, str) else ""

        keywords = {
            'Apartment': ['अपार्टमेंट','अपार्टमेन्ट'],
            'Row_House': ['रो हाऊस','रो हाउस'],
            'Bunglow': ['बंगलो','बंगला'],
            'Warehouse': ['गोडाऊन'],
            'Industrial':['इंडस्ट्रियल'],
            'Office': ['ऑफीस','ऑॅफीस','ऑफिस'],
            'Shop': ['गाळा','गाला','गाळे','शॉप','शॉपिंग','दुकाने','दुकान'],
            'Commercial': ['कमर्शियल','कमर्शिअल'],
            'Flat': ['सदनिका','सदनिकेचे','सदनिक','सदनीका','सदनिकाचे',
                    'फ्लॅट','फ़लॅट','फ़लॅटचे','फलॅट','फ्लेट','फ्लैट']
        }

        for prop_type, prop_keywords in keywords.items():
            for keyword in prop_keywords:
                if keyword in property_string:
                    return prop_type

        if any(x in property_string for x in ['युनिट','यूनिट','युनीट']):
            if any(x in property_string for x in ['गाळा','दुकान','शॉप']):
                return 'Shop'
            if floor_string:
                return 'Flat'
            return 'Unit'

        if 'चाळ' in property_string:
            return 'Chawl'

        if any(x in property_string for x in ['रूम','रुम']):
            return 'Room'

        return None
    
    df['Property Type'] = df.apply(
        lambda row: assign_property_type(row['propertydescription'], row['floor_clean']),
        axis=1
    )


    # Operation for Unit number column
    df['unit_clean'] = df['unit_clean'].str.extract(r'(\d+)', expand=False).astype(float).astype(pd.Int64Dtype(), errors='ignore')


    mask = ~df['unit_clean'].isnull() & ~df['Property Type'].isnull()


    df['unit_clean'] = np.where(mask,
                                        df['Property Type'].astype(str) + ' no. ' + df['unit_clean'].astype(str),
                                        df['unit_clean'].astype(str))

    df.drop(columns=['UNIT_FLOOR', 'UNIT_NO'],inplace=True)
    df.rename(columns={'registrationdate': 'Transaction Date', 'consideration_amt': 'Agreement Price(INR)', 'unit_clean': 'Unit No', 'floor_clean': 'Floor No',
                    'docname': 'Document Type', 'propertydescription': 'Bhumapan','village_Name_Eng': 'IGR Village', 
                    'sroname': 'SRO Name', 'docno': 'Document No', 'marketvalue': 'Bajarbhav','sellerparty':'Seller Name','purchaserparty':'Purchaser Name'}, inplace=True)



    print(f"\nStage 1 complete. Shape: {df.shape}")
    print(df['property_category'].value_counts())
    df_Sale = df[df['property_category'] == 'Sale']
    df_lease = df[df['property_category'] == 'Lease']
    df_other = df[df['property_category'] == 'Other']
    
    # df.to_excel(OUTPUT_PA TH, index=False)
    df_Sale.to_excel(OUTPUT_PATH_SALE, index=False)
    df_lease.to_excel(OUTPUT_PATH_LEASE, index=False)
    df_other.to_excel(OUTPUT_PATH_OTHER, index=False)
    print(f"Saved → {OUTPUT_PATH_SALE}")
    print(f"Saved → {OUTPUT_PATH_LEASE}")
    print(f"Saved → {OUTPUT_PATH_OTHER}")
    return df


if __name__ == "__main__":
    run_stage1()
