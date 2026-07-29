# =============================================================================
# DB1 PIPELINE — STAGE 3: Post-Manual Processing
# =============================================================================
# Input:  Manually reviewed Excel (output of Stage 2, must have
#           llm_processed and manual_processed columns)
# Output: result_df — final cleaned DataFrame ready for DB1
# =============================================================================

import ast
import re
import warnings
from dataclasses import dataclass
from pathlib import Path

import numpy as np
import pandas as pd
from geopy.geocoders import ArcGIS
from joblib import Parallel, delayed

from static import result_dict, word_number_dict

warnings.filterwarnings("ignore")


# =============================================================================
# CONFIG
# =============================================================================

@dataclass(frozen=True)
class Stage3Config:
    city: str = "mumbai"  # "mumbai" | "thane" | "pune"
    input_path: Path = Path(r"manual processed Sale data for manual.xlsx")
    rera_grand_path: Path = Path(r".\Excels Required for DB1\mumbai RERA GRAND EXCEL VERSION.xlsx")
    rera_keywords_path: Path = Path(r".\Excels Required for DB1\RERA_All_Keywords_BHK_Prop_Type.xlsx")
    coordinates_path: Path = Path(r".\Excels Required for DB1\project address and its coordinates.xlsx")
    village_dir_path: Path = Path(r".\Excels Required for DB1\Mumbai IGR Village Directory.xlsx")
    postal_csv_path: Path = Path(r".\Excels Required for DB1\postal_pincode.csv")
    output_path: Path = Path(r"sample_file_for_db1.xlsx")
    bhk_max_diff: float = 5


DEFAULT_CONFIG = Stage3Config()


# =============================================================================
# CONSTANTS
# =============================================================================

SKIP_BHK = {'UNDEFINED FLATS', 'SHOP', 'OFFICE', 'OTHERS'}
NULL_KEYS = {'', 'NA', 'NULL'}

# Built-Up -> Carpet divisor is fixed across cities.
BUILDUP_TO_CARPET_DIVISOR = 1.2

# Saleable -> Carpet divisor varies by city.
SALEABLE_TO_CARPET_DIVISOR_BY_CITY = {
    "mumbai": 1.45,
    "thane": 1.4,
    "pune": 1.35,
}

SRO_DICT = {
    "thane": {
        "दु.नि. ठाणे 1": "73",       "सह दु.नि.ठाणे 2": "74",
        'सह दु.नि. ठाणे 3': "75",    'सह दु.नि. ठाणे 4': "76",
        "सह दु.नि.ठाणे 5": '335',    'सह दु.नि.ठाणे 6': "336",
        'सह दु.नि.ठाणे 7': "337",    'सह दु.नि. ठाणे 8': "392",
        "दु.नि. ठाणे 9": "536",      'सह दु.नि. ठाणे 10': "393",
        'सह दु.नि. ठाणे 11': "394",  "सह दु.नि.ठाणे 12": "530",
        'दु.नि. कल्याण 1': '70',     'सह दु.नि. कल्याण 2': "71",
        'सह दु.नि. कल्याण 3': "72",  'सह दु.नि.कल्याण 4': "338",
        'सह दु.नि. कल्याण 5': "507", "दु.नि. उल्हासनगर 1": "77",
        "सह दु.नि. उल्हासनगर 4": "541", "दु.नि. वसई 1": "79",
        "सह दु.नि. वसई 3": "350",    "सह दु.नि. वसई 6": "535",
        "सह दु.नि. वसई 4": "533",    "सह दु.नि. वसई 2": "80",
        "सह दु.नि. वसई 5": "534",
    },
    "mumbai": {
        'दु.नि.मुंबई शहर 1': '318',       'सह दु.नि.मुंबई शहर 2': '319',
        'सह दु. नि. मुंबई शहर 3': '450',  'सह दु.नि.मुंबई शहर 4': '508',
        'सह दु.नि.मुंबई शहर 5': '509',    'सह दु.नि. अंधेरी 1': '322',
        'सह दु.नि. अंधेरी 2': '323',      'सह दु.नि. बोरीवली 1': '324',
        'सह दु.नि. बोरीवली 2': '367',     'सह दु.नि. बोरीवली 3': '368',
        'सह दु.नि. कुर्ला 1': '369',      'सह दु.नि. कुर्ला 2': '370',
        'सह दु.नि. अंधेरी 3': '378',      'सह दु.नि. बोरीवली 4': '387',
        'सह दु.नि. बोरीवली 5': '388',     'सह दु.नि. बोरीवली 6': '389',
        'सह दु.नि. कुर्ला 3': '390',      'सह दु.नि. कुर्ला 4': '391',
        'सह दु.नि. अंधेरी 4': '401',      'सह दु.नि. बोरीवली 7': '451',
        'सह दु.नि. अंधेरी 5': '512',      'सह दु.नि. अंधेरी 6': '513',
        'सह दु.नि. अंधेरी 7': '514',      'सह दु.नि.बोरीवली 8': '516',
        'सह दु.नि.बोरीवली 9': '517',      'सह दु.नि.कुर्ला 5': '520',
        "सह दु.नि.मुंबई 8": "322",        "Joint S.R. Mumbai 8": "322",
        "सह दु.नि.मुंबई 9": "323",        "Joint S.R. Mumbai 9": "323",
        "सह दु.नि.मुंबई 10": "378",       "Joint S.R. Mumbai 10": "378",
        "सह दु.नि.मुंबई 11": "401",       "Joint S.R. Mumbai 11": "401",
        "सह दु.नि.मुंबई 12": "512",       "Joint S.R. Mumbai 12": "512",
        "सह दु.नि.मुंबई 13": "513",       "Joint S.R. Mumbai 13": "513",
        "सह दु.नि.मुंबई 14": "514",       "Joint S.R. Mumbai 14": "514",
        "Joint S.R. Mumbai 15": "515",     "सह दु.नि.मुंबई 16": "324",
        "Joint S.R. Mumbai 16": "324",     "सह दु.नि.मुंबई 17": "367",
        "Joint S.R. Mumbai 17": "367",     "सह दु.नि.मुंबई 18": "368",
        "Joint S.R. Mumbai 18": "368",     "सह दु.नि.मुंबई 19": "387",
        "Joint S.R. Mumbai 19": "387",     "सह दु.नि.मुंबई 20": "388",
        "Joint S.R. Mumbai 20": "388",     "सह दु.नि.मुंबई 21": "389",
        "Joint S.R. Mumbai 21": "389",     "सह दु.नि.मुंबई 22": "451",
        "Joint S.R. Mumbai 22": "451",     "सह दु.नि.मुंबई 23": "516",
        "Joint S.R. Mumbai 23": "516",     "सह दु.नि.मुंबई 24": "517",
        "Joint S.R. Mumbai 24": "517",     "Joint S.R. Mumbai 25": "518",
        "Joint S.R. Mumbai 26": "519",     "सह दु.नि.मुंबई 27": "369",
        "Joint S.R. Mumbai 27": "369",     "सह दु.नि.मुंबई 28": "370",
        "Joint S.R. Mumbai 28": "370",     "सह दु.नि.मुंबई 29": "390",
        "Joint S.R. Mumbai 29": "390",     "सह दु.नि.मुंबई 30": "391",
        "Joint S.R. Mumbai 30": "391",     "सह दु.नि.मुंबई 31": "520",
        "Joint S.R. Mumbai 31": "520",     "Joint S.R. Mumbai 32": "521",
    },
    "pune": {
        'सह दु.नि.हवेली 24': '525',        'सह दु.नि.हवेली 25': '526',
        'सह दु.नि. वडगांव-मावळ-२': '454',  'सह दु.नि. हवेली 17': '385',
        'दु.नि.हवेली 1': '1',              'सह दु.नि.हवेली 23': '524',
        'सह दु.नि. हवेली 10': '326',       'सह दु.नि. हवेली 4': '4',
        'सह दु.नि.हवेली 22': '523',        'सह दु.नि. हवेली 11': '329',
        'सह दु.नि. हवेली 18': '386',       'सह दु.नि. हवेली 8': '8',
        'सह दु.नि. हवेली 15': '333',       'सह दु.नि.हवेली 21': '522',
        'सह दु.नि. हवेली 9': '9',          'सह दु.नि. हवेली 2': '2',
        'सह दु.नि. हवेली 13': '331',       'सह दु.नि. हवेली 26': '527',
        'सह दु.नि. मुळशी-२': '453',        'सह दु.नि. हवेली 19': '396',
        'सह दु.नि. हवेली 12': '330',       'सह दु.नि. हवेली 7': '7',
        'सह दु.नि. हवेली 5': '5',          'सह दु.नि. हवेली 20': '397',
        'सह दु.नि. हवेली 14': '332',       'सह दु.नि. हवेली 6': '6',
        'दु.नि. मुळशी': '18',             'सह दु.नि. हवेली 27': '540',
        'सह दु.नि. हवेली 3': '3',          'सह दु.नि. हवेली 16': '334',
        'सह दु.नि. खेड-3': '545',          'दु.नि. खेड': '16',
        'सह दु.नि. खेड-२': '452',          "दु.नि. मावळ": "17",
        "सह दु.नि. लोणावळा": "427",
    },
}


# =============================================================================
# UTILITY FUNCTIONS
# =============================================================================

def standardise_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Replace special characters in column names with underscores, lowercase."""
    df.columns = [
        re.sub(r'[()/\s.,*-]', '_', col).strip('_').lower()
        for col in df.columns
    ]
    return df


def convert_area_to_sqmt(x):
    """Convert Marathi area unit strings to sq metres."""
    if isinstance(x, str):
        for unit, factor in [
            ("आर.चौ.मीटर", 1), ("चौ.मीटर", 1),
            ("चौ.फूट", 1 / 10.764), ("हेक्टर . आर", 10_000),
        ]:
            if unit in x:
                try:
                    return round(float(x.replace(unit, "")) * factor, 2)
                except ValueError:
                    return x
    return x


def safe_parse(x):
    """Parse a stringified Python literal safely; return None on failure."""
    if pd.isna(x):
        return None
    if isinstance(x, str):
        x = x.replace("''", "'")
        try:
            return ast.literal_eval(x)
        except Exception as e:
            print(f"Error parsing: {str(x)[:80]} | {e}")
    return x


def extract_pincode(text):
    if isinstance(text, str):
        m = re.search(r'\b\d{6}\b', text)
        return m.group(0) if m else None
    return None


# =============================================================================
# STAGE 3.0 — Net carpet area cascade (Carpet → BuildUp → Saleable → Other)
# =============================================================================

def derive_net_carpet_area(df: pd.DataFrame, city: str) -> pd.DataFrame:
    """Fill net_carpet_area_sqmt from whichever area column is available,
    preferring Carpet, then BuildUp (÷1.2, fixed), then Saleable
    (÷ city-specific divisor), then Other.
    """
    saleable_divisor = SALEABLE_TO_CARPET_DIVISOR_BY_CITY.get(city.lower())
    if saleable_divisor is None:
        raise ValueError(
            f"No Saleable->Carpet divisor configured for city '{city}'. "
            f"Known cities: {list(SALEABLE_TO_CARPET_DIVISOR_BY_CITY)}"
        )

    df['net_carpet_area_sqmt'] = np.nan

    mask = df['Carpet_Area_sqmt'].notna()
    df.loc[mask, 'net_carpet_area_sqmt'] = df.loc[mask, 'Carpet_Area_sqmt'].values

    mask = df['net_carpet_area_sqmt'].isna() & df['BuildUp_Area_sqmt'].notna()
    df.loc[mask, 'net_carpet_area_sqmt'] = (
        df.loc[mask, 'BuildUp_Area_sqmt'].values / BUILDUP_TO_CARPET_DIVISOR
    )

    mask = df['net_carpet_area_sqmt'].isna() & df['Saleable_Area_sqmt'].notna()
    df.loc[mask, 'net_carpet_area_sqmt'] = (
        df.loc[mask, 'Saleable_Area_sqmt'].values / saleable_divisor
    )

    mask = df['net_carpet_area_sqmt'].isna() & df['Other_Area_sqmt'].notna()
    df.loc[mask, 'net_carpet_area_sqmt'] = df.loc[mask, 'Other_Area_sqmt'].values

    df['net_carpet_area_sqmt'] = df['net_carpet_area_sqmt'].round(2)
    return df


# =============================================================================
# STAGE 3.1 — Rename, unit/floor, column standardisation
# =============================================================================

def rename_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.rename(columns={
        'areaname': 'village',
        'srocode': 'sro_code',
        'Bhumapan': 'property_description',
        'Property Type': 'property_type_raw',
        'Agreement Price(INR)': 'agreement_price',
        'Terrace_Area_sqmt': 'terrace_sqmt',
        'Balcony_Area_sqmt': 'balcony_sqmt',
        "Modified_Project_Name_1": "project_name",
    })
    return df


def process_unit_and_floor(df: pd.DataFrame) -> pd.DataFrame:
    """
    Build 'Flat no. 901' strings from unit number and property_type_raw.
    Uses np.where to avoid Int64 / object dtype conflicts.
    """
    df['Transaction Type'] = df['Document Type'].map(result_dict.get)

    unit_numeric = (
        df['Unit No'].astype(str)
        .str.extract(r'(\d+)', expand=False)
        .pipe(pd.to_numeric, errors='coerce')
        .astype(pd.Int64Dtype())
    )

    mask = unit_numeric.notna() & df['property_type_raw'].notna()

    df['Unit No'] = np.where(
        mask,
        df['property_type_raw'].astype(str) + ' no. ' + unit_numeric.astype(str),
        unit_numeric.astype(str).replace('<NA>', None),
    )

    df['Floor No'] = df['Floor No'].map(word_number_dict)
    return df


def select_required_columns(df: pd.DataFrame) -> pd.DataFrame:
    required = [
        'project_name', 'village', 'igr_village', 'location',
        'sro_code', 'sro_name', 'document_no', 'transaction_type',
        'agreement_price', 'bajarbhav', 'property_description',
        'transaction_date', 'floor_no', 'unit_no', 'property_type_raw',
        'net_carpet_area_sqmt', 'balcony_sqmt', 'terrace_sqmt',
        'seller_name', 'purchaser_name', 'property_category',
        'llm_processed', 'manual_processed', 'internaldocumentnumber',
        'micrno', 'bank_type', 'party_code', 'dateofexecution',
        'stampdutypaid', 'registrationfees',
    ]
    return df[[c for c in required if c in df.columns]]


# =============================================================================
# STAGE 3.2 — Split & clean sale data
# =============================================================================

def split_by_category(df: pd.DataFrame):
    lease = df[df['property_category'] == 'Lease'].copy()
    other = df[df['property_category'] == 'Other'].copy()
    sale = df[df['property_category'] == 'Sale'].copy()
    print(f"Sale: {len(sale)} | Lease: {len(lease)} | Other: {len(other)}")
    return sale, lease, other


def _infer_floor_from_unit(df: pd.DataFrame) -> pd.DataFrame:
    """Infer missing floor_no from unit number pattern e.g. '904' -> floor 9."""
    wrong = []
    for unit in df[df['floor_no'].isna()]['unit_no'].dropna().unique():
        try:
            floor_str = str(unit).split("no.")[-1].strip()
            if len(floor_str) not in (3, 4):
                continue
            floor_int = int(floor_str)
            floor_no = floor_int // 100
            unit_pos = floor_int % 100
            if unit_pos < 50:
                mask = df['floor_no'].isna() & (df['unit_no'] == unit)
                df.loc[mask, 'floor_no'] = floor_no
        except Exception as e:
            wrong.append((unit, e))
    if wrong:
        print(f"[WARN] Floor inference failed for {len(wrong)} units")
    return df


def clean_sale_data(dataframe: pd.DataFrame) -> pd.DataFrame:
    print(f"Before dedup: {dataframe.shape}")
    dataframe['property_description'] = dataframe['property_description'].str.title()
    print("columns in dataframe : ", dataframe.columns)

    dataframe = dataframe.drop_duplicates(
        subset=['village', 'sro_name', 'document_no', 'transaction_type',
                'property_description', 'transaction_date', 'agreement_price'],
        keep='first'
    )
    print(f"After dedup:  {dataframe.shape}")

    dataframe['net_carpet_area_sqmt'] = pd.to_numeric(
        dataframe['net_carpet_area_sqmt'], errors='coerce'
    )
    print(f"Carpet < 5:    {(dataframe['net_carpet_area_sqmt'] < 5).sum()}")
    print(f"Carpet > 3000: {(dataframe['net_carpet_area_sqmt'] > 3000).sum()}")
    print(f"Carpet NaN:    {dataframe['net_carpet_area_sqmt'].isna().sum()}")

    dataframe = _infer_floor_from_unit(dataframe)

    dataframe['location'] = dataframe['igr_village']
    dataframe['transaction_type'] = np.where(
        dataframe['transaction_type'].isna(), 'Sale Agreement', dataframe['transaction_type']
    )

    return dataframe[[
        'project_name', 'village', 'location', 'igr_village',
        'sro_code', 'sro_name', 'document_no', 'transaction_type',
        'agreement_price', 'bajarbhav', 'property_description', 'transaction_date',
        'floor_no', 'unit_no', 'property_type_raw', 'net_carpet_area_sqmt',
        'balcony_sqmt', 'terrace_sqmt', 'seller_name', 'purchaser_name',
        'internaldocumentnumber', 'micrno', 'bank_type', 'party_code',
        'dateofexecution', 'stampdutypaid', 'registrationfees',
        'property_category', 'llm_processed', 'manual_processed',
    ]]


# =============================================================================
# STAGE 3.3 — RERA index
# =============================================================================

def _get_rera_values(rera_grand, project, village):
    """Exact then list-fallback RERA lookup for a project+village pair."""
    mask = (
        (rera_grand['modified_project_name'] == project) &
        (rera_grand['rera_location'] == village)
    )
    match = rera_grand[mask]

    if match.empty:
        def _in_list(x):
            try:
                return isinstance(x, str) and x.startswith('[') and village in ast.literal_eval(x)
            except Exception:
                return False
        mask = (
            (rera_grand['modified_project_name'] == project) &
            rera_grand['rera_location_v1'].apply(_in_list)
        )
        match = rera_grand[mask]

    if not match.empty:
        r = match.iloc[0]
        return r['index'], r['project_lat'], r['project_lng']
    return None, None, None


def assign_rera_index(final_village: pd.DataFrame,
                       rera_grand: pd.DataFrame, city: str) -> pd.DataFrame:
    print("[RERA] Assigning index...")

    # Safe string normalisation — handles NaN and non-string values
    for col in ['igr_village', 'location', 'project_name']:
        final_village[col] = (
            final_village[col].astype(str)
            .str.title().str.strip().str.strip("\n")
            .replace('Nan', np.nan)
        )

    final_village['location'] = final_village['location'].fillna(final_village['igr_village'])

    for col in ['modified_project_name', 'rera_location', 'rera_location_v1']:
        rera_grand[col] = (
            rera_grand[col].astype(str)
            .str.strip().str.title()
            .replace('Nan', np.nan)
        )

    rera_grand = rera_grand.sort_values('index')
    print(f"  Duplicate RERA entries: {rera_grand.duplicated(subset=['modified_project_name', 'rera_location']).sum()}")
    rera_grand.drop_duplicates(subset=['modified_project_name', 'rera_location'], inplace=True)

    # Exact merge on project_name + location
    final_village = final_village.merge(
        rera_grand[['index', 'modified_project_name', 'rera_location', 'project_lat', 'project_lng']],
        left_on=['project_name', 'location'],
        right_on=['modified_project_name', 'rera_location'],
        how='left'
    )

    # Fuzzy fallback for rows still missing index
    lookup = {}
    unmatched = final_village[
        final_village['project_name'].notna() & final_village['index'].isna()
    ].groupby(['project_name', 'location'])

    for (project, village), group in unmatched:
        if (project, village) not in lookup:
            lookup[(project, village)] = _get_rera_values(rera_grand, project, village)
        idx, lat, lng = lookup[(project, village)]
        final_village.loc[group.index, 'index'] = idx
        final_village.loc[group.index, 'project_lat'] = lat
        final_village.loc[group.index, 'project_lng'] = lng

    # Assign non-RERA index codes for manually processed rows
    final_village['index'] = pd.to_numeric(final_village['index'], errors='coerce')
    final_village['city'] = city.title()

    non_rera = (
        final_village[
            final_village['index'].isna() & (final_village['manual_processed'] == 'Yes')
        ][['project_name', 'igr_village']]
        .drop_duplicates()
        .reset_index(drop=True)
    )

    prefix = (city.lower()[0] + "NR").lower()
    non_rera['index'] = [f"{prefix}{str(i).zfill(3)}" for i in range(101, 101 + len(non_rera))]

    merged_df = final_village.merge(
        non_rera[['project_name', 'igr_village', 'index']],
        on=['project_name', 'igr_village'],
        how='left', suffixes=('', '_grouped')
    )
    merged_df['index'] = merged_df['index'].fillna(merged_df.pop('index_grouped'))
    merged_df['index'] = merged_df['index'].apply(
        lambda x: str(int(float(x))) if str(x).replace('.', '').isdigit() else x
    )
    return merged_df


# =============================================================================
# STAGE 3.4 — Geocode
# =============================================================================

def geocode_coordinates(merged_df: pd.DataFrame,
                         coord_path: Path, city: str) -> pd.DataFrame:
    print("[GEO] Filling project coordinates...")
    coordinate_sheet = pd.read_excel(coord_path)
    coordinate_sheet.columns = coordinate_sheet.columns.str.lower()

    merged_df['api_call_input'] = (
        merged_df['project_name'].fillna('') + ", " +
        merged_df['igr_village'].fillna('') + ", " + city
    )

    village_df = merged_df.merge(
        coordinate_sheet[['api_call_input', 'project_lat_x', 'project_lng_x']],
        on='api_call_input', how='left'
    )
    village_df['project_lat'] = village_df['project_lat'].fillna(village_df['project_lat_x'])
    village_df['project_lng'] = village_df['project_lng'].fillna(village_df['project_lng_x'])
    village_df.drop(columns=['project_lat_x', 'project_lng_x'], inplace=True)

    remaining_missing = village_df[village_df['project_lat'].isna()]['api_call_input'].unique()
    print(f"    Still missing: {len(remaining_missing)}")

    if len(remaining_missing):
        geolocator = ArcGIS()
        new_rows = []
        for loc in remaining_missing:
            location = geolocator.geocode(loc)
            if location:
                lat, lng = location.latitude, location.longitude
                village_df.loc[village_df['api_call_input'] == loc, ['project_lat', 'project_lng']] = lat, lng
                new_rows.append({
                    'api_call_input': loc, 'project_lat_x': lat,
                    'project_lng_x': lng, 'manually corrected': 'No'
                })
            else:
                print(f"    [WARN] Could not geocode: {loc}")
        if new_rows:
            coordinate_sheet = pd.concat([coordinate_sheet, pd.DataFrame(new_rows)], ignore_index=True)
            coordinate_sheet.to_excel(coord_path, index=False)
            print(f"    Cache updated -> {coord_path.name}")

    return village_df


# =============================================================================
# STAGE 3.5 — Property type
# =============================================================================

def _map_property_type(value):
    if value in ['Flat', 'Apartment', 'Flat/Shop', 'Duplex']:
        return 'Flat'
    elif value in ['Shop', 'Showroom']:
        return 'Shop'
    elif value == 'Office':
        return 'Office'
    return 'Others'


def classify_property_type(village_df: pd.DataFrame) -> pd.DataFrame:
    village_df['property_type_raw'] = (
        village_df['property_type_raw'].str.strip('\n').str.title().str.strip()
    )
    village_df['property_type'] = village_df['property_type_raw'].apply(_map_property_type)
    print(village_df['property_type'].value_counts().to_string())
    return village_df


# =============================================================================
# STAGE 3.6 — BHK
# =============================================================================

def combine_columns(bhk_ca):
    """Flatten BHK-wise carpet area nested structure into {BHK_KEY: [area_list]}."""
    try:
        if isinstance(bhk_ca, float):
            return None
        result_local = {}
        try:
            old_ca_list = eval(str(bhk_ca).lower())   # noqa: S307
        except Exception:
            old_ca_list = str(bhk_ca).lower()
        try:
            old_ca_list = eval(old_ca_list)            # noqa: S307
        except Exception:
            pass
        try:
            old_ca_list = eval(old_ca_list)            # noqa: S307
        except Exception:
            pass

        for values in old_ca_list.values():
            try:
                for item in eval(values):              # noqa: S307
                    for key in item:
                        new_key = key.replace(" bhk", "bhk").upper()
                        result_local.setdefault(new_key, []).extend(item[key])
            except Exception:
                try:
                    for item in eval(values)[0]:       # noqa: S307
                        for key in item:
                            new_key = key.replace(" bhk", "bhk").upper()
                            result_local.setdefault(new_key, [])
                except Exception:
                    pass
        return str(result_local)
    except Exception:
        print(bhk_ca)
        return None


def _normalize_bhk_key(key, keyword_map):
    k = key.strip().upper()
    if k in NULL_KEYS:
        return 'OTHERS'
    if k in SKIP_BHK:
        return k
    return keyword_map.get(k, k)


def _clean_bhk_values(value_list):
    out = []
    for v in value_list:
        s = re.sub(r'(\d+\.\d+)\.', r'\1', str(v))
        s = re.sub(r'(\d+)\.\.(\d+)', r'\1.\2', s)
        try:
            out.append(round(float(s), 2))
        except ValueError:
            pass
    return out


def _find_closest_bhk(carpet, building_info, max_diff):
    best_bhk, best_val, best_diff = None, None, float('inf')
    for bhk_type, values in building_info.items():
        if not values:
            continue
        arr = np.array(values, dtype=np.float64)
        diffs = np.abs(arr - carpet)
        idx = diffs.argmin()
        diff = float(diffs[idx])
        if diff < best_diff and diff <= max_diff:
            best_diff = diff
            best_val = arr[idx]
            best_bhk = bhk_type
    return best_bhk, (float(best_val) if best_val is not None else None)


def assign_bhk_carpet_match(village_df: pd.DataFrame,
                             rera_grand: pd.DataFrame,
                             rera_keywords: pd.DataFrame,
                             bhk_max_diff: float = DEFAULT_CONFIG.bhk_max_diff) -> pd.DataFrame:
    """Stages 1+2+3: exact -> closest -> skip-type retry BHK matching."""

    keyword_map = rera_keywords.set_index('Keywords')['BHK'].to_dict()

    def normalize_key(key):
        return _normalize_bhk_key(key, keyword_map)

    def process_row(i, carpet_raw, building_raw, parse_cache):
        try:
            carpet = round(float(carpet_raw), 2)
            if np.isnan(carpet):
                return i, None
        except (ValueError, TypeError):
            return i, None

        building_info = parse_cache.get(building_raw)
        if not building_info:
            return i, None

        cleaned = {k: _clean_bhk_values(v) for k, v in building_info.items()}
        bhk = None

        for key, values in cleaned.items():         # Stage 1: exact match
            if carpet in values:
                bhk = normalize_key(key)

        if bhk is None:                             # Stage 2: closest match
            raw_bhk, _ = _find_closest_bhk(carpet, cleaned, bhk_max_diff)
            if raw_bhk is not None:
                bhk = normalize_key(raw_bhk)

        if bhk and bhk.upper() in SKIP_BHK:        # Stage 3: retry without skip types
            filtered = {
                normalize_key(k): v for k, v in cleaned.items()
                if k.strip().upper() not in (SKIP_BHK | NULL_KEYS)
                and normalize_key(k).upper() not in SKIP_BHK
            }
            if filtered:
                new_bhk, _ = _find_closest_bhk(carpet, filtered, bhk_max_diff)
                bhk = new_bhk if (new_bhk and new_bhk.upper() not in SKIP_BHK) else None
            else:
                bhk = None

        final = bhk.upper() if bhk else None
        return i, (None if final in SKIP_BHK else final)

    # Extract first number from mixed strings
    village_df['net_carpet_area_sqmt'] = village_df['net_carpet_area_sqmt'].apply(
        lambda x: re.findall(r"[-+]?\d*\.?\d+|\d+", str(x))[0] if not isinstance(x, float) else x
    )

    rera_grand['index'] = rera_grand['index'].astype(str)
    village_df = village_df.merge(
        rera_grand[['index', 'building_wise_carpet_area']], on='index', how='left'
    )
    village_df.sort_values('index', inplace=True)
    village_df['BHK'] = None

    mask = village_df['building_wise_carpet_area'].notna() & (village_df['property_type'] == 'Flat')
    target = village_df[mask][['net_carpet_area_sqmt', 'building_wise_carpet_area']].copy()

    parse_cache = {}
    for raw in target['building_wise_carpet_area'].unique():
        try:
            parse_cache[raw] = ast.literal_eval(raw)
        except (ValueError, SyntaxError):
            parse_cache[raw] = None

    print("Pre-parsing building info...")
    print(f"Processing {len(target)} rows, {len(parse_cache)} unique building configs...")

    results = Parallel(n_jobs=-1, backend='threading', verbose=1)(
        delayed(process_row)(i, row['net_carpet_area_sqmt'], row['building_wise_carpet_area'], parse_cache)
        for i, row in target.iterrows()
    )
    valid = {i: bhk for i, bhk in results if bhk is not None}
    print(f"Matched {len(valid)} / {len(target)} rows")
    village_df.loc[list(valid.keys()), 'BHK'] = list(valid.values())

    village_df.loc[village_df['BHK'].isin(SKIP_BHK | {'FLAT'}), 'BHK'] = None
    return village_df


def _build_bhk_phase_data(rera_grand: pd.DataFrame) -> pd.DataFrame:
    phase_data_list = []
    for _, row in rera_grand.iterrows():
        data_str = row['carpet_wise_total_sold_units']
        data_list = data_str if isinstance(data_str, dict) else {}
        if isinstance(data_str, str):
            try:
                data_list = ast.literal_eval(data_str)
            except Exception:
                pass
        for phase, phase_data in data_list.items():
            if isinstance(phase_data, list) and phase_data:
                for item in phase_data:
                    phase_data_list.append({
                        "modified_project_name": row["modified_project_name"],
                        "Rera_Location": row["rera_location"],
                        "Phase": phase, "Data": item,
                    })
    return pd.DataFrame(phase_data_list)


def _reshape_bhk_row(row):
    new_rows = []
    if isinstance(row["Data"], dict):
        for bhk, data in row["Data"].items():
            if isinstance(data, dict):
                for carpet, values in data.items():
                    if isinstance(values, list) and len(values) == 2:
                        new_rows.append({
                            "modified_project_name": row["modified_project_name"],
                            "Rera_Location": row["Rera_Location"],
                            "Phase": row["Phase"],
                            "BHK": bhk, "carpet_sqmt": carpet,
                        })
    return pd.DataFrame(new_rows) if new_rows else None


def assign_bhk_range_fallback(village_df: pd.DataFrame,
                               rera_grand: pd.DataFrame,
                               rera_keywords: pd.DataFrame) -> pd.DataFrame:
    """Stage 4: Percentile range fallback for Flat rows still without BHK."""
    try:
        rera_grand['carpet_wise_total_sold_units'] = rera_grand['carpet_wise_total_sold_units'].apply(safe_parse)
        rera_grand['index'] = rera_grand['index'].astype(int)
        rera_keywords = rera_keywords.apply(lambda c: c.str.upper().str.strip() if c.dtype == object else c)
        village_df['BHK'] = village_df['BHK'].str.title()

        df2 = _build_bhk_phase_data(rera_grand)
        if df2.empty:
            print("[WARN] No phase data — skipping range fallback")
            return village_df

        df_expanded = pd.concat(df2.apply(_reshape_bhk_row, axis=1).dropna().tolist(), ignore_index=True)

        df_expanded["BHK"] = df_expanded["BHK"].str.replace(" ", "").str.upper()
        rera_keywords["Keywords"] = rera_keywords["Keywords"].str.replace(" ", "")
        bhk_mapping = dict(zip(rera_keywords["Keywords"], rera_keywords["Final BHK"]))
        df_expanded["BHK_Modified"] = df_expanded["BHK"].map(bhk_mapping).fillna("UNDEFINED OTHERS")

        df_grouped = (
            df_expanded.groupby("BHK_Modified")["carpet_sqmt"]
            .agg(lambda x: [float(i) for i in x if str(i).replace('.', '', 1).isdigit()])
            .reset_index()
        )
        df_grouped = df_grouped[df_grouped["BHK_Modified"].isin(["1BHK", "2BHK", "3BHK"])].reset_index(drop=True)

        if len(df_grouped) < 3:
            print("[WARN] Not enough BHK rows for percentile ranges")
            return village_df

        df_grouped["p10"] = df_grouped["carpet_sqmt"].apply(
            lambda x: round(np.percentile(x, 10), 2) if x else None
        )
        df_grouped["p90"] = df_grouped["carpet_sqmt"].apply(
            lambda x: round(np.percentile(x, 90), 2) if x else None
        )

        ranges = {
            "<1BHK": (0, df_grouped.loc[0, "p10"]),
            "1BHK": (df_grouped.loc[0, "p10"], (df_grouped.loc[0, "p90"] + df_grouped.loc[1, "p10"]) / 2),
            "2BHK": ((df_grouped.loc[0, "p90"] + df_grouped.loc[1, "p10"]) / 2, (df_grouped.loc[1, "p90"] + df_grouped.loc[2, "p10"]) / 2),
            "3BHK": ((df_grouped.loc[1, "p90"] + df_grouped.loc[2, "p10"]) / 2, df_grouped.loc[2, "p90"]),
            ">3BHK": (df_grouped.loc[2, "p90"], float("inf")),
        }
        for label, (lo, hi) in ranges.items():
            print(f"  {label}: {round(lo, 2)} -> {round(hi, 2)}")

        def assign_bhk_range(carpet_area):
            for bhk, (low, high) in ranges.items():
                if float(low) <= carpet_area < float(high):
                    return bhk
            return None

        village_df['BHK'] = village_df.apply(
            lambda row: assign_bhk_range(row['net_carpet_area_sqmt'])
            if pd.isna(row['BHK']) and row['property_type'] == 'Flat'
            else row['BHK'],
            axis=1
        )

    except Exception as e:
        print(f"[WARN] BHK range fallback failed: {e}")

    return village_df


def finalise_bhk(village_df: pd.DataFrame) -> pd.DataFrame:
    village_df['BHK'] = village_df['BHK'].astype(str).str.title().str.strip()
    village_df['BHK'] = village_df['BHK'].replace({'None': None, 'Nan': None, '': None})
    village_df['BHK'] = np.where(village_df['BHK'].isna(), village_df['property_type'], village_df['BHK'])
    village_df['BHK'] = village_df['BHK'].str.upper()

    village_df['unit_no'] = village_df['unit_no'].astype(str)
    village_df.loc[village_df['unit_no'] == 'nan', 'unit_no'] = None
    return village_df


# =============================================================================
# STAGE 3.7 — Buyer pincode
# =============================================================================

def add_buyer_location(village_df: pd.DataFrame, postal_csv_path: Path) -> pd.DataFrame:
    print("[POSTAL] Adding buyer pincode & location...")
    village_df['buyer_pincode'] = village_df['purchaser_name'].apply(extract_pincode)

    postal = (
        pd.read_csv(postal_csv_path)
        .drop_duplicates(subset='Pincode', keep='first')
        .rename(columns={
            'OfficeName_P': 'Locality of buyer',
            'District_P': 'District',
            'StateName_P': 'StateName',
            'Pincode': 'buyer_pincode',
        })
    )
    village_df['buyer_pincode'] = village_df['buyer_pincode'].astype('Int64')
    return pd.merge(village_df, postal, on='buyer_pincode', how='left')


# =============================================================================
# STAGE 3.8 — Final assembly
# =============================================================================

def final_assembly(df_merged: pd.DataFrame, lease_data: pd.DataFrame,
                    other_data: pd.DataFrame, city: str,
                    village_dir_path: Path) -> pd.DataFrame:
    print("[FINAL] Assembling result...")

    df_merged = pd.concat([df_merged, lease_data, other_data], ignore_index=True)

    df_merged['transaction_date'] = pd.to_datetime(df_merged['transaction_date'], dayfirst=True)
    quarter = df_merged['transaction_date'].dt.quarter.astype("Int64")
    year = df_merged['transaction_date'].dt.year.astype("Int64")
    df_merged['quarter'] = "Q" + quarter.astype(str) + "-" + year.astype(str)
    df_merged['year'] = year

    df_merged['Tower'] = None
    df_merged['gross_carpet_sqft'] = None
    df_merged['rate_on_gca_sqft'] = None
    df_merged[['is_duplicate', 'Primary Sale_or_Secondary Sale']] = None

    df_merged['sro_code'] = df_merged['sro_name'].map(SRO_DICT.get(city.lower(), {}))

    result_df = df_merged[[
        'index', 'project_name', 'village', 'location', 'igr_village',
        'year', 'quarter', 'city', 'sro_code', 'sro_name', 'document_no',
        'transaction_type', 'agreement_price', 'bajarbhav', 'property_description',
        'transaction_date', 'floor_no', 'unit_no', 'property_type_raw',
        'net_carpet_area_sqmt', 'balcony_sqmt', 'terrace_sqmt',
        'seller_name', 'purchaser_name', 'property_category',
        'internaldocumentnumber', 'micrno', 'bank_type', 'party_code',
        'dateofexecution', 'stampdutypaid', 'registrationfees',
        'modified_project_name', 'rera_location', 'project_lat', 'project_lng',
        'property_type', 'BHK', 'buyer_pincode',
        'Locality of buyer', 'District', 'StateName',
        'Tower', 'gross_carpet_sqft', 'rate_on_gca_sqft',
        'is_duplicate', 'Primary Sale_or_Secondary Sale',
        'llm_processed', 'manual_processed',
    ]]

    result_df = standardise_columns(result_df)

    # Village-level lat/lng
    village_dir = pd.read_excel(village_dir_path)
    location_lat = village_dir.set_index('village')['Project_location_lat'].to_dict()
    location_lng = village_dir.set_index('village')['Project_location_lng'].to_dict()
    result_df['location_lat'] = result_df['village'].map(location_lat)
    result_df['location_lng'] = result_df['village'].map(location_lng)

    result_df.loc[result_df['year'].isna(), 'quarter'] = None

    result_df = result_df[[
        'index', 'project_name', 'village', 'location', 'igr_village',
        'year', 'quarter', 'city', 'sro_name', 'sro_code', 'document_no',
        'transaction_type', 'agreement_price', 'bajarbhav', 'property_description',
        'transaction_date', 'floor_no', 'unit_no', 'property_type_raw',
        'net_carpet_area_sqmt', 'balcony_sqmt', 'terrace_sqmt',
        'seller_name', 'purchaser_name', 'property_category',
        'internaldocumentnumber', 'micrno', 'bank_type', 'party_code',
        'dateofexecution', 'stampdutypaid', 'registrationfees',
        'project_lat', 'project_lng', 'location_lat', 'location_lng',
        'property_type', 'bhk', 'buyer_pincode',
        'locality_of_buyer', 'district', 'statename',
        'tower', 'gross_carpet_sqft', 'rate_on_gca_sqft',
        'is_duplicate', 'primary_sale_or_secondary_sale',
        'llm_processed', 'manual_processed',
    ]]

    result_df = result_df.rename(columns={
        'village': 'village_mr',
        'bajarbhav': 'market_value',
        'bhk': 'bhk_br',
    })

    return result_df


# =============================================================================
# MAIN
# =============================================================================

def run_stage3(config: Stage3Config = DEFAULT_CONFIG) -> pd.DataFrame:
    city = config.city

    df = pd.read_excel(config.input_path)
    df = derive_net_carpet_area(df, city)

    # 3.1 Rename & reshape
    df = rename_columns(df)
    df = process_unit_and_floor(df)
    df = standardise_columns(df)
    df['net_carpet_area_sqmt'] = df['net_carpet_area_sqmt'].apply(convert_area_to_sqmt)
    df = select_required_columns(df)

    # 3.2 Split by category
    sale_df, lease_df, other_df = split_by_category(df)
    print("columns in sale_df : ", sale_df.columns)

    # 3.3 Clean sale data
    final_village = clean_sale_data(sale_df)

    # 3.4 RERA index
    rera_grand = pd.read_excel(config.rera_grand_path)
    rera_grand = rera_grand[[
        'index', 'modified_project_name', 'rera_location_v1',
        'rera_location', 'project_lat', 'project_lng',
        'bhk_wise_ca', 'carpet_wise_total_sold_units'
    ]]
    rera_grand = rera_grand[rera_grand['modified_project_name'] != 0]

    merged_df = assign_rera_index(final_village, rera_grand, city)

    # 3.5 Geocode
    village_df = geocode_coordinates(merged_df, config.coordinates_path, city)

    # 3.6 Property type
    village_df = classify_property_type(village_df)

    # 3.7 BHK
    rera_grand['bhk_wise_ca'] = rera_grand['bhk_wise_ca'].apply(safe_parse)
    rera_grand['building_wise_carpet_area'] = rera_grand['bhk_wise_ca'].apply(combine_columns)

    rera_keywords = pd.read_excel(config.rera_keywords_path)
    rera_keywords['Keywords'] = rera_keywords['Keywords'].str.upper().str.strip()
    rera_keywords['Final BHK'] = rera_keywords['Final BHK'].str.upper().str.strip()
    rera_keywords['Final Property Type'] = rera_keywords['Final Property Type'].str.upper().str.strip()

    village_df = assign_bhk_carpet_match(village_df, rera_grand, rera_keywords, config.bhk_max_diff)
    village_df = assign_bhk_range_fallback(village_df, rera_grand, rera_keywords)
    village_df = finalise_bhk(village_df)

    # 3.8 Buyer location
    village_df = add_buyer_location(village_df, config.postal_csv_path)

    # 3.9 Final assembly
    result_df = final_assembly(village_df, lease_df, other_df, city, config.village_dir_path)

    print(f"\nStage 3 complete. Shape: {result_df.shape}")
    result_df.to_excel(config.output_path, index=False)
    print(f"Saved -> {config.output_path}")
    return result_df


if __name__ == "__main__":
    result_df = run_stage3()