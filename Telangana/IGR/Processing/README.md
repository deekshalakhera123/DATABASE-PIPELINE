## Overview
This Python script extracts and processes property registration data from Excel files. It performs comprehensive data extraction including property descriptions, party information, dates, document details, and area conversions.

## Features

### 1. **Property Description Extraction**
Extracts the following fields from the "Description of property" column:
- `VILL/COL` - Village/Colony name
- `W-B` - Ward/Block information
- `SURVEY` - Survey number
- `PLOT` - Plot number
- `HOUSE` - House number
- `APARTMENT` - Apartment name
- `BLOCK` - Block number
- `FLAT` - Flat number
- `EXTENT` - Land extent (area)
- `BUILT` - Built-up area
- `Boundires` - Property boundaries

### 2. **Area Conversion**
- **EXTENT in SqFt**: Converts extent from Square Yards to Square Feet (1 Sq.Yd = 9 Sq.Ft)
- **BUILT in SqFt**: Formats built-up area from Square Feet format (no conversion needed)

### 3. **Date Extraction**
Extracts three types of dates from the "Reg.Date Exe.Date Pres.Date" column:
- Registration Date `(R)`
- Execution Date `(E)`  
- Presentation Date `(P)`

### 4. **Document Information**
Extracts from "Nature & Mkt.Value Con. Value" column:
- Document type code (first 4 digits)
- Document Type (full description)
- Market Value
- Consideration Value

### 5. **Party Information Extraction**
Extracts from "Name of Parties Executant(EX) & Claimants(CL)" column:
- **Seller**: Parties with types {EX, MR, DR, RR, PL, LR, FP}
- **Buyer**: Parties with types {CL, ME, DE, RE, AY, LE, SP}
- Supports multiple sellers and buyers
- Handles numbered entries (1., 2., 3., etc.)
- Removes duplicate entries

### 6. **Transaction Type Classification**
Adds a "Transaction Type" column based on Document Type:
- **Sales**: Matches predefined sales document types
- **Lease**: Matches predefined lease document types  
- **Others**: All other document types

## Input File Requirements

The input Excel file should contain the following columns:
- `Description of property` (required)
- `Reg.Date Exe.Date Pres.Date` (optional)
- `Nature & Mkt.Value Con. Value` (optional)
- `Name of Parties Executant(EX) & Claimants(CL)` (optional)

## Output

The script generates a new Excel file with:
- All original columns
- New extracted columns for property fields
- Converted area columns (EXTENT in SqFt, BUILT in SqFt)
- Extracted date columns
- Document information columns
- Seller and Buyer columns
- Transaction Type classification

## Party Type Mappings

### Seller Types:
- `EX` - Executant
- `MR` - Minor
- `DR` - Donor
- `RR` - Releasor
- `PL` - Plaintiff
- `LR` - Land Receiver
- `FP` - Family Pensioner

### Buyer Types:
- `CL` - Claimant
- `ME` - Major
- `DE` - Donee
- `RE` - Releasee
- `AY` - Aayayatdar
- `LE` - Lessee
- `SP` - Service Pensioner

## Transaction Type Categories

### Sales Types (18 types):
- Sale Deed
- AGREEMENT OF SALE CUM GPA
- Sale Agreement Without Possess
- Sale Agreement With Possession
- CONVEYANCE FOR CONSIDERATION
- CONVEYANCE DEED(WITHOUT CONSID
- Sale deed executed by A.P.Hous
- Sale Deeds executed by Courts
- Sale Certificate
- Exchange
- Assignment deed
- RECONVEYANCE DEED EXECUTED BY
- Sale of life interest
- Development Agreement Cum GPA
- DEVELOPMENT AGREEMENT OR CONST
- Sale deed executed by or infav
- Sale deed executed by Society
- Sale deed in favour of State o

### Lease Types (5 types):
- Lease Deed
- Lease in favour of State/Centr
- Lease(others)
- Surrender of Lease
- Transfer of Lease

## Usage

1. **Set file paths** at the top of the script:
```python
INPUT_PATH = r"path\to\your\input_file.xlsx"
OUTPUT_PATH = r"path\to\your\output_file.xlsx"
```

2. **Run the script**:
```bash
python script_name.py
```

3. **Check output**:
- The script will display progress messages
- Shows transaction type summary with counts
- Provides debug samples for verification
- Saves the extracted data to the output path

## Debug Features

The script includes several debug outputs:
- Empty boundaries detection
- BUILT conversion samples (first 10 rows)
- EXTENT conversion samples (first 10 rows)
- Party extraction samples with new mappings
- Transaction Type samples (first 20 rows)

## Dependencies

- pandas
- openpyxl (for Excel file handling)
- re (regular expressions - built-in)

## Notes

- All extracted text is cleaned of extra spaces
- Duplicate party names are removed
- The script handles missing columns gracefully
- Creates output directory if it doesn't exist
- Preserves all original data while adding new columns
