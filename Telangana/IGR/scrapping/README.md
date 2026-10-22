# Telangana IGR Document Scraper

`Telangana_scrapper.py` automates lookups of registered documents on the
[Telangana Registration & Stamps portal](https://registration.telangana.gov.in/).
It signs in, searches document numbers for registration year **2025**, and
appends any returned records to an Excel workbook.

## Requirements

- Python 3
- Google Chrome
- Tesseract OCR installed on the machine
- The Python packages used by the script:

  ```powershell
  python -m pip install pandas openpyxl selenium pillow pytesseract requests
  ```

Selenium uses Chrome to control the browser. Make sure your Selenium and Chrome
installations can work together.

## Configure before running

The script currently has machine-specific settings near the top of
`Telangana_scrapper.py`:

- `pytesseract.pytesseract.tesseract_cmd` points to the Tesseract executable.
  Update it if Tesseract is installed elsewhere.
- `excel_file` is the full path of the output workbook. Update it to a location
  available on your machine.
- The portal login values are currently embedded in the script. Replace them
  with your own authorized credentials, and do not share or commit credentials.
- The document year passed to `scrape_document` is currently `2025`.
- The district and Sub-Registrar Office are selected by fixed option positions
  in the portal. Confirm those selections still correspond to the intended
  district and office before collecting data.

The script accesses the portal using a browser session and attempts to read its
CAPTCHA using Tesseract OCR. Portal changes or CAPTCHA changes may prevent
login or scraping.

## Choose documents to scrape

Create a `last_doc.txt` file in the **current working directory** (the directory
from which you launch the script). Its contents select one of two modes:

### Sequential range mode

Put a single integer in the file. The script starts at that document number and
queues numbers up to (but not including) `100000`.

```text
4911
```

If the file is absent, the script starts at document number `1`. In range mode,
the script updates `last_doc.txt` after saving a document with data. Documents
without data are not recorded as completed.

### Specific document list mode

Put comma-separated document numbers in the file. Only those numbers are
queued, in the order given.

```text
4911, 4912, 4913
```

In list mode, `last_doc.txt` is treated as input and is not updated with
progress.

## Run

From the directory containing the script, run:

```powershell
python .\Telangana_scrapper.py
```

The script starts two Chrome workers. Keep the console open while it runs; it
prints login, scraping, and error messages there.

## Output and progress

- Scraped rows are written to the workbook configured by `excel_file`.
- If the workbook already exists, rows are appended to its active sheet;
  otherwise, a new workbook is created.
- Each result includes the portal table fields plus `Document No`, `District`,
  and `Sub-Registrar Office`.
- `last_doc.txt` stores the most recently saved document number in range mode.
- CAPTCHA images named `captcha_Worker-1.png` and
  `captcha_Worker-2.png` may be created in the current working directory.

Use the portal only in accordance with its terms and applicable laws, and avoid
overloading it with requests.
