# maltego-to-excel

Converts a Maltego CSV graph export into a structured Excel file.

Automatically detects root entities (emails, phone numbers, persons, etc.) and produces one row per root, with all linked platform accounts and leaf nodes spread as columns.

## What it does

- Parses any Maltego CSV export
- Detects root entities purely from graph structure — no entity types are hardcoded
- Excludes platform/affiliation nodes from being misidentified as roots
- Outputs a colour-coded `.xlsx` with one row per root and all relationships flattened horizontally

## Requirements

- Python 3.8+
- pip

## Setup

```bash
git clone https://github.com/nitishchandan/maltego-to-excel.git
cd maltego-to-excel
pip install -r requirements.txt
```

## Usage

### Web app (recommended)

```bash
python app.py
```

Open **http://localhost:5000** in your browser. Drop in a CSV, preview the detected roots, download the Excel.

### Command line

```bash
python maltego_to_excel.py export.csv
python maltego_to_excel.py export.csv --out results.xlsx
```

## Exporting from Maltego

In Maltego desktop: **File → Export → Export Graph as CSV**

## Notes

- All processing happens locally — no data is sent anywhere
- Works with any Maltego CSV export regardless of entity types or graph size
- New platform types are picked up automatically and colour-coded in the output
