# Lookup Automator

Simple Streamlit app that automates Excel-style VLOOKUP/XLOOKUP steps by letting you upload two spreadsheets, pick matching columns, and pull the values you need.

## Quick start

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

## Workflow

1. Upload your **base file** (the sheet that should receive the lookup values).
2. Upload the **lookup file** (contains the reference values you want to bring over).
3. Select the primary key column from each file.
4. Choose one or more columns from the lookup file to append to the base file.
5. Click **Run lookup** to preview the merged data and download it as CSV.
