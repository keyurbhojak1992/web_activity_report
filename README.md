# Diamond Sales Log Processor

## Recommended deployment

- Platform: Streamlit Community Cloud
- Python: **3.12**
- Entrypoint: `app.py`
- Keep `app.py`, `report_builder.py`, and `requirements.txt` in the same GitHub folder.

## Why this version is safer

- Removes the background keep-alive thread.
- Calls `st.set_page_config()` before every other Streamlit command.
- Uses the uploaded Weblog workbook as the output workbook instead of copying every cell and style into a second workbook.
- Writes the final workbook to a temporary disk file before loading the download bytes, reducing peak RAM usage.
- Stores the finished report in session state so clicking Download does not regenerate it.
- Validates required Excel columns and displays useful errors.
- Blocks large Weblog files by default on low-memory cloud hosting.

## Required columns

### Weblog Data

- `PARTY_COMPANY_NAME`
- `EMPLOYEE_SHORT_NAME`
- `description`
- `ipAddress`

### Color Master Data

- `Company Name`
- `Color`
- `Zone` is optional

### Critical Search Data

- `Grp`
- `Name` or `Short Name`

### Sales and Bid Data (optional)

- `Sold Party`
- `Type`
- `AMT`

## Local run

```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

## Important Excel limitation

The app preserves normal worksheet data, formulas, cell styles, dimensions, merged cells, images, and charts supported by openpyxl. Some uncommon Excel extensions, external links, slicers, or unsupported vendor-specific features may be removed when the workbook is saved.

## Critical Search counting

Critical Search is counted by matched master shapes, not only by qualifying log rows.

The Streamlit control **Count Critical Search only in rows containing SEARCHID** selects the matching scope:

- **Enabled:** critical master shapes are checked only in descriptions containing `SEARCHID`.
- **Disabled:** critical master shapes are checked in every Weblog description.

In both modes:

- Every distinct master row under `Grp` = `F3` or `CRITICAL` contributes 1 when either its `Name` or `Short Name` is present.
- `Name` and `Short Name` are aliases of the same shape and cannot double-count that shape in one log row.
- Repeated occurrences of one shape in the same description count once.
- A row containing five distinct critical shapes contributes 5.
- Search-related rows (`SEARCHID` or `SEARCH PERFORMED`) with no critical shape contribute 1 to `Detail`.
- In full-data mode, a non-search row containing critical shapes is classified as `Critical_Search` before other action rules.

Because `Grand Total` is the sum of action columns, multiple critical shapes in one row also increase `Grand Total` by the same amount.
