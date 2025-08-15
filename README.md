# Work Order Reporting App (Streamlit)

A lightweight Streamlit app to generate per–craft group reports from the **Time on Work Order** Excel export.

## Features
- Upload the **Time on Work Order** Excel file (with the same layout as your export).
- Date selector shows only dates present in the uploaded file (formatted **mm/dd/yyyy**).
- Personnel are assigned to craft groups from your **Address Book** file.
- Craft groups are displayed in the order defined by **Craft Group Order**.
- Per-person summary (Hours & unique Work Orders) and detailed rows.
- Export filtered detail to CSV.

## Quick Start
```bash
# 1) Create and activate a virtual environment (optional but recommended)
python -m venv .venv && source .venv/bin/activate   # Windows: .venv\\Scripts\\activate

# 2) Install dependencies
pip install -r requirements.txt

# 3) Run the app
streamlit run app.py
```

## Usage
1. Upload your **Time on Work Order** export (`.xlsx`).  
2. (Optional) Upload **Craft Group Order** and **Address Book** files.  
   - If omitted, the app uses the sample files located in `sample_data/`.
3. Select a date (only dates present in the file will be shown).  
4. View the report per craft group in the defined order.

## Expected File Formats
- **Time on Work Order**: Must include at least these columns once parsed:
  - `AddressBookNumber`, `Name`, `Production Date`, `OrderNumber`, `Sum of Hours.` (renamed to `Hours`), `Description`
  - The app auto-detects the header row even if there is an “Applied filters” banner above the table.
- **Craft Group Order**: Single column named `Craft Description` with the desired display order.
- **Address Book**: Columns `AddressBookNumber`, `Name`, `Craft Description`.

> Columns are case-sensitive as stated above. The app tries to be forgiving with minor variations and will warn if fields are missing.

## Deploying to Streamlit Community Cloud
1. Push this folder to a GitHub repository.
2. In Streamlit Community Cloud, set **Main file** to `app.py`.
3. Add any secrets if needed (none required for this app).

## Notes
- People not found in the Address Book are listed under **Unassigned** and shown in an expandable warning section.
- All dates are displayed for selection in **mm/dd/yyyy** format, while internally parsed with pandas.

## License
MIT
