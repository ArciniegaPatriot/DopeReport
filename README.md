
# Autofill Numbers App (Streamlit)

Upload a CSV/Excel report, map columns once (or load a JSON config), and auto-fill the KPIs:
- Total Calls
- Total Agents Staffed
- Abandonment Rate (total & by custom skills list)
- AHT by Group
- Exports: Markdown, Word (.docx), PDF
- Save/load config (JSON)
- Robust Excel parsing for Streamlit Cloud

## Local Run
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy on Streamlit Community Cloud
1. Push this folder to GitHub.
2. Go to https://share.streamlit.io → **New app**.
3. Select your repo and set **Main file path** to `app.py`.
4. Deploy and share the URL.

## Using a Shared Config
- Map columns & skills → click **Download current config (JSON)**.
- Teammates upload that JSON under **Sidebar → Config** to auto-fill mappings.

## Notes
- If your Excel file is `.xlsx`, we use **openpyxl**; for `.xls`, we use **xlrd**.
- If `python-docx` or `reportlab` fail to install, the app still runs (exports disabled).
