# Mohon B / Mohon S DBF → XLSX Converter

A Streamlit application that converts Mohon B or Mohon S DBF files to XLSX.

## Features

- Choose **Mohon B** or **Mohon S**
- Upload a `.DBF` file
- Reads Visual FoxPro/DBF date and datetime fields
- Converts dates to real Excel date values
- Displays dates as `DD/MM/YYYY`
- Keeps numeric fields numeric
- Freezes the header and enables Excel filters
- Downloads the converted XLSX

## Run locally

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## Deploy to Streamlit Community Cloud

1. Create a GitHub repository.
2. Upload:
   - `streamlit_app.py`
   - `requirements.txt`
   - `README.md`
   - `.gitignore`
3. Do NOT upload actual DBF data containing sensitive information.
4. Open Streamlit Community Cloud.
5. Select the repository, branch `main`, and file `streamlit_app.py`.
6. Deploy.

## Typical files

- Mohon B: `mhn_bpkb.DBF`
- Mohon S: `mhn_stnk.dbf`

The app does not require these files to be stored in the repository; users upload them through the webpage.
