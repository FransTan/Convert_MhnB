# DBF → XLSX Streamlit Converter

A Streamlit app that converts legacy DBF / Visual FoxPro DBF files into XLSX.

## Features
- Upload `.DBF` directly in the browser.
- Reads DBF fields without requiring Microsoft Access or Excel.
- Converts `D` fields to real Excel dates with `DD/MM/YYYY` display format.
- Converts Visual FoxPro `T` DateTime fields; by default only the date is kept.
- Preserves character, numeric and logical fields.
- Skips deleted DBF records.
- Produces an `.xlsx` file named after the original DBF.

## Run locally on Windows

Open Command Prompt in this folder:

```text
py -m pip install -r requirements.txt
py -m streamlit run streamlit_app.py
```

Then open the local URL shown by Streamlit.

## Deploy to Streamlit Community Cloud

1. Create a GitHub repository, for example `dbf-to-xlsx`.
2. Upload these files to the repository:
   - `streamlit_app.py`
   - `requirements.txt`
3. Go to Streamlit Community Cloud and connect your GitHub account.
4. Create an app and select the repository, branch, and `streamlit_app.py`.
5. Deploy.

The DBF data is processed in the running app and is not stored in the GitHub repository by this project.

## Important for sensitive data
If the DBF contains confidential customer/company information, treat the deployed app as a sensitive-data processing service. Do not put sample DBF files containing real data into a public GitHub repository. If the data is highly confidential, consider deploying the same app in an internal/private environment instead of a public Community Cloud app.
