import io
from datetime import datetime
from pathlib import Path
import pandas as pd
import streamlit as st
from dbfread import DBF
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

st.set_page_config(
    page_title="Mohon B/S → XLSX Converter",
    page_icon="📄",
    layout="wide",
)

st.title("📄 Mohon B / Mohon S → XLSX Converter")
st.write("Choose the Mohon type, upload the DBF file, then convert it to Excel.")

# --- Step 1: choose file type ---
mohon_type = st.radio(
    "1. Choose Mohon type",
    ["Mohon B", "Mohon S"],
    horizontal=True,
)

# --- Step 2: upload ---
st.subheader("2. Upload DBF file")
uploaded_file = st.file_uploader(
    f"Upload the DBF file for {mohon_type}",
    type=["dbf", "DBF"],
)

if uploaded_file is not None:
    # Basic filename warning
    expected_hint = "mhn_bpkb" if mohon_type == "Mohon B" else "mhn_stnk"
    st.caption(f"Selected: `{uploaded_file.name}`")

    if expected_hint.lower() not in uploaded_file.name.lower():
        st.warning(
            f"The filename does not look like the usual {mohon_type} file "
            f"(`{expected_hint}.dbf`). You can still continue if this is intentional."
        )

    # DBF -> DataFrame
    try:
        # dbfread expects a filesystem path, not a BytesIO object.
        # Save the Streamlit upload temporarily, then give DBF() the path.
        import tempfile
        import os

        suffix = Path(uploaded_file.name).suffix or ".dbf"
        temp_path = None

        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            tmp.write(uploaded_file.getvalue())
            temp_path = tmp.name

        # load=True reads the complete DBF table into memory.
        # char_decode_errors='ignore' prevents one bad character from stopping
        # the conversion.
        table = DBF(
            temp_path,
            load=True,
            char_decode_errors="ignore",
            ignore_missing_memofile=True,
        )

        records = list(table.records)
        df = pd.DataFrame(records)

        # The DBF has now been read, so the temporary uploaded file can be removed.
        if temp_path and os.path.exists(temp_path):
            os.remove(temp_path)
            temp_path = None

        # DBF deletion flag: dbfread normally excludes deleted records.
        # Convert date/datetime columns into real Excel-compatible datetimes.
        date_columns = []
        datetime_columns = []

        for col in df.columns:
            if df[col].dtype == "object" and len(df) > 0:
                # Detect Python date/datetime objects
                sample = df[col].dropna()
                if not sample.empty:
                    first = sample.iloc[0]
                    if isinstance(first, datetime):
                        datetime_columns.append(col)
                    elif hasattr(first, "year") and hasattr(first, "month") and hasattr(first, "day"):
                        date_columns.append(col)

        # Also catch columns by their DBF field type.
        for field in table.fields:
            if field.name in df.columns:
                if field.type == "D":
                    if field.name not in date_columns:
                        date_columns.append(field.name)
                elif field.type == "T":
                    if field.name not in datetime_columns:
                        datetime_columns.append(field.name)

        # For Excel, use actual datetime/date values and format them as DD/MM/YYYY.
        for col in date_columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

        for col in datetime_columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

        # Some DBF files contain control characters (e.g. stray 0x00-0x1F bytes)
        # that Excel's XML format cannot store. Strip them from every text
        # column so openpyxl doesn't raise IllegalCharacterError.
        text_columns = [
            col for col in df.columns
            if col not in date_columns and col not in datetime_columns
        ]
        for col in text_columns:
            if df[col].dtype == "object":
                df[col] = df[col].apply(
                    lambda v: ILLEGAL_CHARACTERS_RE.sub("", v) if isinstance(v, str) else v
                )

        st.success(
            f"DBF loaded successfully — **{len(df):,} records** and "
            f"**{len(df.columns)} columns**."
        )

        if date_columns:
            st.info("Date fields detected: " + ", ".join(date_columns))
        if datetime_columns:
            st.info(
                "DateTime fields detected: "
                + ", ".join(datetime_columns)
                + ". They will be displayed as DD/MM/YYYY."
            )

        st.subheader("3. Preview")
        st.dataframe(df.head(100), use_container_width=True, height=350)

        # --- Step 3: convert ---
        st.subheader("4. Convert")
        output_name = "mohon_B.xlsx" if mohon_type == "Mohon B" else "mohon_S.xlsx"

        if st.button("🔄 Convert to XLSX", type="primary"):
            output = io.BytesIO()

            with pd.ExcelWriter(
                output,
                engine="openpyxl",
                date_format="DD/MM/YYYY",
                datetime_format="DD/MM/YYYY",
            ) as writer:
                df.to_excel(writer, index=False, sheet_name=mohon_type.replace(" ", "_"))

                ws = writer.book[writer.sheets[mohon_type.replace(" ", "_")].title]

                # Explicitly apply DD/MM/YYYY to every detected date/date-time
                # cell so Excel treats them as real dates rather than text.
                for col_name in date_columns + datetime_columns:
                    col_idx = df.columns.get_loc(col_name) + 1
                    for row_idx in range(2, len(df) + 2):
                        ws.cell(row=row_idx, column=col_idx).number_format = "DD/MM/YYYY"

                # Freeze header and enable Excel filter.
                ws.freeze_panes = "A2"
                ws.auto_filter.ref = ws.dimensions

                # Reasonable column widths without making huge sheets too wide.
                for column_cells in ws.columns:
                    max_length = 0
                    column_letter = column_cells[0].column_letter
                    for cell in column_cells[:200]:
                        value = "" if cell.value is None else str(cell.value)
                        max_length = max(max_length, len(value))
                    ws.column_dimensions[column_letter].width = min(max(max_length + 2, 10), 40)

            output.seek(0)

            st.success("Conversion completed!")
            st.download_button(
                label=f"⬇️ Download {output_name}",
                data=output,
                file_name=output_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    except Exception as e:
        # Clean up the temporary DBF if an error occurs.
        try:
            if temp_path and os.path.exists(temp_path):
                os.remove(temp_path)
        except Exception:
            pass

        st.error("Could not read or convert this DBF file.")
        st.exception(e)

st.divider()
st.caption(
    "The uploaded DBF is processed by this Streamlit session. "
    "Do not upload sensitive files to public repositories such as GitHub."
)
