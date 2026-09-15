import io
import math
import struct
import tempfile
import re
from datetime import date, datetime, timedelta
from pathlib import Path

import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="DBF → XLSX Converter", page_icon="📄", layout="wide")


def decode_text(raw: bytes) -> str:
    # Most legacy Indonesian/FoxPro DBFs use Windows-1252/ANSI.
    for enc in ("cp1252", "latin-1", "utf-8"):
        try:
            text = raw.decode(enc).rstrip()
            # Remove XML-illegal control characters that Excel/openpyxl cannot store.
            return re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F]", "", text)
        except UnicodeDecodeError:
            pass
    text = raw.decode("latin-1", errors="replace").rstrip()
    return re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F]", "", text)


def foxpro_datetime(raw: bytes):
    """Decode an 8-byte Visual FoxPro DateTime value.
    First 4 bytes = Julian day; next 4 bytes = milliseconds since midnight.
    """
    if len(raw) != 8 or raw == b"\x00" * 8:
        return None
    julian_day = struct.unpack("<I", raw[:4])[0]
    milliseconds = struct.unpack("<I", raw[4:])[0]
    if julian_day == 0:
        return None

    # Julian day 0 in FoxPro is 4713-01-01 BCE; this equivalent offset
    # converts the DBF Julian day to a Python date.
    try:
        d = date(1, 1, 1) + timedelta(days=julian_day - 1721426)
        seconds, millis = divmod(milliseconds, 1000)
        hours, rem = divmod(seconds, 3600)
        minutes, seconds = divmod(rem, 60)
        if hours >= 24:
            return d
        return datetime(d.year, d.month, d.day, hours, minutes, seconds, millis * 1000)
    except (OverflowError, ValueError):
        return None


def parse_dbf(file_obj, keep_time=True):
    """Yield (field_names, rows, metadata) from a DBF file without pandas/dbfread."""
    file_obj.seek(0)
    header = file_obj.read(32)
    if len(header) != 32:
        raise ValueError("File is too small to be a valid DBF file.")

    record_count = struct.unpack("<I", header[4:8])[0]
    header_len = struct.unpack("<H", header[8:10])[0]
    record_len = struct.unpack("<H", header[10:12])[0]

    fields = []
    while file_obj.tell() < header_len:
        desc = file_obj.read(32)
        if not desc or desc[0] == 0x0D:
            break
        name = desc[:11].split(b"\x00", 1)[0].decode("ascii", errors="replace")
        field_type = chr(desc[11])
        length = desc[16]
        decimals = desc[17]
        fields.append((name, field_type, length, decimals))

    if not fields:
        raise ValueError("No DBF fields were found.")

    field_names = [x[0] for x in fields]
    date_columns = set()
    datetime_columns = set()
    for name, field_type, _, _ in fields:
        if field_type == "D":
            date_columns.add(name)
        elif field_type == "T":
            datetime_columns.add(name)

    file_obj.seek(header_len)

    def convert(raw, field_type, decimals):
        if field_type == "C":
            value = decode_text(raw)
            return value if value != "" else None
        if field_type == "D":
            text = raw.decode("ascii", errors="ignore").strip()
            if len(text) != 8 or text == "00000000":
                return None
            try:
                return date(int(text[:4]), int(text[4:6]), int(text[6:8]))
            except ValueError:
                return None
        if field_type == "T":
            value = foxpro_datetime(raw)
            if value is None:
                return None
            return value if keep_time else value.date()
        if field_type == "N":
            text = raw.decode("ascii", errors="ignore").strip()
            if not text:
                return None
            try:
                number = float(text) if decimals else int(float(text))
                return number
            except ValueError:
                return text
        if field_type == "L":
            text = raw.decode("ascii", errors="ignore").upper()
            if text in ("Y", "T"):
                return True
            if text in ("N", "F"):
                return False
            return None
        # Generic fallback for other DBF field types.
        value = decode_text(raw)
        return value if value != "" else None

    rows_read = 0
    for _ in range(record_count):
        record = file_obj.read(record_len)
        if len(record) < record_len:
            break
        if record[0:1] == b"*":  # deleted record
            continue

        row = []
        offset = 1
        for _, field_type, length, decimals in fields:
            raw = record[offset:offset + length]
            row.append(convert(raw, field_type, decimals))
            offset += length
        rows_read += 1
        yield row

    metadata = {
        "record_count": record_count,
        "rows_read": rows_read,
        "field_count": len(fields),
        "date_columns": date_columns,
        "datetime_columns": datetime_columns,
        "fields": fields,
    }
    return field_names, metadata


def convert_dbf_to_xlsx(uploaded_file, keep_time=True):
    # Use a temporary file so the DBF does not need to be duplicated in RAM.
    with tempfile.NamedTemporaryFile(suffix=".dbf", delete=False) as tmp:
        tmp.write(uploaded_file.getbuffer())
        dbf_path = tmp.name

    try:
        with open(dbf_path, "rb") as f:
            # Read field definitions once for column names and date indexes.
            header = f.read(32)
            header_len = struct.unpack("<H", header[8:10])[0]
            f.seek(32)
            fields = []
            while f.tell() < header_len:
                desc = f.read(32)
                if not desc or desc[0] == 0x0D:
                    break
                name = desc[:11].split(b"\x00", 1)[0].decode("ascii", errors="replace")
                fields.append((name, chr(desc[11]), desc[16], desc[17]))
            field_names = [x[0] for x in fields]
            date_indexes = {i for i, x in enumerate(fields) if x[1] == "D"}
            datetime_indexes = {i for i, x in enumerate(fields) if x[1] == "T"}
            date_columns = {x[0] for x in fields if x[1] == "D"}
            datetime_columns = {x[0] for x in fields if x[1] == "T"}

        # Normal openpyxl workbook is used so we can apply true Excel date formats.
        # This file has ~77k rows / 50 columns, which is practical for XLSX creation.
        wb = Workbook()
        ws = wb.active
        ws.title = "DATA"
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(field_names))}1"

        for i, name in enumerate(field_names, 1):
            c = ws.cell(1, i, name)
            c.font = Font(bold=True)
            c.alignment = Alignment(horizontal="center", vertical="center")

        with open(dbf_path, "rb") as f:
            gen = parse_dbf(f, keep_time=keep_time)
            count = 0
            try:
                while True:
                    row = next(gen)
                    count += 1
                    for i, value in enumerate(row):
                        c = ws.cell(count + 1, i + 1, value)
                        if i in date_indexes and value is not None:
                            c.number_format = "DD/MM/YYYY"
                        elif i in datetime_indexes and value is not None:
                            c.number_format = "DD/MM/YYYY HH:MM:SS" if keep_time else "DD/MM/YYYY"
            except StopIteration:
                pass

        # Use DBF field lengths as a safe initial width; cap to avoid huge columns.
        for i, (name, field_type, length, _) in enumerate(fields, 1):
            width = min(max(len(name) + 2, min(length, 30)), 35)
            ws.column_dimensions[get_column_letter(i)].width = width

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        return output.getvalue(), count, field_names, date_columns, datetime_columns
    finally:
        try:
            Path(dbf_path).unlink(missing_ok=True)
        except Exception:
            pass


st.title("📄 DBF → XLSX Converter")
st.write("Upload a DBF file and convert it to Excel while preserving the DBF fields and formatting date fields correctly.")

uploaded = st.file_uploader("Choose a DBF file", type=["dbf", "DBF"])
keep_time = st.checkbox("Keep time for FoxPro DateTime (T) fields", value=False,
                        help="If unchecked, T fields are exported as dates only (DD/MM/YYYY).")

if uploaded:
    size_mb = uploaded.size / (1024 * 1024)
    st.info(f"File: **{uploaded.name}** — {size_mb:.1f} MB")

    if st.button("Convert to XLSX", type="primary"):
        try:
            with st.spinner("Reading DBF and creating XLSX…"):
                xlsx_bytes, row_count, columns, date_cols, datetime_cols = convert_dbf_to_xlsx(
                    uploaded, keep_time=keep_time
                )
            output_name = Path(uploaded.name).stem + ".xlsx"
            st.success(f"Conversion completed: **{row_count:,} records** and **{len(columns)} columns**.")
            st.write("**Date fields:**", ", ".join(sorted(date_cols)) or "None")
            if datetime_cols:
                st.write("**FoxPro DateTime fields:**", ", ".join(sorted(datetime_cols)))
            st.download_button(
                "⬇️ Download XLSX",
                data=xlsx_bytes,
                file_name=output_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )
        except Exception as exc:
            st.error(f"Conversion failed: {exc}")
            st.exception(exc)

st.caption("Date cells are stored as real Excel dates, with display format DD/MM/YYYY. The converter also handles Visual FoxPro DateTime fields.")
