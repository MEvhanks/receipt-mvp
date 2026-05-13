import streamlit as st
from PIL import Image, ImageOps
import pillow_heif
pillow_heif.register_heif_opener()
import pytesseract
import re
import pandas as pd
import hashlib
import numpy as np
import cv2
import fitz
import io
from pathlib import Path
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

EXCEL_PATH = "master.xlsx"
HASHES_PATH = "receipt_hashes.txt"
COLUMNS = ["date", "time", "merchant", "subtotal", "gst", "total"]

st.title("Taming the Receipt Monster")
st.caption("Upload receipt images or PDFs to extract GST and expense data.")


def pdf_to_image(uploaded_file):
    uploaded_file.seek(0)
    pdf_bytes = uploaded_file.read()
    uploaded_file.seek(0)
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    page = doc[0]
    mat = fitz.Matrix(2, 2)
    pix = page.get_pixmap(matrix=mat)
    img_bytes = pix.tobytes("png")
    return Image.open(io.BytesIO(img_bytes))


def preprocess_image(image):
    image = ImageOps.exif_transpose(image).convert("RGB")
    img_array = np.array(image)
    if len(img_array.shape) == 2:
        gray = img_array
    else:
        gray = cv2.cvtColor(img_array, cv2.COLOR_RGB2GRAY)
    _, thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    return Image.fromarray(thresh)


def normalize_ocr(text):
    text = re.sub(r'(\d)\s+\.(\d)', r'\1.\2', text)
    text = re.sub(r'(\d)\.\s+(\d)', r'\1.\2', text)
    # fix OCR confusion: l/I/| misread as 1 in numeric contexts
    text = re.sub(r'(?<!\w)[lI|](\d)', r'1\1', text)       # l8.71 → 18.71 (start)
    text = re.sub(r'(\d)[lI](?=[\d.,])', r'\g<1>1', text)  # 18l71 → 18171 (middle)
    text = re.sub(r'(\d)[lI](?!\w)', r'\g<1>1', text)       # 18.7l → 18.71 (end)
    return text


def _gst_rate_ok(total, gst, lo=0.04, hi=0.15):
    subtotal = total - gst
    if subtotal <= 0:
        return False
    return lo <= gst / subtotal <= hi


def extract_fields(raw_text):
    text = normalize_ocr(raw_text)
    lines = [l.strip() for l in text.strip().split('\n') if l.strip()]

    result = {
        'merchant': '',
        'date': '',
        'time': '',
        'subtotal': '',
        'gst': '',
        'total': '',
    }

    for line in lines:
        if len(line) > 3 and not re.match(r'^[\d\s\W]+$', line):
            result['merchant'] = line
            break

    date_match = re.search(r'\b(\d{1,2}[/\-]\d{1,2}[/\-]\d{4})\b', text)
    if not date_match:
        date_match = re.search(r'\b(\d{4}[/\-]\d{2}[/\-]\d{2})\b', text)
    if not date_match:
        date_match = re.search(r'\b(\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\w*\s+\d{4})\b', text, re.IGNORECASE)
    result['date'] = date_match.group(1) if date_match else ''

    time_match = re.search(r'\b(\d{1,2}:\d{2}(?::\d{2})?(?:\s*[AP]M)?)\b', text, re.IGNORECASE)
    result['time'] = time_match.group(1) if time_match else ''

    # Extract subtotal directly from receipt (before service charge and GST)
    subtotal_match = None
    for pattern in [
        # Full word forms
        r'Sub[\s\-]?[Tt]otal\s*:?\s*\$?(\d+[\.,]\d+)',
        r'Food\s+(?:&\s+Bev(?:erage)?s?)?\s+Sub[\s\-]?[Tt]otal\s*:?\s*\$?(\d+[\.,]\d+)',
        # Abbreviated forms: SUBTTL, SUBTL, SUBTOT, SUB TTL, SUB TOT, SUBT
        r'SUB[\s\-]?(?:TTL|TL|TOT(?:AL)?|T)\b[^\n]{0,20}?(\d+[\.,]\d+)',
        # Fuzzy OCR variants: sub-tota], sub-tota1, sub_tota etc. (match first 4 chars of "total")
        r'sub[\s\-_\.]*tota[^\n]{0,15}?(\d+[\.,]\d+)',
        # Before/excl GST
        r'(?:Before\s+GST|Excl\.?\s+GST|Excl\.\s+Tax)\s*:?\s*\$?(\d+[\.,]\d+)',
        # Net amount forms
        r'Nett?\s+(?:Amount|Total)\s*:?\s*\$?(\d+[\.,]\d+)',
    ]:
        subtotal_match = re.search(pattern, text, re.IGNORECASE)
        if subtotal_match:
            break
    if subtotal_match:
        result['subtotal'] = subtotal_match.group(1).replace(',', '.')

    # Extract GST so we can use it to validate the total candidate.
    gst_match = None
    for pattern in [
        r'Inclusive of.*?GST.*?\(\$?(\d+[\.,]\d+)\)',
        r'\d+%\s*GST\s+(\d+[\.,]\d+)',
        r'GST\s+\d+\S*\s+\$?(\d+[\.,]\d+)',
        r'GST.*?Amt.*?\n.*?\$?(\d+[\.,]\d+)',
        r'GST[^\n]{0,40}\$(\d+[\.,]\d+)',
        r'GST[^\n]{0,20}(\d+[\.,]\d+)',
        r'\$(\d+[\.,]\d+)\s*GST',
    ]:
        gst_match = re.search(pattern, text, re.IGNORECASE)
        if gst_match:
            break
    if gst_match:
        result['gst'] = gst_match.group(1).replace(',', '.')

    total_match = None
    for pattern in [
        r'Total\s+Due\s*[^$\d]*(\d+[\.,]\d+)',
        r'(?:VISA|MASTER\s*CARD|NETS|CASH|PAYNOW|GRABPAY|PAYLAH|AMEX|NETS\s*FLASHPAY)\s+\$?(\d+[\.,]\d+)',
        r'(?:Grand\s+)?TOTAL\s+\$?(\d+[\.,]\d+)',
        r'(?:Amount\s+Due|Amount\s+Payable|Paid\s+By|Net\s+Amount)\s*:?\s*\$?(\d+[\.,]\d+)',
        r'(?:Grand\s+)?Total\s*\n+\s*\$?(\d+[\.,]\d+)',
    ]:
        total_match = re.search(pattern, text, re.IGNORECASE)
        if total_match:
            break
    if total_match:
        result['total'] = total_match.group(1).replace(',', '.')
    else:
        amounts = re.findall(r'\$?(\d+[\.,]\d+)', text)
        amounts = [a for a in amounts if float(a.replace(',', '.')) > 1]
        if amounts:
            gst_val = None
            try:
                gst_val = float(result['gst'])
            except (ValueError, KeyError):
                pass

            if gst_val:
                valid = [a for a in amounts if _gst_rate_ok(float(a.replace(',', '.')), gst_val)]
                if valid:
                    result['total'] = min(valid, key=lambda x: float(x.replace(',', '.')))
                else:
                    candidates = [a for a in amounts if float(a.replace(',', '.')) > gst_val]
                    if candidates:
                        result['total'] = min(candidates, key=lambda x: float(x.replace(',', '.')))
            else:
                result['total'] = max(amounts, key=lambda x: float(x.replace(',', '.')))

    # If GST rate looks wrong, try prepending "1" (OCR sometimes drops a leading 1)
    if result['gst'] and result['total']:
        try:
            gst = float(result['gst'])
            total = float(result['total'])
            if not _gst_rate_ok(total, gst):
                candidate = float("1" + result['gst'])
                if _gst_rate_ok(total, candidate):
                    result['gst'] = f"1{result['gst']}"
        except ValueError:
            pass

    has_service_charge = bool(re.search(
        r'serv(?:ice)?\s*ch(?:g|arge)|svc\s*ch(?:g|arge)',
        text, re.IGNORECASE
    ))
    inclusive_gst = bool(re.search(
        r'(?:inclusive\s+of\s+(?:\d+%\s*)?GST|incl\.?\s+GST|GST\s+incl(?:uded)?)',
        text, re.IGNORECASE
    ))

    # If no GST extracted but total is inclusive of GST, back-calculate at 9%
    if not result['gst'] and result['total'] and inclusive_gst:
        try:
            total = float(result['total'])
            result['gst'] = f"{total * 9 / 109:.2f}"
            result['subtotal'] = f"{total * 100 / 109:.2f}"
        except ValueError:
            pass

    # If subtotal still missing and no service charge: compute total - gst
    if not result['subtotal'] and result['total'] and result['gst'] and not has_service_charge:
        try:
            result['subtotal'] = f"{float(result['total']) - float(result['gst']):.2f}"
        except ValueError:
            pass

    return result


def hash_file(uploaded_file):
    uploaded_file.seek(0)
    file_bytes = uploaded_file.read()
    uploaded_file.seek(0)
    return hashlib.md5(file_bytes).hexdigest()


def load_hashes():
    p = Path(HASHES_PATH)
    if p.exists():
        return set(p.read_text().splitlines())
    return set()


def save_hashes(hashes):
    Path(HASHES_PATH).write_text("\n".join(sorted(hashes)))


def load_master():
    if Path(EXCEL_PATH).exists():
        df = pd.read_excel(EXCEL_PATH, dtype=str)
        df = df.rename(columns={
            "gst_amount": "gst",
            "total_before_gst": "subtotal",
            "total_with_gst": "total",
        })
        for col in COLUMNS:
            if col not in df.columns:
                df[col] = ""
        return df
    return pd.DataFrame(columns=COLUMNS)


def _write_excel_table(df, writer, sheet_name="Receipts"):
    df.to_excel(writer, index=False, sheet_name=sheet_name)
    if not df.empty:
        ws = writer.sheets[sheet_name]
        ref = f"A1:{get_column_letter(len(df.columns))}{len(df) + 1}"
        tab = Table(displayName=sheet_name, ref=ref)
        tab.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium9",
            showFirstColumn=False, showLastColumn=False,
            showRowStripes=True, showColumnStripes=False,
        )
        ws.add_table(tab)
        for col in ws.columns:
            max_len = max(
                (len(str(cell.value)) for cell in col if cell.value is not None),
                default=8,
            )
            ws.column_dimensions[col[0].column_letter].width = max_len + 4


def save_master(df):
    with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl") as writer:
        _write_excel_table(df, writer)


uploaded_files = st.file_uploader(
    "Upload receipts (images or PDFs)",
    type=["jpg", "jpeg", "png", "pdf", "heic"],
    accept_multiple_files=True
)

if uploaded_files:
    master_df = load_master()
    existing_hashes = load_hashes()

    new_rows = []
    skipped = []

    for uploaded_file in uploaded_files:
        file_hash = hash_file(uploaded_file)

        if file_hash in existing_hashes:
            skipped.append(uploaded_file.name)
            continue

        if uploaded_file.name.lower().endswith(".pdf"):
            image = pdf_to_image(uploaded_file)
        else:
            image = Image.open(uploaded_file)

        processed_image = preprocess_image(image)

        with st.expander(f"Preview: {uploaded_file.name}"):
            col1, col2 = st.columns(2)
            with col1:
                st.write("Original")
                st.image(image, use_column_width=True)
            with col2:
                st.write("Preprocessed")
                st.image(processed_image, use_column_width=True)

        color_image = ImageOps.exif_transpose(image).convert("RGB")
        text_color = pytesseract.image_to_string(color_image)
        text_bw = pytesseract.image_to_string(processed_image)

        fields_color = extract_fields(text_color)
        fields_bw = extract_fields(text_bw)

        key_fields = ["date", "time", "gst", "total"]
        score_color = sum(1 for f in key_fields if fields_color.get(f))
        score_bw = sum(1 for f in key_fields if fields_bw.get(f))

        # merge: start from lower-scoring result, overlay the better one field by field
        base, overlay = (fields_bw, fields_color) if score_color >= score_bw else (fields_color, fields_bw)
        fields = base.copy()
        for f in key_fields + ["merchant", "subtotal"]:
            if not fields.get(f) and overlay.get(f):
                fields[f] = overlay[f]

        best_text = text_color if score_color >= score_bw else text_bw
        with st.expander(f"Raw OCR text: {uploaded_file.name}"):
            st.code(best_text, language=None)

        fields["_hash"] = file_hash
        new_rows.append(fields)

    if skipped:
        st.warning(f"Skipped {len(skipped)} duplicate(s): {', '.join(skipped)}")

    if new_rows:
        new_df = pd.DataFrame(new_rows, columns=COLUMNS)

        st.subheader("Review and edit before saving")
        st.caption("Click any cell to fix a value before saving.")

        edited_df = st.data_editor(new_df, use_container_width=True, hide_index=True)

        if st.button("Save to master.xlsx", type="primary"):
            combined_df = pd.concat([master_df, edited_df], ignore_index=True)
            combined_df["_d"] = pd.to_datetime(combined_df["date"], dayfirst=True, errors="coerce")
            combined_df = combined_df.sort_values("_d", ascending=True, na_position="last").drop(columns=["_d"]).reset_index(drop=True)
            save_master(combined_df)

            new_hashes = existing_hashes | {r["_hash"] for r in new_rows}
            save_hashes(new_hashes)

            st.success(f"Saved {len(edited_df)} receipt(s) to master.xlsx")

            with open(EXCEL_PATH, "rb") as f:
                st.download_button(
                    label="Download master.xlsx",
                    data=f,
                    file_name="master.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    else:
        if not skipped:
            st.info("No new receipts to process.")

st.divider()

with st.expander("Danger Zone"):
    st.warning("This will permanently delete all saved receipts and duplicate tracking.")
    if st.button("Clear all data", type="primary"):
        for p in [EXCEL_PATH, HASHES_PATH]:
            Path(p).unlink(missing_ok=True)
        st.success("All data cleared.")
        st.rerun()

st.subheader("All Receipts")

master_df = load_master()

master_df["_date_parsed"] = pd.to_datetime(master_df["date"], dayfirst=True, errors="coerce")
valid_dates = master_df["_date_parsed"].dropna()

col_start, col_end, col_sort = st.columns([1, 1, 2])
with col_start:
    start_date = st.date_input("From", value=valid_dates.min().date() if not valid_dates.empty else None)
with col_end:
    end_date = st.date_input("To", value=valid_dates.max().date() if not valid_dates.empty else None)
with col_sort:
    sort_by = st.selectbox("Sort by", [
        "Date (latest first)", "Date (oldest first)",
        "Name (A-Z)", "Name (Z-A)",
        "Price (low → high)", "Price (high → low)",
        "GST (low → high)", "GST (high → low)",
    ])

if master_df.empty:
    st.info("No receipts saved yet.")
else:
    filtered = master_df.copy()
    if start_date and end_date:
        filtered = filtered[
            (filtered["_date_parsed"] >= pd.Timestamp(start_date)) &
            (filtered["_date_parsed"] <= pd.Timestamp(end_date))
        ]

    filtered["_total_num"] = pd.to_numeric(filtered["total"], errors="coerce")
    filtered["_gst_num"] = pd.to_numeric(filtered["gst"], errors="coerce")

    sort_map = {
        "Date (latest first)": ("_date_parsed", False),
        "Date (oldest first)": ("_date_parsed", True),
        "Name (A-Z)": ("merchant", True),
        "Name (Z-A)": ("merchant", False),
        "Price (low → high)": ("_total_num", True),
        "Price (high → low)": ("_total_num", False),
        "GST (low → high)": ("_gst_num", True),
        "GST (high → low)": ("_gst_num", False),
    }
    sort_col, sort_asc = sort_map[sort_by]
    filtered = filtered.sort_values(sort_col, ascending=sort_asc, na_position="last")

    st.dataframe(filtered[COLUMNS], use_container_width=True, hide_index=True)

    export_df = filtered[COLUMNS].reset_index(drop=True)
    export_buf = io.BytesIO()
    with pd.ExcelWriter(export_buf, engine="openpyxl") as writer:
        _write_excel_table(export_df, writer)
    export_buf.seek(0)
    st.download_button(
        label=f"Download filtered ({len(filtered)} rows)",
        data=export_buf,
        file_name="receipts_filtered.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
