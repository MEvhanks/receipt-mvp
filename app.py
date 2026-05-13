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
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

COLUMNS = ["date", "time", "merchant", "subtotal", "gst", "total"]

if "master_df" not in st.session_state:
    st.session_state.master_df = pd.DataFrame(columns=COLUMNS)
if "seen_hashes" not in st.session_state:
    st.session_state.seen_hashes = set()

st.title("Taming the Receipt Monster")
st.caption("Upload receipt images or PDFs to extract GST and expense data.")


def _pdf_to_image(file_bytes: bytes) -> Image.Image:
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    pix = doc[0].get_pixmap(matrix=fitz.Matrix(2, 2))
    return Image.open(io.BytesIO(pix.tobytes("png")))


def _preprocess(image: Image.Image) -> Image.Image:
    img = np.array(ImageOps.exif_transpose(image).convert("RGB"))
    gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY) if img.ndim == 3 else img
    _, thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    return Image.fromarray(thresh)


def _normalize(text: str) -> str:
    text = re.sub(r'(\d)\s+\.(\d)', r'\1.\2', text)
    text = re.sub(r'(\d)\.\s+(\d)', r'\1.\2', text)
    text = re.sub(r'(?<!\w)[lI|](\d)', r'1\1', text)
    text = re.sub(r'(\d)[lI](?=[\d.,])', r'\g<1>1', text)
    text = re.sub(r'(\d)[lI](?!\w)', r'\g<1>1', text)
    return text


def _gst_rate_ok(total, gst, lo=0.04, hi=0.15):
    subtotal = total - gst
    return subtotal > 0 and lo <= gst / subtotal <= hi


def _extract_fields(raw_text: str) -> dict:
    text = _normalize(raw_text)
    lines = [l.strip() for l in text.strip().split('\n') if l.strip()]

    result = {k: '' for k in COLUMNS}

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

    subtotal_match = None
    for pattern in [
        r'Sub[\s\-]?[Tt]otal\s*:?\s*\$?(\d+[\.,]\d+)',
        r'Food\s+(?:&\s+Bev(?:erage)?s?)?\s+Sub[\s\-]?[Tt]otal\s*:?\s*\$?(\d+[\.,]\d+)',
        r'SUB[\s\-]?(?:TTL|TL|TOT(?:AL)?|T)\b[^\n]{0,20}?(\d+[\.,]\d+)',
        r'sub[\s\-_\.]*tota[^\n]{0,15}?(\d+[\.,]\d+)',
        r'(?:Before\s+GST|Excl\.?\s+GST|Excl\.\s+Tax)\s*:?\s*\$?(\d+[\.,]\d+)',
        r'Nett?\s+(?:Amount|Total)\s*:?\s*\$?(\d+[\.,]\d+)',
    ]:
        subtotal_match = re.search(pattern, text, re.IGNORECASE)
        if subtotal_match:
            break
    if subtotal_match:
        result['subtotal'] = subtotal_match.group(1).replace(',', '.')

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
        amounts = [a for a in re.findall(r'\$?(\d+[\.,]\d+)', text) if float(a.replace(',', '.')) > 1]
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

    has_service_charge = bool(re.search(r'serv(?:ice)?\s*ch(?:g|arge)|svc\s*ch(?:g|arge)', text, re.IGNORECASE))
    inclusive_gst = bool(re.search(r'(?:inclusive\s+of\s+(?:\d+%\s*)?GST|incl\.?\s+GST|GST\s+incl(?:uded)?)', text, re.IGNORECASE))

    if not result['gst'] and result['total'] and inclusive_gst:
        try:
            total = float(result['total'])
            result['gst'] = f"{total * 9 / 109:.2f}"
            result['subtotal'] = f"{total * 100 / 109:.2f}"
        except ValueError:
            pass

    if not result['subtotal'] and result['total'] and result['gst'] and not has_service_charge:
        try:
            result['subtotal'] = f"{float(result['total']) - float(result['gst']):.2f}"
        except ValueError:
            pass

    return result


@st.cache_data(show_spinner=False)
def _process_receipt(file_bytes: bytes, filename: str):
    """OCR a receipt. Cached by file content — re-uploads never re-run Tesseract."""
    if filename.lower().endswith(".pdf"):
        image = _pdf_to_image(file_bytes)
    else:
        image = Image.open(io.BytesIO(file_bytes))

    processed = _preprocess(image)
    color_image = ImageOps.exif_transpose(image).convert("RGB")

    text_color = pytesseract.image_to_string(color_image)
    fields_color = _extract_fields(text_color)

    key_fields = ["date", "time", "gst", "total"]
    score_color = sum(1 for f in key_fields if fields_color.get(f))

    if score_color < len(key_fields):
        text_bw = pytesseract.image_to_string(processed)
        fields_bw = _extract_fields(text_bw)
        score_bw = sum(1 for f in key_fields if fields_bw.get(f))

        base, overlay = (fields_bw, fields_color) if score_color >= score_bw else (fields_color, fields_bw)
        fields = base.copy()
        for f in key_fields + ["merchant", "subtotal"]:
            if not fields.get(f) and overlay.get(f):
                fields[f] = overlay[f]
        best_text = text_color if score_color >= score_bw else text_bw
    else:
        fields = fields_color
        best_text = text_color

    orig_buf = io.BytesIO()
    image.save(orig_buf, format="PNG")
    proc_buf = io.BytesIO()
    processed.save(proc_buf, format="PNG")

    return fields, best_text, orig_buf.getvalue(), proc_buf.getvalue()


def _make_excel(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Receipts")
        if not df.empty:
            ws = writer.sheets["Receipts"]
            ref = f"A1:{get_column_letter(len(df.columns))}{len(df) + 1}"
            tab = Table(displayName="Receipts", ref=ref)
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
    return buf.getvalue()


# ── Upload & process ──────────────────────────────────────────────────────────

uploaded_files = st.file_uploader(
    "Upload receipts (images or PDFs)",
    type=["jpg", "jpeg", "png", "pdf", "heic"],
    accept_multiple_files=True,
)

if uploaded_files:
    new_rows = []
    skipped = []

    for uf in uploaded_files:
        file_bytes = uf.read()
        file_hash = hashlib.md5(file_bytes).hexdigest()

        if file_hash in st.session_state.seen_hashes:
            skipped.append(uf.name)
            continue

        with st.spinner(f"Processing {uf.name}…"):
            fields, best_text, orig_bytes, proc_bytes = _process_receipt(file_bytes, uf.name)

        with st.expander(f"Preview: {uf.name}"):
            col1, col2 = st.columns(2)
            with col1:
                st.write("Original")
                st.image(orig_bytes, use_column_width=True)
            with col2:
                st.write("Preprocessed")
                st.image(proc_bytes, use_column_width=True)

        with st.expander(f"Raw OCR text: {uf.name}"):
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

        if st.button("Save to master", type="primary"):
            combined = pd.concat([st.session_state.master_df, edited_df], ignore_index=True)
            combined["_d"] = pd.to_datetime(combined["date"], dayfirst=True, errors="coerce")
            combined = combined.sort_values("_d", ascending=True, na_position="last").drop(columns=["_d"]).reset_index(drop=True)
            st.session_state.master_df = combined
            st.session_state.seen_hashes |= {r["_hash"] for r in new_rows}
            st.success(f"Saved {len(edited_df)} receipt(s).")

            st.download_button(
                label="Download master.xlsx",
                data=_make_excel(combined),
                file_name="master.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    else:
        if not skipped:
            st.info("No new receipts to process.")

st.divider()

# ── Danger zone ───────────────────────────────────────────────────────────────

with st.expander("Danger Zone"):
    st.warning("This will clear all receipts from your current session.")
    if st.button("Clear all data", type="primary"):
        st.session_state.master_df = pd.DataFrame(columns=COLUMNS)
        st.session_state.seen_hashes = set()
        st.success("Session cleared.")
        st.rerun()

# ── All receipts ──────────────────────────────────────────────────────────────

st.subheader("All Receipts")

master_df = st.session_state.master_df.copy()
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

    st.download_button(
        label=f"Download filtered ({len(filtered)} rows)",
        data=_make_excel(filtered[COLUMNS].reset_index(drop=True)),
        file_name="receipts_filtered.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
