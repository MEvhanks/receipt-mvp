# Taming the Receipt Monster

A local web app that reads receipt images and PDFs, extracts GST and expense data, and saves everything to Excel.

## What it does

- Upload multiple receipts at once (JPG, PNG, or PDF)
- Automatically extracts: merchant, date, time, GST amount, subtotal, total
- Shows an editable table so you can fix any mistakes before saving
- Appends to a master Excel file (skips duplicates automatically)
- Download the updated Excel with one click

## How to run

Open Terminal and run these commands:

```
cd ~/Desktop/receipt-mvp
source venv/bin/activate
streamlit run app.py
```

Then open your browser to: http://localhost:8501

## Requirements

- Python 3.9+
- Tesseract OCR (installed via Homebrew: `brew install tesseract`)
- All Python dependencies listed in requirements.txt

## First time setup

```
cd ~/Desktop/receipt-mvp
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

# 🚀 Demo Instructions

---

# 🟡 Before the demo

## 1. Go to your project folder

```bash
cd ~/Desktop/"Claude Code"/Projects/receipt-mvp

rm master.xlsx

source venv/bin/activate
streamlit run app.py

🎬 During the demo

Follow this exact flow:

1. Upload receipts
Upload 2–3 receipts at once
Show multi-upload working
2. Show image previews
Point out that uploaded receipts expand visually
3. Show extracted table

Highlight:

Merchant
Date
GST fields
4. Live edit demo
Edit one cell in the table
Show instant update
5. Save to Excel

Click:

Save to master.xlsx

Explain that data is stored locally in Excel format

6. Download file

Click:

Download master.xlsx

Open in Excel to show final output

7. Duplicate detection test
Upload the same receipt again

Expected result:

Skipped 1 duplicate