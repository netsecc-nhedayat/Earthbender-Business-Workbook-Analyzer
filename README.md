# Earthbender Business Insights (prototype)

Streamlit prototype built against the real Earthbender workbook schema.
Reads CSV / XLSX / XLSM / XLS, builds a dashboard, flags likely data issues,
generates plain-Python recommendations, and exports a PDF + cleaned CSV.

## Install (Ubuntu)

sudo apt update
sudo apt install -y python3 python3-pip python3-venv
mkdir -p ~/earthbender && cd ~/earthbender
# copy app.py, requirements.txt, README.md here
python3 -m venv .venv
source .venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt

## Run

streamlit run app.py

Streamlit opens at http://localhost:8501.

## What was built from your workbook

Auto-detected columns (exact header matches from your file):
    JobID, Customer, Job Start, Job End, Material, Type, Scope, Lead Source,
    Quote $, Total Pay, Total Cost $, Profit $, Profit %, Hours worked,
    Workers?, Suburb, Learnings

Currency strings like "$2,700" are parsed to numbers. Percent strings like
"30%" are parsed to fractions. The original file is never modified.

## Expected flags on your workbook (rough)

- End-before-start: J009, J075, J287 (and possibly J298/J299 which show end
  before start by one day).
- Missing JobID: J239, J240.
- Duplicate JobIDs: J021/J022, J124/J125, J158/J159.
- Unusual margins (>150% or <−50%): rows like J031 (100%), J047 (100%),
  J117 (100%), J151 (100%), J161 (100%), J166 (100%), J171 (100%).
  These are mostly labour-only jobs — flagged for review, not declared wrong.
- "Workers?" empty on ~60 rows in later part of workbook (J170+).
- Some rows show Profit % inconsistent with Profit ÷ Quote (e.g. J192, J305).

## Limitations

- Column detection uses regex against headers; always verify the mapping screen.
- "Unusual margin" uses a fixed heuristic of −50% to +150%.
- PDF charts rendered with matplotlib (no Kaleido/Chrome dependency).
- Prototype, not production.

## Files

- app.py           — everything
- requirements.txt — deps
- README.md        — this file
