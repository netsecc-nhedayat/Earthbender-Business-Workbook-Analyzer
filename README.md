# Earthbender Business workbook analyzer

Streamlit prototype built against the real Earthbender workbook schema.
Reads CSV / XLSX / XLSM / XLS, builds a dashboard, flags likely data issues,
generates plain-Python recommendations, and exports a PDF + cleaned CSV.

## Install (Linux)

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
markdown
## Part 1-B — Installing on Windows

This section is the Windows equivalent of Part 1. You do **not** need any
previous Python experience.

### 1-B.1 Install Python

1. Open <https://www.python.org/downloads/windows/> in your browser.
2. Download the latest **Python 3.12.x** (or 3.11+) **64-bit installer**.
3. Run the installer. On the first screen:
   - Tick **"Add python.exe to PATH"** (bottom of the window — very important).
   - Then click **Install Now**.
4. When it finishes, click **Close**.

### 1-B.2 Open PowerShell and verify Python

1. Press `Win` key, type **PowerShell**, and open **Windows PowerShell**.
2. Run:

   ```powershell
   python --version
Expected output: Python 3.12.x (or similar).

If you see "python is not recognized", close PowerShell, reopen it, and
try again. If it still fails, reinstall Python and make sure the
"Add python.exe to PATH" box is ticked.

1-B.3 Create a project folder
In PowerShell:

powershell
cd $HOME
mkdir earthbender
cd earthbender
Copy app.py, requirements.txt, README.md and USER_MANUAL.md into this
folder (via File Explorer drag-and-drop, or Copy-Item).

1-B.4 Create an isolated environment
powershell
python -m venv .venv
If PowerShell blocks the activation script (you'll see a red "running scripts
is disabled on this system" error), allow it once for your user:

powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
Type Y when prompted. Then activate:

powershell
.\.venv\Scripts\Activate.ps1
Your prompt should now start with (.venv).

Using Command Prompt (cmd.exe) instead of PowerShell?
Activate with .\.venv\Scripts\activate.bat instead.

1-B.5 Install dependencies
powershell
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
This takes 1–3 minutes on the first run and needs an internet connection.

1-B.6 Launch the app
powershell
streamlit run app.py
Your default browser opens at http://localhost:8501. To stop the app, press
Ctrl + C in PowerShell and confirm with Y.

Re-launching later (Windows)
Every time you want to use the app:

powershell
cd $HOME\earthbender
.\.venv\Scripts\Activate.ps1
streamlit run app.py

## Limitations

- Column detection uses regex against headers; always verify the mapping screen.
- "Unusual margin" uses a fixed heuristic of −50% to +150%.
- PDF charts rendered with matplotlib (no Kaleido/Chrome dependency).
- Prototype, not production.

## Files

- app.py           — everything
- requirements.txt — deps
- README.md        — this file
