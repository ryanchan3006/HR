# Contract Generator — Setup & Usage Guide

## What it does
Reads candidate data from Excel, fills a Word contract template, lets HR
review and approve each contract, then exports approved ones as PDFs into
a chosen folder.

---

## Requirements

| Requirement | Purpose |
|---|---|
| Python 3.9+ | Run the script |
| python-docx | Read/write Word templates |
| openpyxl | Read Excel data |
| LibreOffice | Convert Word → PDF (free) |

---

## Setup (first time only)

1. Install Python from https://python.org  
2. Install LibreOffice from https://www.libreoffice.org  
3. Open a terminal in this folder and run:

```
pip install -r requirements.txt
```

---

## Running the app

```
python app.py
```

Or double-click `ContractGenerator.exe` if you have built the executable.

---

## Building the EXE (Windows)

Double-click `build.bat` — it installs PyInstaller and creates
`dist\ContractGenerator.exe`. The EXE bundles Python and all libraries.
LibreOffice still needs to be installed separately for PDF export.

---

## Template rules

- Use `{{Column Name}}` as placeholders in your Word template.
- The placeholder name must exactly match the column header in your Excel file.
- Example: if your Excel column is `Start Date`, use `{{Start Date}}` in the template.

---

## Excel file structure

**Annex B sheet** — candidate data with one row per candidate.  
Required columns (names can vary):
- Full Name / Candidate Name
- Job Title / Position
- Start Date
- Salary
- Rank (used to look up signatory)
- Address

**Annex C sheet** — signatory mapping.  
Required columns:
- Rank
- Signatory Name
- Signatory Title

---

## Workflow

1. **Generate tab** — select Excel file, Word template, output folder → click Generate
2. **Edit Template tab** — optionally edit template text or insert fields
3. **Review & Approve tab** — review each contract, approve or reject
4. Click **Export approved →** — PDFs saved to your output folder, named `Contract_[Name].pdf`

---

## Troubleshooting

| Problem | Fix |
|---|---|
| PDF export fails | Make sure LibreOffice is installed and `soffice` is on your PATH |
| Placeholders not replaced | Check placeholder name exactly matches Excel column header |
| No candidates loaded | Ensure Annex B sheet exists and has a header row |
| App won't start | Run `pip install -r requirements.txt` first |
