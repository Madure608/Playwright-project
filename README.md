# Playwright Test Automation — Assignment 1

This repository contains a Playwright-based test automation script for the Transliteration frontend.

Contents
- `test_automation.py` — main test runner script (Python + Playwright).
- `Assignment 1 - Test cases.xlsx` — input spreadsheet with test cases.
- `Assignment 1 - Test cases_results.xlsx` — results file produced by `test_automation.py`.

Requirements
- Python 3.8+
- Playwright
- openpyxl

Quick setup (Windows / PowerShell):

```powershell
cd "D:\test_automation (1)\test_automation"
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install --upgrade pip
pip install playwright openpyxl
playwright install
```

Run the tests:

```powershell
python test_automation.py --output "Assignment 1 - Test cases_results.xlsx"
```

Useful options:
- `--headless` run browser without UI
- `--wait-ms` maximum wait for a translation to complete (ms)
- `--retries`, `--retry-wait-ms` adjust retry behavior
- `--type-delay-ms` typing delay (ms)
- `--output` path for results workbook

Debugging
- If output cells are empty, the script prints debug info about visible textareas and button state.
- Ensure the frontend page is reachable by network and ads/overlays do not block UI.

License
- Use as needed; no license specified.
