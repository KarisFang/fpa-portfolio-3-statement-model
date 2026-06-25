# Apple Inc. (AAPL) — 3-Statement Financial Model & Forecast

A fully linked 3-statement model (Income Statement, Balance Sheet, Cash Flow Statement) for Apple Inc.,
built from real FY2023–FY2025 SEC filings, with a 3-year driver based forecast and a built in "what if
revenue drops 10%" scenario toggle. 

**Portfolio project #1 of 3** mapping directly to FP&A job descriptions. ([Project 2: automated reporting
pipeline](https://github.com/KarisFang/fpa-portfolio-reporting-pipeline) · [Project 3: budget-vs-actual variance dashboard](#) — coming soon)

## What's in this repo

| File | Description |
|---|---|
| [`model/AAPL_3statement_model.xlsx`](model/AAPL_3statement_model.xlsx) | The model. 7 tabs, 357 live formulas, zero formula errors, balance check = 0 in every period. |
| [`docs/ASSUMPTIONS_AND_SCENARIO_WRITEUP.md`](docs/ASSUMPTIONS_AND_SCENARIO_WRITEUP.md) | 1-page write-up of every assumption, every simplification, and the revenue-shock scenario result. |
| [`build_model.py`](build_model.py) | The Python (openpyxl) script that generates the workbook from scratch — included so the build is fully reproducible and auditable. |

## Highlights

- **Real data, not synthetic:** every historical figure (FY2023A–FY2025A) is sourced directly from Apple's
  10-K filings and 8-K earnings releases on SEC EDGAR, with the source cited in a cell comment next to the figure.
- **A scenario dropdown that actually drives the whole model:** switch `Assumptions!B4` between *Base*,
  *Bear (-10% Rev Shock)*, and *Bull (Upside)* and every number across all three statements recalculates.
- **A live balance check:** `Total Assets − (Total Liabilities + Equity)` equals exactly **zero** in every
  one of the 6 periods (3 actual, 3 forecast) and under all 3 scenarios — proof the statements are properly integrated.
- **Working capital modeled with DSO/DIO/DPO**, not flat percentages — the way it's actually done in practice.
- **A 2-way sensitivity grid** (Net Income & Free Cash Flow vs. Revenue Growth × Gross Margin) on its own tab.

## How to use it

1. Download `model/AAPL_3statement_model.xlsx` and open in Excel (or Google Sheets / LibreOffice).
2. Go to the **Assumptions** tab, cell `B4`, and pick a scenario from the dropdown.
3. Watch the **Income Statement**, **Balance Sheet**, and **Cash Flow Statement** tabs update.
4. See **Scenario Analysis** for a side-by-side Base/Bear/Bull comparison, and **Sensitivity Table** for the
   full revenue-growth × gross-margin grid.

## Rebuilding it from scratch

```bash
pip install openpyxl
python3 build_model.py
```

This regenerates `model/AAPL_3statement_model.xlsx` identically — useful for extending the model (add a
segment-level revenue build, add a debt schedule, etc.) without hand-editing formulas in Excel.

---
Built by Karis Fang as a portfolio project for FP&A roles.
