# Apple Inc. (AAPL) 3-Statement Model — Assumptions & Scenario Write-Up

## Overview

This is a fully-linked 3-statement model (Income Statement, Balance Sheet, Cash Flow Statement) for Apple
Inc., built from real FY2023–FY2025 SEC filings, with a 3-year driver-based forecast (FY2026E–FY2028E) and
a "what if revenue drops 10%" scenario built into the model itself via a dropdown toggle.

**File:** `model/AAPL_3statement_model.xlsx`

## Data sources (all primary)

| Period | Source |
|---|---|
| FY2025A, FY2024A | Apple Inc. Form 10-K (FY2025), SEC EDGAR accession 0000320193-25-000079; Q4 FY2025 8-K Ex-99.1, filed 10/30/2025 |
| FY2023A | Apple Inc. Form 10-K (FY2023), SEC EDGAR accession 0000320193-23-000106; Q4 FY2023 8-K Ex-99.1, filed 11/2/2023 |

Every hardcoded historical figure carries a cell comment citing its exact source. No figures were estimated
or pulled from secondary aggregator sites for the historical periods.

## Forecast methodology

- **Revenue:** grown off a single "Active Growth Rate" assumption per year, itself driven by a scenario
  selector (Base / Bear / Bull) on the Assumptions tab.
- **Margins & opex (Gross Margin %, R&D % of revenue, SG&A % of revenue, tax rate):** explicit blue-input
  assumptions per forecast year, anchored to Apple's recent historical ranges (e.g., gross margin trending
  from 44.1% in FY23A to 46.9% in FY25A; FY26E–FY28E assumes continued modest expansion to 47.2–47.7%).
- **Working capital (AR, Inventory, AP):** forecast using Days Sales Outstanding (33 days), Days Inventory
  Outstanding (9 days) and Days Payable Outstanding (113 days) — derived from Apple's FY2025A actuals —
  rather than a flat percent-of-revenue, which is the standard FP&A approach for working capital.
- **Other current/non-current assets & liabilities, vendor non-trade receivables, deferred revenue:**
  forecast as a fixed % of revenue, based on FY2025A actual ratios.
- **Capex and D&A:** modeled as % of revenue (3.0% and 2.8% respectively), consistent with Apple's recent
  capital intensity.
- **Capital returns:** dividends grow 4%/year off the prior year's actual payment; buybacks are modeled as
  75% of Free Cash Flow (CFO − Capex), and the resulting cash outflow reduces the diluted share count
  (using an assumed average share price), which flows directly into forecasted diluted EPS.

## Simplifications (explicitly called out, not hidden)

1. **Marketable securities (current & non-current) and total debt (commercial paper + term debt) are held
   flat in the forecast.** Apple actively manages its investment portfolio and debt maturities, but modeling
   that activity in detail would require assumptions about future market conditions that are outside the
   scope of an operating-business forecast. This is a standard simplification in a first-pass operating model.
2. **Equity is rolled forward as a single balance** (prior equity + Net Income − Dividends − Buybacks − Tax
   Withholding on Equity Awards + Stock-Based Compensation) rather than split into Common Stock/APIC,
   Retained Earnings, and AOCI sub-lines for the forecast years. Historicals retain the full breakdown.
3. **"Other" non-cash items** in the cash flow statement are assumed to be zero in the forecast (they were
   small and non-recurring in the historical actuals).

## Balance check

Row 40 of the Balance Sheet tab (`Total Assets − Total Liabilities & Equity`) is **zero in every single
period**, including all three forecast years and under all three scenarios. This isn't cosmetic — it's the
proof that every balance sheet account that moves is matched by a corresponding line in the cash flow
statement or the income statement. This is the standard a real FP&A model is held to before it ever goes in
front of a CFO.

## Scenario: "What if revenue drops 10%?"

The Assumptions tab (cell B4) has a dropdown: **Base**, **Bear (-10% Rev Shock)**, **Bull (Upside)**. Picking
"Bear" sets FY2026E revenue growth to **-10%** instead of the Base Case's +8%, while every other assumption
(margins, opex ratios, tax rate, working capital, capex) stays unchanged — isolating the pure effect of the
revenue shock.

**Result (FY2026E, Bear vs. Base):**

| Metric | Base Case | Bear Case (-10% Rev) | Δ |
|---|---:|---:|---:|
| Net Sales | $449.5B | $374.5B | -16.7% |
| Net Income | $121.9B | $101.6B | **-16.7%** |

**Why the bottom line falls by the same ~17%, not 18pp:** because the model holds R&D% and SG&A% of revenue
fixed, operating expenses scale down proportionally with revenue, so the net income hit roughly tracks the
revenue hit in this simplified model. In a real-world downturn, R&D and SG&A are largely *fixed* in the
short run (headcount, leases, committed programs), so an actual 10% revenue shock would compress margins
*and* hit net income by meaningfully more than 10% — which is exactly the kind of insight an FP&A analyst
would flag to leadership alongside the model output. The Sensitivity Table tab makes this trade-off explicit
across a full grid of revenue growth and gross margin combinations.

The balance sheet still balances to zero under the Bear case — cash builds more slowly, buybacks shrink
(since they're tied to Free Cash Flow), and the diluted share count declines less, all flowing through
automatically from the same formulas used in the Base Case.
