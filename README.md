# DCF Valuation Tool

![Python](https://img.shields.io/badge/Python-3.8%2B-blue?logo=python&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green)
![Excel Output](https://img.shields.io/badge/Output-Excel%20%2B%20Terminal-217346?logo=microsoft-excel&logoColor=white)
![Platform](https://img.shields.io/badge/Platform-macOS%20%7C%20Linux%20%7C%20Windows-lightgrey)

A terminal-based equity valuation script that runs four models, prints a structured report, and exports a Excel workbook — no server, no browser, no GUI.

Built as a personal finance project at the intersection of corporate finance theory, discounted cash flow modeling, and software engineering.

---

## What it does

Every time it runs, the tool:

1. Prompts the analyst through 6 structured input sections (company, cost of capital, capital structure, FCF projections, Gordon/multiple inputs, sensitivity)
2. Validates hard constraints before any computation (WACC > g, non-zero FCFs)
3. Runs four valuation models simultaneously and prints a ranked results table
4. Builds a 13×13 sensitivity matrix varying WACC and terminal g independently
5. Generates a four-sheet Excel workbook where every computed cell is a live formula
6. Saves the file as `DCF_{TICKER}_{YYYYMMDD_HHMM}.xlsx` in the working directory

---

## Models

| # | Model | Formula | Discount Rate |
|---|-------|---------|---------------|
| 1 | **Traditional DCF** | PV(FCFs) + Gordon TV | WACC |
| 2 | **Gordon — Dividend DDM** | P = D1 / (Ke − g) | Ke |
| 3 | **Gordon — FCF Terminal** | TV = FCF_n × (1+g) / (WACC−g) | WACC |
| 4 | **Sector Multiple** | EV = Multiple × Metric | — |
| 5 | **Sensitivity Matrix** | 13×13 grid, ±0.01 to ±0.10 shocks | WACC |

**Gordon Dividend DDM** discounts dividend streams using the cost of equity — applicable to mature, dividend-paying firms.

**Gordon FCF Terminal** is the terminal value mechanism inside the DCF presented as a standalone output — isolates how much of the DCF price is driven by the perpetuity assumption alone.

---

## Sensitivity Matrix

The matrix varies WACC and terminal g independently around the base-case inputs using fixed shock steps:

```
±0.01  ±0.02  ±0.03  ±0.04  ±0.05  ±0.10
```

This produces a **13×13 grid (169 cells)**. Row and column headers show the shock relative to the base case, not absolute rates — readable regardless of what the base inputs are.

- The base-case cell is highlighted with a thick green border in Excel
- Cells where WACC ≤ g are marked **N/A** (model undefined)

---

## Excel Output

Four sheets, dark-navy theme, maximum formula referencing.

| Sheet | Content |
|-------|---------|
| **Summary** | Inputs table, valuation results (4 models), DCF breakdown  |
| **FCF Projections** | Year-by-year table: FCF, Discount Factor, PV, % of Total PV, Cumulative PV, FCF/Share — entirely formula-driven |
| **Sensitivity** | Color-coded 13×13 heatmap with base-case highlight |
| **Inputs Log** | Single source of truth — all raw numeric inputs stored here; every other sheet references this one |

> Formula architecture: changing any number in **Inputs Log** cascades to Summary and FCF Projections automatically. No hardcoded values in computed cells.

**Sensitivity color scale:**

| Color | Signal |
|-------|--------|
| 🟢 Green | Upside > 20% |
| 🟩 Light green | Upside 5—20% |
| 🟡 Yellow | Upside ±5% (neutral) |
| 🟠 Orange | Downside 5—20% |
| 🔴 Red | Downside > 20% |

---

## Terminal Output Sample

```
════════════════════════════════════════════════════════════
RESULTS — Apple Inc. (AAPL)
════════════════════════════════════════════════════════════
MODEL                   PRICE       UPSIDE   SIGNAL
────────────────────────────────────────────────────────────
Traditional DCF         USD  13.24   -62.2%  ▼ SELL
Gordon — Dividend DDM   USD  31.25   -10.7%  ▼ SELL
Gordon — FCF Terminal   USD  16.97   -51.5%  ▼ SELL
Sector Multiple         USD  12.00   -65.7%  ▼ SELL
────────────────────────────────────────────────────────────
Market Price            USD  35.00

DCF BREAKDOWN
────────────────────────────────────────────────────────────
PV FCF Year 1                    9,049,773.76
PV FCF Year 2                    9,827,808.60
PV FCF Year 3                   10,376,268.51
PV FCF Year 4                   10,061,023.12
PV FCF Year 5                    9,711,998.18
PV of Projected FCFs            49,026,872.17
PV of Terminal Value           133,378,108.40
Enterprise Value (EV)          182,404,980.57
(-) Net Debt                    50,000,000.00
Equity Value                   132,404,980.57
Implied Price (DCF)                     13.24
TV / EV                                 73.1%
```

---

## Code Architecture

The script is organized into **9 labeled blocks**, each with a standardized comment header:

```python
# BLOCK N — NAME
#
# WHAT IT DOES:  what the function computes and returns
# WHY THIS WAY:  design decisions and trade-offs
# DEPENDENCIES:  which blocks feed into this one
# USED IN:       where this block's output is consumed
```

| Block | Name | Role |
|-------|------|------|
| 0 | Imports & Config | Color constants, library imports |
| 1 | Excel Style Helpers | `sc()`, `mh()`, `ref()` — cell styling primitives |
| 2 | Input Helpers | `ask()`, `ask_str()`, `ask_fcfs_by_year()` |
| 3 | Traditional DCF | Multi-period DCF with Gordon TV |
| 4 | Gordon Growth | Dividend DDM + FCF-based TV variant |
| 5 | Sector Multiple | EV = Multiple × Metric |
| 6 | Sensitivity Analysis | 13×13 fixed-shock matrix |
| 7 | Terminal Report | Formatted terminal output |
| 8 | Excel Builder | Four-sheet workbook with live formulas |
| 9 | Main | Input flow, validation, orchestration |

---

## Input Validation

Two hard constraints enforced before any calculation runs:

- **WACC > g** — the DCF terminal value is mathematically undefined when WACC ≤ g (Gordon denominator collapses to zero or negative). The script exits with an explanation instead of crashing.
- **Non-zero FCFs** — all-zero FCF inputs produce EV = 0 and cascade ZeroDivisionErrors downstream. Caught explicitly.

---

## Requirements

```
Python 3.8+
numpy
pandas
openpyxl
```

Install:

```bash
pip install numpy pandas openpyxl
```

---

## Usage

```bash
python dcf_valuation.py
```

The script walks through **6 input sections** in sequence:

| Section | Prompts |
|---------|---------|
| 1 / 6 — Company Identification | name, ticker, sector, currency |
| 2 / 6 — Cost of Capital | WACC, Ke, Kd, tax rate |
| 3 / 6 — Capital Structure | net debt, shares, market price |
| 4 / 6 — Free Cash Flow | start year, horizon, FCF per year, terminal g |
| 5 / 6 — Gordon & Multiple | D1, Gordon g, multiple type/value, metric |
| 6 / 6 — Sensitivity | (auto) ±0.01 to ±0.10 grid centered on base inputs |

Every prompt has a default — press **Enter** to accept. The FCF section prompts each year individually by calendar label to prevent positional ambiguity.

After the last input the script:
1. Prints the full report to the terminal
2. Saves `DCF_{TICKER}_{YYYYMMDD_HHMM}.xlsx` in the working directory

---

## File Output

```
DCF_{TICKER}_{YYYYMMDD_HHMM}.xlsx
```

Saved in the working directory. The timestamp in the filename allows multiple runs on the same ticker without overwriting prior versions.

---

## Limitations

- **Single-path model** — no Monte Carlo, no scenario branching. Uncertainty is captured only through the sensitivity matrix.
- **Gordon TV assumes steady state** — the terminal growth rate applies forever after year n. Inappropriate for firms in structural decline or high-growth transition.
- **Multiples are user-supplied** — the script does not fetch comparable company data; the analyst must research and input the sector multiple manually.
- **No currency conversion** — all inputs assumed to be in the same currency unit.

---

## Background

I built this during my BS Economics program (minors in Finance and Data Analysis) as a way to apply corporate finance theory, discounted cash flow modeling, and software engineering to a real valuation problem.

DCF is one of the most foundational frameworks in equity research — every CFA candidate learns it, every sell-side analyst uses it — but the raw output creates false precision: a single implied price with no indication of how sensitive it is to the assumptions underneath it.

The goal was to build something rigorous enough to use as an actual analytical tool, not just a class project. That meant going beyond a textbook DCF: adding two Gordon Growth variants to cross-check the terminal value assumption from different angles, building a 13×13 sensitivity matrix with fixed shock steps so the valuation band is immediately visible, and engineering the Excel output so every computed cell is a live formula — change one input and the entire workbook updates.

This tool is for personal use. I use it to run quick valuations on stocks I’m researching, stress-test my own assumptions before forming a position view, and practice the kind of structured financial modeling I’ll be doing professionally in quantitative finance, asset research, or risk management.

> This is not investment advice and makes no claim to production-grade accuracy.
