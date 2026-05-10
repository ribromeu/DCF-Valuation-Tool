# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).

---

## [1.0.0] — 2025-05-09

### Added

- **Traditional DCF model** — multi-period discounted cash flow with Gordon Growth terminal value
- **Gordon — Dividend DDM** — dividend discount model using cost of equity (Ke)
- **Gordon — FCF Terminal** — standalone perpetuity output to isolate terminal value sensitivity
- **Sector Multiple model** — EV = Multiple × Metric for comparable-based valuation
- **13×13 Sensitivity Matrix** — fixed shock steps (±0.01 to ±0.10) varying WACC and terminal g independently
- **Structured terminal report** — ranked results table with upside/downside signals printed to stdout
- **Four-sheet Excel workbook** (dark-navy theme, fully formula-referenced):
  - `Summary` — all valuation results as live `=FORMULA` cells
  - `FCF Projections` — year-by-year DCF table entirely formula-driven
  - `Sensitivity` — color-coded 13×13 heatmap with base-case highlight
  - `Inputs Log` — single source of truth; all other sheets reference this one
- **Hard input validation** before any computation:
  - WACC > g enforcement (Gordon denominator check)
  - Non-zero FCF check (prevents ZeroDivisionError cascade)
- **6-section interactive prompt flow** with per-field defaults (press Enter to accept)
- **Timestamped output filename** — `DCF_{TICKER}_{YYYYMMDD_HHMM}.xlsx` prevents overwrite across runs
- **9-block code architecture** with standardized inline comment headers (`WHAT IT DOES`, `WHY THIS WAY`, `DEPENDENCIES`, `USED IN`)

---

## [Unreleased]

### Planned

- Monte Carlo simulation layer over the DCF
- Automatic comparable fetching for sector multiples
- Currency conversion support
- CLI argument parsing (non-interactive mode)
