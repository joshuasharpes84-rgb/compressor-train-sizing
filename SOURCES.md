# Sources — Compressor Train Sizing

Where each type of information in this project (decks, `Compressor Train Sizing.xlsx`, the portfolio
calculator) comes from. Topic-level, not number-by-number: the aim is to show nothing was invented,
and to point a reader to the right book and section if they want to dig.

**Tiers**
- **Library** — reference textbook in `Downloads\ChemE\Textbooks (Level 1)\` (used across projects)
- **Project PDF** — paper or excerpt held in this folder
- **Web** — public page, URL given
- **Print only** — book cited from a paper copy; no digital copy for page lookup
- **Own work** — derived, fitted or chosen in this project; the basis is stated

Copyrighted PDFs are cited here but never published.

| Topic | Source | Tier | Location |
|---|---|---|---|
| Compressor type selection (coverage map, Q vs adiabatic head) | Perry's *Chemical Engineers' Handbook*, Fig. 10-64; *Compressor Selection* deck | Library | `Textbooks (Level 1)\Handbooks & Quick Reference\Perry's 9th Edition.pdf`, Sec. 10 |
| Compressor types and their service ranges (selection deck) | Cheremisinoff, *Chemical Engineer's Condensed Encyclopedia of Process Equipment*, Gulf, 2000; Perry's | Print only + Library | Encyclopedia: no digital copy yet |
| Polytropic head, temperature rise, polytropic exponent n | Perry's Fig. 10-54 (polytropic curves); *Compressor Train Sizing* deck | Library | Perry's Sec. 10 |
| Allowable discharge temperature ranges by type (300–350 °F recip., 392–446 °F centrifugal, 212–350 °F rotary) | *Compressor Train Sizing* deck table | Own deck | `Compressor Train Sizing.pptx` |
| Efficiency: empirical curves (reciprocating vs pressure ratio; centrifugal and axial vs suction flow) | *Compressor Train Sizing* deck, Figs. 1 and 3 (digitized) | Own deck | `reference_charts\` |
| Efficiency: scale-up method | Casey & Robinson (scale-up efficiency correlation); friction-factor equations not held | Cited in deck | — |
| Efficiency: specific speed / specific diameter | Balje N_s–D_s chart (chart itself not held) | Cited in deck | — |
| Compressor and intercooler purchase cost | Seider, Lewin, Seader, Widagdo, *Product and Process Design Principles*, Ch. 16.5 (Eq. 16.36–16.45, Table 16.25) | Library + Project PDF | `Textbooks (Level 1)\Process & Plant Design\Product and Process Design Principles (Sieder, Lewin) ...pdf`; excerpt `compressor_cost_comparison_PT_vs_Seider.pdf` |
| Cash-flow definitions, cooling-water basis (90/120 °F), U = 50 Btu/hr·ft²·°F precedent (Ex. 16.10) | Seider Ch. 16–17 | Library | as above |
| Payback period and minimum acceptable return (Table 8-1, Eq. 8-2a/8-2c, worked example) | Peters, Timmerhaus & West, *Plant Design and Economics for Chemical Engineers*, Ch. 8 | Library + Project PDF | `Textbooks (Level 1)\Process & Plant Design\Plant Design and Economics ... Peters.pdf`; `PT_payback_period_p345.pdf` |
| Cost-correlation comparison (P&T Ch. 12 vs Seider Ch. 16) | Peters & Timmerhaus Ch. 12, Figs. 12-28 to 12-30; Seider Ch. 16.5 | Project PDF | `compressor_cost_comparison_PT_vs_Seider.pdf` |
| Total capital investment build-up (bare-module factors, contingency 18 %, fee 3 %, working capital 15 %) | Economics project method (Seider-based) | Own work | `ChemE\FEL Projects\Economics\notes\01_capital_cost_estimation.md` |
| CEPCI annual averages 2001–2024 | Chemical Engineering magazine index, via toweringskills.com | Web | `CEPCI_reference.txt` |
| Tax rate 21 % | US federal corporate rate after the 2017 TCJA; Texas has no state corporate income tax | Web | — |
| Intercooler tube/shell side rules | Coulson & Richardson (via *Heat Exchanger Heuristics* deck) | Library | `Textbooks (Level 1)\Process & Plant Design\Chemical Engineering Design (Coulson, Richardson) ...pdf` |
| Gas library MW and k (20 gases, ideal gas near ambient) | Standard property values | Own work | `build_compressor_workbook.py` |
| Constant ηp across the 1–6 stage sweep; Fig. 10-64 boundaries read by eye | Class 5 screening simplifications | Own work | flagged in workbook |
