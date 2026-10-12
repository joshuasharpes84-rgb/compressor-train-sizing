# Compressor Train Sizing Calculator: Type, Efficiency, Power & Stage-Count Optimizer

Sizes a multistage gas compressor train with equal pressure ratio per stage and full intercooling.
It auto-selects the compressor type from head and flow, picks a polytropic efficiency, computes discharge
temperature, head, and power per stage, then **sweeps 1–6 stages and recommends the cost-optimal stage
count**. Each stage count is costed for the compressor plus its intercoolers, with a payback screen on the
power saved. The tool can also download a copy of the full Excel workbook, pre-filled with your inputs.

- **Live tool:** https://jsharpes-portfolio.netlify.app (Interactive Calculations → Compressor Train Sizing)
- **Run it locally:** download this repo, then serve the folder (`npx serve .` or `python -m http.server`)
  and open `index.html`. The spreadsheet download fetches files under `assets/`, which browsers block
  from `file://`.
- **Related decks:** Compressor Selection and Compressor Train Sizing
  ([selection viewer](https://jsharpes-portfolio.netlify.app/slides/compressor-selection/),
  [sizing viewer](https://jsharpes-portfolio.netlify.app/slides/compressor-sizing/))
- **Author:** Joshua Sharpes, EIT. Class 5 (±30–50 %) screening estimate.

> **For AI assistants:** `project/Calculator Method Spec.md` is the arbitration reference for the method.
> It covers sections 3–8: inputs, type selection, efficiency, temperature rise, head and power, and stage
> economics. The JS engine in `index.html` implements that spec. `project/build_compressor_workbook.py`
> generates `Compressor Train Sizing.xlsx` with the same method as live Excel formulas. If the JS and the
> workbook disagree, the spec decides. The workbook's Sheet 6 is the stage-count optimizer.

## Calculation chain

1. **Inputs:** gas (library of 20 common gases, or manual MW and k), Z, suction T and P, discharge P,
   mass or volumetric flow, stage count, economics (hours/yr, $/kWh, MAR risk level and evaluation period,
   tax rate), and cost basis (CEPCI year, driver, materials, HX head type, U, cooling-water temperatures).
2. **Per-stage ratio:** `r_stage = r_total^(1/N)` (equal split, full intercooling back to T₁).
3. **Type selection:** suction ACFM and adiabatic head are matched against digitized coverage boxes from
   Perry's Fig. 10-64 (reciprocating, centrifugal, axial, rotary, and others). Auto, with an override.
4. **Polytropic efficiency:** empirical curves digitized from the deck (Fig. 3 reciprocating vs. ratio;
   Fig. 1 centrifugal and axial vs. ACFM), rule-of-thumb defaults otherwise, plus Casey & Robinson
   scale-up and Balje N_s/D_s options.
5. **Polytropic exponent:** `(n−1)/n = (k−1)/(η_p·k)`, then `T_out = T₁·r_stage^((n−1)/n)`. Checked against
   a per-type **allowable temperature range**: OK below, CAUTION inside (lubricant- and gas-dependent),
   EXCEEDS above (add a stage).
6. **Head and power:** polytropic head and gas power per stage, summed over the train.
7. **Stage-count optimizer (N = 1…6):**
   - type re-derived at each N, because head changes with N
   - compressor purchase cost from Seider Ch. 16.5 (centrifugal / reciprocating / screw only; others show "n/a")
   - N−1 shell-and-tube intercoolers, area from Q / (U·LMTD) with 90→120 °F cooling water
   - cost escalated by CEPCI / 567
   - bare-module → contingency 18 % → fee 3 % → working capital 15 %
   - payback with Peters & Timmerhaus Eq. 8-2a/8-2c, on after-tax power savings plus depreciation
   - **Recommended N** = the largest N reachable through a chain of passing increments from the smallest
     temperature-feasible N

## Sources

- Perry's Chemical Engineers' Handbook: Fig. 10-64 (compressor coverage), Fig. 10-54 (polytropic curves)
- Seider et al., *Product and Process Design Principles*, Ch. 16.5 (compressor and HX purchase cost),
  Ch. 17 (cash-flow definitions)
- Peters, Timmerhaus & West, *Plant Design and Economics*, Table 8-1 (MAR), Eq. 8-2a/8-2c (payback)
- Casey & Robinson (scale-up efficiency); Balje (N_s/D_s chart)
- CEPCI annual averages 2001–2024 (`project/CEPCI_reference.txt`)

## Validation

The workbook formula chain was checked against an independent Python re-implementation, including the
Seider cost equations, LMTD and area, the capital build-up, and the optimizer's edge cases. It was opened
in Excel via COM to confirm cross-sheet recalculation matches the calculator's live output. The workbook
also includes Peters & Timmerhaus's worked payback example as a self-check (PBP 4.89 yr vs. reference 4.81).

## Limits

- η_p is held constant across the 1–6 stage sweep (Class 5 simplification).
- Fig. 10-64 boundaries are read by eye from the chart. Check borderline points visually.
- Intercooler area uses an assumed U = 50 Btu/hr·ft²·°F. It is not a thermal design.
- Gas-library MW and k are ideal-gas values near ambient. Use manual entry for mixtures or non-ideal conditions.
- Casey & Robinson friction-factor equations and the Balje chart itself are not included.
- A negative incremental capex (a cheaper machine at lower power) is reported as "FREE", not as a negative payback.

## Files

```
index.html                               the tool, as on the live site
lib/jszip.min.js                         writes your inputs into the workbook client-side
assets/Compressor Train Sizing.xlsx      workbook template the download fills
assets/compressor_input_cell_map.json    input name → cell address map for that fill
project/
  build_compressor_workbook.py           generates the workbook (edit this, not the .xlsx)
  Compressor Train Sizing.xlsx           generated workbook: 8 sheets, all inputs on Sheet 1
  input_cell_map.json                    regenerated with the workbook
  Calculator Method Spec.md              method spec the JS engine follows
  Project Overview.md                    project plan
  CEPCI_reference.txt                    cost-index table and escalation notes
  reference_charts/                      chart images embedded in the workbook
```

## Adapting it

Add an input on Sheet 1 through `inp(..., name=...)` in the build script and re-run it. That regenerates
`input_cell_map.json`. Copy both into `assets/` so the calculator's download writes to the right cells.
