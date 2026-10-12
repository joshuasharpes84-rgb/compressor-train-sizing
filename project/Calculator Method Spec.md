# Compressor Train Web Calculator — Methodology & Build Handoff

**Audience**: an AI (or developer) building the Netlify web calculator described in
`Project Overview.md` and `CLAUDE.md` (Step 2 of this project). You do not need to open
`build_compressor_workbook.py` line-by-line to use this document — everything needed to
implement the engineering logic is below. That script remains the source of truth for
exact Excel cell addresses (see "Spreadsheet download" section) — read it only when you
get to that part.

**What this is not**: a UI mockup or a tech-stack mandate. It is the engineering
methodology (what to compute, in what order, from what inputs) plus the two hard
requirements the user stated:

1. The web calculator must be **iterative**: the user enters a small set of critical
   inputs and gets critical outputs back, live, including a recommendation to change one
   of those inputs (stage count) and see everything recompute — mirroring the
   spreadsheet's own Sheet 1 ⇄ Sheet 6 feedback loop.
2. Whatever the user typed into the calculator must **carry over into a downloadable
   copy of `Compressor Train Sizing.xlsx`** — the workbook then expands on those same
   inputs with the full multi-sheet detail (coverage charts, digitized efficiency
   curves, per-stage tables, Seider cost build-up) that a single-page calculator
   shouldn't try to replicate in the browser.

---

## 1. What the spreadsheet does (sheet-by-sheet summary)

The workbook analyzes a multi-stage compressor train with full intercooling back to
suction temperature and an equal pressure ratio per stage. It answers two separate
questions:

- **Sheets 1-5**: "given this many stages, what type, efficiency, discharge
  temperature, and power result?" — full detail at *one* chosen `N_stages`.
- **Sheet 6**: "how many stages *should* there be?" — sweeps N = 1..6, cost-optimizes
  independently of whatever N was typed into Sheet 1, and tells the user what to change
  Sheet 1 to.

The web calculator should reproduce both halves: a reactive single-N results panel
(≈ mirrors the 0. Overview sheet) plus a stage-count sweep table/recommendation
(≈ mirrors Sheet 6), with the two visibly reconciled ("your input says N=3; the optimizer
recommends N=4 — update?").

---

## 2. Input taxonomy — what to ask for, and how prominently

Every input in the spreadsheet lives on Sheet 1. Group them the same way for the
calculator's UI, in this priority order:

### 2a. Critical inputs (always visible, required, no defaults hidden)
| Input | Notes |
|---|---|
| Gas | Dropdown of the 20-gas library (§4) + "Custom / Manual Entry" |
| MW, k (manual) | Only shown/used if Gas = Custom |
| Z, compressibility factor | Default 1.00 — always a manual number, never gas-looked-up (Z is condition-dependent) |
| T1, suction temperature (°F) | e.g. 100 |
| P1, suction pressure (psia) | e.g. 100 |
| P2, final discharge pressure (psia) | e.g. 1000 |
| ṁ, mass flow rate (lb/hr) | e.g. 50,000 |
| N_stages | Integer 1-6. This is the one the optimizer (§7) may tell the user to change. |

### 2b. Auto-computed, override-able (shown, but pre-filled by the app — not blank)
| Field | Auto logic | Override |
|---|---|---|
| Compressor type | Box-match against the Fig. 10-64 coverage table (§5) using this N's Q/H_ad | Dropdown of the 10 coverage types, or "Auto" |
| Efficiency method | Empirical Correlation if type ∈ {Reciprocating, Centrifugal, Axial}, else Quick Default | Dropdown: Auto / Empirical Correlation / Quick Default / Scale-Up / Balje / Manual |
| Manual ηp | Only relevant if method = Manual | Number input, default 0.75 |
| Scale-up inputs (η1, Bref, fref, f1, f2) | Only relevant if method = Scale-Up | All default 0 — flag as "needs external friction-factor data" (§6b) |
| Balje inputs (N rpm, D ft, ηc) | Only relevant if method = Balje | Defaults 3600 rpm, 1.5 ft, ηc=0 — flag "chart is external, ηc must be read manually" |

### 2c. Economics (visible, sensible defaults — needed for every output, not "advanced")
| Input | Default |
|---|---|
| Operating hours/year | 8000 |
| Electricity cost ($/kWh) | 0.07 |
| Risk level (Table 8-1, P&T) | "Low - new capacity / established market" |
| N, evaluation period (yr) | 10 |

`mar` (minimum acceptable return) is **not** independently editable — it's the midpoint
of the selected risk band: `mar = AVERAGE(risk_min, risk_max)`. (The spreadsheet's own
label text calls it "editable," but the actual cell is a formula keyed off the risk
dropdown, not a free input — implement it as computed, matching the real behavior, not
the label.) Show the resolved `mar` value read-only next to the risk dropdown.

### 2d. Cost & Sizing Basis (collapsible "Advanced" section — real defaults, rarely changed)
| Input | Default |
|---|---|
| CEPCI (Year dropdown 2001-2024, or Custom/Manual) | Custom / Manual = 567 (Seider's own 2013 basis, i.e. no escalation) |
| Compressor driver | Electric Motor |
| Compressor material | Carbon Steel |
| Compressor bare-module factor | 2.15 |
| Intercooler head type | Floating Head |
| Intercooler material | Carbon steel/Carbon steel |
| Intercooler tube length (ft) | 20 |
| Overall U (Btu/hr·ft²·°F) | 50 |
| Cooling water in/out (°F) | 90 / 120 |
| Intercooler bare-module factor | 3.29 |
| Combined tax rate | 21% |

Every one of these has a real, citable default (Seider's own worked examples / current
US federal rate) — the section can start collapsed without the user feeling blocked.

---

## 3. Derived quantities (compute first, feed everything downstream)

```
T1_R        = T1_F + 459.67
R_prime     = 1545.349 / MW                      # ft·lbf/(lbm·°R)
r_total     = P2 / P1
r_stage(N)  = r_total ^ (1/N)                    # equal split across N stages
rho1        = (P1 * MW) / (10.7316 * Z * T1_R)   # lbm/ft³, R = 10.7316 psia·ft³/(lbmol·°R)
Q_ACFM      = (mdot / 60) / rho1                 # ft³/min at suction — CONSTANT across all N
H_ad(N)     = (k/(k-1)) * R_prime * T1_R * (r_stage(N)^((k-1)/k) - 1)   # ft, adiabatic head, n=k
```

`Q_ACFM` does not depend on N (only suction conditions/flow) — only `r_stage` and
`H_ad` change as N changes. This is why Sheet 6's sweep only has to vary two derived
numbers per row to re-run type selection and everything downstream.

---

## 4. Gas property library (reference data, not computed)

Approximate ideal-gas MW / k near ambient — flag in the UI that these are ballpark
(not flash/EOS results); "Custom / Manual Entry" should always be available and is the
default.

| Gas | MW | k |
|---|---|---|
| Air | 28.97 | 1.400 |
| Nitrogen (N2) | 28.01 | 1.400 |
| Oxygen (O2) | 32.00 | 1.395 |
| Hydrogen (H2) | 2.016 | 1.410 |
| Carbon Monoxide (CO) | 28.01 | 1.400 |
| Carbon Dioxide (CO2) | 44.01 | 1.289 |
| Methane (CH4) | 16.04 | 1.32 |
| Natural Gas (~0.60 SG) | 17.40 | 1.27 |
| Ethane (C2H6) | 30.07 | 1.19 |
| Ethylene (C2H4) | 28.05 | 1.24 |
| Propane (C3H8) | 44.10 | 1.13 |
| Propylene (C3H6) | 42.08 | 1.15 |
| n-Butane (C4H10) | 58.12 | 1.09 |
| Ammonia (NH3) | 17.03 | 1.31 |
| Hydrogen Sulfide (H2S) | 34.08 | 1.32 |
| Chlorine (Cl2) | 70.90 | 1.34 |
| Sulfur Dioxide (SO2) | 64.07 | 1.29 |
| Water Vapor / Steam (H2O) | 18.02 | 1.33 |
| Argon (Ar) | 39.95 | 1.667 |
| Helium (He) | 4.003 | 1.667 |

---

## 5. Compressor type selection (Fig. 10-64 coverage box-match)

Coverage table (`Q_min, Q_max` in ACFM; `H_min, H_max` in ft) — eyeballed from the
source chart, approximate by the workbook's own admission:

| Type | Q_min | Q_max | H_min | H_max |
|---|---|---|---|---|
| Diaphragm | 1 | 10 | 1,000 | 100,000 |
| Reciprocating | 1 | 3,000 | 500 | 100,000 |
| Rotary - Straight Lobe | 100 | 3,000 | 100 | 1,000 |
| Rotary - Screw | 10 | 3,000 | 100 | 10,000 |
| Rotary - Sliding Vane | 10 | 1,000 | 100 | 3,000 |
| Rotary - Liquid Piston | 1 | 300 | 1,000 | 10,000 |
| Centrifugal - Single Stage | 1,000 | 100,000 | 100 | 10,000 |
| Centrifugal - Multistage | 500 | 100,000 | 1,000 | 100,000 |
| Axial - Multistage | 3,000 | 100,000 | 100 | 3,000 |
| Fan / Blower | 1,000 | 100,000 | 10 | 1,000 |

**Match logic**: for a given `(Q, H_ad)`, a type matches if
`Q_min <= Q <= Q_max AND H_min <= H_ad <= H_max`. Zero, one, or several types can match.
Join all matches (comma-separated) as "candidates"; take the **first** match as the
"primary type" fed into everything downstream (Tallow lookup, Seider cost equation
routing). If **no** type matches, the effective type is the literal string `"None"` —
this must propagate as an unknown/blocking state (see §6d), never silently fall through
to a "safe" default.

This is exactly the logic Sheet 2 runs once (at the user's chosen N_stages) and Sheet 6
reruns per sweep row (§7) — implement it as one shared function, `suggestChoices(Q, H_ad)
-> string[]`, called from both places.

---

## 6. Efficiency (ηp)

### 6a. Empirical Correlation — three digitized curves, routed by type
Piecewise-linear interpolation, **clamped at both ends** (never extrapolate past the
table's first/last x-value):

**Reciprocating**, x = r_stage (pressure ratio):
| r_stage | 1.5 | 2.0 | 2.5 | 3.0 | 3.5 | 4.0 | 4.5 | 5.0 | 5.5 | 6.0 | 6.5 |
|---|---|---|---|---|---|---|---|---|---|---|---|
| ηp | 0.740 | 0.790 | 0.815 | 0.825 | 0.833 | 0.840 | 0.843 | 0.846 | 0.848 | 0.846 | 0.843 |

**Centrifugal**, x = Q_ACFM (suction):
| Q | 1,000 | 2,000 | 4,000 | 7,000 | 10,000 | 20,000 | 40,000 | 70,000 | 100,000 | 200,000 |
|---|---|---|---|---|---|---|---|---|---|---|
| ηp | 0.650 | 0.700 | 0.720 | 0.730 | 0.735 | 0.745 | 0.755 | 0.763 | 0.766 | 0.770 |

**Axial**, x = Q_ACFM (suction):
| Q | 75,000 | 100,000 | 200,000 | 400,000 | 600,000 |
|---|---|---|---|---|---|
| ηp | 0.813 | 0.817 | 0.827 | 0.832 | 0.836 |

Routing: if effective type contains "Reciprocating" → recip curve; "Centrifugal" →
centrifugal curve; "Axial" → axial curve; else → `"n/a - no digitized chart for this
type, use Quick Default"`.

### 6b. Quick Default — rule-of-thumb table, fallback for every other type
| Type | ηp |
|---|---|
| Reciprocating | 0.80 |
| Centrifugal - Single Stage | 0.76 |
| Centrifugal - Multistage | 0.75 |
| Axial - Multistage | 0.87 |
| Rotary - Screw | 0.72 |
| Rotary - Sliding Vane | 0.70 |
| Rotary - Straight Lobe | 0.68 |
| Rotary - Liquid Piston | 0.65 |
| Fan / Blower | 0.72 |

Lookup by exact type match; if the type isn't in this table (e.g. "None"), fall back to
0.75.

### 6c. Scale-Up (Casey & Robinson Eq. 17) — needs external data, not in the source decks
```
η2 = η1 - Bref * (f2 - f1) / fref        # 0 if fref = 0
```
Flag in the UI: "friction factors f1/f2 come from Casey & Robinson's Appendix 1
equations 23/24, not available in this project's source material — enter vendor/
external values, or use another method."

### 6d. Balje chart — Ns/Ds computed, ηc is a manual chart-reading
```
Ns = (N_rpm * 2π/60) * sqrt(Q_ACFM/60) / (32.174 * H_ad)^0.75
Ds = D_ft * (32.174 * H_ad)^0.25 / sqrt(Q_ACFM/60)
ηc = <manual entry — chart itself isn't in the source decks>
```

### 6e. Method resolution & final ηp
```
effective_method =
  if override != "Auto": override
  elif type ∈ {Reciprocating, Centrifugal, Axial}: "Empirical Correlation"
  else: "Quick Default"

selected_ηp =
  Empirical Correlation → §6a result
  Quick Default         → §6b result
  Scale-Up              → §6c result
  Balje                 → ηc from §6d
  Manual                → manual ηp input
```
`selected_ηp` is the single number everything from here down depends on.

---

## 7. Per-stage temperature, head, and power (at the user's chosen N_stages)

```
n = 1 / (1 - (k-1)/(k * selected_ηp))          # polytropic exponent, solved from
                                                # ηp = [n/(n-1)] / [k/(k-1)]

# per stage (identical for every stage under equal-split/full-intercool):
Tout_R  = T1_R * r_stage^((n-1)/n)
Tout_F  = Tout_R - 459.67

Hp_ft   = (n/(n-1)) * Z * R_prime * T1_R * (r_stage^((n-1)/n) - 1)   # ft (per stage)
Pgas_ft_lbf_per_hr = mdot * Hp_ft / selected_ηp
Pgas_hp = Pgas_ft_lbf_per_hr / 1,980,000        # 1,980,000 = 33,000 ft-lbf/min-hp * 60 min/hr

train_hp     = N_stages * Pgas_hp               # all stages equal, so a simple multiply
per_stage_hp = Pgas_hp
```

### Tallow (max allowable discharge temperature) — 3-tier flag, per stage
Ranges, not single cutoffs (true limit is lubricant/seal/flammability dependent):

| Type substring match | Min °F | Max °F | Why a range |
|---|---|---|---|
| "Reciprocating" | 300 | 350 | Flammable gases (e.g. H2) push the effective limit toward 300°F |
| "Centrifugal" | 392 | 446 | API 617 + industry practice |
| "Rotary" | 212 | 350 | Heavily oil-type dependent — some oils degrade as low as 212°F |

Flag logic, evaluated in this priority order:
```
if effective_type == "None":
    "UNKNOWN TYPE — cannot verify Tallow; treated as EXCEEDS until inspected manually"
elif type matches a row above AND Tout_F > max:
    "EXCEEDS {max}F — add a stage"
elif type matches a row above AND Tout_F > min:
    "CAUTION: {min}-{max}F is vendor/lubricant dependent — inspect the specific
     compressor rating (outside scope)"
elif type matches a row above:
    "OK"
else:   # Axial, Fan/Blower, Diaphragm, other Rotary sub-types not in the 3 rows
    "n/a (Axial/Fan — no Tallow range given in the deck)"
```
Critical: `"None"` (no coverage-box match at all) must hit the **first** branch, not
fall through to the `"n/a"` branch — an unknown type is a blocking state, not a free
pass. This is a bug the spreadsheet build hit once and fixed explicitly; don't
reintroduce it.

---

## 8. Stage-Count Optimizer (Sheet 6 sweep, N = 1..6) — the second half of the calculator

This is independent of the user's chosen N_stages above — it answers "what N should
this actually be." Implement as `sweepStageCounts(inputs) -> row[6]`, then a reconciler
against the user's N_stages.

**Simplification carried over deliberately**: `selected_ηp` is held constant across
the whole sweep (computed once, at the user's chosen N_stages) rather than re-derived
per row. This is a documented Class-5 (±30-50%) screening approximation — surface it in
the UI near the sweep table, don't silently hide it.

For each `N` in 1..6:
```
r_stage(N) = r_total ^ (1/N)
H_ad(N)    = (k/(k-1)) * R_prime * T1_R * (r_stage(N)^((k-1)/k) - 1)
candidates(N) = suggestChoices(Q_ACFM, H_ad(N))        # §5 — Q constant, only H_ad varies
primary_type(N) = candidates(N)[0]  (or "None")

Tout_F(N) = T1_R * r_stage(N)^((n-1)/n) - 459.67        # same n for every row (ηp constant)
temp_flag(N) = §7's Tallow logic, using primary_type(N)

Pc(N) = N * mdot * Hp_ft(r_stage(N)) / selected_ηp / 1,980,000   # train power at this N,
                                                                   # treated directly as
                                                                   # Seider's "consumed power"
```

### Compressor cost (Seider, Seader, Lewin & Widagdo Ch. 16.5)
One machine, sized at `Pc(N)` (hp), base cost at CE=567:
```
Cb_compressor =
  Centrifugal:    exp(9.1553 + 0.63   * ln(Pc))
  Reciprocating:  exp(4.6762 + 1.23   * ln(Pc))
  Screw:          exp(8.2496 + 0.7243 * ln(Pc))
  other types:    "n/a - outside Seider Ch.16.5 correlation coverage"

FD (driver factor) = Electric Motor: 1.00 | Steam Turbine: 1.15 | Gas Turbine: 1.25
FM (material)       = Carbon Steel: 1.00 | Stainless Steel: 2.5 | Nickel Alloy: 5.0

Cp_compressor  = Cb_compressor * FD * FM * (CEPCI_current / 567)
BMC_compressor = Cp_compressor * BMF_compressor          # default 2.15
```

### Intercooler cost — N-1 shell-and-tube units
```
coolers_qty(N)   = max(N - 1, 0)
Cp_gas_Btu       = (k/(k-1)) * R_prime / 778.169          # Btu/(lbm·°F)
cooling_duty(N)  = mdot * Cp_gas_Btu * (Tout_R(N) - T1_R)  # Btu/hr, per unit
design_psig(N)   = P1 * r_stage(N) - 14.696                # interstage gauge pressure

# Counter-current LMTD: gas Tout(hot in) -> T1(cold out); CW in=90F meets gas cold end,
# CW out=120F meets gas hot end
dt_hot  = Tout_F(N) - CW_out
dt_cold = T1_F      - CW_in
LMTD(N) = dt_hot                              if dt_hot == dt_cold
        = (dt_hot - dt_cold) / ln(dt_hot/dt_cold)   otherwise
                                               # if LMTD <= 0: "n/a - check CW temps"

Area(N) = cooling_duty(N) / (U * LMTD(N))      # ft², U default 50 Btu/(hr·ft²·°F)

Cb_HX(A) =
  Floating Head:    exp(12.0310 - 0.8709*ln(A) + 0.09005*ln(A)^2)
  Fixed Head:       exp(11.4185 - 0.9228*ln(A) + 0.09861*ln(A)^2)
  U-Tube:           exp(11.5510 - 0.9186*ln(A) + 0.09790*ln(A)^2)
  Kettle Vaporizer: exp(12.3310 - 0.8709*ln(A) + 0.09005*ln(A)^2)

FP(N) = 0.9803 + 0.018*(design_psig(N)/100) + 0.0017*(design_psig(N)/100)^2

FM_HX(A) = a + (A/100)^b, where (a, b) by shell/tube material:
  Carbon steel/Carbon steel:      (0,    0)     -> FM_HX = 1
  Carbon steel/Stainless steel:   (1.75, 0.13)
  Stainless steel/Stainless steel:(2.70, 0.07)

FL (tube length) = 8ft: 1.25 | 12ft: 1.12 | 16ft: 1.05 | 20ft: 1.00

Cp_HX_per_unit  = Cb_HX * FP(N) * FM_HX * FL * (CEPCI_current / 567)
BMC_HX_per_unit = Cp_HX_per_unit * BMF_intercooler        # default 3.29
BMC_HX_total(N) = coolers_qty(N) * BMC_HX_per_unit
```

### Build-up to installed cost (Guthrie/Seider module costing)
```
TBM(N)              = BMC_compressor(N) + BMC_HX_total(N)     # "n/a" if either side is n/a
Contingency(N)       = 0.18 * TBM(N)
Fee(N)               = 0.03 * TBM(N)
Total_Module_Cost(N) = TBM(N) + Contingency(N) + Fee(N)
Working_Capital(N)   = 0.15 * Total_Module_Cost(N)
TCI(N)               = Total_Module_Cost(N) + Working_Capital(N)   # "TOTAL INSTALLED COST"
```

### Payback screen (P&T structure, Seider after-tax cash flow) — incremental N-1 → N
```
incr_capex(N)   = TCI(N) - TCI(N-1)
pretax_savings(N) = (Pc(N-1) - Pc(N)) * 0.7457 * hours_per_year * elec_cost_per_kWh
after_tax_CF(N) = (1 - tax_rate) * pretax_savings(N) + incr_capex(N) / N_eval_years
                  # depreciation life reuses N_eval since none was separately specified

incr_PBP(N) =
  if incr_capex(N) <= 0: "FREE - cheaper capital AND lower operating cost"   # auto-pass
  elif after_tax_CF(N) <= 0: "n/a"                                          # auto-fail
  else: incr_capex(N) / after_tax_CF(N)

PBP_ref = 0.85 / (mar + 0.85 / N_eval_years)          # Eq. 8-2c

passes_econ(N) =
  "FREE..." -> TRUE
  "n/a"     -> FALSE
  numeric   -> (incr_PBP(N) <= PBP_ref) AND (incr_PBP(N) <= 5)
```
**Edge case to preserve**: Seider's steep power exponents (1.23 reciprocating, 0.63
centrifugal) mean a *lower*-power train from adding a stage can make the compressor
itself cheaper even after adding an intercooler — `incr_capex` can be negative. Don't
compute `incr_capex / after_tax_CF` blindly in that case (it produces a nonsensical
negative "payback period" that happens to pass); the `"FREE"` branch above exists
specifically to special-case it.

### Recommended optimal N
```
min_feasible_N = smallest N in 1..6 whose temp_flag(N) is NOT "EXCEEDS..."
                 (OK, CAUTION, and the axial/fan "n/a - no range given" all count as
                  feasible; only EXCEEDS blocks)
                 -> if none, "NONE - need more than 6 stages"

chain(N) =
  N < min_feasible_N        -> FALSE
  N == min_feasible_N       -> TRUE   (auto-passing baseline — never evaluate its own
                                        incr_PBP, which is undefined against a
                                        possibly-"n/a" N-1 row)
  N > min_feasible_N        -> chain(N-1) AND passes_econ(N)

recommended_optimal_N = largest N (1..6) where chain(N) == TRUE
                         -> if none pass, "infeasible within 6 stages"

total_project_cost = TCI(recommended_optimal_N)
suggested_type_at_optimum = primary_type(recommended_optimal_N)
```

---

## 9. Iterative calculator UX — the reconciliation loop

This is the part that makes it "iterative," not just a form:

1. User fills in §2a critical inputs (+ §2c economics; §2b/§2d default sensibly).
2. App live-computes §3 (Q, H_ad), auto-suggests type + method (§5, §6e) inline —
   editable, not just displayed.
3. App shows the §7 results at the user's current `N_stages`: ηp, n, Tout, Tallow flag,
   train hp — this is the "critical outputs" panel (≈ Overview sheet).
4. App also runs the full §8 sweep in the background and shows a compact
   **"Recommended optimal N: X — total project cost $Y"** callout, exactly mirroring the
   spreadsheet's `COMBINED RECOMMENDATION` cell:
   - If `recommended_optimal_N == N_stages`: "Your stage count already matches the
     economic optimum."
   - Else: "Optimum is N stage(s); you're currently set to N_stages — [Update to N]"
     with a one-click button that sets `N_stages = recommended_optimal_N` and
     recomputes step 3 — this closes the loop the spreadsheet expects the user to do
     manually (change Sheet 1, watch Sheets 2-5 re-derive).
5. Expandable detail view of the §8 sweep table (one row per N, same columns as Sheet 6)
   for users who want to see why a given N was rejected (Tallow EXCEEDS, or PBP fails).
6. A persistent "Download full spreadsheet" action (see §10) available at every step,
   always reflecting the *current* state of every input/override on screen.

Recompute everything client-side on every input change — there's no reason to round-trip
to a server for this; the whole model is closed-form (interpolation + algebra), no
solver/iteration needed beyond simple formulas.

---

## 10. Carrying inputs into the downloadable spreadsheet

**Requirement**: whatever the user has entered/overridden in the calculator must land
in the *same* Sheet 1 cells `build_compressor_workbook.py` writes as `inp()` calls, so
that when the user opens the downloaded `.xlsx`, every formula on Sheets 2-6 recalculates
from their actual inputs — not the workbook's built-in defaults.

**Recommended architecture**:

1. **Template, not regeneration.** Bundle the already-built `Compressor Train
   Sizing.xlsx` (or a copy of it) as a static asset with the Netlify Function. The
   function opens it, overwrites only the specific input cells (openpyxl in a Python
   function, or `exceljs`/similar in a Node function), and streams the result back —
   it does not re-run `build_compressor_workbook.py`'s cell-layout logic. This is
   simpler and can't accidentally regenerate a differently-laid-out sheet.
2. **Single source of truth for cell addresses.** `build_compressor_workbook.py`
   currently computes every input cell's row via a running `r` counter or a named
   `*_row` variable (e.g. `su_bref_row`, `hx_type_row`, `tax_row` — see that script's
   "Input Architecture" section in this project's `CLAUDE.md` for why row numbers
   aren't hardcoded). **Recommend adding one small change to that script**: right before
   `wb.save(OUT_PATH)`, dump a JSON sidecar (e.g. `input_cell_map.json`) of
   `{logical_input_name: "C15", ...}` for every `inp()` call made. Both the calculator's
   JS engine and the Netlify Function's cell-writer then import that same JSON instead
   of hardcoding cell addresses a second time — eliminating the exact "row drift"
   failure mode this project has already hit twice (documented in `CLAUDE.md`'s closing
   "Conventions" section).
3. **Two implementations of the math will necessarily coexist** — the browser's live JS
   engine (for reactive UI, per §3/§6/§7/§8 above) and the Excel formulas already baked
   into the template (for the downloaded file). They are not the same code and cannot
   be merged (Excel needs real formulas, not precomputed values, so the download stays
   editable/traceable after the user leaves the site). **Treat this markdown document
   as the arbitration reference**: if the engineering method ever changes, update
   `build_compressor_workbook.py`'s formulas, the JS engine, and this document together;
   don't patch one and assume the others still match.
4. Auto-computed values that came from an "Auto" state in the calculator (type,
   method) should still be written as `"Auto"` into the override cell (not the resolved
   value) — so the downloaded spreadsheet keeps behaving exactly like the live
   calculator (re-deriving from Q/H_ad) rather than silently freezing a stale
   suggestion. Only write the resolved value if the user explicitly overrode it in the
   UI.
5. File name should communicate it's a personalized export, e.g.
   `Compressor Train Sizing - <gas>-<N>stage-<timestamp>.xlsx`.

---

## 11. Known simplifications / gaps to surface in the UI (don't silently omit)

Carry these disclaimers into the web calculator's copy — they're already flagged inline
in the spreadsheet and should stay visible, not get lost in translation:

- Fig. 10-64 type-selection boundaries are eyeballed from a chart image, not precisely
  digitized — borderline operating points should note "cross-check against the
  reference chart."
- Gas library MW/k are approximate ideal-gas values near ambient — mixtures/non-ideal
  conditions should use Custom/Manual entry.
- Scale-Up method needs external Casey & Robinson friction-factor data not present in
  this project's source decks.
- Balje method's chart itself isn't available — ηc must come from an external chart
  reading.
- Seider Ch. 16.5 only gives compressor cost equations for Centrifugal, Reciprocating,
  Screw — every other type must show "n/a - outside Seider correlation coverage," not a
  fabricated fit.
- CEPCI defaults to 567 (no escalation) unless the user picks a real year.
- Combined tax rate defaults to 21% flat federal (no state) — user-editable for other
  jurisdictions.
- Depreciation life for the after-tax cash flow reuses the evaluation period N, not a
  separately specified MACRS schedule.
- The Stage-Count Optimizer holds ηp constant across the N=1..6 sweep — a documented
  Class-5 (±30-50%) screening simplification, not per-row-accurate.
- Intercooler sizing uses an assumed overall U and a standard 90/120°F cooling-water
  basis to back into an area for costing — not a full thermal/mechanical HX design.

---

## 12. Validation checklist before calling the port "done"

- [ ] Reference example on Sheet 6 (ΔTCI=$631,880, ΔCOM=$129,163/yr, mar=12%, N=15yr)
      reproduces PBP=4.89 vs PBPref=4.81 (borderline reject) — run this through the new
      JS engine as a standalone unit test, independent of the sweep logic.
  - [Reference example is not derived from user inputs — it's from the annotated PBP
    PDF and exists purely as a spot-check.]
- [ ] For a handful of representative input sets, compare the web calculator's Sheet
      1-5-equivalent outputs (type, ηp, n, Tout, flag, hp) against the actual
      `Compressor Train Sizing.xlsx` opened in real Excel with the same inputs typed
      into Sheet 1.
- [ ] Confirm the "None" (no coverage match) and "EXCEEDS" Tallow paths are reachable
      and render distinctly from "OK"/"CAUTION" in the UI — don't let an unknown type
      silently read as safe.
- [ ] Confirm a negative `incr_capex` case renders "FREE - cheaper capital AND lower
      operating cost," not a nonsensical negative payback period.
- [ ] Download a spreadsheet from the calculator with non-default inputs, open it in
      Excel, and confirm Sheets 2-6 recalculate to the same numbers the web calculator
      showed.
