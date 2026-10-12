"""
Generates Compressor Train Sizing.xlsx from the Compressor Selection / Compressor
Sizing decks + PT_payback_period_p345.pdf. Re-run this script after editing to
regenerate the workbook (the .xlsx itself is a build artifact, not hand-edited).

Workflow implemented (mirrors "Sizing Workflow", Compressor Sizing.pptx slide 4):
  0. Overview           - read-only rollup of every sheet (protected, no edits)
  1. Inputs             - EVERY user input lives here (gas props, stage count,
                          economics, cost-scaling reference points). Compressor
                          type and efficiency method are NOT manual inputs — both
                          are auto-computed from adiabatic head/type (an "Auto"
                          override dropdown exists for each, for one-off cases).
                          All other sheets are read-only / derived from this sheet.
  2. Type Selection     - capacity/head vs Fig. 10-64 (Perry's) coverage chart; the
                          matching coverage-table row(s) are highlighted (conditional
                          formatting) rather than charted — a live ScatterChart was
                          tried and dropped as unreliable/limited for a log-log
                          rectangular region
  3. Efficiency         - empirical correlation / scale-up / Balje methods
  4. Temp Rise & n      - polytropic n from eta_p, Tout vs Tallow per stage
  5. Head & Power       - polytropic head/power per stage and train totals (US units)
  6. Stage Economics    - Stage-Count Optimizer: sweeps N=1..6 in one wide table (one
                          row per N, every cost constituent its own column), costing
                          the compressor + intercoolers actually implied at each N via
                          Seider, Seader, Lewin & Widagdo Ch. 16.5's explicit purchase-
                          cost equations (compressor by consumed power, intercooler by
                          area from an assumed U/LMTD), then Guthrie module costing per
                          ../../../FEL Projects/Economics/notes/01_capital_cost_estimation.md.
                          Compressor type is re-suggested per row (head drops non-
                          linearly with more stages, so a different N can land in a
                          different Fig. 10-64 box). The payback screen keeps Peters &
                          Timmerhaus's Eq. 8-2a/8-2c structure but with an after-tax,
                          depreciation-added-back cash flow per Seider Eq. 17.8 — see
                          compressor_cost_comparison_PT_vs_Seider.pdf for why each
                          textbook was used where it was. Recommends the largest N
                          whose next stage still pays for itself, subject to the
                          Tallow temperature constraint, and reports total project cost.

PURPOSE, for anyone re-deriving this: Sheets 1-5 analyze ONE chosen stage count
(N_stages on Sheet 1) in full detail — physics, efficiency, temperature, power.
Sheet 6 independently answers "how many stages SHOULD there be" by sweeping
N=1..6 and finding the economically optimal count (not just screening the one
N you happened to pick). If Sheet 6's recommendation differs from Sheet 1's
N_stages, update N_stages to match and Sheets 2-5 re-derive automatically.

All engineering units are US customary (psia, °F/°R, ft, ACFM, hp, Btu/hr) to match
the source decks' charts and tables. $/kWh is kept for electricity billing since that
is the standard utility billing unit even in the US, not an SI holdover.
"""
import os
import json
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter

BASE = os.path.dirname(os.path.abspath(__file__))
CHARTS = os.path.join(BASE, "reference_charts")
OUT_PATH = os.path.join(BASE, "Compressor Train Sizing.xlsx")
CELL_MAP_PATH = os.path.join(BASE, "input_cell_map.json")

# Populated by inp(..., name=...) below — one entry per logical input the Netlify
# calculator needs to write into, so cell addresses are never hardcoded a second time
# outside this script (see Netlify_Calculator_Handoff.md §10.2). Dumped to
# input_cell_map.json right before wb.save().
INPUT_CELL_MAP = {}

# --- Styling helpers ---------------------------------------------------
TITLE_FILL = PatternFill("solid", fgColor="1F4E78")
SECTION_FILL = PatternFill("solid", fgColor="D9E1F2")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
CALC_FILL = PatternFill("solid", fgColor="F2F2F2")
FLAG_FILL = PatternFill("solid", fgColor="FCE4D6")
OVERVIEW_FILL = PatternFill("solid", fgColor="E2EFDA")
REQUIRED_FILL = PatternFill("solid", fgColor="F8CBAD")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def title_row(ws, cell, text, span=6):
    ws.merge_cells(f"{cell}:{get_column_letter(ws[cell].column + span - 1)}{ws[cell].row}")
    c = ws[cell]
    c.value = text
    c.font = Font(bold=True, size=14, color="FFFFFF")
    c.fill = TITLE_FILL
    c.alignment = Alignment(vertical="center")
    ws.row_dimensions[ws[cell].row].height = 24


def section(ws, cell, text, span=6):
    ws.merge_cells(f"{cell}:{get_column_letter(ws[cell].column + span - 1)}{ws[cell].row}")
    c = ws[cell]
    c.value = text
    c.font = Font(bold=True, size=11)
    c.fill = SECTION_FILL
    c.border = BORDER


def label(ws, cell, text, bold=False, italic=False, wrap=False, size=10):
    c = ws[cell]
    c.value = text
    c.font = Font(bold=bold, italic=italic, size=size)
    if wrap:
        c.alignment = Alignment(wrap_text=True, vertical="top")


def inp(ws, cell, value, fmt=None, required=False, name=None):
    c = ws[cell]
    c.value = value
    c.fill = REQUIRED_FILL if required else INPUT_FILL
    c.border = BORDER
    if fmt:
        c.number_format = fmt
    if name:
        INPUT_CELL_MAP[name] = f"{ws.title}!{cell}"
    return c


def calc(ws, cell, formula, fmt=None):
    c = ws[cell]
    c.value = formula
    c.fill = CALC_FILL
    c.border = BORDER
    c.font = Font(italic=True)
    if fmt:
        c.number_format = fmt
    return c


def note(ws, cell, text, span=6, size=9):
    ws.merge_cells(f"{cell}:{get_column_letter(ws[cell].column + span - 1)}{ws[cell].row}")
    c = ws[cell]
    c.value = text
    c.font = Font(italic=True, size=size, color="595959")
    c.alignment = Alignment(wrap_text=True, vertical="top")


def hyperlink_note(ws, cell, display_text, target_sheet, target_cell, span=2):
    ws.merge_cells(f"{cell}:{get_column_letter(ws[cell].column + span - 1)}{ws[cell].row}")
    c = ws[cell]
    c.value = f'=HYPERLINK("#\'{target_sheet}\'!{target_cell}","{display_text}")'
    c.font = Font(italic=True, size=9, color="0563C1", underline="single")


def set_widths(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def tallow_flag_formula(row, type_ref, n_stages_ref, tallow, h_col="H"):
    """3-tier flag: OK / CAUTION (within the vendor-dependent range) / EXCEEDS.
    Tallow entries are ranges, not single numbers, because the true limit depends on
    specifics (lubricant/oil type, flammable-gas content, seal material) that this
    spreadsheet does not model — see Compressor Sizing.pptx slide 8.
    type_ref may resolve to the literal text "None" (Sheet 6's per-row auto-suggestion when
    no coverage box matches at all, e.g. an extreme operating point) — that must NOT fall
    through to the axial/fan "no range given" branch, which is silently treated as feasible.
    Unknown-type is the opposite of "no constraint": treat it as blocking until inspected."""
    current = '"n/a (Axial/Fan - no Tallow range given in the deck)"'
    for name, tmin, tmax, _ in reversed(tallow):
        current = (
            f'IF(ISNUMBER(SEARCH("{name}",{type_ref})),'
            f'IF({h_col}{row}>{tmax},"EXCEEDS {tmax}F - add a stage",'
            f'IF({h_col}{row}>{tmin},'
            f'"CAUTION: {tmin}-{tmax}F is vendor/lubricant dependent - inspect the specific '
            f'compressor rating (outside scope of this spreadsheet)","OK")),'
            f'{current})'
        )
    unknown_guard = (
        f'IF({type_ref}="None",'
        f'"UNKNOWN TYPE - no compressor in the coverage table fits these conditions; '
        f'cannot verify Tallow, treated as EXCEEDS until inspected manually",'
        f'{current})'
    )
    return f'=IF(B{row}>{n_stages_ref},"",{unknown_guard})'


MATCH_FILL = PatternFill("solid", fgColor="C6EFCE")  # standard Excel "good" green


def highlight_matching_rows(ws, cell_range, first_row, q_ref, h_ref, qmin_col, qmax_col, hmin_col, hmax_col,
                             fill=MATCH_FILL):
    """Conditional formatting instead of a chart: Excel's native charting has no clean way to
    shade a log-log rectangular region, and an earlier attempt at a live ScatterChart proved
    unreliable in practice. Highlighting the matching row(s) of the coverage-range table is
    simpler and just as informative — formula is written relative to first_row and Excel
    auto-adjusts it for every row in cell_range, exactly like a normal CF rule."""
    formula = (f"AND({q_ref}>={qmin_col}{first_row},{q_ref}<={qmax_col}{first_row},"
               f"{h_ref}>={hmin_col}{first_row},{h_ref}<={hmax_col}{first_row})")
    ws.conditional_formatting.add(cell_range, FormulaRule(formula=[formula], fill=fill))


def suggestion_formula(q_ref, h_ref, cov_sheet_ref, first_row, last_row):
    """Same box-match logic as the Sheet-2 auto-suggestion, generalized to any Q/H cell
    reference and any coverage-table location (used for the per-N sweep on Sheet 6, where
    H_ad varies row to row while Q stays constant)."""
    return "&".join(
        f'IF(AND({q_ref}>={cov_sheet_ref}C{i},{q_ref}<={cov_sheet_ref}D{i},'
        f'{h_ref}>={cov_sheet_ref}E{i},{h_ref}<={cov_sheet_ref}F{i}),{cov_sheet_ref}B{i}&", ","")'
        for i in range(first_row, last_row + 1)
    )


def seider_compressor_cb_formula(type_expr, pc_ref):
    """Seider, Seader, Lewin & Widagdo Ch. 16.5 base f.o.b. purchase cost at CE=567 (2013),
    Pc = consumed power in Hp: Eq. 16.36 (centrifugal), 16.37 (reciprocating), 16.38 (screw).
    No equation is given in the source for diaphragm/other-rotary/axial/fan-blower types —
    those fall through to a plain-text "not covered" flag rather than a fabricated fit."""
    return (
        f'IF(ISNUMBER(SEARCH("Centrifugal",{type_expr})),EXP(9.1553+0.63*LN({pc_ref})),'
        f'IF(ISNUMBER(SEARCH("Reciprocating",{type_expr})),EXP(4.6762+1.23*LN({pc_ref})),'
        f'IF(ISNUMBER(SEARCH("Screw",{type_expr})),EXP(8.2496+0.7243*LN({pc_ref})),'
        f'"n/a - outside Seider Ch.16.5 correlation coverage")))'
    )


def seider_hx_cb_formula(hxtype_ref, a_ref):
    """Seider shell-and-tube base cost at CE=567: Eq. 16.39 (floating head), 16.40 (fixed
    head), 16.41 (U-tube), 16.42 (kettle vaporizer); A = tube outside surface area, ft²."""
    return (
        f'IF({hxtype_ref}="Floating Head",EXP(12.0310-0.8709*LN({a_ref})+0.09005*LN({a_ref})^2),'
        f'IF({hxtype_ref}="Fixed Head",EXP(11.4185-0.9228*LN({a_ref})+0.09861*LN({a_ref})^2),'
        f'IF({hxtype_ref}="U-Tube",EXP(11.5510-0.9186*LN({a_ref})+0.09790*LN({a_ref})^2),'
        f'EXP(12.3310-0.8709*LN({a_ref})+0.09005*LN({a_ref})^2))))'
    )


def hx_fm_formula(hxmat_ref, a_ref):
    """FM = a + (A/100)^b, Seider Table 16.25 / Eq. 16.44 — three common shell/tube combos."""
    a_lookup = (f'IF({hxmat_ref}="Carbon steel/Carbon steel",0,'
                f'IF({hxmat_ref}="Carbon steel/Stainless steel",1.75,2.70))')
    b_lookup = (f'IF({hxmat_ref}="Carbon steel/Carbon steel",0,'
                f'IF({hxmat_ref}="Carbon steel/Stainless steel",0.13,0.07))')
    return f'(({a_lookup})+({a_ref}/100)^({b_lookup}))'


def lmtd_formula(tout_gas_ref, t1_gas_ref, cw_in_ref, cw_out_ref):
    """Counter-current LMTD: gas cools from Tout (hot inlet) to T1 (cold outlet, i.e. back to
    suction temp via the intercooler); cooling water enters at cw_in (coldest, meets the gas's
    cold end) and exits at cw_out (meets the gas's hot end)."""
    dt_hot = f'({tout_gas_ref}-{cw_out_ref})'
    dt_cold = f'({t1_gas_ref}-{cw_in_ref})'
    return f'IF({dt_hot}={dt_cold},{dt_hot},({dt_hot}-{dt_cold})/LN({dt_hot}/{dt_cold}))'


def interp_formula(x_ref, x_first_row, x_last_row, x_col="B", y_col="C"):
    """Piecewise-linear interpolation against an ascending x-vs-y table on the current
    sheet, clamped at both ends. Returns the bare expression (no leading '='), so it can
    be embedded inside a larger IF chain (e.g. routing by compressor type) as well as
    used standalone. Shared by all three digitized efficiency curves on Sheet 3."""
    match_rng = f"${x_col}${x_first_row}:${x_col}${x_last_row}"
    y_rng = f"${y_col}${x_first_row}:${y_col}${x_last_row}"
    return (
        f"IF({x_ref}<={x_col}{x_first_row},{y_col}{x_first_row},"
        f"IF({x_ref}>={x_col}{x_last_row},{y_col}{x_last_row},"
        f"INDEX({y_rng},MATCH({x_ref},{match_rng},1))+"
        f"(INDEX({y_rng},MATCH({x_ref},{match_rng},1)+1)-INDEX({y_rng},MATCH({x_ref},{match_rng},1)))*"
        f"({x_ref}-INDEX({match_rng},MATCH({x_ref},{match_rng},1)))/"
        f"(INDEX({match_rng},MATCH({x_ref},{match_rng},1)+1)-INDEX({match_rng},MATCH({x_ref},{match_rng},1))))"
        f")"
    )


wb = Workbook()

# =========================================================================
# 1. Inputs — every user-editable value in the workbook lives on this sheet.
# Sheets 2-6 read from here; none of them hold their own input cells anymore.
# =========================================================================
ws1 = wb.active
ws1.title = "1. Inputs"
ws = ws1
set_widths(ws, [3, 34, 16, 46, 3, 32, 14, 12])
title_row(ws, "B2", "1. Process & Economic Inputs", span=3)
note(ws, "B3", "Every input in the workbook is on this sheet. Sheets 2-6 are read-only / derived — "
              "change a value here and it carries through everywhere it's used.", span=3)

section(ws, "B5", "Gas Properties", span=3)
label(ws, "B6", "Molecular weight, MW (lb/lbmol) — effective", bold=True)
calc(ws, "C6", '=IF($G$6="Custom / Manual Entry",$G$7,VLOOKUP($G$6,$F$11:$H$31,2,FALSE))', "0.00")
ws["C6"].fill = FLAG_FILL
label(ws, "B7", "Heat capacity ratio, k = Cp/Cv — effective", bold=True)
calc(ws, "C7", '=IF($G$6="Custom / Manual Entry",$G$8,VLOOKUP($G$6,$F$11:$H$31,3,FALSE))', "0.00")
ws["C7"].fill = FLAG_FILL
label(ws, "B8", "Compressibility factor, Z (avg.)")
inp(ws, "C8", 1.00, "0.00", name="Z")
note(ws, "D6", "k is easier to compress at higher MW (fewer degrees of freedom) — "
              "Compressor Selection.pptx slide 5.", span=1)
note(ws, "D7", "Z is process-condition dependent (not gas-identity alone), so it stays a manual "
              "input regardless of Gas selection — see the library at right for MW/k only.", span=1)

# --- Gas library (columns F:H) — select a common gas, or leave "Custom / Manual
# Entry" (the default) to type MW/k directly, exactly as before this feature existed. ---
label(ws, "F6", "Gas (choose from library at right, or \"Custom / Manual Entry\")")
ws.merge_cells("G6:H6")
inp(ws, "G6", "Custom / Manual Entry", name="gas")
gas_table = [
    ("Air", 28.97, 1.400),
    ("Nitrogen (N2)", 28.01, 1.400),
    ("Oxygen (O2)", 32.00, 1.395),
    ("Hydrogen (H2)", 2.016, 1.410),
    ("Carbon Monoxide (CO)", 28.01, 1.400),
    ("Carbon Dioxide (CO2)", 44.01, 1.289),
    ("Methane (CH4)", 16.04, 1.32),
    ("Natural Gas (~0.60 SG)", 17.40, 1.27),
    ("Ethane (C2H6)", 30.07, 1.19),
    ("Ethylene (C2H4)", 28.05, 1.24),
    ("Propane (C3H8)", 44.10, 1.13),
    ("Propylene (C3H6)", 42.08, 1.15),
    ("n-Butane (C4H10)", 58.12, 1.09),
    ("Ammonia (NH3)", 17.03, 1.31),
    ("Hydrogen Sulfide (H2S)", 34.08, 1.32),
    ("Chlorine (Cl2)", 70.90, 1.34),
    ("Sulfur Dioxide (SO2)", 64.07, 1.29),
    ("Water Vapor / Steam (H2O)", 18.02, 1.33),
    ("Argon (Ar)", 39.95, 1.667),
    ("Helium (He)", 4.003, 1.667),
]
dv_gas = DataValidation(type="list", formula1="$F$11:$F$31", allow_blank=False)
ws.add_data_validation(dv_gas)
dv_gas.add(ws["G6"])
label(ws, "F7", "Manual MW (lb/lbmol) — used only if Gas = \"Custom / Manual Entry\"")
inp(ws, "G7", 20.0, "0.00", name="mw_manual")
label(ws, "F8", "Manual k — used only if Gas = \"Custom / Manual Entry\"")
inp(ws, "G8", 1.30, "0.00", name="k_manual")
note(ws, "F9", "MW/k below are approximate ideal-gas values near ambient conditions (ballpark "
              "reference only, not a flash/EOS result) — for mixtures or non-ideal conditions, use "
              "Custom / Manual Entry with your own computed values.", span=2)
label(ws, "F10", "Gas", bold=True)
label(ws, "G10", "MW (lb/lbmol)", bold=True)
label(ws, "H10", "k = Cp/Cv", bold=True)
for c in ("F10", "G10", "H10"):
    ws[c].fill = SECTION_FILL
    ws[c].border = BORDER
label(ws, "F11", "Custom / Manual Entry")
gas_row0 = 12
for i, (name, mw, k) in enumerate(gas_table):
    row = gas_row0 + i
    label(ws, f"F{row}", name)
    gc = ws[f"G{row}"]
    gc.value = mw
    gc.number_format = "0.00"
    gc.border = BORDER
    hc = ws[f"H{row}"]
    hc.value = k
    hc.number_format = "0.000"
    hc.border = BORDER

section(ws, "B10", "Suction / Discharge Conditions", span=3)
label(ws, "B11", "Suction temperature, T1 (°F)")
inp(ws, "C11", 100.0, "0.0", name="T1_F")
label(ws, "B12", "Suction pressure, P1 (psia)")
inp(ws, "C12", 100.0, "0.0", name="P1_psia")
label(ws, "B13", "Final discharge pressure, P2 (psia)")
inp(ws, "C13", 1000.0, "0.0", name="P2_psia")
label(ws, "B14", "Mass flow rate, mdot (lb/hr)")
inp(ws, "C14", 50000.0, "#,##0", name="mdot_lbhr")
label(ws, "B15", "Number of stages to analyze in detail, N_stages (drives Sheets 2-5)")
inp(ws, "C15", 3, "0", name="N_stages")
note(ws, "D15", "Equal pressure ratio per stage, full intercool back to T1. This is the stage "
               "count Sheets 2-5 analyze in detail — it does NOT have to equal Sheet 6's "
               "recommended optimum; update it to match once Sheet 6 tells you what that is.",
     span=1)

section(ws, "B17", "Economics — General & Financial", span=3)
label(ws, "B18", "Operating hours per year")
inp(ws, "C18", 8000, "#,##0", name="hours_per_year")
label(ws, "B19", "Electricity cost ($/kWh)")
inp(ws, "C19", 0.07, "0.00", name="elec_cost_kwh")
note(ws, "D19", "$/kWh is the standard utility billing unit even in US-customary practice.", span=1)
label(ws, "B20", "Risk level (Table 8-1, Peters & Timmerhaus)")
risk_table = [
    ("Safe - cost of capital", 0.04, 0.08),
    ("Low - new capacity / established market", 0.08, 0.16),
    ("Medium - new product/established market or new process tech", 0.16, 0.24),
    ("High - new product/process / new application", 0.24, 0.32),
    ("Very High - all-new / high R&D and marketing", 0.32, 0.48),
]
risk_cell_row = 20
inp(ws, f"C{risk_cell_row}", risk_table[1][0], name="risk_level")
dv_risk_input = DataValidation(type="list", formula1='"' + ",".join(x[0] for x in risk_table) + '"',
                                allow_blank=False)
ws.add_data_validation(dv_risk_input)
dv_risk_input.add(ws[f"C{risk_cell_row}"])
label(ws, "B21", "mar, minimum acceptable return (midpoint of risk level, editable)")
mar_row = 21
# VLOOKUP target (Table 8-1) lives on Sheet 6; filled in with the real row after that
# sheet is built, via a placeholder resolved below (MAR_TABLE_REF).
label(ws, "B22", "N, evaluation period (yr)")
n_eval_row = 22
inp(ws, "C22", 10, "0", name="N_eval_years")
mar_link_row = 23

r = 25
section(ws, f"B{r}", "Compressor Type Selection — automatic", span=3)
r += 1
note(ws, f"B{r}", "Compressor selection depends on adiabatic head, not manual judgment — auto-computed "
                  "below from Q and H_ad at the current N_stages, using the same Fig. 10-64 box-match "
                  "logic as Sheet 2 (all stages share the same head under the equal-split assumption, so "
                  "one type applies to the whole train at this N — see Sheet 6 for how type CAN genuinely "
                  "differ across different N choices, since head there varies a lot more).", span=3)
r += 2
coverage = [
    ("Diaphragm", 1, 10, 1000, 100000),
    ("Reciprocating", 1, 3000, 500, 100000),
    ("Rotary - Straight Lobe", 100, 3000, 100, 1000),
    ("Rotary - Screw", 10, 3000, 100, 10000),
    ("Rotary - Sliding Vane", 10, 1000, 100, 3000),
    ("Rotary - Liquid Piston", 1, 300, 1000, 10000),
    ("Centrifugal - Single Stage", 1000, 100000, 100, 10000),
    ("Centrifugal - Multistage", 500, 100000, 1000, 100000),
    ("Axial - Multistage", 3000, 100000, 100, 3000),
    ("Fan / Blower", 1000, 100000, 10, 1000),
]
type_list = [c[0] for c in coverage]
label(ws, f"B{r}", "Auto-suggested compressor type (from adiabatic head — patched in once Sheet 2 exists)",
      bold=True)
auto_type_row = r
ws.merge_cells(f"C{r}:D{r}")
r += 1
label(ws, f"B{r}", "Manual override (pick a type for a one-off case, or leave \"Auto\")")
override_type_row = r
inp(ws, f"C{r}", "Auto", name="type_override")
ws.merge_cells(f"C{r}:D{r}")
dv_type = DataValidation(type="list", formula1='"Auto,' + ",".join(type_list) + '"', allow_blank=False)
ws.add_data_validation(dv_type)
dv_type.add(ws[f"C{r}"])
r += 1
label(ws, f"B{r}", "Effective selected type (used everywhere)", bold=True)
sel_type_row = r
calc(ws, f"C{r}", f'=IF(C{override_type_row}="Auto",C{auto_type_row},C{override_type_row})')
ws.merge_cells(f"C{r}:D{r}")
ws[f"C{r}"].fill = FLAG_FILL
r += 2

section(ws, f"B{r}", "Efficiency Method — auto-selected by type", span=3)
r += 1
note(ws, f"B{r}", "Reciprocating, Centrifugal, and Axial each have a real digitized correlation chart in "
                  "the source deck (Fig. 3 and Fig. 1) — every other type (Rotary sub-types, Diaphragm, "
                  "Fan/Blower) auto-falls-back to the rule-of-thumb Quick Default. Override to Scale-Up, "
                  "Balje, or Manual below for a one-off case that needs more precision than either gives; "
                  "see Sheet 3 for what data each of those needs.", span=3)
r += 2
label(ws, f"B{r}", "Method override (pick one, or leave \"Auto\" for the type-based default)")
override_method_row = r
inp(ws, f"C{r}", "Auto", name="method_override")
dv_method = DataValidation(
    type="list", formula1='"Auto,Empirical Correlation,Quick Default,Scale-Up,Balje,Manual"',
    allow_blank=False)
ws.add_data_validation(dv_method)
dv_method.add(ws[f"C{r}"])
r += 1
label(ws, f"B{r}", "Effective method (used everywhere)", bold=True)
method_row = r
calc(ws, f"C{r}",
     f'=IF(C{override_method_row}<>"Auto",C{override_method_row},'
     f'IF(OR(ISNUMBER(SEARCH("Reciprocating",C{sel_type_row})),'
     f'ISNUMBER(SEARCH("Centrifugal",C{sel_type_row})),'
     f'ISNUMBER(SEARCH("Axial",C{sel_type_row}))),"Empirical Correlation","Quick Default"))')
ws[f"C{r}"].fill = FLAG_FILL
r += 2
label(ws, f"B{r}", "Manual ηp (used only if Method = Manual)")
manual_eta_row = r
inp(ws, f"C{r}", 0.75, "0.0%", name="eta_p_manual")
r += 2

label(ws, f"B{r}", "Scale-Up method inputs (Casey & Robinson Eq. 17 — needs external friction-"
                   "factor data, see Sheet 3)", italic=True)
r += 1
label(ws, f"B{r}", "η1 (reference efficiency)")
su_eta1_row = r
inp(ws, f"C{r}", 0, "0.000", name="su_eta1")
r += 1
label(ws, f"B{r}", "Bref")
su_bref_row = r
inp(ws, f"C{r}", 0, "0.000", name="su_bref")
r += 1
label(ws, f"B{r}", "fref")
su_fref_row = r
inp(ws, f"C{r}", 0, "0.000", name="su_fref")
r += 1
label(ws, f"B{r}", "f1")
su_f1_row = r
inp(ws, f"C{r}", 0, "0.000", name="su_f1")
r += 1
label(ws, f"B{r}", "f2")
su_f2_row = r
inp(ws, f"C{r}", 0, "0.000", name="su_f2")
r += 2

label(ws, f"B{r}", "Balje method inputs (chart itself is external, see Sheet 3)", italic=True)
r += 1
label(ws, f"B{r}", "Shaft speed, N (rpm)")
balje_n_row = r
inp(ws, f"C{r}", 3600, "#,##0", name="balje_n_rpm")
r += 1
label(ws, f"B{r}", "Impeller diameter, D (ft)")
balje_d_row = r
inp(ws, f"C{r}", 1.5, "0.00", name="balje_d_ft")
r += 1
label(ws, f"B{r}", "ηc read from Balje chart (manual entry)")
balje_etac_row = r
inp(ws, f"C{r}", 0, "0.0%", name="balje_etac")
r += 2
section(ws, f"B{r}", "Stage-Count Optimizer — Equipment Cost Basis (Seider correlations)", span=3)
r += 1
note(ws, f"B{r}", "Compressor and intercooler purchased cost use Seider, Seader, Lewin & Widagdo, Product "
                  "and Process Design Principles, Ch. 16.5 — explicit equations (not charts), sized by the "
                  "same horsepower/duty this workbook already computes. Chosen over Peters & Timmerhaus for "
                  "equipment costing specifically; see compressor_cost_comparison_PT_vs_Seider.pdf for why. "
                  "P&T's own PBP screen (Eq. 8-2a/8-2c) is still used below for the accept/reject structure.",
     span=3)
r += 2
label(ws, f"B{r}", "CEPCI Year (annual averages — pick a year, or \"Custom / Manual Entry\")")
cepci_year_row = r
ws.merge_cells(f"C{r}:D{r}")
inp(ws, f"C{r}", "Custom / Manual Entry", name="cepci_year")
r += 1
label(ws, f"B{r}", "Manual CEPCI (used only if Year = \"Custom / Manual Entry\")")
cepci_manual_row = r
inp(ws, f"C{cepci_manual_row}", 567, "0.0", required=True, name="cepci_manual")
r += 1
label(ws, f"B{r}", "Current CEPCI — effective (Seider eqs. calibrated to CE=567, 2013 avg)", bold=True)
cepci_row = r
calc(ws, f"C{cepci_row}",
     f'=IF(C{cepci_year_row}="Custom / Manual Entry",C{cepci_manual_row},'
     f'VLOOKUP(C{cepci_year_row},$F${cepci_year_row + 1}:$G${cepci_year_row + 25},2,FALSE))', "0.0")
ws[f"C{cepci_row}"].fill = FLAG_FILL
note(ws, f"D{cepci_row}", "Defaults to 567 (Seider's own 2013 basis = no escalation) via Custom / Manual "
                          "Entry above. Pick a Year for a real, current-dollar update, or type any other "
                          "value manually.", span=1)
r += 1

# --- CEPCI reference table (columns F:G), anchored at cepci_year_row so it lines up
# with the dropdown. Source: CEPCI_reference.txt (Chemical Engineering magazine annual
# averages, as compiled at toweringskills.com/financial-analysis/cost-indices/). ---
cepci_table = [
    (2001, 394.3), (2002, 395.6), (2003, 402.0), (2004, 444.2), (2005, 468.2),
    (2006, 499.6), (2007, 525.4), (2008, 575.4), (2009, 521.9), (2010, 550.8),
    (2011, 585.7), (2012, 584.6), (2013, 567.3), (2014, 576.1), (2015, 556.8),
    (2016, 541.7), (2017, 567.5), (2018, 603.1), (2019, 607.5), (2020, 596.2),
    (2021, 708.8), (2022, 816.0), (2023, 797.9), (2024, 795.4),
]
label(ws, f"F{cepci_year_row}", "Year", bold=True)
label(ws, f"G{cepci_year_row}", "CEPCI", bold=True)
for c in (f"F{cepci_year_row}", f"G{cepci_year_row}"):
    ws[c].fill = SECTION_FILL
    ws[c].border = BORDER
label(ws, f"F{cepci_year_row + 1}", "Custom / Manual Entry")
for i, (year, value) in enumerate(cepci_table):
    row = cepci_year_row + 2 + i
    yc = ws[f"F{row}"]
    yc.value = year
    yc.number_format = "0"
    yc.border = BORDER
    vc = ws[f"G{row}"]
    vc.value = value
    vc.number_format = "0.0"
    vc.border = BORDER
dv_cepci_year = DataValidation(type="list", formula1=f"$F${cepci_year_row + 1}:$F${cepci_year_row + 25}",
                                allow_blank=False)
ws.add_data_validation(dv_cepci_year)
dv_cepci_year.add(ws[f"C{cepci_year_row}"])
note(ws, f"F{cepci_year_row + 26}",
     "2024 is a Jan-2024 monthly value (annual average not yet finalized in the source) — treat as "
     "approximate. Only escalate a few years at a time; the index's cost-component weighting drifts over "
     "long spans, so accuracy degrades the further apart the two years are.", span=2)
label(ws, f"B{r}", "Compressor driver")
driver_row = r
inp(ws, f"C{driver_row}", "Electric Motor", name="driver")
dv_driver = DataValidation(type="list", formula1='"Electric Motor,Steam Turbine,Gas Turbine"', allow_blank=False)
ws.add_data_validation(dv_driver)
dv_driver.add(ws[f"C{driver_row}"])
r += 1
label(ws, f"B{r}", "Compressor material of construction")
comp_mat_row = r
inp(ws, f"C{comp_mat_row}", "Carbon Steel", name="comp_material")
dv_compmat = DataValidation(type="list", formula1='"Carbon Steel,Stainless Steel,Nickel Alloy"',
                            allow_blank=False)
ws.add_data_validation(dv_compmat)
dv_compmat.add(ws[f"C{comp_mat_row}"])
r += 1
label(ws, f"B{r}", "Compressor bare-module factor, BMF")
comp_bmf_row = r
inp(ws, f"C{comp_bmf_row}", 2.15, "0.00", name="comp_bmf")
r += 2

label(ws, f"B{r}", "Intercooler: shell-and-tube head type", bold=True)
hx_type_row = r
inp(ws, f"C{hx_type_row}", "Floating Head", name="hx_head_type")
dv_hxtype = DataValidation(type="list", formula1='"Floating Head,Fixed Head,U-Tube,Kettle Vaporizer"',
                           allow_blank=False)
ws.add_data_validation(dv_hxtype)
dv_hxtype.add(ws[f"C{hx_type_row}"])
r += 1
label(ws, f"B{r}", "Intercooler shell/tube material")
hx_mat_row = r
inp(ws, f"C{hx_mat_row}", "Carbon steel/Carbon steel", name="hx_material")
dv_hxmat = DataValidation(
    type="list",
    formula1='"Carbon steel/Carbon steel,Carbon steel/Stainless steel,Stainless steel/Stainless steel"',
    allow_blank=False)
ws.add_data_validation(dv_hxmat)
dv_hxmat.add(ws[f"C{hx_mat_row}"])
r += 1
label(ws, f"B{r}", "Intercooler tube length (ft)")
hx_len_row = r
inp(ws, f"C{hx_len_row}", 20, "0", name="hx_tube_length_ft")
dv_hxlen = DataValidation(type="list", formula1='"8,12,16,20"', allow_blank=False)
ws.add_data_validation(dv_hxlen)
dv_hxlen.add(ws[f"C{hx_len_row}"])
r += 1
label(ws, f"B{r}", "Overall heat-transfer coefficient, U (Btu/(hr·ft²·°F))")
u_row = r
inp(ws, f"C{u_row}", 50, "0", name="U_btu")
note(ws, f"D{u_row}", "50 per Seider Example 16.10 (gas-side service) — adjust if you have a better "
                      "estimate for this duty.", span=1)
r += 1
label(ws, f"B{r}", "Cooling water inlet temperature (°F)")
cw_in_row = r
inp(ws, f"C{cw_in_row}", 90, "0", name="cw_in_F")
r += 1
label(ws, f"B{r}", "Cooling water outlet temperature (°F)")
cw_out_row = r
inp(ws, f"C{cw_out_row}", 120, "0", name="cw_out_F")
note(ws, f"D{cw_out_row}", "90/120°F is Seider's own standard cooling-water assumption (Ch. 17.2).", span=1)
r += 1
label(ws, f"B{r}", "Intercooler bare-module factor, BMF")
hx_bmf_row = r
inp(ws, f"C{hx_bmf_row}", 3.29, "0.00", name="hx_bmf")
r += 2

label(ws, f"B{r}", "Combined income tax rate, t (for after-tax cash flow, Seider Eq. 17.8)", bold=True)
tax_row = r
inp(ws, f"C{tax_row}", 0.21, "0%", name="tax_rate")
note(ws, f"D{tax_row}", "21% flat federal (current US corporate rate, post-2017 TCJA) + 0% state — Texas "
                        "has no state corporate income tax. Seider's own Table 17.4 graduated schedule "
                        "(up to 35% federal + 5% state = 40%) predates the 2018 tax-law change; using it "
                        "here would overstate the tax shield.", span=1)
r += 2

section(ws, f"B{r}", "Derived Quantities (do not edit)", span=3)
r += 1
label(ws, f"B{r}", "T1 (°R)")
calc(ws, f"C{r}", "=C11+459.67", "0.00")
t1r_row = r
r += 1
label(ws, f"B{r}", "Specific gas constant, R' = 1545.349/MW (ft·lbf/(lbm·°R))")
calc(ws, f"C{r}", "=1545.349/C6", "0.000")
rprime_row = r
r += 1
label(ws, f"B{r}", "Overall pressure ratio, r_total = P2/P1")
calc(ws, f"C{r}", "=C13/C12", "0.00")
rtotal_row = r
r += 1
label(ws, f"B{r}", "Per-stage pressure ratio, r_stage = r_total^(1/N_stages)")
calc(ws, f"C{r}", f"=C{rtotal_row}^(1/C15)", "0.000")
rstage_row = r
r += 1
label(ws, f"B{r}", "Suction density, rho1 (lbm/ft³) [ideal: rho=P/(R′ZT), P in psia, R′ in psia·ft³/(lbm·R) form]")
calc(ws, f"C{r}", f"=(C12*C6)/(10.7316*C8*C{t1r_row})", "0.000")
rho1_row = r
r += 1
label(ws, f"B{r}", "Inlet actual volumetric flow, Q (ACFM = ft³/min)")
calc(ws, f"C{r}", f"=(C14/60)/C{rho1_row}", "#,##0.0")
q_acfm_row = r
note(ws, f"D{rho1_row}", "rho uses R = 10.7316 psia·ft³/(lbmol·°R) ÷ MW; R' above (1545.349 ft·lbf/(lbm·R)) "
                          "is the head-equation form of the same specific gas constant, just different "
                          "units.", span=1)

IN = "'1. Inputs'!"
MW_C, K_C, Z_C = f"{IN}$C${6}", f"{IN}$C${7}", f"{IN}$C${8}"
T1F_C, P1_C, P2_C, MDOT_C, NST_C = f"{IN}$C${11}", f"{IN}$C${12}", f"{IN}$C${13}", f"{IN}$C${14}", f"{IN}$C${15}"
HOURS_C, ELEC_C = f"{IN}$C${18}", f"{IN}$C${19}"
NEVAL_C = f"{IN}$C${n_eval_row}"
SELTYPE_C = f"{IN}$C${sel_type_row}"
METHOD_C = f"{IN}$C${method_row}"
MANUAL_ETA_C = f"{IN}$C${manual_eta_row}"
SU_ETA1_C, SU_BREF_C, SU_FREF_C, SU_F1_C, SU_F2_C = (f"{IN}$C${su_eta1_row}", f"{IN}$C${su_bref_row}",
                                                       f"{IN}$C${su_fref_row}", f"{IN}$C${su_f1_row}",
                                                       f"{IN}$C${su_f2_row}")
BALJE_N_C, BALJE_D_C, BALJE_ETAC_C = f"{IN}$C${balje_n_row}", f"{IN}$C${balje_d_row}", f"{IN}$C${balje_etac_row}"
T1R_C, RPRIME_C, RTOTAL_C, RSTAGE_C = (f"{IN}$C${t1r_row}", f"{IN}$C${rprime_row}",
                                        f"{IN}$C${rtotal_row}", f"{IN}$C${rstage_row}")
Q_ACFM_C = f"{IN}$C${q_acfm_row}"
CEPCI_C = f"{IN}$C${cepci_row}"
DRIVER_C, COMPMAT_C, COMP_BMF_C = f"{IN}$C${driver_row}", f"{IN}$C${comp_mat_row}", f"{IN}$C${comp_bmf_row}"
HXTYPE_C, HXMAT_C, HXLEN_C = f"{IN}$C${hx_type_row}", f"{IN}$C${hx_mat_row}", f"{IN}$C${hx_len_row}"
U_C, CWIN_C, CWOUT_C, HX_BMF_C = f"{IN}$C${u_row}", f"{IN}$C${cw_in_row}", f"{IN}$C${cw_out_row}", f"{IN}$C${hx_bmf_row}"
TAX_C = f"{IN}$C${tax_row}"

# =========================================================================
# 2. Type Selection
# =========================================================================
ws2 = wb.create_sheet("2. Type Selection")
ws = ws2
set_widths(ws, [3, 30, 14, 14, 14, 14, 45])
title_row(ws, "B2", "2. Compressor Type Selection", span=6)
note(ws, "B3", "\"Compressor selection is primarily based on capacity and head requirements\" — "
              "Compressor Selection.pptx slide 7. Boundaries below are approximate, eyeballed from the embedded "
              "Fig. 10-64 coverage chart. Change the selected type on Sheet 1, not here.", span=6)

section(ws, "B5", "Computed Operating Point", span=3)
label(ws, "B6", "Capacity, Q (ACFM at suction)")
calc(ws, "C6", f"={Q_ACFM_C}", "#,##0.0")
label(ws, "B7", "Adiabatic head, H_ad (ft) [n=k, per stage]")
calc(ws, "C7", f"=({K_C}/({K_C}-1))*{RPRIME_C}*{T1R_C}*"
               f"({RSTAGE_C}^(({K_C}-1)/{K_C})-1)", "#,##0")
label(ws, "B8", "Effective compressor type (auto-computed on Sheet 1 from Q/H_ad above; "
                "override available there)", bold=True)
calc(ws, "C8", f"={SELTYPE_C}")
ws.merge_cells("C8:D8")

section(ws, "B10", "Approx. Coverage Ranges (digitized from Fig. 10-64) — the highlighted "
                   "row(s) contain your current operating point", span=6)
note(ws, "B11", "A live chart was tried here and dropped — Excel's charting has no clean way to "
                "shade a log-log rectangular region, and a ScatterChart approximation proved "
                "unreliable in practice. Row highlighting (conditional formatting, driven by the "
                "same Q/H comparison) is simpler and equally informative.", span=6)
hdr = ["Type", "Q_min (ACFM)", "Q_max (ACFM)", "H_min (ft)", "H_max (ft)"]
for i, h in enumerate(hdr):
    label(ws, f"{get_column_letter(2+i)}13", h, bold=True)
first_cov_row = 14
r = first_cov_row
for name, qmin, qmax, hmin, hmax in coverage:
    label(ws, f"B{r}", name)
    for col, val in zip("CDEF", [qmin, qmax, hmin, hmax]):
        c = ws[f"{col}{r}"]
        c.value = val
        c.number_format = "#,##0"
        c.border = BORDER
    r += 1
last_cov_row = r - 1
highlight_matching_rows(ws, f"B{first_cov_row}:F{last_cov_row}", first_cov_row,
                         "$C$6", "$C$7", "C", "D", "E", "F")

section(ws, f"B{r+1}", "Auto-Suggestion (candidates whose box contains the operating point)", span=6)
r += 2
label(ws, f"B{r}", "Candidate type(s)")
# NOTE: a previous version used TEXTJOIN(), which openpyxl writes without the
# "_xlfn." prefix Excel needs for post-2007 functions -> #NAME? error. Plain "&"
# concatenation (universally supported) replaces it, via a hidden helper cell
# that a visible cell trims/falls back on.
raw_parts = "&".join(
    f'IF(AND($C$6>=C{i},$C$6<=D{i},$C$7>=E{i},$C$7<=F{i}),B{i}&", ","")'
    for i in range(first_cov_row, last_cov_row + 1)
)
candidate_row = r
helper_cell = f"H{candidate_row}"
calc(ws, helper_cell, "=" + raw_parts)
ws[helper_cell].font = Font(size=1, color="F2F2F2")
trimmed_formula = (
    f'=IF(LEN({helper_cell})=0,"(no boundary match - inspect Fig. 10-64 chart manually)",'
    f'LEFT({helper_cell},LEN({helper_cell})-2))'
)
calc(ws, f"C{candidate_row}", trimmed_formula)
ws.merge_cells(f"C{candidate_row}:F{candidate_row}")

note(ws, f"B{candidate_row+2}", "Additional selection parameters mentioned in the deck (Reynolds number, specific "
                          "speed Ns, specific diameter Ds, flow coefficient φ, pressure coefficient Ψ) "
                          "are standard turbomachinery similarity parameters used mainly to refine centrifugal-"
                          "stage/impeller choice once a vendor curve is available — not computed here since "
                          "the deck names them without giving the specific formulas used.", span=6)

qual_row = candidate_row + 5
section(ws, f"B{qual_row}", "Qualitative Notes by Type (from Compressor Selection.pptx)", span=6)
qual = [
    ("Centrifugal [Continuous, Dynamic]", "Multiple stages for lower power input; backward-curved blades most "
     "common; surges at min flow, chokes at max flow."),
    ("Reciprocating [Positive Displacement]", "High pressure ratio, low flow; discharge temperature limits "
     "pressure ratio (lubricant degradation, thermal expansion); multi-stage needs interstage coolers; "
     "~80% efficient; rod-packing failure is an overpressure scenario — PSV required on vent piping."),
    ("Rotary [Positive Displacement]", "Straight lobe, screw, sliding vane, or liquid piston sub-types."),
    ("Axial [Continuous, Dynamic]", "Low head, high flow, small operating range; used for gas turbines/aircraft."),
    ("Fans / Blowers", "Fans <0.5 psi, blowers <1.5 psi delivery; axial fans for high flow/low pressure; "
     "centrifugal blowers most common, no cooling required (ratio ~3.5)."),
    ("Ejectors [Continuous]", "Steam-driven vacuum alternative to a compressor; low maintenance/cost, "
     "environmental tradeoff; not sized on this sheet."),
]
r = qual_row + 1
for name, txt in qual:
    label(ws, f"B{r}", name, bold=True)
    note(ws, f"C{r}", txt, span=4, size=9)
    ws.row_dimensions[r].height = 40
    r += 2

img_row = r + 1
label(ws, f"B{img_row}", "Fig. 10-64 — Compressor coverage chart (Perry's, via Compressor Selection.pptx), "
                          "for visual cross-check against the live plot above", bold=True)
img = XLImage(os.path.join(CHARTS, "fig10-64_compressor_coverage_chart.png"))
img.width, img.height = 560, 400
ws.add_image(img, f"B{img_row+1}")

# --- Now that Sheet 2's coverage-table location is known, patch in Sheet 1's auto-
# suggested compressor type formula (Sheet 1 was built first, so this cell was left
# blank there — same deferred-patch pattern as the mar hyperlink further below). Reuses
# Sheet 2's own already-computed Q (C6) and H_ad (C7) at Sheet 1's current N_stages. ---
auto_type_raw = suggestion_formula("'2. Type Selection'!$C$6", "'2. Type Selection'!$C$7",
                                    "'2. Type Selection'!", first_cov_row, last_cov_row)
calc(ws1, f"J{auto_type_row}", "=" + auto_type_raw)
ws1[f"J{auto_type_row}"].font = Font(size=1, color="F2F2F2")
calc(ws1, f"C{auto_type_row}",
     f'=IF(LEN(J{auto_type_row})=0,"None",LEFT(J{auto_type_row},SEARCH(", ",J{auto_type_row})-1))')
ws1[f"C{auto_type_row}"].fill = FLAG_FILL

# =========================================================================
# 3. Efficiency
# =========================================================================
ws3 = wb.create_sheet("3. Efficiency")
ws = ws3
set_widths(ws, [3, 34, 12, 12, 12, 12, 12, 40])
title_row(ws, "B2", "3. Polytropic Efficiency Estimation", span=7)
note(ws, "B3", "Three methods from Compressor Sizing.pptx slide 4-7. The method choice and all "
              "manual inputs (Scale-Up, Balje, Manual ηp) are set on Sheet 1 — this sheet shows the "
              "resulting computation for each.", span=7)

section(ws, "B5", "Method 1: Empirical Correlation — three digitized curves by type", span=7)
note(ws, "B6", "ηp increases with flow rate and pressure ratio; general range 70-80% (Compressor Sizing.pptx "
              "slide 5). Reciprocating uses Fig. 3 (vs. pressure ratio); Centrifugal and Axial use the two "
              "curves on Fig. 1 (vs. suction ACFM) — added when the deck was updated with that chart. Any "
              "other type still has no digitized chart and falls back to Quick Defaults below.", span=7)
r = 8
label(ws, f"B{r}", "Reciprocating (Fig. 3, vs. pressure ratio)", bold=True)
r += 1
label(ws, f"B{r}", "Pressure ratio", bold=True)
label(ws, f"C{r}", "ηp", bold=True)
pr_table = [1.5, 2.0, 2.5, 3.0, 3.5, 4.0, 4.5, 5.0, 5.5, 6.0, 6.5]
eff_table = [0.74, 0.79, 0.815, 0.825, 0.833, 0.840, 0.843, 0.846, 0.848, 0.846, 0.843]
first_eff_row = r + 1
r = first_eff_row
for pr, eff in zip(pr_table, eff_table):
    c1 = ws[f"B{r}"]; c1.value = pr; c1.number_format = "0.0"; c1.border = BORDER
    c2 = ws[f"C{r}"]; c2.value = eff; c2.number_format = "0.0%"; c2.border = BORDER
    r += 1
last_eff_row = r - 1
label(ws, f"B{r}", "Interpolated ηp (Reciprocating)")
recip_interp_row = r
calc(ws, f"C{r}", "=" + interp_formula(RSTAGE_C, first_eff_row, last_eff_row), "0.0%")
img = XLImage(os.path.join(CHARTS, "fig3_reciprocating_efficiency_chart.png"))
img.width, img.height = 300, 230
ws.add_image(img, f"E{first_eff_row}")
r += 2

label(ws, f"B{r}", "Centrifugal (Fig. 1, vs. suction ACFM)", bold=True)
r += 1
label(ws, f"B{r}", "Suction ACFM", bold=True)
label(ws, f"C{r}", "ηp", bold=True)
cent_acfm = [1000, 2000, 4000, 7000, 10000, 20000, 40000, 70000, 100000, 200000]
cent_eff = [0.650, 0.700, 0.720, 0.730, 0.735, 0.745, 0.755, 0.763, 0.766, 0.770]
first_cent_row = r + 1
r = first_cent_row
for q, eff in zip(cent_acfm, cent_eff):
    c1 = ws[f"B{r}"]; c1.value = q; c1.number_format = "#,##0"; c1.border = BORDER
    c2 = ws[f"C{r}"]; c2.value = eff; c2.number_format = "0.0%"; c2.border = BORDER
    r += 1
last_cent_row = r - 1
label(ws, f"B{r}", "Interpolated ηp (Centrifugal)")
cent_interp_row = r
calc(ws, f"C{r}", "=" + interp_formula(Q_ACFM_C, first_cent_row, last_cent_row), "0.0%")
r += 2

label(ws, f"B{r}", "Axial (Fig. 1, vs. suction ACFM)", bold=True)
r += 1
label(ws, f"B{r}", "Suction ACFM", bold=True)
label(ws, f"C{r}", "ηp", bold=True)
axial_acfm = [75000, 100000, 200000, 400000, 600000]
axial_eff = [0.813, 0.817, 0.827, 0.832, 0.836]
first_axial_row = r + 1
r = first_axial_row
for q, eff in zip(axial_acfm, axial_eff):
    c1 = ws[f"B{r}"]; c1.value = q; c1.number_format = "#,##0"; c1.border = BORDER
    c2 = ws[f"C{r}"]; c2.value = eff; c2.number_format = "0.0%"; c2.border = BORDER
    r += 1
last_axial_row = r - 1
label(ws, f"B{r}", "Interpolated ηp (Axial)")
axial_interp_row = r
calc(ws, f"C{r}", "=" + interp_formula(Q_ACFM_C, first_axial_row, last_axial_row), "0.0%")
img = XLImage(os.path.join(CHARTS, "fig1_centrifugal_axial_efficiency_chart.png"))
img.width, img.height = 340, 260
ws.add_image(img, f"E{first_cent_row}")
r += 2

label(ws, f"B{r}", "Empirical Correlation result (routes by effective type)", bold=True)
method1_row = r
method1_formula = (
    f'=IF(ISNUMBER(SEARCH("Reciprocating",{SELTYPE_C})),C{recip_interp_row},'
    f'IF(ISNUMBER(SEARCH("Centrifugal",{SELTYPE_C})),C{cent_interp_row},'
    f'IF(ISNUMBER(SEARCH("Axial",{SELTYPE_C})),C{axial_interp_row},'
    f'"n/a - no digitized chart for this type, use Quick Default")))'
)
calc(ws, f"C{method1_row}", method1_formula, "0.0%")
r += 2

section(ws, f"B{r}", "Quick Defaults by Type (rule-of-thumb, 65-90% range per selection deck)", span=7)
r += 1
defaults = [
    ("Reciprocating", 0.80), ("Centrifugal - Single Stage", 0.76), ("Centrifugal - Multistage", 0.75),
    ("Axial - Multistage", 0.87), ("Rotary - Screw", 0.72), ("Rotary - Sliding Vane", 0.70),
    ("Rotary - Straight Lobe", 0.68), ("Rotary - Liquid Piston", 0.65), ("Fan / Blower", 0.72),
]
default_first = r
for name, eff in defaults:
    label(ws, f"B{r}", name)
    c = ws[f"C{r}"]; c.value = eff; c.number_format = "0%"; c.border = BORDER
    r += 1
default_last = r - 1
method2_row = r + 1
label(ws, f"B{method2_row}", "Default ηp for Sheet-1 selected type", bold=True)
calc(ws, f"C{method2_row}",
     f"=IFERROR(INDEX($C${default_first}:$C${default_last},MATCH({SELTYPE_C},$B${default_first}:$B${default_last},0)),0.75)",
     "0.0%")

r = default_last + 3
section(ws, f"B{r}", "Method 2: Casey & Robinson Scale-Up (Eq. 17) — needs external friction-factor data",
        span=7)
r += 1
note(ws, f"B{r}", "Δη = η2 - η1 = -Bref · (f2-f1)/fref. Friction factors f come from equations "
                  "23/24 in Appendix 1 of the Casey & Robinson paper — not included in Compressor Sizing.pptx. "
                  "Inputs (η1, Bref, fref, f1, f2) are set on Sheet 1.", span=7)
r += 1
label(ws, f"B{r}", "η1 (reference efficiency)"); calc(ws, f"C{r}", f"={SU_ETA1_C}", "0.000")
label(ws, f"B{r+1}", "Bref"); calc(ws, f"C{r+1}", f"={SU_BREF_C}", "0.000")
label(ws, f"B{r+2}", "fref"); calc(ws, f"C{r+2}", f"={SU_FREF_C}", "0.000")
label(ws, f"B{r+3}", "f1"); calc(ws, f"C{r+3}", f"={SU_F1_C}", "0.000")
label(ws, f"B{r+4}", "f2"); calc(ws, f"C{r+4}", f"={SU_F2_C}", "0.000")
scaleup_row = r + 5
label(ws, f"B{scaleup_row}", "η2 = η1 + Δη", bold=True)
calc(ws, f"C{scaleup_row}", f"=C{r}-IF(C{r+2}=0,0,C{r+1}*(C{r+4}-C{r+3})/C{r+2})", "0.0%")

r = scaleup_row + 2
section(ws, f"B{r}", "Method 3: Balje Chart (centrifugal) — chart is external, Ns/Ds computed only", span=7)
r += 1
note(ws, f"B{r}", "Specific speed Ns = N·√Q / (g·Had)^0.75 ; specific diameter Ds = D·(g·Had)^0.25 / √Q "
                  "(standard turbomachinery similarity parameters, N in rad/s, Q in ft³/s, Had in ft, D in ft). "
                  "N, D, and ηc are set on Sheet 1 — the Balje chart image itself is not in the deck.", span=7)
r += 1
label(ws, f"B{r}", "Shaft speed, N (rpm)"); calc(ws, f"C{r}", f"={BALJE_N_C}", "#,##0")
n_row = r
label(ws, f"B{r+1}", "Impeller diameter, D (ft)"); calc(ws, f"C{r+1}", f"={BALJE_D_C}", "0.00")
label(ws, f"B{r+2}", "Ns (dimensionless)")
calc(ws, f"C{r+2}",
     f"=(C{n_row}*2*PI()/60)*SQRT('2. Type Selection'!$C$6/60)/(32.174*'2. Type Selection'!$C$7)^0.75", "0.000")
label(ws, f"B{r+3}", "Ds (dimensionless)")
calc(ws, f"C{r+3}",
     f"=C{n_row+1}*(32.174*'2. Type Selection'!$C$7)^0.25/SQRT('2. Type Selection'!$C$6/60)", "0.000")
balje_row = r + 5
label(ws, f"B{balje_row}", "ηc read from Balje chart (set on Sheet 1)", bold=True)
calc(ws, f"C{balje_row}", f"={BALJE_ETAC_C}", "0.0%")

r = balje_row + 2
section(ws, f"B{r}", "Selected ηp (feeds Sheet 4 and 5) — method chosen on Sheet 1", span=7)
r += 1
label(ws, f"B{r}", "Method used")
calc(ws, f"C{r}", f"={METHOD_C}")
method_display_row = r
final_eta_row = r + 1
label(ws, f"B{final_eta_row}", "Selected ηp", bold=True)
sel_formula = (
    f'=IF(C{method_display_row}="Empirical Correlation",C{method1_row},'
    f'IF(C{method_display_row}="Quick Default",C{method2_row},'
    f'IF(C{method_display_row}="Scale-Up",C{scaleup_row},'
    f'IF(C{method_display_row}="Balje",C{balje_row},{MANUAL_ETA_C}))))'
)
calc(ws, f"C{final_eta_row}", sel_formula, "0.0%")
ETA_P_REF = f"'3. Efficiency'!$C${final_eta_row}"

# =========================================================================
# 4. Temp Rise & n
# =========================================================================
ws4 = wb.create_sheet("4. Temp Rise & n")
ws = ws4
set_widths(ws, [3, 4, 14, 14, 14, 14, 14, 14, 44])
title_row(ws, "B2", "4. Polytropic Coefficient n and Discharge Temperature per Stage", span=8)
note(ws, "B3", "ηp = [n/(n-1)] / [k/(k-1)]  (Compressor Selection.pptx slide 6) solved for n. Discharge temp "
              "Tout = T1·r_stage^((n-1)/n), assuming each stage's suction is intercooled back to T1. Compare "
              "against Tallow (Compressor Sizing.pptx slide 8, Table below).", span=8)

section(ws, "B5", "Max Allowable Discharge Temperature, Tallow (°F) — RANGES, not single limits", span=8)
tallow = [
    ("Reciprocating", 300, 350, "Flammable gases (e.g. H2) push the effective limit toward 300F"),
    ("Centrifugal", 392, 446, "API 617 + industry practice"),
    ("Rotary", 212, 350, "Heavily dependent on oil type — some oils degrade at temperatures as low as 212F"),
]
r = 6
label(ws, f"B{r}", "Type", bold=True); label(ws, f"C{r}", "Min (°F)", bold=True)
label(ws, f"D{r}", "Max (°F)", bold=True); label(ws, f"E{r}", "Notes", bold=True)
tallow_first = r + 1
r += 1
for name, tmin, tmax, notes_txt in tallow:
    label(ws, f"B{r}", name)
    for col, val in zip("CD", [tmin, tmax]):
        c = ws[f"{col}{r}"]; c.value = val; c.number_format = "0"; c.border = BORDER
    note(ws, f"E{r}", notes_txt, span=3, size=9)
    r += 1
tallow_last = r - 1
note(ws, f"B{r+1}", "The range reflects real variability by construction/lubricant, not measurement "
                    "uncertainty. Tout between Min and Max does not mean \"safe\" — it means the specific "
                    "vendor/oil/seal rating must be checked; this spreadsheet flags it rather than silently "
                    "picking an endpoint.", span=8)

r += 3
section(ws, f"B{r}", "Per-Stage Calculation (up to 6 stages; blank rows beyond N_stages)", span=8)
r += 1
hdrs = ["Stage", "Suction P (psia)", "Discharge P (psia)", "T1 (°R)", "n", "Tout (°R)", "Tout (°F)"]
for i, h in enumerate(hdrs):
    label(ws, f"{get_column_letter(2+i)}{r}", h, bold=True)
stage_first = r + 1
r += 1
for i in range(1, 7):
    row = stage_first + i - 1
    label(ws, f"B{row}", i)
    calc(ws, f"C{row}", f'=IF(B{row}<={NST_C},{P1_C}*{RSTAGE_C}^(B{row}-1),"")', "#,##0.0")
    calc(ws, f"D{row}", f'=IF(B{row}<={NST_C},{P1_C}*{RSTAGE_C}^B{row},"")', "#,##0.0")
    calc(ws, f"E{row}", f'=IF(B{row}<={NST_C},{T1R_C},"")', "0.0")
    n_formula = f'=IF(B{row}<={NST_C},1/(1-(({K_C}-1)/({K_C}*{ETA_P_REF}))),"")'
    calc(ws, f"F{row}", n_formula, "0.000")
    calc(ws, f"G{row}", f'=IF(B{row}<={NST_C},E{row}*{RSTAGE_C}^((F{row}-1)/F{row}),"")', "#,##0.0")
    calc(ws, f"H{row}", f'=IF(B{row}<={NST_C},G{row}-459.67,"")', "#,##0.0")
r = stage_first + 6

label(ws, f"I{stage_first-1}", "vs. Tallow", bold=True)
for i in range(1, 7):
    row = stage_first + i - 1
    flag_formula = tallow_flag_formula(row, SELTYPE_C, NST_C, tallow)
    fc = calc(ws, f"I{row}", flag_formula)
    fc.font = Font(bold=True)
    fc.alignment = Alignment(wrap_text=True, vertical="top")
ws.row_dimensions[stage_first].height = 30

img = XLImage(os.path.join(CHARTS, "fig10-54_polytropic_compression_curves.png"))
img.width, img.height = 300, 250
ws.add_image(img, f"B{stage_first+8}")
note(ws, f"E{stage_first+8}", "n=1 isothermal, n=k adiabatic, n>k polytropic — higher n gives greater "
                              "temperature rise (Compressor Selection.pptx slide 5).", span=4)

# =========================================================================
# 5. Head & Power  (US units only — hp)
# =========================================================================
ws5 = wb.create_sheet("5. Head & Power")
ws = ws5
set_widths(ws, [3, 4, 16, 16, 16, 16])
title_row(ws, "B2", "5. Polytropic Head and Power per Stage", span=5)
note(ws, "B3", "Hp = [n/(n-1)]·Z·R'·T1·[(P2/P1)^((n-1)/n) - 1]  (ft·lbf/lbm ≡ ft of gas). "
              "Pgas = mdot·Hp/ηp (Compressor Selection.pptx slide 6). All US units — hp only, no kW column.",
     span=5)
r = 5
hdrs = ["Stage", "Hp (ft)", "ηp", "Pgas (ft·lbf/hr)", "Pgas (hp)"]
for i, h in enumerate(hdrs):
    label(ws, f"{get_column_letter(2+i)}{r}", h, bold=True)
hp_first = r + 1
r += 1
for i in range(1, 7):
    row = hp_first + i - 1
    src = f"'4. Temp Rise & n'!"
    label(ws, f"B{row}", i)
    calc(ws, f"C{row}",
         f'=IF(B{row}>{NST_C},"",(({src}F{stage_first+i-1})/({src}F{stage_first+i-1}-1))*'
         f'{Z_C}*{RPRIME_C}*{src}E{stage_first+i-1}*'
         f'({RSTAGE_C}^(({src}F{stage_first+i-1}-1)/{src}F{stage_first+i-1})-1))', "#,##0")
    calc(ws, f"D{row}", f"={ETA_P_REF}", "0.0%")
    calc(ws, f"E{row}", f'=IF(B{row}>{NST_C},"",{MDOT_C}*C{row}/D{row})', "#,##0")
    calc(ws, f"F{row}", f'=IF(B{row}>{NST_C},"",E{row}/1980000)', "#,##0.0")
hp_last = hp_first + 5

r = hp_last + 2
section(ws, f"B{r}", "Train Totals (at Sheet 1's N_stages)", span=5)
r += 1
label(ws, f"B{r}", "Total gas power, sum(Pgas)")
calc(ws, f"C{r}", f"=SUM(F{hp_first}:F{hp_last})", "#,##0.0")
label(ws, f"D{r}", "hp")
train_hp_row = r
label(ws, f"B{r+1}", "Per-stage gas power (all stages equal under the equal-split / full-intercool "
                     "assumption)")
calc(ws, f"C{r+1}", f"=F{hp_first}", "#,##0.0")
label(ws, f"D{r+1}", "hp/stage")
per_stage_hp_row = r + 1

# =========================================================================
# 6. Stage Economics — Stage-Count Optimizer (sweeps N=1..6)
# =========================================================================
ws6 = wb.create_sheet("6. Stage Economics")
ws = ws6
set_widths(ws, [3, 46, 13, 13, 13, 13, 13, 13, 30])
title_row(ws, "B2", "6. Stage-Count Optimizer — How Many Stages Should There Be?", span=8)
note(ws, "B3", "This sheet answers a DIFFERENT question than Sheets 2-5: not \"is my chosen N_stages OK\" "
              "but \"what is the economically optimal N, from 1 to 6\". It sweeps every stage count, costs "
              "the compressor + intercoolers actually implied at each one, and finds the largest N whose "
              "next stage still pays for itself — subject to the Tallow temperature constraint always "
              "winning if it's violated. Compressor Sizing.pptx slide 4's \"back to step 1 with 1 "
              "additional stage\" loop is what this sweep automates.", span=8)
note(ws, f"B{4}", "Simplification: ηp is held constant across the sweep at Sheet 3's Selected ηp (computed "
                  "at Sheet 1's current N_stages) — re-deriving ηp per row would need the full method-"
                  "selection logic duplicated 6 times. Acceptable at this Class-5 (±30-50%) estimate "
                  "accuracy; ηp does shift slightly with pressure ratio in reality.", span=8)

section(ws, "B6", "Reference Example (from the annotated PBP PDF — self-check only, do not edit)", span=8)
label(ws, "B7", "ΔTCI: compressor stage (BMF 2.15) + intercooler HX (BMF 3.29), built up through "
                "contingency/fee/working capital")
ws["C7"].value = 631880; ws["C7"].number_format = "$#,##0"; ws["C7"].border = BORDER
label(ws, "B8", "ΔCOM: annual power-savings cash flow")
ws["C8"].value = 129163; ws["C8"].number_format = "$#,##0"; ws["C8"].border = BORDER
label(ws, "B9", "mar (Low-Medium risk, Table 8-1)")
ws["C9"].value = 0.12; ws["C9"].number_format = "0%"; ws["C9"].border = BORDER
label(ws, "B10", "N, evaluation period (yr)")
ws["C10"].value = 15; ws["C10"].number_format = "0"; ws["C10"].border = BORDER
label(ws, "B11", "PBP = ΔTCI/ΔCOM  (Eq. 8-2a)", bold=True)
calc(ws, "C11", "=C7/C8", "0.00")
label(ws, "B12", "PBPref = 0.85/(mar+0.85/N)  (Eq. 8-2c)", bold=True)
calc(ws, "C12", "=0.85/(C9+0.85/C10)", "0.00")
label(ws, "B13", "Decision")
calc(ws, "C13", '=IF(AND(C11<=C12,C11<=5),"ACCEPT",IF(C11<=C12*1.05,"BORDERLINE - REJECT","REJECT"))')
note(ws, "B14", "Expected result: PBP=4.89 yr vs PBPref=4.81 yr → borderline reject. If this cell doesn't "
               "show that, a formula above is broken.", span=8)

section(ws, "B16", "Table 8-1 — mar by Risk Level (Peters & Timmerhaus) — reference for Sheet 1's mar input",
        span=8)
mar_table_anchor = "B16"
r = 17
label(ws, f"B{r}", "Risk level", bold=True); label(ws, f"C{r}", "mar min", bold=True); label(ws, f"D{r}", "mar max", bold=True)
risk_first = r + 1
r += 1
for name, lo, hi in risk_table:
    label(ws, f"B{r}", name)
    ws[f"C{r}"].value = lo; ws[f"C{r}"].number_format = "0%"; ws[f"C{r}"].border = BORDER
    ws[f"D{r}"].value = hi; ws[f"D{r}"].number_format = "0%"; ws[f"D{r}"].border = BORDER
    r += 1
risk_last = r - 1
RISK_TABLE_RNG = f"'6. Stage Economics'!$B${risk_first}:$D${risk_last}"

r += 1
section(ws, f"B{r}", "Stage-Count Sweep — one row per N, every cost constituent visible "
                     "(what stage cost what)", span=8)
r += 1
note(ws, f"B{r}", "Compressor selection is re-done for each N, not fixed at Sheet 1's choice: adiabatic "
                  "head drops non-linearly as stages are added (Compressor Sizing.pptx slide 4), so a "
                  "different N can genuinely land in a different Fig. 10-64 box. \"Suggested Type\" below "
                  "recomputes the box match per row, the same way Sheet 2 does for a single N. Within one "
                  "fixed N (Sheets 2-5), all individual stages share the same head under the equal-split/"
                  "full-intercool assumption, so type does not differ stage-to-stage there — it's only "
                  "across different N choices, here, that it can.", span=8)
r += 2

# Column widths for this wide table (overrides the sheet-wide defaults set above for B:I)
col_widths = {"C": 9, "D": 10, "E": 30, "F": 9, "G": 32, "H": 12, "I": 12, "J": 12, "K": 12,
              "L": 8, "M": 13, "N": 13, "O": 13, "P": 12, "Q": 12, "R": 12, "S": 12, "T": 13,
              "U": 13, "V": 11, "W": 9, "X": 13, "Y": 11, "Z": 15, "AA": 12, "AB": 12,
              "AC": 12, "AD": 11, "AE": 10}
for col, w in col_widths.items():
    ws.column_dimensions[col].width = w

sweep_hdr_row = r
sweep_hdrs = ["N", "r_stage", "H_ad (ft)", "Suggested Type(s)", "Tout (°F)", "Temp Flag",
              "Train Power = Pc (hp)", "Compressor Cb @CE567 ($)", "Compressor Cp ($)",
              "Compressor BMC ($)", "Coolers (qty)", "Cooling Duty/unit (Btu/hr)",
              "Design Pressure (psig)", "LMTD (°F)", "HX Area/unit (ft²)", "HX Cb @CE567 ($)",
              "HX Cp/unit ($)", "HX BMC/unit ($)", "HX BMC Total ($)", "Total Bare Module ($)",
              "Contingency 18% ($)", "Fee 3% ($)", "Total Module Cost ($)", "Working Capital 15% ($)",
              "TOTAL INSTALLED COST ($)", "Incr. Capex ($)", "Pretax Savings ($/yr)",
              "After-Tax Cash Flow ($/yr)", "Incr. PBP vs N-1", "Passes Econ?"]
for i, h in enumerate(sweep_hdrs):
    c = ws.cell(row=sweep_hdr_row, column=2 + i, value=h)
    c.font = Font(bold=True, size=9)
    c.alignment = Alignment(wrap_text=True, vertical="bottom")
ws.row_dimensions[sweep_hdr_row].height = 48

sweep_first = sweep_hdr_row + 1
n_expr = f"(1/(1-(({K_C}-1)/({K_C}*{ETA_P_REF}))))"  # same n for every row (eta_p held constant, see note)
cp_expr = f"(({K_C}/({K_C}-1))*{RPRIME_C}/778.169)"
FD_EXPR = f'IF({DRIVER_C}="Electric Motor",1,IF({DRIVER_C}="Steam Turbine",1.15,1.25))'
FM_COMP_EXPR = f'IF({COMPMAT_C}="Carbon Steel",1,IF({COMPMAT_C}="Stainless Steel",2.5,5))'
FL_EXPR = f'IF({HXLEN_C}=8,1.25,IF({HXLEN_C}=12,1.12,IF({HXLEN_C}=16,1.05,1)))'
RAW_COL = "AG"  # hidden-ish raw suggestion concatenation, off to the side of the visible table
for i in range(1, 7):
    row = sweep_first + i - 1
    label(ws, f"B{row}", i)
    # r_stage(N), H_ad(N) [adiabatic, n=k — same definition as Sheet 2's C7]
    calc(ws, f"C{row}", f"={RTOTAL_C}^(1/B{row})", "0.000")
    calc(ws, f"D{row}", f"=({K_C}/({K_C}-1))*{RPRIME_C}*{T1R_C}*(C{row}^(({K_C}-1)/{K_C})-1)", "#,##0")
    # Suggested Type(s) at this N's H_ad (Q is constant across N, only H_ad changes)
    # BUG FIX: this used to pass literal "$C$6", which on Sheet 2 correctly means "this
    # sheet's own Q cell" — but reused verbatim here it resolved to SHEET 6's own C6 (the
    # Reference Example's $631,880), not capacity, so every row silently matched nothing.
    raw_formula = suggestion_formula(Q_ACFM_C, f"D{row}", "'2. Type Selection'!", first_cov_row, last_cov_row)
    calc(ws, f"{RAW_COL}{row}", "=" + raw_formula)
    ws[f"{RAW_COL}{row}"].font = Font(size=1, color="F2F2F2")
    calc(ws, f"E{row}",
         f'=IF(LEN({RAW_COL}{row})=0,"(no match)",LEFT({RAW_COL}{row},LEN({RAW_COL}{row})-2))')
    ws[f"E{row}"].alignment = Alignment(wrap_text=True, vertical="top")
    primary_type_expr = (
        f'IF(LEN({RAW_COL}{row})=0,"None",LEFT({RAW_COL}{row},SEARCH(", ",{RAW_COL}{row})-1))'
    )
    # Tout(N) — same eta_p/n for every row (documented simplification)
    calc(ws, f"F{row}", f"={T1R_C}*C{row}^(({n_expr}-1)/{n_expr})-459.67", "#,##0.0")
    # Temp flag(N) uses THIS row's suggested type, not Sheet 1's fixed selection
    flag_formula_n = tallow_flag_formula(row, primary_type_expr, "999", tallow, h_col="F")
    calc(ws, f"G{row}", flag_formula_n)
    ws[f"G{row}"].alignment = Alignment(wrap_text=True, vertical="top")
    # Train power(N) = N * mdot*Hp(r_stage(N))/(etap*1980000) — treated as Seider's "consumed
    # power, Pc" directly (this workbook doesn't model a separate motor efficiency on top of
    # ηp; ηp already plays an analogous role to Seider's combined efficiency treatment)
    hp_expr = f"(({n_expr}/({n_expr}-1))*{Z_C}*{RPRIME_C}*{T1R_C}*(C{row}^((({n_expr})-1)/({n_expr}))-1))"
    calc(ws, f"H{row}", f"=B{row}*{MDOT_C}*{hp_expr}/{ETA_P_REF}/1980000", "#,##0.0")

    # --- Compressor: Seider Ch.16.5 (Eq. 16.36/16.37/16.38), ONE machine at this row's Pc ---
    cb_comp_formula = seider_compressor_cb_formula(primary_type_expr, f"H{row}")
    calc(ws, f"I{row}", f"={cb_comp_formula}")
    calc(ws, f"J{row}", f'=IF(ISNUMBER(I{row}),I{row}*({FD_EXPR})*({FM_COMP_EXPR})*({CEPCI_C}/567),'
                         f'"n/a - outside Seider correlation")', "$#,##0")
    calc(ws, f"K{row}", f'=IF(ISNUMBER(J{row}),J{row}*{COMP_BMF_C},"n/a")', "$#,##0")

    # --- Intercoolers: N-1 of them, each at this row's per-stage cooling duty, costed via
    # Seider shell-and-tube (Eq. 16.39-16.45) sized by area from Q=U*A*LMTD ---
    calc(ws, f"L{row}", f"=MAX(B{row}-1,0)", "0")
    calc(ws, f"M{row}", f"={MDOT_C}*{cp_expr}*((F{row}+459.67)-{T1R_C})", "#,##0")
    calc(ws, f"N{row}", f"={P1_C}*C{row}-14.696", "#,##0")
    lmtd_expr = lmtd_formula(f"F{row}", T1F_C, CWIN_C, CWOUT_C)
    calc(ws, f"O{row}", f"={lmtd_expr}", "0.0")
    calc(ws, f"P{row}", f'=IF(O{row}<=0,"n/a - LMTD<=0, check cooling water temps",M{row}/({U_C}*O{row}))',
         "#,##0")
    cb_hx_formula = seider_hx_cb_formula(HXTYPE_C, f"P{row}")
    calc(ws, f"Q{row}", f'=IF(ISNUMBER(P{row}),{cb_hx_formula},"n/a")')
    fm_hx_expr = hx_fm_formula(HXMAT_C, f"P{row}")
    fp_hx_expr = f"(0.9803+0.018*(N{row}/100)+0.0017*(N{row}/100)^2)"
    calc(ws, f"R{row}", f'=IF(ISNUMBER(Q{row}),Q{row}*{fp_hx_expr}*{fm_hx_expr}*({FL_EXPR})*'
                         f'({CEPCI_C}/567),"n/a")', "$#,##0")
    calc(ws, f"S{row}", f'=IF(ISNUMBER(R{row}),R{row}*{HX_BMF_C},"n/a")', "$#,##0")
    calc(ws, f"T{row}", f'=IF(ISNUMBER(S{row}),L{row}*S{row},"n/a")', "$#,##0")

    # --- Build-up: TBM -> contingency/fee -> total module cost -> working capital -> TCI ---
    calc(ws, f"U{row}", f'=IF(AND(ISNUMBER(K{row}),ISNUMBER(T{row})),K{row}+T{row},'
                         f'"n/a - unsupported type for Seider costing")', "$#,##0")
    calc(ws, f"V{row}", f'=IF(ISNUMBER(U{row}),0.18*U{row},"n/a")', "$#,##0")
    calc(ws, f"W{row}", f'=IF(ISNUMBER(U{row}),0.03*U{row},"n/a")', "$#,##0")
    calc(ws, f"X{row}", f'=IF(AND(ISNUMBER(U{row}),ISNUMBER(V{row}),ISNUMBER(W{row})),'
                         f'U{row}+V{row}+W{row},"n/a")', "$#,##0")
    calc(ws, f"Y{row}", f'=IF(ISNUMBER(X{row}),0.15*X{row},"n/a")', "$#,##0")
    calc(ws, f"Z{row}", f'=IF(AND(ISNUMBER(X{row}),ISNUMBER(Y{row})),X{row}+Y{row},"n/a")', "$#,##0")
    ws[f"Z{row}"].font = Font(italic=True, bold=True)
sweep_last = sweep_first + 5

# Incremental capex/savings (AA/AB), Seider-style after-tax + depreciation cash flow (AC, Eq.
# 17.8: (1-t)*pretax savings + straight-line depreciation on the incremental capital, using
# Sheet 1's N (evaluation period) as the depreciation life since no separate life was given),
# incremental PBP in P&T's structure but fed by that cash flow (AD), and Passes-Econ (AE).
for i in range(2, 7):
    row = sweep_first + i - 1
    prow = row - 1
    calc(ws, f"AA{row}", f'=IF(AND(ISNUMBER(Z{row}),ISNUMBER(Z{prow})),Z{row}-Z{prow},"n/a")', "$#,##0")
    calc(ws, f"AB{row}", f"=(H{prow}-H{row})*0.7457*{HOURS_C}*{ELEC_C}", "$#,##0")
    calc(ws, f"AC{row}",
         f'=IF(ISNUMBER(AA{row}),(1-{TAX_C})*AB{row}+AA{row}/{NEVAL_C},"n/a")', "$#,##0")
    # Seider's steep power exponent on reciprocating/centrifugal cost (1.23/0.63) means a lower-
    # power train from adding a stage can make the COMPRESSOR CHEAPER outright, so incr. capex
    # can be negative even after adding an intercooler. A plain capex/cash-flow ratio would then
    # produce a nonsensical negative "PBP" (e.g. -13.6 yr) that happens to pass the <=5 check for
    # the wrong reason — call it out explicitly instead of displaying that number.
    calc(ws, f"AD{row}",
         f'=IF(NOT(ISNUMBER(AA{row})),"n/a",'
         f'IF(AA{row}<=0,"FREE - cheaper capital AND lower operating cost",'
         f'IF(OR(NOT(ISNUMBER(AC{row})),AC{row}<=0),"n/a",AA{row}/AC{row})))')
label(ws, f"AA{sweep_first}", "n/a (baseline)"); ws[f"AA{sweep_first}"].font = Font(italic=True, size=9)
label(ws, f"AB{sweep_first}", "n/a (baseline)"); ws[f"AB{sweep_first}"].font = Font(italic=True, size=9)
label(ws, f"AC{sweep_first}", "n/a (baseline)"); ws[f"AC{sweep_first}"].font = Font(italic=True, size=9)
label(ws, f"AD{sweep_first}", "n/a (baseline)"); ws[f"AD{sweep_first}"].font = Font(italic=True, size=9)

MAR_C = f"{IN}$C${mar_row}"
for i in range(2, 7):
    row = sweep_first + i - 1
    pass_formula = (
        f'=IF(OR(AD{row}="n/a"),FALSE,'
        f'IF(AD{row}="FREE - cheaper capital AND lower operating cost",TRUE,'
        f'IF(AND(ISNUMBER(AD{row}),AD{row}<=0.85/({MAR_C}+0.85/{NEVAL_C}),AD{row}<=5),TRUE,FALSE)))'
    )
    calc(ws, f"AE{row}", pass_formula)
label(ws, f"AE{sweep_first}", "TRUE (baseline)")
ws[f"AE{sweep_first}"].font = Font(italic=True, size=9)

r = sweep_last + 2
section(ws, f"B{r}", "Optimal Stage Count", span=8)
r += 1
label(ws, f"B{r}", "Smallest N with an acceptable discharge temperature (EXCEEDS = infeasible; "
                   "OK, CAUTION, and n/a-no-range-given all count as feasible)")
min_feasible_formula = "=" + "".join(
    f'IF(NOT(ISNUMBER(SEARCH("EXCEEDS",G{sweep_first+i-1}))),{i},'
    for i in range(1, 7)
) + '"NONE - need more than 6 stages"' + ")" * 6
calc(ws, f"C{r}", min_feasible_formula, "0")
min_feasible_row = r
label(ws, f"B{r+1}", "Cumulative economic pass (chained from the min-feasible N upward)")
r += 2
cum_row_first = r
for i in range(1, 7):
    row = sweep_first + i - 1
    cum_cell = f"C{r}"
    if i == 1:
        formula = f'=IF(B{row}<$C${min_feasible_row},FALSE,IF(B{row}=$C${min_feasible_row},TRUE,FALSE))'
    else:
        prev_cum = f"C{r-1}"
        formula = (
            f'=IF(B{row}<$C${min_feasible_row},FALSE,'
            f'IF(B{row}=$C${min_feasible_row},TRUE,AND({prev_cum},AE{row})))'
        )
    label(ws, f"B{r}", f"  N={i} passes chain?", size=8, italic=True)
    calc(ws, cum_cell, formula)
    r += 1
cum_row_last = r - 1

max_pass_row_expr = (
    f"SUMPRODUCT(MAX((C{cum_row_first}:C{cum_row_last}=TRUE)*"
    f"ROW(C{cum_row_first}:C{cum_row_last})))"
)
optimal_row = r + 1
label(ws, f"B{optimal_row}", "RECOMMENDED OPTIMAL N", bold=True)
optimal_formula = (
    f'=IF({max_pass_row_expr}=0,"infeasible within 6 stages",'
    f'{max_pass_row_expr}-{cum_row_first}+1)'
)
calc(ws, f"C{optimal_row}", optimal_formula, "0")
ws[f"C{optimal_row}"].fill = FLAG_FILL
ws[f"C{optimal_row}"].font = Font(bold=True, size=12)

suggested_type_optimal_row = optimal_row + 1
label(ws, f"B{suggested_type_optimal_row}", "Suggested compressor type at that N")
calc(ws, f"C{suggested_type_optimal_row}",
     f'=IFERROR(INDEX($E${sweep_first}:$E${sweep_last},MATCH(C{optimal_row},$B${sweep_first}:$B${sweep_last},0)),"n/a")')
ws.merge_cells(f"C{suggested_type_optimal_row}:E{suggested_type_optimal_row}")

total_project_cost_row = suggested_type_optimal_row + 1
label(ws, f"B{total_project_cost_row}", "TOTAL PROJECT COST at that N (installed compressor + "
                                        "intercoolers, fully built up)", bold=True)
calc(ws, f"C{total_project_cost_row}",
     f'=IFERROR(INDEX($Z${sweep_first}:$Z${sweep_last},MATCH(C{optimal_row},$B${sweep_first}:$B${sweep_last},0)),"n/a")',
     "$#,##0")
ws[f"C{total_project_cost_row}"].fill = FLAG_FILL
ws[f"C{total_project_cost_row}"].font = Font(bold=True, size=13)
TOTAL_PROJECT_COST_REF = f"'6. Stage Economics'!$C${total_project_cost_row}"

compare_row = total_project_cost_row + 2
label(ws, f"B{compare_row}", "Sheet 1's current N_stages, for comparison")
calc(ws, f"C{compare_row}", f"={NST_C}", "0")

combined_row = compare_row + 2
label(ws, f"B{combined_row}", "COMBINED RECOMMENDATION", bold=True)
combined_formula = (
    f'=IF(C{optimal_row}=C{compare_row},'
    f'"Sheet 1\'s N_stages ("&C{compare_row}&") already matches the optimum.",'
    f'"Optimum is "&C{optimal_row}&" stage(s); Sheet 1 is currently set to "&C{compare_row}&'
    f'" — update N_stages on Sheet 1 to "&C{optimal_row}&" and re-check Sheets 2-5.")'
)
calc(ws, f"C{combined_row}", combined_formula)
ws[f"C{combined_row}"].fill = FLAG_FILL
ws[f"C{combined_row}"].font = Font(bold=True)
ws.merge_cells(f"C{combined_row}:F{combined_row}")
ws.row_dimensions[combined_row].height = 30

note(ws, f"B{combined_row+2}", "Column AG (tiny font, off to the side of the visible table) holds the raw "
                                "type-match concatenation per row — deliberately not hidden (Excel drops "
                                "hidden columns from charts, and shrunk-to-invisible is enough for a table "
                                "this isn't charted from).", span=8)

# =========================================================================
# 0. Overview  (protected — read-only rollup of every sheet)
# =========================================================================
ws0 = wb.create_sheet("0. Overview", 0)
ws = ws0
set_widths(ws, [3, 34, 20, 40])
title_row(ws, "B2", "Overview — Compressor Train Selection & Sizing (read-only)", span=3)
note(ws, "B3", "This sheet is protected — every cell here pulls live from the working sheets, it does not "
              "hold its own inputs. All inputs live on Sheet 1; this sheet updates automatically.",
     span=3)

r = 5
section(ws, f"B{r}", "1. Process Inputs", span=3)
r += 1
overview_rows_sheet1 = [
    ("MW (lb/lbmol)", f"={MW_C}", "0.00"),
    ("k = Cp/Cv", f"={K_C}", "0.00"),
    ("T1 (°F)", f"={T1F_C}", "0.0"),
    ("P1 (psia)", f"={P1_C}", "0.0"),
    ("P2 (psia)", f"={P2_C}", "0.0"),
    ("mdot (lb/hr)", f"={MDOT_C}", "#,##0"),
    ("N_stages (Sheets 2-5 basis)", f"={NST_C}", "0"),
    ("Per-stage pressure ratio", f"={RSTAGE_C}", "0.000"),
    ("Inlet flow, Q (ACFM)", f"={Q_ACFM_C}", "#,##0.0"),
]
for lbl, f, fmt in overview_rows_sheet1:
    label(ws, f"B{r}", lbl)
    c = calc(ws, f"C{r}", f, fmt); c.fill = OVERVIEW_FILL
    r += 1

r += 1
section(ws, f"B{r}", "2. Type Selection", span=3)
r += 1
label(ws, f"B{r}", "Adiabatic head, H_ad (ft)")
c = calc(ws, f"C{r}", "='2. Type Selection'!C7", "#,##0"); c.fill = OVERVIEW_FILL
r += 1
label(ws, f"B{r}", "Auto-suggested candidate(s)")
c = calc(ws, f"C{r}", f"='2. Type Selection'!C{candidate_row}"); c.fill = OVERVIEW_FILL
ws.merge_cells(f"C{r}:D{r}")
r += 1
label(ws, f"B{r}", "Selected compressor type", bold=True)
c = calc(ws, f"C{r}", f"={SELTYPE_C}"); c.fill = OVERVIEW_FILL
r += 2

section(ws, f"B{r}", "3. Efficiency", span=3)
r += 1
label(ws, f"B{r}", "Method used")
c = calc(ws, f"C{r}", f"={METHOD_C}"); c.fill = OVERVIEW_FILL
r += 1
label(ws, f"B{r}", "Selected ηp", bold=True)
c = calc(ws, f"C{r}", f"={ETA_P_REF}", "0.0%"); c.fill = OVERVIEW_FILL
r += 2

section(ws, f"B{r}", "4. Temp Rise & n (per stage — all stages equal)", span=3)
r += 1
label(ws, f"B{r}", "Polytropic n")
c = calc(ws, f"C{r}", f"='4. Temp Rise & n'!F{stage_first}", "0.000"); c.fill = OVERVIEW_FILL
r += 1
label(ws, f"B{r}", "Discharge temp, Tout (°F)")
c = calc(ws, f"C{r}", f"='4. Temp Rise & n'!H{stage_first}", "#,##0.0"); c.fill = OVERVIEW_FILL
r += 1
label(ws, f"B{r}", "Temperature flag", bold=True)
c = calc(ws, f"C{r}", f"='4. Temp Rise & n'!I{stage_first}"); c.fill = OVERVIEW_FILL
c.alignment = Alignment(wrap_text=True, vertical="top")
ws.merge_cells(f"C{r}:D{r}")
ws.row_dimensions[r].height = 30
r += 2

section(ws, f"B{r}", "5. Head & Power", span=3)
r += 1
label(ws, f"B{r}", "Per-stage gas power (hp)")
c = calc(ws, f"C{r}", f"='5. Head & Power'!C{per_stage_hp_row}", "#,##0.0"); c.fill = OVERVIEW_FILL
r += 1
label(ws, f"B{r}", "Train total gas power (hp)", bold=True)
c = calc(ws, f"C{r}", f"='5. Head & Power'!C{train_hp_row}", "#,##0.0"); c.fill = OVERVIEW_FILL
r += 2

section(ws, f"B{r}", "6. Stage-Count Optimizer", span=3)
r += 1
label(ws, f"B{r}", "Recommended optimal N", bold=True)
c = calc(ws, f"C{r}", f"='6. Stage Economics'!C{optimal_row}", "0"); c.fill = OVERVIEW_FILL
r += 1
label(ws, f"B{r}", "Suggested compressor type at that N")
c = calc(ws, f"C{r}", f"='6. Stage Economics'!C{suggested_type_optimal_row}"); c.fill = OVERVIEW_FILL
ws.merge_cells(f"C{r}:D{r}")
r += 1
label(ws, f"B{r}", "TOTAL PROJECT COST at that N", bold=True)
c = calc(ws, f"C{r}", f"={TOTAL_PROJECT_COST_REF}", "$#,##0")
c.fill = FLAG_FILL
c.font = Font(bold=True, size=13)
r += 1
label(ws, f"B{r}", "Sheet 1's current N_stages")
c = calc(ws, f"C{r}", f"={NST_C}", "0"); c.fill = OVERVIEW_FILL
r += 2

section(ws, f"B{r}", "COMBINED RECOMMENDATION", span=3)
r += 1
c = calc(ws, f"C{r}", f"='6. Stage Economics'!C{combined_row}")
c.fill = FLAG_FILL
c.font = Font(bold=True, size=11)
c.alignment = Alignment(wrap_text=True, vertical="top")
ws.merge_cells(f"C{r}:D{r}")
ws.row_dimensions[r].height = 40
r += 2

ws.protection.sheet = True
ws.protection.formatCells = True
ws.protection.formatColumns = True
ws.protection.formatRows = True
ws.protection.insertColumns = True
ws.protection.insertRows = True
ws.protection.deleteColumns = True
ws.protection.deleteRows = True
ws.protection.sort = True
ws.protection.autoFilter = True

# --- Now that Sheet 6's Table 8-1 location is known, add the mar hyperlink on Sheet 1 ---
hyperlink_note(ws1, f"B{mar_link_row}", "What is mar? -> Table 8-1 guidance (Sheet 6)",
               "6. Stage Economics", mar_table_anchor, span=2)
calc(ws1, f"C{mar_row}",
     f'=AVERAGE(VLOOKUP(C{risk_cell_row},{RISK_TABLE_RNG},2,FALSE),'
     f'VLOOKUP(C{risk_cell_row},{RISK_TABLE_RNG},3,FALSE))', "0.0%")

# =========================================================================
# README
# =========================================================================
ws_readme = wb.create_sheet("README", 1)
ws = ws_readme
set_widths(ws, [3, 32, 60, 14, 14, 14, 14])
title_row(ws, "B2", "Compressor Train Selection & Sizing", span=5)
label(ws, "B4", "Purpose", bold=True)
note(ws, "B5", "Sheets 1-5 analyze ONE stage count in full physics detail (Sheet 1's N_stages). Sheet 6 "
              "independently sweeps N=1..6, costs the equipment each implies, and recommends the largest "
              "stage count whose next stage still pays for itself — i.e. it calculates the economically "
              "optimal number of stages, it does not just screen whatever N you happened to pick. If Sheet "
              "6's recommendation differs from Sheet 1's N_stages, update Sheet 1 to match and Sheets 2-5 "
              "re-derive automatically.", span=5)
label(ws, "B7", "Source decks", bold=True)
note(ws, "B8", "Compressor Selection.pptx, Compressor Sizing.pptx, Heat Exchanger Heuristics.pptx "
              "(all in this folder), PT_payback_period_p345.pdf (Peters & Timmerhaus, Ch. 8), and "
              "..\\FEL Projects\\Economics\\notes\\01_capital_cost_estimation.md (Guthrie/Seider module "
              "costing method reused for Sheet 6). These are the source of truth for the method used here "
              "— update this workbook if the decks change rather than hand-patching formulas.", span=5)

label(ws, "B10", "Sheet order = Sizing Workflow (Compressor Sizing.pptx, slide 4)", bold=True)
rows = [
    ("0. Overview", "Protected, read-only rollup of every sheet's key result, including the recommended "
                     "optimal stage count and the total project cost at that count. Cannot be edited — "
                     "change inputs on Sheet 1 instead."),
    ("1. Inputs", "EVERY input in the workbook: gas properties, process conditions, N_stages, economics "
                  "(hours, elec. cost, mar, evaluation period), and the cost-scaling reference points for "
                  "Sheet 6. Compressor type and efficiency method are auto-computed here (not manual "
                  "inputs) — an \"Auto\" override dropdown for each lets you force a specific choice for a "
                  "one-off case. Sheets 2-6 are read-only / derived from here."),
    ("2. Type Selection", "Computes capacity (ACFM) and adiabatic head (ft); the coverage-range table's "
                           "matching row(s) are highlighted (conditional formatting) rather than plotted — "
                           "a live chart was tried and dropped, since Excel has no clean way to shade a "
                           "log-log rectangular region and the ScatterChart approximation proved unreliable "
                           "in practice. This sheet's Q/H_ad feed Sheet 1's auto-suggestion directly — the "
                           "effective type itself lives on Sheet 1, not here."),
    ("3. Efficiency", "Three methods per the Sizing deck: empirical correlation — three digitized curves "
                       "now (Fig. 3 Reciprocating vs. pressure ratio; Fig. 1's two curves, Centrifugal and "
                       "Axial, vs. suction ACFM) — Casey & Robinson scale-up (needs friction-factor data "
                       "not in the deck), and Balje chart (Ns/Ds computed; chart itself is external, needs "
                       "a manual ηc reading). Sheet 1 auto-picks Empirical Correlation when the (also auto) "
                       "type is Reciprocating/Centrifugal/Axial, else Quick Default (rule-of-thumb) — "
                       "override to Scale-Up/Balje/Manual there for a one-off case needing more precision."),
    ("4. Temp Rise & n", "Solves polytropic n from ηp and k, computes stage discharge temperature. Tallow is "
                          "a RANGE per type (not a single cutoff) — a 3-tier flag (OK / CAUTION-inspect-"
                          "specifics / EXCEEDS-add-a-stage) reflects that the true limit depends on "
                          "lubricant/oil type and similar specifics outside this spreadsheet's scope."),
    ("5. Head & Power", "Polytropic head and gas power per stage, summed to train totals — US units only "
                         "(hp); no SI kW column."),
    ("6. Stage Economics", "The Stage-Count Optimizer. Sweeps N=1..6 (not just Sheet 1's current N_stages) "
                            "in one wide table, one row per N, with every cost constituent broken out as "
                            "its own column: compressor cost via Seider Ch. 16.5 (Eq. 16.36/37/38, by "
                            "consumed power, driver and material factors, CEPCI-escalated), intercooler "
                            "cost via Seider's shell-and-tube equations (area from an assumed U and "
                            "LMTD against a standard 90/120°F cooling-water assumption), then Guthrie "
                            "bare-module + contingency + fee + working capital to the final installed cost "
                            "— so it's traceable what each stage count actually costs. Compressor type is "
                            "re-suggested per row, not fixed at Sheet 1's choice — adiabatic head drops "
                            "non-linearly as stages are added, so a different N can land in a different "
                            "Fig. 10-64 box (within one fixed N on Sheets 2-5, all stages share the same "
                            "head and so the same type, since the model splits pressure ratio equally — "
                            "it's only across different N here that type can genuinely change). The payback "
                            "screen keeps Peters & Timmerhaus's Eq. 8-2a/8-2c structure, but its cash flow "
                            "is now after-tax with straight-line depreciation added back (Seider Eq. 17.8) "
                            "instead of the earlier pre-tax gross savings. Finds the largest N where every "
                            "step up to it still passed that screen, subject to the Tallow temperature "
                            "constraint as a hard floor, and reports the TOTAL PROJECT COST at that N. "
                            "Includes a verified worked reference example as a self-check (unaffected by "
                            "the cost-engine change — it exercises Eq. 8-2a/8-2c directly, not the sweep)."),
]
r = 11
for name, desc in rows:
    label(ws, f"B{r}", name, bold=True)
    note(ws, f"C{r}", desc, span=3, size=9)
    ws.row_dimensions[r].height = 56
    r += 2

label(ws, f"B{r+1}", "Known gaps (not in the source decks — flagged inline, not silently guessed)", bold=True)
r += 2
gaps = [
    "Casey & Robinson friction-factor equations 23/24 (Appendix 1 of their paper) are referenced but not "
    "included in Compressor Sizing.pptx — the scale-up method needs Bref/f values from that external source "
    "or vendor data.",
    "The Balje chart image itself is not embedded in the deck (only the method is named) — Ns/Ds are "
    "computed here but ηc must be read off the chart elsewhere and entered manually.",
    "Intercooler sizing here uses an assumed overall U (50 Btu/hr-ft²-F, Seider's own Ex. 16.10 precedent) "
    "and a standard 90/120°F cooling-water assumption to get an area for Seider's shell-and-tube cost "
    "equations — not a full thermal/mechanical HX design (that's still deferred per Compressor Sizing.pptx "
    "slide 10).",
    "Seider Ch. 16.5 gives explicit compressor cost equations only for Centrifugal, Reciprocating, and "
    "Screw — Diaphragm, other Rotary sub-types, Axial, and Fan/Blower selections show 'n/a - outside "
    "Seider correlation coverage' rather than a fabricated fit; that row's TCI reads 'n/a' too.",
    "Current CEPCI defaults to 567 (Seider's own 2013 basis, i.e. no escalation) — update it on Sheet 1 "
    "from chemengonline.com for a real, current-dollar estimate.",
    "Combined tax rate defaults to 21% (flat federal, Texas has no state corporate income tax) — edit on "
    "Sheet 1 if a different jurisdiction or the pre-2018 graduated schedule applies.",
    "Depreciation life for the after-tax cash flow reuses Sheet 1's N (evaluation period) rather than a "
    "separately-specified MACRS/straight-line life — a deliberate simplification since no depreciation "
    "life was specified; revisit if a real one is known.",
    "The Stage-Count Optimizer holds ηp constant across the N=1..6 sweep (Sheet 3's value at Sheet 1's "
    "current N_stages) rather than re-deriving it per row — a documented Class-5 (±30-50%) screening "
    "simplification, not an error.",
    "Fig. 10-64 type-selection boundaries are eyeballed from the chart image, not precisely digitized — "
    "cross-check visually (static image) for borderline operating points.",
]
for g in gaps:
    note(ws, f"B{r}", "• " + g, span=5, size=9)
    ws.row_dimensions[r].height = 32
    r += 1

with open(CELL_MAP_PATH, "w") as f:
    json.dump(INPUT_CELL_MAP, f, indent=2, sort_keys=True)
print(f"Saved {CELL_MAP_PATH} ({len(INPUT_CELL_MAP)} inputs)")

wb.save(OUT_PATH)
print(f"Saved {OUT_PATH}")
