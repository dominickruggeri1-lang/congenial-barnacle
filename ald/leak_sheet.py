"""Build the ALD field leak sheet template (corrected version of Template Leak Sheet.xlsx).

    python -m ald.leak_sheet [-o output/] [--company "Plant name"] [--rate 0.18]

Keeps the legacy column layout so the leak-detection-report skill reads it unchanged:
row 7 headers, data from row 8, B gas, C pressure, D dB, E location, I description,
M repaired, N cost, O CFM, P kWh, Q unrepaired CFM; A6/B6 = Air CFM and cost;
Cost!D6 = $/kWh, Cost!D7 = air $/1000 cf.

Fixes against the 2023-2025 templates:
  * Any pressure from 10 to 150 psi works. Between the seven UE Systems curves the
    CFM is linearly interpolated, and the row is marked INTERPOLATED.
  * CFM is computed from the UE Systems power curves (CFM = a x dB^b) at the actual dB.
    The old VLOOKUP truncated fractional dB down to the integer row.
  * % repaired counts only "Y". The old COUNTA counted "N" as repaired.
  * kWh is shown for Air and Open Blow only. The old sheet divided nitrogen, argon etc.
    dollars by the electricity rate and reported that as kWh.
  * Year, electricity rate and run hours are labeled inputs, not a sheet named "2023".
  * No FILTER(): totals use SUMIFS so the sheet works in every Excel version.
"""
import argparse
import datetime as dt
import os

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FONT = "Arial"
INPUT = Font(name=FONT, color="0000FF")
BODY = Font(name=FONT)
BOLD = Font(name=FONT, bold=True)
HEAD = Font(name=FONT, bold=True, color="FFFFFF")
TITLE = Font(name=FONT, bold=True, size=13)
YELLOW = PatternFill("solid", fgColor="FFFF00")
BLACK = PatternFill("solid", fgColor="0D0D0D")
REDF = PatternFill("solid", fgColor="F8D7DA")
AMBER = PatternFill("solid", fgColor="FFF3CD")
THIN = Border(bottom=Side(style="thin", color="D8D8D8"))
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)

FIRST, LAST = 8, 500     # leak rows, same range the legacy sheet summed
LOG = "Leak Log"
L = f"'{LOG}'"

# UE Systems Inc. Ultraprobe "Compressed Gas Loss Guess-Timator" (c) 2008, as carried in the
# hidden Flow Rate Table of ALD's leak sheet templates: CFM = a * dB ^ b at each pressure.
CURVES = [  # psi, a, b   (ascending pressure, required by MATCH)
    (10, 0.0004, 2.0766),
    (25, 0.001, 1.9185),
    (50, 0.0031, 1.6865),
    (75, 0.0082, 1.4978),
    (100, 0.014, 1.4004),
    (125, 0.02, 1.3399),
    (150, 0.0297, 1.2712),
]

# Gas, $/1000 cf. Air is a formula on the electricity rate. Others are the values carried in the
# ALD templates; their source is not documented there, so they are flagged as such.
GASES = [
    ("Air", None),
    ("Open Blow", None),
    ("Carbon Dioxide", 0.01),
    ("Argon", 22.80),
    ("Helium", 99.95),
    ("Hydrogen", 11.77),
    ("Nitrogen", 3.80),
]


def build(company="Company name here", rate=None, year=None):
    wb = Workbook()
    cost = wb.active
    cost.title = "Cost"
    log = wb.create_sheet(LOG)
    curves = wb.create_sheet("Flow Curves")

    # ---------- Cost (inputs) ----------
    cost.column_dimensions["B"].width = 18
    cost.column_dimensions["C"].width = 10
    cost.column_dimensions["D"].width = 12
    cost.column_dimensions["E"].width = 24
    cost.column_dimensions["F"].width = 90
    cost["B1"] = "ALD Field Leak Sheet · Cost Inputs"
    cost["B1"].font = TITLE
    cost["B2"] = "Blue on yellow = input. Edit only these cells. Basis for each value is in column F."
    cost["C3"] = "This is for the year"
    cost["D3"] = year or dt.date.today().year
    cost["D3"].font, cost["D3"].fill = INPUT, YELLOW
    cost["B5"], cost["D5"] = "Cubic Feet =", 1000
    cost["B6"], cost["C6"] = "Electricity", "Cost ="
    cost["D6"] = rate if rate is not None else 0.13
    cost["D6"].number_format = "$0.0000"
    cost["D6"].font, cost["D6"].fill = INPUT, YELLOW
    cost["E6"] = "Per kWh"
    cost["F6"] = ("INPUT REQUIRED. Use the plant's blended $/kWh from 12 months of bills. "
                  "0.13 is a placeholder and is flagged on the Leak Log until changed.")
    cost["B7"], cost["C7"] = "Air", "Cost ="
    cost["D7"] = "=$D$13*$D$6"
    cost["E7"] = '=_xlfn.CONCAT("Per ",$D$5," Cubic Feet")'
    cost["F7"] = "Calc: kWh per 1000 cf (D13) x electricity rate."
    r = 8
    gas_cost_cell = {"Air": "Cost!$D$7", "Open Blow": "Cost!$D$7"}
    for gas, val in GASES[2:]:
        cost.cell(r, 2, gas)
        cost.cell(r, 3, "Cost =")
        c = cost.cell(r, 4, val)
        c.font, c.fill, c.number_format = INPUT, YELLOW, "$0.00"
        cost.cell(r, 5, f'=_xlfn.CONCAT("Per ",$D$5," Cubic Feet")')
        cost.cell(r, 6, "Carried from ALD template 2023-2025. Source not documented: confirm with the "
                        "plant's gas supplier invoice before quoting savings on this gas.")
        gas_cost_cell[gas] = f"Cost!$D${r}"
        r += 1
    cost["B13"] = "kWh per 1000 cf of air"
    cost["D13"] = 3.29
    cost["D13"].font, cost["D13"].fill = INPUT, YELLOW
    cost["F13"] = ("ALD template factor (3.29 kWh/1000 cf = 0.197 kW/CFM, DOE Sourcebook planning value "
                   "at about 100 psi). Replace with the system's measured specific power when known.")
    cost["B15"] = "Operational Times"
    cost["B15"].font = BOLD
    cost["B16"], cost["D16"] = "Hours Per Day", 24
    cost["B17"], cost["D17"] = "Days Per Year", 365
    for a in ("D16", "D17"):
        cost[a].font, cost[a].fill = INPUT, YELLOW
    cost["F16"] = "ALD default 24/7. Leaks bleed whenever the header is pressurized, so use pressurized hours, not production hours."
    cost["B19"] = "Plant kW per leak CFM"
    cost["D19"] = "=$D$13*60/1000"
    cost["D19"].number_format = "0.0000"
    cost["F19"] = "Calc. Used for kW and bhp."
    # gas lookup table for the log formulas (hidden helper, H:I)
    cost["H4"], cost["I4"] = "Gas", "$/1000 cf"
    for i, (gas, _) in enumerate(GASES, start=5):
        cost.cell(i, 8, gas)
        cost.cell(i, 9, f"={gas_cost_cell[gas]}")
    cost.column_dimensions["H"].hidden = True
    cost.column_dimensions["I"].hidden = True
    GAS_N = f"Cost!$H$5:$H${4 + len(GASES)}"
    GAS_C = f"Cost!$I$5:$I${4 + len(GASES)}"

    # ---------- Flow Curves ----------
    curves["A1"] = "UE Systems Ultraprobe Compressed Gas Loss Guess-Timator, (c) 2008 UE Systems Inc."
    curves["A1"].font = BOLD
    curves["A2"] = ("CFM = a x dB^b. Values are a general guideline, not a measured flow (UE Systems notice). "
                    "Readings compensated for atmospheric pressure; probe held 15 in from the leak.")
    curves["A4"], curves["B4"], curves["C4"] = "Pressure (psi)", "a", "b"
    for c in ("A4", "B4", "C4"):
        curves[c].font = BOLD
    for i, (p, a, b) in enumerate(CURVES, start=5):
        curves.cell(i, 1, p)
        curves.cell(i, 2, a)
        curves.cell(i, 3, b)
    P = "'Flow Curves'!$A$5:$A$11"
    A = "'Flow Curves'!$B$5:$B$11"
    B = "'Flow Curves'!$C$5:$C$11"
    curves.column_dimensions["A"].width = 16
    curves.sheet_state = "hidden"

    # ---------- Leak Log ----------
    ws = log
    for col, w in zip("ABCDEFGHIJKLMNOPQRSTU", [9, 15, 10, 9, 10, 10, 10, 10, 12, 12, 12, 12, 10, 14, 10, 13, 13, 3, 16, 14, 26]):
        ws.column_dimensions[col].width = w
    ws["A3"] = company
    ws["A3"].font, ws["A3"].fill = Font(name=FONT, bold=True, size=13, color="0000FF"), YELLOW
    ws["K3"] = "=_xlfn.CONCAT(\"Survey year \",Cost!$D$3)"
    ws["O3"] = "Total Cost Avoidance"
    ws["O3"].font = BOLD
    groups = [("A", "B", "Air CFM and Cost"), ("C", "D", "Carbon Dioxide CFM and Cost"),
              ("E", "F", "Argon CFM and Cost"), ("G", "H", "Helium CFM and Cost"),
              ("I", "J", "Hydrogen CFM and Cost"), ("K", "L", "Nitrogen CFM and Cost"),
              ("M", "N", "All Leaks Identified")]
    gas_of = ["Air", "Carbon Dioxide", "Argon", "Helium", "Hydrogen", "Nitrogen"]
    rng = lambda col: f"${col}${FIRST}:${col}${LAST}"
    for i, (c1, c2, label) in enumerate(groups):
        ws.merge_cells(f"{c1}4:{c2}4")
        ws[f"{c1}4"] = label
        ws[f"{c1}5"], ws[f"{c2}5"] = "CFM", "Cost"
        if i < 6:
            g = gas_of[i]
            ws[f"{c1}6"] = f'=SUMIFS({rng("O")},{rng("B")},"{g}")'
            ws[f"{c2}6"] = f'=SUMIFS({rng("N")},{rng("B")},"{g}")'
        else:  # all leaks: every gas except open blows
            ws[f"{c1}6"] = f'=SUM({rng("O")})-SUMIFS({rng("O")},{rng("B")},"Open Blow")'
            ws[f"{c2}6"] = f'=SUM({rng("N")})-SUMIFS({rng("N")},{rng("B")},"Open Blow")'
    ws["O4"] = "% of Leaks Repaired"
    ws.merge_cells("O4:O5")
    ws["O6"] = (f'=IFERROR(COUNTIFS({rng("M")},"Y",{rng("D")},">0",{rng("B")},"<>Open Blow")/'
                f'COUNTIFS({rng("D")},">0",{rng("B")},"<>Open Blow"),0)')
    ws.merge_cells("P4:Q4")
    ws["P4"] = "Annual Savings (repaired)"
    ws["P5"], ws["Q5"] = "CFM", "Cost"
    ws["P6"] = f'=SUMIFS({rng("O")},{rng("M")},"Y",{rng("B")},"<>Open Blow")'
    ws["Q6"] = f'=SUMIFS({rng("N")},{rng("M")},"Y",{rng("B")},"<>Open Blow")'
    for col in "ABCDEFGHIJKLMNOPQ":
        for rr in (4, 5):
            c = ws[f"{col}{rr}"]
            c.font, c.fill, c.alignment = HEAD, BLACK, CENTER
        c = ws[f"{col}6"]
        c.font = BOLD
        c.alignment = Alignment(horizontal="center")
        c.number_format = "0.0" if col in "ACEGIKMP" else '$#,##0.00'
    ws["O6"].number_format = "0.0%"

    headers = {"A": "Record Number", "B": "Type of Gas", "C": "Pressure at Leak (psi)", "D": "dB Reading",
               "E": "Problem Location", "I": "Problem Description", "M": "Repaired (Y/N)",
               "N": "Identified Leaks Cost Avoidance", "O": "Size of Leak CFM", "P": "Energy Avoidance (kWh)",
               "Q": "Unrepaired from Previous Audit (ft³/min)", "S": "Leaks Identified Summary", "U": "CFM basis / check"}
    for col, h in headers.items():
        ws[f"{col}7"] = h
    ws.merge_cells("E7:H7")
    ws.merge_cells("I7:L7")
    for col in "ABCDEFGHIJKLMNOPQU":
        c = ws[f"{col}7"]
        c.font, c.fill, c.alignment = HEAD, BLACK, CENTER
    ws.row_dimensions[7].height = 42

    HRS = "Cost!$D$16*Cost!$D$17"
    for r in range(FIRST, LAST + 1):
        ws[f"A{r}"] = r - FIRST + 1
        # Helper columns W:Z (hidden): lower curve index, upper index, CFM on each curve.
        ws[f"W{r}"] = f'=IF(OR($D{r}="",$C{r}=""),"",MATCH(MAX(10,MIN(150,$C{r})),{P},1))'
        ws[f"X{r}"] = f'=IF($W{r}="","",MIN($W{r}+1,7))'
        ws[f"Y{r}"] = f'=IF($W{r}="","",INDEX({A},$W{r})*POWER($D{r},INDEX({B},$W{r})))'
        ws[f"Z{r}"] = f'=IF($W{r}="","",INDEX({A},$X{r})*POWER($D{r},INDEX({B},$X{r})))'
        pl, ph = f"INDEX({P},$W{r})", f"INDEX({P},$X{r})"
        pc = f"MAX(10,MIN(150,$C{r}))"
        ws[f"O{r}"] = (f'=IF($W{r}="",0,IF({ph}={pl},$Y{r},'
                       f'$Y{r}+($Z{r}-$Y{r})*({pc}-{pl})/({ph}-{pl})))')
        ws[f"N{r}"] = (f'=IF($O{r}=0,0,$O{r}*60*{HRS}*'
                       f'IFERROR(INDEX({GAS_C},MATCH($B{r},{GAS_N},0)),0)/Cost!$D$5)')
        ws[f"P{r}"] = (f'=IF(OR($B{r}="Air",$B{r}="Open Blow"),'
                       f'$O{r}*60*{HRS}/Cost!$D$5*Cost!$D$13,0)')
        ws[f"U{r}"] = (f'=IF($D{r}="","",IF(ISNA(MATCH($B{r},{GAS_N},0)),"CHECK GAS TYPE",'
                       f'IF(OR($C{r}<10,$C{r}>150),"CHECK PRESSURE: outside 10-150 psi table",'
                       f'IF(OR($D{r}<1,$D{r}>120),"CHECK dB: outside 1-120 table",'
                       f'IF(ISNA(MATCH($C{r},{P},0)),"INTERPOLATED between curves","Table curve")))))')
        ws.merge_cells(f"E{r}:H{r}")
        ws.merge_cells(f"I{r}:L{r}")
        for col in "BCDEIMQ":
            ws[f"{col}{r}"].font = INPUT
        ws[f"N{r}"].number_format = '$#,##0.00'
        ws[f"O{r}"].number_format = "0.00"
        ws[f"P{r}"].number_format = "#,##0"
        for col in "ABCDEFGHIJKLMNOPQU":
            ws[f"{col}{r}"].border = THIN
    for col in "WXYZ":
        ws.column_dimensions[col].hidden = True

    # Example row (format reference). Delete or overwrite on the first real leak.
    ws["B8"], ws["C8"], ws["D8"] = "Air", 100, 30
    ws["E8"], ws["I8"], ws["M8"] = "EXAMPLE · Line 3 filler, FRL inlet", "EXAMPLE · Push-to-connect fitting leaking", "N"
    ws["A8"].comment = Comment("Example row showing the expected format. Overwrite it with the first real leak.", "ALD")

    dv_gas = DataValidation(type="list", formula1=f"={GAS_N}", allow_blank=True)
    dv_rep = DataValidation(type="list", formula1='"Y,N"', allow_blank=True)
    dv_psi = DataValidation(type="decimal", operator="between", formula1="10", formula2="150", allow_blank=True,
                            error="Pressure must be 10-150 psi (UE Systems table range).", showErrorMessage=True)
    for dv, col in ((dv_gas, "B"), (dv_rep, "M"), (dv_psi, "C")):
        ws.add_data_validation(dv)
        dv.add(f"{col}{FIRST}:{col}{LAST}")

    # Summary block (S:T), as in the legacy sheet
    kwh_all = f'SUMIFS({rng("P")},{rng("B")},"Air")'
    kwh_rep = f'SUMIFS({rng("P")},{rng("B")},"Air",{rng("M")},"Y")'
    summary = [
        (8, f"={kwh_all}", "kWh", "#,##0"), (9, f"=S8/({HRS})", "kW", "0.0"), (10, "=S9/0.746", "bhp", "0.0"),
        (12, "Leaks Repaired Summary", None, None),
        (13, f"={kwh_rep}", "kWh", "#,##0"), (14, f"=S13/({HRS})", "kW", "0.0"), (15, "=S14/0.746", "bhp", "0.0"),
        (17, "Open Blows (excluded from leak totals)", None, None),
        (18, f'=SUMIFS({rng("O")},{rng("B")},"Open Blow")', "CFM", "0.0"),
        (19, f'=SUMIFS({rng("N")},{rng("B")},"Open Blow")', "$/yr", '$#,##0'),
        (21, "Rows needing review", None, None),
        (22, f'=COUNTIF({rng("U")},"CHECK*")', "rows", "0"),
        (23, f'=COUNTIF({rng("U")},"INTERPOLATED*")', "interpolated", "0"),
        (24, '=IF(Cost!$D$6=0.13,"SET PLANT $/kWh ON COST SHEET","Rate set")', None, None),
    ]
    for rr, val, unit, fmt in summary:
        ws[f"S{rr}"] = val
        if unit:
            ws[f"T{rr}"] = unit
            ws[f"S{rr}"].number_format = fmt
        else:
            ws[f"S{rr}"].font = BOLD
    ws["S7"].font = BOLD
    ws.conditional_formatting.add("S22", FormulaRule(formula=["S22>0"], fill=REDF))
    ws.conditional_formatting.add("S24", FormulaRule(formula=['S24<>"Rate set"'], fill=REDF))
    ws.conditional_formatting.add(f"U{FIRST}:U{LAST}", FormulaRule(formula=[f'LEFT(U{FIRST},5)="CHECK"'], fill=REDF))
    ws.conditional_formatting.add(f"U{FIRST}:U{LAST}", FormulaRule(formula=[f'LEFT(U{FIRST},5)="INTER"'], fill=AMBER))
    ws.freeze_panes = "B8"

    for sheet in wb.worksheets:
        for row in sheet.iter_rows():
            for c in row:
                if c.value is not None and c.font.name != FONT:
                    c.font = Font(name=FONT, bold=c.font.bold, color=c.font.color, size=c.font.size)
    wb.active = 1
    return wb


def main():
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("-o", "--outdir", default=os.path.join(ROOT, "output"))
    ap.add_argument("--company", default="Company name here")
    ap.add_argument("--rate", type=float, help="plant $/kWh")
    ap.add_argument("--year", type=int)
    a = ap.parse_args()
    wb = build(a.company, a.rate, a.year)
    os.makedirs(a.outdir, exist_ok=True)
    name = "ALD Leak Sheet Template.xlsx" if a.company == "Company name here" else f"{a.company} Leak Sheet.xlsx"
    path = os.path.join(a.outdir, name)
    wb.save(path)
    print(path)


if __name__ == "__main__":
    main()
