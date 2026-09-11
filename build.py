#!/usr/bin/env python3
"""Build Northline_OpEx_ZBB_Pack.xlsx — zero-based OpEx / decision-package pack."""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

yellow = PatternFill("solid", fgColor="FFF2CC")
header_fill = PatternFill("solid", fgColor="1F4E79")
section_fill = PatternFill("solid", fgColor="D6E3F0")
green_fill = PatternFill("solid", fgColor="C6EFCE")
amber_fill = PatternFill("solid", fgColor="FFE699")
red_fill = PatternFill("solid", fgColor="F8CBAD")
tile_fill = PatternFill("solid", fgColor="E9EDF4")
light_gray = PatternFill("solid", fgColor="F5F5F5")
must_fill = PatternFill("solid", fgColor="C6EFCE")
disc_fill = PatternFill("solid", fgColor="FFE699")
input_font = Font(name="Calibri", size=11, color="0000FF")
black = Font(name="Calibri", size=11, color="000000")
header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
title_font = Font(name="Calibri", size=16, bold=True, color="1F4E79")
section_font = Font(name="Calibri", size=12, bold=True, color="1F4E79")
bold = Font(name="Calibri", size=11, bold=True)
bold_black = Font(name="Calibri", size=11, bold=True, color="000000")
italic_grey = Font(name="Calibri", size=10, italic=True, color="666666")
small_grey = Font(name="Calibri", size=9, italic=True, color="666666")
link_font = Font(name="Calibri", size=11, color="0563C1", underline="single")
tile_label = Font(name="Calibri", size=9, color="666666")
tile_value = Font(name="Calibri", size=14, bold=True, color="1F4E79")
thin = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)
money = '_($* #,##0_);_($* (#,##0);_($* "-"??_);_(@_)'
pct = "0.0%"
int_fmt = "#,##0"

ASSUMP = "01_Assumptions"
CC = "02_Cost_Centers"
DP = "03_Decision_Packages"
BR = "04_Savings_Bridge"
DASH = "05_Dashboard"
DD = "06_Data_Dictionary"

# Cost centers: (code, name, owner, prior_yr, request, fte_prior, fte_req, driver_note)
COST_CENTERS = [
    ("CC-110", "Field Sales", "VP Sales", 12800, 14100, 86, 92, "Headcount + travel"),
    ("CC-120", "Trade Marketing", "VP Sales", 9200, 9800, 18, 19, "Shopper programs"),
    ("CC-210", "Brand Marketing", "CMO", 15400, 17200, 24, 26, "Media + agency"),
    ("CC-220", "Digital / E-comm", "CMO", 6100, 7400, 14, 17, "Performance media"),
    ("CC-310", "Distribution / Logistics", "VP Supply", 7600, 8100, 42, 44, "Freight rates"),
    ("CC-410", "Plant Quality", "VP Ops", 4200, 4500, 28, 29, "Lab + audits"),
    ("CC-510", "G&A — Finance", "CFO", 8900, 9200, 38, 39, "Systems + close"),
    ("CC-520", "G&A — HR / People", "CHRO", 5400, 5800, 22, 24, "Talent programs"),
    ("CC-530", "G&A — IT / Digital", "CIO", 11200, 12800, 31, 34, "SaaS + cyber"),
    ("CC-610", "R&D / Innovation", "CTO", 17800, 19100, 48, 51, "Projects + trials"),
    ("CC-710", "Corporate / Legal", "GC", 4800, 5100, 12, 12, "Counsel + insurance"),
    ("CC-810", "Facilities / RE", "CFO", 3600, 3700, 8, 8, "Lease + utilities"),
]

# Decision packages: (pkg_id, cc_code, title, tier, amount, fund_default, rationale)
# tier: Must-Have | Should-Have | Nice-to-Have
PACKAGES = [
    ("DP-01", "CC-110", "Core field sales coverage (base territory)", "Must-Have", 11200, "Y", "Revenue protection"),
    ("DP-02", "CC-110", "New Club channel SE expansion (6 FTEs)", "Should-Have", 1800, "Y", "Growth bet"),
    ("DP-03", "CC-110", "National sales meeting + incentive trip", "Nice-to-Have", 1100, "N", "Deferrable"),
    ("DP-04", "CC-120", "Base trade calendar & broker fees", "Must-Have", 7800, "Y", "Plan delivery"),
    ("DP-05", "CC-120", "Incremental Club shipper program", "Should-Have", 1400, "N", "ROI unclear"),
    ("DP-06", "CC-120", "Holiday display contests", "Nice-to-Have", 600, "N", "Cut"),
    ("DP-07", "CC-210", "Core brand media (TV + social always-on)", "Must-Have", 12800, "Y", "Brand equity"),
    ("DP-08", "CC-210", "New flavor launch campaign", "Should-Have", 3200, "Y", "Innovation support"),
    ("DP-09", "CC-210", "Sponsorship / experiential", "Nice-to-Have", 1200, "N", "Cut"),
    ("DP-10", "CC-220", "Performance media base + tools", "Must-Have", 5200, "Y", "E-comm growth"),
    ("DP-11", "CC-220", "Incremental TikTok / CTV test", "Should-Have", 1400, "Y", "Test & learn"),
    ("DP-12", "CC-220", "Agency retainer uplift", "Nice-to-Have", 800, "N", "Renegotiate"),
    ("DP-13", "CC-310", "Base DC ops + contracted freight", "Must-Have", 7200, "Y", "Service levels"),
    ("DP-14", "CC-310", "Surge capacity / peak overtime", "Should-Have", 600, "N", "Absorb in base"),
    ("DP-15", "CC-310", "Green fleet pilot", "Nice-to-Have", 300, "N", "Defer"),
    ("DP-16", "CC-410", "QA lab + regulatory compliance", "Must-Have", 4000, "Y", "Food safety"),
    ("DP-17", "CC-410", "Extra 3rd-party audits", "Should-Have", 350, "Y", "Customer req"),
    ("DP-18", "CC-410", "New rapid-test equipment", "Nice-to-Have", 150, "N", "Capex alt"),
    ("DP-19", "CC-510", "Core FP&A / controllership", "Must-Have", 8200, "Y", "Controls"),
    ("DP-20", "CC-510", "Planning system enhancement", "Should-Have", 700, "N", "Defer 1 yr"),
    ("DP-21", "CC-510", "External benchmarking study", "Nice-to-Have", 300, "N", "Cut"),
    ("DP-22", "CC-520", "Core HR ops + benefits admin", "Must-Have", 4800, "Y", "People ops"),
    ("DP-23", "CC-520", "Leadership development cohort", "Should-Have", 700, "Y", "Retention"),
    ("DP-24", "CC-520", "Offsite culture summit", "Nice-to-Have", 300, "N", "Cut"),
    ("DP-25", "CC-530", "Run-the-business IT + cyber", "Must-Have", 9800, "Y", "Security"),
    ("DP-26", "CC-530", "ERP module upgrade phase 1", "Should-Have", 2200, "Y", "Efficiency"),
    ("DP-27", "CC-530", "AI assistant licenses (pilot)", "Nice-to-Have", 800, "N", "Pilot later"),
    ("DP-28", "CC-610", "Core R&D pipeline (must projects)", "Must-Have", 14200, "Y", "Pipeline"),
    ("DP-29", "CC-610", "Adjacent category exploration", "Should-Have", 3200, "N", "Focus core"),
    ("DP-30", "CC-610", "University research grants", "Nice-to-Have", 1700, "N", "Cut"),
    ("DP-31", "CC-710", "Legal, insurance, compliance base", "Must-Have", 4800, "Y", "Risk"),
    ("DP-32", "CC-710", "M&A diligence retainer", "Should-Have", 200, "N", "As-needed"),
    ("DP-33", "CC-710", "Board offsite production", "Nice-to-Have", 100, "N", "Cut"),
    ("DP-34", "CC-810", "Lease, utilities, security base", "Must-Have", 3400, "Y", "Facilities"),
    ("DP-35", "CC-810", "HQ refresh / furniture", "Nice-to-Have", 300, "N", "Defer"),
]


def style_input(cell):
    cell.fill = yellow
    cell.font = input_font
    cell.border = thin
    cell.alignment = Alignment(horizontal="center")


def style_formula(cell, key=False):
    cell.font = bold_black if key else black
    cell.border = thin
    cell.alignment = Alignment(horizontal="center")
    if key:
        cell.fill = green_fill


def style_header_cell(cell, value):
    cell.value = value
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = Alignment(horizontal="center", wrap_text=True)
    cell.border = thin


def set_col_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def shade_section(ws, row, until=8):
    ws.cell(row, 2).fill = section_fill
    ws.cell(row, 2).font = section_font
    for c in range(3, until + 1):
        ws.cell(row, c).fill = section_fill


def landscape(ws, fit_height=0):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = fit_height
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.paperSize = ws.PAPERSIZE_TABLOID
    ws.sheet_view.showGridLines = False
    ws.page_setup.horizontalCentered = True
    ws.oddFooter.left.text = "Northline Consumer Products  |  fictional sample  |  $000s"
    ws.oddFooter.right.text = "Page &P of &N"


wb = Workbook()
wb.properties.creator = "Sai Siri Bandaru"
wb.properties.title = "Northline OpEx ZBB Pack"

# ========== 00_Cover ==========
ws = wb.active
ws.title = "00_Cover"
ws.sheet_properties.tabColor = "1F4E79"
set_col_widths(ws, [4, 96])
landscape(ws, fit_height=1)

ws["B2"] = "Northline Consumer Products"
ws["B2"].font = title_font
ws["B3"] = "OpEx zero-based budget pack  ·  Cost centers, decision packages, savings bridge"
ws["B3"].font = section_font
ws["B4"] = "FY26 ZBB cycle  ·  must-have vs discretionary  ·  fictional CPG company  ·  $000s"
ws["B4"].font = italic_grey

ws["B6"] = "What this file is"
ws["B6"].font = bold
ws["B7"] = (
    "A zero-based OpEx pack for Northline's FY26 budget. Cost centers submit prior-year actuals "
    "and a fresh ask. Spend is rebuilt as decision packages (Must-Have / Should-Have / Nice-to-Have). "
    "A fund Y/N toggle and an OpEx target produce a recommended budget and a Prior → Ask → ZBB cuts "
    "→ Target savings bridge for the CFO review."
)
ws["B7"].alignment = Alignment(wrap_text=True)
ws.row_dimensions[7].height = 68

ws["B9"] = "How to use"
ws["B9"].font = bold
ws["B10"] = "1. Open 01_Assumptions. Set FY26 OpEx target and ZBB policy (must-fund floor, discretionary cut bias)."
ws["B11"] = "2. Review prior / request / FTE drivers on 02_Cost_Centers."
ws["B12"] = "3. Fund or defer packages on 03_Decision_Packages (yellow Fund Y/N cells)."
ws["B13"] = "4. Walk Prior → Ask → Cuts → Recommended / Target on 04_Savings_Bridge."
ws["B14"] = "5. Use 05_Dashboard for ask vs target, tier mix, and gap-to-target."

ws["B16"] = "File conventions"
ws["B16"].font = bold
ws["B17"] = "Yellow cells with blue font = inputs. Black font = formulas. Green cells = key outputs."
ws["B17"].fill = yellow
ws["B17"].font = Font(name="Calibri", size=11, color="0000FF", bold=True)
ws["B18"] = "All figures are fictional. No employer data. Dollar figures in $000s."

ws["B20"] = "Portfolio"
ws["B20"].font = bold
ws["B21"] = "Sai Siri Bandaru — Financial Analyst | FP&A | forecasting, variance analysis, Excel"
ws["B22"] = "https://github.com/saisiri-bandaru"
ws["B22"].font = link_font

# ========== 01_Assumptions ==========
ws = wb.create_sheet(ASSUMP)
ws.sheet_properties.tabColor = "F7C948"
set_col_widths(ws, [4, 48, 16, 16, 28, 14])
landscape(ws)
ws.freeze_panes = "C6"

ws["B2"] = "Assumptions"
ws["B2"].font = title_font
ws["B3"] = "Yellow + blue = inputs. Target and ZBB policy drive the savings bridge and dashboard gap."
ws["B3"].font = italic_grey

ws["B5"] = "Control panel"
shade_section(ws, 5, until=4)

ws["B6"] = "Budget year"
ws["C6"] = 2026
style_input(ws["C6"])
ws["C6"].number_format = "0"

ws["B7"] = "Company"
ws["C7"] = "Northline Consumer Products"
style_input(ws["C7"])
ws["C7"].alignment = Alignment(horizontal="left")

ws["B8"] = "FY26 OpEx target ($000s)"
ws["C8"] = 112000
style_input(ws["C8"])
ws["C8"].number_format = money
ws["D8"] = "CFO stretch vs prior"
ws["D8"].font = small_grey

ws["B9"] = "Must-Have fund floor (% of Must $)"
ws["C9"] = 1.00
style_input(ws["C9"])
ws["C9"].number_format = pct
ws["D9"] = "Policy: fund 100% of Must-Have"
ws["D9"].font = small_grey

ws["B10"] = "Should-Have default fund bias"
ws["C10"] = 0.70
style_input(ws["C10"])
ws["C10"].number_format = pct
ws["D10"] = "Informational — actual fund flags on packages tab"
ws["D10"].font = small_grey

ws["B11"] = "Nice-to-Have default fund bias"
ws["C11"] = 0.00
style_input(ws["C11"])
ws["C11"].number_format = pct
ws["D11"] = "Policy: defer Nice-to-Have unless exception"
ws["D11"].font = small_grey

ws["B13"] = "Tier legend"
shade_section(ws, 13, until=4)
style_header_cell(ws.cell(14, 2), "Tier")
style_header_cell(ws.cell(14, 3), "Meaning")
style_header_cell(ws.cell(14, 4), "Default stance")

tiers = [
    ("Must-Have", "Required to run the business / compliance / revenue protection", "Fund"),
    ("Should-Have", "High-ROI growth or efficiency; fund if within target", "Case-by-case"),
    ("Nice-to-Have", "Discretionary / deferrable / low near-term ROI", "Defer"),
]
for r, (t, m, d) in enumerate(tiers, 15):
    ws.cell(r, 2).value = t
    ws.cell(r, 2).border = thin
    ws.cell(r, 2).font = bold_black
    if t == "Must-Have":
        ws.cell(r, 2).fill = must_fill
    elif t == "Nice-to-Have":
        ws.cell(r, 2).fill = disc_fill
    else:
        ws.cell(r, 2).fill = amber_fill
    ws.cell(r, 3).value = m
    ws.cell(r, 3).border = thin
    ws.cell(r, 3).font = small_grey
    ws.cell(r, 4).value = d
    ws.cell(r, 4).border = thin

ws["B19"] = "Roll-up checks (formulas from other tabs)"
shade_section(ws, 19, until=4)
ws["B20"] = "Prior-year OpEx (sum of cost centers)"
ws["C20"] = f"={CC}!E20"
style_formula(ws["C20"], key=True)
ws["C20"].number_format = money

ws["B21"] = "Total ask (sum of cost centers)"
ws["C21"] = f"={CC}!F20"
style_formula(ws["C21"], key=True)
ws["C21"].number_format = money

ws["B22"] = "Package ask (sum of decision packages)"
ws["C22"] = f"={DP}!F42"
style_formula(ws["C22"], key=True)
ws["C22"].number_format = money

ws["B23"] = "Funded packages (recommended)"
ws["C23"] = f"={DP}!H42"
style_formula(ws["C23"], key=True)
ws["C23"].number_format = money

ws["B24"] = "Gap: funded vs target (negative = under target)"
ws["C24"] = "=C23-C8"
style_formula(ws["C24"], key=True)
ws["C24"].number_format = money

ws["B26"] = "Ask should ≈ package total. Fund flags on 03_Decision_Packages are the ZBB decisions."
ws["B26"].font = italic_grey

# ========== 02_Cost_Centers ==========
ws = wb.create_sheet(CC)
ws.sheet_properties.tabColor = "5B9BD5"
set_col_widths(ws, [4, 12, 26, 14, 12, 12, 12, 10, 10, 12, 12, 22])
landscape(ws)
ws.freeze_panes = "C6"

ws["B2"] = "Cost centers"
ws["B2"].font = title_font
ws["B3"] = "Prior-year actuals, FY26 request, and headcount drivers. Yellow = editable inputs."
ws["B3"].font = italic_grey

ws["B5"] = "Cost-center ledger ($000s)"
shade_section(ws, 5, until=11)

headers = [
    "Code", "Cost center", "Owner", "Prior FY", "FY26 ask", "Δ $", "Δ %",
    "FTE prior", "FTE ask", "Δ FTE", "Driver note",
]
for i, h in enumerate(headers, 2):
    style_header_cell(ws.cell(6, i), h)

for i, (code, name, owner, prior, ask, fte_p, fte_a, note) in enumerate(COST_CENTERS):
    r = 7 + i
    ws.cell(r, 2).value = code
    ws.cell(r, 2).border = thin
    ws.cell(r, 2).font = bold_black
    ws.cell(r, 3).value = name
    ws.cell(r, 3).border = thin
    ws.cell(r, 4).value = owner
    ws.cell(r, 4).border = thin
    ws.cell(r, 4).font = small_grey
    ws.cell(r, 5).value = prior
    style_input(ws.cell(r, 5))
    ws.cell(r, 5).number_format = money
    ws.cell(r, 6).value = ask
    style_input(ws.cell(r, 6))
    ws.cell(r, 6).number_format = money
    ws.cell(r, 7).value = f"=F{r}-E{r}"
    style_formula(ws.cell(r, 7))
    ws.cell(r, 7).number_format = money
    ws.cell(r, 8).value = f"=F{r}/E{r}-1"
    style_formula(ws.cell(r, 8))
    ws.cell(r, 8).number_format = pct
    ws.cell(r, 9).value = fte_p
    style_input(ws.cell(r, 9))
    ws.cell(r, 9).number_format = int_fmt
    ws.cell(r, 10).value = fte_a
    style_input(ws.cell(r, 10))
    ws.cell(r, 10).number_format = int_fmt
    ws.cell(r, 11).value = f"=J{r}-I{r}"
    style_formula(ws.cell(r, 11))
    ws.cell(r, 11).number_format = int_fmt
    ws.cell(r, 12).value = note
    ws.cell(r, 12).border = thin
    ws.cell(r, 12).font = small_grey

# Totals row 19? 7+12-1=18, so total at 19... wait 7+11=18 for last, total at 19
# Actually 12 centers: rows 7-18, total row 19. But C20 referenced in Assumptions - I'll use row 20 for total with blank 19 or put total at 20.

# Put total at row 20 (row 19 blank for spacing) to match Assumptions refs C20/D20
ws["B19"] = ""
ws["B20"] = "TOTAL"
ws["B20"].font = bold_black
ws["B20"].fill = section_fill
ws["B20"].border = thin
ws["C20"] = ""
ws["C20"].fill = section_fill
ws["C20"].border = thin
ws["D20"] = ""
ws["D20"].fill = section_fill
ws["D20"].border = thin
ws["E20"] = "=SUM(E7:E18)"
style_formula(ws["E20"], key=True)
ws["E20"].number_format = money
# Wait - Assumptions refs C20 for prior and D20 for ask. I need to fix either Assumptions or layout.
# Simpler: put Prior total in C20 and Ask in D20 by restructuring columns... 
# Actually Assumptions says:
# C20 = CC!C20 prior
# C21 = CC!D20 ask
# But my columns are E=prior, F=ask. Let me fix Assumptions to use E20 and F20.

ws["F20"] = "=SUM(F7:F18)"
style_formula(ws["F20"], key=True)
ws["F20"].number_format = money
ws["G20"] = "=F20-E20"
style_formula(ws["G20"], key=True)
ws["G20"].number_format = money
ws["H20"] = "=F20/E20-1"
style_formula(ws["H20"])
ws["H20"].number_format = pct
ws["I20"] = "=SUM(I7:I18)"
style_formula(ws["I20"])
ws["I20"].number_format = int_fmt
ws["J20"] = "=SUM(J7:J18)"
style_formula(ws["J20"])
ws["J20"].number_format = int_fmt
ws["K20"] = "=J20-I20"
style_formula(ws["K20"])
ws["K20"].number_format = int_fmt

# Must vs disc summary placeholders via package rolls (linked later conceptually)
ws["B22"] = "Ask vs prior"
ws["B22"].font = bold
ws["B23"] = "YoY ask growth $"
ws["C23"] = "=G20"
style_formula(ws["C23"], key=True)
ws["C23"].number_format = money
ws["B24"] = "YoY ask growth %"
ws["C24"] = "=H20"
style_formula(ws["C24"], key=True)
ws["C24"].number_format = pct

ws["B26"] = "Note: package rebuild on 03_Decision_Packages should reconcile to FY26 ask (± rounding)."
ws["B26"].font = italic_grey

# ========== 03_Decision_Packages ==========
ws = wb.create_sheet(DP)
ws.sheet_properties.tabColor = "70AD47"
set_col_widths(ws, [4, 10, 12, 44, 14, 12, 10, 12, 28])
landscape(ws)
ws.freeze_panes = "C6"

ws["B2"] = "Decision packages"
ws["B2"].font = title_font
ws["B3"] = "ZBB rebuild of OpEx. Toggle Fund (Y/N) on yellow cells. Funded $ rolls to the savings bridge."
ws["B3"].font = italic_grey

ws["B5"] = "Package register ($000s)"
shade_section(ws, 5, until=8)

headers = [
    "Pkg ID", "Cost center", "Package title", "Tier", "Ask $", "Fund (Y/N)", "Funded $", "Rationale",
]
for i, h in enumerate(headers, 2):
    style_header_cell(ws.cell(6, i), h)

dv_fund = DataValidation(type="list", formula1='"Y,N"', allow_blank=False)
ws.add_data_validation(dv_fund)

for i, (pid, cc, title, tier, amt, fund, rationale) in enumerate(PACKAGES):
    r = 7 + i
    ws.cell(r, 2).value = pid
    ws.cell(r, 2).border = thin
    ws.cell(r, 2).font = bold_black
    ws.cell(r, 3).value = cc
    ws.cell(r, 3).border = thin
    ws.cell(r, 4).value = title
    ws.cell(r, 4).border = thin
    ws.cell(r, 4).alignment = Alignment(wrap_text=True)
    ws.cell(r, 5).value = tier
    ws.cell(r, 5).border = thin
    ws.cell(r, 5).font = bold_black
    if tier == "Must-Have":
        ws.cell(r, 5).fill = must_fill
    elif tier == "Nice-to-Have":
        ws.cell(r, 5).fill = disc_fill
    else:
        ws.cell(r, 5).fill = amber_fill
    ws.cell(r, 6).value = amt
    style_input(ws.cell(r, 6))
    ws.cell(r, 6).number_format = money
    ws.cell(r, 7).value = fund
    style_input(ws.cell(r, 7))
    dv_fund.add(ws.cell(r, 7))
    ws.cell(r, 8).value = f'=IF(G{r}="Y",F{r},0)'
    style_formula(ws.cell(r, 8))
    ws.cell(r, 8).number_format = money
    ws.cell(r, 9).value = rationale
    ws.cell(r, 9).border = thin
    ws.cell(r, 9).font = small_grey

# Last package row = 7+35-1 = 41
last = 6 + len(PACKAGES)  # 41
tot = last + 1  # 42

ws.cell(tot, 2).value = "TOTAL"
ws.cell(tot, 2).font = bold_black
ws.cell(tot, 2).fill = section_fill
ws.cell(tot, 2).border = thin
for c in range(3, 6):
    ws.cell(tot, c).fill = section_fill
    ws.cell(tot, c).border = thin
ws.cell(tot, 6).value = f"=SUM(F7:F{last})"
style_formula(ws.cell(tot, 6), key=True)
ws.cell(tot, 6).number_format = money
ws.cell(tot, 7).fill = section_fill
ws.cell(tot, 7).border = thin
ws.cell(tot, 8).value = f"=SUM(H7:H{last})"
style_formula(ws.cell(tot, 8), key=True)
ws.cell(tot, 8).number_format = money
ws.cell(tot, 9).value = "Ask vs funded"
ws.cell(tot, 9).font = small_grey
ws.cell(tot, 9).fill = section_fill
ws.cell(tot, 9).border = thin

# Tier rollups
ws["B44"] = "Tier rollup"
shade_section(ws, 44, until=5)
style_header_cell(ws.cell(45, 2), "Tier")
style_header_cell(ws.cell(45, 3), "Ask $")
style_header_cell(ws.cell(45, 4), "Funded $")
style_header_cell(ws.cell(45, 5), "Cut / deferred $")
style_header_cell(ws.cell(45, 6), "% funded")

for r, tier in enumerate(["Must-Have", "Should-Have", "Nice-to-Have"], 46):
    ws.cell(r, 2).value = tier
    ws.cell(r, 2).border = thin
    ws.cell(r, 2).font = bold_black
    if tier == "Must-Have":
        ws.cell(r, 2).fill = must_fill
    elif tier == "Nice-to-Have":
        ws.cell(r, 2).fill = disc_fill
    else:
        ws.cell(r, 2).fill = amber_fill
    ws.cell(r, 3).value = f'=SUMIF($E$7:$E${last},B{r},$F$7:$F${last})'
    style_formula(ws.cell(r, 3))
    ws.cell(r, 3).number_format = money
    ws.cell(r, 4).value = f'=SUMIF($E$7:$E${last},B{r},$H$7:$H${last})'
    style_formula(ws.cell(r, 4))
    ws.cell(r, 4).number_format = money
    ws.cell(r, 5).value = f"=C{r}-D{r}"
    style_formula(ws.cell(r, 5))
    ws.cell(r, 5).number_format = money
    ws.cell(r, 6).value = f"=IF(C{r}=0,0,D{r}/C{r})"
    style_formula(ws.cell(r, 6))
    ws.cell(r, 6).number_format = pct

ws["B50"] = "Cut / deferred total"
ws["C50"] = f"=E{tot}-G{tot}"  # Wait columns: F=ask total at tot col6, H=funded at tot col8
# Fix: Ask total is F42 (=col 6), Funded is H42 (=col 8)
# In openpyxl cell tot,6 and tot,8
ws["C50"] = f"=F{tot}-H{tot}"
style_formula(ws["C50"], key=True)
ws["C50"].number_format = money
ws["B50"].border = thin

ws["B52"] = f"Package rows 7–{last}. Change Fund Y/N to rebuild recommended OpEx instantly."
ws["B52"].font = italic_grey

# ========== 04_Savings_Bridge ==========
ws = wb.create_sheet(BR)
ws.sheet_properties.tabColor = "ED7D31"
set_col_widths(ws, [4, 36, 14, 14, 40])
landscape(ws)
ws.freeze_panes = "C6"

ws["B2"] = "Savings bridge"
ws["B2"].font = title_font
ws["B3"] = "Prior → Ask → ZBB cuts → Recommended budget, vs CFO target."
ws["B3"].font = italic_grey

ws["B5"] = "OpEx walk ($000s)"
shade_section(ws, 5, until=4)
style_header_cell(ws.cell(6, 2), "Step")
style_header_cell(ws.cell(6, 3), "Amount")
style_header_cell(ws.cell(6, 4), "Running")
style_header_cell(ws.cell(6, 5), "Notes")

ws["B7"] = "Prior-year OpEx"
ws["C7"] = f"={CC}!E20"
ws["D7"] = "=C7"
ws["E7"] = "Starting point (FY25 actuals)"
style_formula(ws["C7"], key=True)
style_formula(ws["D7"], key=True)
ws["C7"].number_format = money
ws["D7"].number_format = money
ws["B7"].border = thin
ws["E7"].font = small_grey
ws["E7"].border = thin

ws["B8"] = "Gross ask uplift"
ws["C8"] = f"={CC}!G20"
ws["D8"] = "=D7+C8"
ws["E8"] = "Cost-center requests − prior"
style_formula(ws["C8"])
style_formula(ws["D8"])
ws["C8"].number_format = money
ws["D8"].number_format = money
ws["B8"].border = thin
ws["E8"].font = small_grey
ws["E8"].border = thin

ws["B9"] = "FY26 ask (pre-ZBB)"
ws["C9"] = f"={CC}!F20"
ws["D9"] = "=D8"
ws["E9"] = "Should equal prior + uplift"
style_formula(ws["C9"], key=True)
style_formula(ws["D9"], key=True)
ws["C9"].number_format = money
ws["D9"].number_format = money
ws["B9"].border = thin
ws["E9"].font = small_grey
ws["E9"].border = thin

ws["B10"] = "ZBB cuts — Must-Have deferred"
ws["C10"] = f"=-{DP}!E46"
ws["D10"] = "=D9+C10"
ws["E10"] = "Should be ~0 under policy"
style_formula(ws["C10"])
style_formula(ws["D10"])
ws["C10"].number_format = money
ws["D10"].number_format = money
ws["B10"].border = thin
ws["E10"].font = small_grey
ws["E10"].border = thin

ws["B11"] = "ZBB cuts — Should-Have deferred"
ws["C11"] = f"=-{DP}!E47"
ws["D11"] = "=D10+C11"
ws["E11"] = "Case-by-case deferrals"
style_formula(ws["C11"])
style_formula(ws["D11"])
ws["C11"].number_format = money
ws["D11"].number_format = money
ws["B11"].border = thin
ws["E11"].font = small_grey
ws["E11"].border = thin

ws["B12"] = "ZBB cuts — Nice-to-Have deferred"
ws["C12"] = f"=-{DP}!E48"
ws["D12"] = "=D11+C12"
ws["E12"] = "Discretionary deferrals"
style_formula(ws["C12"])
style_formula(ws["D12"])
ws["C12"].number_format = money
ws["D12"].number_format = money
ws["B12"].border = thin
ws["E12"].font = small_grey
ws["E12"].border = thin

ws["B13"] = "Recommended (funded packages)"
ws["C13"] = f"={DP}!H42"
ws["D13"] = "=D12"
ws["E13"] = "Running should ≈ funded total"
style_formula(ws["C13"], key=True)
style_formula(ws["D13"], key=True)
ws["C13"].number_format = money
ws["D13"].number_format = money
ws["B13"].border = thin
ws["E13"].font = small_grey
ws["E13"].border = thin

ws["B14"] = "Bridge check (should be ~0)"
ws["C14"] = "=D13-C13"
style_formula(ws["C14"], key=True)
ws["C14"].number_format = money
ws["B14"].border = thin

ws["B16"] = "Target reconciliation"
shade_section(ws, 16, until=4)
ws["B17"] = "CFO OpEx target"
ws["C17"] = f"={ASSUMP}!C8"
style_formula(ws["C17"], key=True)
ws["C17"].number_format = money
ws["B17"].border = thin

ws["B18"] = "Recommended vs target"
ws["C18"] = "=C13-C17"
style_formula(ws["C18"], key=True)
ws["C18"].number_format = money
ws["D18"] = '=IF(C18>0,"Over target — more cuts needed","At / under target")'
style_formula(ws["D18"])
ws["B18"].border = thin

ws["B19"] = "Total ZBB savings vs ask"
ws["C19"] = f"={DP}!C50"
style_formula(ws["C19"], key=True)
ws["C19"].number_format = money
ws["B19"].border = thin

ws["B20"] = "Savings vs prior"
ws["C20"] = "=C7-C13"
style_formula(ws["C20"], key=True)
ws["C20"].number_format = money
ws["D20"] = "Positive = OpEx down vs prior"
ws["D20"].font = small_grey
ws["B20"].border = thin

# Chart helper data
ws["B22"] = "Chart data — bridge steps"
shade_section(ws, 22, until=3)
style_header_cell(ws.cell(23, 2), "Step")
style_header_cell(ws.cell(23, 3), "Value")
for i, (lab, src) in enumerate(
    [
        ("Prior", "C7"),
        ("Ask uplift", "C8"),
        ("Must cuts", "C10"),
        ("Should cuts", "C11"),
        ("Nice cuts", "C12"),
        ("Recommended", "C13"),
        ("Target", "C17"),
    ],
    24,
):
    ws.cell(i, 2).value = lab
    ws.cell(i, 2).border = thin
    ws.cell(i, 3).value = f"={src}"
    style_formula(ws.cell(i, 3))
    ws.cell(i, 3).number_format = money

# ========== 05_Dashboard ==========
ws = wb.create_sheet(DASH)
ws.sheet_properties.tabColor = "7030A0"
set_col_widths(ws, [4, 22, 14, 14, 14, 14, 14, 14, 14])
landscape(ws, fit_height=1)

ws["B2"] = "Dashboard"
ws["B2"].font = title_font
ws["B3"] = "Ask vs target, tier mix, and ZBB outcome for the CFO review."
ws["B3"].font = italic_grey

ws["B5"] = "Headline tiles ($000s)"
shade_section(ws, 5, until=3)

tiles = [
    (6, "Prior OpEx", f"={CC}!E20"),
    (7, "FY26 ask", f"={CC}!F20"),
    (8, "Recommended", f"={DP}!H42"),
    (9, "CFO target", f"={ASSUMP}!C8"),
    (10, "Gap vs target", f"={BR}!C18"),
    (11, "ZBB savings vs ask", f"={DP}!C50"),
]
for row, lab, formul in tiles:
    ws.cell(row, 2).value = lab
    ws.cell(row, 2).font = tile_label
    ws.cell(row, 2).fill = tile_fill
    ws.cell(row, 2).border = thin
    ws.cell(row, 3).value = formul
    ws.cell(row, 3).font = tile_value
    ws.cell(row, 3).fill = tile_fill
    ws.cell(row, 3).border = thin
    ws.cell(row, 3).number_format = money
    ws.cell(row, 3).alignment = Alignment(horizontal="center")

ws["E5"] = "Tier mix (funded $)"
shade_section(ws, 5, until=7)
# fix E5
ws["E5"].fill = section_fill
ws["E5"].font = section_font
for c in range(6, 8):
    ws.cell(5, c).fill = section_fill

style_header_cell(ws.cell(6, 5), "Tier")
style_header_cell(ws.cell(6, 6), "Funded $")
style_header_cell(ws.cell(6, 7), "% of funded")

for r, tier, src in [
    (7, "Must-Have", f"={DP}!D46"),
    (8, "Should-Have", f"={DP}!D47"),
    (9, "Nice-to-Have", f"={DP}!D48"),
]:
    ws.cell(r, 5).value = tier
    ws.cell(r, 5).border = thin
    ws.cell(r, 6).value = src
    style_formula(ws.cell(r, 6))
    ws.cell(r, 6).number_format = money
    ws.cell(r, 7).value = f"=IF($C$8=0,0,F{r}/$C$8)"
    # Better: % of recommended funded
    ws.cell(r, 7).value = f"=IF($C$8=0,0,F{r}/$C$8)"
    # C8 is FY26 ask tile - wrong. Use recommended C8? Wait C8 is ask. Recommended is C8? 
    # tiles: C6 prior, C7 ask, C8 recommended, C9 target
    ws.cell(r, 7).value = f"=IF($C$8=0,0,F{r}/$C$8)"
    style_formula(ws.cell(r, 7))
    ws.cell(r, 7).number_format = pct

# Fix % of funded to use recommended (C8)
# C8 = Recommended — correct!

ws["B13"] = "Ask vs target summary"
shade_section(ws, 13, until=4)
style_header_cell(ws.cell(14, 2), "Metric")
style_header_cell(ws.cell(14, 3), "Amount")
style_header_cell(ws.cell(14, 4), "% of ask")

rows_sum = [
    (15, "Ask", f"={CC}!F20", 1),
    (16, "Funded / recommended", f"={DP}!H42", f"=C16/C15"),
    (17, "Deferred / cut", f"={DP}!C50", f"=C17/C15"),
    (18, "Target", f"={ASSUMP}!C8", f"=C18/C15"),
    (19, "Over/(under) target", f"={BR}!C18", f"=C19/C15"),
]
for row, lab, formul, pctf in rows_sum:
    ws.cell(row, 2).value = lab
    ws.cell(row, 2).border = thin
    ws.cell(row, 3).value = formul
    style_formula(ws.cell(row, 3), key=(row in (16, 19)))
    ws.cell(row, 3).number_format = money
    if pctf == 1:
        ws.cell(row, 4).value = 1
    else:
        ws.cell(row, 4).value = pctf
    style_formula(ws.cell(row, 4))
    ws.cell(row, 4).number_format = pct

# Bar chart ask / recommended / target
ws["B21"] = "Comparison"
style_header_cell(ws.cell(21, 2), "Item")
style_header_cell(ws.cell(21, 3), "$000s")
ws["B22"] = "Prior"
ws["C22"] = f"={CC}!E20"
ws["B23"] = "Ask"
ws["C23"] = f"={CC}!F20"
ws["B24"] = "Recommended"
ws["C24"] = f"={DP}!H42"
ws["B25"] = "Target"
ws["C25"] = f"={ASSUMP}!C8"
for r in range(22, 26):
    ws.cell(r, 2).border = thin
    style_formula(ws.cell(r, 3), key=(r in (24, 25)))
    ws.cell(r, 3).number_format = money

chart = BarChart()
chart.type = "col"
chart.title = "Prior / Ask / Recommended / Target ($000s)"
chart.y_axis.title = "$000s"
chart.style = 10
data = Reference(ws, min_col=3, min_row=21, max_col=3, max_row=25)
cats = Reference(ws, min_col=2, min_row=22, max_col=2, max_row=25)
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)
chart.shape = 4
chart.width = 14
chart.height = 8
chart.legend = None
ws.add_chart(chart, "E13")

# Pie of funded by tier
pie = PieChart()
pie.title = "Funded $ by tier"
labels = Reference(ws, min_col=5, min_row=7, max_row=9)
pdata = Reference(ws, min_col=6, min_row=6, max_row=9)
pie.add_data(pdata, titles_from_data=True)
pie.set_categories(labels)
pie.dataLabels = DataLabelList()
pie.dataLabels.showPercent = True
pie.dataLabels.showVal = False
pie.dataLabels.showCatName = False
pie.width = 10
pie.height = 8
ws.add_chart(pie, "E28")

ws["B40"] = "Screen-share test: flip a Should-Have Fund from Y→N on 03_Decision_Packages — Recommended and gap-to-target update; bridge cuts deepen."
ws["B40"].font = italic_grey

# ========== 06_Data_Dictionary ==========
ws = wb.create_sheet(DD)
ws.sheet_properties.tabColor = "7F7F7F"
set_col_widths(ws, [4, 28, 18, 58, 16])
landscape(ws)
ws.freeze_panes = "C5"

ws["B2"] = "Data dictionary"
ws["B2"].font = title_font
ws["B3"] = "Field definitions for inheriting analysts. Amounts in $000s unless noted."
ws["B3"].font = italic_grey

headers = ["Field", "Tab", "Definition", "Type"]
for i, h in enumerate(headers, 2):
    style_header_cell(ws.cell(5, i), h)

entries = [
    ("OpEx target", ASSUMP, "CFO FY26 OpEx ceiling ($000s).", "Input"),
    ("Must-Have fund floor", ASSUMP, "Policy % of Must-Have dollars expected to be funded.", "Input"),
    ("Tier", ASSUMP + " / " + DP, "Must-Have / Should-Have / Nice-to-Have ranking.", "Input"),
    ("Prior FY", CC, "Prior-year OpEx actual by cost center ($000s).", "Input"),
    ("FY26 ask", CC, "Cost-center requested budget before ZBB decisions.", "Input"),
    ("FTE prior / ask", CC, "Headcount drivers supporting the ask.", "Input"),
    ("Package ask $", DP, "Dollar size of each decision package.", "Input"),
    ("Fund (Y/N)", DP, "ZBB decision: include package in recommended budget.", "Input"),
    ("Funded $", DP, "Ask $ if Fund=Y, else 0.", "Formula"),
    ("Tier rollup", DP, "Ask / funded / deferred by Must / Should / Nice.", "Formula"),
    ("Gross ask uplift", BR, "Ask − Prior (pre-ZBB growth).", "Formula"),
    ("ZBB cuts by tier", BR, "Deferred package $ by tier (negative steps).", "Formula"),
    ("Recommended", BR, "Sum of funded packages.", "Formula"),
    ("Gap vs target", BR + " / " + DASH, "Recommended − CFO target (positive = over).", "Formula"),
    ("ZBB savings vs ask", BR, "Ask − Recommended.", "Formula"),
]

for r, (field, tab, defin, typ) in enumerate(entries, 6):
    ws.cell(r, 2).value = field
    ws.cell(r, 2).border = thin
    ws.cell(r, 2).font = bold_black if typ == "Input" else black
    ws.cell(r, 3).value = tab
    ws.cell(r, 3).border = thin
    ws.cell(r, 3).font = small_grey
    ws.cell(r, 4).value = defin
    ws.cell(r, 4).border = thin
    ws.cell(r, 4).alignment = Alignment(wrap_text=True)
    ws.cell(r, 5).value = typ
    ws.cell(r, 5).border = thin
    if typ == "Input":
        ws.cell(r, 5).fill = yellow
        ws.cell(r, 5).font = input_font
    else:
        ws.cell(r, 5).font = black

ws["B23"] = "Fictional company and sample data for portfolio demonstration only. No VBA."
ws["B23"].font = italic_grey

out = Path(__file__).resolve().parent / "Northline_OpEx_ZBB_Pack.xlsx"
wb.save(out)
print(f"Wrote {out}")
print(f"Decision packages: {len(PACKAGES)}, last row {6+len(PACKAGES)}, total row {7+len(PACKAGES)}")
