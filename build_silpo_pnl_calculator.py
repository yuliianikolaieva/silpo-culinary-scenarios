from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter


OUTPUT = Path(__file__).parent / "Silpo_2027_DK_Resto_Stores_PnL_Calculator.xlsx"

GREEN = "0B8F50"
DARK = "17231E"
LIGHT_GREEN = "EDF7F1"
YELLOW = "FFF2CC"
AMBER = "FFF7E7"
RED = "FFF1EE"
LINE = "D9E3DD"
MUTED = "64726A"
WHITE = "FFFFFF"

thin = Side(style="thin", color=LINE)
border = Border(left=thin, right=thin, top=thin, bottom=thin)


def style_title(ws, cell, text, width=14):
    ws.merge_cells(start_row=cell.row, start_column=cell.column, end_row=cell.row, end_column=cell.column + width)
    cell.value = text
    cell.fill = PatternFill("solid", fgColor=DARK)
    cell.font = Font(name="Aptos Display", size=18, bold=True, color=WHITE)
    cell.alignment = Alignment(vertical="center")
    ws.row_dimensions[cell.row].height = 34


def section(ws, row, text, end_col=6):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=end_col)
    c = ws.cell(row, 1, text)
    c.fill = PatternFill("solid", fgColor=GREEN)
    c.font = Font(name="Aptos", size=11, bold=True, color=WHITE)
    c.alignment = Alignment(vertical="center")
    ws.row_dimensions[row].height = 22


def header_row(ws, row, values):
    for col, value in enumerate(values, 1):
        c = ws.cell(row, col, value)
        c.fill = PatternFill("solid", fgColor=DARK)
        c.font = Font(name="Aptos", size=10, bold=True, color=WHITE)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = border
    ws.row_dimensions[row].height = 30


def normal(ws, cell, value=None, fill=None, bold=False, align=None):
    if value is not None:
        cell.value = value
    if fill:
        cell.fill = PatternFill("solid", fgColor=fill)
    cell.font = Font(name="Aptos", size=10, bold=bold, color=DARK)
    cell.alignment = Alignment(vertical="center", wrap_text=True, horizontal=align or "left")
    cell.border = border
    return cell


def input_cell(ws, row, col, value, fmt=None):
    c = normal(ws, ws.cell(row, col), value, YELLOW, align="right")
    if fmt:
        c.number_format = fmt
    return c


def formula_cell(ws, row, col, formula, fmt=None, fill=None, bold=False):
    c = normal(ws, ws.cell(row, col), formula, fill, bold, "right")
    if fmt:
        c.number_format = fmt
    return c


def set_widths(ws, widths):
    for col, width in widths.items():
        ws.column_dimensions[col].width = width


def build_readme(wb):
    ws = wb.active
    ws.title = "Read me"
    style_title(ws, ws["A1"], "Silpo 2027 — DK, Resto & Stores P&L calculator", 9)
    ws["A3"] = "Як користуватися"
    ws["A3"].font = Font(name="Aptos", size=13, bold=True, color=GREEN)
    items = [
        "1. У вкладці Inputs змінюйте лише жовті клітинки: обсяг, ставки, DF campaign, Bolt+, Public Launch та Sponsored Listing.",
        "2. У Scenarios редагуйте ставки та прапорці інвестицій. Зелений рядок — базовий поточний сценарій.",
        "3. У P&L by vertical оберіть сценарій у жовтій клітинці B3. Вкладка показує DK, resto, stores та portfolio поруч.",
        "4. CP L1 = commission revenue − direct costs before demand incentives. CP L2 = CP L1 − existing demand incentives.",
        "5. Final P&L after cash investment окремо віднімає DF campaign, Bolt+, Public Launch cash, Sponsored Listing discount та інші cash costs.",
        "6. In-kind marketing показано окремо: це інвестиція Bolt, але не зменшує cash P&L.",
    ]
    for i, item in enumerate(items, 5):
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=9)
        c = ws.cell(i, 1, item)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        c.font = Font(name="Aptos", size=11, color=DARK)
        ws.row_dimensions[i].height = 32
    ws["A13"] = "База моделі"
    ws["A13"].font = Font(name="Aptos", size=13, bold=True, color=GREEN)
    ws.merge_cells("A14:J16")
    ws["A14"] = (
        "DK/resto: 12-місячна база зі вкладки «Спільна ставка» — Dark kitchen €905.3k GMV / 58,154 orders; "
        "Resto €32.0k / 1,982 orders. Stores: 2× scenario — €855.6k GMV / 61,220 orders. "
        "Historical unit economics використані як базові, а зміна ставки впливає на commission revenue."
    )
    ws["A14"].alignment = Alignment(wrap_text=True, vertical="top")
    ws["A14"].fill = PatternFill("solid", fgColor=LIGHT_GREEN)
    ws["A14"].border = border
    set_widths(ws, {"A": 34, "B": 16, "C": 16, "D": 16, "E": 16, "F": 16, "G": 16, "H": 16, "I": 16, "J": 16})
    ws.sheet_view.showGridLines = False


def build_inputs(wb):
    ws = wb.create_sheet("Inputs")
    style_title(ws, ws["A1"], "Inputs — editable assumptions", 6)
    ws.merge_cells("A3:G3")
    ws["A3"] = "Yellow cells are inputs. White cells contain formulas and should not be overwritten."
    ws["A3"].font = Font(name="Aptos", italic=True, color=MUTED)

    section(ws, 5, "Global incremental investment assumptions", 4)
    header_row(ws, 6, ["Input", "Value", "Unit / calculation", "Comment"])
    globals_ = [
        ("DF campaign cost · stores", 28000, "€ / year", "Cash cost; editable"),
        ("Bolt+ duration", 6 / 12, "% of year", "6 months = 50%"),
        ("Bolt+ share · DK/resto", 0.248, "% of orders", "Historical planning assumption"),
        ("Bolt+ share · stores", 0.18, "% of orders", "Planning assumption"),
        ("DF discount per Bolt+ order", 1.96, "€ / order", "Discount only; SF/SOF not discounted"),
        ("Public Launch cash", 10000, "€ / year", "Allocated by GMV"),
        ("In-kind marketing", 11200, "€ / year", "Allocated by GMV; not deducted from cash P&L"),
        ("Sponsored Listing discount", 0, "€ / year", "Manual input; allocated by GMV"),
        ("Other cash investment", 0, "€ / year", "Manual input; allocated by GMV"),
    ]
    for row, (name, value, unit, comment) in enumerate(globals_, 7):
        normal(ws, ws.cell(row, 1), name)
        fmt = "0.0%" if "%" in unit else '€#,##0'
        input_cell(ws, row, 2, value, fmt)
        normal(ws, ws.cell(row, 3), unit)
        normal(ws, ws.cell(row, 4), comment)

    section(ws, 18, "Annual base P&L by vertical", 6)
    header_row(ws, 19, ["Metric", "Dark kitchen", "Resto", "Stores", "Portfolio", "Source / note"])
    metrics = [
        ("Annual delivered orders", 58154, 1982, 61220, "orders"),
        ("Annual GMV", 905300, 32000, 855600, "€"),
        ("Current commission rate", 0.2243, 0.2503, 0.12, "%"),
        ("Current commission revenue", "=B21*B22", "=C21*C22", "=D21*D22", "€"),
        ("Direct cost before demand incentives", 53500, 2100, 59472, "€; commission − historical CP L1"),
        ("Historical CP L1", "=B23-B24", "=C23-C24", "=D23-D24", "€"),
        ("Existing demand incentives", 96000, 3000, 0, "€"),
        ("Historical CP L2", "=B25-B26", "=C25-C26", "=D25-D26", "€"),
        ("Bolt+ cost", "=B20*$B$9*$B$11*$B$8", "=C20*$B$9*$B$11*$B$8", "=D20*$B$10*$B$11*$B$8", "€; formula"),
        ("Public Launch allocation", "=B21/SUM($B$21:$D$21)", "=C21/SUM($B$21:$D$21)", "=D21/SUM($B$21:$D$21)", "% of portfolio GMV"),
    ]
    for row, (metric, dk, resto, stores, note) in enumerate(metrics, 20):
        normal(ws, ws.cell(row, 1), metric, bold=metric in {"Annual GMV", "Historical CP L1", "Historical CP L2"})
        for col, value in enumerate([dk, resto, stores], 2):
            is_input = row in {20, 21, 22, 24, 26}
            fmt = "0.00%" if row in {22, 29} else ('€#,##0' if row not in {20} else '#,##0')
            if isinstance(value, str) and value.startswith("="):
                formula_cell(ws, row, col, value, fmt)
            elif is_input:
                input_cell(ws, row, col, value, fmt)
            else:
                normal(ws, ws.cell(row, col), value, None, False, "right")
                ws.cell(row, col).number_format = fmt
        if row in {20, 21, 23, 24, 25, 26, 27, 28}:
            formula = f"=SUM(B{row}:D{row})"
            fmt = "#,##0" if row == 20 else ("0.00%" if row in {22, 29} else '€#,##0')
            formula_cell(ws, row, 5, formula, fmt, LIGHT_GREEN if row in {21, 25, 27} else None, row in {21, 25, 27})
        else:
            formula_cell(ws, row, 5, f"=SUM(B{row}:D{row})", '€#,##0')
        normal(ws, ws.cell(row, 6), note)

    section(ws, 32, "Calculation notes", 6)
    ws.merge_cells("A33:F35")
    ws["A33"] = (
        "Current commission rates, direct costs and historical demand incentives are editable reference inputs. "
        "The calculator assumes direct costs before demand incentives remain unchanged when commission changes. "
        "DF campaign applies to stores only. Bolt+ and Public Launch can be switched on/off per scenario."
    )
    ws["A33"].alignment = Alignment(wrap_text=True, vertical="top")
    ws["A33"].fill = PatternFill("solid", fgColor=AMBER)
    ws["A33"].border = border
    set_widths(ws, {"A": 35, "B": 17, "C": 17, "D": 17, "E": 17, "F": 34})
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "B20"


def build_scenarios(wb):
    ws = wb.create_sheet("Scenarios")
    style_title(ws, ws["A1"], "Rate scenarios — portfolio comparison", 13)
    ws.merge_cells("A3:N3")
    ws["A3"] = "Edit yellow rate and investment flags. P&L includes existing DI, DF campaign, Bolt+, Public Launch cash and Sponsored Listing discount."
    ws["A3"].font = Font(name="Aptos", italic=True, color=MUTED)
    headers = [
        "Scenario", "DK rate", "Resto rate", "Stores rate", "Use DF",
        "Use Bolt+", "Use PL cash", "Use in-kind", "Use SL / other",
        "Commission revenue", "CP L1", "CP L2 before investments",
        "Cash P&L after investment", "Cash P&L % GMV", "Partner saving vs baseline",
    ]
    header_row(ws, 5, headers)
    scenarios = [
        ("Baseline current", 0.2243, 0.2503, 0.12, 1, 1, 1, 1, 1),
        ("Tier 1 — stores 12% / DK-resto 25%", 0.25, 0.25, 0.12, 1, 1, 1, 1, 1),
        ("Tier 2 — stores 12% / DK-resto 20%", 0.20, 0.20, 0.12, 1, 1, 1, 1, 1),
        ("Tier 3a — stores 12% / DK-resto 16%", 0.16, 0.16, 0.12, 1, 1, 1, 1, 1),
        ("Tier 3b — shared 16%", 0.16, 0.16, 0.16, 1, 1, 1, 1, 1),
        ("Revenue-neutral shared 17.50%", 0.175, 0.175, 0.175, 1, 1, 1, 1, 1),
    ]
    for row, values in enumerate(scenarios, 6):
        for col, value in enumerate(values, 1):
            normal(ws, ws.cell(row, col), value, YELLOW if col >= 2 else (LIGHT_GREEN if row == 6 else None), col == 1)
            if col in {2, 3, 4}:
                ws.cell(row, col).number_format = "0.00%"
            if col in {5, 6, 7, 8, 9}:
                ws.cell(row, col).number_format = "0"
        # Portfolio formulas: Inputs annual GMV B21:D21, direct costs B24:D24, DI B26:D26.
        formula_cell(ws, row, 10, f"=SUMPRODUCT($B{row}:$D{row},Inputs!$B$21:$D$21)", '€#,##0')
        formula_cell(ws, row, 11, f"=J{row}-SUM(Inputs!$B$24:$D$24)", '€#,##0')
        formula_cell(ws, row, 12, f"=K{row}-SUM(Inputs!$B$26:$D$26)", '€#,##0')
        investments = (
            f"($E{row}*Inputs!$B$7)+($F{row}*SUM(Inputs!$B$28:$D$28))"
            f"+($G{row}*Inputs!$B$12)+($I{row}*(Inputs!$B$14+Inputs!$B$15))"
        )
        formula_cell(ws, row, 13, f"=L{row}-{investments}", '€#,##0', LIGHT_GREEN if row == 6 else None, row == 6)
        formula_cell(ws, row, 14, f"=M{row}/SUM(Inputs!$B$21:$D$21)", '0.00%')
        formula_cell(ws, row, 15, f"=$J$6-J{row}", '€#,##0')
    ws.conditional_formatting.add("M6:M11", CellIsRule(operator="lessThan", formula=["0"], fill=PatternFill("solid", fgColor=RED)))
    ws.conditional_formatting.add("M6:M11", CellIsRule(operator="greaterThan", formula=["0"], fill=PatternFill("solid", fgColor=LIGHT_GREEN)))
    ws.auto_filter.ref = "A5:O11"
    ws.freeze_panes = "J6"
    set_widths(ws, {"A": 38, "B": 12, "C": 12, "D": 12, "E": 10, "F": 11, "G": 12, "H": 12, "I": 13, "J": 17, "K": 15, "L": 21, "M": 24, "N": 18, "O": 24})
    ws.sheet_view.showGridLines = False


def build_pnl(wb):
    ws = wb.create_sheet("P&L by vertical")
    style_title(ws, ws["A1"], "P&L by vertical — selected scenario", 6)
    normal(ws, ws["A3"], "Selected scenario", bold=True)
    input_cell(ws, 3, 2, "Tier 3b — shared 16%")
    ws.merge_cells("C3:G3")
    ws["C3"] = "Select a scenario to compare Dark kitchen, resto, stores and portfolio. Yellow cells are editable."
    ws["C3"].font = Font(name="Aptos", italic=True, color=MUTED)
    dv = DataValidation(type="list", formula1="=Scenarios!$A$6:$A$11", allow_blank=False)
    ws.add_data_validation(dv)
    dv.add(ws["B3"])
    headers = ["P&L item", "Dark kitchen", "Resto", "Stores", "Portfolio", "Formula / treatment"]
    header_row(ws, 5, headers)
    rows = [
        ("Delivered orders", "orders"),
        ("GMV", "gmv"),
        ("AOV", "aov"),
        ("Selected commission rate", "rate"),
        ("Commission revenue", "commission"),
        ("Direct cost before demand incentives", "direct"),
        ("CP L1", "l1"),
        ("Existing demand incentives", "di"),
        ("CP L2 before incremental investment", "l2"),
        ("DF campaign", "df"),
        ("Bolt+ free delivery", "plus"),
        ("Public Launch cash", "pl"),
        ("Sponsored Listing + other cash", "sl"),
        ("Final P&L after cash investment", "final"),
        ("Final P&L / GMV", "margin"),
        ("In-kind marketing investment", "inkind"),
        ("Total Bolt cash investment", "cash_inv"),
        ("Total Bolt investment incl. in-kind", "total_inv"),
        ("Partner saving vs baseline commission", "saving"),
    ]
    for r, (label, key) in enumerate(rows, 6):
        normal(ws, ws.cell(r, 1), label, bold=key in {"l1", "l2", "final", "total_inv"})
    # Formula mapping by entity column. Inputs verticals B/C/D; scenario rates B/C/D.
    entity_inputs = {2: "B", 3: "C", 4: "D"}
    rate_cols = {2: "B", 3: "C", 4: "D"}
    for col, input_col in entity_inputs.items():
        col_letter = get_column_letter(col)
        rate_col = rate_cols[col]
        rate_lookup = f'=INDEX(Scenarios!${rate_col}$6:${rate_col}$11,MATCH($B$3,Scenarios!$A$6:$A$11,0))'
        flag = lambda flag_col: f'=INDEX(Scenarios!${flag_col}$6:${flag_col}$11,MATCH($B$3,Scenarios!$A$6:$A$11,0))'
        formula_cell(ws, 6, col, f"=Inputs!{input_col}20", '#,##0')
        formula_cell(ws, 7, col, f"=Inputs!{input_col}21", '€#,##0')
        formula_cell(ws, 8, col, f"={col_letter}7/{col_letter}6", '€0.00')
        formula_cell(ws, 9, col, rate_lookup, '0.00%', YELLOW)
        formula_cell(ws, 10, col, f"={col_letter}7*{col_letter}9", '€#,##0')
        formula_cell(ws, 11, col, f"=Inputs!{input_col}24", '€#,##0')
        formula_cell(ws, 12, col, f"={col_letter}10-{col_letter}11", '€#,##0', LIGHT_GREEN, True)
        formula_cell(ws, 13, col, f"=Inputs!{input_col}26", '€#,##0')
        formula_cell(ws, 14, col, f"={col_letter}12-{col_letter}13", '€#,##0', LIGHT_GREEN, True)
        df_formula = f"=IF({flag('E')}=0,0,{ 'Inputs!$B$7' if col == 4 else '0'})"
        formula_cell(ws, 15, col, df_formula, '€#,##0', AMBER)
        formula_cell(ws, 16, col, f"=IF({flag('F')}=0,0,Inputs!{input_col}28)", '€#,##0', AMBER)
        formula_cell(ws, 17, col, f"=IF({flag('G')}=0,0,Inputs!$B$12*Inputs!{input_col}29)", '€#,##0', AMBER)
        formula_cell(ws, 18, col, f"=IF({flag('I')}=0,0,(Inputs!$B$14+Inputs!$B$15)*Inputs!{input_col}29)", '€#,##0', AMBER)
        formula_cell(ws, 19, col, f"={col_letter}14-SUM({col_letter}15:{col_letter}18)", '€#,##0', LIGHT_GREEN, True)
        formula_cell(ws, 20, col, f"={col_letter}19/{col_letter}7", '0.00%', LIGHT_GREEN)
        formula_cell(ws, 21, col, f"=IF({flag('H')}=0,0,Inputs!$B$13*Inputs!{input_col}29)", '€#,##0')
        formula_cell(ws, 22, col, f"=SUM({col_letter}15:{col_letter}18)", '€#,##0')
        formula_cell(ws, 23, col, f"={col_letter}21+{col_letter}22", '€#,##0', AMBER, True)
        formula_cell(ws, 24, col, f"=Inputs!{input_col}23-{col_letter}10", '€#,##0')
    for row in range(6, 25):
        if row == 8:
            formula_cell(ws, row, 5, "=E7/E6", '€0.00')
        elif row == 9:
            formula_cell(ws, row, 5, "=E10/E7", '0.00%', YELLOW)
        elif row == 20:
            formula_cell(ws, row, 5, "=E19/E7", '0.00%', LIGHT_GREEN)
        else:
            fill = LIGHT_GREEN if row in {12, 14, 19} else (AMBER if row in {15, 16, 17, 18, 23} else None)
            formula_cell(ws, row, 5, f"=SUM(B{row}:D{row})", '€#,##0' if row != 6 else '#,##0', fill, row in {12, 14, 19, 23})
        normal(ws, ws.cell(row, 6), {
            6: "Input",
            7: "Input",
            8: "GMV / orders",
            9: "Scenario lookup",
            10: "GMV × selected rate",
            11: "Input; held constant",
            12: "Commission − direct cost",
            13: "Historical input",
            14: "CP L1 − incentives",
            15: "Stores only",
            16: "Scenario flag × input",
            17: "Allocated by GMV",
            18: "Allocated by GMV",
            19: "CP L2 − cash costs",
            20: "Final P&L / GMV",
            21: "Allocated by GMV",
            22: "Cash costs only",
            23: "Cash + in-kind",
            24: "Baseline commission − selected commission",
        }[row])
    ws.conditional_formatting.add("B19:E19", CellIsRule(operator="lessThan", formula=["0"], fill=PatternFill("solid", fgColor=RED)))
    set_widths(ws, {"A": 38, "B": 18, "C": 18, "D": 18, "E": 18, "F": 28})
    ws.freeze_panes = "B6"
    ws.sheet_view.showGridLines = False


def build_investments(wb):
    ws = wb.create_sheet("Investment & partner value")
    style_title(ws, ws["A1"], "Investment & partner value — selected scenario", 7)
    ws.merge_cells("A3:H3")
    ws["A3"] = "Values are linked to the scenario selected in P&L by vertical. Cash costs affect Final P&L; in-kind is shown separately."
    ws["A3"].font = Font(name="Aptos", italic=True, color=MUTED)
    header_row(ws, 5, ["Item", "Dark kitchen", "Resto", "Stores", "Portfolio", "Cash / in-kind", "P&L treatment", "Partner value"])
    data = [
        ("DF campaign", "=\'P&L by vertical\'!B15", "=\'P&L by vertical\'!C15", "=\'P&L by vertical\'!D15", "=SUM(B6:D6)", "Cash", "Deducted from Final P&L", "Demand activation"),
        ("Bolt+ free delivery", "=\'P&L by vertical\'!B16", "=\'P&L by vertical\'!C16", "=\'P&L by vertical\'!D16", "=SUM(B7:D7)", "Cash", "Deducted from Final P&L", "Higher-frequency Plus cohort"),
        ("Public Launch", "=\'P&L by vertical\'!B17", "=\'P&L by vertical\'!C17", "=\'P&L by vertical\'!D17", "=SUM(B8:D8)", "Cash", "Deducted from Final P&L", "Launch visibility"),
        ("Sponsored Listing + other", "=\'P&L by vertical\'!B18", "=\'P&L by vertical\'!C18", "=\'P&L by vertical\'!D18", "=SUM(B9:D9)", "Cash", "Deducted from Final P&L", "Manual input"),
        ("In-kind marketing", "=\'P&L by vertical\'!B21", "=\'P&L by vertical\'!C21", "=\'P&L by vertical\'!D21", "=SUM(B10:D10)", "In-kind", "Not deducted from cash P&L", "Banner, push/email, hero, modal"),
        ("Total Bolt cash investment", "=\'P&L by vertical\'!B22", "=\'P&L by vertical\'!C22", "=\'P&L by vertical\'!D22", "=SUM(B11:D11)", "Cash", "Deducted", "—"),
        ("Total Bolt investment", "=\'P&L by vertical\'!B23", "=\'P&L by vertical\'!C23", "=\'P&L by vertical\'!D23", "=SUM(B12:D12)", "Cash + in-kind", "Full investment view", "—"),
        ("Partner commission saving", "=\'P&L by vertical\'!B24", "=\'P&L by vertical\'!C24", "=\'P&L by vertical\'!D24", "=SUM(B13:D13)", "Partner value", "Does not affect Bolt P&L directly", "vs baseline commission"),
    ]
    for row, values in enumerate(data, 6):
        for col, value in enumerate(values, 1):
            formula_cell(ws, row, col, value, '€#,##0', LIGHT_GREEN if row in {11, 12, 13} else None, row in {11, 12, 13}) if isinstance(value, str) and value.startswith("=") else normal(ws, ws.cell(row, col), value, None, row in {11, 12, 13})
    set_widths(ws, {"A": 30, "B": 16, "C": 16, "D": 16, "E": 16, "F": 15, "G": 30, "H": 30})
    ws.sheet_view.showGridLines = False


def build_source(wb):
    ws = wb.create_sheet("Source & checks")
    style_title(ws, ws["A1"], "Source references & checks", 8)
    header_row(ws, 4, ["Source", "Used for", "Reference", "Status"])
    sources = [
        ("SILPO business model (6).xlsx", "Stores baseline and formats", "Provided by user", "Reference only; source contains legacy #REF cells"),
        ("Silpo 2027 — DK & Resto profitability calculator.xlsx", "DK / resto historical P&L format and source data", "Provided by user", "Used for structure and historical economics"),
        ("silpo-culinary-scenarios", "Offer thresholds and scenarios", "GitHub Pages v6a97416", "Rates/thresholds reflected in Scenarios"),
        ("Databricks / fact_provider_weekly", "Underlying DK/resto actuals", "Existing model source", "Not refreshed inside this workbook"),
    ]
    for row, values in enumerate(sources, 5):
        for col, value in enumerate(values, 1):
            normal(ws, ws.cell(row, col), value, LIGHT_GREEN if row == 7 else None)
    section(ws, 11, "Model limitations", 4)
    ws.merge_cells("A12:D15")
    ws["A12"] = (
        "This calculator keeps direct costs before demand incentives constant when commission rates change. "
        "DF campaign, Bolt+, Public Launch, Sponsored Listing and in-kind marketing are separately visible and editable. "
        "Before external use, confirm commission recognition, 2027 growth assumptions, promo overlap and final contractual rate definitions."
    )
    ws["A12"].alignment = Alignment(wrap_text=True, vertical="top")
    ws["A12"].fill = PatternFill("solid", fgColor=AMBER)
    ws["A12"].border = border
    set_widths(ws, {"A": 38, "B": 38, "C": 32, "D": 42})
    ws.sheet_view.showGridLines = False


def build_workbook():
    wb = Workbook()
    build_readme(wb)
    build_inputs(wb)
    build_scenarios(wb)
    build_pnl(wb)
    build_investments(wb)
    build_source(wb)
    for ws in wb.worksheets:
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_view.zoomScale = 90
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.calculation.calcMode = "auto"
    wb.save(OUTPUT)
    print(OUTPUT)


if __name__ == "__main__":
    build_workbook()
