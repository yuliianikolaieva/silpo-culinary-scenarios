from pathlib import Path
import xlsxwriter


OUT = Path(__file__).parent / "Silpo_2027_DK_Resto_Stores_PnL_Calculator.xlsx"

# Historical source: "Silpo 2027 — DK & Resto profitability calculator.xlsx", Oct-2025–Sep-2026
# Stores source: "SILPO business model (6).xlsx", column L "Спільна ставка - 16% ДК + Stores"
verticals = {
    "DK": {"orders": 56959, "gmv": 1029023.57, "eater": 115473.31, "commission": 207245.02, "cpl1": 137132.27, "di": 99019.71, "df_hist": 14297.19},
    "Resto": {"orders": 1931, "gmv": 34036.47, "eater": 3810.16, "commission": 7668.47, "cpl1": 4983.58, "di": 2001.89, "df_hist": 120.98},
    "Stores": {"orders": 50400, "gmv": 852768.00, "eater": 96768.00, "commission": 120960.00, "cpl1": 46330.28644, "di": 28728.00, "df_hist": 0.00},
}
for data in verticals.values():
    data["direct"] = data["commission"] + data["eater"] - data["cpl1"]
    data["cpl2"] = data["cpl1"] - data["di"]
    data["rate"] = data["commission"] / data["gmv"]

SCENARIOS = [
    ("Current effective rates", verticals["DK"]["rate"], verticals["DK"]["rate"], verticals["Resto"]["rate"], verticals["Resto"]["rate"], 0.16, 0.16, "Current historical / Shared16 stores", 0),
    ("Tier 1 · launch", 0.25, 0.25, 0.25, 0.25, 0.12, 0.12, "12 months at launch rates", 0),
    ("Tier 2 · €33k stores GMV / month", 0.25, 0.20, 0.25, 0.20, 0.12, 0.12, "6m DK/resto 25%, then 6m 20%; threshold €33k", 33000),
    ("Tier 3a · step-down to 16%", 0.20, 0.16, 0.20, 0.16, 0.12, 0.12, "6m DK/resto 20%, then 6m 16%; stores stays 12%", 65000),
    ("Tier 3a · 12m DK/resto 16%", 0.16, 0.16, 0.16, 0.16, 0.12, 0.12, "12 months: stores 12%, DK/resto 16%", 65000),
    ("Tier 3b · step to shared 16%", 0.20, 0.16, 0.20, 0.16, 0.12, 0.16, "6m tier 2, then 6m shared 16%", 65000),
    ("Tier 3b · 12m shared 16%", 0.16, 0.16, 0.16, 0.16, 0.16, 0.16, "12 months shared 16%", 65000),
]


def build():
    wb = xlsxwriter.Workbook(OUT)
    wb.set_calc_mode("auto")
    fmt = {
        "title": wb.add_format({"bold": True, "font_size": 18, "font_color": "#FFFFFF", "bg_color": "#17231E", "valign": "vcenter"}),
        "section": wb.add_format({"bold": True, "font_color": "#FFFFFF", "bg_color": "#0B8F50", "valign": "vcenter"}),
        "header": wb.add_format({"bold": True, "font_color": "#FFFFFF", "bg_color": "#17231E", "align": "center", "valign": "vcenter", "text_wrap": True, "border": 1, "border_color": "#D9E3DD"}),
        "label": wb.add_format({"font_color": "#17231E", "border": 1, "border_color": "#D9E3DD"}),
        "input": wb.add_format({"bg_color": "#FFF2CC", "border": 1, "border_color": "#D9E3DD", "num_format": "#,##0"}),
        "input_eur": wb.add_format({"bg_color": "#FFF2CC", "border": 1, "border_color": "#D9E3DD", "num_format": '€#,##0'}),
        "input_pct": wb.add_format({"bg_color": "#FFF2CC", "border": 1, "border_color": "#D9E3DD", "num_format": "0.00%"}),
        "num": wb.add_format({"border": 1, "border_color": "#D9E3DD", "num_format": "#,##0"}),
        "eur": wb.add_format({"border": 1, "border_color": "#D9E3DD", "num_format": '€#,##0'}),
        "pct": wb.add_format({"border": 1, "border_color": "#D9E3DD", "num_format": "0.00%"}),
        "green_eur": wb.add_format({"bg_color": "#EDF7F1", "border": 1, "border_color": "#D9E3DD", "num_format": '€#,##0', "bold": True}),
        "green_pct": wb.add_format({"bg_color": "#EDF7F1", "border": 1, "border_color": "#D9E3DD", "num_format": "0.00%", "bold": True}),
        "amber_eur": wb.add_format({"bg_color": "#FFF7E7", "border": 1, "border_color": "#D9E3DD", "num_format": '€#,##0'}),
        "note": wb.add_format({"font_color": "#64726A", "italic": True, "text_wrap": True}),
        "text": wb.add_format({"border": 1, "border_color": "#D9E3DD", "text_wrap": True, "valign": "top"}),
        "bad_eur": wb.add_format({"bg_color": "#FFF1EE", "font_color": "#BD3D31", "border": 1, "border_color": "#D9E3DD", "num_format": '€#,##0', "bold": True}),
    }

    # Read me
    ws = wb.add_worksheet("Read me")
    ws.hide_gridlines(2)
    ws.merge_range("A1:J1", "Silpo 2027 — DK, Resto & Stores P&L calculator", fmt["title"])
    ws.write("A3", "How this workbook works", fmt["section"])
    notes = [
        "Inputs: yellow cells are editable. DK/resto is populated from the provided 12-month historical calculator; Stores is populated from SILPO business model (6), scenario “Спільна ставка - 16% ДК + Stores”.",
        "Scenarios: annual total only. Each tier shows H1 and H2 commission, the GMV threshold, CP L1 and CP L2 both in € and % of GMV.",
        "P&L by vertical: displays Dark Kitchen, Resto, Stores and Portfolio separately for every scenario, including eater fees and all cost lines.",
        "DF campaign for DK/resto is historical information inside existing demand incentives and is not deducted twice. Incremental DF campaign is a separate Stores cash investment for 6 of 12 months.",
        "Bolt+ discount uses delivery base only. SF and SOF are excluded from the discount input.",
    ]
    for i, note in enumerate(notes, 5):
        ws.merge_range(i - 1, 0, i - 1, 9, note, fmt["text"])
        ws.set_row(i - 1, 32)
    ws.set_column("A:J", 16)

    # Inputs
    ws = wb.add_worksheet("Inputs")
    ws.hide_gridlines(2)
    ws.merge_range("A1:G1", "Inputs — historical data and 2027 projection", fmt["title"])
    ws.merge_range("A3:G3", "Yellow cells are editable. 2027 projection = historical 12-month baseline × volume multiplier.", fmt["note"])
    ws.merge_range("A5:D5", "Incremental investments and promotional assumptions", fmt["section"])
    ws.write_row("A6", ["Input", "Value", "Unit", "Treatment"], fmt["header"])
    global_inputs = [
        ("DF campaign · Stores", 28000, "€ / year", "Cash; activation 6 of 12 months"),
        ("DF campaign · DK + Resto", 14418.17, "€ / year", "Historical info only; already included in existing DI"),
        ("Bolt+ · DK + Resto", 14615.45, "€ / year", "Cash; delivery base only"),
        ("Bolt+ · Stores", 10800, "€ / year", "Cash; delivery base only"),
        ("Public Launch cash", 10000, "€ / year", "Cash; allocated by GMV"),
        ("In-kind marketing", 11200, "€ / year", "Not deducted from cash P&L; allocated by GMV"),
        ("Sponsored Listing discount", 0, "€ / year", "Cash impact; allocated by GMV"),
        ("Other cash investment", 0, "€ / year", "Cash impact; allocated by GMV"),
    ]
    for r, (name, value, unit, treatment) in enumerate(global_inputs, 7):
        ws.write(r - 1, 0, name, fmt["label"])
        ws.write(r - 1, 1, value, fmt["input_eur"])
        ws.write(r - 1, 2, unit, fmt["label"])
        ws.write(r - 1, 3, treatment, fmt["text"])

    ws.merge_range("A17:F17", "Historical 12-month base and 2027 projection", fmt["section"])
    ws.write_row("A18", ["Metric", "Dark Kitchen", "Resto", "Stores", "Portfolio", "Source / treatment"], fmt["header"])
    metrics = [
        ("Historical delivered orders", "orders", "#,##0"),
        ("Historical GMV", "gmv", '€#,##0'),
        ("Historical eater-fee revenue", "eater", '€#,##0'),
        ("Historical commission revenue", "commission", '€#,##0'),
        ("Historical direct costs before DI", "direct", '€#,##0'),
        ("Historical CP L1", "cpl1", '€#,##0'),
        ("Historical demand incentives", "di", '€#,##0'),
        ("Historical CP L2", "cpl2", '€#,##0'),
        ("2027 volume multiplier", None, "0.00%"),
        ("2027 projected orders", "orders", "#,##0"),
        ("2027 projected GMV", "gmv", '€#,##0'),
        ("2027 projected eater fees", "eater", '€#,##0'),
        ("2027 projected direct costs", "direct", '€#,##0'),
        ("2027 projected existing DI", "di", '€#,##0'),
    ]
    for r, (label, key, number_fmt) in enumerate(metrics, 19):
        ws.write(r - 1, 0, label, fmt["label"])
        for c, name in enumerate(["DK", "Resto", "Stores"], 1):
            if label == "2027 volume multiplier":
                ws.write(r - 1, c, 1.0, fmt["input_pct"])
            elif r >= 28:
                hist_row = {28: 19, 29: 20, 30: 21, 31: 23, 32: 25}[r]
                formula = f"={chr(66+c-1)}{hist_row}*{chr(66+c-1)}27"
                # cached result matches the same input value at multiplier 1
                ws.write_formula(r - 1, c, formula, fmt["green_eur"] if "GMV" in label or "CP" in label else fmt["num"], verticals[name][key])
            else:
                ws.write(r - 1, c, verticals[name][key], fmt["input_eur"] if number_fmt.startswith("€") else fmt["input"])
        if label == "2027 volume multiplier":
            ws.write_formula(r - 1, 4, "=AVERAGE(B27:D27)", fmt["pct"], 1.0)
        elif number_fmt == "#,##0":
            ws.write_formula(r - 1, 4, f"=SUM(B{r}:D{r})", fmt["num"], sum(v[key] for v in verticals.values()))
        else:
            cached = sum(v[key] for v in verticals.values())
            ws.write_formula(r - 1, 4, f"=SUM(B{r}:D{r})", fmt["green_eur"] if "CP" in label or "GMV" in label else fmt["eur"], cached)
        source = "Provided DK/resto calculator" if r < 27 else "Historical × volume multiplier"
        if "Stores" in label and r < 27:
            source = "SILPO business model (6) · Shared 16%"
        ws.write(r - 1, 5, source, fmt["text"])
    ws.set_column("A:A", 34)
    ws.set_column("B:E", 18)
    ws.set_column("F:F", 38)

    # Scenarios annual
    ws = wb.add_worksheet("Scenarios")
    ws.hide_gridlines(2)
    ws.merge_range("A1:Z1", "Scenarios — annual total, H1/H2 rate schedule and GMV thresholds", fmt["title"])
    ws.merge_range("A3:Z3", "Rate cells are yellow and editable. Threshold is a commercial trigger; current Stores average monthly GMV is visible in Inputs / 2027 projected GMV.", fmt["note"])
    headers = ["Scenario", "DK H1", "DK H2", "Resto H1", "Resto H2", "Stores H1", "Stores H2", "GMV threshold / month", "Model comment", "Orders", "GMV", "Eater fees", "Commission", "Direct costs", "CP L1 €", "CP L1 %", "Existing DI", "Historical DF DK/resto\ninside DI", "CP L2 €", "CP L2 %", "DF Stores", "Bolt+", "PL cash", "Final CP L2", "Final %", "In-kind", "Partner saving"]
    ws.write_row("A5", headers, fmt["header"])
    total_orders = sum(v["orders"] for v in verticals.values())
    total_gmv = sum(v["gmv"] for v in verticals.values())
    total_eater = sum(v["eater"] for v in verticals.values())
    total_direct = sum(v["direct"] for v in verticals.values())
    total_di = sum(v["di"] for v in verticals.values())
    baseline_commission = sum(v["commission"] for v in verticals.values())
    for r, scenario in enumerate(SCENARIOS, 6):
        name, dkh1, dkh2, rh1, rh2, sh1, sh2, comment, threshold = scenario
        values = [name, dkh1, dkh2, rh1, rh2, sh1, sh2, threshold, comment]
        for c, value in enumerate(values):
            style = fmt["input_pct"] if 1 <= c <= 6 else (fmt["input_eur"] if c == 7 else fmt["text"])
            ws.write(r - 1, c, value, style)
        # Inputs sheet 2027 rows: orders 28, GMV 29, eater 30, direct 31, DI 32
        commission = (verticals["DK"]["gmv"] * (dkh1 + dkh2) / 2 +
                      verticals["Resto"]["gmv"] * (rh1 + rh2) / 2 +
                      verticals["Stores"]["gmv"] * (sh1 + sh2) / 2)
        cpl1 = commission + total_eater - total_direct
        cpl2 = cpl1 - total_di
        final = cpl2 - 28000 - 25415.45 - 10000
        partner_saving = baseline_commission - commission
        formulas = [
            ("=SUM(Inputs!$B$28:$D$28)", total_orders, fmt["num"]),
            ("=SUM(Inputs!$B$29:$D$29)", total_gmv, fmt["eur"]),
            ("=SUM(Inputs!$B$30:$D$30)", total_eater, fmt["eur"]),
            (f"=Inputs!$B$29*(B{r}+C{r})/2+Inputs!$C$29*(D{r}+E{r})/2+Inputs!$D$29*(F{r}+G{r})/2", commission, fmt["eur"]),
            ("=SUM(Inputs!$B$31:$D$31)", total_direct, fmt["eur"]),
            (f"=M{r}-N{r}+L{r}", cpl1, fmt["green_eur"]),
            (f"=O{r}/K{r}", cpl1 / total_gmv, fmt["green_pct"]),
            ("=SUM(Inputs!$B$32:$D$32)", total_di, fmt["eur"]),
            ("=Inputs!$B$8", 14418.17, fmt["amber_eur"]),
            (f"=O{r}-Q{r}", cpl2, fmt["green_eur"]),
            (f"=S{r}/K{r}", cpl2 / total_gmv, fmt["green_pct"]),
            ("=Inputs!$B$7", 28000, fmt["amber_eur"]),
            ("=Inputs!$B$9+Inputs!$B$10", 25415.45, fmt["amber_eur"]),
            ("=Inputs!$B$11", 10000, fmt["amber_eur"]),
            (f"=S{r}-SUM(U{r}:W{r})", final, fmt["green_eur"] if final >= 0 else fmt["bad_eur"]),
            (f"=X{r}/K{r}", final / total_gmv, fmt["green_pct"]),
            ("=Inputs!$B$12", 11200, fmt["eur"]),
            (f"=$M$6-M{r}", partner_saving, fmt["eur"]),
        ]
        for c, (formula, cached, style) in enumerate(formulas, 9):
            ws.write_formula(r - 1, c, formula, style, cached)
    ws.autofilter(4, 0, 4 + len(SCENARIOS), 26)
    ws.freeze_panes(5, 9)
    ws.set_column("A:A", 35)
    ws.set_column("B:G", 11)
    ws.set_column("H:H", 15)
    ws.set_column("I:I", 38)
    ws.set_column("J:AA", 15)

    # P&L by vertical - all scenarios, no dependent dropdowns
    ws = wb.add_worksheet("P&L by vertical")
    ws.hide_gridlines(2)
    ws.merge_range("A1:T1", "P&L by vertical — all annual scenarios", fmt["title"])
    ws.merge_range("A3:T3", "Each row has an internal formula and a cached opening value. Eater fees are shown separately but already included in historical unit economics; they are not added twice.", fmt["note"])
    headers = ["Scenario", "Vertical", "Orders", "GMV", "AOV", "Eater fees", "H1 rate", "H2 rate", "Commission", "Direct costs", "CP L1 €", "CP L1 %", "Existing DI", "Historical DF\ninside DI", "CP L2 €", "CP L2 %", "DF Stores", "Bolt+", "PL cash", "Final CP L2", "Final %"]
    ws.write_row("A5", headers, fmt["header"])
    row = 6
    for scenario_index, scenario in enumerate(SCENARIOS):
        name, dkh1, dkh2, rh1, rh2, sh1, sh2, _, _ = scenario
        source_scenario_row = 6 + scenario_index
        rates = {"DK": (dkh1, dkh2, "B", "C", "B"), "Resto": (rh1, rh2, "D", "E", "C"), "Stores": (sh1, sh2, "F", "G", "D")}
        for key in ["DK", "Resto", "Stores"]:
            data = verticals[key]
            rate1, rate2, h1_col, h2_col, input_col = rates[key]
            commission = data["gmv"] * (rate1 + rate2) / 2
            cpl1 = commission + data["eater"] - data["direct"]
            cpl2 = cpl1 - data["di"]
            df = 28000 if key == "Stores" else 0
            plus = 14615.45 * data["gmv"] / (verticals["DK"]["gmv"] + verticals["Resto"]["gmv"]) if key in {"DK", "Resto"} else 10800
            pl = 10000 * data["gmv"] / total_gmv
            final = cpl2 - df - plus - pl
            ws.write(row - 1, 0, name, fmt["text"])
            ws.write(row - 1, 1, key, fmt["text"])
            ws.write_formula(row - 1, 2, f"=Inputs!${input_col}$28", fmt["num"], data["orders"])
            ws.write_formula(row - 1, 3, f"=Inputs!${input_col}$29", fmt["eur"], data["gmv"])
            ws.write_formula(row - 1, 4, f"=D{row}/C{row}", fmt["eur"], data["gmv"] / data["orders"])
            ws.write_formula(row - 1, 5, f"=Inputs!${input_col}$30", fmt["eur"], data["eater"])
            ws.write_formula(row - 1, 6, f"=Scenarios!${h1_col}${source_scenario_row}", fmt["pct"], rate1)
            ws.write_formula(row - 1, 7, f"=Scenarios!${h2_col}${source_scenario_row}", fmt["pct"], rate2)
            ws.write_formula(row - 1, 8, f"=D{row}*(G{row}+H{row})/2", fmt["eur"], commission)
            ws.write_formula(row - 1, 9, f"=Inputs!${input_col}$31", fmt["eur"], data["direct"])
            ws.write_formula(row - 1, 10, f"=I{row}+F{row}-J{row}", fmt["green_eur"], cpl1)
            ws.write_formula(row - 1, 11, f"=K{row}/D{row}", fmt["pct"], cpl1 / data["gmv"])
            ws.write_formula(row - 1, 12, f"=Inputs!${input_col}$32", fmt["eur"], data["di"])
            hist_df_formula = f"=Inputs!$B$8*D{row}/(Inputs!$B$29+Inputs!$C$29)" if key in {"DK", "Resto"} else "=0"
            ws.write_formula(row - 1, 13, hist_df_formula, fmt["amber_eur"], data["df_hist"])
            ws.write_formula(row - 1, 14, f"=K{row}-M{row}", fmt["green_eur"], cpl2)
            ws.write_formula(row - 1, 15, f"=O{row}/D{row}", fmt["pct"], cpl2 / data["gmv"])
            ws.write_formula(row - 1, 16, "=Inputs!$B$7" if key == "Stores" else "=0", fmt["amber_eur"], df)
            plus_formula = f"=Inputs!$B$9*D{row}/(Inputs!$B$29+Inputs!$C$29)" if key in {"DK", "Resto"} else "=Inputs!$B$10"
            ws.write_formula(row - 1, 17, plus_formula, fmt["amber_eur"], plus)
            ws.write_formula(row - 1, 18, f"=Inputs!$B$11*D{row}/SUM(Inputs!$B$29:$D$29)", fmt["amber_eur"], pl)
            ws.write_formula(row - 1, 19, f"=O{row}-Q{row}-R{row}-S{row}", fmt["bad_eur"] if final < 0 else fmt["green_eur"], final)
            ws.write_formula(row - 1, 20, f"=T{row}/D{row}", fmt["pct"], final / data["gmv"])
            row += 1
        row += 1
    ws.autofilter(4, 0, row - 2, 20)
    ws.freeze_panes(5, 2)
    ws.set_column("A:A", 35)
    ws.set_column("B:B", 14)
    ws.set_column("C:U", 14)

    # Source / assumptions
    ws = wb.add_worksheet("Source & assumptions")
    ws.hide_gridlines(2)
    ws.merge_range("A1:F1", "Source data and calculation assumptions", fmt["title"])
    ws.write_row("A4", ["Source", "Metric / assumption", "Value", "Treatment", "Notes", "Status"], fmt["header"])
    rows = [
        ("Silpo 2027 — DK & Resto profitability calculator.xlsx", "Historical DK + resto", "Oct-2025–Sep-2026", "Pulled into Inputs", "DK_A + DK_B + Resto", "Used"),
        ("SILPO business model (6).xlsx", "Stores scenario", "Shared rate 16%", "Pulled into Inputs", "Column L", "Used"),
        ("Eater fees", "DK/resto", "€119.3k", "Displayed separately", "Already contained in historical CP unit economics", "Informational"),
        ("DF campaign", "DK/resto", "€14.4k", "Inside historical DI", "Visible but not deducted a second time", "Informational"),
        ("DF campaign", "Stores", "€28.0k", "Incremental cash cost", "6 of 12 months", "Editable"),
        ("Bolt+", "All verticals", "€25.4k", "Incremental cash cost", "Delivery base only; no SF/SOF discount", "Editable"),
    ]
    for r, data in enumerate(rows, 5):
        for c, value in enumerate(data):
            ws.write(r - 1, c, value, fmt["text"])
    ws.set_column("A:A", 48)
    ws.set_column("B:B", 28)
    ws.set_column("C:C", 18)
    ws.set_column("D:D", 28)
    ws.set_column("E:E", 45)
    ws.set_column("F:F", 18)
    wb.close()
    print(OUT)


if __name__ == "__main__":
    build()
