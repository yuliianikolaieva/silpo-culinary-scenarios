from pathlib import Path

import xlsxwriter


OUT = Path(__file__).parent / "Silpo_Commission_Mix_and_Profitability_Calculator.xlsx"

# 12-month factual base used in the existing Silpo calculator.
DK_RESTO_GMV = 1_063_060.04
STORES_GMV = 852_768.00
STORES_ORDERS = 50_400
STORES_EATER = 96_768.00
STORES_DIRECT = 171_397.71356
STORES_DI = 28_728.00
DK_RESTO_EATER = 119_283.47
DK_RESTO_DIRECT = 192_081.11
DK_RESTO_DI = 101_021.60
PORTFOLIO_GMV = DK_RESTO_GMV + STORES_GMV
PORTFOLIO_ORDERS = 56_959 + 1_931 + STORES_ORDERS


def build() -> None:
    wb = xlsxwriter.Workbook(OUT)
    wb.set_calc_mode("auto")

    fmt = {
        "title": wb.add_format({"bold": True, "font_size": 18, "font_color": "#FFFFFF", "bg_color": "#17231E", "valign": "vcenter"}),
        "section": wb.add_format({"bold": True, "font_color": "#FFFFFF", "bg_color": "#0B8F50", "valign": "vcenter"}),
        "header": wb.add_format({"bold": True, "font_color": "#FFFFFF", "bg_color": "#17231E", "align": "center", "valign": "vcenter", "text_wrap": True, "border": 1, "border_color": "#D9E3DD"}),
        "label": wb.add_format({"border": 1, "border_color": "#D9E3DD", "valign": "top"}),
        "text": wb.add_format({"border": 1, "border_color": "#D9E3DD", "valign": "top", "text_wrap": True}),
        "note": wb.add_format({"font_color": "#64726A", "italic": True, "text_wrap": True}),
        "input_eur": wb.add_format({"bg_color": "#FFF2CC", "border": 1, "border_color": "#D9E3DD", "num_format": '€#,##0'}),
        "input_pct": wb.add_format({"bg_color": "#FFF2CC", "border": 1, "border_color": "#D9E3DD", "num_format": "0.0%"}),
        "eur": wb.add_format({"border": 1, "border_color": "#D9E3DD", "num_format": '€#,##0'}),
        "pct": wb.add_format({"border": 1, "border_color": "#D9E3DD", "num_format": "0.0%"}),
        "num": wb.add_format({"border": 1, "border_color": "#D9E3DD", "num_format": "#,##0"}),
        "green_eur": wb.add_format({"bg_color": "#EDF7F1", "border": 1, "border_color": "#B9DCC7", "num_format": '€#,##0', "bold": True}),
        "green_pct": wb.add_format({"bg_color": "#EDF7F1", "border": 1, "border_color": "#B9DCC7", "num_format": "0.0%", "bold": True}),
        "amber": wb.add_format({"bg_color": "#FFF7E7", "border": 1, "border_color": "#ECD9AB", "text_wrap": True}),
        "bad_eur": wb.add_format({"bg_color": "#FFF1EE", "font_color": "#BD3D31", "border": 1, "border_color": "#E7BCB5", "num_format": '€#,##0', "bold": True}),
        "bad_pct": wb.add_format({"bg_color": "#FFF1EE", "font_color": "#BD3D31", "border": 1, "border_color": "#E7BCB5", "num_format": "0.0%", "bold": True}),
    }

    # Read me
    ws = wb.add_worksheet("Read me")
    ws.hide_gridlines(2)
    ws.merge_range("A1:H1", "Silpo — commission mix & profitability calculator", fmt["title"])
    ws.merge_range("A3:H3", "Purpose: test whether a change in commission is a real partner saving or only a shift of GMV from DK/resto to Stores.", fmt["section"])
    notes = [
        "Yellow cells are editable. Green cells show a positive saving / contribution; red cells flag a loss or a worse outcome for the relevant comparison.",
        "Mix matrix: tests the same total GMV at different Stores vs DK/resto shares. It shows why a shared 16% is not automatically a better deal for a Stores-heavy mix.",
        "DK 16% + flexible Stores: holds DK/resto at 16% and varies the Stores rate. Compare both saving for Silpo and contribution profit for Bolt.",
        "Bolt unit economics: the profitability test uses the existing 12-month Stores unit economics. It includes eater fees, direct costs and existing demand incentives, before any new launch investment.",
        "Key guardrail: commission / investment changes should be triggered by incremental combined GMV and preserved DK/resto volume, not by orders moving from one vertical to another.",
    ]
    for row, note in enumerate(notes, 5):
        ws.merge_range(row - 1, 0, row - 1, 7, note, fmt["text"])
        ws.set_row(row - 1, 32)
    ws.set_column("A:H", 20)

    # Inputs
    ws = wb.add_worksheet("Inputs")
    ws.hide_gridlines(2)
    ws.merge_range("A1:E1", "Inputs — editable commercial and unit-economics assumptions", fmt["title"])
    ws.write("A3", "Commission assumptions", fmt["section"])
    ws.write_row("A4", ["Input", "Value", "Unit", "Use in calculation", "Source / note"], fmt["header"])
    inputs = [
        ("Reference total GMV for mix examples", PORTFOLIO_GMV, "€ / year", "2027 portfolio projection used in sensitivity tables", "DK + Resto + Stores 12-month projection"),
        ("Current DK + resto commission", 0.25, "%", "Partner baseline", "Commercial current rate"),
        ("Current Stores commission", 0.12, "%", "Partner baseline", "Bolt couriers minimum rate"),
        ("Growth DK + resto commission", 0.20, "%", "After €33k Stores GMV / month", "Scenario 1"),
        ("Growth Stores commission", 0.12, "%", "After €33k Stores GMV / month", "Scenario 1"),
        ("Shared commission", 0.16, "%", "Shared-rate comparison", "Scenario 2 / test only"),
        ("DK + resto floor for alternative", 0.16, "%", "Flexible Stores-rate table", "Do not go below without separate approval"),
    ]
    for row, (name, value, unit, use, source) in enumerate(inputs, 5):
        ws.write(row - 1, 0, name, fmt["label"])
        ws.write(row - 1, 1, value, fmt["input_eur"] if unit.startswith("€") else fmt["input_pct"])
        ws.write(row - 1, 2, unit, fmt["label"])
        ws.write(row - 1, 3, use, fmt["text"])
        ws.write(row - 1, 4, source, fmt["text"])

    ws.write("A14", "Stores unit economics — 12-month base", fmt["section"])
    ws.write_row("A15", ["Metric", "Value", "Unit", "Treatment", "Source / note"], fmt["header"])
    store_inputs = [
        ("Stores GMV", STORES_GMV, "€ / year", "Volume base", "SILPO business model (6) · shared 16% scenario"),
        ("Stores delivered orders", STORES_ORDERS, "orders / year", "Volume base", "SILPO business model (6)"),
        ("Eater-fee revenue", STORES_EATER, "€ / year", "Revenue", "€1.92 / order"),
        ("Direct costs before DI", STORES_DIRECT, "€ / year", "Costs", "Courier, payment, operations, support etc.; excludes DI"),
        ("Existing demand incentives", STORES_DI, "€ / year", "Costs", "Existing DI / promotions"),
    ]
    for row, (name, value, unit, treatment, source) in enumerate(store_inputs, 16):
        ws.write(row - 1, 0, name, fmt["label"])
        ws.write(row - 1, 1, value, fmt["input_eur"] if "€" in unit else fmt["num"])
        ws.write(row - 1, 2, unit, fmt["label"])
        ws.write(row - 1, 3, treatment, fmt["text"])
        ws.write(row - 1, 4, source, fmt["text"])

    ws.write("A23", "DK + resto 12-month base", fmt["section"])
    ws.write_row("A24", ["Metric", "Value", "Unit", "Treatment", "Source / note"], fmt["header"])
    dk_inputs = [
        ("DK + resto GMV", DK_RESTO_GMV, "€ / year", "Volume base", "Provided DK & Resto profitability calculator"),
        ("Eater-fee revenue", DK_RESTO_EATER, "€ / year", "Revenue", "Historical"),
        ("Direct costs before DI", DK_RESTO_DIRECT, "€ / year", "Costs", "Historical; excludes DI"),
        ("Existing demand incentives", DK_RESTO_DI, "€ / year", "Costs", "Historical"),
    ]
    for row, (name, value, unit, treatment, source) in enumerate(dk_inputs, 25):
        ws.write(row - 1, 0, name, fmt["label"])
        ws.write(row - 1, 1, value, fmt["input_eur"])
        ws.write(row - 1, 2, unit, fmt["label"])
        ws.write(row - 1, 3, treatment, fmt["text"])
        ws.write(row - 1, 4, source, fmt["text"])
    ws.set_column("A:A", 34)
    ws.set_column("B:B", 18)
    ws.set_column("C:C", 16)
    ws.set_column("D:E", 34)

    # Projection base: the exact user-provided 12-month table.
    ws = wb.add_worksheet("Projection base")
    ws.hide_gridlines(2)
    ws.merge_range("A1:F1", "Historical 12-month base and 2027 projection", fmt["title"])
    ws.merge_range("A3:F3", "2027 projection is set to 100% of the provided 12-month DK, Resto and Stores base. All sensitivity sheets use this portfolio GMV as their default.", fmt["note"])
    ws.write_row("A5", ["Metric", "Dark Kitchen", "Resto", "Stores", "Portfolio", "Source / treatment"], fmt["header"])
    projection = [
        ("Historical delivered orders", 56_959, 1_931, 50_400, PORTFOLIO_ORDERS, "Provided DK/resto calculator"),
        ("Historical GMV", 1_029_023.57, 34_036.47, STORES_GMV, PORTFOLIO_GMV, "Provided DK/resto calculator"),
        ("Historical eater-fee revenue", 115_473.31, 3_810.16, STORES_EATER, 216_051.47, "Provided DK/resto calculator"),
        ("Historical commission revenue", 207_245.02, 7_668.47, 120_960.00, 335_873.49, "Provided DK/resto calculator"),
        ("Historical direct costs before DI", 185_586.06, 6_495.05, STORES_DIRECT, 363_479.82356, "Provided DK/resto calculator"),
        ("Historical CP L1", 137_132.27, 4_983.58, 46_330.28644, 188_446.13644, "Provided DK/resto calculator"),
        ("Historical demand incentives", 99_019.71, 2_001.89, STORES_DI, 129_749.60, "Provided DK/resto calculator"),
        ("Historical CP L2", 38_112.56, 2_981.70, 17_602.28644, 58_696.54644, "Provided DK/resto calculator"),
        ("2027 volume multiplier", 1.0, 1.0, 1.0, 1.0, "Historical × volume multiplier"),
        ("2027 projected orders", 56_959, 1_931, 50_400, PORTFOLIO_ORDERS, "Historical × volume multiplier"),
        ("2027 projected GMV", 1_029_023.57, 34_036.47, STORES_GMV, PORTFOLIO_GMV, "Historical × volume multiplier"),
        ("2027 projected eater fees", 115_473.31, 3_810.16, STORES_EATER, 216_051.47, "Historical × volume multiplier"),
        ("2027 projected direct costs", 185_586.06, 6_495.05, STORES_DIRECT, 363_479.82356, "Historical × volume multiplier"),
        ("2027 projected existing DI", 99_019.71, 2_001.89, STORES_DI, 129_749.60, "Historical × volume multiplier"),
    ]
    for row, (metric, dk, resto, stores, portfolio, source) in enumerate(projection, 6):
        ws.write(row - 1, 0, metric, fmt["label"])
        is_rate = "multiplier" in metric
        number_format = fmt["pct"] if is_rate else (fmt["num"] if "orders" in metric.lower() else fmt["eur"])
        ws.write_row(row - 1, 1, [dk, resto, stores, portfolio], number_format)
        ws.write(row - 1, 5, source, fmt["text"])
    ws.set_column("A:A", 34)
    ws.set_column("B:E", 18)
    ws.set_column("F:F", 36)

    # Unified partner/Bolt view. Source-scenario values are taken from the
    # "Scenarios" tab of Silpo_2027_DK_Resto_Stores_PnL_Calculator (6).xlsx.
    ws = wb.add_worksheet("Scenario summary")
    ws.hide_gridlines(2)
    ws.merge_range("A1:S1", "Scenario summary — partner saving and Bolt CP L1 / CP L2", fmt["title"])
    ws.merge_range("A3:S3", "Source rows replicate the supplied Scenarios tab. The rows “Stores rate test” hold DK/resto GMV and rate at the shared-16% base, then reduce Stores commission from 16% to 5%.", fmt["note"])
    headers = [
        "Scenario", "Type", "DK\nrate", "Resto\nrate", "Stores\nrate", "DK\nGMV", "Resto\nGMV", "Stores\nGMV",
        "Portfolio\nGMV", "Weighted\ncommission", "Commission\nrevenue", "Eater fees", "Direct costs",
        "CP L1", "CP L1\n%", "Existing DI", "CP L2", "CP L2\n%", "Partner saving\nvs current",
    ]
    ws.write_row("A5", headers, fmt["header"])

    source_scenarios = [
        ("Current effective rates", 0.2013996822, 0.2253015662, 0.16, 852_768.00, 216_051.47, 363_478.8236, 129_749.60),
        ("Tier 1 · launch", 0.25, 0.25, 0.12, 427_800.00, 178_054.67, 296_178.0150, 118_469.30),
        ("Tier 2 · €33k stores GMV / month", 0.225, 0.225, 0.12, 427_800.00, 178_054.67, 296_178.0150, 118_469.30),
        ("Tier 3a · step-down to 16%", 0.18, 0.18, 0.12, 852_768.00, 216_051.47, 363_478.8236, 129_749.60),
        ("Tier 3a · 12m DK/resto 16%", 0.16, 0.16, 0.14, 852_768.00, 216_051.47, 363_478.8236, 129_749.60),
        ("Tier 3b · step to shared 16%", 0.18, 0.18, 0.14, 852_768.00, 216_051.47, 363_478.8236, 129_749.60),
        ("Tier 3b · 12m shared 16%", 0.16, 0.16, 0.16, 852_768.00, 216_051.47, 363_478.8236, 129_749.60),
    ]
    scenarios = [(name, "Source · Scenarios", dk, resto, stores, store_gmv, eater, direct, di)
                 for name, dk, resto, stores, store_gmv, eater, direct, di in source_scenarios]
    scenarios += [
        (f"Stores rate test · {rate:.0%}", "Test · DK/resto fixed 16%", 0.16, 0.16, rate,
         STORES_GMV, 216_051.47, 363_478.8236, 129_749.60)
        for rate in [0.16, 0.15, 0.14, 0.13, 0.12, 0.11, 0.10, 0.09, 0.08, 0.07, 0.06, 0.05]
    ]
    baseline_commission = 351_356.37
    for row, (name, scenario_type, dk_rate, resto_rate, stores_rate, stores_gmv, eater, direct, di) in enumerate(scenarios, 6):
        dk_gmv = 1_029_023.57
        resto_gmv = 34_036.47
        portfolio = dk_gmv + resto_gmv + stores_gmv
        commission = dk_gmv * dk_rate + resto_gmv * resto_rate + stores_gmv * stores_rate
        weighted = commission / portfolio
        cpl1 = commission + eater - direct
        cpl2 = cpl1 - di
        partner_saving = baseline_commission - commission
        ws.write(row - 1, 0, name, fmt["text"])
        ws.write(row - 1, 1, scenario_type, fmt["text"])
        ws.write(row - 1, 2, dk_rate, fmt["input_pct"])
        ws.write(row - 1, 3, resto_rate, fmt["input_pct"])
        ws.write(row - 1, 4, stores_rate, fmt["input_pct"])
        ws.write(row - 1, 5, dk_gmv, fmt["input_eur"])
        ws.write(row - 1, 6, resto_gmv, fmt["input_eur"])
        ws.write(row - 1, 7, stores_gmv, fmt["input_eur"])
        ws.write_formula(row - 1, 8, f"=SUM(F{row}:H{row})", fmt["eur"], portfolio)
        ws.write_formula(row - 1, 9, f"=(F{row}*C{row}+G{row}*D{row}+H{row}*E{row})/I{row}", fmt["pct"], weighted)
        ws.write_formula(row - 1, 10, f"=F{row}*C{row}+G{row}*D{row}+H{row}*E{row}", fmt["eur"], commission)
        ws.write(row - 1, 11, eater, fmt["eur"])
        ws.write(row - 1, 12, direct, fmt["eur"])
        ws.write_formula(row - 1, 13, f"=K{row}+L{row}-M{row}", fmt["green_eur"] if cpl1 >= 0 else fmt["bad_eur"], cpl1)
        ws.write_formula(row - 1, 14, f"=N{row}/I{row}", fmt["green_pct"] if cpl1 >= 0 else fmt["bad_pct"], cpl1 / portfolio)
        ws.write(row - 1, 15, di, fmt["eur"])
        ws.write_formula(row - 1, 16, f"=N{row}-P{row}", fmt["green_eur"] if cpl2 >= 0 else fmt["bad_eur"], cpl2)
        ws.write_formula(row - 1, 17, f"=Q{row}/I{row}", fmt["green_pct"] if cpl2 >= 0 else fmt["bad_pct"], cpl2 / portfolio)
        ws.write_formula(row - 1, 18, f"=$K$6-K{row}", fmt["green_eur"] if partner_saving >= 0 else fmt["bad_eur"], partner_saving)
    ws.autofilter(4, 0, 4 + len(scenarios), 18)
    ws.freeze_panes(5, 5)
    ws.merge_range("A26:S26", "How to read: source rows are the supplied calculator’s Scenarios tab. The Stores rate test keeps DK/resto GMV unchanged at the 2026 annual run-rate and increases Stores GMV independently to €852.8k. CP L1 and CP L2 are Bolt results before any new DF, Bolt+, free-delivery or launch investment.", fmt["amber"])
    ws.set_column("A:A", 34)
    ws.set_column("B:B", 25)
    ws.set_column("C:E", 12)
    ws.set_column("F:I", 16)
    ws.set_column("J:J", 14)
    ws.set_column("K:Q", 16)
    ws.set_column("R:R", 14)
    ws.set_column("S:S", 18)

    # Mix matrix
    ws = wb.add_worksheet("Mix matrix")
    ws.hide_gridlines(2)
    ws.merge_range("A1:M1", "Commission mix matrix — does shared 16% create a genuine saving?", fmt["title"])
    ws.merge_range("A3:M3", "All calculations use the same total GMV from Inputs. Change the yellow assumptions there to test another commercial baseline.", fmt["note"])
    headers = [
        "Stores share\nof combined GMV", "DK + resto share", "Stores GMV", "DK + resto GMV",
        "Current blended\n25% / 12%", "Growth blended\n20% / 12%", "Shared\n16%",
        "Saving: growth\nvs current", "Saving: shared\nvs current", "Δ shared vs growth",
        "Growth commission\n€", "Shared commission\n€", "Read-out",
    ]
    ws.write_row("A5", headers, fmt["header"])
    shares = [STORES_GMV / PORTFOLIO_GMV, 0.10, 0.25, 0.40, 0.50, 0.60, 0.75, 0.90]
    total = PORTFOLIO_GMV
    for row, store_share in enumerate(shares, 6):
        dk_share = 1 - store_share
        current = dk_share * 0.25 + store_share * 0.12
        growth = dk_share * 0.20 + store_share * 0.12
        shared = 0.16
        saving_growth = current - growth
        saving_shared = current - shared
        delta = shared - growth
        readout = "Shared 16% is better than 20% / 12%" if delta < 0 else "Shared 16% is worse than 20% / 12%"
        ws.write(row - 1, 0, store_share, fmt["input_pct"])
        ws.write_formula(row - 1, 1, f"=1-A{row}", fmt["pct"], dk_share)
        ws.write_formula(row - 1, 2, f"=Inputs!$B$5*A{row}", fmt["eur"], total * store_share)
        ws.write_formula(row - 1, 3, f"=Inputs!$B$5*B{row}", fmt["eur"], total * dk_share)
        ws.write_formula(row - 1, 4, f"=B{row}*Inputs!$B$6+A{row}*Inputs!$B$7", fmt["pct"], current)
        ws.write_formula(row - 1, 5, f"=B{row}*Inputs!$B$8+A{row}*Inputs!$B$9", fmt["pct"], growth)
        ws.write_formula(row - 1, 6, "=Inputs!$B$10", fmt["pct"], shared)
        ws.write_formula(row - 1, 7, f"=E{row}-F{row}", fmt["green_pct"], saving_growth)
        ws.write_formula(row - 1, 8, f"=E{row}-G{row}", fmt["green_pct"] if saving_shared >= 0 else fmt["bad_pct"], saving_shared)
        ws.write_formula(row - 1, 9, f"=G{row}-F{row}", fmt["bad_pct"] if delta > 0 else fmt["green_pct"], delta)
        ws.write_formula(row - 1, 10, f"=Inputs!$B$5*F{row}", fmt["eur"], total * growth)
        ws.write_formula(row - 1, 11, f"=Inputs!$B$5*G{row}", fmt["eur"], total * shared)
        ws.write(row - 1, 12, readout, fmt["amber"] if delta > 0 else fmt["text"])
    ws.merge_range("A15:M15", "Interpretation: shared 16% is better than 20% DK/resto + 12% Stores only when DK + resto is more than 50% of combined GMV. At 90% Stores / 10% DK + resto, 20% / 12% gives 12.8%; shared 16% is +3.2 p.p. worse.", fmt["amber"])
    ws.set_column("A:B", 15)
    ws.set_column("C:D", 16)
    ws.set_column("E:J", 14)
    ws.set_column("K:L", 16)
    ws.set_column("M:M", 34)

    # Fixed 16% DK / Stores-rate sensitivity
    ws = wb.add_worksheet("DK 16% + Stores rate")
    ws.hide_gridlines(2)
    ws.merge_range("A1:O1", "Alternative: DK + resto fixed at 16%, flexible Stores commission", fmt["title"])
    ws.merge_range("A3:O3", "This sheet shows both partner saving and Bolt contribution. A Stores rate below 12% is lower than the current Bolt-courier minimum and should be treated as a subsidy.", fmt["note"])
    headers = [
        "Stores share", "DK + resto share", "Stores rate", "DK + resto rate",
        "Blended rate", "Partner saving\nvs 25% / 12%", "Partner saving\nvs 20% / 12%",
        "Partner saving €\nvs current", "Portfolio commission €", "Stores CP L2\nbefore new investment",
        "DK + resto CP L2\nbefore new investment", "Portfolio CP L2", "Status", "Stores GMV", "DK + resto GMV",
    ]
    ws.write_row("A5", headers, fmt["header"])
    row = 6
    total = PORTFOLIO_GMV
    for store_share in [0.25, 0.50, 0.75, 0.90]:
        for stores_rate in [0.12, 0.13, 0.14, 0.15, 0.16]:
            dk_share = 1 - store_share
            current = dk_share * 0.25 + store_share * 0.12
            growth = dk_share * 0.20 + store_share * 0.12
            blended = dk_share * 0.16 + store_share * stores_rate
            saving_current = current - blended
            saving_growth = growth - blended
            stores_gmv = total * store_share
            dk_gmv = total * dk_share
            stores_factor = stores_gmv / STORES_GMV
            dk_factor = dk_gmv / DK_RESTO_GMV
            stores_cpl2 = stores_gmv * stores_rate + STORES_EATER * stores_factor - STORES_DIRECT * stores_factor - STORES_DI * stores_factor
            dk_cpl2 = dk_gmv * 0.16 + DK_RESTO_EATER * dk_factor - DK_RESTO_DIRECT * dk_factor - DK_RESTO_DI * dk_factor
            portfolio_cpl2 = stores_cpl2 + dk_cpl2
            status = "Bolt CP L2 positive" if portfolio_cpl2 >= 0 else "Bolt CP L2 negative before new investment"
            ws.write(row - 1, 0, store_share, fmt["input_pct"])
            ws.write_formula(row - 1, 1, f"=1-A{row}", fmt["pct"], dk_share)
            ws.write(row - 1, 2, stores_rate, fmt["input_pct"])
            ws.write_formula(row - 1, 3, "=Inputs!$B$11", fmt["pct"], 0.16)
            ws.write_formula(row - 1, 4, f"=A{row}*C{row}+B{row}*D{row}", fmt["pct"], blended)
            ws.write_formula(row - 1, 5, f"=B{row}*Inputs!$B$6+A{row}*Inputs!$B$7-E{row}", fmt["green_pct"] if saving_current >= 0 else fmt["bad_pct"], saving_current)
            ws.write_formula(row - 1, 6, f"=B{row}*Inputs!$B$8+A{row}*Inputs!$B$9-E{row}", fmt["green_pct"] if saving_growth >= 0 else fmt["bad_pct"], saving_growth)
            ws.write_formula(row - 1, 7, f"=F{row}*Inputs!$B$5", fmt["green_eur"] if saving_current >= 0 else fmt["bad_eur"], saving_current * total)
            ws.write_formula(row - 1, 8, f"=E{row}*Inputs!$B$5", fmt["eur"], blended * total)
            ws.write_formula(row - 1, 9, f"=N{row}/Inputs!$B$16*(C{row}*Inputs!$B$16+Inputs!$B$18-Inputs!$B$19-Inputs!$B$20)", fmt["green_eur"] if stores_cpl2 >= 0 else fmt["bad_eur"], stores_cpl2)
            ws.write_formula(row - 1, 10, f"=O{row}/Inputs!$B$25*(D{row}*Inputs!$B$25+Inputs!$B$26-Inputs!$B$27-Inputs!$B$28)", fmt["green_eur"] if dk_cpl2 >= 0 else fmt["bad_eur"], dk_cpl2)
            ws.write_formula(row - 1, 11, f"=J{row}+K{row}", fmt["green_eur"] if portfolio_cpl2 >= 0 else fmt["bad_eur"], portfolio_cpl2)
            ws.write(row - 1, 12, status, fmt["text"])
            ws.write_formula(row - 1, 13, f"=Inputs!$B$5*A{row}", fmt["eur"], stores_gmv)
            ws.write_formula(row - 1, 14, f"=Inputs!$B$5*B{row}", fmt["eur"], dk_gmv)
            row += 1
    ws.autofilter(4, 0, row - 2, 14)
    ws.freeze_panes(5, 4)
    ws.set_column("A:D", 14)
    ws.set_column("E:G", 15)
    ws.set_column("H:L", 18)
    ws.set_column("M:M", 32)
    ws.set_column("N:O", 17)

    # Stores profitability break-even
    ws = wb.add_worksheet("Stores profitability")
    ws.hide_gridlines(2)
    ws.merge_range("A1:H1", "Stores unit economics — break-even commission rate", fmt["title"])
    ws.merge_range("A3:H3", "This answers the Bolt question: at which Stores rate does CP L2 turn positive before new DF, Bolt+, free-delivery or launch costs?", fmt["note"])
    ws.write_row("A5", ["Metric", "Formula / treatment", "Value", "€/order", "% GMV", "Commercial meaning", "Editable?", "Note"], fmt["header"])
    aov = STORES_GMV / STORES_ORDERS
    eater_order = STORES_EATER / STORES_ORDERS
    direct_order = STORES_DIRECT / STORES_ORDERS
    di_order = STORES_DI / STORES_ORDERS
    breakeven_rate = (STORES_DIRECT + STORES_DI - STORES_EATER) / STORES_GMV
    rows = [
        ("AOV", "GMV / orders", aov, aov, 1.0, "Average delivered order value", "No", "Historical volume base"),
        ("Eater fees", "Eater-fee revenue", eater_order, eater_order, STORES_EATER / STORES_GMV, "Offsets delivery economics", "No", "€1.92 / order"),
        ("Direct costs", "Courier, payment, operations, support", direct_order, direct_order, STORES_DIRECT / STORES_GMV, "Costs before DI", "No", "Excludes demand incentives"),
        ("Existing DI", "Existing demand incentives", di_order, di_order, STORES_DI / STORES_GMV, "Existing promo cost", "No", "Before incremental campaign"),
        ("Break-even Stores commission", "(Direct costs + DI − eater fees) / GMV", breakeven_rate, breakeven_rate * aov, breakeven_rate, "CP L2 = €0 before new investment", "No", "Below this rate Bolt loses before incremental investment"),
        ("Current proposed Stores rate", "Commercial input", 0.12, 0.12 * aov, 0.12, "Minimum launch rate", "Yes", "Compare to break-even"),
        ("Shared rate", "Commercial input", 0.16, 0.16 * aov, 0.16, "Scale rate test", "Yes", "Compare to break-even"),
    ]
    for row, (metric, treatment, value, per_order, pct_gmv, meaning, editable, note) in enumerate(rows, 6):
        ws.write(row - 1, 0, metric, fmt["label"])
        ws.write(row - 1, 1, treatment, fmt["text"])
        is_rate = "rate" in metric.lower() or "commission" in metric.lower()
        ws.write(row - 1, 2, value, fmt["input_pct"] if editable == "Yes" else (fmt["pct"] if is_rate else fmt["eur"]))
        ws.write(row - 1, 3, per_order, fmt["eur"])
        ws.write(row - 1, 4, pct_gmv, fmt["pct"])
        ws.write(row - 1, 5, meaning, fmt["text"])
        ws.write(row - 1, 6, editable, fmt["label"])
        ws.write(row - 1, 7, note, fmt["text"])
    ws.merge_range("A15:H15", f"Initial result: Stores break-even is {breakeven_rate:.1%}. The 12% rate is below this pre-investment break-even; a lower rate is not commercially viable without a compensating fee, cost improvement or incremental volume efficiency.", fmt["amber"])
    ws.set_column("A:A", 31)
    ws.set_column("B:B", 34)
    ws.set_column("C:E", 15)
    ws.set_column("F:F", 30)
    ws.set_column("G:G", 11)
    ws.set_column("H:H", 30)

    # Executive recommendation
    ws = wb.add_worksheet("Recommendation")
    ws.hide_gridlines(2)
    ws.merge_range("A1:H1", "Recommendation — commercial logic to discuss with Silpo", fmt["title"])
    ws.write("A3", "What the model proves", fmt["section"])
    points = [
        ("1. Shared 16% is not universally better", "With 90% Stores / 10% DK + resto, the 20% / 12% model has a 12.8% blended rate. Shared 16% adds 3.2 p.p.; it cannot be offered as an automatic growth reward."),
        ("2. Avoid channel-shift incentives", "Any next step should require incremental combined GMV and non-declining DK + resto GMV / orders versus an agreed 3-month baseline. Growth in Stores alone is insufficient if it cannibalises DK/resto."),
        ("3. Stores rate has a hard floor", f"On the historical unit economics, Stores pre-investment CP L2 break-even is {breakeven_rate:.1%}. A rate below 12% loses money before new DF, free delivery or Bolt+ costs."),
        ("4. A workable path", "Keep 12% Stores + 20% DK/resto after the €33k Stores threshold. At sustained combined scale, discuss a Stores-rate review only after delivery costs, combined incremental GMV and investment payback are verified."),
        ("5. Investment is portfolio-led", "The larger share of Bolt demand investment belongs to DK + resto because Smart Promo, Sponsored Listing, co-funding and restaurant tools sit there. Investment should grow total BF demand, not subsidise a move into Stores."),
    ]
    for row, (title, body) in enumerate(points, 5):
        ws.write(row - 1, 0, title, fmt["label"])
        ws.merge_range(row - 1, 1, row - 1, 7, body, fmt["text"])
        ws.set_row(row - 1, 44)
    ws.set_column("A:A", 31)
    ws.set_column("B:H", 18)

    wb.close()
    print(OUT)


if __name__ == "__main__":
    build()
