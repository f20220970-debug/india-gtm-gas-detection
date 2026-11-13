"""Build india_gtm_model.xlsx: go-to-market model for portable multi-gas detectors in Indian factories.

Who to sell to first, at what price, through which channel, and when it breaks even.
Blue = sourced input, yellow = assumption, black = formula. Run: python build.py
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.utils import get_column_letter as L

F = "Arial"
BLUE = Font(name=F, color="0000FF")
BLACK = Font(name=F)
GREEN = Font(name=F, color="008000")
BOLD = Font(name=F, bold=True)
WHITE_B = Font(name=F, bold=True, color="FFFFFF")
TITLE = Font(name=F, bold=True, size=14)
HDR = PatternFill("solid", fgColor="1F3A5F")
YEL = PatternFill("solid", fgColor="FFFF00")
thin = Side(style="thin", color="BFBFBF")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
WRAP = Alignment(wrap_text=True, vertical="top")
INR = '#,##0;(#,##0);"-"'
LAKH = '#,##0.0;(#,##0.0);"-"'
PCT = '0.0%'

wb = Workbook()


def header(ws, row, labels, widths=None, start=1):
    for i, lab in enumerate(labels, start):
        c = ws.cell(row=row, column=i, value=lab)
        c.font, c.fill, c.border = WHITE_B, HDR, BOX
        c.alignment = Alignment(wrap_text=True, vertical="center")
    if widths:
        for i, w in enumerate(widths, start):
            ws.column_dimensions[L(i)].width = w


def put(ws, ref, value, font=BLACK, fmt=None, fill=None):
    c = ws[ref]
    c.value, c.font, c.border = value, font, BOX
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    return c


# ---------------------------------------------------------------- README
ws = wb.active
ws.title = "README"
ws.column_dimensions["A"].width = 115
for i, (t, f) in enumerate([
    ("India go-to-market model: portable multi-gas detectors for factories", TITLE),
    ("Built by Pavitra Seth (BITS Pilani, B.E. Mechanical + M.Sc. Chemistry) as preparation for Honeywell PS-II. Public data only; not affiliated with Honeywell.", BLACK),
    ("", BLACK),
    ("Job to be done: before a worker enters a tank, pit or reactor, the safety officer needs to know whether the air is safe (oxygen, flammable gas, CO, H2S).", BOLD),
    ("Buyer: EHS / safety manager or plant head. User: operator or maintenance technician. Triggers: statutory audits, near-misses, customer and export audits, insurance.", BLACK),
    ("", BLACK),
    ("Tabs: Inputs > Segments (how many factories, how many detectors) > Pricing (13 market price points, 3 tiers) > Channel (mix and margins) > "
     "Plan (quarterly funnel and P&L for 3 years) > UnitEconomics (CAC, payback, lifetime value) > Sensitivity > Launch (90-day plan, KPIs, risks) > Sources.", BLACK),
    ("Colour code: blue = sourced; yellow = assumption (change it); black = formula; green = link to another tab.", BLACK),
    ("All money in INR. Lakh = 100,000; crore = 10,000,000.", BLACK),
], 1):
    c = ws.cell(row=i, column=1, value=t)
    c.font = f
    c.alignment = Alignment(wrap_text=True)

# ---------------------------------------------------------------- Inputs
inp = wb.create_sheet("Inputs")
inp["A1"] = "Inputs"
inp["A1"].font = TITLE
header(inp, 3, ["Input", "Value", "Unit", "Type", "Source / reasoning"], [50, 12, 14, 12, 90])
R = {}
rows = [
    ("sec", "PRODUCT AND RECURRING REVENUE"),
    ("cogs", "Device cost of goods as % of list price", 0.45, "%", "Assumption", "Imported sensors + local assembly"),
    ("svc_price", "Calibration and bump-test service per device per year", 4000, "INR", "Assumption", "Annual service contract"),
    ("svc_margin", "Service gross margin", 0.60, "%", "Assumption", "Mostly technician time"),
    ("svc_attach", "Share of devices sold that take a service contract", 0.50, "%", "Assumption", ""),
    ("sensor_price", "Sensor replacement per device every 2 years", 8000, "INR", "Assumption", "Electrochemical sensors typically last 2-3 years"),
    ("sensor_margin", "Sensor gross margin", 0.50, "%", "Assumption", ""),
    ("life", "Customer life", 5, "years", "Assumption", "Device life before replacement"),
    ("sec", "FUNNEL"),
    ("leads_q", "Qualified leads per active cluster per quarter", 250, "leads", "Assumption", "Industrial-area lists, safety-officer associations, distributor referrals"),
    ("demo_rate", "Lead to demo", 0.30, "%", "Assumption", ""),
    ("win_rate", "Demo to order", 0.30, "%", "Assumption", ""),
    ("units_order", "Detectors per first order", 4, "units", "Assumption", "Matches the units-per-factory estimate on Segments"),
    ("repeat", "Repeat orders per customer per year (extra units)", 1, "units", "Assumption", "New lines, lost or damaged units"),
    ("sec", "COSTS"),
    ("rep_cost", "Cost per sales engineer per year (loaded)", 1200000, "INR", "Assumption", "Salary, travel, demo kit"),
    ("reps_cluster", "Sales engineers per active cluster", 2, "people", "Assumption", ""),
    ("mkt_q", "Marketing per active cluster per quarter", 300000, "INR", "Assumption", "Safety workshops, trade fairs, digital"),
    ("fixed_q", "Fixed overhead per quarter (product, certification, support)", 1500000, "INR", "Assumption", "Small launch team"),
    ("launch", "One-time launch cost (certification, training, demo fleet)", 4000000, "INR", "Assumption", "Hazardous-area certification and local approvals"),
]
r = 4
for row in rows:
    if row[0] == "sec":
        inp.cell(row=r, column=1, value=row[1]).font = BOLD
        r += 1
        continue
    key, lab, v, unit, typ, src = row
    R[key] = r
    put(inp, f"A{r}", lab)
    put(inp, f"B{r}", v, BLUE, PCT if unit == "%" else INR, YEL if typ == "Assumption" else None)
    put(inp, f"C{r}", unit)
    put(inp, f"D{r}", typ)
    put(inp, f"E{r}", src).alignment = WRAP
    r += 1


def I(k):
    return f"Inputs!$B${R[k]}"


# ---------------------------------------------------------------- Segments
sg = wb.create_sheet("Segments")
sg["A1"] = "Who needs detectors: factories in operation (ASI 2023-24) and detector demand"
sg["A1"].font = TITLE
header(sg, 3, ["Segment (NIC)", "Factories in operation", "Share with confined-space / gas risk", "Target factories",
               "Detectors per factory", "Detector need (units)", "Share already equipped", "Annual market (units)",
               "Priority score (1-5)", "Why"], [34, 14, 14, 12, 12, 13, 12, 13, 11, 60])
segs = [
    ("Chemicals and chemical products (20)", 12148, 0.80, 4, 0.50, 5, "Solvents, reactors, tanks; MSIHC Rules apply; dense clusters (Vapi, Ankleshwar, Dahej)"),
    ("Pharmaceuticals (21)", 4923, 0.70, 3, 0.60, 4, "Solvent recovery and API plants; strong audit culture from export customers"),
    ("Coke and refined petroleum (19)", 1538, 1.00, 20, 0.90, 2, "High need, but large refiners already buy global brands direct"),
    ("Food products (10): ammonia cold storage", 35081, 0.05, 2, 0.30, 3, "Ammonia refrigeration leaks; many small, price-sensitive owners"),
    ("Basic metals (24)", 10834, 0.30, 4, 0.50, 3, "CO in blast furnace and gas-plant areas"),
]
for i, (name, n, share, per, eq, pri, why) in enumerate(segs):
    rr = 4 + i
    put(sg, f"A{rr}", name, BOLD)
    put(sg, f"B{rr}", n, BLUE, INR)
    put(sg, f"C{rr}", share, BLUE, PCT, YEL)
    put(sg, f"D{rr}", f"=B{rr}*C{rr}", BLACK, INR)
    put(sg, f"E{rr}", per, BLUE, INR, YEL)
    put(sg, f"F{rr}", f"=D{rr}*E{rr}", BLACK, INR)
    put(sg, f"G{rr}", eq, BLUE, PCT, YEL)
    put(sg, f"H{rr}", f"=F{rr}*G{rr}/$B$13+F{rr}*(1-G{rr})*$B$14", BLACK, INR)
    put(sg, f"I{rr}", pri, BLUE, "0", YEL)
    put(sg, f"J{rr}", why).alignment = WRAP
put(sg, "A9", "Total", BOLD)
for col in "BDFH":
    put(sg, f"{col}9", f"=SUM({col}4:{col}8)", BLACK, INR)
put(sg, "A11", "Market assumptions", BOLD)
put(sg, "A13", "Device replacement cycle (years)")
put(sg, "B13", 4, BLUE, "0", YEL)
put(sg, "A14", "Share of unequipped need converting each year")
put(sg, "B14", 0.08, BLUE, PCT, YEL)
put(sg, "A15", "Annual market value at standard-tier price (INR crore)", BOLD)
put(sg, "B15", "=H9*Pricing!C17/10000000", BLACK, LAKH)
put(sg, "A16", "Cross-check: hazardous-process factories registered (DGFASLI, 2023)")
put(sg, "B16", 45374, BLUE, INR)
put(sg, "C16", "=D9/B16", BLACK, PCT)
put(sg, "D16", "target factories as share of hazardous-process factories")
put(sg, "A18", "First segment to attack: highest priority score", BOLD)
put(sg, "B18", "=INDEX(A4:A8,MATCH(MAX(I4:I8),I4:I8,0))")
sg["A20"] = ("Factory counts: MoSPI Annual Survey of Industries 2023-24 (sample-based estimates). Shares, detectors per factory and "
             "equipped share are my assumptions; test them in customer interviews.")
sg["A20"].font = Font(name=F, italic=True)

# ---------------------------------------------------------------- Pricing
pr = wb.create_sheet("Pricing")
pr["A1"] = "Price benchmarks (portable gas detectors on IndiaMART, Nov 2026) and chosen tiers"
pr["A1"].font = TITLE
header(pr, 3, ["Product (as listed)", "Brand", "Price (INR)", "Tier", "Notes"], [46, 22, 12, 12, 50])
bench = [
    ("Atmospheric oxygen meter, multi-gas detector 8812", "Unbranded", 12500, "Value", "H2S, O2, CO, LEL; pump"),
    ("Multi-gas monitor for CO, H2S, LEL, O2", "Unbranded (8811)", 13000, "Value", ""),
    ("Multi gas detector 8812 Workzone", "Workzone", 13500, "Value", ""),
    ("4-gas multi gas detector with LCD", "Honeywell", 20000, "Mid", "Listed for confined-space use"),
    ("Bosean multi gas detector", "Bosean", 28000, "Mid", "LEL, CO, O2, H2S"),
    ("Honeywell BW Flex multi gas detector", "Honeywell", 39000, "Premium", ""),
    ("Multi gas detector portable (pump)", "Unbranded", 45000, "Premium", "H2S, O2, LEL, CO"),
    ("Draeger X-am 5600", "Draeger", 45890, "Premium", "Up to 6 gases"),
    ("Venits MX4 multi gas detector", "Industrial Scientific", 70000, "Premium", ""),
]
for i, (p, b, v, t, n) in enumerate(bench):
    rr = 4 + i
    put(pr, f"A{rr}", p)
    put(pr, f"B{rr}", b)
    put(pr, f"C{rr}", v, BLUE, INR)
    put(pr, f"D{rr}", t)
    put(pr, f"E{rr}", n)
put(pr, "A14", "Benchmark median / min / max", BOLD)
put(pr, "C14", "=MEDIAN(C4:C12)", BLACK, INR)
put(pr, "D14", "=MIN(C4:C12)", BLACK, INR)
put(pr, "E14", "=MAX(C4:C12)", BLACK, INR)
header(pr, 16, ["Our tier", "Positioning", "List price (INR)", "Share of unit sales", "Why this price"])
tiers = [
    ("Standard", "Certified 4-gas, rugged, bump-test dock", 32000, 0.60, "Below BW Flex and Draeger; above unbranded with certification and service"),
    ("Value", "4-gas, diffusion, no dock", 18000, 0.30, "Answers unbranded at 12.5-13.5k with certification and local service"),
    ("Connected", "Standard + Bluetooth / cloud log for audits", 52000, 0.10, "Audit trail sells to pharma exporters"),
]
for i, (t, pos, price, share, why) in enumerate(tiers):
    rr = 17 + i
    put(pr, f"A{rr}", t, BOLD)
    put(pr, f"B{rr}", pos)
    put(pr, f"C{rr}", price, BLUE, INR, YEL)
    put(pr, f"D{rr}", share, BLUE, PCT, YEL)
    put(pr, f"E{rr}", why).alignment = WRAP
put(pr, "A21", "Blended list price per unit", BOLD)
put(pr, "C21", "=SUMPRODUCT(C17:C19,D17:D19)", BLACK, INR)
put(pr, "D21", '=IF(ABS(SUM(D17:D19)-1)<0.001,"mix sums to 100%","CHECK MIX")')
put(pr, "A23", "Value ceiling: what one avoided incident is worth (illustrative)", BOLD)
put(pr, "A24", "Days of plant shutdown after a serious incident")
put(pr, "C24", 10, BLUE, "0", YEL)
put(pr, "A25", "Contribution lost per day, small chemical unit (INR)")
put(pr, "C25", 300000, BLUE, INR, YEL)
put(pr, "A26", "Value of avoiding one incident vs a 4-detector order", BOLD)
put(pr, "C26", "=C24*C25/(4*C21)", BLACK, '0.0"x"')
put(pr, "E26", "Even one avoided shutdown pays for the detectors many times over: sell on risk, not on spec")

# ---------------------------------------------------------------- Channel
ch = wb.create_sheet("Channel")
ch["A1"] = "Channel mix and economics"
ch["A1"].font = TITLE
header(ch, 3, ["Channel", "Margin given away", "Share Y1", "Share Y2", "Share Y3", "Best for"], [34, 14, 10, 10, 10, 60])
chans = [("Direct key accounts", 0.0, 0.50, 0.35, 0.25, "Large pharma and chemical sites; multi-site deals"),
         ("Industrial safety distributors", 0.25, 0.40, 0.50, 0.55, "Cluster coverage at low cost to serve"),
         ("GeM and e-commerce", 0.10, 0.10, 0.15, 0.20, "Government units and repeat buys")]
for i, (n, m, a, b, c, why) in enumerate(chans):
    rr = 4 + i
    put(ch, f"A{rr}", n, BOLD)
    put(ch, f"B{rr}", m, BLUE, PCT, YEL)
    for col, v in zip("CDE", (a, b, c)):
        put(ch, f"{col}{rr}", v, BLUE, PCT, YEL)
    put(ch, f"F{rr}", why)
put(ch, "A8", "Blended margin given to channel", BOLD)
for col in "CDE":
    put(ch, f"{col}8", f"=SUMPRODUCT($B$4:$B$6,{col}4:{col}6)", BLACK, PCT)
put(ch, "A9", "Mix check", BOLD)
for col in "CDE":
    put(ch, f"{col}9", f'=IF(ABS(SUM({col}4:{col}6)-1)<0.001,"OK","CHECK")')

# ---------------------------------------------------------------- Plan (quarterly)
pl = wb.create_sheet("Plan")
pl["A1"] = "Quarterly funnel and P&L, 3 years (INR lakh unless stated)"
pl["A1"].font = TITLE
Q = 12
qcols = [L(2 + i) for i in range(Q)]  # B..M
header(pl, 3, ["Line"] + [f"Y{q // 4 + 1} Q{q % 4 + 1}" for q in range(Q)] + ["Total"], [44] + [10] * Q + [11])
clusters = [1, 1, 1, 1, 3, 3, 3, 3, 5, 5, 5, 5]
lines = [
    ("Active clusters", None, "0"),
    ("Qualified leads", "={c}4*" + I("leads_q"), INR),
    ("Demos", "={c}5*" + I("demo_rate"), INR),
    ("New customers", "={c}6*" + I("win_rate"), INR),
    ("Cumulative customers", "={c}7", INR),
    ("Units: first orders", "={c}7*" + I("units_order"), INR),
    ("Units: repeat orders", "={p}8*" + I("repeat") + "/4", INR),
    ("Units sold", "={c}9+{c}10", INR),
    ("Installed base (units)", "={c}11", INR),
    ("Channel margin (year mix)", None, PCT),
    ("Device revenue, net of channel", "={c}11*Pricing!$C$21*(1-{c}13)/100000", LAKH),
    ("Service revenue", "={p}12*" + I("svc_attach") + "*" + I("svc_price") + "/4/100000", LAKH),
    ("Sensor revenue", "={p}12*" + I("sensor_price") + "/8/100000", LAKH),
    ("TOTAL REVENUE", "={c}14+{c}15+{c}16", LAKH),
    ("Gross profit", "={c}11*Pricing!$C$21*(1-{c}13-" + I("cogs") + ")/100000+{c}15*" + I("svc_margin") + "+{c}16*" + I("sensor_margin"), LAKH),
    ("Gross margin", "=IFERROR({c}18/{c}17,0)", PCT),
    ("Sales engineers", "={c}4*" + I("reps_cluster") + "*" + I("rep_cost") + "/4/100000", LAKH),
    ("Marketing", "={c}4*" + I("mkt_q") + "/100000", LAKH),
    ("Fixed overhead", "=" + I("fixed_q") + "/100000", LAKH),
    ("One-time launch cost", None, LAKH),
    ("Operating profit", "={c}18-{c}20-{c}21-{c}22-{c}23", LAKH),
    ("Cumulative operating profit", "={c}24", LAKH),
]
for i, (lab, f, fmt) in enumerate(lines):
    rr = 4 + i
    put(pl, f"A{rr}", lab, BOLD if lab.isupper() or lab.startswith("Operating") else BLACK)
    for j, c in enumerate(qcols):
        p = qcols[j - 1] if j else None
        if lab == "Active clusters":
            v, font = clusters[j], BLUE
            put(pl, f"{c}{rr}", v, font, fmt, YEL)
            continue
        if lab == "Channel margin (year mix)":
            put(pl, f"{c}{rr}", f"=Channel!{'CDE'[j // 4]}8", GREEN, fmt)
            continue
        if lab == "One-time launch cost":
            put(pl, f"{c}{rr}", "=" + I("launch") + "/100000" if j == 0 else 0, BLACK, fmt)
            continue
        if p is None and ("{p}" in f):
            ff = f.replace("{p}8", "0").replace("{p}12", "0")
        else:
            ff = f
        if lab == "Cumulative customers":
            ff = f"={c}7" if j == 0 else f"={p}8+{c}7"
        if lab == "Installed base (units)":
            ff = f"={c}11" if j == 0 else f"={p}12+{c}11"
        if lab == "Cumulative operating profit":
            ff = f"={c}24" if j == 0 else f"={p}25+{c}24"
        put(pl, f"{c}{rr}", ff.format(c=c, p=p or "B"), BLACK, fmt)
    if fmt in (LAKH, INR) and lab not in ("Cumulative customers", "Installed base (units)", "Cumulative operating profit", "Active clusters"):
        put(pl, f"N{rr}", f"=SUM(B{rr}:M{rr})", BLACK, fmt)
assert pl["A24"].value == "Operating profit" and pl["A17"].value == "TOTAL REVENUE"
put(pl, "A27", "Annual summary", BOLD)
header(pl, 28, ["Metric", "Year 1", "Year 2", "Year 3"], None)
summ = [("Revenue (INR lakh)", 17, LAKH), ("Gross profit", 18, LAKH), ("Operating profit", 24, LAKH), ("Units sold", 11, INR), ("New customers", 7, INR)]
for i, (lab, src, fmt) in enumerate(summ):
    rr = 29 + i
    put(pl, f"A{rr}", lab, BOLD)
    for y in range(3):
        a, b = qcols[4 * y], qcols[4 * y + 3]
        put(pl, f"{'BCD'[y]}{rr}", f"=SUM({a}{src}:{b}{src})", BLACK, fmt)
put(pl, "A34", "Gross margin", BOLD)
for y in range(3):
    put(pl, f"{'BCD'[y]}34", f"={'BCD'[y]}30/{'BCD'[y]}29", BLACK, PCT)
put(pl, "A35", "Year-3 share of annual market (units)", BOLD)
put(pl, "D35", "=D32/Segments!H9", BLACK, PCT)
put(pl, "A36", "First quarter with positive operating profit", BOLD)
put(pl, "B36", '=IFERROR(INDEX($B$3:$M$3,MATCH(TRUE,INDEX($B$24:$M$24>0,0),0)),"not in 3 years")')
put(pl, "A37", "Quarter cumulative profit turns positive (payback)", BOLD)
put(pl, "B37", '=IFERROR(INDEX($B$3:$M$3,MATCH(TRUE,INDEX($B$25:$M$25>0,0),0)),"not in 3 years")')
put(pl, "A38", "Peak cash need (most negative cumulative, INR lakh)", BOLD)
put(pl, "B38", "=MIN(B25:M25)", BLACK, LAKH)

lc = LineChart()
lc.title = "Quarterly revenue vs operating profit (INR lakh)"
lc.add_data(Reference(pl, min_col=1, max_col=13, min_row=17), from_rows=True, titles_from_data=True)
lc.add_data(Reference(pl, min_col=1, max_col=13, min_row=24), from_rows=True, titles_from_data=True)
lc.add_data(Reference(pl, min_col=1, max_col=13, min_row=25), from_rows=True, titles_from_data=True)
lc.set_categories(Reference(pl, min_col=2, max_col=13, min_row=3))
lc.height, lc.width = 9, 22
lc.x_axis.delete = False
lc.y_axis.delete = False
pl.add_chart(lc, "F27")

# ---------------------------------------------------------------- UnitEconomics
ue = wb.create_sheet("UnitEconomics")
ue["A1"] = "Unit economics per customer (Year 3 run-rate)"
ue["A1"].font = TITLE
header(ue, 3, ["Metric", "Value", "Unit", "How"], [46, 14, 12, 60])
uel = [
    ("Sales + marketing spend, Year 3", "=SUM(Plan!J20:M21)*100000", INR, "Plan rows: sales engineers + marketing"),
    ("New customers, Year 3", "=SUM(Plan!J7:M7)", INR, ""),
    ("Customer acquisition cost (CAC)", "=B4/B5", INR, "spend / new customers"),
    ("First-order device gross profit", f"={I('units_order')}*Pricing!C21*(1-Channel!E8-{I('cogs')})", INR, "units x price x (1 - channel - COGS)"),
    ("Recurring gross profit per customer per year", f"={I('units_order')}*({I('svc_attach')}*{I('svc_price')}*{I('svc_margin')}+{I('sensor_price')}/2*{I('sensor_margin')})+{I('repeat')}*Pricing!C21*(1-Channel!E8-{I('cogs')})", INR, "service + sensors + repeat units"),
    ("Lifetime gross profit per customer", f"=B7+B8*({I('life')}-1)", INR, "first order + recurring for remaining years"),
    ("LTV / CAC", "=B9/B6", '0.0"x"', "above 3x is healthy"),
    ("Months to recover CAC", "=IF(B7>=B6,0,(B6-B7)/(B8/12))", '0.0', "0 = first order already pays back CAC; else months of recurring profit needed"),
    ("CAC as % of first-order value", f"=B6/({I('units_order')}*Pricing!C21)", PCT, "should be well below 100%"),
]
for i, (lab, f, fmt, how) in enumerate(uel):
    rr = 4 + i
    put(ue, f"A{rr}", lab, BOLD)
    put(ue, f"B{rr}", f, BLACK, fmt)
    put(ue, f"D{rr}", how)

# ---------------------------------------------------------------- Sensitivity
se = wb.create_sheet("Sensitivity")
se["A1"] = "Year-3 operating profit (INR lakh) by standard-tier price and demo-to-order win rate"
se["A1"].font = TITLE
se["A2"] = "Each cell re-scales Year-3 device gross profit and units from the Plan tab (service and costs held), so it stays live."
prices = [24000, 28000, 32000, 36000, 40000]
wins = [0.20, 0.25, 0.30, 0.35, 0.40]
put(se, "A4", "Win rate (down) / standard price (across)", BOLD)
se.column_dimensions["A"].width = 36
for j, p in enumerate(prices):
    put(se, f"{L(2 + j)}4", p, BLUE, INR)
    se.column_dimensions[L(2 + j)].width = 12
for i, w in enumerate(wins):
    rr = 5 + i
    put(se, f"A{rr}", w, BLUE, PCT)
    for j in range(len(prices)):
        c = L(2 + j)
        # blended price with this standard price
        bp = f"(Pricing!$C$18*Pricing!$D$18+{c}$4*Pricing!$D$17+Pricing!$C$19*Pricing!$D$19)"
        units = f"(Plan!$D$32*$A{rr}/{I('win_rate')})"
        dev_gp = f"({units}*{bp}*(1-Channel!$E$8-{I('cogs')})/100000)"
        base_dev_gp = f"(Plan!$D$32*Pricing!$C$21*(1-Channel!$E$8-{I('cogs')})/100000)"
        put(se, f"{c}{rr}", f"=Plan!$D$31-{base_dev_gp}+{dev_gp}", BLACK, LAKH)
se.conditional_formatting.add("B5:F9", ColorScaleRule(start_type="min", start_color="F8696B", mid_type="num", mid_value=0, mid_color="FFFFFF", end_type="max", end_color="63BE7B"))
put(se, "A11", "Check: base cell (32,000 / 30%) equals Plan Year-3 operating profit", BOLD)
put(se, "D11", "=D7-Plan!D31", BLACK, LAKH)

# ---------------------------------------------------------------- Launch
la = wb.create_sheet("Launch")
la["A1"] = "Launch plan, KPIs and risks"
la["A1"].font = TITLE
header(la, 3, ["Phase", "When", "What", "KPI to pass before moving on"], [18, 14, 70, 50])
phases = [
    ("Pilot", "Days 0-90", "One cluster (Vapi-Ankleshwar-Dahej chemical belt). 2 sales engineers, 1 distributor, 20 free 30-day trials with EHS managers, safety workshops with industrial associations.", "30 paying customers; trial-to-order >= 40%; service attach >= 50%"),
    ("Expand", "Year 2", "Add Hyderabad pharma and Baddi pharma clusters; sign 2 more distributors; list on GeM.", "CAC payback < 12 months; gross margin >= 45%"),
    ("Scale", "Year 3", "5 clusters; launch the Connected tier for pharma exporters (audit logs).", "Year-3 operating profit positive; share of annual market < 10% (headroom)"),
]
for i, row in enumerate(phases):
    for j, v in enumerate(row):
        c = la.cell(row=4 + i, column=1 + j, value=v)
        c.font, c.alignment, c.border = (BOLD if j == 0 else BLACK), WRAP, BOX
    la.row_dimensions[4 + i].height = 45
header(la, 9, ["Risk", "Likelihood", "Mitigation", ""], None)
risks = [
    ("Price war with unbranded detectors at 12-14k", "High", "Sell certification, service and audit logs; Value tier at 18k"),
    ("Hazardous-area certification delays", "Medium", "Start approvals before launch; launch cost budgets for it"),
    ("Distributor pushes competitor brands", "Medium", "Exclusive territories, margin tied to service attach"),
    ("Customers buy once, skip calibration", "High", "Bundle first-year service; bump-test dock; audit reminders"),
    ("Import duty or sensor supply shocks", "Medium", "Dual-source sensors; local assembly"),
]
for i, row in enumerate(risks):
    for j, v in enumerate(row):
        c = la.cell(row=10 + i, column=1 + j, value=v)
        c.font, c.alignment, c.border = BLACK, WRAP, BOX

# ---------------------------------------------------------------- Sources
so = wb.create_sheet("Sources")
header(so, 1, ["Source", "Used for", "Link"], [50, 40, 100])
srcs = [
    ("MoSPI: Annual Survey of Industries 2023-24, Summary Results", "Factories in operation by NIC 2-digit and state", "https://mospi.gov.in/uploads/publications_reports/publications_reports1759229789833_2b8ed8ad-314e-46b7-b123-b08cb17f1233_ASI_Summary_Results_2023-24C.pdf"),
    ("Dataful (DGFASLI data): workplace injuries in Indian factories", "45,374 hazardous-process factories (2023); injury rates", "https://insights.dataful.in/articles/fewer-workplace-injuries-but-fatal-risks-persist-in-indian-factories"),
    ("IndiaMART: multi gas detector listings", "Price benchmarks (Nov 2026)", "https://m.indiamart.com/impcat/multi-gas-detector.html"),
    ("CPCB: Manufacture, Storage and Import of Hazardous Chemical Rules, 1989", "Regulatory driver for chemical units", "https://cpcb.nic.in"),
    ("PESO (Petroleum and Explosives Safety Organisation)", "Hazardous-area approvals", "https://peso.gov.in"),
    ("GeM (Government e-Marketplace)", "Government channel", "https://gem.gov.in"),
]
for i, s in enumerate(srcs):
    for j, v in enumerate(s):
        c = so.cell(row=2 + i, column=j + 1, value=v)
        c.font, c.alignment, c.border = BLACK, WRAP, BOX

for sh in wb.worksheets:
    sh.sheet_view.showGridLines = False
wb.save("india_gtm_model.xlsx")
print("saved india_gtm_model.xlsx")
