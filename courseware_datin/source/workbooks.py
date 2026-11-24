"""Build labelled starter workbooks and an interactive classroom checkpoint."""

from collections import defaultdict

import xlsxwriter

from datasets import CHANNELS, CHECKPOINT, FIRST, FRESH, HEADERS, PRODUCTS, REGIONS, records, select, totals

MONEY = '"GHS " #,##0'
NAVY, PANEL, GOLD, WHITE, MUTED = "#0B1730", "#152642", "#E9BE64", "#FFFFFF", "#B9C9E2"
PRODUCT_NAMES = [product for category in PRODUCTS for product in category]


def write_sales(book, rows):
    sheet = book.add_worksheet("Sales Data")
    sheet.freeze_panes(1, 3)
    sheet.set_column("A:A", 16)
    sheet.set_column("B:C", 17)
    sheet.set_column("D:F", 19)
    sheet.set_column("G:G", 10)
    sheet.set_column("H:L", 23)
    sheet.set_column("M:M", 22)
    date = book.add_format({"num_format": "dd mmm yyyy"})
    amount = book.add_format({"num_format": MONEY})
    formats = {1: date, 2: date, **dict.fromkeys(range(7, 12), amount)}
    sheet.add_table(0, 0, len(rows), len(HEADERS) - 1, {
        "name": "tblDatinSales", "style": "Table Style Medium 2", "data": rows,
        "columns": [{"header": label, "format": formats.get(index)} for index, label in enumerate(HEADERS)],
    })


def write_checks(book, period, rows):
    sheet = book.add_worksheet("Data Guide")
    sheet.hide_gridlines(2)
    title = book.add_format({"font_size": 18, "bold": True, "font_color": NAVY})
    text = book.add_format({"text_wrap": True, "valign": "vcenter", "font_color": NAVY})
    header = book.add_format({"bg_color": NAVY, "font_color": WHITE, "bold": True, "text_wrap": True})
    amount = book.add_format({"num_format": MONEY})
    rate = book.add_format({"num_format": "0.00%"})
    sheet.set_column("A:A", 43)
    sheet.set_column("B:G", 18)
    sheet.merge_range("A1:G2", "Datin practice data | " + ("Jan–Jun 2026" if period == FIRST else "Jul–Dec 2026"), title)
    sheet.merge_range("A4:G5", "Synthetic orders for a fictional Ghana distributor. Source is tblDatinSales on the Sales Data sheet. One row per Order_ID. All money values are numeric and denominated in Ghana cedi (GHS).", text)
    sheet.merge_range("A7:G8", "Net sales = units × list price − discount. Gross profit = net sales − cost of goods sold (COGS). Gross margin = total gross profit ÷ total net sales. Check the dashboard against the totals below.", text)
    sheet.write_row("A10", ["Filter selection", "Order count", "Units", "Net sales", "COGS", "Gross profit", "Gross margin"], header)
    for index, (label, filtered) in enumerate([
        ("All", rows), ("Greater Accra / Wholesale", select(rows, region=REGIONS[0], channel=CHANNELS[0])),
        ("Greater Accra / Wholesale / Rice", select(rows, region=REGIONS[0], channel=CHANNELS[0], product="Rice")),
    ], start=10):
        results = totals(filtered)
        sheet.write(index, 0, label)
        sheet.write_row(index, 1, [results["orders"], results["units"]])
        for col, key in enumerate(("net", "cost", "gross"), 3):
            sheet.write_number(index, col, results[key], amount)
        sheet.write_number(index, 6, results["margin"], rate)
    sheet.merge_range("A15:G16", "Open your working copy in Excel with Copilot. Run the prompt listed in the lab PDF. Keep the source table and this guide available to reconcile the result.", text)


def write_dashboard(book, rows):
    sheet = book.add_worksheet("Sales Dashboard")
    sheet.activate()
    sheet.hide_gridlines(2)
    sheet.set_zoom(85)
    sheet.set_column("A:A", 3)
    sheet.set_column("B:M", 11)
    sheet.set_column("N:N", 3)
    sheet.set_column("O:V", 17, None, {"hidden": True})
    fill = book.add_format({"bg_color": NAVY})
    for row in range(42):
        sheet.set_row(row, 19)
        for col in range(14):
            sheet.write_blank(row, col, None, fill)

    def style(bg=PANEL, fg=WHITE, size=11, bold=False, **options):
        return book.add_format({
            "bg_color": bg, "font_color": fg, "font_name": "Calibri",
            "font_size": size, "bold": bold, "valign": "vcenter", **options,
        })

    sheet.merge_range("B2:M3", "Sales performance | January–June 2026", style(NAVY, WHITE, 23, True))
    sheet.merge_range("B4:M4", "Datin class example · fictional wholesale distributor in Tema, Ghana · GHS", style(NAVY, MUTED))
    for label_range, label, cell_range, cell, choices in [
        ("B6:C6", "Region", "D6:E6", "D6", REGIONS),
        ("F6:G6", "Channel", "H6:I6", "H6", CHANNELS),
        ("J6:K6", "Product", "L6:M6", "L6", PRODUCT_NAMES),
    ]:
        sheet.merge_range(label_range, label, style(NAVY, MUTED, bold=True))
        sheet.merge_range(cell_range, "All", style(PANEL, GOLD, bold=True, indent=1))
        sheet.data_validation(cell, {"validate": "list", "source": ["All", *choices], "error_type": "stop"})

    # Every measure, summary and chart reads the same three filters. The All
    # condition matches all non-empty labels in the source table.
    filt = ",".join(
        f'tblDatinSales[{field}],IF(${cell}$6="All","<>",${cell}$6)'
        for field, cell in (("Region", "D"), ("Channel", "H"), ("Product", "L"))
    )
    all_values = totals(rows)
    for label_band, value_band, cell, label, formula, key, number in [
        ("B8:D8", "B9:D11", "B9", "Net sales (revenue)", f"=SUMIFS(tblDatinSales[Net_Sales_GHS],{filt})", "net", MONEY),
        ("E8:G8", "E9:G11", "E9", "Gross profit", f"=SUMIFS(tblDatinSales[Gross_Profit_GHS],{filt})", "gross", MONEY),
        ("H8:J8", "H9:J11", "H9", "Gross margin", '=IF(B9=0,"–",E9/B9)', "margin", "0.00%"),
        ("K8:M8", "K9:M11", "K9", "Order count", f"=COUNTIFS({filt})", "orders", "#,##0"),
    ]:
        sheet.merge_range(label_band, label, style(PANEL, MUTED, bold=True, align="center"))
        value = style(PANEL, GOLD, 17, True, align="center", num_format=number)
        sheet.merge_range(value_band, "", value)
        sheet.write_formula(cell, formula, value, all_values[key])

    month_totals = defaultdict(int)
    product_totals = defaultdict(int)
    region_totals = defaultdict(int)
    channel_totals = defaultdict(int)
    for row in rows:
        month_totals[row[2]] += row[9]
        product_totals[row[12]] += row[9]
        region_totals[row[3]] += row[9]
        channel_totals[row[4]] += row[9]
    for row, (month, total) in enumerate(sorted(month_totals.items()), 1):
        sheet.write_datetime(row, 14, month, book.add_format({"num_format": "mmm"}))
        formula = f"=SUMIFS(tblDatinSales[Net_Sales_GHS],{filt},tblDatinSales[Month_Start],O{row + 1})"
        sheet.write_formula(row, 15, formula, None, total)
    for names, field, column, letter, summary in [
        (PRODUCT_NAMES, "Product", 16, "Q", product_totals),
        (REGIONS, "Region", 18, "S", region_totals),
        (CHANNELS, "Channel", 20, "U", channel_totals),
    ]:
        for row, name in enumerate(names, 1):
            sheet.write_string(row, column, name)
            formula = f"=SUMIFS(tblDatinSales[Net_Sales_GHS],{filt},tblDatinSales[{field}],{letter}{row + 1})"
            sheet.write_formula(row, column + 1, formula, None, summary[name])

    for kind, column, end, title, anchor in [
        ("line", 14, 6, "Monthly net revenue", "B13"),
        ("doughnut", 16, len(PRODUCT_NAMES), "Revenue mix by product", "H13"),
        ("bar", 18, len(REGIONS), "Net revenue by region", "B26"),
        ("doughnut", 20, len(CHANNELS), "Revenue mix by channel", "H26"),
    ]:
        chart = book.add_chart({"type": kind})
        series = {
            "categories": [sheet.name, 1, column, end, column],
            "values": [sheet.name, 1, column + 1, end, column + 1],
        }
        if kind == "line":
            series.update({"line": {"color": GOLD, "width": 2.75}, "marker": {
                "type": "circle", "size": 5, "border": {"color": GOLD}, "fill": {"color": GOLD},
            }})
        elif kind == "doughnut":
            swatches = [GOLD, "#76BCB4", "#F69D8B", "#AB9FE2", "#C0C780", "#83AEEF"][:end]
            series["points"] = [{"fill": {"color": color}, "border": {"color": PANEL}} for color in swatches]
            series["data_labels"] = {"percentage": True, "font": {"color": NAVY, "size": 10, "bold": True}}
            chart.set_hole_size(62)
        else:
            series.update({"fill": {"color": GOLD}, "border": {"none": True}})
        chart.add_series(series)
        chart.show_hidden_data()
        chart.set_title({"name": title, "name_font": {"color": WHITE, "size": 12}})
        chart.set_legend({"position": "right", "font": {"color": WHITE, "size": 9}} if kind == "doughnut" else {"none": True})
        chart.set_chartarea({"fill": {"color": PANEL}, "border": {"none": True}})
        chart.set_plotarea({"fill": {"color": PANEL}, "border": {"none": True}})
        label_axis = {"num_font": {"color": WHITE, "size": 9}, "line": {"color": MUTED}}
        value_axis = {**label_axis, "num_format": '"GHS "0,"k"', "min": 0,
                      "major_gridlines": {"visible": True, "line": {"color": "#314463"}}}
        if kind == "line":
            chart.set_x_axis({**label_axis, "num_format": "mmm"})
            chart.set_y_axis(value_axis)
        elif kind == "bar":
            chart.set_x_axis(value_axis)
            chart.set_y_axis(label_axis)
        chart.set_size({"width": 481, "height": 250 if anchor.endswith("26") else 290})
        sheet.insert_chart(anchor, chart, {"x_scale": 0.96})
    sheet.merge_range("B37:M37", "Use Region, Channel and Product to filter the four KPIs and all four charts.", style(NAVY, MUTED))
    sheet.set_print_scale(77)
    sheet.set_landscape()
    sheet.print_area("A1:N38")
    sheet.fit_to_pages(1, 1)


def build_workbooks(lab_folders):
    for period, folder in [(FIRST, lab_folders[2]), (FRESH, lab_folders[3])]:
        rows = list(records(period))
        data = folder / "data"
        data.mkdir(exist_ok=True)
        with xlsxwriter.Workbook(data / period.filename) as book:
            write_sales(book, rows)
            write_checks(book, period, rows)
        if period == FIRST:
            examples = folder / "examples"
            examples.mkdir(exist_ok=True)
            with xlsxwriter.Workbook(examples / CHECKPOINT) as book:
                write_sales(book, rows)
                write_checks(book, period, rows)
                write_dashboard(book, rows)
