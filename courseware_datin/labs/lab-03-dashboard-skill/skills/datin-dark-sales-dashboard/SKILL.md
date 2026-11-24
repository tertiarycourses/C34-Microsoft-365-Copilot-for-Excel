---
name: datin-dark-sales-dashboard
description: Create an interactive Datin sales dashboard in Excel or restyle an existing sales dashboard with the Datin navy, gold and white theme. Use for requests to create an interactive Datin sales dashboard or apply its dark theme.
---

# Datin dark sales dashboard

## Choose the operation

- If the user asks to restyle an existing dashboard, identify the intended dashboard sheet. Keep its layout, chart objects and series, KPI formulas, source table, filter connections and current filter selections.
- If the user asks to create the dashboard and a dashboard is absent, create one new worksheet named Sales Dashboard. Prefer tblDatinSales. Inspect the current workbook and confirm the source and mapping if several plausible tables exist.
- If a dashboard already exists and the request to create another is ambiguous, ask before replacing the existing sheet.

## Build when needed

1. Use one row per order, keyed by Order_ID. Use Net_Sales_GHS and Gross_Profit_GHS for totals. Count Order_ID. Compute Gross Margin as total Gross Profit divided by total Net Sales. Show a dash for zero sales. Format money with the GHS code and margin as a percentage.
2. Add a monthly net revenue line chart using Month_Start, a regional revenue bar chart, a doughnut chart by Product and another doughnut chart by Channel. Use net sales after discounts as net revenue on all four charts.
3. Add working Region, Channel and Product controls, such as connected slicers or drop-down selectors driving formulas. Use the same filters for all charts and all four KPIs. Start with All. Keep source values and table structure intact. Keep chart support calculations on the dashboard.
4. If the Excel session cannot create working controls, identify the unfinished control and explain the next manual step. Report completion after filter behavior is verified.

## Theme from the Datin redesign prompt

- Use navy #0B1730 for the dashboard canvas and #152642 for cards and chart panels.
- Use gold #E9BE64 for key accents and primary series. Use white #FFFFFF for titles, axes, labels and important text. Make gridlines subtle and keep filter states distinguishable with text and contrast.
- For a restyle, edit only the dashboard’s presentation. Preserve the existing chart series, units, formulas, positions and controls. Keep currency, labels and legends legible at normal zoom.

## Verify

- Compare unfiltered sales, gross profit, margin and order count with the source totals.
- Filter Greater Accra and Wholesale, then add Rice when those options exist. Check that all KPIs and charts respond. Restore the starting selections after testing.
- For a restyle, confirm the same selection displays the same numbers before and after. Report what changed and identify any control that still needs manual verification.
