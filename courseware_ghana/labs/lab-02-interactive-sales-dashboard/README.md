# Lab 2 — Create and Restyle an Interactive Sales Dashboard

[Instructions PDF](Lab-02-Instructions.pdf) · [Copy prompts](../../prompts/index.html#lab-2)

## Goal

Build an interactive GHS dashboard, then apply the dark theme while preserving every result and control.

## Prerequisites

- Excel for Microsoft 365 and a work account with Copilot editing enabled.
- Open data/Datin_Sales_Finance_H1_GHS.xlsx as a working copy in OneDrive. Turn AutoSave on.

## Steps

### 1. Open the practice workbook in Excel. On the Sales Data sheet, click inside tblDatinSales. Open Copilot, select Plan if available to review the proposed changes, then allow editing to apply them.

### 2. Watch the trainer run Prompt 1, then run it in your own H1 copy. Check that Copilot creates a new dashboard sheet and links every chart and KPI to the data. Compare the structure with the reference video.

```text
Create a professional, interactive sales dashboard using the data in this workbook. Include key sales metrics, charts showing sales trends and performance, and interactive filters so I can explore the data. Put everything on a new worksheet.
```

### 3. Inspect the dashboard. It should show Net Sales, Gross Profit, Gross Margin and Order Count, plus a monthly trend, regional bars and doughnuts for product and channel mix. If any elements are missing, run the follow-up. Test All, Greater Accra plus Wholesale, and then add Rice. Record each result.

```text
On the dashboard sheet, use only tblDatinSales. Show Net Sales, Gross Profit, Gross Margin (total Gross Profit divided by total Net Sales) and Order Count (count Order_ID). Format money in GHS. Show a monthly net revenue trend, regional revenue bars, and doughnut charts for Product and Channel mix. Connect Region, Channel and Product controls to every KPI and chart. Keep the Sales Data table unchanged and tell me which controls you implemented.
```

### 4. Learner activity. Save a pre-theme copy. Select the dashboard sheet, run Prompt 2 and inspect the result at 100% zoom.

```text
Redesign the dashboard with a modern dark theme using navy blue backgrounds, gold accents, and high-contrast white text. Keep the existing layout, charts, metrics, and interactive filters fully functional. Make sure all labels and charts remain easy to read.
```

### 5. Repeat all three filter states. Check that all four KPI values match the pre-theme checks, every chart responds, and the labels and controls remain readable. Clear the filters and save a working copy.

## Test it

- **All regions and channels** · GHS 1,448,550 net sales, GHS 563,084 gross profit, 38.87% margin, 432 orders
- **Greater Accra plus Wholesale** · GHS 302,775 net sales, GHS 115,670 gross profit, 38.20% margin, 72 orders
- **Greater Accra plus Wholesale plus Rice** · GHS 32,080 net sales, GHS 11,629 gross profit, 36.25% margin, 12 orders
- **Dark theme** · Navy panels, gold accents and readable white labels. The original layout, charts and controls still work.

## Checkpoint

Use examples/Datin_Sales_Dashboard_Example_GHS.xlsx to check the finished look and test the Region, Channel and Product drop-downs. For the theme activity, use your Prompt 1 dashboard or a trainer-provided pre-theme copy. Keep the theme and verification steps. In Lab 3 you will save them as a skill and test them with Tema’s fresh H2 orders.

## Files

- [Datin_Sales_Finance_H1_GHS.xlsx](data/Datin_Sales_Finance_H1_GHS.xlsx)
- [Finished dark dashboard example](examples/Datin_Sales_Dashboard_Example_GHS.xlsx)
- [Kevin Stratvert’s dashboard walkthrough](https://www.youtube.com/watch?v=YgcEesCDIkw)
- [Microsoft’s slicer guide](https://support.microsoft.com/en-us/excel/get-started/use-slicers-to-filter-data)
