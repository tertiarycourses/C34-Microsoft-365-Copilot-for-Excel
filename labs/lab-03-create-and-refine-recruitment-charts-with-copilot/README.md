# Lab 3 — Create and Refine Recruitment Charts with Copilot

**Course:** Microsoft 365 Copilot for Excel  
**Course Code:** C34  
**Version:** v2.0 (8 October 2026)  
**Topic 2:** Use Copilot to Plot Charts  
**Maps to:** LO2: use Copilot to build PivotTables and trend, comparison, funnel and combo charts from prompts, refine them with follow-up prompts and tie each chart to a checked total  
**Duration:** 45 minutes  
**Tools:** Excel · Copilot (Allow editing) · PivotTables · line, bar, funnel and combo charts · BrightPath_Recruitment_Metrics.xlsx

> 📄 **Instruction PDF (with sample prompts):** [Lab-03-Instructions.pdf](Lab-03-Instructions.pdf)  
> 💬 **Copy-ready prompts:** [sample-prompts.txt](sample-prompts.txt)  
> 🌐 **Prompts web page (one-click copy):** [Lab 3 prompts](https://tertiarycourses.github.io/C34-Microsoft-365-Copilot-for-Excel/prompts/#lab-3)  
> 📊 **Mock data:** [data/BrightPath_Recruitment_Metrics.xlsx](data/BrightPath_Recruitment_Metrics.xlsx)  

---

## Goal

Produce four decision-ready recruitment charts - hiring trend, hires by sector, the H1 2026 funnel and source effectiveness - each backed by a reconciled summary.

## What You Will Do

BrightPath's leadership wants to see how hiring is moving and which sectors and sourcing channels perform best. You will ask Copilot to build PivotTables and charts from 18 months of funnel metrics and a sourcing summary, choose each chart from the question it answers, refine titles, labels, sort order and axes with follow-up prompts, and reconcile every chart to an independent SUMIFS check.

## What You Will Build

BrightPath_Recruitment_Metrics.xlsx with PivotTables and four refined charts on the Charts sheet: a monthly hires line, a sorted 2026 hires-by-sector bar, a 2026 recruitment funnel and a hire-rate vs cost-per-hire combo chart

## Prerequisites

- Topic 1 completed: you can open a OneDrive workbook, use Copilot modes and verify with formulas.
- This lab folder's data/BrightPath_Recruitment_Metrics.xlsx.

> **Data note.** Use only the supplied fictional BrightPath data or information you are authorised to process. Do not paste credentials, confidential business material or real employees' personal data into Copilot.

## Steps

### 1. Go to onedrive.com (or your SharePoint site), open or create the folder work-c34 and select Upload > Files. Upload BrightPath_Recruitment_Metrics.xlsx from this lab's data folder, then open it in Excel for the web or select Open in Desktop App. Confirm the title bar shows the cloud location and AutoSave is On, and that Formulas > Calculation Options is set to Automatic. Read the About_This_Data sheet before you prompt. Each chart you build must answer one of the four questions below.

```text
Tables : tblMonthly (108 rows = 18 months x 6 sectors) | tblSources (7 rows)

Q1 Trend      : How have total hires moved month by month since Jan 2025?
Q2 Comparison : Which sector produced the most hires in Jan-Jun 2026?
Q3 Funnel     : Where do candidates drop out between application and hire in 2026?
Q4 Efficiency : Which sources convert best, and at what cost per hire?
```

### 2. Choose Allow editing. Ask Copilot to write the four questions at the top of the Charts sheet, then ask for the trend PivotTable and line chart. Ask Copilot to check the Pivot grand total against =SUM(tblMonthly[Hires]) and name the peak month.

```text
ALLOW EDITING:
In Charts!A1:A4, type these four questions:
Q1 Trend: How have total hires moved month by month since Jan 2025?
Q2 Comparison: Which sector produced the most hires in Jan-Jun 2026?
Q3 Funnel: Where do candidates drop out between application and hire in 2026?
Q4 Efficiency: Which sources convert best, and at what cost per hire?

ALLOW EDITING:
Create a PivotTable from tblMonthly on a new sheet named Pivot_Trend with Month_Label in rows and
Sum of Hires as values. Then create a line chart of total hires by month from that PivotTable.
Show data labels and gridlines. Title: "Monthly hires, Jan 2025 - Jun 2026".

CHECK (Allow editing):
In an empty cell on Pivot_Trend, add this formula. Tell me whether it matches the PivotTable grand total,
and which month had the most hires:
=SUM(tblMonthly[Hires])
```

### 3. Ask for the 2026 sector comparison. Then refine it with a follow-up prompt. Ask Copilot to add the SUMIFS check for Technology's 2026 hires.

```text
Create a PivotTable on a new sheet Pivot_Sector from tblMonthly with Sector in rows and Sum of Hires,
filtered to months in 2026 only. Create a bar chart from it.

FOLLOW-UP: Sort the bars from highest to lowest and change the title to state the leading sector.

CHECK (Allow editing):
In an empty cell on Pivot_Sector, add this formula and tell me if it matches Technology's bar:
=SUMIFS(tblMonthly[Hires],tblMonthly[Sector],"Technology",tblMonthly[Month],">="&DATE(2026,1,1))
```

### 4. Ask Copilot to build a 2026 funnel summary with formulas, then a funnel chart. Ask Copilot to show the formula in each value cell to confirm the summary uses SUMIFS, not typed numbers.

```text
On the Charts sheet, build a small table of 2026 totals for Applications, Screened, Interviews, Offers
and Hires from tblMonthly using SUMIFS on Month >= 1 Jan 2026. Then create a funnel chart from it
with data labels. Add a column showing each stage as a % of Applications.

FOLLOW-UP: Show the formula in each value cell of the funnel table.
```

### 5. Ask Copilot to add rate columns to tblSources and a combo chart with two axes. Then ask it for the best hire rate and the lowest cost per hire.

```text
In tblSources add Hire_Rate = Hires / Applications (percentage, 1 decimal) and
Cost_per_Hire = Sourcing_Spend / Hires (S$, no decimals). Create a combo chart: Hire_Rate as
clustered columns sorted highest first, Cost_per_Hire as a line on a secondary axis.
Label both axes with units.

FOLLOW-UP: Which source has the highest Hire_Rate, and which has the lowest Cost_per_Hire? Give both values.
```

### 6. Refine all four charts with short follow-up prompts, one change at a time, then move them onto the Charts sheet in a 2 x 2 grid. Add alt text to each chart.

```text
Make every chart title state the finding in one sentence.
Use the same blue colour for all hires series and format numbers with thousands separators.
Move all four charts to the Charts sheet in a 2 x 2 grid of equal size.
Add alt text to each chart describing what it shows.
```

### 7. Ask Copilot (Chat only) to review the charts against the four questions, then ask it to log your prompts and checks on Prompt_Log.

```text
CHAT ONLY:
Review the four charts on the Charts sheet. For each, tell me which question at the top of the sheet
it answers, whether the chart type suits that question, and one improvement. Do not change anything.

LOG (Allow editing):
Prompt_Log already lists every prompt in this lab. For each one I ran, fill in What Copilot did and
How I verified, and set Decision to Keep. I will tell you any row to change to Revise or Undo.
```

## Sample Prompts

| # | Purpose | Prompt |
|---|---|---|
| 1 | PivotTable | Create a PivotTable from tblMonthly with Month_Label in rows, Sector in columns and Sum of Hires as values. |
| 2 | Line chart | Create a line chart showing total hires by month. Show data labels and gridlines. |
| 3 | Bar chart | Create a bar chart of 2026 hires by Sector, sorted from highest to lowest. |
| 4 | Funnel | Create a funnel chart of 2026 Applications, Screened, Interviews, Offers and Hires with data labels. |
| 5 | Combo chart | Create a combo chart of Hire_Rate as columns and Cost_per_Hire as a line on a secondary axis. |
| 6 | Refine | Change the chart title to: Technology led hiring in H1 2026. |
| 7 | Refine | Sort the bars from highest to lowest and add data labels formatted with thousands separators. |
| 8 | Accessibility | Add alt text to this chart that describes the trend it shows. |
| 9 | Review | Is a line chart the right choice for this data? Suggest a better chart if not, and explain why. |

## Test It

Pivot_Trend grand total = 1,065 hires (= SUM(tblMonthly[Hires])); the peak month is 2025-11 with 68 hires. The 2026 sector bar chart shows Technology first with 84 hires and its Pivot total is 361. The funnel shows 6,616 > 3,229 > 1,385 > 495 > 361. In tblSources, Referral has the highest hire rate (7.4%) and MyCareersFuture the lowest cost per hire (S$186). All four charts sit on the Charts sheet with finding-style titles, labelled axes and alt text.

## Checkpoint and Rejoin Point

Keep the workbook in work-c34. Lab 4 uses its own placements workbook.

## Troubleshooting

| If this happens | Fix |
|---|---|
| The 2026 filter includes 2025 months. | Ask Copilot to filter on the Month date field (>= 1 Jan 2026), not on Month_Label text, or add a Year field. |
| Copilot inserts a picture of a chart instead of a chart. | That came from a Chat answer and is static. Switch to Allow editing and ask it to create a chart linked to the data. |
| The funnel chart type is not offered. | Funnel charts need Excel 2019 or Microsoft 365. Ask for a sorted horizontal bar chart of the stages instead. |
| Both series in the combo chart share one axis and the line is flat. | Ask: 'Put Cost_per_Hire on a secondary axis and format it as S$'. |

## Challenge

Ask Copilot for a 100% stacked column chart showing each sector's share of hires by quarter for 2025-2026, then decide whether it communicates better or worse than the bar chart, and why.

## Reflection

Which of your four charts would most change a BrightPath leadership decision, and what check makes you trust it?

---

[← Lab 2](../lab-02-generate-and-verify-formula-columns-and-lookups-for-job-requisitions/README.md) · [Lab 4 →](../lab-04-visualise-pay-and-placement-distributions-with-python-in-excel/README.md)
