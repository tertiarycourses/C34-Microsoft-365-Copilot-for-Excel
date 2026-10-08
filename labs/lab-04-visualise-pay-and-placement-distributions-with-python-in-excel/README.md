# Lab 4 — Visualise Pay and Placement Distributions with Python in Excel

**Course:** Microsoft 365 Copilot for Excel  
**Course Code:** C34  
**Version:** v2.0 (8 October 2026)  
**Topic 2:** Use Copilot to Plot Charts  
**Maps to:** LO2: use Copilot with Python in Excel to create distribution, relationship and heatmap visuals, review the generated code and verify plotted statistics with native formulas  
**Duration:** 35 minutes  
**Tools:** Excel · Copilot (Allow editing, Chat only) · Python in Excel (=PY) · BrightPath_Contract_Placements.xlsx

> 📄 **Instruction PDF (with sample prompts):** [Lab-04-Instructions.pdf](Lab-04-Instructions.pdf)  
> 💬 **Copy-ready prompts:** [sample-prompts.txt](sample-prompts.txt)  
> 📊 **Mock data:** [data/BrightPath_Contract_Placements.xlsx](data/BrightPath_Contract_Placements.xlsx)  

---

## Goal

Show how contract bill rates, pay and margins are distributed across job levels and sectors, and explain one outlier, using Python visuals you have checked.

## What You Will Do

BrightPath's flexible-staffing desk places contract associates with clients. You will ask Copilot to add a gross margin column, then create Python in Excel visuals - a box plot of bill rates by job level, a scatter of experience against pay with a trend line, and a sector-by-level margin heatmap. You will read the =PY code, re-compute a median and a correlation with Excel formulas, and investigate the highest pay rate.

## What You Will Build

BrightPath_Contract_Placements.xlsx with a Gross_Margin_Pct column and, on Python_Charts, a box plot, a scatter with trend line, a margin heatmap and an outlier table - each with a native-formula check beside it

## Prerequisites

- Lab 3 completed or equivalent charting experience.
- Python in Excel available: Formulas > Insert Python is visible (Windows, Mac or web; internet required).
- This lab folder's data/BrightPath_Contract_Placements.xlsx.

> **Data note.** Use only the supplied fictional BrightPath data or information you are authorised to process. Do not paste credentials, confidential business material or real employees' personal data into Copilot.

## Steps

### 1. Go to onedrive.com (or your SharePoint site), open or create the folder work-c34 and select Upload > Files. Upload BrightPath_Contract_Placements.xlsx from this lab's data folder, then open it in Excel for the web or select Open in Desktop App. Confirm the title bar shows the cloud location and AutoSave is On, and that Formulas > Calculation Options is set to Automatic. Read the About_This_Data sheet before you prompt. Check that the Formulas tab shows Insert Python. If it does not, follow the fallback in Troubleshooting.

```text
Table  : tblPlacements (180 rows)
Levels : Junior | Mid | Senior | Lead
Output sheet : Python_Charts
```

### 2. Choose Allow editing and ask for the margin column. Check one row by hand and the average with AVERAGE.

```text
Add a column Gross_Margin_Pct to tblPlacements = (Bill_Rate_Hourly - Pay_Rate_Hourly) / Bill_Rate_Hourly,
formatted as a percentage with 1 decimal.

CHECK: =AVERAGE(tblPlacements[Gross_Margin_Pct])
```

### 3. Ask Copilot to create the box plot with Python. Click the =PY cell and read the code: which table and columns it uses and the order of levels. Then verify the Lead median with a native formula in the cell beside it.

```text
Using Python in Excel, create a box plot of Bill_Rate_Hourly by Job_Level from tblPlacements,
ordered Junior, Mid, Senior, Lead. Put it in Python_Charts!B2. Title it
"Contract bill rate by job level (S$/hour)" and label both axes.

CHECK: =MEDIAN(FILTER(tblPlacements[Bill_Rate_Hourly],tblPlacements[Job_Level]="Lead"))
```

### 4. Ask for a scatter plot with a trend line and the correlation coefficient. Verify r with CORREL.

```text
Using Python, create a scatter plot of Years_Experience (x) against Pay_Rate_Hourly (y) from
tblPlacements, coloured by Job_Level, with a linear trend line. Also return the Pearson correlation
coefficient in the cell below the chart.

CHECK: =CORREL(tblPlacements[Years_Experience],tblPlacements[Pay_Rate_Hourly])
```

### 5. Ask for a heatmap of average margin. Read the code to confirm it uses the mean of Gross_Margin_Pct per cell.

```text
Using Python, create a heatmap of the average Gross_Margin_Pct by Sector (rows) and Job_Level (columns,
ordered Junior to Lead) with the values printed in each cell as percentages.
```

### 6. Ask Copilot (Chat only) to find outliers with a stated rule, then look at the highest pay rate row. Decide in Prompt_Log whether it is an error, a genuine specialist rate, or needs checking with the desk - an outlier is not automatically wrong.

```text
Using the IQR rule (above Q3 + 1.5 x IQR), list the Pay_Rate_Hourly outliers in tblPlacements with
Placement_ID, Sector, Job_Level and Pay_Rate_Hourly. State Q1, Q3 and the upper fence. Show the code.
```

### 7. Ask Copilot to explain what the Python code in the box-plot cell does, then record all prompts and checks in Prompt_Log.

```text
Explain the Python code in Python_Charts!B2 line by line in plain English.
```

## Sample Prompts

| # | Purpose | Prompt |
|---|---|---|
| 1 | Calculated column | Add Gross_Margin_Pct = (Bill_Rate_Hourly - Pay_Rate_Hourly) / Bill_Rate_Hourly as a percentage. |
| 2 | Box plot | Using Python, create a box plot of Bill_Rate_Hourly by Job_Level ordered Junior, Mid, Senior, Lead. |
| 3 | Histogram | Using Python, create a histogram of Pay_Rate_Hourly with 15 bins and mark the median. |
| 4 | Scatter | Using Python, plot Years_Experience against Pay_Rate_Hourly with a linear trend line and return the correlation. |
| 5 | Heatmap | Using Python, create a heatmap of average Gross_Margin_Pct by Sector and Job_Level with values shown. |
| 6 | Outliers | Using the IQR rule, list Pay_Rate_Hourly outliers with Placement_ID, Sector and Job_Level, and state the upper fence. |
| 7 | Explain code | Explain the Python code in the selected cell line by line. |
| 8 | Direct answer | What are the median and mean Bill_Rate_Hourly by Job_Level? Show the code. |

## Test It

Average Gross_Margin_Pct = 22.6%. The box plot shows four levels in order and the Lead median bill rate check returns S$65.74 (Junior S$22.72, Mid S$29.53, Senior S$41.70). CORREL returns about 0.49 - a moderate positive relationship. The heatmap shows Technology with the lowest average margin (~19.1%). The IQR rule flags 8 pay-rate outliers above about S$68.44; the highest is PL-5060 at S$173.47, and your Prompt_Log records a decision about it.

## Checkpoint and Rejoin Point

Keep the workbook in work-c34. Topic 3 starts from the dashboard workbook in the Lab 5 folder.

## Troubleshooting

| If this happens | Fix |
|---|---|
| Insert Python is not visible or =PY returns an error. | Python in Excel needs a supported plan, Windows/Mac/web and internet. Fallback: ask Copilot in Chat only for the same visual (it is returned as a static image) and use Insert > Charts > Statistical > Box and Whisker for a native box plot. |
| The =PY cell shows a Python object instead of a picture. | Select the cell and change its output (the icon beside the formula bar) from Python Object to Excel Value, or click the card icon to view the image. |
| #BUSY! or a slow calculation. | Python runs in the Microsoft cloud. Wait, or select Formulas > Reset (Python) and recalculate. |
| Levels appear in alphabetical order (Junior, Lead, Mid, Senior). | Ask Copilot to set the category order explicitly to Junior, Mid, Senior, Lead. |

## Challenge

Ask Copilot, using Python, whether contracts that were Extended have a higher average Client_Rating than those that were not, with counts. Re-check the two averages with AVERAGEIFS.

## Reflection

Which Python visual told you something a column chart of averages would have hidden?

---

[← Lab 3](../lab-03-create-and-refine-recruitment-charts-with-copilot/README.md) · [Lab 5 →](../lab-05-build-a-recruitment-kpi-dashboard-with-copilot/README.md)
