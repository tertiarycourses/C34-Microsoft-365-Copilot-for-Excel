# Lab 2 — Generate and Verify Formula Columns and Lookups for Job Requisitions

**Course:** Microsoft 365 Copilot for Excel  
**Course Code:** C34  
**Version:** v2.0 (8 October 2026)  
**Topic 1:** Use Copilot to Analyse Data  
**Maps to:** LO1: use Copilot to generate lookup and calculated columns from plain-English rules, explain them, and reconcile them to independent control formulas  
**Duration:** 35 minutes  
**Tools:** Excel · Copilot (Allow editing, Chat only) · XLOOKUP · structured references · BrightPath_Job_Requisitions.xlsx

> 📄 **Instruction PDF (with sample prompts):** [Lab-02-Instructions.pdf](Lab-02-Instructions.pdf)  
> 💬 **Copy-ready prompts:** [sample-prompts.txt](sample-prompts.txt)  
> 🌐 **Prompts web page (one-click copy):** [Lab 2 prompts](https://tertiarycourses.github.io/C34-Microsoft-365-Copilot-for-Excel/prompts/#lab-2)  
> 📊 **Mock data:** [data/BrightPath_Job_Requisitions.xlsx](data/BrightPath_Job_Requisitions.xlsx)  

---

## Goal

Add recruiter names, days open, fill rate, estimated placement fee and an ageing flag to a requisitions table, then prove each column is right.

## What You Will Do

BrightPath's permanent-recruitment desk tracks 60 client job orders. You will ask Copilot to add an XLOOKUP for the recruiter's name and four calculated columns written from business rules, ask it to explain each formula, hand-check one requisition and reconcile column totals to independent formulas on the Control sheet. You will then ask Copilot which recruiter generated the most estimated fees and verify the answer.

## What You Will Build

BrightPath_Job_Requisitions.xlsx with Recruiter_Name, Days_Open, Fill_Rate, Est_Fee and Ageing_Flag columns in tblRequisitions, overdue requisitions highlighted, and a Control sheet that reconciles the totals

## Prerequisites

- Lab 1 completed, or Copilot access confirmed on a OneDrive file with AutoSave On.
- This lab folder's data/BrightPath_Job_Requisitions.xlsx.

> **Data note.** Use only the supplied fictional BrightPath data or information you are authorised to process. Do not paste credentials, confidential business material or real employees' personal data into Copilot.

## Steps

### 1. Go to onedrive.com (or your SharePoint site), open or create the folder work-c34 and select Upload > Files. Upload BrightPath_Job_Requisitions.xlsx from this lab's data folder, then open it in Excel for the web or select Open in Desktop App. Confirm the title bar shows the cloud location and AutoSave is On, and that Formulas > Calculation Options is set to Automatic. Read the About_This_Data sheet before you prompt. Note the fee rule: Est_Fee = Filled x Monthly_Salary_Budget x 12 x Fee_Rate, and the report date in Control!B1.

```text
Tables : tblRequisitions (60 rows) | tblRecruiters (6 rows)
Control!B1 Report_Date = 2026-06-30
```

### 2. Choose Allow editing. Ask Copilot to add the lookup column. Select a cell in the new column and ask Copilot to explain the formula. Check one row by hand against tblRecruiters.

```text
In tblRequisitions, add a column Recruiter_Name that looks up Recruiter_Name from tblRecruiters
using Recruiter_ID. Use XLOOKUP and return "Not found" if there is no match.

Then: Explain the formula in the selected cell.
```

### 3. Ask for Days_Open with an explicit rule for open requisitions. Read the explanation. Test an open row (blank Close_Date) and a closed row by subtracting the dates yourself.

```text
Add a column Days_Open to tblRequisitions. If Close_Date is blank, use Control!B1 minus Open_Date;
otherwise use Close_Date minus Open_Date. Show whole numbers.
```

### 4. Ask for Fill_Rate and Est_Fee, one prompt each. Format Fill_Rate as a percentage and Est_Fee as S$ with no decimals. Hand-check JOB-2002: Filled 4 x S$4,900 x 12 x 25% = S$58,800.

```text
Add a column Fill_Rate = Filled divided by Openings, formatted as a percentage.

Add a column Est_Fee = Filled x Monthly_Salary_Budget x 12 x Fee_Rate, formatted as S$ with no decimals.
Explain the formula.
```

### 5. Ask for the ageing flag and a highlight. Count the flags with the status bar or a COUNTIF.

```text
Add a column Ageing_Flag that shows "Overdue" when Status is "Open" and Days_Open is greater than 45,
otherwise leave it blank. Highlight rows where Ageing_Flag is "Overdue" with a light amber fill.
```

### 6. Build independent controls on the Control sheet using the formulas below (type them yourself - do not reuse Copilot's columns for the control). Each pair must match.

```text
A3 Total Est_Fee (column)       B3 =SUM(tblRequisitions[Est_Fee])
A4 Total fee (independent)      B4 =SUMPRODUCT(tblRequisitions[Filled],tblRequisitions[Monthly_Salary_Budget],tblRequisitions[Fee_Rate])*12
A5 Avg days to fill (Filled)    B5 =AVERAGEIFS(tblRequisitions[Days_Open],tblRequisitions[Status],"Filled")
A6 Overdue (independent)        B6 =COUNTIFS(tblRequisitions[Status],"Open",tblRequisitions[Days_Open],">45")
A7 Overdue flags (column)       B7 =COUNTIF(tblRequisitions[Ageing_Flag],"Overdue")
```

### 7. In Chat only, ask which recruiter generated the most estimated fees. Verify with SUMIFS on Control, then log all prompts on Prompt_Log.

```text
Using tblRequisitions, which Recruiter_Name has the highest total Est_Fee? Show the total for every recruiter.

CHECK: =SUMIFS(tblRequisitions[Est_Fee],tblRequisitions[Recruiter_Name],"<name>")
```

## Sample Prompts

| # | Purpose | Prompt |
|---|---|---|
| 1 | Lookup | Add a column Recruiter_Name that looks up Recruiter_Name from tblRecruiters using Recruiter_ID with XLOOKUP. Return "Not found" if missing. |
| 2 | Duration | Add a column Days_Open: if Close_Date is blank use Control!B1 minus Open_Date, otherwise Close_Date minus Open_Date. |
| 3 | Ratio | Add a column Fill_Rate = Filled / Openings formatted as a percentage. |
| 4 | Business rule | Add a column Est_Fee = Filled x Monthly_Salary_Budget x 12 x Fee_Rate, formatted as S$ with no decimals. |
| 5 | Flag | Add a column Ageing_Flag that shows "Overdue" when Status is Open and Days_Open > 45, otherwise blank. |
| 6 | Explain | Explain the formula in the selected cell step by step, and tell me how blanks are handled. |
| 7 | Summary row | Add a total row to tblRequisitions that sums Openings, Filled and Est_Fee and averages Days_Open. |
| 8 | Question | Which recruiter has the highest total Est_Fee? Show every recruiter's total, highest first. |

## Test It

Control!B3 and B4 both show S$1,030,512. B5 shows 34.6 days (29 Filled requisitions). B6 and B7 both show 13. JOB-2002 shows Est_Fee S$58,800 and Days_Open 36. Every Recruiter_Name resolves (no "Not found"). The top recruiter by estimated fees is Aisha Rahman with S$221,760.

## Checkpoint and Rejoin Point

Keep the workbook in work-c34. Topic 2 starts from a new metrics workbook in the Lab 3 folder.

## Troubleshooting

| If this happens | Fix |
|---|---|
| Days_Open is negative or huge for open requisitions. | The formula used today's date or a blank as zero. Ask Copilot to use Control!B1 explicitly, or edit the formula. |
| Est_Fee is 12 times too small or too large. | Re-read the rule: fee is on first-year salary (monthly x 12). Ask Copilot to explain the formula and correct it. |
| The new column is outside the table. | Click inside tblRequisitions before prompting and name the table in the prompt; or drag the table's resize handle. |
| Copilot cannot add a formula column. | Confirm Calculation Options is Automatic and the data is an Excel table; try Excel for the web if desktop fails. |

## Challenge

Ask Copilot to add a Fee_Achieved column that is Est_Fee for Filled requisitions only, then build a small summary of Fee_Achieved by Sector with SUMIFS. Check that the sector totals add up to the grand total.

## Reflection

Which of today's formulas would have looked correct but been wrong without a hand-check, and how did you catch it?

---

[← Lab 1](../lab-01-explore-and-question-recruitment-pipeline-data/README.md) · [Lab 3 →](../lab-03-create-and-refine-recruitment-charts-with-copilot/README.md)
