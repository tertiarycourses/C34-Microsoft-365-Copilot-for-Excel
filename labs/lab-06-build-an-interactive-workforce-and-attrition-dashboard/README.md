# Lab 6 — Build an Interactive Workforce and Attrition Dashboard

**Course:** Microsoft 365 Copilot for Excel  
**Course Code:** C34  
**Version:** v2.0 (8 October 2026)  
**Topic 3:** Use Copilot to Create Dashboards  
**Maps to:** LO3: use Copilot to build a refreshable HR workforce dashboard with helper columns, KPI cards, charts and slicers, prove it refreshes with new rows, and check it for responsible sharing  
**Duration:** 35 minutes  
**Tools:** Excel · Copilot (Plan, Allow editing) · PivotTables, slicers, refresh · BrightPath_Workforce_Headcount.xlsx

> 📄 **Instruction PDF (with sample prompts):** [Lab-06-Instructions.pdf](Lab-06-Instructions.pdf)  
> 💬 **Copy-ready prompts:** [sample-prompts.txt](sample-prompts.txt)  
> 🌐 **Prompts web page (one-click copy):** [Lab 6 prompts](https://tertiarycourses.github.io/C34-Microsoft-365-Copilot-for-Excel/prompts/#lab-6)  
> 📊 **Mock data:** [data/BrightPath_Workforce_Headcount.xlsx](data/BrightPath_Workforce_Headcount.xlsx)  

---

## Goal

Show headcount, 12-month attrition, new hires and tenure by department and location on one interactive page that updates correctly when new employees are added.

## What You Will Do

BrightPath's HR Director needs a workforce dashboard for the July leadership meeting. You will ask Copilot to add period helper columns, type independent checks for headcount and attrition, have Copilot plan and build the dashboard with slicers, then append eight July new joiners and refresh to prove the dashboard is live. Finally you will check that no slicer combination exposes a group too small to stay anonymous.

## What You Will Build

BrightPath_Workforce_Headcount.xlsx with helper columns in tblEmployees, a Checks area, and an HR_Dashboard sheet showing six KPI cards, headcount and attrition-rate charts by department, hires vs exits by quarter, and Location and Job_Level slicers - refresh-tested with new rows

## Prerequisites

- Lab 5 completed: you have built and verified a dashboard with Copilot.
- This lab folder's data/BrightPath_Workforce_Headcount.xlsx.

> **Data note.** Use only the supplied fictional BrightPath data or information you are authorised to process. Do not paste credentials, confidential business material or real employees' personal data into Copilot.

## Steps

### 1. Go to onedrive.com (or your SharePoint site), open or create the folder work-c34 and select Upload > Files. Upload BrightPath_Workforce_Headcount.xlsx from this lab's data folder, then open it in Excel for the web or select Open in Desktop App. Confirm the title bar shows the cloud location and AutoSave is On, and that Formulas > Calculation Options is set to Automatic. Read the About_This_Data sheet before you prompt. Do not touch the New_Joiners_Jul2026 sheet until Step 6.

```text
Table : tblEmployees (420 employees hired since 2019; Status Active or Exited)
Period: 2025-07-01 to 2026-06-30
Attrition rate = exits in period / ((active on 2025-07-01 + active on 2026-06-30) / 2)
```

### 2. Choose Allow editing and ask Copilot for three helper columns. Then ask it to explain the Active_Start formula and show one employee who exited before 2025-07-01 (Active_Start must be FALSE).

```text
Add three columns to tblEmployees:
1) Active_Start = TRUE if Hire_Date <= 2025-07-01 and (Exit_Date is blank or Exit_Date > 2025-07-01).
2) Exit_In_Period = TRUE if Exit_Date is between 2025-07-01 and 2026-06-30 inclusive.
3) Tenure_Years = years from Hire_Date to Exit_Date, or to 2026-06-30 if Exit_Date is blank, 1 decimal.

FOLLOW-UP: Explain the Active_Start formula, and show one employee who exited before 2025-07-01
with Hire_Date, Exit_Date and Active_Start.
```

### 3. Ask Copilot to create a Checks sheet with these independent checks (label in column A, formula in column B), exactly as written. These are the numbers the dashboard must match.

```text
Create a sheet named Checks. Put these labels in column A and formulas in column B exactly as written,
then tell me each result. Build the Attrition rate formula from the cells above it:
Active headcount     =COUNTIFS(tblEmployees[Status],"Active")
Headcount 2025-07-01 =COUNTIFS(tblEmployees[Active_Start],TRUE)
Exits in period      =COUNTIFS(tblEmployees[Exit_In_Period],TRUE)
Voluntary exits      =COUNTIFS(tblEmployees[Exit_In_Period],TRUE,tblEmployees[Exit_Type],"Voluntary")
Attrition rate       = Exits in period / ((Headcount 2025-07-01 + Active headcount) / 2)
Hires in period      =COUNTIFS(tblEmployees[Hire_Date],">="&DATE(2025,7,1),tblEmployees[Hire_Date],"<="&DATE(2026,6,30))
```

### 4. Choose Plan, paste the brief, review the plan against the definitions, then switch to Allow editing to build.

```text
Plan an HR workforce dashboard on the HR_Dashboard sheet from tblEmployees for the HR Director.
KPI cards (formulas): Active headcount, Exits (last 12 months), Attrition rate, Voluntary attrition rate,
Avg tenure of active staff (years), Hires (last 12 months). Use the definitions on About_This_Data.
Charts: active headcount by Department (bar, sorted); attrition rate by Department (bar); hires vs exits
by quarter for the last 12 months (clustered column); active headcount by Location (bar).
Slicers: Location and Job_Level, connected to all PivotTables. Title and "Data as at 2026-06-30".
Do not show Employee_ID, salary or individual records on the dashboard.
```

### 5. Ask Copilot to compare the KPI cards with Checks, and fix any difference with a follow-up prompt. Then ask which Location, Job_Level and Department combinations contain fewer than 5 people.

```text
CHAT ONLY:
Compare each KPI card on HR_Dashboard with the Checks sheet and tell me which, if any, do not match.

FOLLOW-UP (if needed):
The attrition rate must use average headcount = (Active_Start count + active count) / 2.

CHAT ONLY:
Which combinations of Location, Job_Level and Department contain fewer than 5 active employees?
```

### 6. Refresh test: ask Copilot to append the 8 rows from New_Joiners_Jul2026 to tblEmployees and refresh every PivotTable. Every KPI and chart that counts active staff must update.

```text
ALLOW EDITING:
Copy the 8 data rows in New_Joiners_Jul2026!A2:L9 to the first empty rows directly below tblEmployees
so the table expands. Then refresh all PivotTables in the workbook and tell me the new Active headcount
and the number of rows in tblEmployees.

Expected after refresh:
Active headcount : 237 -> 245
tblEmployees rows: 420 -> 428
New joiners by department: Recruitment +2, Sales & Client Services +2, IT +4
```

### 7. Ask Copilot (Chat only) to review the dashboard for privacy and clarity, ask it to act on its valid points, and ask it to log all prompts on Prompt_Log.

```text
CHAT ONLY:
Review HR_Dashboard before I share it with the leadership team. Does any element expose individual
employee data or a group smaller than 5? Is each KPI labelled with its period and definition?

LOG (Allow editing):
Add one row to Prompt_Log for each prompt I ran in this lab: Mode, Prompt, What Copilot did,
How it was verified, and Decision = Keep. I will tell you any row to change to Revise or Undo.
```

## Sample Prompts

| # | Purpose | Prompt |
|---|---|---|
| 1 | Helper column | Add Active_Start = TRUE if Hire_Date <= 2025-07-01 and Exit_Date is blank or after 2025-07-01. |
| 2 | Helper column | Add Tenure_Years from Hire_Date to Exit_Date, or to 2026-06-30 if Exit_Date is blank, rounded to 1 decimal. |
| 3 | Plan | Plan an HR workforce dashboard with KPI cards for active headcount, exits, attrition rate, voluntary attrition, average tenure and hires. |
| 4 | Chart | Create a bar chart of 12-month attrition rate by Department, sorted highest first, with data labels as percentages. |
| 5 | Slicers | Add slicers for Location and Job_Level and connect them to every PivotTable on the dashboard. |
| 6 | Refresh | I added rows to tblEmployees. Refresh all PivotTables and tell me which KPI values changed. |
| 7 | Privacy | Which slicer combinations show a group of fewer than 5 employees? Suggest how to suppress them. |
| 8 | Explain | Explain how the Attrition rate card is calculated, step by step. |

## Test It

Before the refresh: Active headcount 237, headcount on 2025-07-01 210, exits in period 29 (25 voluntary), attrition rate 13.0% and voluntary attrition 11.2%, hires in period 56, average tenure of active staff about 3.1 years. Recruitment has the highest attrition rate (21.8%, 12 exits) and the largest active headcount (60). After appending the July joiners and refreshing, Active headcount shows 245 and tblEmployees has 428 rows. The dashboard shows no Employee_ID, salary or individual record.

## Checkpoint and Rejoin Point

Keep the workbook in work-c34. Topic 4 starts from the attrition-analysis workbook in the Lab 7 folder.

## Troubleshooting

| If this happens | Fix |
|---|---|
| The new rows did not join the table and nothing changed on refresh. | Ask Copilot to resize tblEmployees to include the 8 new rows, then refresh all PivotTables again. |
| Attrition rate is far too high (e.g. 29 / 237). | The card divides by closing headcount only. Ask Copilot to use the average of opening and closing headcount. |
| Helper columns show #VALUE! for blank Exit_Date. | Ask Copilot to treat blank Exit_Date explicitly, e.g. OR([@Exit_Date]="",[@Exit_Date]>DATE(2025,7,1)). |
| A chart did not update after the refresh. | It is a static picture or a range copy. Ask Copilot to replace it with a PivotChart linked to tblEmployees. |

## Challenge

Ask Copilot to add a Timeline control on Hire_Date and a chart of average Engagement_Score by Department for active staff, then check one department's average with AVERAGEIFS.

## Reflection

What would happen to trust in this dashboard if one leader could filter it down to a single named team member?

---

[← Lab 5](../lab-05-build-a-recruitment-kpi-dashboard-with-copilot/README.md) · [Lab 7 →](../lab-07-surface-attrition-insights-with-copilot-and-advanced-analysis/README.md)
