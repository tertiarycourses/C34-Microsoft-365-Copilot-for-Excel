# Lab 1 — Explore and Question Recruitment Pipeline Data

**Course:** Microsoft 365 Copilot for Excel  
**Course Code:** C34  
**Version:** v2.0 (8 October 2026)  
**Topic 1:** Use Copilot to Analyse Data  
**Maps to:** LO1: confirm Copilot access, prepare a named table, and use Chat and Edit modes to question, check, highlight, sort and filter recruitment pipeline data  
**Duration:** 35 minutes  
**Tools:** Excel for Microsoft 365 or Excel for the web · Copilot (Chat only, Allow editing) · BrightPath_Candidate_Pipeline.xlsx

> 📄 **Instruction PDF (with sample prompts):** [Lab-01-Instructions.pdf](Lab-01-Instructions.pdf)  
> 💬 **Copy-ready prompts:** [sample-prompts.txt](sample-prompts.txt)  
> 📊 **Mock data:** [data/BrightPath_Candidate_Pipeline.xlsx](data/BrightPath_Candidate_Pipeline.xlsx)  

---

## Goal

Turn a raw candidate-pipeline export into a named Excel table, find its data-quality issues with Copilot, and answer pipeline questions with checked numbers.

## What You Will Do

You are a talent analyst at BrightPath Talent Solutions, a fictional Singapore recruitment and staffing firm. You will put the H1 2026 candidate-pipeline export into a named table, ask Copilot (Chat only) to describe it, use Copilot to find duplicate IDs and inconsistent Source labels, then answer pipeline questions and build a highlighted, sorted and filtered review view. Every number is checked with a formula.

## What You Will Build

BrightPath_Candidate_Pipeline.xlsx in OneDrive with table tblApplications, a Source_Clean column, duplicate IDs highlighted, Hired rows highlighted, a verified source-by-hires answer and a filled Prompt_Log

## Prerequisites

- A Microsoft 365 account with Copilot in Excel enabled, signed in to Excel or Excel for the web.
- A OneDrive or SharePoint folder named work-c34 that you can upload to.
- This lab folder downloaded, including data/BrightPath_Candidate_Pipeline.xlsx.

> **Data note.** Use only the supplied fictional BrightPath data or information you are authorised to process. Do not paste credentials, confidential business material or real employees' personal data into Copilot.

## Steps

### 1. Go to onedrive.com (or your SharePoint site), open or create the folder work-c34 and select Upload > Files. Upload BrightPath_Candidate_Pipeline.xlsx from this lab's data folder, then open it in Excel for the web or select Open in Desktop App. Confirm the title bar shows the cloud location and AutoSave is On, and that Formulas > Calculation Options is set to Automatic. Read the About_This_Data sheet before you prompt.

```text
Required state:
File      : OneDrive/work-c34/BrightPath_Candidate_Pipeline.xlsx
AutoSave  : On
Calculation Options : Automatic
Sheets    : About_This_Data | Applications | Prompt_Log
```

### 2. Select cell A1 on the Applications sheet and press Ctrl+T (Home > Format as Table). Tick My table has headers and select OK. On the Table Design tab, set Table Name to tblApplications. Then select the Copilot button. If Copilot is unavailable, use the Troubleshooting table before continuing.

```text
Table name : tblApplications
Grain      : one row = one candidate application (Application_ID)
Expected   : 13 columns, 242 data rows
```

### 3. In the Copilot pane, choose Chat only (no workbook changes). Paste the prompt below. Compare Copilot's row count and date range with the table (select the Application_ID column and read Count on the status bar). Record the prompt, what Copilot said and your check on the Prompt_Log sheet.

```text
GOAL: Describe this candidate pipeline before I analyse it.
CONTEXT: BrightPath recruitment desk, H1 2026. One row should be one application.
DATA: Use only tblApplications.
CONSTRAINTS: Do not change the workbook.
OUTPUT: What a row represents, each column with its data type, the number of rows,
        the Applied_Date range and the distinct values of Stage and Source.
REVIEW: Tell me how you counted the rows.
```

### 4. Still in Chat only, ask Copilot to check data quality. It should report 2 repeated Application_IDs and Source values that are not in the canonical list on About_This_Data. Then switch to Allow editing (Edit) and ask Copilot to highlight the duplicate IDs and add a Source_Clean column. Check that no row was deleted - the table must still have 242 rows.

```text
CHAT ONLY:
Check tblApplications for data-quality issues. List any Application_ID that appears more than once,
and any Source value that is not exactly one of: LinkedIn, JobStreet, MyCareersFuture, Referral,
BrightPath Website, Job Fair, Talent Database. Give the count of each issue. Do not change anything.

ALLOW EDITING:
1) Highlight Application_ID cells that appear more than once in tblApplications with an orange fill.
2) Add a column Source_Clean to tblApplications that trims spaces from Source and changes 'Linked In'
   to 'LinkedIn'. Do not delete or edit any existing values.
```

### 5. Ask Copilot a direct question about which sources produce hires. Expand or ask Copilot to summarise any code it shows. Then verify the hires total in an empty cell on Prompt_Log with the formula below, and verify one source with COUNTIFS. Duplicate rows are Rejected applications, so they do not change the hire counts.

```text
PROMPT (Chat only):
Using tblApplications, count applications and hires (Stage = "Hired") by Source_Clean.
Show a table sorted by hires, highest first, with hire rate = hires / applications.

CHECK FORMULAS:
=COUNTIFS(tblApplications[Stage],"Hired")
=COUNTIFS(tblApplications[Source_Clean],"LinkedIn",tblApplications[Stage],"Hired")
```

### 6. Switch to Allow editing and build a review view, one change at a time. After each change, inspect the sheet and keep it or Undo. Finally apply the filter, read the visible-row count on the status bar, record it, and clear the filter so all 242 rows return.

```text
1) Highlight entire rows of tblApplications where Stage is "Hired" with a light green fill.
2) Sort tblApplications by Interview_Score from highest to lowest, then by Days_in_Pipeline
   from lowest to highest.
3) Filter tblApplications to show only Client_Sector = "Technology" and Stage = "Interview".
```

### 7. Complete the Prompt_Log for every prompt you ran: mode, prompt, what Copilot did, how you verified it and your decision (Keep, Revise or Undo). Save - AutoSave keeps the cloud copy current.

## Sample Prompts

| # | Purpose | Prompt |
|---|---|---|
| 1 | Describe | Describe tblApplications: what one row represents, each column's data type, the row count and the Applied_Date range. Do not change the workbook. |
| 2 | Data quality | List any Application_ID that appears more than once in tblApplications and any Source value that is not in the canonical list on About_This_Data. |
| 3 | Clean helper column | Add a column Source_Clean that trims spaces from Source and replaces 'Linked In' with 'LinkedIn'. Do not change the original Source column. |
| 4 | Question | Using tblApplications, count applications and hires by Source_Clean and show the hire rate. Sort by hires, highest first. |
| 5 | Question | Which Client_Sector received the most applications, and how many? Show the count for every sector. |
| 6 | Highlight | Highlight rows where Stage is Hired in light green. |
| 7 | Sort | Sort tblApplications by Interview_Score from highest to lowest, then by Days_in_Pipeline from lowest to highest. |
| 8 | Filter | Filter to show only Technology applications at the Interview stage. |
| 9 | Insight | What's interesting about this pipeline data? Give three observations, each with the numbers behind it. |

## Test It

tblApplications exists with 242 data rows and =COUNTA(UNIQUE(tblApplications[Application_ID])) returns 240. The repeated IDs are APP-0018 and APP-0020 (both Rejected) and are orange. Copilot found 3 'Linked In' and 2 'JobStreet ' (trailing space) values; =COUNTIF(tblApplications[Source_Clean],"LinkedIn") returns 69. Hires total 22 and LinkedIn has the most hires (8). The Technology + Interview filter shows 9 rows. Technology received the most applications (49 unique). After clearing the filter all 242 rows are visible, and Prompt_Log has at least 5 entries.

## Checkpoint and Rejoin Point

Keep this workbook in work-c34. Lab 2 uses a separate requisitions workbook, so a learner who is behind can start Lab 2 from its own data folder.

## Troubleshooting

| If this happens | Fix |
|---|---|
| The Copilot button is missing or greyed out. | Save the file as .xlsx in OneDrive or SharePoint, turn AutoSave On, close and reopen it. Check your licence in your Microsoft 365 account and ask IT if Copilot is disabled for your tenant. Pair with a neighbour meanwhile. |
| 'Unsupported file state' message. | If the file is checked out in SharePoint, check it in and reopen it, or open it in Excel for the web. Strict Open XML files are not supported - save as a normal .xlsx. |
| Copilot answers about the whole sheet instead of the table. | Name the table (tblApplications) and the exact columns in the prompt, and click inside the table before asking. |
| Copilot's count differs from your formula. | Ask Copilot how it counted (rows vs unique IDs, which Stage values). Trust the formula, record the difference and refine the prompt. |

## Challenge

Ask Copilot to add a Pipeline_Status column that groups Stage into Active (Applied, Screening, Interview, Offer), Hired and Closed (Rejected, Withdrawn). Verify the three counts add up to the row count.

## Reflection

Which data-quality issue would most distort a source-effectiveness report if you had not found it, and why?

---

[← Labs index](../README.md) · [Lab 2 →](../lab-02-generate-and-verify-formula-columns-and-lookups-for-job-requisitions/README.md)
