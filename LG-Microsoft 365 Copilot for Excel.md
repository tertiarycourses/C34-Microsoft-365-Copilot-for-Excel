# Microsoft 365 Copilot for Excel — Learner Guide

**Course Code:** C34  |  **Conducted by:** Tertiary Infotech Academy Pte Ltd (UEN 201200696W)  |  **Version v2.0 · 8 October 2026**

## Contents

- [Introduction](#introduction)
- [Course Learning Outcomes](#course-learning-outcomes)
- [Before You Start — Preparation](#before-you-start--preparation)
- [Topic 01 — Use Copilot to Analyse Data  (Labs 1-2)](#topic-01--use-copilot-to-analyse-data--labs-1-2)
  - [What Copilot in Excel Is and How It Works](#what-copilot-in-excel-is-and-how-it-works)
  - [Requirements, Licensing, Platforms and Data Privacy](#requirements-licensing-platforms-and-data-privacy)
  - [Preparing HR and Recruitment Data for Copilot](#preparing-hr-and-recruitment-data-for-copilot)
  - [Prompting Copilot for Data Analysis](#prompting-copilot-for-data-analysis)
  - [Direct Answers, Highlighting, Sorting and Filtering](#direct-answers-highlighting-sorting-and-filtering)
  - [Formula Columns, Lookups and Formula Explanations](#formula-columns-lookups-and-formula-explanations)
  - [Lab 1 — Explore and Question Recruitment Pipeline Data](#lab-1--explore-and-question-recruitment-pipeline-data)
  - [Lab 2 — Generate and Verify Formula Columns and Lookups for Job Requisitions](#lab-2--generate-and-verify-formula-columns-and-lookups-for-job-requisitions)
  - [Topic 01 Recap - Mapped to Learning Outcomes](#topic-01-recap---mapped-to-learning-outcomes)
- [Topic 02 — Use Copilot to Plot Charts  (Labs 3-4)](#topic-02--use-copilot-to-plot-charts--labs-3-4)
  - [Choosing the Right Chart for an HR Question](#choosing-the-right-chart-for-an-hr-question)
  - [Creating Charts and PivotTables from a Prompt](#creating-charts-and-pivottables-from-a-prompt)
  - [Refining and Formatting Charts with Follow-up Prompts](#refining-and-formatting-charts-with-follow-up-prompts)
  - [Advanced Visuals with Python in Excel](#advanced-visuals-with-python-in-excel)
  - [Lab 3 — Create and Refine Recruitment Charts with Copilot](#lab-3--create-and-refine-recruitment-charts-with-copilot)
  - [Lab 4 — Visualise Pay and Placement Distributions with Python in Excel](#lab-4--visualise-pay-and-placement-distributions-with-python-in-excel)
  - [Topic 02 Recap - Mapped to Learning Outcomes](#topic-02-recap---mapped-to-learning-outcomes)
- [Topic 03 — Use Copilot to Create Dashboards  (Labs 5-6)](#topic-03--use-copilot-to-create-dashboards--labs-5-6)
  - [Designing HR and Recruitment KPIs](#designing-hr-and-recruitment-kpis)
  - [Building a Dashboard with Copilot Plan and Edit Modes](#building-a-dashboard-with-copilot-plan-and-edit-modes)
  - [Interactivity, Refresh and Maintenance](#interactivity-refresh-and-maintenance)
  - [Sharing People Dashboards Responsibly](#sharing-people-dashboards-responsibly)
  - [Lab 5 — Build a Recruitment KPI Dashboard with Copilot](#lab-5--build-a-recruitment-kpi-dashboard-with-copilot)
  - [Lab 6 — Build an Interactive Workforce and Attrition Dashboard](#lab-6--build-an-interactive-workforce-and-attrition-dashboard)
  - [Topic 03 Recap - Mapped to Learning Outcomes](#topic-03-recap---mapped-to-learning-outcomes)
- [Topic 04 — Use Copilot to Draw Insights and Recommendations  (Labs 7-8)](#topic-04--use-copilot-to-draw-insights-and-recommendations--labs-7-8)
  - [Surfacing Trends, Outliers and Drivers](#surfacing-trends-outliers-and-drivers)
  - [Advanced Analysis Mode with Python](#advanced-analysis-mode-with-python)
  - [Analysing Survey Comments: Themes and Sentiment](#analysing-survey-comments-themes-and-sentiment)
  - [From Verified Insight to Owned Recommendation](#from-verified-insight-to-owned-recommendation)
  - [Lab 7 — Surface Attrition Insights with Copilot and Advanced Analysis](#lab-7--surface-attrition-insights-with-copilot-and-advanced-analysis)
  - [Lab 8 — Turn Engagement Survey Feedback into Recommendations](#lab-8--turn-engagement-survey-feedback-into-recommendations)
  - [Topic 04 Recap - Mapped to Learning Outcomes](#topic-04-recap---mapped-to-learning-outcomes)
- [Wrap-Up - The Complete C34 Workflow](#wrap-up---the-complete-c34-workflow)
- [Next Steps](#next-steps)
- [Glossary](#glossary)


## Introduction

This Learner Guide accompanies Microsoft 365 Copilot for Excel (C34). It follows the same four topics - analyse data, plot charts, create dashboards, and draw insights and recommendations - and the same eight connected labs as the slide deck, Lesson Plan and lab folders. Every lab uses fictional HR and recruitment data from BrightPath Talent Solutions, a made-up Singapore recruitment and staffing firm.

Use the guide as a self-contained study text before, during and after class. Each concept explains what the capability is, why it matters, how it works, a worked example and when to use or avoid it, drawing on Microsoft Support and Microsoft Learn guidance. Each lab folder contains an instruction PDF with sample prompts and a mock-data workbook. Copilot output remains a proposal until you inspect it and complete the Test It checks.


## Course Learning Outcomes

- LO1: Explain how Copilot in Excel works (modes, requirements and limits) and use it to analyse, highlight, sort, filter and calculate HR and recruitment data with verified formulas.
- LO2: Use Copilot to create, refine and verify charts, PivotCharts and Python visuals that answer specific HR and recruitment questions.
- LO3: Use Copilot to build an interactive, refreshable HR KPI dashboard with PivotTables, KPI cards, charts and slicers, and check every KPI against its definition.
- LO4: Use Copilot and Python-based advanced analysis to draw evidence-backed insights and actionable, owned recommendations from workforce and survey data.


## Before You Start — Preparation

**What you need**

- A Windows or Mac laptop with Excel for Microsoft 365, or a browser for Excel for the web.
- A Microsoft 365 account with Copilot in Excel enabled (commercial Copilot licence, Copilot Chat-eligible business plan, or Microsoft 365 Personal / Family / Premium with AI credits).
- A OneDrive or SharePoint folder named work-c34 where you can save .xlsx files with AutoSave on.
- Python in Excel for Labs 4 and 7 (Windows, Mac or web; internet required). A fallback path is given if unavailable.
- The course labs folder: one sub-folder per lab with its instruction PDF, sample-prompts.txt and data/ workbook.

**Verify your setup**

Upload labs/lab-01-*/data/BrightPath_Candidate_Pipeline.xlsx to your OneDrive work-c34 folder and open it in Excel. Confirm the title bar shows the cloud location and AutoSave is On, and that Formulas > Calculation Options is Automatic. Select the Copilot button and note which modes (Allow editing, Plan, Chat only) you can see; labels may change as Microsoft updates the product.

```bash
labs/
  lab-01-<name>/
    Lab-01-Instructions.pdf     instructions + sample prompts
    README.md                   same instructions in Markdown
    sample-prompts.txt          copy-ready prompts
    data/BrightPath_<dataset>.xlsx
OneDrive or SharePoint/work-c34/   (your cloud working copies)
```

**Conventions used in every lab**

- All BrightPath companies, people, IDs and values are fictional training data.
- Prompts are shown in shaded boxes; copy them from sample-prompts.txt and adapt them.
- OBSERVED = supported by workbook evidence; HYPOTHESIS = possible explanation to test.
- S$ is used for currency, dates are yyyy-mm-dd and the reporting date is 2026-06-30.
- Log each prompt, what Copilot did, how you verified it and your decision on the Prompt_Log sheet.
- Copilot output can vary between runs; the Test It values come from the data, so check against them.


## Topic 01 — Use Copilot to Analyse Data  (Labs 1-2)

Copilot in Excel modes and requirements · preparing HR data as tables · prompting for analysis · direct answers, highlight, sort and filter · formula and lookup columns

**Key concepts**

- Three modes — Allow editing (Edit) changes the workbook, Plan proposes steps first, Chat only answers without editing.
- Requirements — Eligible license, .xlsx saved to OneDrive or SharePoint, AutoSave on, automatic calculation.
- Tables first — One header row, one record per row, one meaning per column, a named Excel table.
- Specific prompts — Goal, Context, Source and Expectations - and name the exact columns to analyse.
- Direct answers — Python-based answers in Chat; outputs inserted from Chat are static, not refreshable.
- Verify formulas — Ask Copilot to explain each formula, test edge rows and reconcile to a control total.


### What Copilot in Excel Is and How It Works

Copilot in Excel is an AI assistant, opened from the Copilot button, that plans and carries out workbook tasks in natural language using Excel's own features: formulas, tables, PivotTables, charts, conditional formatting, sorting, filtering and Python. It offers three modes - Allow editing (Edit), Plan and Chat only - and its results stay as ordinary, editable Excel content.

After a prompt, Copilot builds a step-by-step plan, executes it, then checks whether the result matches your intent; you can watch its reasoning and stop it at any time. Because Edit mode changes a shared workbook directly, choosing the mode deliberately and reviewing Show Changes or version history keeps you in control. Microsoft states it is an assistant, not an autopilot.

**How it works**

- Open the workbook, select the Copilot button and choose Chat only, Plan or Allow editing.
- Describe the task; Copilot plans, uses Excel tools, and reports what it did in the pane.
- Review the cells, objects and changes; keep, refine with a follow-up prompt, or Undo.

**Worked example**

- A BrightPath HR analyst opens the candidate pipeline workbook in Excel for the web.
- Chat only: 'Describe what each row in this table represents' - nothing in the sheet changes.
- Plan, then Allow editing: Copilot proposes and then adds a Days_Open column, which is reviewed.

**Decision guide**

| Use when | Avoid when |
|---|---|
| A task maps to Excel features you could inspect and edit yourself afterwards. | A decision about an individual (hiring, pay, exit) would rest on an unchecked AI output. |
| You can review each change and restore an earlier version if the result is wrong. | The data is not approved for use with Copilot under your organisation's policy. |

**Practitioner quality lens**

- Mode: Chat for questions, Plan to preview, Edit only for an approved change.
- Traceability: Use Show Changes, the Prompt_Log sheet and version history.
- Ownership: A named person reviews and approves the workbook before it is shared.

**Authoritative references**

- https://support.microsoft.com/en-us/excel/copilot/get-started-with-copilot-in-excel
- https://support.microsoft.com/en-us/excel/copilot/frequently-asked-questions-about-copilot-in-excel
- https://www.microsoft.com/en-us/copilot/blog/2026/06/25/copilot-in-excel-built-for-the-era-of-frontier-finance/
- https://techcommunity.microsoft.com/t5/s/gxcuf89792/attachments/gxcuf89792/HealthcareAndLifeSciencesBlog/2034.1/1/Copilot%20for%20M365%20HR%20-%20Starter%20Series.pdf

---


### Requirements, Licensing, Platforms and Data Privacy

Copilot appears in Excel for Microsoft 365 on Windows, Mac, the web and mobile when the account has an eligible plan: a commercial Microsoft 365 Copilot licence, a Copilot Chat-eligible business plan, or Microsoft 365 Personal, Family or Premium with AI credits. The workbook must be saved to OneDrive or SharePoint with AutoSave on, and Copilot editing needs Automatic calculation.

Most 'Copilot is missing' problems are setup problems: an unsaved local file, AutoSave off, a Strict Open XML file, a checked-out SharePoint file or a licence or tenant setting. For HR data, work accounts keep prompts and data inside the Microsoft 365 service boundary and respect existing permissions - but the organisation's data policy still decides what personal data may be used.

**How it works**

- Check the licence in your Microsoft 365 account; an in-app label may read M365 Copilot (Premium or Basic).
- Save as .xlsx in OneDrive or SharePoint, turn AutoSave on and set Formulas > Calculation Options > Automatic.
- Use pseudonymised IDs instead of names, and share only aggregated HR results.

**Worked example**

- A learner opens a CSV from Downloads: the Copilot button is grey.
- They save it to OneDrive/work-c34 as .xlsx, AutoSave switches on, and Copilot becomes available.
- Employee names are not needed for analysis, so the workbook uses Employee_ID only.

**Decision guide**

| Use when | Avoid when |
|---|---|
| The file is cloud-saved, AutoSave is on and your plan or tenant enables Copilot. | Edit mode on iPhone or Android, where Chat is currently the supported mode. |
| HR data is minimised to the columns the analysis actually needs. | Pasting sensitive personal data into prompts when an ID or aggregate would do. |

**Practitioner quality lens**

- Setup: Cloud .xlsx, AutoSave On, Automatic calculation, eligible licence.
- Privacy: Minimise personal data; pseudonymise; aggregate before sharing.
- Model awareness: Edit tasks may use Claude or GPT models; outputs still need review.

**Authoritative references**

- https://support.microsoft.com/en-us/excel/copilot/frequently-asked-questions-about-copilot-in-excel
- https://support.microsoft.com/en-us/microsoft-365-copilot/frequently-asked-questions-about-copilot-in-microsoft-365-subscriptions
- https://www.microsoft.com/en-us/windows/learning-center/copilot-in-microsoft-365-on-windows-11
- https://www.microsoft.com/en-us/microsoft-365-copilot/copilot-for-work
- https://learn.microsoft.com/hr-hr/copilot/microsoft-365/microsoft-365-copilot-overview

---


### Preparing HR and Recruitment Data for Copilot

Copilot works with tables, table-like ranges or several such ranges; it does not analyse unstructured data. An analysis-ready HR table has one header row, one record per row (one application, one job order, one employee), one meaning per column and consistent types. Formatting the range as a named Excel table makes its boundary explicit to people, formulas and Copilot.

Merged headings, blank rows, subtotal rows, duplicated IDs and inconsistent labels such as 'LinkedIn', 'Linked In' and 'LinkedIn ' split categories and corrupt counts. If you cannot say what one row means, no prompt can fix the logic. Recruiting data also changes daily, so a table that expands automatically keeps formulas, PivotTables and charts in step.

**How it works**

- Write the grain: 'one row = one candidate application identified by Application_ID'.
- Press Ctrl+T (Home > Format as Table), confirm headers, name it e.g. tblApplications.
- Record row count and unique-ID count; list duplicates and label variants before analysing.

**Worked example**

- BrightPath's pipeline export has 242 rows but only 240 unique Application_IDs.
- Five Source values use non-standard spellings that would split the LinkedIn and JobStreet totals.
- The analyst records these issues and asks Copilot to flag - not delete - them.

**Decision guide**

| Use when | Avoid when |
|---|---|
| The data can be expressed as a rectangular list with a stable row meaning. | A sheet mixes report titles, notes, subtotals and several tables in one range. |
| Each column has one data type and a documented list of valid values. | Fixing an issue would require guessing a missing business fact. |

**Practitioner quality lens**

- Grain: One sentence that says exactly what a row represents.
- Schema: Unique headers, consistent types and canonical category values.
- Controls: Row count, unique-ID count and key totals recorded before any change.

**Authoritative references**

- https://support.microsoft.com/en-us/excel/get-direct-answers-to-your-data-analysis-questions
- https://support.microsoft.com/en-us/office/create-and-format-tables-2f0ce0e7-500b-4fe1-9b5b-a16ba4b9f34a
- https://support.microsoft.com/en-us/excel/copilot/get-started-with-copilot-in-excel

---


### Prompting Copilot for Data Analysis

Microsoft's prompt ingredients are Goal (what you want), Context (why and for whom), Source (which data) and Expectations (how to respond). This course extends them to G-C-D-C-O-R: Goal, Context, Data, Constraints, Output and Review - so every prompt names the table and columns, what must not change, the output format and how you will check it.

Microsoft's own guidance is to be specific: naming the columns to analyse produces more accurate results, and 'Compare Q1 and Q2 hires' beats 'Tell me about hiring'. 'Summarise column C' names a task but not the goal. Follow-up prompts keep context, so refine one dimension at a time, and type 'new topic' when you switch tasks.

**How it works**

- State one goal and the HR decision it supports.
- Name the table, columns, filters, period and units; say what must stay unchanged.
- Ask for a specific output (number, table, column, chart) and for the logic to check it.

**Worked example**

- Weak: 'Analyse my recruitment data.'
- Strong: 'Using tblApplications, count applications and hires by Source for Jan-Jun 2026.'
- 'Show the result as a table sorted by hires, and explain how Hired was identified.'

**Decision guide**

| Use when | Avoid when |
|---|---|
| You can name the data boundary and the expected output. | The request is 'make this better' with no audience, rule or decision. |
| The answer has a checkable acceptance test (a count, a total, a rule). | Copilot would have to invent business thresholds or facts not in the data. |

**Practitioner quality lens**

- Specific: Name the table, columns, period, units and output.
- Bounded: State exclusions, privacy limits and what must not change.
- Reviewable: Ask for logic, assumptions and a success check.

**Authoritative references**

- https://support.microsoft.com/en-us/excel/copilot/data-insights-with-copilot-in-excel
- https://support.microsoft.com/en-us/excel/get-direct-answers-to-your-data-analysis-questions
- https://www.microsoft.com/en-us/windows/learning-center/copilot-in-microsoft-365-on-windows-11
- https://techcommunity.microsoft.com/t5/s/gxcuf89792/attachments/gxcuf89792/HealthcareAndLifeSciencesBlog/2034.1/1/Copilot%20for%20M365%20HR%20-%20Starter%20Series.pdf

---


### Direct Answers, Highlighting, Sorting and Filtering

In Chat, Copilot answers data questions directly using Python-based analysis: it interprets the question, finds the relevant data, computes the answer and can show commented code. Separately, Copilot can apply conditional formatting, sorting and filtering - for example 'Highlight the lowest values in units sold' or 'Sort engagement rate from lowest to highest' - describing each change before you apply it.

Direct answers do not modify the workbook, and tables or images inserted from them are static - they will not refresh. Highlights, sorts and filters change only what you see, not the data, yet a filter can hide the denominator. Writing the exact rule (field, operator, value) makes the view reproducible and lets a colleague check it.

**How it works**

- Ask a specific question, expand the code or ask Copilot to summarise its logic.
- Turn the HR question into explicit rules: 'Stage = Hired' or 'Interview_Score >= 4'.
- Apply one change at a time, then count visible rows and clear filters to restore the view.

**Worked example**

- 'How many applications came from each Source?' returns a direct answer with code.
- 'Highlight rows where Stage is Hired in light green' adds a conditional format rule.
- 'Filter to Technology applications at the Interview stage' shows 9 of 240 records.

**Decision guide**

| Use when | Avoid when |
|---|---|
| A quick, checkable answer or a reversible review view is needed. | Colour alone would be the only record of a status. |
| The rule can be written with named fields, operators and values. | A static inserted result will be mistaken for a live figure. |

**Practitioner quality lens**

- Rule: Write the condition before applying colour or a filter.
- Denominator: Record total rows and visible rows together.
- Static vs live: Label inserted answers 'static as at <date>'.

**Authoritative references**

- https://support.microsoft.com/en-us/excel/get-direct-answers-to-your-data-analysis-questions
- https://support.microsoft.com/en-US/Excel/copilot/visualize-your-data-with-copilot-in-excel
- https://support.microsoft.com/en-us/excel/copilot/get-started-with-copilot-in-excel

---


### Formula Columns, Lookups and Formula Explanations

Copilot can add calculated columns or summary rows (totals, durations, margins), suggest a formula for one cell, build lookups such as XLOOKUP to pull matching data from another sheet, and explain any existing formula. In a table it uses structured references such as [@Filled] and [@Fee_Rate] that fill the whole column and adjust as rows are added.

A formula can be syntactically valid and still wrong: Days_Open may ignore open requisitions, a fee may use the monthly instead of the annual salary, or blanks may silently become zero. Asking Copilot to 'Explain the formula in the selected cell', testing a few rows by hand and reconciling totals turns a generated formula from a guess into evidence.

**How it works**

- Write the business rule in words, including blanks, units and the report date.
- Ask for one column at a time and for an explanation of the formula.
- Hand-check two rows (one normal, one edge case) and compare a column total with an independent formula.

**Worked example**

- 'Add a column Est_Fee = Filled x Monthly_Salary_Budget x 12 x Fee_Rate.'
- 'Create a column that looks up Recruiter_Name from tblRecruiters using Recruiter_ID.'
- JOB-2002: 4 x S$4,900 x 12 x 25% = S$58,800 - matches the generated value.

**Decision guide**

| Use when | Avoid when |
|---|---|
| The rule is deterministic and expressible with native Excel functions. | A subjective judgement is dressed up as a precise calculation. |
| A hand-worked sample or control total can verify the result. | The formula hides missing data or an unapproved threshold. |

**Practitioner quality lens**

- Semantics: Check numerator, denominator, units and the date used.
- Edge cases: Test blanks, zero, open items and one large value.
- Reconcile: Tie the column total to an independent control.

**Authoritative references**

- https://support.microsoft.com/en-us/excel/copilot/data-insights-with-copilot-in-excel
- https://support.microsoft.com/en-us/excel/copilot/get-started-with-copilot-in-excel
- https://support.microsoft.com/en-us/office/create-and-format-tables-2f0ce0e7-500b-4fe1-9b5b-a16ba4b9f34a

---


### Lab 1 — Explore and Question Recruitment Pipeline Data

Learning outcome: LO1: confirm Copilot access, prepare a named table, and use Chat and Edit modes to question, check, highlight, sort and filter recruitment pipeline data.

Goal: Turn a raw candidate-pipeline export into a named Excel table, find its data-quality issues with Copilot, and answer pipeline questions with checked numbers.

You are a talent analyst at BrightPath Talent Solutions, a fictional Singapore recruitment and staffing firm. You will put the H1 2026 candidate-pipeline export into a named table, ask Copilot (Chat only) to describe it, use Copilot to find duplicate IDs and inconsistent Source labels, then answer pipeline questions and build a highlighted, sorted and filtered review view. Every number is checked with a formula.

**What you'll build**

BrightPath_Candidate_Pipeline.xlsx in OneDrive with table tblApplications, a Source_Clean column, duplicate IDs highlighted, Hired rows highlighted, a verified source-by-hires answer and a filled Prompt_Log   (Tools: Excel for Microsoft 365 or Excel for the web · Copilot (Chat only, Allow editing) · BrightPath_Candidate_Pipeline.xlsx.)

**Prerequisites**

- A Microsoft 365 account with Copilot in Excel enabled, signed in to Excel or Excel for the web.
- A OneDrive or SharePoint folder named work-c34 that you can upload to.
- This lab folder downloaded, including data/BrightPath_Candidate_Pipeline.xlsx.

**Step-by-step**

1. Go to onedrive.com (or your SharePoint site), open or create the folder work-c34 and select Upload > Files. Upload BrightPath_Candidate_Pipeline.xlsx from this lab's data folder, then open it in Excel for the web or select Open in Desktop App. Confirm the title bar shows the cloud location and AutoSave is On, and that Formulas > Calculation Options is set to Automatic. Read the About_This_Data sheet before you prompt.

   ```bash
   Required state:
File      : OneDrive/work-c34/BrightPath_Candidate_Pipeline.xlsx
AutoSave  : On
Calculation Options : Automatic
Sheets    : About_This_Data | Applications | Prompt_Log
   ```

2. Select cell A1 on the Applications sheet and press Ctrl+T (Home > Format as Table). Tick My table has headers and select OK. On the Table Design tab, set Table Name to tblApplications. Then select the Copilot button. If Copilot is unavailable, use the Troubleshooting table before continuing.

   ```bash
   Table name : tblApplications
Grain      : one row = one candidate application (Application_ID)
Expected   : 13 columns, 242 data rows
   ```

3. In the Copilot pane, choose Chat only (no workbook changes). Paste the prompt below. Compare Copilot's row count and date range with the table (select the Application_ID column and read Count on the status bar). Record the prompt, what Copilot said and your check on the Prompt_Log sheet.

   ```bash
   GOAL: Describe this candidate pipeline before I analyse it.
CONTEXT: BrightPath recruitment desk, H1 2026. One row should be one application.
DATA: Use only tblApplications.
CONSTRAINTS: Do not change the workbook.
OUTPUT: What a row represents, each column with its data type, the number of rows,
        the Applied_Date range and the distinct values of Stage and Source.
REVIEW: Tell me how you counted the rows.
   ```

4. Still in Chat only, ask Copilot to check data quality. It should report 2 repeated Application_IDs and Source values that are not in the canonical list on About_This_Data. Then switch to Allow editing (Edit) and ask Copilot to highlight the duplicate IDs and add a Source_Clean column. Check that no row was deleted - the table must still have 242 rows.

   ```bash
   CHAT ONLY:
Check tblApplications for data-quality issues. List any Application_ID that appears more than once,
and any Source value that is not exactly one of: LinkedIn, JobStreet, MyCareersFuture, Referral,
BrightPath Website, Job Fair, Talent Database. Give the count of each issue. Do not change anything.

ALLOW EDITING:
1) Highlight Application_ID cells that appear more than once in tblApplications with an orange fill.
2) Add a column Source_Clean to tblApplications that trims spaces from Source and changes 'Linked In'
   to 'LinkedIn'. Do not delete or edit any existing values.
   ```

5. Ask Copilot a direct question about which sources produce hires. Expand or ask Copilot to summarise any code it shows. Then verify the hires total in an empty cell on Prompt_Log with the formula below, and verify one source with COUNTIFS. Duplicate rows are Rejected applications, so they do not change the hire counts.

   ```bash
   PROMPT (Chat only):
Using tblApplications, count applications and hires (Stage = "Hired") by Source_Clean.
Show a table sorted by hires, highest first, with hire rate = hires / applications.

CHECK FORMULAS:
=COUNTIFS(tblApplications[Stage],"Hired")
=COUNTIFS(tblApplications[Source_Clean],"LinkedIn",tblApplications[Stage],"Hired")
   ```

6. Switch to Allow editing and build a review view, one change at a time. After each change, inspect the sheet and keep it or Undo. Finally apply the filter, read the visible-row count on the status bar, record it, and clear the filter so all 242 rows return.

   ```bash
   1) Highlight entire rows of tblApplications where Stage is "Hired" with a light green fill.
2) Sort tblApplications by Interview_Score from highest to lowest, then by Days_in_Pipeline
   from lowest to highest.
3) Filter tblApplications to show only Client_Sector = "Technology" and Stage = "Interview".
   ```

7. Complete the Prompt_Log for every prompt you ran: mode, prompt, what Copilot did, how you verified it and your decision (Keep, Revise or Undo). Save - AutoSave keeps the cloud copy current.

**Sample prompts**

| Purpose | Prompt |
|---|---|
| Describe | Describe tblApplications: what one row represents, each column's data type, the row count and the Applied_Date range. Do not change the workbook. |
| Data quality | List any Application_ID that appears more than once in tblApplications and any Source value that is not in the canonical list on About_This_Data. |
| Clean helper column | Add a column Source_Clean that trims spaces from Source and replaces 'Linked In' with 'LinkedIn'. Do not change the original Source column. |
| Question | Using tblApplications, count applications and hires by Source_Clean and show the hire rate. Sort by hires, highest first. |
| Question | Which Client_Sector received the most applications, and how many? Show the count for every sector. |
| Highlight | Highlight rows where Stage is Hired in light green. |
| Sort | Sort tblApplications by Interview_Score from highest to lowest, then by Days_in_Pipeline from lowest to highest. |
| Filter | Filter to show only Technology applications at the Interview stage. |
| Insight | What's interesting about this pipeline data? Give three observations, each with the numbers behind it. |

**Test it**

tblApplications exists with 242 data rows and =COUNTA(UNIQUE(tblApplications[Application_ID])) returns 240. The repeated IDs are APP-0018 and APP-0020 (both Rejected) and are orange. Copilot found 3 'Linked In' and 2 'JobStreet ' (trailing space) values; =COUNTIF(tblApplications[Source_Clean],"LinkedIn") returns 69. Hires total 22 and LinkedIn has the most hires (8). The Technology + Interview filter shows 9 rows. Technology received the most applications (49 unique). After clearing the filter all 242 rows are visible, and Prompt_Log has at least 5 entries.

**Checkpoint and rejoin point**

Keep this workbook in work-c34. Lab 2 uses a separate requisitions workbook, so a learner who is behind can start Lab 2 from its own data folder.

**Troubleshooting**

| If this happens | Fix |
|---|---|
| The Copilot button is missing or greyed out. | Save the file as .xlsx in OneDrive or SharePoint, turn AutoSave On, close and reopen it. Check your licence in your Microsoft 365 account and ask IT if Copilot is disabled for your tenant. Pair with a neighbour meanwhile. |
| 'Unsupported file state' message. | If the file is checked out in SharePoint, check it in and reopen it, or open it in Excel for the web. Strict Open XML files are not supported - save as a normal .xlsx. |
| Copilot answers about the whole sheet instead of the table. | Name the table (tblApplications) and the exact columns in the prompt, and click inside the table before asking. |
| Copilot's count differs from your formula. | Ask Copilot how it counted (rows vs unique IDs, which Stage values). Trust the formula, record the difference and refine the prompt. |

**Challenge**

Ask Copilot to add a Pipeline_Status column that groups Stage into Active (Applied, Screening, Interview, Offer), Hired and Closed (Rejected, Withdrawn). Verify the three counts add up to the row count.

**Reflection**

Which data-quality issue would most distort a source-effectiveness report if you had not found it, and why?

> **Note:** The lab folder labs/lab-01-*/ holds the instruction PDF with sample prompts, sample-prompts.txt and the mock-data workbook (data/BrightPath_Candidate_Pipeline.xlsx). Use only the supplied fictional BrightPath data or information you are authorised to process. Do not paste credentials, confidential business material or real employees' personal data into Copilot.

---


### Lab 2 — Generate and Verify Formula Columns and Lookups for Job Requisitions

Learning outcome: LO1: use Copilot to generate lookup and calculated columns from plain-English rules, explain them, and reconcile them to independent control formulas.

Goal: Add recruiter names, days open, fill rate, estimated placement fee and an ageing flag to a requisitions table, then prove each column is right.

BrightPath's permanent-recruitment desk tracks 60 client job orders. You will ask Copilot to add an XLOOKUP for the recruiter's name and four calculated columns written from business rules, ask it to explain each formula, hand-check one requisition and reconcile column totals to independent formulas on the Control sheet. You will then ask Copilot which recruiter generated the most estimated fees and verify the answer.

**What you'll build**

BrightPath_Job_Requisitions.xlsx with Recruiter_Name, Days_Open, Fill_Rate, Est_Fee and Ageing_Flag columns in tblRequisitions, overdue requisitions highlighted, and a Control sheet that reconciles the totals   (Tools: Excel · Copilot (Allow editing, Chat only) · XLOOKUP · structured references · BrightPath_Job_Requisitions.xlsx.)

**Prerequisites**

- Lab 1 completed, or Copilot access confirmed on a OneDrive file with AutoSave On.
- This lab folder's data/BrightPath_Job_Requisitions.xlsx.

**Step-by-step**

1. Go to onedrive.com (or your SharePoint site), open or create the folder work-c34 and select Upload > Files. Upload BrightPath_Job_Requisitions.xlsx from this lab's data folder, then open it in Excel for the web or select Open in Desktop App. Confirm the title bar shows the cloud location and AutoSave is On, and that Formulas > Calculation Options is set to Automatic. Read the About_This_Data sheet before you prompt. Note the fee rule: Est_Fee = Filled x Monthly_Salary_Budget x 12 x Fee_Rate, and the report date in Control!B1.

   ```bash
   Tables : tblRequisitions (60 rows) | tblRecruiters (6 rows)
Control!B1 Report_Date = 2026-06-30
   ```

2. Choose Allow editing. Ask Copilot to add the lookup column. Select a cell in the new column and ask Copilot to explain the formula. Check one row by hand against tblRecruiters.

   ```bash
   In tblRequisitions, add a column Recruiter_Name that looks up Recruiter_Name from tblRecruiters
using Recruiter_ID. Use XLOOKUP and return "Not found" if there is no match.

Then: Explain the formula in the selected cell.
   ```

3. Ask for Days_Open with an explicit rule for open requisitions. Read the explanation. Test an open row (blank Close_Date) and a closed row by subtracting the dates yourself.

   ```bash
   Add a column Days_Open to tblRequisitions. If Close_Date is blank, use Control!B1 minus Open_Date;
otherwise use Close_Date minus Open_Date. Show whole numbers.
   ```

4. Ask for Fill_Rate and Est_Fee, one prompt each. Format Fill_Rate as a percentage and Est_Fee as S$ with no decimals. Hand-check JOB-2002: Filled 4 x S$4,900 x 12 x 25% = S$58,800.

   ```bash
   Add a column Fill_Rate = Filled divided by Openings, formatted as a percentage.

Add a column Est_Fee = Filled x Monthly_Salary_Budget x 12 x Fee_Rate, formatted as S$ with no decimals.
Explain the formula.
   ```

5. Ask for the ageing flag and a highlight. Count the flags with the status bar or a COUNTIF.

   ```bash
   Add a column Ageing_Flag that shows "Overdue" when Status is "Open" and Days_Open is greater than 45,
otherwise leave it blank. Highlight rows where Ageing_Flag is "Overdue" with a light amber fill.
   ```

6. Build independent controls on the Control sheet using the formulas below (type them yourself - do not reuse Copilot's columns for the control). Each pair must match.

   ```bash
   A3 Total Est_Fee (column)       B3 =SUM(tblRequisitions[Est_Fee])
A4 Total fee (independent)      B4 =SUMPRODUCT(tblRequisitions[Filled],tblRequisitions[Monthly_Salary_Budget],tblRequisitions[Fee_Rate])*12
A5 Avg days to fill (Filled)    B5 =AVERAGEIFS(tblRequisitions[Days_Open],tblRequisitions[Status],"Filled")
A6 Overdue (independent)        B6 =COUNTIFS(tblRequisitions[Status],"Open",tblRequisitions[Days_Open],">45")
A7 Overdue flags (column)       B7 =COUNTIF(tblRequisitions[Ageing_Flag],"Overdue")
   ```

7. In Chat only, ask which recruiter generated the most estimated fees. Verify with SUMIFS on Control, then log all prompts on Prompt_Log.

   ```bash
   Using tblRequisitions, which Recruiter_Name has the highest total Est_Fee? Show the total for every recruiter.

CHECK: =SUMIFS(tblRequisitions[Est_Fee],tblRequisitions[Recruiter_Name],"<name>")
   ```


**Sample prompts**

| Purpose | Prompt |
|---|---|
| Lookup | Add a column Recruiter_Name that looks up Recruiter_Name from tblRecruiters using Recruiter_ID with XLOOKUP. Return "Not found" if missing. |
| Duration | Add a column Days_Open: if Close_Date is blank use Control!B1 minus Open_Date, otherwise Close_Date minus Open_Date. |
| Ratio | Add a column Fill_Rate = Filled / Openings formatted as a percentage. |
| Business rule | Add a column Est_Fee = Filled x Monthly_Salary_Budget x 12 x Fee_Rate, formatted as S$ with no decimals. |
| Flag | Add a column Ageing_Flag that shows "Overdue" when Status is Open and Days_Open > 45, otherwise blank. |
| Explain | Explain the formula in the selected cell step by step, and tell me how blanks are handled. |
| Summary row | Add a total row to tblRequisitions that sums Openings, Filled and Est_Fee and averages Days_Open. |
| Question | Which recruiter has the highest total Est_Fee? Show every recruiter's total, highest first. |

**Test it**

Control!B3 and B4 both show S$1,030,512. B5 shows 34.6 days (29 Filled requisitions). B6 and B7 both show 13. JOB-2002 shows Est_Fee S$58,800 and Days_Open 36. Every Recruiter_Name resolves (no "Not found"). The top recruiter by estimated fees is Aisha Rahman with S$221,760.

**Checkpoint and rejoin point**

Keep the workbook in work-c34. Topic 2 starts from a new metrics workbook in the Lab 3 folder.

**Troubleshooting**

| If this happens | Fix |
|---|---|
| Days_Open is negative or huge for open requisitions. | The formula used today's date or a blank as zero. Ask Copilot to use Control!B1 explicitly, or edit the formula. |
| Est_Fee is 12 times too small or too large. | Re-read the rule: fee is on first-year salary (monthly x 12). Ask Copilot to explain the formula and correct it. |
| The new column is outside the table. | Click inside tblRequisitions before prompting and name the table in the prompt; or drag the table's resize handle. |
| Copilot cannot add a formula column. | Confirm Calculation Options is Automatic and the data is an Excel table; try Excel for the web if desktop fails. |

**Challenge**

Ask Copilot to add a Fee_Achieved column that is Est_Fee for Filled requisitions only, then build a small summary of Fee_Achieved by Sector with SUMIFS. Check that the sector totals add up to the grand total.

**Reflection**

Which of today's formulas would have looked correct but been wrong without a hand-check, and how did you catch it?

> **Note:** The lab folder labs/lab-02-*/ holds the instruction PDF with sample prompts, sample-prompts.txt and the mock-data workbook (data/BrightPath_Job_Requisitions.xlsx). Use only the supplied fictional BrightPath data or information you are authorised to process. Do not paste credentials, confidential business material or real employees' personal data into Copilot.

---


### Topic 01 Recap - Mapped to Learning Outcomes

Use this checkpoint to explain what you can now do and identify any lab evidence you still need to repair.

| Learning outcome | Evidence from this topic |
|---|---|
| LO1 - Analyse data with Copilot | Confirm license, OneDrive/AutoSave and mode; format data as a named table; ask specific questions that name columns; highlight, sort and filter with explicit rules; generate formula and lookup columns and reconcile them to independent checks. |

---


## Topic 02 — Use Copilot to Plot Charts  (Labs 3-4)

Choosing the right chart for an HR question · charts and PivotCharts from a prompt · refining charts with follow-up prompts · advanced visuals with Python in Excel

**Key concepts**

- Question first — Comparison, trend, composition, distribution or relationship decides the chart type.
- Describe the chart — Name x-axis, y-axis, measure, aggregation, filters and chart type in the prompt.
- Pivot behind the chart — Summarise with a PivotTable first, so the chart has a checkable total.
- Refine by follow-up — Titles, labels, sort order, colours and axis scales are prompt-able edits.
- Python visuals — =PY cells run in the Microsoft cloud with Anaconda libraries for box plots and heatmaps.
- Honest charts — Units, full categories, zero-based bars and a source note on every chart.


### Choosing the Right Chart for an HR Question

Every chart answers one kind of question. Comparison across categories uses a column or bar chart; change over time uses a line chart; a recruitment funnel uses a funnel or ordered bar chart; distribution of salaries or rates uses a histogram or box plot; the relationship between two measures uses a scatter plot; and share of a whole uses a stacked bar or, sparingly, a pie.

Copilot will draw whatever you ask for, quickly and polished - but a pie chart of 18 months, or an average of averages, can mislead an HR leadership team. Deciding the question, the grain and the measure before prompting is the analyst's job; Copilot then removes the clicking.

**How it works**

- Write the question: 'Which client sector produced the most hires in H1 2026?'
- Classify it (comparison) and choose the form (sorted bar chart).
- Name the measure and aggregation (Sum of Hires) and the filter (2026 only).

**Worked example**

- Trend: monthly hires Jan 2025 - Jun 2026 as a line chart.
- Funnel: Applications > Screened > Interviews > Offers > Hires for H1 2026.
- Efficiency: hire rate and cost per hire by sourcing channel as a bar or combo chart.

**Decision guide**

| Use when | Avoid when |
|---|---|
| The audience needs one clear message from the visual. | A pie with many slices or a 3-D effect would distort comparison. |
| The categories and time order are complete and consistent. | Rates are averaged row by row instead of recomputed from totals. |

**Practitioner quality lens**

- Fit: Chart type matches comparison, trend, funnel, distribution or relationship.
- Message: Title states the finding, e.g. 'Technology led H1 hires'.
- Integrity: Zero-based bars, units shown, no hidden categories.

**Authoritative references**

- https://support.microsoft.com/en-US/Excel/copilot/visualize-your-data-with-copilot-in-excel
- https://support.microsoft.com/en-us/excel/overview-of-pivottables-and-pivotcharts
- https://www.microsoft.com/en-gb/microsoft-365/excel/ai-for-excel

---


### Creating Charts and PivotTables from a Prompt

Describe the chart you want and Copilot generates it, linked to the source data so it stays editable. Microsoft's examples include 'Create a line chart showing change over time. Show data labels and gridlines.' and 'Create a PivotTable with the sum of sales, with month in rows and category in columns.' The more detail you include, the more useful the result.

A PivotTable summarises large data sets without manual steps and gives the chart a visible, checkable summary. If the first result does not match your intent, add detail - the field, aggregation, filter or chart type - and ask again rather than starting from scratch.

**How it works**

- Prompt for the PivotTable: rows, columns, values, aggregation and filters.
- Prompt for the chart from that PivotTable and say where to place it (the Charts sheet).
- Check the Pivot grand total against a SUM or SUMIFS of the source table.

**Worked example**

- 'Create a PivotTable from tblMonthly with Month_Label in rows, Sector in columns and Sum of Hires.'
- 'Create a line chart of total hires by month from that PivotTable, with data labels.'
- Check: Pivot grand total for 2026 = 361 hires = SUMIFS(tblMonthly[Hires], ...).

**Decision guide**

| Use when | Avoid when |
|---|---|
| The data is in a table and the measure has a defined aggregation. | The question, denominator or aggregation has not been agreed yet. |
| You need an editable chart that updates when the PivotTable refreshes. | A static image from Chat would be pasted into a recurring report. |

**Practitioner quality lens**

- Grain: Know what one Pivot cell represents.
- Tie-out: Pivot grand total equals the source-table control.
- Editable: Chart is linked to data, not a pasted picture.

**Authoritative references**

- https://support.microsoft.com/en-US/Excel/copilot/visualize-your-data-with-copilot-in-excel
- https://support.microsoft.com/en-us/excel/copilot/get-started-with-copilot-in-excel
- https://support.microsoft.com/en-us/excel/overview-of-pivottables-and-pivotcharts

---


### Refining and Formatting Charts with Follow-up Prompts

Follow-up prompts keep the conversation's context, so you can refine a chart step by step: change the type, sort the bars, add a secondary axis, rename the title to state the finding, format axis numbers as S$ or %, apply consistent colours and add alt text. Each refinement is still an ordinary Excel chart setting you can inspect in Format Chart.

First drafts rarely communicate. A combo chart that puts cost per hire (S$) and hire rate (%) on one axis is unreadable; an unsorted bar chart hides the ranking; a generic title such as 'Sum of Hires' tells leadership nothing. Small, specific follow-ups turn a default chart into a decision visual.

**How it works**

- Ask for one refinement per prompt: type, sort, axis, labels, title or colours.
- Use a secondary axis only when measures have different units, and label both axes.
- Add a source note and alt text, then check the chart against the PivotTable values.

**Worked example**

- 'Sort the bars from highest to lowest hire rate.'
- 'Change the title to: Referrals convert best - hire rate by source, Jan-Jun 2026.'
- 'Put Cost per Hire on a secondary axis as a line and format it as S$.'

**Decision guide**

| Use when | Avoid when |
|---|---|
| The chart is right in substance but not yet clear to its audience. | A refinement would change the meaning, e.g. truncating a bar axis. |
| You want consistent formatting across several charts or a dashboard. | Many edits are bundled into one vague prompt such as 'make it pretty'. |

**Practitioner quality lens**

- Readable: Finding-style title, labelled axes, units and legend.
- Accessible: Colour-blind-safe palette and alt text.
- Accurate: Values still match the PivotTable after formatting.

**Authoritative references**

- https://support.microsoft.com/en-US/Excel/copilot/visualize-your-data-with-copilot-in-excel
- https://support.microsoft.com/en-us/excel/copilot/get-started-with-copilot-in-excel
- https://support.microsoft.com/en-us/excel/copilot/data-insights-with-copilot-in-excel

---


### Advanced Visuals with Python in Excel

Python in Excel runs Python code in worksheet cells: type =PY or use Formulas > Insert Python. The code runs in the Microsoft cloud on a standard Anaconda distribution, so no local install is needed, and results return to the grid. Copilot can write the Python for you - including box plots, histograms, heatmaps and regression lines that native charts cannot easily produce.

HR questions are often about spread, not averages: pay-rate ranges by job level, outliers, or whether experience relates to pay. Python visuals show distributions honestly. But Python in Excel needs internet access, is available on Windows, Mac and the web (not mobile), and its code and data choices must be reviewed like any formula.

**How it works**

- Ask Copilot for the visual and to use Python, naming the table and columns.
- Open the generated =PY cell and read the code: data range, grouping, filters.
- Cross-check one plotted statistic (e.g. a median) with an Excel formula.

**Worked example**

- 'Using Python, create a box plot of Bill_Rate_Hourly by Job_Level in tblPlacements.'
- The Lead median bill rate in the plot is checked with MEDIAN(IF(...)) in Excel.
- 'Create a heatmap of average gross margin % by Sector and Job_Level.'

**Decision guide**

| Use when | Avoid when |
|---|---|
| The question is about distribution, correlation or many groups at once. | The workbook will be used offline or on a phone or tablet. |
| Your licence and platform support Python in Excel and a reviewer can read the code. | A simple native column or line chart would answer the question. |

**Practitioner quality lens**

- Inputs: Code references the right table, columns and filters.
- Method: Statistic (median, IQR, correlation) suits the question.
- Reproducible: One plotted value re-computed with a native formula.

**Authoritative references**

- https://support.microsoft.com/en-us/excel/python/introduction-to-python-in-excel
- https://support.microsoft.com/en-us/excel/get-direct-answers-to-your-data-analysis-questions
- https://support.microsoft.com/en-us/office/get-advanced-data-analysis-using-copilot-in-excel-app-skills-364e4ae9-9343-4d56-952a-5f62b0f70db6

---


### Lab 3 — Create and Refine Recruitment Charts with Copilot

Learning outcome: LO2: use Copilot to build PivotTables and trend, comparison, funnel and combo charts from prompts, refine them with follow-up prompts and tie each chart to a checked total.

Goal: Produce four decision-ready recruitment charts - hiring trend, hires by sector, the H1 2026 funnel and source effectiveness - each backed by a reconciled summary.

BrightPath's leadership wants to see how hiring is moving and which sectors and sourcing channels perform best. You will ask Copilot to build PivotTables and charts from 18 months of funnel metrics and a sourcing summary, choose each chart from the question it answers, refine titles, labels, sort order and axes with follow-up prompts, and reconcile every chart to an independent SUMIFS check.

**What you'll build**

BrightPath_Recruitment_Metrics.xlsx with PivotTables and four refined charts on the Charts sheet: a monthly hires line, a sorted 2026 hires-by-sector bar, a 2026 recruitment funnel and a hire-rate vs cost-per-hire combo chart   (Tools: Excel · Copilot (Allow editing) · PivotTables · line, bar, funnel and combo charts · BrightPath_Recruitment_Metrics.xlsx.)

**Prerequisites**

- Topic 1 completed: you can open a OneDrive workbook, use Copilot modes and verify with formulas.
- This lab folder's data/BrightPath_Recruitment_Metrics.xlsx.

**Step-by-step**

1. Go to onedrive.com (or your SharePoint site), open or create the folder work-c34 and select Upload > Files. Upload BrightPath_Recruitment_Metrics.xlsx from this lab's data folder, then open it in Excel for the web or select Open in Desktop App. Confirm the title bar shows the cloud location and AutoSave is On, and that Formulas > Calculation Options is set to Automatic. Read the About_This_Data sheet before you prompt. Write the four questions below at the top of the Charts sheet - each chart must answer one of them.

   ```bash
   Tables : tblMonthly (108 rows = 18 months x 6 sectors) | tblSources (7 rows)

Q1 Trend      : How have total hires moved month by month since Jan 2025?
Q2 Comparison : Which sector produced the most hires in Jan-Jun 2026?
Q3 Funnel     : Where do candidates drop out between application and hire in 2026?
Q4 Efficiency : Which sources convert best, and at what cost per hire?
   ```

2. Choose Allow editing. Ask Copilot for the trend PivotTable and line chart. Check the Pivot grand total against =SUM(tblMonthly[Hires]) and find the peak month.

   ```bash
   Create a PivotTable from tblMonthly on a new sheet named Pivot_Trend with Month_Label in rows and
Sum of Hires as values. Then create a line chart of total hires by month from that PivotTable.
Show data labels and gridlines. Title: "Monthly hires, Jan 2025 - Jun 2026".
   ```

3. Ask for the 2026 sector comparison. Then refine it with a follow-up prompt. Verify Technology's 2026 hires with the SUMIFS shown.

   ```bash
   Create a PivotTable on a new sheet Pivot_Sector from tblMonthly with Sector in rows and Sum of Hires,
filtered to months in 2026 only. Create a bar chart from it.

FOLLOW-UP: Sort the bars from highest to lowest and change the title to state the leading sector.

CHECK: =SUMIFS(tblMonthly[Hires],tblMonthly[Sector],"Technology",tblMonthly[Month],">="&DATE(2026,1,1))
   ```

4. Ask Copilot to build a 2026 funnel summary with formulas, then a funnel chart. Confirm the summary uses SUMIFS (not typed numbers) by clicking a value cell.

   ```bash
   On the Charts sheet, build a small table of 2026 totals for Applications, Screened, Interviews, Offers
and Hires from tblMonthly using SUMIFS on Month >= 1 Jan 2026. Then create a funnel chart from it
with data labels. Add a column showing each stage as a % of Applications.
   ```

5. Ask Copilot to add rate columns to tblSources and a combo chart with two axes. Check the best hire rate and the lowest cost per hire.

   ```bash
   In tblSources add Hire_Rate = Hires / Applications (percentage, 1 decimal) and
Cost_per_Hire = Sourcing_Spend / Hires (S$, no decimals). Create a combo chart: Hire_Rate as
clustered columns sorted highest first, Cost_per_Hire as a line on a secondary axis.
Label both axes with units.
   ```

6. Refine all four charts with short follow-up prompts, one change at a time, then move them onto the Charts sheet in a 2 x 2 grid. Add alt text to each chart.

   ```bash
   Make every chart title state the finding in one sentence.
Use the same blue colour for all hires series and format numbers with thousands separators.
Move all four charts to the Charts sheet in a 2 x 2 grid of equal size.
Add alt text to each chart describing what it shows.
   ```

7. Ask Copilot (Chat only) to review the charts against the four questions, then log your prompts and checks on Prompt_Log.

   ```bash
   Review the four charts on the Charts sheet. For each, tell me which question at the top of the sheet
it answers, whether the chart type suits that question, and one improvement. Do not change anything.
   ```


**Sample prompts**

| Purpose | Prompt |
|---|---|
| PivotTable | Create a PivotTable from tblMonthly with Month_Label in rows, Sector in columns and Sum of Hires as values. |
| Line chart | Create a line chart showing total hires by month. Show data labels and gridlines. |
| Bar chart | Create a bar chart of 2026 hires by Sector, sorted from highest to lowest. |
| Funnel | Create a funnel chart of 2026 Applications, Screened, Interviews, Offers and Hires with data labels. |
| Combo chart | Create a combo chart of Hire_Rate as columns and Cost_per_Hire as a line on a secondary axis. |
| Refine | Change the chart title to: Technology led hiring in H1 2026. |
| Refine | Sort the bars from highest to lowest and add data labels formatted with thousands separators. |
| Accessibility | Add alt text to this chart that describes the trend it shows. |
| Review | Is a line chart the right choice for this data? Suggest a better chart if not, and explain why. |

**Test it**

Pivot_Trend grand total = 1,065 hires (= SUM(tblMonthly[Hires])); the peak month is 2025-11 with 68 hires. The 2026 sector bar chart shows Technology first with 84 hires and its Pivot total is 361. The funnel shows 6,616 > 3,229 > 1,385 > 495 > 361. In tblSources, Referral has the highest hire rate (7.4%) and MyCareersFuture the lowest cost per hire (S$186). All four charts sit on the Charts sheet with finding-style titles, labelled axes and alt text.

**Checkpoint and rejoin point**

Keep the workbook in work-c34. Lab 4 uses its own placements workbook.

**Troubleshooting**

| If this happens | Fix |
|---|---|
| The 2026 filter includes 2025 months. | Ask Copilot to filter on the Month date field (>= 1 Jan 2026), not on Month_Label text, or add a Year field. |
| Copilot inserts a picture of a chart instead of a chart. | That came from a Chat answer and is static. Switch to Allow editing and ask it to create a chart linked to the data. |
| The funnel chart type is not offered. | Funnel charts need Excel 2019 or Microsoft 365. Ask for a sorted horizontal bar chart of the stages instead. |
| Both series in the combo chart share one axis and the line is flat. | Ask: 'Put Cost_per_Hire on a secondary axis and format it as S$'. |

**Challenge**

Ask Copilot for a 100% stacked column chart showing each sector's share of hires by quarter for 2025-2026, then decide whether it communicates better or worse than the bar chart, and why.

**Reflection**

Which of your four charts would most change a BrightPath leadership decision, and what check makes you trust it?

> **Note:** The lab folder labs/lab-03-*/ holds the instruction PDF with sample prompts, sample-prompts.txt and the mock-data workbook (data/BrightPath_Recruitment_Metrics.xlsx). Use only the supplied fictional BrightPath data or information you are authorised to process. Do not paste credentials, confidential business material or real employees' personal data into Copilot.

---


### Lab 4 — Visualise Pay and Placement Distributions with Python in Excel

Learning outcome: LO2: use Copilot with Python in Excel to create distribution, relationship and heatmap visuals, review the generated code and verify plotted statistics with native formulas.

Goal: Show how contract bill rates, pay and margins are distributed across job levels and sectors, and explain one outlier, using Python visuals you have checked.

BrightPath's flexible-staffing desk places contract associates with clients. You will ask Copilot to add a gross margin column, then create Python in Excel visuals - a box plot of bill rates by job level, a scatter of experience against pay with a trend line, and a sector-by-level margin heatmap. You will read the =PY code, re-compute a median and a correlation with Excel formulas, and investigate the highest pay rate.

**What you'll build**

BrightPath_Contract_Placements.xlsx with a Gross_Margin_Pct column and, on Python_Charts, a box plot, a scatter with trend line, a margin heatmap and an outlier table - each with a native-formula check beside it   (Tools: Excel · Copilot (Allow editing, Chat only) · Python in Excel (=PY) · BrightPath_Contract_Placements.xlsx.)

**Prerequisites**

- Lab 3 completed or equivalent charting experience.
- Python in Excel available: Formulas > Insert Python is visible (Windows, Mac or web; internet required).
- This lab folder's data/BrightPath_Contract_Placements.xlsx.

**Step-by-step**

1. Go to onedrive.com (or your SharePoint site), open or create the folder work-c34 and select Upload > Files. Upload BrightPath_Contract_Placements.xlsx from this lab's data folder, then open it in Excel for the web or select Open in Desktop App. Confirm the title bar shows the cloud location and AutoSave is On, and that Formulas > Calculation Options is set to Automatic. Read the About_This_Data sheet before you prompt. Check that the Formulas tab shows Insert Python. If it does not, follow the fallback in Troubleshooting.

   ```bash
   Table  : tblPlacements (180 rows)
Levels : Junior | Mid | Senior | Lead
Output sheet : Python_Charts
   ```

2. Choose Allow editing and ask for the margin column. Check one row by hand and the average with AVERAGE.

   ```bash
   Add a column Gross_Margin_Pct to tblPlacements = (Bill_Rate_Hourly - Pay_Rate_Hourly) / Bill_Rate_Hourly,
formatted as a percentage with 1 decimal.

CHECK: =AVERAGE(tblPlacements[Gross_Margin_Pct])
   ```

3. Ask Copilot to create the box plot with Python. Click the =PY cell and read the code: which table and columns it uses and the order of levels. Then verify the Lead median with a native formula in the cell beside it.

   ```bash
   Using Python in Excel, create a box plot of Bill_Rate_Hourly by Job_Level from tblPlacements,
ordered Junior, Mid, Senior, Lead. Put it in Python_Charts!B2. Title it
"Contract bill rate by job level (S$/hour)" and label both axes.

CHECK: =MEDIAN(FILTER(tblPlacements[Bill_Rate_Hourly],tblPlacements[Job_Level]="Lead"))
   ```

4. Ask for a scatter plot with a trend line and the correlation coefficient. Verify r with CORREL.

   ```bash
   Using Python, create a scatter plot of Years_Experience (x) against Pay_Rate_Hourly (y) from
tblPlacements, coloured by Job_Level, with a linear trend line. Also return the Pearson correlation
coefficient in the cell below the chart.

CHECK: =CORREL(tblPlacements[Years_Experience],tblPlacements[Pay_Rate_Hourly])
   ```

5. Ask for a heatmap of average margin. Read the code to confirm it uses the mean of Gross_Margin_Pct per cell.

   ```bash
   Using Python, create a heatmap of the average Gross_Margin_Pct by Sector (rows) and Job_Level (columns,
ordered Junior to Lead) with the values printed in each cell as percentages.
   ```

6. Ask Copilot (Chat only) to find outliers with a stated rule, then look at the highest pay rate row. Decide in Prompt_Log whether it is an error, a genuine specialist rate, or needs checking with the desk - an outlier is not automatically wrong.

   ```bash
   Using the IQR rule (above Q3 + 1.5 x IQR), list the Pay_Rate_Hourly outliers in tblPlacements with
Placement_ID, Sector, Job_Level and Pay_Rate_Hourly. State Q1, Q3 and the upper fence. Show the code.
   ```

7. Ask Copilot to explain what the Python code in the box-plot cell does, then record all prompts and checks in Prompt_Log.

   ```bash
   Explain the Python code in Python_Charts!B2 line by line in plain English.
   ```


**Sample prompts**

| Purpose | Prompt |
|---|---|
| Calculated column | Add Gross_Margin_Pct = (Bill_Rate_Hourly - Pay_Rate_Hourly) / Bill_Rate_Hourly as a percentage. |
| Box plot | Using Python, create a box plot of Bill_Rate_Hourly by Job_Level ordered Junior, Mid, Senior, Lead. |
| Histogram | Using Python, create a histogram of Pay_Rate_Hourly with 15 bins and mark the median. |
| Scatter | Using Python, plot Years_Experience against Pay_Rate_Hourly with a linear trend line and return the correlation. |
| Heatmap | Using Python, create a heatmap of average Gross_Margin_Pct by Sector and Job_Level with values shown. |
| Outliers | Using the IQR rule, list Pay_Rate_Hourly outliers with Placement_ID, Sector and Job_Level, and state the upper fence. |
| Explain code | Explain the Python code in the selected cell line by line. |
| Direct answer | What are the median and mean Bill_Rate_Hourly by Job_Level? Show the code. |

**Test it**

Average Gross_Margin_Pct = 22.6%. The box plot shows four levels in order and the Lead median bill rate check returns S$65.74 (Junior S$22.72, Mid S$29.53, Senior S$41.70). CORREL returns about 0.49 - a moderate positive relationship. The heatmap shows Technology with the lowest average margin (~19.1%). The IQR rule flags 8 pay-rate outliers above about S$68.44; the highest is PL-5060 at S$173.47, and your Prompt_Log records a decision about it.

**Checkpoint and rejoin point**

Keep the workbook in work-c34. Topic 3 starts from the dashboard workbook in the Lab 5 folder.

**Troubleshooting**

| If this happens | Fix |
|---|---|
| Insert Python is not visible or =PY returns an error. | Python in Excel needs a supported plan, Windows/Mac/web and internet. Fallback: ask Copilot in Chat only for the same visual (it is returned as a static image) and use Insert > Charts > Statistical > Box and Whisker for a native box plot. |
| The =PY cell shows a Python object instead of a picture. | Select the cell and change its output (the icon beside the formula bar) from Python Object to Excel Value, or click the card icon to view the image. |
| #BUSY! or a slow calculation. | Python runs in the Microsoft cloud. Wait, or select Formulas > Reset (Python) and recalculate. |
| Levels appear in alphabetical order (Junior, Lead, Mid, Senior). | Ask Copilot to set the category order explicitly to Junior, Mid, Senior, Lead. |

**Challenge**

Ask Copilot, using Python, whether contracts that were Extended have a higher average Client_Rating than those that were not, with counts. Re-check the two averages with AVERAGEIFS.

**Reflection**

Which Python visual told you something a column chart of averages would have hidden?

> **Note:** The lab folder labs/lab-04-*/ holds the instruction PDF with sample prompts, sample-prompts.txt and the mock-data workbook (data/BrightPath_Contract_Placements.xlsx). Use only the supplied fictional BrightPath data or information you are authorised to process. Do not paste credentials, confidential business material or real employees' personal data into Copilot.

---


### Topic 02 Recap - Mapped to Learning Outcomes

Use this checkpoint to explain what you can now do and identify any lab evidence you still need to repair.

| Learning outcome | Evidence from this topic |
|---|---|
| LO2 - Plot charts with Copilot | Match the HR question to comparison, trend, composition, distribution or relationship; prompt for the exact fields, aggregation and chart type; refine titles, labels and axes; and tie every chart back to a PivotTable or source total. |

---


## Topic 03 — Use Copilot to Create Dashboards  (Labs 5-6)

HR and recruitment KPI design · building a dashboard with Plan and Edit modes · slicers, timelines and refresh · sharing people dashboards responsibly

**Key concepts**

- KPI definitions — Time to hire, offer acceptance, cost per hire, fill rate, headcount and attrition - defined first.
- Plan before build — Ask Copilot to plan the dashboard, review it, then let it build in Edit mode.
- Dashboard anatomy — KPI cards on top, 3-4 charts in the middle, slicers at the side, source note below.
- Interactivity — Slicers and timelines filter every PivotTable and chart connected to them.
- Refreshable by design — Tables + PivotTables + formulas refresh; pasted images and static answers do not.
- Responsible sharing — Aggregate, suppress small groups, control permissions and track changes.


### Designing HR and Recruitment KPIs

A KPI is a defined measure tied to a decision, with a formula, a period and a target. Common recruitment KPIs are applications, hires, time to hire, offer acceptance rate and cost per hire; workforce KPIs are headcount, new hires, attrition rate and average tenure. Microsoft's HR adoption scenarios cite cost per hire, onboarding time, retention and eNPS as measures Copilot can help improve.

Two people can compute 'offer acceptance rate' differently - including or excluding pending offers - and both will look right. Writing each definition before building lets you check Copilot's formula against it and keeps the dashboard consistent month after month.

**How it works**

- List the 4-6 decisions the dashboard supports and one KPI for each.
- Write the formula, period and exclusions for each KPI in a definitions note.
- Pick the dimensions users will filter by: month, sector, source, recruiter, department.

**Worked example**

- Offer acceptance rate = accepted / (accepted + declined); pending offers excluded.
- Cost per hire = total Sourcing_Cost / hires in the period.
- Attrition rate = exits in 12 months / average of opening and closing headcount.

**Decision guide**

| Use when | Avoid when |
|---|---|
| The same measures are reviewed regularly by the same audience. | A metric has no agreed definition or no one acts on it. |
| Each KPI has an owner who can act on a change. | Too many KPIs crowd the page and bury the important ones. |

**Practitioner quality lens**

- Defined: Formula, period and exclusions written down.
- Actionable: Every KPI maps to a decision and an owner.
- Focused: Four to six headline KPIs, not twenty.

**Authoritative references**

- https://adoption.microsoft.com/en-us/scenario-library/human-resources/implement-self-service-policy-questions-and-updates/
- https://techcommunity.microsoft.com/t5/s/gxcuf89792/attachments/gxcuf89792/HealthcareAndLifeSciencesBlog/2034.1/1/Copilot%20for%20M365%20HR%20-%20Starter%20Series.pdf
- https://learn.microsoft.com/en-us/training/modules/empower-workforce-copilot-hr/

---


### Building a Dashboard with Copilot Plan and Edit Modes

Copilot can create and edit charts, PivotTables and shapes and build dashboards, planning tools and templates from scratch, with editable links back to the source data. In Plan mode it outlines the sheets, ranges, formulas and objects it intends to create and can ask clarifying questions; in Allow editing mode it carries out that multi-step plan and checks the result.

A dashboard is a multi-step build - KPI formulas, several PivotTables, charts, slicers and layout - which is exactly where an agentic plan-execute-verify loop saves time. Reviewing the plan first catches a wrong KPI definition before it is copied into five objects.

**How it works**

- Plan: give the KPI definitions, the dimensions, the target sheet and the layout.
- Approve or correct the plan, then switch to Allow editing and let Copilot build it.
- Inspect each KPI card formula and reconcile it to its PivotTable or source total.

**Worked example**

- 'Plan a recruitment dashboard on the Dashboard sheet using tblApps: KPI cards for applications, hires, offer acceptance rate, average time to hire and cost per hire; charts for hires by month vs target, hires by sector and applications by source; slicers for Sector and Recruiter.'
- The plan is checked against the KPI definitions in About_This_Data before approval.

**Decision guide**

| Use when | Avoid when |
|---|---|
| The data is a clean table and the KPI definitions are written down. | The workbook is shared live with others and you have not saved a version. |
| You can review the plan and each created object afterwards. | You would accept the KPI values without checking a single one. |

**Practitioner quality lens**

- Plan review: KPI definitions in the plan match the written ones.
- Linked: Cards use formulas or GETPIVOTDATA, not typed numbers.
- Verified: Each KPI ties to an independent check.

**Authoritative references**

- https://support.microsoft.com/en-us/excel/copilot/get-started-with-copilot-in-excel
- https://www.microsoft.com/en-us/copilot/blog/2026/06/25/copilot-in-excel-built-for-the-era-of-frontier-finance/
- https://support.microsoft.com/en-us/excel/copilot/frequently-asked-questions-about-copilot-in-excel

---


### Interactivity, Refresh and Maintenance

Slicers and timelines are on-sheet filter buttons that can control several PivotTables and charts at once. Because the source is an Excel table, new rows are included automatically; Data > Refresh All updates the PivotTables and charts. Copilot can also store structure, naming and formula conventions as workbook rules, and reusable skills, so the dashboard is rebuilt the same way each month.

A dashboard that only works for one month's data is a report, not a dashboard. Testing refresh with new rows proves the formulas and PivotTables are linked correctly - and exposes any static images or typed numbers that Copilot or a user inserted.

**How it works**

- Ask Copilot to connect slicers to all PivotTables on the dashboard.
- Append test rows to the source table, then Data > Refresh All.
- Confirm every KPI and chart changed by the expected amount; fix any that did not.

**Worked example**

- Eight July 2026 new joiners are appended to tblEmployees.
- After Refresh All, active headcount rises by exactly 8 and every chart updates.
- A pasted chart image did not update, so it is replaced with a linked PivotChart.

**Decision guide**

| Use when | Avoid when |
|---|---|
| The dashboard will be reused with new data each period. | Slicers are left unconnected to some charts, giving inconsistent views. |
| Users need to explore by department, sector, recruiter or month. | Numbers are typed into KPI cards instead of calculated. |

**Practitioner quality lens**

- Connected: Each slicer filters every related object.
- Refresh-tested: Added rows change KPIs by the expected amount.
- Documented: Definitions, source and refresh date shown on the sheet.

**Authoritative references**

- https://support.microsoft.com/en-us/excel/copilot/get-started-with-copilot-in-excel
- https://www.microsoft.com/en-us/copilot/blog/2026/06/25/copilot-in-excel-built-for-the-era-of-frontier-finance/
- https://support.microsoft.com/en-us/excel/overview-of-pivottables-and-pivotcharts

---


### Sharing People Dashboards Responsibly

People data - salaries, performance, engagement and exits - is sensitive. A responsible HR dashboard shows aggregated results, suppresses groups too small to protect anonymity (for example fewer than 5 people), limits access through OneDrive or SharePoint permissions, and keeps a record of who changed what. Copilot's Show Changes pane attributes edits made by Copilot alongside collaborators.

Copilot does not give anyone access to data they could not already open - but anyone who can open the workbook sees everything in it, including hidden sheets. Saved changes are visible to co-authors. Responsible sharing protects employees and keeps the dashboard trusted.

**How it works**

- Remove or pseudonymise identifying columns and keep raw data in a restricted file if needed.
- Check that every slicer combination still shows groups of at least 5 people.
- Share with specific people, review Show Changes and keep version history.

**Worked example**

- Salary detail stays in the source file; the leadership copy shows only averages by level.
- Filtering IT + Director would show one person, so that slicer combination is removed.
- Copilot can then help draft the summary email or slides from the verified dashboard.

**Decision guide**

| Use when | Avoid when |
|---|---|
| The dashboard leaves the HR team or is used in a leadership meeting. | Individual-level records are visible to a broad audience. |
| The data includes pay, performance, health or exit details. | A small group's result could identify a person. |

**Practitioner quality lens**

- Minimise: Only the columns the audience needs.
- Protect: Small-group suppression and specific-people sharing.
- Audit: Show Changes and version history retained.

**Authoritative references**

- https://support.microsoft.com/en-us/excel/copilot/frequently-asked-questions-about-copilot-in-excel
- https://www.microsoft.com/en-us/microsoft-365-copilot/copilot-for-work
- https://www.microsoft.com/en-us/copilot/blog/2026/06/25/copilot-in-excel-built-for-the-era-of-frontier-finance/
- https://techcommunity.microsoft.com/t5/s/gxcuf89792/attachments/gxcuf89792/HealthcareAndLifeSciencesBlog/2034.1/1/Copilot%20for%20M365%20HR%20-%20Starter%20Series.pdf

---


### Lab 5 — Build a Recruitment KPI Dashboard with Copilot

Learning outcome: LO3: define recruitment KPIs, use Copilot Plan and Allow editing modes to build a dashboard of KPI cards, PivotTables, charts and slicers, and verify every KPI against an independent check.

Goal: Deliver a one-page H1 2026 recruitment dashboard whose five KPI cards, three charts and two slicers all agree with independently calculated checks.

BrightPath's Head of Recruitment wants one page that shows H1 2026 applications, hires, offer acceptance rate, time to hire and cost per hire, by month, sector, source and recruiter. You will write the checks first, ask Copilot to plan the dashboard, review the plan against the KPI definitions, let Copilot build it in Allow editing mode, and then prove each KPI and slicer works.

**What you'll build**

BrightPath_Recruitment_Dashboard_Data.xlsx with a Checks sheet and a Dashboard sheet holding five formula-driven KPI cards, charts for hires vs target by month, hires by sector and applications by source, and Sector and Recruiter slicers connected to every PivotTable   (Tools: Excel · Copilot (Plan, Allow editing) · PivotTables, PivotCharts, slicers · BrightPath_Recruitment_Dashboard_Data.xlsx.)

**Prerequisites**

- Topic 2 completed: you can build and verify PivotTables and charts with Copilot.
- This lab folder's data/BrightPath_Recruitment_Dashboard_Data.xlsx.

**Step-by-step**

1. Go to onedrive.com (or your SharePoint site), open or create the folder work-c34 and select Upload > Files. Upload BrightPath_Recruitment_Dashboard_Data.xlsx from this lab's data folder, then open it in Excel for the web or select Open in Desktop App. Confirm the title bar shows the cloud location and AutoSave is On, and that Formulas > Calculation Options is set to Automatic. Read the About_This_Data sheet before you prompt. Read the KPI definitions row on About_This_Data carefully - the dashboard must follow them exactly.

   ```bash
   Tables : tblApps (600 applications, Jan-Jun 2026) | tblTargets (monthly hire targets)
KPI definitions:
  Applications          = number of rows
  Hires                 = Furthest_Stage = "Hired"
  Offer acceptance rate = accepted / (accepted + declined)   (Pending excluded)
  Avg time to hire      = average Time_to_Hire_Days of hires
  Cost per hire         = total Sourcing_Cost / hires
   ```

2. Before Copilot builds anything, create a sheet named Checks and type these independent KPI formulas yourself. These are the numbers the dashboard must match.

   ```bash
   A1 Applications        B1 =ROWS(tblApps[Application_ID])
A2 Hires               B2 =COUNTIFS(tblApps[Furthest_Stage],"Hired")
A3 Offer acceptance    B3 =COUNTIFS(tblApps[Offer_Accepted],"Yes")/(COUNTIFS(tblApps[Offer_Accepted],"Yes")+COUNTIFS(tblApps[Offer_Accepted],"No"))
A4 Avg time to hire    B4 =AVERAGEIFS(tblApps[Time_to_Hire_Days],tblApps[Furthest_Stage],"Hired")
A5 Cost per hire       B5 =SUM(tblApps[Sourcing_Cost])/B2
   ```

3. Choose Plan. Paste the dashboard brief below. Read Copilot's plan: does it use the same KPI definitions (especially excluding Pending offers) and count hires by Hire_Date month? Reply with corrections until the plan matches. Do not approve yet.

   ```bash
   GOAL: Build a one-page H1 2026 recruitment KPI dashboard on the Dashboard sheet.
CONTEXT: For BrightPath's Head of Recruitment; monthly review meeting.
DATA: tblApps and tblTargets only. Use the KPI definitions on About_This_Data.
CONSTRAINTS: Do not change tblApps or tblTargets. KPI cards must be formulas, not typed numbers.
OUTPUT:
  Row 1: title "Recruitment Dashboard - Jan to Jun 2026" and "Data as at 2026-06-30".
  Row 3: five KPI cards - Applications, Hires, Offer acceptance rate, Avg time to hire (days),
         Cost per hire (S$).
  Below: (1) column chart of hires by Hire_Date month with a line for Hire_Target from tblTargets,
         (2) bar chart of hires by Sector sorted highest first,
         (3) bar chart of applications by Source.
  Slicers for Sector and Recruiter connected to all PivotTables. Put PivotTables on a sheet Dash_Data.
REVIEW: List each KPI formula you will use before building.
   ```

4. When the plan is right, switch to Allow editing and tell Copilot to carry it out. Watch the steps in the pane; you can select Stop if it goes off track. When it finishes, open Dash_Data and the Dashboard and review what was created.

   ```bash
   The plan is approved. Build the dashboard exactly as planned.
   ```

5. Compare each KPI card with Checks!B1:B5 (with no slicer selected). Click each card to see its formula. Fix any difference with a targeted follow-up prompt - the most common error is including Pending offers.

   ```bash
   FOLLOW-UP EXAMPLE:
The Offer acceptance rate card must exclude Offer_Accepted = "Pending". Change its formula to
accepted / (accepted + declined) and show it as a percentage with 1 decimal.
   ```

6. Test interactivity. Select Technology on the Sector slicer and check that every chart and KPI card changes. Compare with COUNTIFS for Technology. Then clear the slicer.

   ```bash
   CHECK (Technology):
=COUNTIFS(tblApps[Sector],"Technology")
=COUNTIFS(tblApps[Sector],"Technology",tblApps[Furthest_Stage],"Hired")

If a KPI card does not change, ask: "Make the KPI cards respond to the Sector and Recruiter slicers."
   ```

7. Polish with one prompt per change, then ask Copilot (Chat only) to review the dashboard for clarity. Log all prompts on Prompt_Log.

   ```bash
   Add a note under the title: "Source: tblApps; KPI definitions on About_This_Data".
Use one consistent colour scheme and align the charts in a grid.
CHAT ONLY: Review this dashboard for a Head of Recruitment. What is unclear or missing?
   ```


**Sample prompts**

| Purpose | Prompt |
|---|---|
| Plan | Plan a recruitment KPI dashboard on the Dashboard sheet from tblApps with KPI cards for applications, hires, offer acceptance rate, average time to hire and cost per hire. List the formulas before building. |
| KPI card | Create a KPI card that shows Cost per hire = total Sourcing_Cost / hires, formatted as S$ with 2 decimals. |
| Chart vs target | Create a column chart of hires by Hire_Date month with Hire_Target from tblTargets as a line. |
| Slicers | Add slicers for Sector and Recruiter and connect them to all PivotTables on Dash_Data. |
| Fix a KPI | The offer acceptance rate must exclude Pending offers. Update the formula and explain it. |
| Layout | Arrange the KPI cards in one row and the three charts in a grid below them with equal sizes. |
| Review | Review this dashboard for a Head of Recruitment. What is unclear, missing or potentially misleading? |
| Question | Which recruiter made the most hires in H1 2026, and what was their average time to hire? |

**Test it**

With no slicer selected the KPI cards match Checks: Applications 600, Hires 41, Offer acceptance 78.8% (41 accepted of 52 decided; 10 pending excluded), Avg time to hire 27.8 days, Cost per hire S$348.20. Hires by month: 2026-02 = 9, 2026-03 = 10, 2026-04 = 9, 2026-05 = 7, 2026-06 = 6 (none in 2026-01). Manufacturing has the most hires (11). With Sector = Technology, Applications = 124, Hires = 5 and Avg time to hire = 42.0 days. Siti Hamid made the most hires (13). No KPI card contains a typed number.

**Checkpoint and rejoin point**

Keep the workbook in work-c34. Lab 6 builds a second dashboard from the workforce workbook in its own folder.

**Troubleshooting**

| If this happens | Fix |
|---|---|
| Copilot builds before you approve the plan. | You were in Allow editing mode. Undo (Ctrl+Z) or restore from version history, switch to Plan and resubmit. |
| KPI cards do not change when a slicer is used. | Cards built with COUNTIFS on tblApps ignore slicers. Ask Copilot to base them on the PivotTables (GETPIVOTDATA) or accept that cards show the H1 total and label them so. |
| Hires appear in January. | The chart is counting by Applied_Month. Ask Copilot to group hires by Hire_Date month as the KPI definition says. |
| Copilot stops part-way or times out. | Ask it to continue from the last completed step, or split the build into KPI cards, then charts, then slicers. |

**Challenge**

Ask Copilot to add a sixth KPI card for 'Hires vs target %' (hires / sum of targets) and conditional formatting that turns it red below 90%. Verify it with your own formula.

**Reflection**

Which KPI definition was easiest to get subtly wrong, and how would you stop that happening next month?

> **Note:** The lab folder labs/lab-05-*/ holds the instruction PDF with sample prompts, sample-prompts.txt and the mock-data workbook (data/BrightPath_Recruitment_Dashboard_Data.xlsx). Use only the supplied fictional BrightPath data or information you are authorised to process. Do not paste credentials, confidential business material or real employees' personal data into Copilot.

---


### Lab 6 — Build an Interactive Workforce and Attrition Dashboard

Learning outcome: LO3: use Copilot to build a refreshable HR workforce dashboard with helper columns, KPI cards, charts and slicers, prove it refreshes with new rows, and check it for responsible sharing.

Goal: Show headcount, 12-month attrition, new hires and tenure by department and location on one interactive page that updates correctly when new employees are added.

BrightPath's HR Director needs a workforce dashboard for the July leadership meeting. You will ask Copilot to add period helper columns, type independent checks for headcount and attrition, have Copilot plan and build the dashboard with slicers, then append eight July new joiners and refresh to prove the dashboard is live. Finally you will check that no slicer combination exposes a group too small to stay anonymous.

**What you'll build**

BrightPath_Workforce_Headcount.xlsx with helper columns in tblEmployees, a Checks area, and an HR_Dashboard sheet showing six KPI cards, headcount and attrition-rate charts by department, hires vs exits by quarter, and Location and Job_Level slicers - refresh-tested with new rows   (Tools: Excel · Copilot (Plan, Allow editing) · PivotTables, slicers, Refresh All · BrightPath_Workforce_Headcount.xlsx.)

**Prerequisites**

- Lab 5 completed: you have built and verified a dashboard with Copilot.
- This lab folder's data/BrightPath_Workforce_Headcount.xlsx.

**Step-by-step**

1. Go to onedrive.com (or your SharePoint site), open or create the folder work-c34 and select Upload > Files. Upload BrightPath_Workforce_Headcount.xlsx from this lab's data folder, then open it in Excel for the web or select Open in Desktop App. Confirm the title bar shows the cloud location and AutoSave is On, and that Formulas > Calculation Options is set to Automatic. Read the About_This_Data sheet before you prompt. Do not touch the New_Joiners_Jul2026 sheet until Step 6.

   ```bash
   Table : tblEmployees (420 employees hired since 2019; Status Active or Exited)
Period: 2025-07-01 to 2026-06-30
Attrition rate = exits in period / ((active on 2025-07-01 + active on 2026-06-30) / 2)
   ```

2. Choose Allow editing and ask Copilot for three helper columns. Ask it to explain the Active_Start formula and test one employee who exited before 2025-07-01 (Active_Start must be FALSE).

   ```bash
   Add three columns to tblEmployees:
1) Active_Start = TRUE if Hire_Date <= 2025-07-01 and (Exit_Date is blank or Exit_Date > 2025-07-01).
2) Exit_In_Period = TRUE if Exit_Date is between 2025-07-01 and 2026-06-30 inclusive.
3) Tenure_Years = years from Hire_Date to Exit_Date, or to 2026-06-30 if Exit_Date is blank, 1 decimal.
   ```

3. Create a sheet named Checks and type these independent checks in columns A:B (label in A, formula in B). These are the numbers the dashboard must match.

   ```bash
   Active headcount     =COUNTIFS(tblEmployees[Status],"Active")
Headcount 2025-07-01 =COUNTIFS(tblEmployees[Active_Start],TRUE)
Exits in period      =COUNTIFS(tblEmployees[Exit_In_Period],TRUE)
Voluntary exits      =COUNTIFS(tblEmployees[Exit_In_Period],TRUE,tblEmployees[Exit_Type],"Voluntary")
Attrition rate       =<exits> / ((<headcount 2025-07-01> + <active headcount>) / 2)
Hires in period      =COUNTIFS(tblEmployees[Hire_Date],">="&DATE(2025,7,1),tblEmployees[Hire_Date],"<="&DATE(2026,6,30))
   ```

4. Choose Plan, paste the brief, review the plan against the definitions, then switch to Allow editing to build.

   ```bash
   Plan an HR workforce dashboard on the HR_Dashboard sheet from tblEmployees for the HR Director.
KPI cards (formulas): Active headcount, Exits (last 12 months), Attrition rate, Voluntary attrition rate,
Avg tenure of active staff (years), Hires (last 12 months). Use the definitions on About_This_Data.
Charts: active headcount by Department (bar, sorted); attrition rate by Department (bar); hires vs exits
by quarter for the last 12 months (clustered column); active headcount by Location (bar).
Slicers: Location and Job_Level, connected to all PivotTables. Title and "Data as at 2026-06-30".
Do not show Employee_ID, salary or individual records on the dashboard.
   ```

5. Compare the KPI cards with Checks and fix any difference with a follow-up prompt. Then select single Location and Job_Level slicer values and look for any department bar based on fewer than 5 people.

   ```bash
   FOLLOW-UP (if needed):
The attrition rate must use average headcount = (Active_Start count + active count) / 2.

CHAT ONLY:
Which combinations of Location, Job_Level and Department contain fewer than 5 active employees?
   ```

6. Refresh test: copy the 8 data rows (A2:L9) from New_Joiners_Jul2026 and paste them into the first empty row directly below tblEmployees so the table expands. Select Data > Refresh All. Every KPI and chart that counts active staff must update.

   ```bash
   Expected after refresh:
Active headcount : 237 -> 245
tblEmployees rows: 420 -> 428
New joiners by department: Recruitment +2, Sales & Client Services +2, IT +4
   ```

7. Ask Copilot (Chat only) to review the dashboard for privacy and clarity, act on its valid points, and log all prompts on Prompt_Log.

   ```bash
   Review HR_Dashboard before I share it with the leadership team. Does any element expose individual
employee data or a group smaller than 5? Is each KPI labelled with its period and definition?
   ```


**Sample prompts**

| Purpose | Prompt |
|---|---|
| Helper column | Add Active_Start = TRUE if Hire_Date <= 2025-07-01 and Exit_Date is blank or after 2025-07-01. |
| Helper column | Add Tenure_Years from Hire_Date to Exit_Date, or to 2026-06-30 if Exit_Date is blank, rounded to 1 decimal. |
| Plan | Plan an HR workforce dashboard with KPI cards for active headcount, exits, attrition rate, voluntary attrition, average tenure and hires. |
| Chart | Create a bar chart of 12-month attrition rate by Department, sorted highest first, with data labels as percentages. |
| Slicers | Add slicers for Location and Job_Level and connect them to every PivotTable on the dashboard. |
| Refresh | I added rows to tblEmployees. Refresh all PivotTables and tell me which KPI values changed. |
| Privacy | Which slicer combinations show a group of fewer than 5 employees? Suggest how to suppress them. |
| Explain | Explain how the Attrition rate card is calculated, step by step. |

**Test it**

Before the refresh: Active headcount 237, headcount on 2025-07-01 210, exits in period 29 (25 voluntary), attrition rate 13.0% and voluntary attrition 11.2%, hires in period 56, average tenure of active staff about 3.1 years. Recruitment has the highest attrition rate (21.8%, 12 exits) and the largest active headcount (60). After appending the July joiners and Refresh All, Active headcount shows 245 and tblEmployees has 428 rows. The dashboard shows no Employee_ID, salary or individual record.

**Checkpoint and rejoin point**

Keep the workbook in work-c34. Topic 4 starts from the attrition-analysis workbook in the Lab 7 folder.

**Troubleshooting**

| If this happens | Fix |
|---|---|
| Pasted rows did not join the table and nothing changed on refresh. | Paste into the row directly under the last table row, or use Table Design > Resize Table to include them. |
| Attrition rate is far too high (e.g. 29 / 237). | The card divides by closing headcount only. Ask Copilot to use the average of opening and closing headcount. |
| Helper columns show #VALUE! for blank Exit_Date. | Ask Copilot to treat blank Exit_Date explicitly, e.g. OR([@Exit_Date]="",[@Exit_Date]>DATE(2025,7,1)). |
| A chart did not update after Refresh All. | It is a static picture or a range copy. Ask Copilot to replace it with a PivotChart linked to tblEmployees. |

**Challenge**

Ask Copilot to add a Timeline control on Hire_Date and a chart of average Engagement_Score by Department for active staff, then check one department's average with AVERAGEIFS.

**Reflection**

What would happen to trust in this dashboard if one leader could filter it down to a single named team member?

> **Note:** The lab folder labs/lab-06-*/ holds the instruction PDF with sample prompts, sample-prompts.txt and the mock-data workbook (data/BrightPath_Workforce_Headcount.xlsx). Use only the supplied fictional BrightPath data or information you are authorised to process. Do not paste credentials, confidential business material or real employees' personal data into Copilot.

---


### Topic 03 Recap - Mapped to Learning Outcomes

Use this checkpoint to explain what you can now do and identify any lab evidence you still need to repair.

| Learning outcome | Evidence from this topic |
|---|---|
| LO3 - Create dashboards with Copilot | Define each KPI before building; ask Copilot to plan, then build KPI cards, PivotTables, charts and slicers; prove the dashboard refreshes when rows are added; and share only aggregated, permissioned views. |

---


## Topic 04 — Use Copilot to Draw Insights and Recommendations  (Labs 7-8)

Surfacing trends, outliers and drivers · advanced analysis mode with Python · themes and sentiment in feedback · from verified insight to owned recommendation

**Key concepts**

- Insight ladder — Observation, evidence, implication, recommendation, owner and KPI.
- What's interesting? — A good opening prompt - followed by specific, checkable follow-ups.
- Advanced analysis — Python on a new sheet for deeper, editable, refreshable analysis.
- Text analysis — Themes and sentiment from comments as a new column, spot-checked by a human.
- Observed vs hypothesis — A pattern is not a cause; label explanations as hypotheses to test.
- Human decision — People decisions stay with people; Copilot informs, it does not decide.


### Surfacing Trends, Outliers and Drivers

Ask Copilot about your data in plain language and it surfaces summaries, trends and outliers as text, charts or PivotTables. Microsoft's examples include 'What's interesting about this data?', 'Are there any outliers or anomalies in this data?' and 'What are the mean and median of revenue by region?'. A driver analysis compares an outcome, such as attrition, across groups or measures.

Open questions are a good start, but they can return selective or causal-sounding stories. Specific follow-ups with a baseline and denominator - 'attrition rate for engagement below 3 versus 3 and above, with counts' - produce evidence you can check. Small groups (an IT team of 25) can show extreme rates by chance.

**How it works**

- Start broad ('What's interesting?'), then ask for specific comparisons with counts.
- Always ask for the denominator and group size behind each rate.
- Verify the headline numbers with a PivotTable or COUNTIFS before recording them.

**Worked example**

- Overall 12-month attrition in the study group: 56 of 400 = 14.0%.
- Engagement below 3: 29.6% left (n = 81) vs 10.0% (n = 319) at 3 and above.
- IT shows 28.0% - but only 25 people, so it is flagged as a small-sample signal.

**Decision guide**

| Use when | Avoid when |
|---|---|
| You need candidate patterns to investigate further. | A short period or small group is used to claim a lasting trend. |
| Each claimed pattern can be traced to rows and a calculation. | A coincident pattern is presented as the proven cause. |

**Practitioner quality lens**

- Baseline: Show the comparison group, denominator and n.
- Method: State how a trend or outlier was defined.
- Language: OBSERVED for evidence, HYPOTHESIS for explanations.

**Authoritative references**

- https://support.microsoft.com/en-us/excel/copilot/data-insights-with-copilot-in-excel
- https://support.microsoft.com/en-us/excel/get-direct-answers-to-your-data-analysis-questions
- https://learn.microsoft.com/en-us/training/modules/empower-workforce-copilot-hr/

---


### Advanced Analysis Mode with Python

After a direct answer, choose 'Get deeper results using advanced analysis mode' or ask Copilot to 'Enter advanced analysis mode'. Copilot opens a new sheet and runs Python for deeper, customisable analysis - group comparisons, correlations, clustering or regression - which you can modify and insert as a refreshable Python cell. Select 'Stop advanced analysis' to return.

Python widens what you can ask - 'Which factors are most associated with leaving?' - but correlation is not causation, and the model's choices (which rows, how blanks are handled, which test) shape the answer. Microsoft advises reviewing the commented code and asking Copilot to summarise it to check its logic.

**How it works**

- Ask a precise question naming the table, outcome, candidate factors and output.
- Read the code: filters, grouping, missing values and the statistic used.
- Re-compute one headline value in Excel and record the method and limitations.

**Worked example**

- 'Compare employees who left with those who stayed on Overtime_Hours_Month, Engagement_Score, Years_Since_Promotion and Compa_Ratio. Show means and counts.'
- Leavers average 15.9 overtime hours a month vs 10.6 for stayers - checked with AVERAGEIFS.

**Decision guide**

| Use when | Avoid when |
|---|---|
| The question needs statistics beyond a PivotTable. | The output would be used to predict or decide about a named individual. |
| A reviewer can read the code and reproduce the headline numbers. | Nobody can explain the method used. |

**Practitioner quality lens**

- Inputs: Right table, rows, columns and blank handling.
- Method: Statistic fits the question; association not causation.
- Reproducible: Headline values re-computed with native formulas.

**Authoritative references**

- https://support.microsoft.com/en-us/excel/get-direct-answers-to-your-data-analysis-questions
- https://support.microsoft.com/en-us/excel/python/introduction-to-python-in-excel
- https://support.microsoft.com/en-us/office/get-advanced-data-analysis-using-copilot-in-excel-app-skills-364e4ae9-9343-4d56-952a-5f62b0f70db6

---


### Analysing Survey Comments: Themes and Sentiment

For text such as survey results, reviews or feedback, Copilot can summarise comments, identify themes and sentiment, and add the results as a new column. In the grid, the COPILOT function - for example =COPILOT("Classify the theme of this comment into one of", D2, F2:F8) - classifies text row by row where your licence and channel support it. Hover over a superscript citation to see which rows Copilot used.

Hundreds of comments can be themed in minutes, which makes employee voice usable in an engagement review. But AI classification can be wrong, inconsistent between runs, and is not suited to exact calculation - Microsoft recommends native functions such as SUM, AVERAGE and IF for numbers. A fixed theme list and a human spot-check keep it honest.

**How it works**

- Give Copilot a fixed list of themes and sentiments to choose from, plus 'Unclear'.
- Add the classification as new columns; keep the original comment unchanged.
- Spot-check at least 10 rows; if more than 1 in 10 is wrong, refine the prompt and rerun.

**Worked example**

- Themes: Workload, Career Growth, Manager Support, Pay & Benefits, Recognition, Tools & Systems, Flexibility.
- Workload is the most frequent theme in BrightPath's June pulse survey.
- eNPS itself is computed with COUNTIFS, not with the COPILOT function.

**Decision guide**

| Use when | Avoid when |
|---|---|
| Comments are many and the theme list is agreed in advance. | Comments could identify the respondent and would be shown verbatim. |
| A person reviews a sample and owns the final counts. | The AI label would be used as an exact metric without review. |

**Practitioner quality lens**

- Taxonomy: Fixed theme list plus Unclear.
- Spot-check: Sample reviewed and error rate recorded.
- Anonymity: Do not quote comments that reveal a person.

**Authoritative references**

- https://support.microsoft.com/en-us/excel/copilot/data-insights-with-copilot-in-excel
- https://support.microsoft.com/en-us/excel/copilot-function
- https://adoption.microsoft.com/en-us/scenario-library/human-resources/implement-self-service-policy-questions-and-updates/

---


### From Verified Insight to Owned Recommendation

A recommendation turns a verified finding into an action: Finding (OBSERVED) - Evidence (cells and numbers) - Implication - Recommendation - Owner - Priority - Success KPI and review date. Copilot can draft each part, an executive summary, and - through Copilot in Word, PowerPoint or Outlook - the memo, slides or email that communicates it.

AI output can contain mistakes; Microsoft advises reviewing, editing and verifying everything before relying on it, and not relying on it alone for high-stakes decisions. Recommendations about people affect livelihoods: they should target teams, processes and policies, cite evidence, state limitations and be approved by an accountable person.

**How it works**

- Select 3-5 verified findings and write each as an OBSERVED statement with numbers.
- Ask Copilot to propose actions, owners and KPIs; edit them to fit your organisation.
- Draft the summary, check every number against the workbook, then decide Share, Revise or Hold.

**Worked example**

- Finding: Workload balance is the lowest-scoring item (2.71 of 5) and the top comment theme.
- Recommendation: review recruiter desk targets and add peak-season support - owner: Head of Recruitment.
- KPI: raise Q3_Workload_Balance to 3.2 and eNPS above 0 by the December pulse.

**Decision guide**

| Use when | Avoid when |
|---|---|
| Findings are verified and an accountable owner exists. | An action targets an identifiable individual based on model output. |
| The recommendation can be measured and reviewed. | A hypothesis is presented as a proven cause. |

**Practitioner quality lens**

- Evidence-backed: Every recommendation cites a verified finding.
- Owned: Named owner, priority and review date.
- Measurable: A KPI and target that will show if it worked.

**Authoritative references**

- https://support.microsoft.com/en-us/excel/copilot/frequently-asked-questions-about-copilot-in-excel
- https://learn.microsoft.com/en-us/training/modules/empower-workforce-copilot-hr/
- https://adoption.microsoft.com/en-us/scenario-library/human-resources/implement-self-service-policy-questions-and-updates/
- https://www.randgroup.com/insights/copilot/enhancing-productivity-with-ai-in-microsoft-365/

---


### Lab 7 — Surface Attrition Insights with Copilot and Advanced Analysis

Learning outcome: LO4: use Copilot direct answers and Python-based advanced analysis to surface attrition patterns, verify them with native formulas and record each as OBSERVED evidence or a HYPOTHESIS with its limitation.

Goal: Identify the work-pattern factors most associated with leaving BrightPath, with verified rates, group sizes and honest limitations in an Insight_Log.

BrightPath lost 14% of a 400-person study group in 12 months. You will start with an open 'What's interesting?' prompt, set the baseline attrition rate, ask Copilot for specific group comparisons with counts, enter advanced analysis mode to compare leavers and stayers with Python, re-compute the headline numbers with COUNTIFS and AVERAGEIFS, and write four labelled insights - including why a high rate in a small team is only a signal.

**What you'll build**

BrightPath_Attrition_Analysis.xlsx with verified attrition rates by department and factor, an advanced-analysis sheet with reviewed Python output, and an Insight_Log of four OBSERVED or HYPOTHESIS insights with evidence and limitations   (Tools: Excel · Copilot (Chat only, advanced analysis mode) · Python in Excel · BrightPath_Attrition_Analysis.xlsx.)

**Prerequisites**

- Labs 1-6 completed, or equivalent skill with Copilot prompts, PivotTables and COUNTIFS.
- Python in Excel available for advanced analysis (fallback in Troubleshooting).
- This lab folder's data/BrightPath_Attrition_Analysis.xlsx.

**Step-by-step**

1. Go to onedrive.com (or your SharePoint site), open or create the folder work-c34 and select Upload > Files. Upload BrightPath_Attrition_Analysis.xlsx from this lab's data folder, then open it in Excel for the web or select Open in Desktop App. Confirm the title bar shows the cloud location and AutoSave is On, and that Formulas > Calculation Options is set to Automatic. Read the About_This_Data sheet before you prompt. Note the warning on About_This_Data: these are team-level signals, never a basis for decisions about an individual.

   ```bash
   Table   : tblAttrition (400 employees active on 2025-07-01)
Outcome : Left_Company = Yes / No (left within 12 months)
Template: Insight_Log sheet
   ```

2. Choose Chat only and ask the open question. Copy Copilot's candidate patterns into Prompt_Log - do not record them as insights yet; each must be verified first.

   ```bash
   What's interesting about tblAttrition in relation to Left_Company? Give up to five patterns, each with
the numbers and group sizes behind it. Do not change the workbook.
   ```

3. Set the baseline: ask for the overall rate, then check it with formulas in two empty cells.

   ```bash
   What is the overall attrition rate in tblAttrition (Left_Company = "Yes" divided by all employees)?

CHECK:
=COUNTIFS(tblAttrition[Left_Company],"Yes")
=COUNTIFS(tblAttrition[Left_Company],"Yes")/ROWS(tblAttrition[Employee_ID])
   ```

4. Ask for specific comparisons with counts. Verify at least the engagement and overtime splits with COUNTIFS, and note which departments have fewer than 30 people.

   ```bash
   Using tblAttrition, show attrition rate AND number of employees for:
1) each Department;
2) Engagement_Score below 3 vs 3 and above;
3) Overtime_Hours_Month above 20 vs 20 or below;
4) Years_Since_Promotion 3 or more vs below 3;
5) Compa_Ratio below 0.9 vs 0.9 and above;
6) each Work_Mode.

CHECK (engagement < 3):
=COUNTIFS(tblAttrition[Engagement_Score],"<3",tblAttrition[Left_Company],"Yes")/COUNTIFS(tblAttrition[Engagement_Score],"<3")
   ```

5. Go deeper: ask Copilot to enter advanced analysis mode (or select 'Get deeper results using advanced analysis mode'). Copilot opens a new sheet with Python. Read the code: which rows, which columns, how blanks are handled. Verify two of the means with AVERAGEIFS, then select Stop advanced analysis.

   ```bash
   Enter advanced analysis mode.
Compare employees who left with those who stayed on Overtime_Hours_Month, Engagement_Score,
Years_Since_Promotion, Compa_Ratio and Training_Hours. Show the mean for each group, the group sizes,
and the correlation between Overtime_Hours_Month and Engagement_Score. Explain the method.

CHECK:
=AVERAGEIFS(tblAttrition[Overtime_Hours_Month],tblAttrition[Left_Company],"Yes")
=AVERAGEIFS(tblAttrition[Engagement_Score],tblAttrition[Left_Company],"Yes")
   ```

6. Ask Copilot to help label your findings, then edit them yourself into four rows of the Insight_Log sheet. Each row needs the verified numbers, a label, a possible explanation to test and a limitation.

   ```bash
   Here are my verified findings: <paste your checked rates and means>.
For each, write one OBSERVED sentence with the numbers, one HYPOTHESIS that could explain it, and
one limitation (sample size, correlation is not causation, overlapping factors). Keep each under 30 words.
   ```

7. Ask Copilot (Chat only) to challenge your Insight_Log, then log all prompts on Prompt_Log.

   ```bash
   Review my Insight_Log. Which statement overstates the evidence or implies causation? How would
you reword it?
   ```


**Sample prompts**

| Purpose | Prompt |
|---|---|
| Open question | What's interesting about this data in relation to Left_Company? Give five patterns with numbers and group sizes. |
| Baseline | What is the overall attrition rate in tblAttrition? Show the numerator and denominator. |
| Comparison | Show attrition rate and headcount for Engagement_Score below 3 vs 3 and above. |
| By group | Show attrition rate and number of employees by Department, sorted highest first. |
| Advanced analysis | Enter advanced analysis mode. Compare leavers and stayers on overtime, engagement, promotion, compa-ratio and training. |
| Correlation | What is the correlation between Overtime_Hours_Month and Engagement_Score, and what does it mean for interpreting the attrition patterns? |
| Outliers | Are there any managers (Manager_ID) whose teams have unusually high attrition? Only include teams with at least 10 people. |
| Challenge | Which of my insights could be a coincidence or caused by another factor? Explain. |

**Test it**

Overall attrition = 56 of 400 = 14.0%. Engagement below 3: 29.6% (n = 81) vs 10.0% (n = 319). Overtime above 20 hours: 34.0% (n = 47) vs 11.3% (n = 353). Promotion 3+ years ago: 21.3% vs 13.0%. On-site staff: 19.3%. IT shows the highest department rate (28.0%) but has only 25 people, so it is recorded as a small-sample signal. Leavers average 15.9 overtime hours and engagement 3.10, vs 10.6 and 3.54 for stayers. Overtime and engagement correlate at about -0.25, so the two factors overlap. Insight_Log has four rows, each labelled and with a limitation.

**Checkpoint and rejoin point**

Keep the workbook in work-c34 and keep your Insight_Log open - Lab 8 turns verified insights like these into recommendations.

**Troubleshooting**

| If this happens | Fix |
|---|---|
| Advanced analysis mode is not offered. | It needs Python in Excel. Stay in Chat only: direct answers also use Python and can show code. Or build the comparisons with a PivotTable (Left_Company in columns, Count of Employee_ID, % of row total). |
| Copilot gives a rate without a group size. | Follow up: 'Show the number of employees in each group next to each rate.' Never record a rate without n. |
| Copilot says a factor 'causes' attrition. | Ask it to restate as an association, and record any causal idea as a HYPOTHESIS to test (e.g. exit interviews). |
| Your COUNTIFS rate differs slightly from Copilot's. | Check the boundary (< 3 vs <= 3, > 20 vs >= 20) and whether blanks were excluded. Use your formula. |

**Challenge**

In advanced analysis mode, ask Copilot to fit a simple logistic regression of Left_Company on the five factors and explain the output in plain English. Record why this still does not prove cause.

**Reflection**

Which insight would you be most careful presenting to managers, and what extra evidence would you want first?

> **Note:** The lab folder labs/lab-07-*/ holds the instruction PDF with sample prompts, sample-prompts.txt and the mock-data workbook (data/BrightPath_Attrition_Analysis.xlsx). Use only the supplied fictional BrightPath data or information you are authorised to process. Do not paste credentials, confidential business material or real employees' personal data into Copilot.

---


### Lab 8 — Turn Engagement Survey Feedback into Recommendations

Learning outcome: LO4: use Copilot to quantify survey scores and eNPS, classify comment themes and sentiment with a human spot-check, and draft owned, measurable recommendations and an executive summary from verified findings.

Goal: Produce three evidence-backed recommendations and a 150-word executive summary from BrightPath's June engagement pulse survey, with every number verified.

BrightPath ran an anonymous pulse survey of 160 staff. You will have Copilot calculate eNPS and question averages and verify them with COUNTIFS and AVERAGE, compare departments while respecting small groups, add Theme and Sentiment columns from a fixed list and spot-check them, then use the Recommendations template and Copilot to turn verified findings into owned actions and an executive summary you approve.

**What you'll build**

BrightPath_Engagement_Survey.xlsx with verified eNPS and question averages, Theme and Sentiment columns with a recorded spot-check, a Recommendations sheet with three owned, measurable actions, and an Executive_Summary   (Tools: Excel · Copilot (Chat only, Allow editing) · COPILOT function (where available) · BrightPath_Engagement_Survey.xlsx.)

**Prerequisites**

- Lab 7 completed: you can verify an insight and label it OBSERVED or HYPOTHESIS.
- This lab folder's data/BrightPath_Engagement_Survey.xlsx.

**Step-by-step**

1. Go to onedrive.com (or your SharePoint site), open or create the folder work-c34 and select Upload > Files. Upload BrightPath_Engagement_Survey.xlsx from this lab's data folder, then open it in Excel for the web or select Open in Desktop App. Confirm the title bar shows the cloud location and AutoSave is On, and that Formulas > Calculation Options is set to Automatic. Read the About_This_Data sheet before you prompt. Responses are anonymous: do not try to identify anyone, and do not report a group with fewer than 5 responses.

   ```bash
   Table : tblSurvey (160 responses; Q1-Q6 scored 1-5; eNPS_Score 0-10; Comment)
eNPS  = % promoters (9-10) - % detractors (0-6)
Templates: Recommendations | Executive_Summary
   ```

2. Choose Chat only and ask for eNPS and the question averages. Then verify with the formulas below in an empty area of Executive_Summary (or a new Checks sheet).

   ```bash
   Using tblSurvey, calculate the number of promoters (9-10), passives (7-8) and detractors (0-6) on
eNPS_Score, and the eNPS. Then give the average of Q1 to Q6, lowest to highest.

CHECK:
=COUNTIFS(tblSurvey[eNPS_Score],">=9")
=COUNTIFS(tblSurvey[eNPS_Score],"<=6")
=(COUNTIFS(tblSurvey[eNPS_Score],">=9")-COUNTIFS(tblSurvey[eNPS_Score],"<=6"))/ROWS(tblSurvey[Response_ID])*100
=AVERAGE(tblSurvey[Q3_Workload_Balance])
   ```

3. Ask for eNPS by department with response counts. Mark any department with fewer than 10 responses as 'interpret with caution' and never report one with fewer than 5.

   ```bash
   Show eNPS and number of responses for each Department in tblSurvey, sorted lowest eNPS first.
   ```

4. Choose Allow editing and ask Copilot to classify comments into a fixed list. If your Excel has the COPILOT function you may use the formula option instead. Keep the original Comment column unchanged.

   ```bash
   OPTION A (Copilot pane):
Add two columns to tblSurvey. Theme: classify each Comment into exactly one of Workload, Career Growth,
Manager Support, Pay & Benefits, Recognition, Tools & Systems, Flexibility, or Unclear. Sentiment:
Positive, Negative or Mixed. Leave both blank where Comment is blank. Do not change Comment.

OPTION B (COPILOT function, where available), in the first Theme cell:
=COPILOT("Classify this staff comment into exactly one theme from the list. Reply with the theme only.",
         [@Comment], "List:", "Workload, Career Growth, Manager Support, Pay & Benefits, Recognition,
         Tools & Systems, Flexibility, Unclear")
   ```

5. Spot-check: read 10 classified comments yourself and count disagreements. If more than 1 in 10 is wrong, refine the prompt (for example, add an example per theme) and rerun. Then ask for theme counts and record the error rate on Prompt_Log.

   ```bash
   Count comments by Theme and by Sentiment in tblSurvey, highest first. For the top theme, summarise
what staff are saying in three bullet points without quoting anything that could identify a person.

CHECK: =COUNTIF(tblSurvey[Theme],"Workload")
   ```

6. Fill the Recommendations sheet: ask Copilot to draft three rows from your verified findings, then edit the owners, targets and dates so they are realistic. Every Evidence cell must cite numbers you verified.

   ```bash
   Using only these verified findings: <paste eNPS, lowest question, top theme, department results>,
draft three recommendations in the Recommendations sheet columns: Finding (OBSERVED), Evidence,
Recommendation (a team, process or policy action - not about any individual), Owner (role),
Priority, Success KPI and target, Review date (by Dec 2026).
   ```

7. Ask Copilot to draft the executive summary on Executive_Summary. Check every number against your checks, edit the wording, and end with your decision: Share, Revise or Hold. Log all prompts. Optionally ask Copilot in Word or PowerPoint to turn the summary into a memo or two slides.

   ```bash
   Draft a 150-word executive summary for BrightPath's leadership from the Recommendations sheet and the
verified results. Include eNPS, the lowest-scoring question, the top comment theme and the three
recommendations with owners. Use plain English, no jargon, and do not claim causes.
   ```


**Sample prompts**

| Purpose | Prompt |
|---|---|
| Score | Calculate eNPS from eNPS_Score in tblSurvey: promoters 9-10, passives 7-8, detractors 0-6. Show the counts. |
| Averages | Give the average of each question Q1 to Q6, lowest to highest, with the number of responses. |
| Groups | Show eNPS and number of responses by Department, sorted lowest first. |
| Themes | Classify each Comment into exactly one theme from my list (or Unclear) in a new Theme column. Do not change Comment. |
| Sentiment | Add a Sentiment column: Positive, Negative or Mixed for each non-blank Comment. |
| COPILOT function | =COPILOT("Classify the sentiment of this comment as Positive, Negative or Mixed", [@Comment]) |
| Summarise text | Summarise what staff say about Workload in three bullet points without quoting anything identifying. |
| Recommend | From these verified findings, draft three recommendations with owner, priority, success KPI and review date. |
| Executive summary | Draft a 150-word executive summary for leadership. Do not claim causes; cite the numbers. |

**Test it**

Promoters 39, passives 61, detractors 60 of 160; eNPS = -13.1. Lowest question Workload Balance (2.71); highest Manager Support (3.76). Human Resources has the lowest department eNPS (-36.4) from only 11 responses, so it is flagged 'interpret with caution'. 134 responses have a comment; after classification Workload is the most frequent theme (about 39 comments - AI labels can vary by a few, which is why the spot-check error rate is recorded). The Recommendations sheet has three rows, each with verified evidence, an owner role, a KPI target and a review date, and Executive_Summary ends with a Share, Revise or Hold decision.

**Checkpoint and rejoin point**

This is the final lab. Keep all eight workbooks in work-c34 as your personal reference set.

**Troubleshooting**

| If this happens | Fix |
|---|---|
| Copilot invents new themes outside the list. | Repeat the exact list, add 'reply with one theme from the list only', and ask it to use Unclear when unsure. |
| The COPILOT function returns #NAME? or is missing. | It is not available on every licence or channel yet. Use Option A in the Copilot pane instead. |
| Copilot's eNPS uses 7-10 as promoters or averages the scores. | Restate the rule: promoters 9-10, detractors 0-6, eNPS = % promoters - % detractors. Trust your COUNTIFS. |
| The summary quotes a comment that could identify someone. | Remove it and ask Copilot to paraphrase themes without quotes. Anonymity comes before vividness. |

**Challenge**

Combine Lab 7 and Lab 8: ask Copilot whether the departments with the highest attrition also have the lowest Workload Balance scores, then write one recommendation that uses both sources - labelled as a HYPOTHESIS.

**Reflection**

Which recommendation would you be confident defending to the HR Director, and which number makes it defensible?

> **Note:** The lab folder labs/lab-08-*/ holds the instruction PDF with sample prompts, sample-prompts.txt and the mock-data workbook (data/BrightPath_Engagement_Survey.xlsx). Use only the supplied fictional BrightPath data or information you are authorised to process. Do not paste credentials, confidential business material or real employees' personal data into Copilot.

---


### Topic 04 Recap - Mapped to Learning Outcomes

Use this checkpoint to explain what you can now do and identify any lab evidence you still need to repair.

| Learning outcome | Evidence from this topic |
|---|---|
| LO4 - Draw insights and recommendations | Ask for descriptive evidence before explanations; review Python logic in advanced analysis; classify comments and spot-check them; label OBSERVED versus HYPOTHESIS; and convert verified findings into owned, measurable recommendations. |

---


## Wrap-Up - The Complete C34 Workflow

Across eight labs you moved from questions to charts, dashboards and recommendations - always with a verification step between Copilot's output and a people decision.

**Analyse and plot**

- Set up Copilot correctly: eligible licence, cloud .xlsx, AutoSave and the right mode.
- Prepare named tables, prompt with named columns, and verify formulas and lookups.
- Choose the chart from the question, build it from a PivotTable and refine it by follow-up.

**Dashboard**

- Define KPIs first, plan with Copilot, then build KPI cards, PivotTables, charts and slicers.
- Prove refresh with new rows and remove any static or typed values.
- Share aggregated, permissioned views and keep change history.

**Insight and recommend**

- Ask broad, then specific; always see the denominator and group size.
- Review Python logic in advanced analysis and re-compute headline values.
- Turn verified findings into owned, measurable recommendations and a reviewed summary.

---


## Next Steps

- Re-run all eight labs from the original data workbooks without looking at your finished files.
- Adapt the sample prompts and the Prompt_Log to one authorised HR or recruitment workbook in your role.
- Turn your best dashboard into a monthly template, and use a Copilot skill or workbook rules to keep it consistent.
- Check Microsoft Support regularly: Copilot in Excel's modes, labels and capabilities change frequently.


## Glossary

- **Advanced analysis mode** — Copilot mode that opens a new sheet and runs editable Python for deeper analysis.
- **Allow editing (Edit) mode** — The Copilot mode that changes the workbook directly.
- **Attrition rate** — Exits in a period divided by average headcount in that period.
- **Chat only mode** — The Copilot mode that answers and analyses without changing the workbook.
- **COPILOT function** — An Excel function that sends a prompt plus cell context to AI and returns the result to the grid.
- **Cost per hire** — Total recruiting cost in a period divided by hires in that period.
- **eNPS** — Employee Net Promoter Score: % promoters (9-10) minus % detractors (0-6).
- **G-C-D-C-O-R** — Goal, Context, Data, Constraints, Output and Review - the C34 prompt framework.
- **Grain** — What one row or one aggregated cell represents.
- **KPI card** — A dashboard cell or shape that shows one headline measure, calculated by formula.
- **Offer acceptance rate** — Accepted offers divided by offers with a decision (accepted + declined).
- **Plan mode** — The Copilot mode that proposes steps for review before changing the workbook.
- **Python in Excel** — Python code run in the Microsoft cloud from =PY cells, returning results to the grid.
- **Slicer** — An on-sheet filter button that can control several PivotTables and charts.
- **Static output** — An inserted table or image that does not update when the source data changes.
- **Structured reference** — A formula reference that uses table and column names, e.g. tblApps[Hires].
- **Time to hire** — Days from application to accepted offer, averaged over hires.
