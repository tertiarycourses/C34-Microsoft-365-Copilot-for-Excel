# Microsoft 365 Copilot for Excel — Learner Guide

**Course Code:** C34  |  **Conducted by:** Tertiary Infotech Academy Pte Ltd (UEN 201200696W)  |  **Version v3.0 · 11 October 2026**

## Contents

- [Introduction](#introduction)
- [Course Learning Outcomes](#course-learning-outcomes)
- [Before You Start — Preparation](#before-you-start--preparation)
- [Topic 01 — Excel Fundamentals with Copilot  (Labs 1-3)](#topic-01--excel-fundamentals-with-copilot--labs-1-3)
  - [What Is Data Prep?](#what-is-data-prep)
  - [Why Data Quality Matters](#why-data-quality-matters)
  - [Chat Only, Plan and Allow Editing](#chat-only-plan-and-allow-editing)
  - [Setup, Data Sources and the Model Picker](#setup-data-sources-and-the-model-picker)
  - [Cleaning Data with Copilot](#cleaning-data-with-copilot)
  - [Lab 1 · Data Prep & Data Quality — Explore and Question Recruitment Pipeline Data](#lab-1--data-prep--data-quality--explore-and-question-recruitment-pipeline-data)
  - [What Is a Formula in Excel?](#what-is-a-formula-in-excel)
  - [Table References: What Does @ Mean?](#table-references-what-does--mean)
  - [Reading Excel References: !, :, $ and [@[...]]](#reading-excel-references----and-)
  - [Reading a Formula: IF, AND and @](#reading-a-formula-if-and-and-)
  - [What Is XLOOKUP?](#what-is-xlookup)
  - [Writing Formulas and XLOOKUPs with Copilot](#writing-formulas-and-xlookups-with-copilot)
  - [Lab 2 · Formulas & XLOOKUP — Generate and Verify Formula Columns and XLOOKUPs for Job Requisitions](#lab-2--formulas--xlookup--generate-and-verify-formula-columns-and-xlookups-for-job-requisitions)
  - [What Are PivotTables and PivotCharts?](#what-are-pivottables-and-pivotcharts)
  - [PivotTables and PivotCharts from a Prompt](#pivottables-and-pivotcharts-from-a-prompt)
  - [Choosing and Refining Charts](#choosing-and-refining-charts)
  - [Lab 3 · PivotTables, PivotCharts & Visuals — Build PivotTables, PivotCharts and Recruitment Charts with Copilot](#lab-3--pivottables-pivotcharts--visuals--build-pivottables-pivotcharts-and-recruitment-charts-with-copilot)
  - [Topic 01 Recap - Mapped to Learning Outcomes](#topic-01-recap---mapped-to-learning-outcomes)
- [Topic 02 — Skills and Rules  (Labs 4-6)](#topic-02--skills-and-rules--labs-4-6)
  - [Personalisation: Preferences That Follow You](#personalisation-preferences-that-follow-you)
  - [Workbook Rules: The .Rules Sheet](#workbook-rules-the-rules-sheet)
  - [Lab 4 · Workbook Rules & Personalisation — Clean Messy Client Orders with Cleanup, Personalisation and Rules](#lab-4--workbook-rules--personalisation--clean-messy-client-orders-with-cleanup-personalisation-and-rules)
  - [Key Finance Measures: CAC, LTV, Churn and Payback](#key-finance-measures-cac-ltv-churn-and-payback)
  - [Built-in Skills: Formatting, Visualisation and Finance](#built-in-skills-formatting-visualisation-and-finance)
  - [Lab 5 · Built-in Skills: Visual & Finance — Analyse HR-Outsourcing Finances with Built-in Skills](#lab-5--built-in-skills-visual--finance--analyse-hr-outsourcing-finances-with-built-in-skills)
  - [Custom Skills: The SKILL.md Recipe](#custom-skills-the-skillmd-recipe)
  - [Using Copilot to Write and Improve a Skill](#using-copilot-to-write-and-improve-a-skill)
  - [Lab 6 · Custom Skills: Calendar — Use Copilot to Build a Custom Interview-Calendar Skill](#lab-6--custom-skills-calendar--use-copilot-to-build-a-custom-interview-calendar-skill)
  - [Topic 02 Recap - Mapped to Learning Outcomes](#topic-02-recap---mapped-to-learning-outcomes)
- [Topic 03 — Interactive Dashboards with Copilot, Skills and Rules  (Labs 7-8)](#topic-03--interactive-dashboards-with-copilot-skills-and-rules--labs-7-8)
  - [Designing a Highly Visual Dashboard](#designing-a-highly-visual-dashboard)
  - [Plan and Build with Rules and Skills](#plan-and-build-with-rules-and-skills)
  - [Interactivity: Slicers, Timelines and KPI Cards](#interactivity-slicers-timelines-and-kpi-cards)
  - [Lab 7 · Interactive Dashboard — Build an Interactive Recruitment Dashboard with Rules and Skills](#lab-7--interactive-dashboard--build-an-interactive-recruitment-dashboard-with-rules-and-skills)
  - [Workforce Measures: Headcount, Attrition and Tenure](#workforce-measures-headcount-attrition-and-tenure)
  - [A Reusable Dashboard Skill and the Refresh Test](#a-reusable-dashboard-skill-and-the-refresh-test)
  - [Sharing People Dashboards Responsibly](#sharing-people-dashboards-responsibly)
  - [Lab 8 · Dashboard Skill, Refresh & Share — Turn the Dashboard into a Reusable Skill, Refresh It and Share It Safely](#lab-8--dashboard-skill-refresh--share--turn-the-dashboard-into-a-reusable-skill-refresh-it-and-share-it-safely)
  - [Topic 03 Recap - Mapped to Learning Outcomes](#topic-03-recap---mapped-to-learning-outcomes)
- [Wrap-Up - The Complete C34 Workflow](#wrap-up---the-complete-c34-workflow)
- [Next Steps](#next-steps)
- [Glossary](#glossary)


## Introduction

This Learner Guide accompanies Microsoft 365 Copilot for Excel (C34). It follows the same three topics - Excel fundamentals with Copilot, skills and rules, and interactive dashboards - and the same eight connected labs as the slide deck, Lesson Plan and lab folders. Every lab uses fictional data from BrightPath Talent Solutions, a made-up Singapore recruitment, staffing and HR-outsourcing firm.

Use the guide as a self-contained study text before, during and after class. Each concept explains what the capability is, why it matters, how it works, a worked example and when to use or avoid it, drawing on Microsoft Support guidance current in October 2026. Each lab folder contains an instruction PDF with sample prompts, a mock-data workbook and, where needed, a rules template or a ready-made SKILL.md. Copilot output remains a proposal until you inspect it and complete the Test It checks.


## Course Learning Outcomes

- LO1: Use Copilot in Excel's Chat only, Plan and Allow editing modes to prepare and clean data, write formulas and XLOOKUPs, and build PivotTables, PivotCharts and charts, verifying every result.
- LO2: Make Copilot's work consistent and repeatable with personalisation, workbook rules (.Rules), built-in skills and custom skills that you create with Copilot.
- LO3: Combine Copilot, rules and skills to build a highly visual, interactive and refreshable dashboard, and verify and share it responsibly.


## Before You Start — Preparation

**What you need**

- A Windows or Mac laptop with a current Excel for Microsoft 365, or a browser for Excel for the web.
- A Microsoft 365 work account with a Microsoft 365 Copilot licence (personal and free accounts do not include every feature used in the labs).
- Access to your OneDrive in a browser, with a folder named work-c34. Custom skills (Labs 6 and 8) are stored in a OneDrive skills folder.
- Office display language set to English: skills and workbook rules are fully supported in English only.
- The course labs folder: one sub-folder per lab with its instruction PDF, sample-prompts.txt, data/ workbook and any rules-template.txt or skills/ files.

**Verify your setup**

Upload labs/lab-01-*/data/BrightPath_Candidate_Pipeline.xlsx to your OneDrive work-c34 folder and open it in Excel. Confirm that Formulas > Calculation Options is Automatic and open Copilot. Check the mode menu above the prompt box (Chat only, Plan, Allow editing), the model picker at the top, the + (Add work content) menu and Settings (...) > Personalization and Manage plugins & skills. Labels and availability may change as Microsoft updates the product.

```bash
labs/
  lab-01-<name>/
    Lab-01-Instructions.pdf     instructions + sample prompts
    sample-prompts.txt          copy-ready prompts
    data/BrightPath_<dataset>.xlsx
    rules-template.txt          (Labs 4, 7, 8) .Rules text to paste
    skills/<name>/SKILL.md      (Labs 6, 8) ready-made skill
OneDrive/work-c34/              your cloud working copies
OneDrive/.../skills/            your custom skills
```

**Conventions used in every lab**

- All BrightPath companies, people, IDs and values are fictional training data.
- Prompts are shown in shaded boxes; copy them from sample-prompts.txt or the prompts web page and adapt them.
- Labels such as CHAT ONLY:, PLAN MODE: and ALLOW EDITING: tell you which mode to choose first.
- S$ is used for currency, dates are yyyy-mm-dd and the reporting date is 2026-06-30.
- Log each prompt, what Copilot did, how you verified it and your decision on the Prompt_Log sheet.
- Copilot output can vary between runs and models; the Test It values come from the data, so check against them.


## Topic 01 — Excel Fundamentals with Copilot  (Labs 1-3)

Chat only, Plan and Allow editing · setup, data sources and the model picker · data prep and automated cleanup · formulas and XLOOKUP · PivotTables, PivotCharts and charts

**Key concepts**

- Pick the right mode — Chat only answers without touching the sheet. Plan shows its steps first. Allow editing makes the change.
- Keep or undo every change — After Copilot edits, select Done to keep it or Undo to reverse it. Version history is your safety net.
- Turn your data into a table — Ask Copilot to format the range as a named table, e.g. tblApplications. Copilot gives better answers on tables.
- Name the columns in your prompt — 'Count hires by Source in tblApplications for 2026' beats 'analyse my data'.
- Describe the formula in plain English — 'Look up Recruiter_Name from tblRecruiters' - Copilot writes the XLOOKUP and explains it.
- Check every number — Compare Copilot's answer with your own COUNTIFS, SUMIFS or PivotTable total.

Concepts for Lab 1: What Is Data Prep?, Why Data Quality Matters, Chat Only, Plan and Allow Editing, Setup, Data Sources and the Model Picker, Cleaning Data with Copilot.


### What Is Data Prep?

Data preparation is the work of turning a raw export into a table you can analyse: one header row, one record per row, one kind of value per column, and no blank rows, merged cells, subtotals or notes inside the data. In Excel you finish by formatting the range as a named table (Home > Format as Table, or Ctrl+T, then a name such as tblApplications), so formulas, PivotTables, charts and Copilot all see the same, clearly bounded data.

Most reports start as an export from another system - an applicant tracking system, payroll or a survey tool - and exports are rarely analysis-ready. A clean table expands automatically when rows are added, keeps formulas consistent down the column, and lets you say exactly what one row means (the grain). If you cannot say what a row means, no formula or prompt can fix the logic.

**How it works**

- Move report titles, notes and subtotal rows out of the data range; keep one header row with unique names.
- Select the data and format it as a table (Ctrl+T); give it a clear name in Table Design > Table Name.
- Write the grain in one sentence, e.g. 'one row = one candidate application'.

**Worked example**

- BrightPath's pipeline export has 13 columns and one row per application.
- It is formatted as tblApplications, so new applications join the table automatically.
- Grain: one row = one candidate application, identified by Application_ID.

**Decision guide**

| Use when | Avoid when |
|---|---|
| You receive an export or a hand-built sheet that you will count, total or chart. | Summary rows, notes and data share one range - split them first. |
| Several people or tools (formulas, PivotTables, Copilot) will use the same data. | Cells are merged across rows or columns - unmerge before analysing. |

**Practitioner quality lens**

- Grain: One sentence that says what a row represents.
- Headers: One header row, unique names, no blanks.
- Table: Formatted as a named Excel table.

**Authoritative references**

- https://support.microsoft.com/en-us/office/create-and-format-tables-2f0ce0e7-500b-4fe1-9b5b-a16ba4b9f34a
- https://support.microsoft.com/en-us/excel/copilot/get-started-with-copilot-in-excel

---


### Why Data Quality Matters

Data quality is whether data is complete, consistent, unique, valid and accurate enough for the decision it supports. Typical problems are duplicate IDs, the same category spelt several ways ('LinkedIn', 'Linked In', 'JobStreet ' with a trailing space), numbers or dates stored as text, unexpected blanks, and values outside the allowed list.

Every count, total and chart inherits the errors in its source. Two duplicate rows inflate the application count; three spellings split LinkedIn into three sources; a date stored as text drops out of a monthly trend. Errors are cheap to fix before analysis and expensive after a report has been shared. Garbage in, garbage out - and AI does not change that: Copilot analyses the data you give it.

**How it works**

- Profile first: compare the row count with the unique-ID count, and list the distinct values of each category column.
- Flag problems before changing anything, and agree the valid values (the canonical list).
- Fix in a helper column or a copy, keep the original, then re-count to prove nothing was lost.

**Worked example**

- 242 rows but only 240 unique Application_IDs: 2 duplicates.
- 3 'Linked In' and 2 'JobStreet ' values would have under-counted two sources.
- A Source_Clean helper column fixes them; the original Source column is kept for audit.

**Decision guide**

| Use when | Avoid when |
|---|---|
| Before any count, rate, chart or dashboard is built on new data. | Fixing a value would mean guessing a missing business fact - flag it instead. |
| Data is typed by hand or merged from several systems. | Deleting rows you cannot explain - mark them and ask the data owner. |

**Practitioner quality lens**

- Unique: Row count equals unique-ID count, or the duplicates are explained.
- Consistent: Each category uses only the canonical spellings.
- Valid: Numbers and dates are real numbers and dates, not text.

**Authoritative references**

- https://support.microsoft.com/en-us/office/create-and-format-tables-2f0ce0e7-500b-4fe1-9b5b-a16ba4b9f34a
- https://support.microsoft.com/en-us/excel/copilot/copilot-in-excel-tips

---


### Chat Only, Plan and Allow Editing

Copilot in Excel is an AI assistant in a side pane that can answer questions about a workbook and make changes to it using Excel's own features: tables, formulas, PivotTables, charts, formatting, sorting and filtering. Above the prompt box you choose one of three modes. Chat only answers in the pane and never changes the workbook. Plan reasons through the task, shows you the steps and waits for your approval. Allow editing, the default, plans and makes the changes in one go.

Copilot is not just a chatbot: in Plan and Allow editing it works as an agent that edits your file. Choosing the mode on purpose keeps you in control. After every edit Copilot highlights what it changed and offers Done or Undo, and a cloud workbook keeps version history. Results stay ordinary, editable Excel content that you can inspect - and Copilot can still make mistakes, so you check before you keep.

**How it works**

- Open Copilot from Home > Copilot (or the floating Copilot icon) and choose the mode above the prompt box.
- Describe one task; in Plan mode read the steps and approve or correct them.
- Inspect the highlighted changes, then select Done to keep them or Undo to reverse them.

**Worked example**

- Chat only: 'What does one row in tblApplications represent?' - the sheet does not change.
- Plan: 'Flag duplicate Application_IDs and standardise Source.' Copilot lists four steps; you approve.
- Allow editing: 'Format this range as a table named tblApplications.' You review and select Done.

**Decision guide**

| Use when | Avoid when |
|---|---|
| Chat only for questions and analysis; Plan for big or multi-step changes; Allow editing for small, clear edits. | A decision about an individual (hiring, pay, exit) would rest on an unchecked AI output. |
| You can review each change and restore an earlier version if the result is wrong. | The data is not approved for use with Copilot under your organisation's policy. |

**Practitioner quality lens**

- Mode: Chosen before the prompt, not after the change.
- Control: Every edit reviewed, then kept with Done or reversed with Undo.
- Traceability: Prompt_Log, change history and version history kept.

**Authoritative references**

- https://support.microsoft.com/en-us/excel/copilot/get-started-with-copilot-in-excel
- https://support.microsoft.com/en-us/excel/copilot/copilot-in-excel-tips
- https://support.microsoft.com/en-us/excel/copilot/copilot-in-excel-change-history
- https://www.youtube.com/watch?v=4z4eIaEEH9A
- https://www.youtube.com/watch?v=wZ8QAKiIfuo

**The Copilot Pane — What Each Control Does**

- Copilot button or floating icon — Home > Copilot, or the floating icon at the bottom right. Right-click the icon to dock it or move it to the ribbon.
- Model picker (top) — Auto picks the model for each task. Or choose a Claude (Anthropic) or GPT (OpenAI) model for this session.
- Mode (above the prompt box) — Chat only answers. Plan shows its steps and waits. Allow editing (the default) makes the change.
- + Add work content — Attach a file, choose plugins & skills, create workbook rules, or switch web and work sources.
- Settings (...) — Personalization for your preferences; Manage plugins & skills to turn skills on or off.
- New chat and history — Start a new chat when you switch task or change a setting; reopen earlier chats from history.

**Choosing a Model**

- Auto (recommended) — Copilot chooses the model for each task. Start here.
- Claude - Anthropic — e.g. Claude Opus and Sonnet models. Deep reasoning models suit complex analysis.
- GPT - OpenAI — OpenAI's GPT models.
- Which one? — Try another model if an answer is weak. Premium models have fair-use limits. Always check the result.

---


### Setup, Data Sources and the Model Picker

Copilot appears in Excel for Microsoft 365 on Windows, Mac and the web when your account has an eligible plan, such as a Microsoft 365 Copilot licence on a work account. It needs an internet connection; saving the workbook to OneDrive or SharePoint adds AutoSave and version history, and recent builds also work on local files. From the + (Add work content) menu Copilot can use another file, web results or your work content as context, and the model picker at the top of the pane lets you choose Auto or a specific model.

Most 'Copilot is missing' problems are setup problems: an old Excel build, a licence or tenant setting, or an unsupported file. Bringing in the right file gives Copilot facts it cannot guess - but naming the file beats letting it search. Models have different strengths: Auto suits most tasks, a deeper reasoning model can catch errors in a complex analysis, and premium models have fair-use limits.

**How it works**

- Save the workbook as .xlsx in OneDrive, set Formulas > Calculation Options to Automatic and open Copilot.
- Use + > attach a file (or type / and the file name) to give Copilot another file as context.
- Leave the model on Auto; switch model and start a New chat to compare answers on a hard question.

**Worked example**

- A recruiter attaches the client rate-card PDF and asks Copilot to add the agreed fee rate to each job order.
- Auto says 'revise' on a business case; a deeper reasoning model finds the calculation error behind it.
- Both models' hire counts are checked against the same COUNTIFS formula.

**Decision guide**

| Use when | Avoid when |
|---|---|
| The extra file or web source is approved and you can name it precisely. | Several versions of a file exist and Copilot would have to guess which one you meant. |
| An answer is weak or high-stakes and a second model's view is worth comparing. | You would accept a premium model's answer without checking it - a better model is not always right. |

**Practitioner quality lens**

- Setup: Current Excel build, eligible licence, Automatic calculation, cloud .xlsx.
- Sources: Named file or source; web and work toggles set on purpose.
- Models: Auto by default; compare models only against a known check.

**Authoritative references**

- https://support.microsoft.com/en-us/excel/copilot/frequently-asked-questions-about-copilot-in-excel
- https://support.microsoft.com/en-us/microsoft-365-copilot/frequently-asked-questions-about-copilot-in-microsoft-365-subscriptions
- https://support.microsoft.com/en-us/excel/copilot/import-data-from-the-web-and-internal-sources-with-copilot-in-excel
- https://support.microsoft.com/en-us/excel/copilot/import-data-from-another-excel-file-with-copilot-in-excel
- https://www.youtube.com/watch?v=4z4eIaEEH9A
- https://www.youtube.com/watch?v=wZ8QAKiIfuo
- https://www.youtube.com/watch?v=gh8gs2eSc0c

---


### Cleaning Data with Copilot

Copilot works best on tables: one header row, one record per row, one meaning per column and consistent values. Formatting a range as a named Excel table makes its boundary clear to people, formulas and Copilot. Clean data with Copilot (Data > Clean Data, or the suggestion Excel shows when it opens a messy file) finds extra spaces, inconsistent capitalisation and similar problems, and lets you apply each fix.

Duplicated IDs and labels such as 'LinkedIn', 'Linked In' and 'JobStreet ' split categories and corrupt every count built on them. Automated cleanup removes the typing problems quickly, but it cannot know your business meaning - that WIP means In Progress, or that a note saying 'paused' makes an order On Hold. That knowledge belongs in a prompt or, better, in workbook rules (Topic 2).

**How it works**

- Ask Copilot to format the data as a named table and to report the row count and unique-ID count.
- Ask Copilot (Chat only) to list duplicates and value variants before anything is changed.
- Fix with helper columns or Clean Data suggestions, then re-count to prove nothing was lost.

**Worked example**

- BrightPath's pipeline export has 242 rows but only 240 unique Application_IDs.
- Five Source values use non-standard spellings that would split the LinkedIn and JobStreet totals.
- Copilot adds a Source_Clean column and highlights the cells it fixed; the original Source stays unchanged.

**Decision guide**

| Use when | Avoid when |
|---|---|
| The data can be expressed as a rectangular list with a stable row meaning. | A fix would require guessing a missing business fact. |
| Each column has one data type and a documented list of valid values. | Applying every suggestion blindly, without hovering to see which cells change. |

**Practitioner quality lens**

- Grain: One sentence that says exactly what a row represents.
- Schema: Unique headers, consistent types and canonical category values.
- Controls: Row and unique-ID counts recorded before and after any change.

**Authoritative references**

- https://support.microsoft.com/en-us/office/create-and-format-tables-2f0ce0e7-500b-4fe1-9b5b-a16ba4b9f34a
- https://support.microsoft.com/en-us/excel/copilot/get-started-with-copilot-in-excel
- https://support.microsoft.com/en-us/excel/copilot/copilot-in-excel-tips
- https://www.youtube.com/watch?v=4z4eIaEEH9A

---


### Lab 1 · Data Prep & Data Quality — Explore and Question Recruitment Pipeline Data

Learning outcome: LO1: confirm Copilot access and use the Chat only, Plan and Allow editing modes to prepare, check, highlight, sort and filter recruitment pipeline data.

Goal: Turn a raw candidate-pipeline export into a named Excel table, find its data-quality issues with Copilot, and answer pipeline questions with checked numbers.

You are a talent analyst at BrightPath Talent Solutions, a fictional Singapore recruitment and staffing firm. You will meet Copilot's three modes on a real task: Allow editing turns the H1 2026 pipeline export into a named table, Chat only describes it without touching the sheet, and Plan shows you its steps before it fixes duplicate IDs and inconsistent Source labels. You will then answer pipeline questions, compare two AI models on the same question, and build a highlighted, sorted and filtered review view. Every number is checked with a formula.

**What you'll build**

BrightPath_Candidate_Pipeline.xlsx in OneDrive with table tblApplications, a Source_Clean column, duplicate IDs highlighted, Hired rows highlighted, a verified source-by-hires answer and a filled Prompt_Log   (Tools: Excel for Microsoft 365 or Excel for the web · Copilot (Chat only, Plan, Allow editing, model picker) · BrightPath_Candidate_Pipeline.xlsx.)

**Prerequisites**

- A Microsoft 365 account with Copilot in Excel enabled, signed in to Excel or Excel for the web.
- A OneDrive or SharePoint folder named work-c34 that you can upload to.
- This lab folder downloaded, including data/BrightPath_Candidate_Pipeline.xlsx.

**Step-by-step**

1. Go to onedrive.com (or your SharePoint site), open or create the folder work-c34 and select Upload > Files. Upload BrightPath_Candidate_Pipeline.xlsx from this lab's data folder, then open it in Excel for the web or select Open in Desktop App. OneDrive gives you AutoSave and version history (recent Excel builds can also run Copilot on local files). Check that Formulas > Calculation Options is Automatic, then open Copilot from the Copilot button on the Home tab or the floating Copilot icon at the bottom right. Read the About_This_Data sheet before you prompt.

   ```text
   Required state:
   File      : OneDrive/work-c34/BrightPath_Candidate_Pipeline.xlsx
   AutoSave  : On (recommended, not required)
   Calculation Options : Automatic
   Sheets    : About_This_Data | Applications | Prompt_Log
   ```

2. Open Copilot and find the mode menu above the prompt box: Chat only (answers, no changes), Plan (shows its steps first) and Allow editing (makes the changes). Choose Allow editing and ask Copilot to turn the Applications data into a named table. Its reply must confirm tblApplications, 13 columns and 242 data rows. Select Done to keep the change, or Undo to reverse it. If Excel offers 'Clean data with Copilot', select Dismiss for now - you will use it in Lab 4.

   ```text
   ALLOW EDITING:
   Format the data on the Applications sheet (from A1 to the last row and column) as an Excel table with headers. Name the table tblApplications. Do not change, sort or delete any values.
   Then tell me the table name, the number of columns and the number of data rows.
   ```

3. Switch to Chat only (no workbook changes). Paste the prompt below. Copilot's row count must match the 242 rows it reported in Step 2. If Copilot already points out Source labels such as 'Linked In' or 'JobStreet ' (trailing space), note them - Step 4 checks and fixes them. You will log every prompt in Step 7.

   ```text
   GOAL: Describe this candidate pipeline before I analyse it.
   CONTEXT: BrightPath recruitment desk, H1 2026. One row should be one application.
   DATA: Use only tblApplications.
   CONSTRAINTS: Do not change the workbook.
   OUTPUT: What a row represents, each column with its data type, the number of rows, the Applied_Date range and the distinct values of Stage and Source.
   REVIEW: Tell me how you counted the rows.
   ```

4. Still in Chat only, ask Copilot to check data quality. It should report 2 repeated Application_IDs and Source values that are not in the canonical list on About_This_Data. Then switch to Plan mode and send the two changes: Copilot first shows its plan - read it, approve it, then select Done. Orange means 'this ID appears more than once': each duplicate ID is coloured in both places it occurs, so you see one orange cell near the top of the table and its twin further down (press Ctrl+End to jump to the last rows). Check that no row was deleted - the table must still have 242 rows.

   ```text
   CHAT ONLY:
   Check tblApplications for data-quality issues. List any Application_ID that appears more than once, and any Source value that is not exactly one of: LinkedIn, JobStreet, MyCareersFuture, Referral, BrightPath Website, Job Fair, Talent Database. Give the count of each issue. Do not change anything.

   PLAN MODE:
   1) Highlight Application_ID cells that appear more than once in tblApplications with an orange fill.
   2) Add a column Source_Clean to tblApplications that trims spaces from Source and changes 'Linked In' to 'LinkedIn'. Do not delete or edit any existing values.

   FOLLOW-UP: For each duplicated Application_ID, give the row numbers of both copies and say whether the two rows are identical.

   FOLLOW-UP: Highlight the Source cells in tblApplications where Source_Clean is different from Source with a yellow fill. Tell me how many cells you highlighted and list their Application_IDs.

   FOLLOW-UP: List the distinct Source_Clean values with a count of each. Confirm there is no 'Linked In' and no value with a trailing space, and that tblApplications still has 242 rows.
   ```

5. Ask Copilot a direct question about which sources produce hires, then ask Copilot (Allow editing) to add the two check formulas to a Checks sheet. Duplicate rows are Rejected applications, so they do not change the hire counts. Finally, open the model picker at the top of the pane (Auto by default), choose a different model, select New chat and ask the same question again. Both answers must match the Checks formulas - if a model disagrees, the formula wins.

   ```text
   PROMPT (Chat only):
   Using tblApplications, count applications and hires (Stage = "Hired") by Source_Clean.
   Show a table sorted by hires, highest first, with hire rate = hires / applications.

   CHECK (Allow editing):
   On a sheet named Checks (create it if needed), add these formulas with a label beside each, then tell me the results:
   =COUNTIFS(tblApplications[Stage],"Hired")
   =COUNTIFS(tblApplications[Source_Clean],"LinkedIn",tblApplications[Stage],"Hired")

   CHAT ONLY (new chat, different model):
   Using tblApplications, count applications and hires (Stage = "Hired") by Source_Clean.
   Show a table sorted by hires, highest first.
   ```

6. Switch to Allow editing and build a review view, one change at a time. After each change, look at the sheet and keep it or select Undo in the Copilot pane. Finally ask Copilot how many rows the filter shows, then ask it to clear the filter so all 242 rows return.

   ```text
   1) Highlight entire rows of tblApplications where Stage is "Hired" with a light green fill.
   2) Sort tblApplications by Interview_Score from highest to lowest, then by Days_in_Pipeline from lowest to highest.
   3) Filter tblApplications to show only Client_Sector = "Technology" and Stage = "Interview".
   4) How many rows of tblApplications are visible with this filter?
   5) Clear all filters on tblApplications and confirm all 242 rows are visible.
   ```

7. Prompt_Log already lists every prompt in this lab with its step and mode. Ask Copilot to fill in what it did, how it was verified and your decision (Keep, Revise or Undo). With AutoSave on, the cloud copy and its version history stay current.

   ```text
   LOG (Allow editing):
   Prompt_Log already lists every prompt in this lab. For each one I ran, fill in What Copilot did and How I verified, and set Decision to Keep. I will tell you any row to change to Revise or Undo.
   ```


**Sample prompts**

| Purpose | Prompt |
|---|---|
| Describe | Describe tblApplications: what one row represents, each column's data type, the row count and the Applied_Date range. Do not change the workbook. |
| Data quality | List any Application_ID that appears more than once in tblApplications and any Source value that is not in the canonical list on About_This_Data. |
| Plan first | Plan how you would flag duplicate Application_IDs and standardise Source. Show me the steps before you change anything. |
| Clean helper column | Add a column Source_Clean that trims spaces from Source and replaces 'Linked In' with 'LinkedIn'. Do not change the original Source column. |
| Question | Using tblApplications, count applications and hires by Source_Clean and show the hire rate. Sort by hires, highest first. |
| Question | Which Client_Sector received the most applications, and how many? Show the count for every sector. |
| Highlight | Highlight rows where Stage is Hired in light green. |
| Sort | Sort tblApplications by Interview_Score from highest to lowest, then by Days_in_Pipeline from lowest to highest. |
| Filter | Filter to show only Technology applications at the Interview stage. |
| Insight | What's interesting about this pipeline data? Give three observations, each with the numbers behind it. |

**Test it**

| Check | Expected result |
|---|---|
| Rows in tblApplications | 242 rows, 240 unique Application_IDs |
| Duplicate IDs (orange) | APP-0018 and APP-0020 - 4 orange cells: 2 near the top, 2 in the last rows |
| Fixed Source cells (yellow) | 5 cells: 3 'Linked In' + 2 'JobStreet ' |
| LinkedIn count in Source_Clean | 69 |
| Hires | 22 in total; LinkedIn has the most (8) |
| Technology + Interview filter | 9 rows; all 242 back after clearing it |

**Checkpoint and rejoin point**

Keep this workbook in work-c34. Lab 2 uses a separate requisitions workbook, so a learner who is behind can start Lab 2 from its own data folder.

**Troubleshooting**

| If this happens | Fix |
|---|---|
| The Copilot button is missing or greyed out. | Save the file as .xlsx in OneDrive or SharePoint, close and reopen it. Check your licence in your Microsoft 365 account and ask IT if Copilot is disabled for your tenant. Pair with a neighbour meanwhile. |
| There is no Plan option in the mode menu. | Some Excel builds show only Allow editing and Chat only. Use Chat only, end the prompt with 'Show me the plan; do not change anything yet', then switch to Allow editing to carry it out. |
| 'Unsupported file state' message. | If the file is checked out in SharePoint, check it in and reopen it, or open it in Excel for the web. Strict Open XML files are not supported - save as a normal .xlsx. |
| Copilot answers about the whole sheet instead of the table. | Name the table (tblApplications) and the exact columns in the prompt, and click inside the table before asking. |
| Copilot's count differs from your formula. | Ask Copilot how it counted (rows vs unique IDs, which Stage values). Trust the formula, record the difference and refine the prompt. |

**Challenge**

Ask Copilot to add a Pipeline_Status column that groups Stage into Active (Applied, Screening, Interview, Offer), Hired and Closed (Rejected, Withdrawn). Verify the three counts add up to the row count.

**Reflection**

Which data-quality issue would most distort a source-effectiveness report if you had not found it, and why?

> **Note:** The lab folder labs/lab-01-*/ holds the instruction PDF with sample prompts, sample-prompts.txt and the mock-data workbook (data/BrightPath_Candidate_Pipeline.xlsx). Use only the supplied fictional BrightPath data or information you are authorised to process. Do not paste credentials, confidential business material or real employees' personal data into Copilot, rules or skills.

---

Concepts for Lab 2: What Is a Formula in Excel?, Table References: What Does @ Mean?, Reading Excel References: !, :, $ and [@[...]], Reading a Formula: IF, AND and @, What Is XLOOKUP?, Writing Formulas and XLOOKUPs with Copilot.


### What Is a Formula in Excel?

A formula starts with = and calculates a value from cells, constants and functions such as SUM, IF, COUNTIFS and SUMIFS. Inside a table, structured references use names instead of cell addresses: =[@Filled]/[@Openings] works on the current row, and =SUM(tblRequisitions[Filled]) on a whole column. A formula typed once in a table column fills the entire column automatically.

Formulas make results reproducible and auditable: anyone can click a cell, read how it was calculated, and trust that it recalculates when the data changes - a typed number does neither. Structured references read like the business rule and do not break when rows are added or moved.

**How it works**

- Write the business rule in words first, including what happens with blanks and open items.
- Build it from functions: IF for conditions, COUNTIFS and SUMIFS for conditional counts and totals.
- Check one row by hand and compare a column total with an independent formula.

**Worked example**

- Fill rate: =[@Filled]/[@Openings], formatted as a percentage.
- Open or closed: =IF([@Close_Date]="","Open","Closed").
- Open requisitions: =COUNTIFS(tblRequisitions[Status],"Open").

**Decision guide**

| Use when | Avoid when |
|---|---|
| A value can be calculated from other values by a fixed rule. | Typing a number that could be calculated - it will go out of date. |
| The result must update when the data changes. | Hard-coding a cell such as C2 when a table column name would be clearer. |

**Practitioner quality lens**

- Readable: Structured references and clear column names.
- Complete: Blanks, zero and open items handled on purpose.
- Checked: One row hand-worked, one total reconciled.

**Authoritative references**

- https://support.microsoft.com/en-us/office/create-and-format-tables-2f0ce0e7-500b-4fe1-9b5b-a16ba4b9f34a
- https://support.microsoft.com/en-us/excel/copilot/get-started-with-copilot-in-excel

---


### Table References: What Does @ Mean?

Inside an Excel table, formulas use column names instead of cell addresses. The @ sign means 'this row': [@Filled] is the Filled value in the same row as the formula, so =[@Filled]/[@Openings] divides this row's Filled by this row's Openings. Without the @, tblRequisitions[Filled] means the whole Filled column - the form used for totals such as =SUM(tblRequisitions[Filled]).

Copilot writes table formulas this way, so you need to read them. A row formula with @ is typed once and fills every row, each row calculating with its own numbers. It reads like the business rule instead of =G2/F2, and it keeps working when rows are added, sorted or moved, because it refers to column names, not positions. You only see @ inside an Excel table; in a normal range Excel uses addresses like G2.

**How it works**

- Read [@Column] as 'this row's Column' and Table[Column] as 'the whole Column'.
- Use @ for a calculated column (one result per row); drop it for totals and counts over the column.
- Click a cell in the column and check the formula bar: the same formula appears in every row.

**Worked example**

- Row 3 (JOB-2002): =[@Filled]/[@Openings] means the same as =G3/F3 - 4 / 4 = 100%.
- =SUM(tblRequisitions[Filled]) adds the Filled column for all 60 requisitions.
- Copilot reports 'Added Fill_Rate for all 60 rows' because one @ formula filled the column.

**Decision guide**

| Use when | Avoid when |
|---|---|
| Your data is an Excel table and the formula works row by row. | Mixing @ into a total - =SUM([@Filled]) only adds one row. |
| You want a formula that survives sorting, filtering and new rows. | Expecting @ in a plain range - convert it to a table first (Ctrl+T). |

**Practitioner quality lens**

- Row vs column: @ for one row; no @ for the whole column.
- Consistent: One formula down the whole column, no overrides.
- Readable: Column names that say what the value is.

**Authoritative references**

- https://support.microsoft.com/en-us/office/create-and-format-tables-2f0ce0e7-500b-4fe1-9b5b-a16ba4b9f34a

---


### Reading Excel References: !, :, $ and [@[...]]

A few symbols tell you where a formula looks. The ! separates a sheet name from a cell address: Control!B1 is cell B1 on the Control sheet, and a sheet name with spaces gets quotes, as in 'Dash Data'!B2. The colon makes a range: P2:P61 is every cell from P2 down to P61. The $ locks part of an address so it does not shift when the formula is copied: $B$1 always points at B1. And [@[Column]] is the same as [@Column] - this row's value - with extra brackets that Excel adds around a column name, and that are required when the name contains spaces or special characters.

Copilot's replies and formulas are full of these references - 'Added Est_Fee for all 60 rows in Requisitions!P2:P61' tells you exactly which cells changed. Being able to read them is how you check that a formula points at the right sheet, the right range and the right fixed input, such as the report date. A missing $ is a classic copy error: the reference slides down the sheet with the formula.

**How it works**

- Read Sheet!Cell as 'this cell on that sheet', and A:B ranges as 'from A to B'.
- Use $ (press F4 while editing a reference) for a fixed input every row must use, such as Control!$B$1.
- Read [@[Column]] exactly like [@Column]: this row's value in that column.

**Worked example**

- Requisitions!P2:P61 - the Est_Fee column, rows 2 to 61, on the Requisitions sheet.
- Days_Open for open roles uses Control!$B$1, the report date, so every row uses the same date.
- =[@Filled]*[@[Monthly_Salary_Budget]]*12*[@[Fee_Rate]] - this row's Filled x salary x 12 x fee rate.

**Decision guide**

| Use when | Avoid when |
|---|---|
| Checking where a Copilot formula or change points. | Copying a formula that should use one fixed cell without a $ - it will shift. |
| A formula needs a value from another sheet or one fixed cell. | Typing a value into many formulas instead of referring to one input cell. |

**Practitioner quality lens**

- Right sheet: The sheet name before ! is the one you meant.
- Right range: The range covers every row - no more, no fewer.
- Fixed inputs: Shared inputs are locked with $ or held in one named cell.

**Authoritative references**

- https://support.microsoft.com/en-us/office/create-and-format-tables-2f0ce0e7-500b-4fe1-9b5b-a16ba4b9f34a

---


### Reading a Formula: IF, AND and @

Lab 2's Ageing_Flag column uses =IF(AND([@Status]="Open", [@[Days_Open]]>45), "Overdue", ""). Read it from the inside out. [@Status]="Open" tests whether this row's Status is Open; [@[Days_Open]]>45 tests whether this row has been open more than 45 days; AND(...) is TRUE only when both tests are TRUE; and IF(test, "Overdue", "") shows Overdue when the test is TRUE and leaves the cell blank when it is FALSE.

Copilot writes formulas like this for you, but you still own the result - so you need to be able to read one. The @ makes the formula check each requisition on its own row; without it, [Status] would mean all 60 values at once and the formula would return an error or a wrong result. Reading the formula is also how you confirm the business rule: Open only, and strictly more than 45 days.

**How it works**

- Start with the innermost tests and say each one in words.
- Then read the function that combines them (AND, OR) and the IF that chooses the result.
- Check the rule with a count: =COUNTIF(tblRequisitions[Ageing_Flag],"Overdue").

**Worked example**

- A requisition with Status Open and Days_Open 52 shows Overdue.
- Open and 30 days, or Filled and 80 days, stays blank - both tests must be TRUE.
- BrightPath has 13 Overdue requisitions, matching =COUNTIFS(... "Open", ... ">45").

**Decision guide**

| Use when | Avoid when |
|---|---|
| A flag depends on two or more conditions at the same time. | The rule is 'either' condition - use OR instead of AND. |
| You need to confirm what a Copilot-written formula really tests. | The threshold (45 days) is unclear - agree it before the formula is written. |

**Practitioner quality lens**

- Rule: Each test matches the business rule, including > vs >=.
- Row by row: @ on every column reference inside a calculated column.
- Counted: The flag count equals an independent COUNTIFS.

**Authoritative references**

- https://support.microsoft.com/en-us/office/create-and-format-tables-2f0ce0e7-500b-4fe1-9b5b-a16ba4b9f34a

---


### What Is XLOOKUP?

XLOOKUP finds a value in one column and returns the matching value from another column: =XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found]). For example =XLOOKUP([@Recruiter_ID], tblRecruiters[Recruiter_ID], tblRecruiters[Recruiter_Name], "Not found") brings each recruiter's name into the requisitions table. It replaces VLOOKUP and INDEX/MATCH for most lookups.

The data you need often lives in another table - names in a staff list, rates in a rate card, targets in a plan. A lookup joins the two by a shared key instead of copy-paste, and updates when either table changes. XLOOKUP is easier and safer than VLOOKUP: it matches exactly by default, can look left, does not break when columns are inserted, and lets you choose what to show when nothing is found.

**How it works**

- Find the key both tables share (e.g. Recruiter_ID) and make sure it is unique in the lookup table.
- Write =XLOOKUP(key, lookup column, return column, "Not found") in a new table column.
- Count "Not found" results - each one is a missing or mistyped key to fix at the source.

**Worked example**

- Requisition JOB-2002 has Recruiter_ID R05; XLOOKUP returns Siti Hamid from tblRecruiters.
- An ID typed as 'R05 ' (trailing space) would return Not found - the key must match exactly.
- =COUNTIF(tblRequisitions[Recruiter_Name],"Not found") returns 0 when every key matches.

**Decision guide**

| Use when | Avoid when |
|---|---|
| Two tables share a key and you need a value from one in the other. | The key is not unique in the lookup table - XLOOKUP returns only the first match. |
| The lookup table is maintained separately and changes over time. | Keys differ in type or spacing (text vs number, trailing spaces) - clean them first. |

**Practitioner quality lens**

- Key: Unique in the lookup table, identical in both tables.
- Not found: A clear fallback value, then counted and explained.
- Spot-check: One looked-up value compared with the source table.

**Authoritative references**

- https://support.microsoft.com/en-us/office/xlookup-function-b7fd680e-6d10-43e6-84f9-88eae8bf5929
- https://support.microsoft.com/en-us/office/create-and-format-tables-2f0ce0e7-500b-4fe1-9b5b-a16ba4b9f34a

---


### Writing Formulas and XLOOKUPs with Copilot

Describe a calculation in business language and Copilot writes the formula: calculated columns, total rows, conditional flags and lookups such as XLOOKUP that pull a matching value from another table. In a table it uses structured references such as [@Filled] that fill the whole column. Copilot can also explain any formula, and it can do several related steps from one prompt.

You no longer need to remember XLOOKUP's arguments - but a formula can be valid and still wrong: Days_Open may ignore open requisitions, a fee may use the monthly instead of the annual salary, or a blank may become zero. Using the exact table and column names in the prompt, asking for an explanation and hand-checking a row turns a generated formula into evidence.

**How it works**

- Write the business rule in words, using the exact sheet, table and column names.
- Ask for the column (or two or three related steps) and for an explanation of the formula.
- Hand-check one normal and one edge-case row, and tie the column total to an independent formula.

**Worked example**

- 'Add Recruiter_Name to tblRequisitions by looking it up from tblRecruiters using Recruiter_ID.'
- 'Add Est_Fee = Filled x Monthly_Salary_Budget x 12 x Fee_Rate, then add a total row.'
- JOB-2002: 4 x S$4,900 x 12 x 25% = S$58,800 - matches the generated value.

**Decision guide**

| Use when | Avoid when |
|---|---|
| The rule is deterministic and can be written with native Excel functions. | A subjective judgement is dressed up as a precise calculation. |
| A hand-worked sample or control total can verify the result. | The formula hides missing data or an unapproved threshold. |

**Practitioner quality lens**

- Semantics: Numerator, denominator, units and report date are right.
- Edge cases: Blanks, zero, open items and one large value tested.
- Reconcile: Column total equals an independent control.

**Authoritative references**

- https://support.microsoft.com/en-us/excel/copilot/get-started-with-copilot-in-excel
- https://support.microsoft.com/en-us/office/xlookup-function-b7fd680e-6d10-43e6-84f9-88eae8bf5929
- https://support.microsoft.com/en-us/office/create-and-format-tables-2f0ce0e7-500b-4fe1-9b5b-a16ba4b9f34a
- https://www.youtube.com/watch?v=4z4eIaEEH9A

---


### Lab 2 · Formulas & XLOOKUP — Generate and Verify Formula Columns and XLOOKUPs for Job Requisitions

Learning outcome: LO1: use Copilot to write XLOOKUP and calculated columns from plain-English rules, explain them, and reconcile them to independent control formulas.

Goal: Add recruiter names, days open, fill rate, estimated placement fee and an ageing flag to a requisitions table, then prove each column is right.

BrightPath's permanent-recruitment desk tracks 60 client job orders. You will ask Copilot to add an XLOOKUP for the recruiter's name and four calculated columns written from business rules, ask it to explain each formula, hand-check one requisition and reconcile column totals to independent formulas on the Control sheet. You will then ask Copilot which recruiter generated the most estimated fees and verify the answer.

**What you'll build**

BrightPath_Job_Requisitions.xlsx with Recruiter_Name, Days_Open, Fill_Rate, Est_Fee and Ageing_Flag columns in tblRequisitions, overdue requisitions highlighted, and a Control sheet that reconciles the totals   (Tools: Excel · Copilot (Allow editing, Chat only) · XLOOKUP · structured references · BrightPath_Job_Requisitions.xlsx.)

**Prerequisites**

- Lab 1 completed, or Copilot access confirmed on a OneDrive .xlsx file.
- This lab folder's data/BrightPath_Job_Requisitions.xlsx.

**Step-by-step**

1. Go to onedrive.com (or your SharePoint site), open or create the folder work-c34 and select Upload > Files. Upload BrightPath_Job_Requisitions.xlsx from this lab's data folder, then open it in Excel for the web or select Open in Desktop App. OneDrive gives you AutoSave and version history (recent Excel builds can also run Copilot on local files). Check that Formulas > Calculation Options is Automatic, then open Copilot from the Copilot button on the Home tab or the floating Copilot icon at the bottom right. Read the About_This_Data sheet before you prompt. Note the fee rule: Est_Fee = Filled x Monthly_Salary_Budget x 12 x Fee_Rate, and the report date in Control!B1.

   ```text
   Tables : tblRequisitions (60 rows) | tblRecruiters (6 rows)
   Control!B1 Report_Date = 2026-06-30
   ```

2. Choose Allow editing. Ask Copilot to add the lookup column, then ask it to explain the formula and show one row next to the matching row in tblRecruiters.

   ```text
   In tblRequisitions, add a column Recruiter_Name that looks up Recruiter_Name from tblRecruiters using Recruiter_ID. Use XLOOKUP and return "Not found" if there is no match.

   Then: Explain the Recruiter_Name formula, and show one row's Recruiter_ID and Recruiter_Name next to the matching row in tblRecruiters.
   ```

3. Ask for Days_Open with an explicit rule for open requisitions. Read the explanation, then ask Copilot to show an open row (blank Close_Date) and a closed row so you can confirm the rule.

   ```text
   Add a column Days_Open to tblRequisitions. If Close_Date is blank, use Control!B1 minus Open_Date; otherwise use Close_Date minus Open_Date. Show whole numbers.

   FOLLOW-UP: Show one open requisition (blank Close_Date) and one closed requisition with Open_Date, Close_Date and Days_Open, and the subtraction behind each.
   ```

4. Ask for Fill_Rate and Est_Fee, one prompt each. Format Fill_Rate as a percentage and Est_Fee as S$ with no decimals. Then ask Copilot to show the calculation for JOB-2002: Filled 4 x S$4,900 x 12 x 25% = S$58,800.

   ```text
   Add a column Fill_Rate = Filled divided by Openings, formatted as a percentage.

   Add a column Est_Fee = Filled x Monthly_Salary_Budget x 12 x Fee_Rate, formatted as S$ with no decimals.
   Explain the formula.

   FOLLOW-UP: Show the Est_Fee calculation for JOB-2002 step by step.
   ```

5. Ask for the ageing flag and a highlight, then ask Copilot how many requisitions are flagged.

   ```text
   Add a column Ageing_Flag that shows "Overdue" when Status is "Open" and Days_Open is greater than 45, otherwise leave it blank. Highlight rows where Ageing_Flag is "Overdue" with a light amber fill.

   FOLLOW-UP: How many rows have Ageing_Flag = "Overdue"?
   ```

6. Ask Copilot to build independent controls on the Control sheet with the exact formulas below. The independent formulas recalculate from the raw columns, so they do not rely on Copilot's new columns. Each pair must match.

   ```text
   Enter these labels and formulas on the Control sheet exactly as written. Then tell me each result and whether each pair (B3 and B4, B6 and B7) matches:
   A3 Total Est_Fee (column)       B3 =SUM(tblRequisitions[Est_Fee])
   A4 Total fee (independent)      B4 =SUMPRODUCT(tblRequisitions[Filled],tblRequisitions[Monthly_Salary_Budget],tblRequisitions[Fee_Rate])*12
   A5 Avg days to fill (Filled)    B5 =AVERAGEIFS(tblRequisitions[Days_Open],tblRequisitions[Status],"Filled")
   A6 Overdue (independent)        B6 =COUNTIFS(tblRequisitions[Status],"Open",tblRequisitions[Days_Open],">45")
   A7 Overdue flags (column)       B7 =COUNTIF(tblRequisitions[Ageing_Flag],"Overdue")
   ```

7. In Chat only, ask which recruiter generated the most estimated fees. Then ask Copilot (Allow editing) to add the SUMIFS check to the Control sheet and to log all prompts on Prompt_Log.

   ```text
   CHAT ONLY:
   Using tblRequisitions, which Recruiter_Name has the highest total Est_Fee? Show the total for every recruiter.

   CHECK (Allow editing):
   On the Control sheet, add this check for the top recruiter and tell me if it matches your answer:
   =SUMIFS(tblRequisitions[Est_Fee],tblRequisitions[Recruiter_Name],"<name>")

   LOG (Allow editing):
   Prompt_Log already lists every prompt in this lab. For each one I ran, fill in What Copilot did and How I verified, and set Decision to Keep. I will tell you any row to change to Revise or Undo.
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
| Several steps | Add Fill_Rate and Est_Fee columns, add a total row, then tell me which recruiter has the highest total Est_Fee. |

**Test it**

| Check | Expected result |
|---|---|
| Total estimated fees (Control B3 and B4) | Both S$1,030,512 |
| Average days to fill (B5) | 34.6 days (29 filled requisitions) |
| Overdue requisitions (B6 and B7) | Both 13 |
| Sample row JOB-2002 | Est_Fee S$58,800, Days_Open 36 |
| Recruiter_Name lookup | Every row has a name - no "Not found" |
| Top recruiter by fees | Aisha Rahman - S$221,760 |

**Checkpoint and rejoin point**

Keep the workbook in work-c34. Lab 3 uses its own metrics workbook in the Lab 3 folder.

**Troubleshooting**

| If this happens | Fix |
|---|---|
| Days_Open is negative or huge for open requisitions. | The formula used today's date or a blank as zero. Ask Copilot to use Control!B1 explicitly, or edit the formula. |
| Est_Fee is 12 times too small or too large. | Re-read the rule: fee is on first-year salary (monthly x 12). Ask Copilot to explain the formula and correct it. |
| The new column is outside the table. | Click inside tblRequisitions before prompting and name the table in the prompt; or drag the table's resize handle. |
| Copilot cannot add a formula column. | Confirm Calculation Options is Automatic and the data is an Excel table; try Excel for the web if desktop fails. |

**Challenge**

Ask Copilot to add a Fee_Achieved column that is Est_Fee for Filled requisitions only, then build a small summary of Fee_Achieved by Sector with SUMIFS. Check that the sector totals add up to the grand total. Then bring in other data: upload Lab 1's BrightPath_Candidate_Pipeline.xlsx to work-c34, select + (Add work content) and attach it (or type / and its name), and ask: 'Add a sheet Sector_View that compares open requisitions by Sector in this workbook with hires by Client_Sector in the attached pipeline.' Check one sector's hires with COUNTIFS in the pipeline workbook.

**Reflection**

Which of today's formulas would have looked correct but been wrong without a hand-check, and how did you catch it?

> **Note:** The lab folder labs/lab-02-*/ holds the instruction PDF with sample prompts, sample-prompts.txt and the mock-data workbook (data/BrightPath_Job_Requisitions.xlsx). Use only the supplied fictional BrightPath data or information you are authorised to process. Do not paste credentials, confidential business material or real employees' personal data into Copilot, rules or skills.

---

Concepts for Lab 3: What Are PivotTables and PivotCharts?, PivotTables and PivotCharts from a Prompt, Choosing and Refining Charts.


### What Are PivotTables and PivotCharts?

A PivotTable summarises a table without writing formulas: you place fields in Rows, Columns, Values (Sum, Count or Average) and Filters, and Excel groups and totals the data. A PivotChart is a chart drawn from a PivotTable; it changes when the PivotTable is filtered, rearranged or refreshed. Create them from Insert > PivotTable, then PivotTable Analyze > PivotChart.

PivotTables answer 'how many' and 'how much by group' questions in seconds, and you can reshape the view by dragging a field instead of rewriting formulas. The grand total gives every chart a number you can check. One catch: a PivotTable does not update by itself - after the source data changes, select Data > Refresh All.

**How it works**

- Click inside the source table and select Insert > PivotTable > New Worksheet.
- Drag fields into Rows, Columns and Values; check the aggregation is Sum, Count or Average as intended.
- Insert a PivotChart from it, and refresh both after the data changes.

**Worked example**

- tblMonthly: Month_Label in Rows and Sum of Hires in Values gives hires for each of 18 months.
- Adding Sector to Columns splits each month by sector - no new formulas needed.
- A line PivotChart of that PivotTable shows the hiring trend.

**Decision guide**

| Use when | Avoid when |
|---|---|
| You need totals, counts or averages by one or two categories. | The value must update live without a refresh - use a formula such as SUMIFS. |
| You want to explore the data by rearranging the summary. | The aggregation is wrong for the measure, e.g. summing a rate or an ID. |

**Practitioner quality lens**

- Aggregation: Sum, Count or Average chosen on purpose.
- Tie-out: Grand total equals a SUM of the source column.
- Fresh: Refreshed after every data change.

**Authoritative references**

- https://support.microsoft.com/en-us/excel/overview-of-pivottables-and-pivotcharts

---


### PivotTables and PivotCharts from a Prompt

A PivotTable summarises a table by the fields you choose - rows, columns, values and filters - and a PivotChart draws that summary and stays linked to it. Describe both in one prompt, for example 'Create a PivotTable with Month in rows and Sum of Hires, then a line PivotChart from it', and Copilot builds them on the sheet you name. Slicers and timelines added later filter them together.

A PivotTable gives every chart a visible, checkable total, and refreshes when the source table grows. Its grand total can be reconciled to a SUM or SUMIFS of the source in seconds. If the first result does not match your intent, add the missing detail - field, aggregation, filter or chart type - rather than starting over.

**How it works**

- Prompt for the PivotTable: source table, rows, columns, values, aggregation, filter and target sheet.
- Prompt for the PivotChart from that PivotTable, with the chart type and labels you need.
- Check the grand total against a SUM or SUMIFS of the source table.

**Worked example**

- 'Create a PivotTable from tblMonthly on Pivot_Trend with Month_Label in rows and Sum of Hires.'
- 'Create a line PivotChart of total hires by month from that PivotTable, with data labels.'
- Check: the 2026 Pivot total is 361 hires = SUMIFS(tblMonthly[Hires], ...).

**Decision guide**

| Use when | Avoid when |
|---|---|
| The data is in a table and the measure has a defined aggregation. | The question, denominator or aggregation has not been agreed yet. |
| You need a summary and chart that update when the data changes. | A static picture from a Chat answer would be pasted into a recurring report. |

**Practitioner quality lens**

- Grain: You know what one Pivot cell represents.
- Tie-out: Pivot grand total equals the source-table control.
- Linked: PivotChart linked to its PivotTable, not a pasted picture.

**Authoritative references**

- https://support.microsoft.com/en-us/excel/copilot/visualize-your-data-with-copilot-in-excel
- https://support.microsoft.com/en-us/excel/overview-of-pivottables-and-pivotcharts
- https://support.microsoft.com/en-us/excel/copilot/get-started-with-copilot-in-excel

---


### Choosing and Refining Charts

Every chart answers one kind of question: comparison across categories uses a column or bar chart, change over time a line chart, a recruitment funnel a funnel or ordered bar chart, and two measures with different units a combo chart with a secondary axis. Follow-up prompts keep the conversation's context, so you can refine a chart one change at a time: sort, title, axis format, colours, labels and alt text.

Copilot will draw whatever you ask for, quickly and polished - but a pie chart of 18 months or a combo chart with S$ and % on one axis can mislead a leadership team. Deciding the question first, then refining with small, specific follow-ups, turns a default chart into a decision visual. Copilot can also create a whole new spreadsheet - such as a tracker with a Gantt chart - from a blank workbook.

**How it works**

- Write the question, choose the chart type that answers it, and name the fields in the prompt.
- Refine with one follow-up per change: sort, finding-style title, axis units, colours, alt text.
- Check the chart's values against its PivotTable or source total after formatting.

**Worked example**

- 'Create a combo chart: Hire_Rate as columns sorted highest first, Cost_per_Hire as a line on a secondary axis.'
- 'Change the title to: Referrals convert best - hire rate by source, Jan-Jun 2026.'
- Blank workbook: 'Create a spreadsheet to plan and track July hiring, with a Gantt chart visual.'

**Decision guide**

| Use when | Avoid when |
|---|---|
| The audience needs one clear message from the visual. | A pie with many slices or a 3-D effect would distort comparison. |
| The chart is right in substance but not yet clear to its audience. | Many edits are bundled into one vague prompt such as 'make it pretty'. |

**Practitioner quality lens**

- Fit: Chart type matches comparison, trend, funnel or two-measure question.
- Message: Title states the finding; axes labelled with units.
- Integrity: Zero-based bars, no hidden categories, values still match the source.

**Authoritative references**

- https://support.microsoft.com/en-us/excel/copilot/visualize-your-data-with-copilot-in-excel
- https://support.microsoft.com/en-us/excel/copilot/data-insights-with-copilot-in-excel
- https://support.microsoft.com/en-us/excel/overview-of-pivottables-and-pivotcharts
- https://www.youtube.com/watch?v=4z4eIaEEH9A

---


### Lab 3 · PivotTables, PivotCharts & Visuals — Build PivotTables, PivotCharts and Recruitment Charts with Copilot

Learning outcome: LO1: use Copilot to build PivotTables, PivotCharts and trend, comparison, funnel and combo charts from prompts, refine them with follow-up prompts and tie each chart to a checked total.

Goal: Produce four decision-ready recruitment charts - hiring trend, hires by sector, the H1 2026 funnel and source effectiveness - each backed by a reconciled summary.

BrightPath's leadership wants to see how hiring is moving and which sectors and sourcing channels perform best. You will ask Copilot to build PivotTables and PivotCharts from 18 months of funnel metrics and a sourcing summary, choose each chart from the question it answers, refine titles, labels, sort order and axes with follow-up prompts, and reconcile every chart to an independent SUMIFS check.

**What you'll build**

BrightPath_Recruitment_Metrics.xlsx with PivotTables and four refined charts on the Charts sheet: a monthly hires line PivotChart, a sorted 2026 hires-by-sector bar PivotChart, a 2026 recruitment funnel and a hire-rate vs cost-per-hire combo chart   (Tools: Excel · Copilot (Allow editing) · PivotTables, PivotCharts · line, bar, funnel and combo charts · BrightPath_Recruitment_Metrics.xlsx.)

**Prerequisites**

- Labs 1-2 completed: you can open a OneDrive workbook, switch Copilot modes and verify with formulas.
- This lab folder's data/BrightPath_Recruitment_Metrics.xlsx.

**Step-by-step**

1. Go to onedrive.com (or your SharePoint site), open or create the folder work-c34 and select Upload > Files. Upload BrightPath_Recruitment_Metrics.xlsx from this lab's data folder, then open it in Excel for the web or select Open in Desktop App. OneDrive gives you AutoSave and version history (recent Excel builds can also run Copilot on local files). Check that Formulas > Calculation Options is Automatic, then open Copilot from the Copilot button on the Home tab or the floating Copilot icon at the bottom right. Read the About_This_Data sheet before you prompt. The Charts sheet is empty: you will create each chart there with a prompt, in a 2 x 2 grid. The data is on Monthly_Metrics and Source_Summary. Each chart answers one of the four questions below.

   ```text
   Tables : tblMonthly (108 rows = 18 months x 6 sectors) | tblSources (7 rows)

   Q1 Trend      : How have total hires moved month by month since Jan 2025?
   Q2 Comparison : Which sector produced the most hires in Jan-Jun 2026?
   Q3 Funnel     : Where do candidates drop out between application and hire in 2026?
   Q4 Efficiency : Which sources convert best, and at what cost per hire?
   ```

2. Choose Allow editing. Ask Copilot to create the trend PivotTable and a line PivotChart on the Charts sheet. Then ask Copilot to check the Pivot grand total against =SUM(tblMonthly[Hires]) and name the peak month.

   ```text
   Create a PivotTable from tblMonthly on a new sheet named Pivot_Trend with Month_Label in rows and Sum of Hires as values. Then create a line PivotChart of total hires by month on the Charts sheet, placed over A1:H18. Format it cleanly: white background, no border or effects, one blue line with markers, data labels above the points, light grey gridlines, axis titles Month and Hires.
   Title: "Monthly hires, Jan 2025 - Jun 2026".

   CHECK (Allow editing):
   In an empty cell on Pivot_Trend, add this formula. Tell me whether it matches the PivotTable grand total, and which month had the most hires:
   =SUM(tblMonthly[Hires])
   ```

3. Ask Copilot to create the 2026 sector comparison chart on the Charts sheet. Then refine it with a follow-up prompt. Ask Copilot to add the SUMIFS check for Technology's 2026 hires.

   ```text
   Create a PivotTable on a new sheet Pivot_Sector from tblMonthly with Sector in rows and Sum of Hires, filtered to months in 2026 only. Then create a bar PivotChart from it on the Charts sheet, placed over J1:Q18, in the same clean style as the first chart.

   FOLLOW-UP: Sort the bars from highest to lowest and change the title to state the leading sector.

   CHECK (Allow editing):
   In an empty cell on Pivot_Sector, add this formula and tell me if it matches Technology's bar:
   =SUMIFS(tblMonthly[Hires],tblMonthly[Sector],"Technology",tblMonthly[Month],">="&DATE(2026,1,1))
   ```

4. Ask Copilot to build a 2026 funnel summary with formulas on its own sheet, then create a funnel chart on the Charts sheet. Ask Copilot to show the formula in each value cell to confirm the summary uses SUMIFS, not typed numbers.

   ```text
   On a new sheet named Funnel_2026, build a small table of 2026 totals for Applications, Screened, Interviews, Offers and Hires from tblMonthly using SUMIFS on Month >= 1 Jan 2026. Add a column showing each stage as a % of Applications. Then create a funnel chart from it with data labels on the Charts sheet, placed over A20:H37, in the same clean style as the first chart.

   FOLLOW-UP: Show the formula in each value cell of the funnel table.
   ```

5. Ask Copilot to add rate columns to tblSources and create a combo chart with two axes on the Charts sheet. Then ask it for the best hire rate and the lowest cost per hire.

   ```text
   In tblSources add Hire_Rate = Hires / Applications (percentage, 1 decimal) and Cost_per_Hire = Sourcing_Spend / Hires (S$, no decimals). Create a combo chart on the Charts sheet:
   Hire_Rate as clustered columns sorted highest first, Cost_per_Hire as a line on a secondary axis.
   Label both axes with units. Place it over J20:Q37 in the same clean style as the first chart.

   FOLLOW-UP: Which source has the highest Hire_Rate, and which has the lowest Cost_per_Hire? Give both values.
   ```

6. Refine the four charts on the Charts sheet with short follow-up prompts, one change at a time. Add alt text to each chart.

   ```text
   Make every chart title state the finding in one sentence.
   Use the same blue colour for all hires series and format numbers with thousands separators.
   Make the four charts on the Charts sheet the same size and line them up in a 2 x 2 grid.
   Add alt text to each chart describing what it shows.
   ```

7. Ask Copilot (Chat only) to review the charts against the four questions, then ask it to log your prompts and checks on Prompt_Log.

   ```text
   CHAT ONLY:
   Review the four charts on the Charts sheet. For each, tell me which question it answers (Q1 Trend, Q2 Comparison, Q3 Funnel, Q4 Efficiency), whether the chart type suits that question, and one improvement. Do not change anything.

   LOG (Allow editing):
   Prompt_Log already lists every prompt in this lab. For each one I ran, fill in What Copilot did and How I verified, and set Decision to Keep. I will tell you any row to change to Revise or Undo.
   ```


**Sample prompts**

| Purpose | Prompt |
|---|---|
| PivotTable | Create a PivotTable from tblMonthly with Month_Label in rows, Sector in columns and Sum of Hires as values. |
| PivotChart | Create a line PivotChart from the Pivot_Trend PivotTable showing total hires by month. Show data labels and gridlines. |
| Bar chart | Create a bar chart of 2026 hires by Sector, sorted from highest to lowest. |
| Funnel | Create a funnel chart of 2026 Applications, Screened, Interviews, Offers and Hires with data labels. |
| Combo chart | Create a combo chart of Hire_Rate as columns and Cost_per_Hire as a line on a secondary axis. |
| Refine | Change the chart title to: Technology led hiring in H1 2026. |
| Refine | Sort the bars from highest to lowest and add data labels formatted with thousands separators. |
| Accessibility | Add alt text to this chart that describes the trend it shows. |
| Review | Is a line chart the right choice for this data? Suggest a better chart if not, and explain why. |

**Test it**

| Check | Expected result |
|---|---|
| Total hires (Pivot_Trend) | 1,065 = SUM(tblMonthly[Hires]) |
| Peak month | 2025-11 with 68 hires |
| 2026 hires by sector | Technology first (84); Pivot total 361 |
| Funnel | 6,616 > 3,229 > 1,385 > 495 > 361 |
| Best source | Referral: highest hire rate (7.4%); MyCareersFuture: lowest cost per hire (S$186) |
| Charts sheet | 2 PivotCharts + funnel + combo in a 2 x 2 grid (A1:Q37), clean style, finding-style titles, labelled axes and alt text |

**Checkpoint and rejoin point**

Keep the workbook in work-c34. Topic 2 starts from the client orders workbook in the Lab 4 folder.

**Troubleshooting**

| If this happens | Fix |
|---|---|
| The 2026 filter includes 2025 months. | Ask Copilot to filter on the Month date field (>= 1 Jan 2026), not on Month_Label text, or add a Year field. |
| Copilot inserts a picture of a chart instead of a chart. | That came from a Chat answer and is static. Switch to Allow editing and ask it to create a chart linked to the data. |
| The funnel chart type is not offered. | Funnel charts need Excel 2019 or Microsoft 365. Ask for a sorted horizontal bar chart of the stages instead. |
| A chart lands on the PivotTable sheet, not on the Charts sheet. | Ask: 'Move this chart to the Charts sheet over A1:H18' (use that chart's range), or select the chart and use Move Chart > Object in: Charts. |
| The chart has a glow or gradient border, or the data labels overlap. | Ask: 'Remove the chart border and effects, use a white background and put the data labels above the points.' |
| Both series in the combo chart share one axis and the line is flat. | Ask: 'Put Cost_per_Hire on a secondary axis and format it as S$'. |

**Challenge**

Ask Copilot for a 100% stacked column chart showing each sector's share of hires by quarter for 2025-2026, then decide whether it communicates better or worse than the bar chart, and why. Then try creating a spreadsheet from scratch: in a new blank workbook ask Copilot 'Create a spreadsheet I can use to plan and track BrightPath's July hiring campaign, with a Gantt chart visual.' Change one start date and check that the Gantt bar moves.

**Reflection**

Which of your four charts would most change a BrightPath leadership decision, and what check makes you trust it?

> **Note:** The lab folder labs/lab-03-*/ holds the instruction PDF with sample prompts, sample-prompts.txt and the mock-data workbook (data/BrightPath_Recruitment_Metrics.xlsx). Use only the supplied fictional BrightPath data or information you are authorised to process. Do not paste credentials, confidential business material or real employees' personal data into Copilot, rules or skills.

---


### Topic 01 Recap - Mapped to Learning Outcomes

Use this checkpoint to explain what you can now do and identify any lab evidence you still need to repair.

| Learning outcome | Evidence from this topic |
|---|---|
| LO1 · Modes | Chat only to ask, Plan to review the steps, Allow editing to change - then Done or Undo. |
| LO1 · Data prep | Named tables, duplicates and spelling variants flagged and fixed, row counts unchanged. |
| LO1 · Formulas | Formulas and XLOOKUPs described in plain English, explained, and hand-checked on one row. |
| LO1 · PivotTables and charts | PivotTables, PivotCharts and refined charts, each tied to a SUM or SUMIFS total. |

---


## Topic 02 — Skills and Rules  (Labs 4-6)

Personalisation · workbook rules on a .Rules sheet · built-in formatting, visualisation and finance skills · custom skills in SKILL.md · using Copilot to write a skill

**Key concepts**

- Personalisation = you — Preferences that follow you to every workbook, e.g. 'Never merge cells'. Settings (...) > Personalization.
- Rules = this workbook — A visible sheet named .Rules with one rule per cell in column A. The rules travel with the file.
- Skills = a repeatable recipe — Saved step-by-step instructions you call by name with @, in any workbook.
- Start with built-in skills — @chart-design for charts, @theme-design for styling, and finance skills such as @unit-economics and @check-model.
- Let Copilot write your skill — Describe the task; Copilot writes the SKILL.md. Save it in its own OneDrive folder and select Refresh.
- Still check the result — Rules and skills guide Copilot. They do not guarantee the numbers are right.

Concepts for Lab 4: Personalisation: Preferences That Follow You, Workbook Rules: The .Rules Sheet.


### Personalisation: Preferences That Follow You

Personalization stores your own working preferences for Copilot in Excel - chart types, colours, number and date formats, layout habits - and applies them in every workbook you open. You add them in the Copilot pane under Settings (...) > Personalization, one short preference per line. They belong to your account, not to the file, so they are not passed on when you share a workbook.

If you keep typing the same corrections - don't merge cells, use column charts, give PivotTable fields proper names - personalisation lets you say it once. Preferences are guidance, not locks: if a prompt asks for something different, the prompt wins. Vague preferences such as 'make it look nice' give Copilot nothing to work with.

**How it works**

- Write down the last three corrections you gave Copilot - that is your first list.
- Add each one on its own line in Settings (...) > Personalization, short and specific, and Save.
- Start a New chat so the preferences apply, then check a new chart or table follows them.

**Worked example**

- 'Use column charts rather than bar charts unless the category labels are long.'
- 'Use navy #1F4E79 and teal #2A9D8F as the main chart colours.'
- 'Show dates as dd-mmm-yyyy and money as S$ with no decimals.'

**Decision guide**

| Use when | Avoid when |
|---|---|
| The preference is about how you like to work, in any workbook. | The standard belongs to one workbook or team - put it in that workbook's .Rules instead. |
| You have corrected Copilot the same way at least twice. | The preference is vague ('make it professional') or conflicts with your organisation's brand kit. |

**Practitioner quality lens**

- Specific: One short, testable preference per line.
- Scoped: Personal habits here; workbook standards in .Rules.
- Refreshed: New chat started after every change.

**Authoritative references**

- https://support.microsoft.com/en-us/excel/copilot/copilot-in-excel-personalization
- https://techcommunity.microsoft.com/blog/excelblog/new-ways-to-customize-how-copilot-edits-your-workbooks/4527307
- https://www.youtube.com/watch?v=4z4eIaEEH9A
- https://www.youtube.com/watch?v=wZ8QAKiIfuo

---


### Workbook Rules: The .Rules Sheet

Workbook rules teach Copilot how to handle one specific workbook: definitions, valid values, calculation methods, formatting and layout conventions, and sheets it must not change. They live on a worksheet named exactly .Rules, one rule per cell in column A, and Copilot reads them every time it works in that file. Select + (Add work content) > Create workbook rules to add the sheet with a template, or ask Copilot to create it.

Rules remove the need to repeat the same context in every prompt, so prompts get short and results get consistent. Because the sheet is part of the file, a colleague who opens the workbook gets the same rules with no setup. The sheet must stay visible - Copilot ignores a hidden .Rules sheet - and rules work best in English, short and specific, with an example.

**How it works**

- Select + > Create workbook rules (or create a sheet named .Rules) and keep it visible.
- Write one rule per cell in column A, group them under ## headings and include an example.
- Start a New chat and give a short prompt; check the result follows each rule.

**Worked example**

- 'Status must be one of: Open, In Progress, On Hold, Completed, Cancelled. WIP = In Progress.'
- 'If Notes say the work is paused, the Status is On Hold, whatever the Status column says.'
- With the rules in place: 'Create Orders_Clean with a clean copy of Orders_Raw, following the workbook rules.'

**Decision guide**

| Use when | Avoid when |
|---|---|
| A workbook has definitions or conventions that every prompt should respect. | The instruction is a personal preference for all your files - use Personalization. |
| Several people use the same workbook and need consistent results. | The process must run across many different workbooks - make it a skill. |

**Practitioner quality lens**

- Clear: One rule per cell, ## headings, examples of good output.
- Visible: Sheet named exactly .Rules and not hidden.
- Tested: A short prompt produces output that follows every rule.

**Authoritative references**

- https://support.microsoft.com/en-us/excel/copilot/copilot-in-excel-rules
- https://techcommunity.microsoft.com/blog/excelblog/new-ways-to-customize-how-copilot-edits-your-workbooks/4527307
- https://www.youtube.com/watch?v=4z4eIaEEH9A
- https://www.youtube.com/watch?v=wZ8QAKiIfuo

---


### Lab 4 · Workbook Rules & Personalisation — Clean Messy Client Orders with Cleanup, Personalisation and Rules

Learning outcome: LO2: set personal preferences, use Copilot's automated cleanup, and write a .Rules worksheet that makes Copilot clean and calculate a messy export the same way every time.

Goal: Turn a messy client-orders export into a clean, consistent table with short prompts, because the workbook's .Rules sheet already holds the standards.

BrightPath's order system exports client orders that several people typed by hand: 'WIP', 'in progress' and 'Ongoing' all mean In Progress, regions are written 23 ways and client names carry 'Pte Ltd' suffixes. You will save your own preferences in Personalization, try Copilot's automated cleanup, then create a .Rules sheet that defines the workbook's standards. With the rules in place, a one-line prompt builds a clean table, and the same rules shape the XLOOKUP and budget columns you add next.

**What you'll build**

Your Personalization (saved to your account) and BrightPath_Client_Orders_Raw.xlsx with a .Rules sheet, a clean tblOrdersClean table (5 statuses, 5 regions, 12 clients), Rate, Delivered_Value, Budget_Used and Over_Budget columns, and a Checks sheet that proves the clean-up   (Tools: Excel · Copilot (Allow editing, Chat only) · Personalization · Clean data with Copilot · .Rules worksheet · XLOOKUP.)

**Prerequisites**

- Topic 1 completed: you can switch Copilot modes, keep or undo changes, and verify with formulas.
- This lab folder's data/BrightPath_Client_Orders_Raw.xlsx and rules-template.txt.

**Step-by-step**

1. Go to onedrive.com (or your SharePoint site), open or create the folder work-c34 and select Upload > Files. Upload BrightPath_Client_Orders_Raw.xlsx from this lab's data folder, then open it in Excel for the web or select Open in Desktop App. OneDrive gives you AutoSave and version history (recent Excel builds can also run Copilot on local files). Check that Formulas > Calculation Options is Automatic, then open Copilot from the Copilot button on the Home tab or the floating Copilot icon at the bottom right. Read the About_This_Data sheet before you prompt. Scroll through Orders_Raw and spot three ways the same status or region is written.

   ```text
   Sheets : About_This_Data | Orders_Raw (64 orders, not a table) | Rate_Card (tblRateCard) | Prompt_Log
   Messy : 20 Status spellings | 23 Region spellings | 51 Client spellings
   Clean : 5 statuses | 5 regions | 12 clients
   ```

2. Personalisation is about you, not this file. In the Copilot pane select Settings (...) > Personalization, type the preferences below one per line, select Save and then start a New chat. They now apply in every workbook you open. A prompt that asks for something different still wins.

   ```text
   Charts      : Use column charts rather than bar charts unless the category labels are long.
   Colours     : Use navy #1F4E79 and teal #2A9D8F as the main chart colours.
   Layout      : Never merge cells; leave one empty column between blocks of content.
   PivotTables : Give fields readable names, e.g. Hires instead of Sum of Hires.
   ```

3. Try automated cleanup. Click inside Orders_Raw and select Data > Clean Data (or Show suggestions if Excel offers 'Clean data with Copilot'). Hover over each suggestion to see the cells it affects, then Apply the extra-space and capitalisation fixes. Then ask Copilot what is still inconsistent: cleanup fixes typing, but it cannot know that WIP means In Progress or that C means CENTRAL - that needs rules.

   ```text
   CHAT ONLY:
   On Orders_Raw, list the distinct values in Status, Region and Client with a count of each. Group the values that probably mean the same thing. Do not change anything.
   ```

4. Create the rules sheet. In the prompt box select + (Add work content) > Create workbook rules. Copilot adds a .Rules sheet with suggested rules. Replace them with BrightPath's rules: open rules-template.txt from this lab folder, copy every line, click .Rules!A1 and paste (each line lands in its own cell in column A), or send the prompt below. Keep the .Rules sheet visible - Copilot ignores a hidden one.

   ```text
   ALLOW EDITING:
   On the .Rules sheet, replace everything in column A with these rules, one per cell from A1 down.
   Keep each ## heading in its own cell. Do not change any other sheet.
   ## ABOUT THIS WORKBOOK
   Orders_Raw is the original export. Do not sort, delete or add rows on Orders_Raw. Build cleaned data on new sheets.
   One row = one client order, identified by Order_ID.
   Put "Created by Copilot" and today's date in cell A1 of every new sheet you create (except the Checks sheet), and start the data in row 3.
   ## DATA VALUES
   Status must be one of: Open, In Progress, On Hold, Completed, Cancelled. Map every variation: WIP and Ongoing = In Progress; Done and Closed - filled = Completed; Not started and New = Open; Paused and Hold = On Hold; Canceled and Client cancelled = Cancelled.
   If Notes say the work is paused, the Status is On Hold, whatever the Status column says.
   Region must be one of: CENTRAL, EAST, NORTH, NORTH-EAST, WEST (upper case). For example C = CENTRAL and Northeast = NORTH-EAST.
   Client is the short trading name with no legal suffix, in title case. For example QUILLON SYSTEMS PTE. LTD. = Quillon Systems.
   Owner names are trimmed and in title case, exactly as in tblRateCard[Owner].
   ## DATES AND MONEY
   Text dates in Orders_Raw are day-first. For example 09/03/2026 is 9 March 2026.
   Show dates as dd-mmm-yyyy, for example 09-Mar-2026.
   Budget is in Singapore dollars. Convert text such as 12.5k, SGD 12,500 and $12,500 to the number 12500.
   Format money as S$ with thousands separators and no decimals, e.g. S$12,500. Show negative values in red in brackets.
   ## TABLES AND FORMULAS
   Format every new data range as an Excel table named tbl + the sheet name without underscores, e.g. tblOrdersClean.
   Use XLOOKUP for lookups and return "Not found" when there is no match.
   Never merge cells.
   ```

5. Apply the rules with a one-line prompt in a New chat. Copilot reads .Rules before it works, so you do not repeat the standards. Inspect the result: A1 says Created by Copilot, the table is tblOrdersClean, ORD-302 and ORD-303 are On Hold (their notes say the work is paused), dates read like 09-Mar-2026 and budgets are numbers formatted as S$. Select Done to keep it.

   ```text
   ALLOW EDITING:
   Create a new sheet Orders_Clean with a clean copy of every order on Orders_Raw, following the workbook rules.
   ```

6. Prove the clean-up. Ask Copilot to add independent checks on a Checks sheet. Each result must equal the expected value in Test It.

   ```text
   CHECK (Allow editing):
   Create a sheet named Checks. Enter these labels and formulas exactly as written, then tell me each result:
   A1 Orders              B1 =ROWS(tblOrdersClean[Order_ID])
   A2 Distinct statuses   B2 =COUNTA(UNIQUE(tblOrdersClean[Status]))
   A3 Distinct regions    B3 =COUNTA(UNIQUE(tblOrdersClean[Region]))
   A4 Distinct clients    B4 =COUNTA(UNIQUE(tblOrdersClean[Client]))
   A5 In Progress         B5 =COUNTIF(tblOrdersClean[Status],"In Progress")
   A6 Total budget        B6 =SUM(tblOrdersClean[Budget])
   ```

7. Add calculated columns. The rules already say to use XLOOKUP with 'Not found' and how to format money, so the prompt only states the business logic. Check one row by hand: ORD-326 is 358 hours x S$115 (Priya Nair) = S$41,170 against a budget of S$36,500, so it is over budget.

   ```text
   ALLOW EDITING:
   In tblOrdersClean add: Rate from tblRateCard by Owner; Delivered_Value = Hours_Logged x Rate; Budget_Used = Delivered_Value / Budget as a percentage with 1 decimal; Over_Budget = "Yes" when Delivered_Value is greater than Budget, otherwise "No".

   FOLLOW-UP:
   On the Checks sheet, add the label Total delivered value in A7 with =SUM(tblOrdersClean[Delivered_Value]) in B7, and the label Over budget in A8 with =COUNTIF(tblOrdersClean[Over_Budget],"Yes") in B8.
   Then list the over-budget orders.
   ```

8. Add a rule, then reuse it. Rules can grow: ask Copilot to add one, then apply it with a short prompt. Rules are saved in the file, so a colleague who opens this workbook gets them too. Log your prompts.

   ```text
   ALLOW EDITING:
   Add a rule under ## TABLES AND FORMULAS on the .Rules sheet: Highlight rows where Over_Budget is Yes in light red.

   ALLOW EDITING:
   Apply the workbook rules to tblOrdersClean.

   LOG (Allow editing):
   Prompt_Log already lists every prompt in this lab. For each one I ran, fill in What Copilot did and How I verified, and set Decision to Keep. I will tell you any row to change to Revise or Undo.
   ```


**Sample prompts**

| Purpose | Prompt |
|---|---|
| Create rules | Create a .Rules worksheet based on the formatting already used in this workbook. |
| Add a rule | Add a rule to my .Rules worksheet that all charts should use navy #1F4E79 and teal #2A9D8F. |
| Apply rules | Create a new sheet Orders_Clean with a clean copy of Orders_Raw, following the workbook rules. |
| Find variants | List the distinct values in Status and Region with counts, and group the ones that mean the same thing. |
| Rules + lookup | Add Rate from tblRateCard by Owner and Delivered_Value = Hours_Logged x Rate to tblOrdersClean. |
| Question | Which client has the largest total budget, and how many orders does it have? |
| Explain | Which rules on the .Rules sheet did you follow to clean the Status column? Give two examples. |
| New sheet | Create a new sheet showing orders with a budget over S$50,000 by Owner. |

**Test it**

| Check | Expected result |
|---|---|
| Orders and distinct values (Checks B1:B4) | 64 orders; 5 statuses, 5 regions, 12 clients |
| Status counts | Open 10, In Progress 24, On Hold 6, Completed 21, Cancelled 3 |
| Paused orders | ORD-302 and ORD-303 are On Hold |
| Total budget (B6) | S$3,250,500 |
| Delivered value and over budget (B7, B8) | S$1,961,030; 10 orders over budget |
| Sample ORD-326 | 358 h x S$115 = S$41,170 vs budget S$36,500 - Over_Budget Yes |

**Checkpoint and rejoin point**

Keep the workbook in work-c34. Your Personalization stays with your account for the rest of the course.

**Troubleshooting**

| If this happens | Fix |
|---|---|
| Copilot ignores the rules. | The sheet must be named exactly .Rules (with the full stop), be visible, and have the rules in column A. Start a New chat after changing it. |
| Create workbook rules is not in the + menu. | Create a sheet named .Rules yourself and paste rules-template.txt into A1, or ask Copilot: 'Create a .Rules worksheet with these rules ...'. |
| There is no Clean Data button. | It is rolling out and may be missing in your build. Skip Step 3 - the rules in Steps 4-5 do the same job. |
| ORD-302 or ORD-303 still shows Open. | Copilot missed the Notes rule. Ask: 'Apply the rule about paused work in Notes to tblOrdersClean.' |

**Challenge**

Make a rule respond to the workbook: put a dropdown in .Rules!B1 with Executive and Analyst, and in A20 enter =IF(B1="Executive","New summary sheets show totals only","New summary sheets include a detail table"). Switch the dropdown and ask for an order summary by Region each time.

**Reflection**

Which three corrections do you give Copilot most often? Which belong in Personalization, and which in a workbook's .Rules?

> **Note:** The lab folder labs/lab-04-*/ holds the instruction PDF with sample prompts, sample-prompts.txt and the mock-data workbook (data/BrightPath_Client_Orders_Raw.xlsx), plus rules-template.txt (the .Rules text to paste into column A). Use only the supplied fictional BrightPath data or information you are authorised to process. Do not paste credentials, confidential business material or real employees' personal data into Copilot, rules or skills.

---

Concepts for Lab 5: Key Finance Measures: CAC, LTV, Churn and Payback, Built-in Skills: Formatting, Visualisation and Finance.


### Key Finance Measures: CAC, LTV, Churn and Payback

A subscription business such as BrightPath's HR outsourcing is judged on a few measures: MRR (monthly recurring revenue); churn (the share of clients lost each month); ARPA (average revenue per account = MRR / active clients); CAC (customer acquisition cost = sales and marketing spend / new clients); LTV (lifetime value = ARPA x gross margin / monthly churn); LTV:CAC; and payback months (CAC / (ARPA x gross margin)).

Together they answer one question: does winning a client pay off, and how fast? A common benchmark is an LTV:CAC of 3 or more and a payback within about 12 months. Each measure involves choices - which period, average or end-of-period clients, gross margin or revenue - so two analysts (or two skills) can disagree unless the definitions are written down first.

**How it works**

- Write each definition and the period (e.g. H1 2026) before calculating anything.
- Calculate the inputs with formulas: new clients, spend, MRR, active clients, margin, churn.
- Combine them into CAC, LTV, LTV:CAC and payback, and compare by channel.

**Worked example**

- H1 2026: S$560,420 spend / 38 new clients = CAC of S$14,748.
- ARPA S$3,681 x 44.9% margin / 2.06% churn = LTV of about S$80,433.
- LTV:CAC = 5.5 and payback = 8.9 months - healthy by the common benchmark.

**Decision guide**

| Use when | Avoid when |
|---|---|
| Deciding where to spend on marketing or which clients to win. | The period or definitions are not agreed - numbers will not be comparable. |
| Checking a finance skill's output against your own definitions. | Very few new clients in a channel - its CAC will swing wildly. |

**Practitioner quality lens**

- Defined: Formula and period written for each measure.
- Consistent: Same definitions every month and every channel.
- Reproduced: Headline numbers rebuilt with your own formulas.

**Authoritative references**

- https://support.microsoft.com/en-us/excel/copilot/copilot-in-excel-built-in-skills
- https://www.microsoft.com/en-us/copilot/blog/2026/06/25/copilot-in-excel-built-for-the-era-of-frontier-finance/

---


### Built-in Skills: Formatting, Visualisation and Finance

A skill gives Copilot ready-made instructions for a repeatable task. Copilot in Excel ships with built-in skills that you manage in Settings (...) > Manage plugins & skills: Formatting (@brandkit, @theme-design), Visualization (@chart-design) and a Finance library of about 20 skills such as @unit-economics, @check-model, @3-statements and @comps-analysis. Call one by typing @ and its name, or choose it from + > Choose plugins & skills.

Built-in skills package expert know-how: @chart-design applies the Excel design team's chart guidance, @check-model audits a model for hardcoded overrides and broken formulas, and @unit-economics works out CAC, LTV and payback. They save hours - but a skill uses its own assumptions, so you still hold it to your definitions and check the headline numbers. Skills need Excel's display language set to English.

**How it works**

- Open Settings (...) > Manage plugins & skills and turn on the skills you need (chart-design is off by default).
- Type @ and the skill name in the prompt, plus the table, period and definitions to use.
- Check the output against independent formulas; ask the skill to show its working if they differ.

**Worked example**

- '@chart-design Create a combo chart of MRR and active clients by month from tblClientMonths.'
- '@check-model Audit Headcount_Budget for typed numbers, odd formulas and totals that miss rows.'
- '@theme-design Apply one professional theme to the Charts and Unit_Economics sheets.'

**Decision guide**

| Use when | Avoid when |
|---|---|
| A built-in skill matches the task (styling, charting, a standard finance analysis or model audit). | The skill's default definitions differ from yours and you have not said which to use. |
| You can state the definitions and check the key results. | Nobody on the team can judge whether the finance output is sensible. |

**Practitioner quality lens**

- Right skill: Chosen for the task; enabled or called directly with @.
- Your definitions: Period and formulas stated in the prompt or .Rules.
- Verified: Headline numbers reproduced with native formulas.

**Authoritative references**

- https://support.microsoft.com/en-us/excel/copilot/copilot-in-excel-skills
- https://support.microsoft.com/en-us/excel/copilot/copilot-in-excel-built-in-skills
- https://www.microsoft.com/en-us/copilot/blog/2026/06/25/copilot-in-excel-built-for-the-era-of-frontier-finance/
- https://www.youtube.com/watch?v=4z4eIaEEH9A
- https://www.youtube.com/watch?v=wZ8QAKiIfuo

---


### Lab 5 · Built-in Skills: Visual & Finance — Analyse HR-Outsourcing Finances with Built-in Skills

Learning outcome: LO2: find, enable and call Copilot's built-in skills - chart-design, unit-economics, check-model and theme-design - and verify what each skill produces.

Goal: Use four built-in skills to chart BrightPath's recurring revenue, calculate client unit economics, audit a headcount budget model and style the results - then prove the numbers.

BrightPath's HR-outsourcing business sells monthly payroll and HR services to client companies. Finance wants to know whether winning clients pays off and whether next year's headcount budget is right. You will switch on skills in Manage plugins & skills, call them with @, and let each one do a specialist job: @chart-design builds the revenue chart, @unit-economics works out CAC, LTV and payback, @check-model finds the errors in the budget model, and @theme-design styles the workbook. A Checks sheet proves every headline number.

**What you'll build**

BrightPath_HR_Outsourcing_Finance.xlsx with a skill-designed MRR chart, a Unit_Economics sheet, a corrected Headcount_Budget, a Checks sheet that agrees with the skills, and a consistent theme   (Tools: Excel · Copilot (Allow editing, Plan, Chat only) · @chart-design · @unit-economics · @check-model · @theme-design.)

**Prerequisites**

- Lab 4 completed: you know how rules guide Copilot inside one workbook.
- This lab folder's data/BrightPath_HR_Outsourcing_Finance.xlsx. Skills need Excel's display language set to English.

**Step-by-step**

1. Go to onedrive.com (or your SharePoint site), open or create the folder work-c34 and select Upload > Files. Upload BrightPath_HR_Outsourcing_Finance.xlsx from this lab's data folder, then open it in Excel for the web or select Open in Desktop App. OneDrive gives you AutoSave and version history (recent Excel builds can also run Copilot on local files). Check that Formulas > Calculation Options is Automatic, then open Copilot from the Copilot button on the Home tab or the floating Copilot icon at the bottom right. Read the About_This_Data sheet before you prompt. Read the Definitions row on About_This_Data - you will hold the finance skill to these.

   ```text
   Tables : tblClientMonths (18 months) | tblClients | tblSpend | Headcount_Budget (cost model)
   H1 2026 = Jan-Jun 2026
   CAC     = sales & marketing spend / new clients
   ARPA    = MRR / active clients at Jun 2026
   LTV     = ARPA x gross margin / monthly churn
   Payback = CAC / (ARPA x gross margin), in months
   ```

2. Find the skills. In the Copilot pane select Settings (...) > Manage plugins & skills and open each folder. Turn on chart-design (it is off by default), unit-economics and check-model. Close the dialog, type @ in the prompt box and confirm they are listed. A skill that is off can still be called with an @mention.

   ```text
   Formatting    : brandkit, theme-design
   Visualization : chart-design
   Finance       : 3-statements, check-model, comps-analysis, unit-economics and more (20 skills)
   Call a skill  : type @ and its name, e.g. @chart-design

   CHAT ONLY:
   Which of your skills would help me analyse client acquisition cost and churn in this workbook, and which would help me audit the Headcount_Budget sheet? Do not change anything.
   ```

3. Call @chart-design (Visualization) in Allow editing. It chooses and formats the chart using the Excel design team's guidance. Check the values: MRR grows from S$238,200 in Jan 2025 to S$482,200 in Jun 2026, with 131 active clients at the end.

   ```text
   ALLOW EDITING:
   @chart-design Create a combo chart on a new sheet Charts from tblClientMonths: MRR as columns and Clients_End as a line on a secondary axis, Jan 2025 to Jun 2026.

   FOLLOW-UP:
   @chart-design Improve this chart for a leadership slide: a title that states the finding, axis titles with units, and a data label on the last month only.
   ```

4. Call the finance skill @unit-economics in Plan mode, so you can review its approach before it builds. Check that the plan uses H1 2026 and the About_This_Data definitions; reply with corrections if not, then approve and select Done.

   ```text
   PLAN MODE:
   @unit-economics Analyse the unit economics of BrightPath's HR-outsourcing clients for H1 2026 using tblClientMonths, tblClients and tblSpend. Use the definitions on About_This_Data. Put the results on a new sheet Unit_Economics: CAC, ARPA, gross margin, monthly churn, LTV, LTV:CAC and payback months, plus CAC by Channel. List your assumptions.
   ```

5. Never trust a skill's numbers without a check. Ask Copilot to add these independent formulas, then compare them with the Unit_Economics sheet. If they differ, ask the skill to show its calculation and use the definition.

   ```text
   CHECK (Allow editing):
   Create a sheet named Checks. Enter these labels and formulas exactly as written, then tell me each result:
   A1 New clients H1      B1 =SUMIFS(tblClientMonths[New_Clients],tblClientMonths[Month],">="&DATE(2026,1,1))
   A2 S&M spend H1        B2 =SUMIFS(tblClientMonths[Sales_Marketing_Spend],tblClientMonths[Month],">="&DATE(2026,1,1))
   A3 CAC                 B3 =B2/B1
   A4 ARPA                B4 =INDEX(tblClientMonths[MRR],18)/INDEX(tblClientMonths[Clients_End],18)
   A5 Gross margin        B5 =AVERAGEIFS(tblClientMonths[Gross_Margin_Pct],tblClientMonths[Month],">="&DATE(2026,1,1))
   A6 Monthly churn       B6 =AVERAGE(FILTER(tblClientMonths[Churned_Clients]/tblClientMonths[Clients_Start],tblClientMonths[Month]>=DATE(2026,1,1)))
   A7 LTV                 B7 =B4*B5/B6
   A8 LTV to CAC          B8 =B7/B3
   A9 Payback months      B9 =B3/(B4*B5)

   FOLLOW-UP:
   Compare Checks!B3:B9 with the Unit_Economics sheet. Show any value that differs and why.
   ```

6. Audit a model with @check-model. Ask for the findings first in Chat only, read them, then switch to Allow editing and ask for the fixes. It should find three problems: a typed number in E3, an 11-month formula in E6 and a total in E10 that misses the last role.

   ```text
   CHAT ONLY:
   @check-model Audit the Headcount_Budget sheet: look for typed numbers in formula columns, formulas that differ from their neighbours, and totals that miss rows. List each issue with its cell, what is wrong and the fix. Do not change anything yet.

   ALLOW EDITING:
   Fix the issues you found on Headcount_Budget and tell me the new Total FY2027 cost.
   ```

7. Style the workbook with @theme-design (Formatting). It applies one coordinated theme to cells, tables and charts. If your organisation has set up a brand kit, try @brandkit instead. Then log your prompts.

   ```text
   ALLOW EDITING:
   @theme-design Apply one clean, professional theme to the Charts, Unit_Economics, Checks and Headcount_Budget sheets.

   LOG (Allow editing):
   Prompt_Log already lists every prompt in this lab. For each one I ran, fill in What Copilot did and How I verified, and set Decision to Keep. I will tell you any row to change to Revise or Undo.
   ```


**Sample prompts**

| Purpose | Prompt |
|---|---|
| Find skills | Which skills do you have for charts, formatting and finance? When would I use each one? |
| Chart skill | @chart-design Recommend the best chart for new vs churned clients by month, then create it. |
| Finance skill | @unit-economics Compare CAC and payback months by acquisition Channel for H1 2026. |
| Audit skill | @check-model Check this sheet for hardcoded numbers, broken formulas and totals that miss rows. |
| Style skill | @theme-design Apply a coordinated theme to this workbook. |
| Brand skill | @brandkit Apply our organisation's brand colours and fonts to the charts. |
| Explain | Explain how you calculated LTV, step by step, with the cells you used. |
| Question | Which acquisition channel has the lowest CAC in H1 2026, and how many clients did it bring? |

**Test it**

| Check | Expected result |
|---|---|
| New clients and spend, H1 2026 | 38 clients; S$560,420 |
| CAC and ARPA | S$14,748; S$3,681 |
| Margin and monthly churn | 44.9%; 2.06% |
| LTV, LTV:CAC, payback | About S$80,433; 5.5; 8.9 months |
| CAC by channel | Lowest Referral S$5,180; highest Events S$33,445 |
| @check-model fixes | E3, E6, E10: total S$1,437,934 -> S$1,703,052 |

**Checkpoint and rejoin point**

Keep the workbook in work-c34. Lab 6 needs your OneDrive, because custom skills are stored there.

**Troubleshooting**

| If this happens | Fix |
|---|---|
| Manage plugins & skills is missing or greyed out. | Wait 10 seconds - it loads after the pane. Then check Excel is up to date and the Office display language is English. |
| A skill is not listed when I type @. | Turn it on in Manage plugins & skills, or type its full name: an @mention works even when the skill is off. |
| The skill's CAC or LTV differs from Checks. | Skills use their own defaults. Ask: 'Use the definitions on About_This_Data and recalculate CAC and LTV.' |
| @check-model changes the sheet straight away. | You were in Allow editing. Select Undo, switch to Chat only and ask for the findings first. |

**Challenge**

Ask @unit-economics for a client cohort table by Sign_Date quarter showing how many clients from each cohort are still active. Check one cohort with COUNTIFS on tblClients.

**Reflection**

Which skill saved you the most time, and which of its numbers would you not have trusted without the Checks sheet?

> **Note:** The lab folder labs/lab-05-*/ holds the instruction PDF with sample prompts, sample-prompts.txt and the mock-data workbook (data/BrightPath_HR_Outsourcing_Finance.xlsx). Use only the supplied fictional BrightPath data or information you are authorised to process. Do not paste credentials, confidential business material or real employees' personal data into Copilot, rules or skills.

---

Concepts for Lab 6: Custom Skills: The SKILL.md Recipe, Using Copilot to Write and Improve a Skill.


### Custom Skills: The SKILL.md Recipe

A custom skill is a plain-text file named SKILL.md inside its own folder in your OneDrive skills folder. The file starts with front matter between --- lines: a name, which must match the folder name, and a description that tells Copilot when to use the skill. Below it, the body lists the steps in plain English - no code needed. Create the folder from Manage plugins & skills > Custom skills > Create OneDrive folder.

Rules belong to one workbook; a skill is a recipe you can run in any workbook. If you ask Copilot for the same report every week, a skill turns a long prompt into '@weekly-report' and gives the same instructions every time. Your OneDrive comes with a sample calendar skill to copy, and after you add or edit a skill you select Refresh so Copilot sees the change.

**How it works**

- Create the skills folder, then a sub-folder named after the skill, e.g. interview-calendar.
- Save SKILL.md in it: front matter (name, description), then numbered steps and a final check.
- In Excel select Custom skills > Refresh, turn the skill on and call it with @.

**Worked example**

- Folder: skills/interview-calendar/SKILL.md, with name: interview-calendar.
- Description: 'Use when I ask for an interview calendar, schedule or clashes for a week.'
- Steps: read tblInterviews, build a Mon-Fri grid, colour by Mode, flag clashes, add a summary.

**Decision guide**

| Use when | Avoid when |
|---|---|
| You repeat the same multi-step task on new data every week or month. | It is a one-off task - a good prompt is enough. |
| The steps can be written down clearly enough for a new colleague to follow. | The folder name and the name: line differ - Copilot skips the skill. |

**Practitioner quality lens**

- Findable: Folder name = name: line; description says when to use it.
- Followable: Numbered steps, exact sheet and column names, no guessing.
- Self-checking: Ends with checks Copilot must run and report.

**Authoritative references**

- https://support.microsoft.com/en-us/excel/copilot/copilot-in-excel-skills
- https://agentskills.io/specification
- https://www.youtube.com/watch?v=4z4eIaEEH9A
- https://www.youtube.com/watch?v=wZ8QAKiIfuo

---


### Using Copilot to Write and Improve a Skill

You do not have to write SKILL.md yourself. In Chat only, describe the task, the input table, the output you want and the checks, and ask Copilot to return a complete SKILL.md with front matter. You copy the reply into a plain-text file, save it in the skill's OneDrive folder, refresh, and call it. To improve the skill, ask Copilot to update the file and save the new version over the old one.

Copilot knows the SKILL.md format and writes clear, numbered steps, so a non-technical user can build a skill in minutes. Treat the skill as a living recipe: each time you correct its output twice, add the correction to the steps. Skills run inside what Excel normally allows, but test a new skill on a copy of the workbook before you rely on it.

**How it works**

- Chat only: give the skill name, description, input, numbered steps, output and checks; ask for the full file.
- Save it as SKILL.md (plain text) in skills/<name>/, select Refresh and turn it on.
- Run it on a copy, check the result, then ask Copilot to update the file with any fixes.

**Worked example**

- 'Write a SKILL.md named interview-calendar that turns tblInterviews into a weekly calendar and flags clashes.'
- '@interview-calendar Week starting 13 July 2026.' builds next week's calendar from one line.
- 'Update the skill so the summary shows interview hours per Interviewer.' - save, refresh, rerun.

**Decision guide**

| Use when | Avoid when |
|---|---|
| You can describe the output and the checks precisely. | You would save Copilot's draft without reading every step. |
| Someone will own the skill and keep it up to date. | The skill would process data your organisation has not approved for Copilot. |

**Practitioner quality lens**

- Reviewed: Every step read before saving.
- Tested: Run on a copy and checked against a known answer.
- Owned: A named owner improves it after repeated corrections.

**Authoritative references**

- https://support.microsoft.com/en-us/excel/copilot/copilot-in-excel-skills
- https://support.microsoft.com/en-us/excel/copilot/frequently-asked-questions-about-copilot-in-excel
- https://www.youtube.com/watch?v=4z4eIaEEH9A
- https://www.youtube.com/watch?v=wZ8QAKiIfuo

---


### Lab 6 · Custom Skills: Calendar — Use Copilot to Build a Custom Interview-Calendar Skill

Learning outcome: LO2: use Copilot to write a custom SKILL.md, save it in the OneDrive skills folder, call it with @, and improve it so a weekly task runs from a one-line prompt.

Goal: Create an interview-calendar skill that turns BrightPath's interview list into a weekly calendar with clashes flagged - and reuse it for the next week with one short prompt.

Every Friday a BrightPath coordinator turns next week's interview bookings into a calendar and hunts for double-bookings. That is a repeatable recipe, so it should be a skill. You will open the sample calendar skill in your OneDrive skills folder, ask Copilot to write a new SKILL.md for interview-calendar, save it in its own folder, refresh the skills list and call it with @. The same one-line prompt then builds the next week's calendar.

**What you'll build**

An interview-calendar skill in your OneDrive skills folder, and BrightPath_Interview_Schedule.xlsx with Cal_2026-07-06 and Cal_2026-07-13 calendar sheets, each with clashes flagged and a summary   (Tools: Excel · Copilot (Chat only, Allow editing) · Manage plugins & skills · OneDrive skills folder · SKILL.md.)

**Prerequisites**

- Lab 5 completed: you can find, enable and @mention skills.
- Access to your OneDrive in a browser (custom skills are stored in a OneDrive folder).
- This lab folder's data workbook and skills/interview-calendar/SKILL.md (fallback).

**Step-by-step**

1. Go to onedrive.com (or your SharePoint site), open or create the folder work-c34 and select Upload > Files. Upload BrightPath_Interview_Schedule.xlsx from this lab's data folder, then open it in Excel for the web or select Open in Desktop App. OneDrive gives you AutoSave and version history (recent Excel builds can also run Copilot on local files). Check that Formulas > Calculation Options is Automatic, then open Copilot from the Copilot button on the Home tab or the floating Copilot icon at the bottom right. Read the About_This_Data sheet before you prompt. For this lab the OneDrive copy matters: Copilot finds custom skills in your OneDrive.

   ```text
   Table  : tblInterviews (41 interviews, 6-17 July 2026)
   Week 1 : Mon 6 July - Fri 10 July 2026
   Week 2 : Mon 13 July - Fri 17 July 2026
   Clash  : same Interviewer overlapping, or two Onsite interviews in the same Room overlapping
   ```

2. Create your skills folder. In the Copilot pane select Settings (...) > Manage plugins & skills > Custom skills > Create OneDrive folder, then Open skills folder. OneDrive opens a folder named skills, usually under Documents > Copilot > Microsoft Excel. Open the sample calendar skill folder (e.g. make-calendar.example) and read its SKILL.md: a front matter block between --- lines, then plain-English steps.

   ```text
   Folder      : skills/<skill-name>/   (folder name = the skill name, lower case with hyphens)
   File        : SKILL.md
   Front matter: name (the skill name) and description (when Copilot should use it)
   Body        : the steps Copilot follows - plain English, no code needed
   ```

3. Let Copilot write the skill. In Excel, choose Chat only and send the prompt below. Read the reply: it should start with --- name: interview-calendar, have a description that says when to use it, and list clear steps. Ask for changes until you are happy.

   ```text
   CHAT ONLY:
   Write a Copilot in Excel custom skill as a complete SKILL.md file.
   Name: interview-calendar
   Description: Use when I ask for an interview calendar, schedule or clashes for a week.
   The skill must:
   1) Read tblInterviews (Interview_ID, Role, Interviewer, Interview_Date, Start_Time, Duration_Min, Mode, Room).
   2) Ask for the week start date (a Monday) if I do not give one.
   3) Create a sheet Cal_<yyyy-mm-dd> with Monday to Friday as columns and 30-minute rows from 09:00 to 17:30.
   4) In each slot show Interviewer - Role (Mode); colour Onsite light blue and Video light green.
   5) Flag clashes: overlapping interviews with the same Interviewer, or Onsite interviews in the same Room.
      Colour them red and list them in a Clashes table under the calendar.
   6) Add a summary of interviews per day and per Interviewer.
   7) Never change tblInterviews.
   Return only the SKILL.md text, starting with the --- front matter.
   ```

4. Save the skill. Copy Copilot's reply into a plain-text editor (Notepad, or TextEdit in plain-text mode) and save it as SKILL.md. In the OneDrive skills folder select New > Folder, name it interview-calendar, open it and Upload SKILL.md. Short of time? Upload skills/interview-calendar/SKILL.md from this lab folder instead.

   ```text
   OneDrive : .../skills/interview-calendar/SKILL.md
   Check    : the folder name and the name: line are both exactly interview-calendar
   Check    : the file is called SKILL.md (not SKILL.md.txt)
   ```

5. Make Excel see the new skill: Settings (...) > Manage plugins & skills > Custom skills > Refresh, then turn interview-calendar on. Close the dialog and type @inter in the prompt box - the skill should appear.
6. Run the skill in Allow editing with a one-line prompt. When it finishes, inspect the calendar and select Done. Then ask Copilot to add the check formula and compare it with the number of interviews on the calendar.

   ```text
   ALLOW EDITING:
   @interview-calendar Build the calendar for the week starting Monday 6 July 2026.

   CHECK (Allow editing):
   In an empty cell to the right of the calendar, add this formula and tell me whether it matches the number of interviews on the calendar:
   =COUNTIFS(tblInterviews[Interview_Date],">="&DATE(2026,7,6),tblInterviews[Interview_Date],"<="&DATE(2026,7,10))
   ```

7. Reuse it. This is the point of a skill: next week's calendar needs only the skill name and a date. Week 2 has no clashes - check the skill says so.

   ```text
   ALLOW EDITING:
   @interview-calendar Week starting 13 July 2026.
   ```

8. Improve the skill. Ask Copilot to update the file, save it over SKILL.md in OneDrive, select Refresh in Custom skills and rerun week 1. A skill is a living recipe: improve it whenever you correct Copilot twice. Log your prompts.

   ```text
   CHAT ONLY:
   Update the interview-calendar SKILL.md so the summary also shows total interview hours per Interviewer, and the reply lists the busiest day. Return the full updated file.

   LOG (Allow editing):
   Prompt_Log already lists every prompt in this lab. For each one I ran, fill in What Copilot did and How I verified, and set Decision to Keep. I will tell you any row to change to Revise or Undo.
   ```


**Sample prompts**

| Purpose | Prompt |
|---|---|
| Write a skill | Write a SKILL.md for a skill named weekly-recruiter-report that summarises interviews by Interviewer and Client for a given week, with a small bar chart. |
| Explain a skill | Read this SKILL.md and tell me what Copilot will do, step by step, and anything that is unclear. |
| Call a skill | @interview-calendar Build the calendar for the week starting Monday 6 July 2026. |
| Reuse | @interview-calendar Week starting 13 July 2026. |
| Clash check | List every interview clash in tblInterviews for 6-10 July 2026 with both Interview_IDs. |
| Improve | Update the skill so Video interviews show the Teams icon text [Teams] and the calendar prints on one A4 page. |
| Question | Which interviewer has the most interviews in the week of 6 July, and on which day? |

**Test it**

| Check | Expected result |
|---|---|
| Week of 6 July | 23 interviews = the COUNTIFS check |
| By day | Mon 7, Tue 5, Wed 5, Thu 3, Fri 3 |
| Clashes flagged | INT-010/INT-011 Kevin Ong; INT-015/INT-016 Orchid Room; INT-018/INT-019 Aisha Rahman |
| Busiest interviewer | Aisha Rahman (9 interviews) |
| Week of 13 July | 18 interviews, no clashes |
| Skill | interview-calendar appears when you type @ and runs from a one-line prompt |

**Checkpoint and rejoin point**

Keep the skill: it stays in your OneDrive and works in any workbook with an interviews table. Topic 3 uses rules and skills to build dashboards.

**Troubleshooting**

| If this happens | Fix |
|---|---|
| The skill does not appear after Refresh. | The folder name must match the name: line exactly (lower case, hyphens) and the file must be SKILL.md inside that folder. Select Refresh again. |
| There is no Create OneDrive folder button. | Custom skills are rolling out and depend on your licence and update channel. Use the ready-made SKILL.md to read the format, and run the same steps as a normal prompt instead. |
| The calendar misses a clash. | Ask: 'Check again for overlapping interviews with the same Interviewer or the same Onsite Room on each day.' Then add that wording to the skill's steps. |
| The file saved as SKILL.md.txt or with formatting. | Save as plain text with the exact name SKILL.md. In TextEdit choose Format > Make Plain Text first. |

**Challenge**

Ask Copilot to write a second skill, weekly-recruiter-report, that summarises interviews by Interviewer and Client for a given week with a small bar chart. Save it, refresh and call it.

**Reflection**

Which weekly or monthly task in your own job would make the best first custom skill, and what would its description say?

> **Note:** The lab folder labs/lab-06-*/ holds the instruction PDF with sample prompts, sample-prompts.txt and the mock-data workbook (data/BrightPath_Interview_Schedule.xlsx), plus skills/interview-calendar/SKILL.md (ready-made skill to upload if you run short of time). Use only the supplied fictional BrightPath data or information you are authorised to process. Do not paste credentials, confidential business material or real employees' personal data into Copilot, rules or skills.

---


### Topic 02 Recap - Mapped to Learning Outcomes

Use this checkpoint to explain what you can now do and identify any lab evidence you still need to repair.

| Learning outcome | Evidence from this topic |
|---|---|
| LO2 · Personalisation | Your own preferences, saved once in Settings and applied in every workbook. |
| LO2 · .Rules | One workbook's standards in a visible .Rules sheet, so a one-line prompt cleans the data. |
| LO2 · Built-in skills | @chart-design, @unit-economics, @check-model and @theme-design - checked against formulas. |
| LO2 · Custom skills | A SKILL.md written with Copilot, saved in OneDrive and reused with one @mention. |

---


## Topic 03 — Interactive Dashboards with Copilot, Skills and Rules  (Labs 7-8)

Designing a highly visual dashboard · KPI definitions and style in .Rules · plan, build and polish with skills · slicers, timelines and refresh · a reusable dashboard skill · sharing people dashboards responsibly

**Key concepts**

- Define each KPI first — Write the formula in .Rules, e.g. Cost per hire = total sourcing cost / hires.
- Put the standards in .Rules — KPI definitions, layout and colours in the rules keep the build prompt short.
- Plan, then build — Plan mode shows the dashboard plan. Fix a wrong KPI before Copilot copies it into five charts.
- Make it visual — KPI cards on top, charts in a grid, traffic-light colours, then @chart-design and @theme-design to polish.
- Make it interactive — Slicers and a timeline filter every PivotTable and chart at once, and the KPI cards must follow.
- Make it repeatable and safe — Save the recipe as a skill, test refresh with new rows, and hide groups under 5 people.

Concepts for Lab 7: Designing a Highly Visual Dashboard, Plan and Build with Rules and Skills, Interactivity: Slicers, Timelines and KPI Cards.


### Designing a Highly Visual Dashboard

A dashboard is one page that answers a leader's recurring questions at a glance. A highly visual design puts four to six KPI cards across the top, three or four charts in a grid below, slicers and a timeline at the side, and the data source and date in the header. Each KPI is a defined measure with a formula, a period and a target, and colour is used to signal status, not to decorate.

Two people can compute 'offer acceptance rate' differently and both look right, so the definitions come first. Visual hierarchy, consistent colours and finding-style titles let a leader read the page in seconds; traffic-light colours on a target card show where to act. Writing the design down once - in .Rules - means Copilot builds it the same way every month.

**How it works**

- List the 4-6 decisions the dashboard supports and one KPI, formula and target for each.
- Sketch the layout: KPI row, chart grid, slicers and timeline, title and data date.
- Choose a small palette and when to use it (e.g. amber for below target).

**Worked example**

- Offer acceptance rate = accepted / (accepted + declined); pending offers excluded.
- Hires vs target % card: green at 100% or more, amber 90-99.9%, red below 90%.
- Navy for hires, teal for targets, amber for warnings - written as a rule.

**Decision guide**

| Use when | Avoid when |
|---|---|
| The same measures are reviewed regularly by the same audience. | A metric has no agreed definition or no one acts on it. |
| Each KPI has an owner who can act on a change. | Twenty KPIs and ten colours crowd out the message. |

**Practitioner quality lens**

- Defined: Formula, period, exclusions and target written down.
- Readable: Clear hierarchy; titles state the finding.
- Purposeful colour: Colour signals status and follows one palette.

**Authoritative references**

- https://support.microsoft.com/en-us/excel/copilot/visualize-your-data-with-copilot-in-excel
- https://support.microsoft.com/en-us/excel/copilot/data-insights-with-copilot-in-excel
- https://support.microsoft.com/en-us/excel/copilot/copilot-in-excel-rules
- https://www.youtube.com/watch?v=4z4eIaEEH9A

---


### Plan and Build with Rules and Skills

Copilot can build a whole dashboard - KPI formulas, PivotTables, PivotCharts, slicers and layout - from one brief. With the KPI definitions, layout and colours already on the .Rules sheet, the brief only says what to show. Plan mode lets you review the build steps before anything changes; afterwards, @chart-design polishes the charts and @theme-design applies a coordinated theme.

A dashboard is a multi-step build, exactly where an agentic plan-execute-check loop saves time - and where one wrong definition gets copied into five objects. Reviewing the plan catches it early. Rules keep every rebuild consistent, and skills bring design expertise that most analysts do not have time to apply by hand.

**How it works**

- Write the dashboard rules and an independent Checks sheet first.
- Plan mode: send a short brief; check the plan's definitions and layout; approve.
- Polish one prompt at a time with @chart-design, conditional formatting and @theme-design.

**Worked example**

- Plan mode: 'Build a one-page H1 2026 recruitment dashboard following the workbook rules: six KPI cards, four charts, slicers and a timeline.'
- The plan counts hires by Applied_Month; you reply 'Use Hire_Date month, as the rules say' before approving.
- '@chart-design Review every chart: finding-style titles, axis titles and the rule colours.'

**Decision guide**

| Use when | Avoid when |
|---|---|
| The data is a clean table and the KPI definitions are written down. | The workbook is shared live with others and you have not saved a version. |
| You can review the plan and each created object afterwards. | You would accept the KPI values without checking a single one. |

**Practitioner quality lens**

- Plan review: Definitions in the plan match the rules.
- Linked: KPI cards use formulas or GETPIVOTDATA, not typed numbers.
- Verified: Each KPI ties to the Checks sheet.

**Authoritative references**

- https://support.microsoft.com/en-us/excel/copilot/get-started-with-copilot-in-excel
- https://support.microsoft.com/en-us/excel/copilot/copilot-in-excel-rules
- https://support.microsoft.com/en-us/excel/copilot/copilot-in-excel-built-in-skills
- https://www.microsoft.com/en-us/copilot/blog/2026/06/25/copilot-in-excel-built-for-the-era-of-frontier-finance/

---


### Interactivity: Slicers, Timelines and KPI Cards

Slicers are on-sheet filter buttons for a field such as Sector; a timeline is a slider for a date field such as Hire_Date. Both can be connected to several PivotTables at once, so every PivotChart on the page changes together. KPI cards follow the slicers only when they read from the PivotTables, for example with GETPIVOTDATA, rather than from the raw table.

Interactivity lets a leader answer their own follow-up questions - 'and for Technology only?' - without a new report. But a slicer connected to only some charts, or KPI cards that ignore it, gives a page that contradicts itself. Testing one slicer value and one date range against independent COUNTIFS proves the page is wired correctly.

**How it works**

- Ask Copilot to add slicers and a timeline and connect them to every PivotTable on the data sheet.
- Make the KPI cards read from the PivotTables (GETPIVOTDATA) so they respond.
- Test one slicer value and one date range against COUNTIFS checks, then clear them.

**Worked example**

- Sector = Technology: the Applications and Hires cards match two COUNTIFS checks.
- Timeline Apr-Jun 2026: the Hires card equals COUNTIFS on Hire_Date between 1 Apr and 30 Jun.
- A card that does not move is re-linked to its PivotTable with GETPIVOTDATA.

**Decision guide**

| Use when | Avoid when |
|---|---|
| Users need to explore by sector, source, recruiter, department or date. | Slicers are left unconnected to some charts. |
| All objects on the page come from PivotTables on one data sheet. | A filter combination would expose a group small enough to identify a person. |

**Practitioner quality lens**

- Connected: Every slicer and the timeline filter every related object.
- Responsive: KPI cards change with the filters.
- Tested: One slicer and one date range checked with COUNTIFS.

**Authoritative references**

- https://support.microsoft.com/en-us/excel/overview-of-pivottables-and-pivotcharts
- https://support.microsoft.com/en-us/office/create-a-pivottable-timeline-to-filter-dates-d3956083-01be-408c-906d-6fc99d9fadfa
- https://support.microsoft.com/en-us/excel/copilot/get-started-with-copilot-in-excel

---


### Lab 7 · Interactive Dashboard — Build an Interactive Recruitment Dashboard with Rules and Skills

Learning outcome: LO3: combine a .Rules sheet, Plan mode and the chart-design and theme-design skills to build a highly visual recruitment dashboard with KPI cards, PivotCharts, slicers and a timeline, and verify every KPI.

Goal: Deliver a one-page H1 2026 recruitment dashboard that looks professional, responds to slicers and a timeline, and agrees with independent checks.

BrightPath's Head of Recruitment wants one page that shows H1 2026 applications, hires, offer acceptance, time to hire, cost per hire and hires against target, and can be filtered by sector, source, recruiter and date. You will put the KPI definitions, layout and colours in a .Rules sheet, so the build prompt stays short. Copilot plans the dashboard in Plan mode, builds it, then @chart-design and @theme-design make it look professional. Finally you prove that every KPI, slicer and the timeline give the right numbers.

**What you'll build**

BrightPath_Recruitment_Dashboard_Data.xlsx with a .Rules sheet, a Checks sheet, PivotTables on Dash_Data and a Dashboard sheet with six formula-driven KPI cards (one with traffic-light colours), four PivotCharts, Sector, Source and Recruiter slicers and a Hire_Date timeline   (Tools: Excel · Copilot (Plan, Allow editing, Chat only) · .Rules · @chart-design · @theme-design · PivotCharts, slicers, timeline.)

**Prerequisites**

- Topic 2 completed: you can write .Rules and call skills with @.
- This lab folder's data workbook and rules-template.txt.

**Step-by-step**

1. Go to onedrive.com (or your SharePoint site), open or create the folder work-c34 and select Upload > Files. Upload BrightPath_Recruitment_Dashboard_Data.xlsx from this lab's data folder, then open it in Excel for the web or select Open in Desktop App. OneDrive gives you AutoSave and version history (recent Excel builds can also run Copilot on local files). Check that Formulas > Calculation Options is Automatic, then open Copilot from the Copilot button on the Home tab or the floating Copilot icon at the bottom right. Read the About_This_Data sheet before you prompt. Read the KPI definitions on About_This_Data - they go into the rules next.

   ```text
   Tables : tblApps (600 applications, Jan-Jun 2026) | tblTargets (monthly hire targets)
   KPI definitions:
     Applications          = number of rows
     Hires                 = Furthest_Stage = "Hired", by Hire_Date month
     Offer acceptance rate = accepted / (accepted + declined)   (Pending excluded)
     Avg time to hire      = average Time_to_Hire_Days of hires
     Cost per hire         = total Sourcing_Cost / hires
   ```

2. Write the dashboard's rules first. Select + (Add work content) > Create workbook rules, then replace the suggested rules with BrightPath's: paste rules-template.txt into .Rules!A1, or send the prompt below. From now on every prompt in this workbook follows these definitions, the layout and the colours.

   ```text
   ALLOW EDITING:
   On the .Rules sheet, replace everything in column A with these rules, one per cell from A1 down.
   Keep each ## heading in its own cell. Do not change any other sheet.
   ## DASHBOARD STANDARDS
   This workbook builds BrightPath's H1 2026 recruitment dashboard. Never change tblApps or tblTargets.
   Applications = rows of tblApps. Hires = Furthest_Stage "Hired", counted by Hire_Date month.
   Offer acceptance rate = accepted / (accepted + declined). Pending offers are excluded.
   Avg time to hire = average Time_to_Hire_Days of hires. Cost per hire = total Sourcing_Cost / hires.
   KPI cards are formulas linked to the PivotTables (GETPIVOTDATA) so they respond to slicers - never typed numbers.
   ## LAYOUT
   Dashboard sheet: title in A1, "Data as at 2026-06-30" in A2, KPI cards in row 4, charts in a grid below, slicers and the timeline on the right.
   Put every PivotTable on a sheet named Dash_Data.
   Never merge cells. Hide gridlines on the Dashboard sheet.
   ## VISUAL STYLE
   Colours: navy #1F4E79 for hires, teal #2A9D8F for targets and comparisons, amber #F4A261 for warnings.
   Every chart title states the finding, e.g. Technology led H1 hiring.
   Money as S$ with no decimals, except Cost per hire with 2 decimals. Rates as % with 1 decimal. Days with 1 decimal.
   ```

3. Build the independent checks before the dashboard. These are the numbers the KPI cards must match.

   ```text
   CHECK (Allow editing):
   Create a sheet named Checks. Enter these labels and formulas exactly as written, then tell me each result:
   A1 Applications        B1 =ROWS(tblApps[Application_ID])
   A2 Hires               B2 =COUNTIFS(tblApps[Furthest_Stage],"Hired")
   A3 Offer acceptance    B3 =COUNTIFS(tblApps[Offer_Accepted],"Yes")/(COUNTIFS(tblApps[Offer_Accepted],"Yes")+COUNTIFS(tblApps[Offer_Accepted],"No"))
   A4 Avg time to hire    B4 =AVERAGEIFS(tblApps[Time_to_Hire_Days],tblApps[Furthest_Stage],"Hired")
   A5 Cost per hire       B5 =SUM(tblApps[Sourcing_Cost])/B2
   A6 Hires vs target     B6 =B2/SUM(tblTargets[Hire_Target])
   ```

4. Switch to Plan mode and send the brief. Because the rules hold the definitions, layout and colours, the brief only says what to show. Read the plan: does it exclude Pending offers, count hires by Hire_Date month and put PivotTables on Dash_Data? Reply with corrections until it does, then approve it. Watch the steps, and select Done when Copilot finishes.

   ```text
   PLAN MODE:
   Build a one-page, highly visual H1 2026 recruitment dashboard on the Dashboard sheet for BrightPath's Head of Recruitment, following the workbook rules.
   KPI cards: Applications, Hires, Offer acceptance rate, Avg time to hire, Cost per hire, Hires vs target %.
   Charts: hires by Hire_Date month vs Hire_Target (columns + line); hires by Sector (bar, sorted); applications by Source (bar, sorted); hires by Recruiter (column).
   Interactivity: slicers for Sector, Source and Recruiter and a timeline on Hire_Date, connected to every PivotTable.
   ```

5. Make it visual with skills, one prompt at a time, and keep or undo each change. @chart-design polishes every chart; the traffic-light rule makes the target card readable at a glance; @theme-design pulls the page together while keeping the rule colours.

   ```text
   ALLOW EDITING:
   @chart-design Review every chart on the Dashboard: titles that state the finding, axis titles, data labels where they help, and the colours in the workbook rules.

   FOLLOW-UP:
   Colour the Hires vs target % card green at 100% or more, amber from 90% to 99.9% and red below 90%.

   FOLLOW-UP:
   @theme-design Apply a clean, modern theme to the Dashboard that keeps the colours in the workbook rules.
   ```

6. Verify. With no slicer or timeline selection, ask Copilot to compare each card with the Checks sheet and show the card formulas. Fix any difference with a targeted prompt - the most common error is including Pending offers.

   ```text
   CHAT ONLY:
   Compare each KPI card on the Dashboard with Checks!B1:B6, with no slicer selected. Show each card's formula, its value, the Checks value and whether they match.

   ALLOW EDITING (if needed):
   The Offer acceptance rate card must exclude Offer_Accepted = "Pending", as the workbook rules say.
   Fix the card formula.
   ```

7. Test the interactivity. Set the Sector slicer to Technology and read the cards, then clear it. Then select April to June 2026 on the Hire_Date timeline and read the Hires card, then clear it. Ask Copilot to add the checks and compare.

   ```text
   CHECK (Allow editing):
   On the Checks sheet, add these formulas in D1:D3 with labels in C1:C3, and tell me if they match the cards when Sector = Technology, and the Hires card when the timeline shows Apr-Jun 2026:
   =COUNTIFS(tblApps[Sector],"Technology")
   =COUNTIFS(tblApps[Sector],"Technology",tblApps[Furthest_Stage],"Hired")
   =COUNTIFS(tblApps[Hire_Date],">="&DATE(2026,4,1),tblApps[Hire_Date],"<="&DATE(2026,6,30))

   If a KPI card does not change, ask: "Link the KPI cards to the PivotTables with GETPIVOTDATA so they respond to the slicers and timeline."
   ```

8. Ask Copilot (Chat only) to review the dashboard as the Head of Recruitment would, act on the valid points, and log your prompts.

   ```text
   CHAT ONLY:
   Review this dashboard for a Head of Recruitment. What is unclear, missing or potentially misleading?
   Does it follow every rule on the .Rules sheet?

   LOG (Allow editing):
   Prompt_Log already lists every prompt in this lab. For each one I ran, fill in What Copilot did and How I verified, and set Decision to Keep. I will tell you any row to change to Revise or Undo.
   ```


**Sample prompts**

| Purpose | Prompt |
|---|---|
| Plan | Plan a recruitment dashboard following the workbook rules. Show me the plan first. |
| KPI card | Create a KPI card for Cost per hire linked to the PivotTable with GETPIVOTDATA, formatted as S$. |
| Chart skill | @chart-design Turn the hires-by-month chart into columns for hires and a line for Hire_Target. |
| Timeline | Add a timeline on Hire_Date and connect it to every PivotTable on Dash_Data. |
| Visual cue | Add data bars to the hires-by-recruiter table and traffic-light colours to the target card. |
| Theme skill | @theme-design Apply a modern theme that keeps the colours in the workbook rules. |
| Review | Does this dashboard follow every rule on the .Rules sheet? List any that it breaks. |
| Question | Which recruiter made the most hires in H1 2026, and what was their average time to hire? |

**Test it**

| Check | Expected result |
|---|---|
| Applications and hires | 600 applications, 41 hires (no slicer selected) |
| Offer acceptance rate | 78.8% = 41 of 52; 10 pending excluded |
| Time to hire, cost per hire | 27.8 days; S$348.20 |
| Hires vs target (amber) | 93.2% = 41 of 44 |
| Slicer: Sector = Technology | 124 applications, 5 hires, 42.0 days |
| Timeline: Apr-Jun 2026 | 22 hires; no KPI card holds a typed number |

**Checkpoint and rejoin point**

Keep the workbook in work-c34. Lab 8 turns this dashboard recipe into a reusable skill for workforce data.

**Troubleshooting**

| If this happens | Fix |
|---|---|
| Copilot builds before you approve the plan. | You were in Allow editing. Select Undo, switch to Plan mode and resubmit. |
| KPI cards do not change with the slicers or timeline. | Cards built with COUNTIFS on tblApps ignore slicers. Ask Copilot to link them to the PivotTables with GETPIVOTDATA. |
| The timeline cannot be added. | A timeline needs a date field in a PivotTable. Ask Copilot to make sure Hire_Date is a date column and is in a PivotTable on Dash_Data. |
| Hires appear in January. | The chart counts by Applied_Month. Ask Copilot to group hires by Hire_Date month, as the rules say. |

**Challenge**

Ask Copilot to add a small 'What changed' panel under the KPI cards with two sentences that use formulas (for example ="Hires: "&Checks!B2&" of "&SUM(tblTargets[Hire_Target])&" target"), so the text updates with the data.

**Reflection**

Which rule on your .Rules sheet saved you the most prompting, and which visual element makes the dashboard easiest to read?

> **Note:** The lab folder labs/lab-07-*/ holds the instruction PDF with sample prompts, sample-prompts.txt and the mock-data workbook (data/BrightPath_Recruitment_Dashboard_Data.xlsx), plus rules-template.txt (the .Rules text to paste into column A). Use only the supplied fictional BrightPath data or information you are authorised to process. Do not paste credentials, confidential business material or real employees' personal data into Copilot, rules or skills.

---

Concepts for Lab 8: Workforce Measures: Headcount, Attrition and Tenure, A Reusable Dashboard Skill and the Refresh Test, Sharing People Dashboards Responsibly.


### Workforce Measures: Headcount, Attrition and Tenure

Headcount is the number of active employees on a date. Attrition rate is the number of exits in a period divided by the average headcount over that period, usually (opening headcount + closing headcount) / 2; voluntary attrition counts resignations only. Tenure is the time from hire date to exit date, or to the report date for staff who are still active.

These words are calculated differently across organisations, and small choices change the answer: dividing by closing headcount instead of the average, or counting people hired after the period, moves the rate. Writing the definitions into the workbook's .Rules sheet makes every rebuild - and every skill run - use the same ones. Small groups give volatile rates that can also identify individuals.

**How it works**

- Fix the period (e.g. 2025-07-01 to 2026-06-30) and define 'active on a date'.
- Count opening and closing headcount, exits and voluntary exits with COUNTIFS.
- Divide exits by the average headcount, and report groups of fewer than 5 people only in total.

**Worked example**

- Opening headcount 210, closing 237, average 223.5.
- 29 exits / 223.5 = 13.0% attrition; 25 voluntary exits = 11.2%.
- Recruitment has the highest attrition, so it is reviewed as a team, never person by person.

**Decision guide**

| Use when | Avoid when |
|---|---|
| Reporting workforce trends to HR leaders over a fixed period. | The period or the headcount basis differs between reports. |
| Comparing departments or locations that are large enough to stay anonymous. | A group is so small that its rate could identify a person. |

**Practitioner quality lens**

- Defined: Period, headcount basis and exit types written down.
- Reconciled: Rate rebuilt from COUNTIFS on the source table.
- Private: No group under 5 people shown on its own.

**Authoritative references**

- https://support.microsoft.com/en-us/excel/copilot/frequently-asked-questions-about-copilot-in-excel
- https://support.microsoft.com/en-us/excel/copilot/copilot-in-excel-rules

---


### A Reusable Dashboard Skill and the Refresh Test

Once a dashboard recipe works, write it down as a custom skill: read .Rules, add helper columns, build a Checks sheet, create the PivotTables, cards, charts and slicers, then compare every card with the checks. Because the source is an Excel table, new rows are included automatically; refreshing all PivotTables updates the cards and charts. The refresh test proves the dashboard is live, not a one-off report.

A monthly dashboard should take one line to rebuild, not a long brief - and every rebuild should follow the same definitions. A skill plus workbook rules gives exactly that. The refresh test exposes static pictures, typed numbers and tables that did not expand, which are the most common reasons a dashboard silently goes out of date.

**How it works**

- Ask Copilot (Chat only) to write the dashboard SKILL.md; save it, refresh and turn it on.
- Run '@skill-name' in Plan mode on the workbook with its .Rules sheet; review, approve, verify.
- Append test rows to the source table, refresh all, and confirm each KPI moves by the expected amount.

**Worked example**

- '@hr-workforce-dashboard Build the workforce dashboard for the 12 months ending 2026-06-30.'
- Eight July joiners are appended to tblEmployees; active headcount rises by exactly 8.
- A chart that did not update was a pasted picture; it is replaced with a linked PivotChart.

**Decision guide**

| Use when | Avoid when |
|---|---|
| The dashboard will be rebuilt or refreshed with new data each period. | The skill builds its own checks and nobody compares them with known values. |
| The recipe has been tested once by hand and its checks pass. | Numbers are typed into KPI cards instead of calculated. |

**Practitioner quality lens**

- Repeatable: One-line skill call plus .Rules rebuilds the page.
- Refresh-tested: Added rows change KPIs by the expected amount.
- Self-checking: Skill compares cards with Checks and reports differences.

**Authoritative references**

- https://support.microsoft.com/en-us/excel/copilot/copilot-in-excel-skills
- https://support.microsoft.com/en-us/excel/copilot/copilot-in-excel-rules
- https://support.microsoft.com/en-us/excel/overview-of-pivottables-and-pivotcharts
- https://www.youtube.com/watch?v=4z4eIaEEH9A

---


### Sharing People Dashboards Responsibly

People data - salaries, performance, engagement and exits - is sensitive. A responsible HR dashboard shows aggregated results, suppresses groups too small to protect anonymity (for example fewer than 5 people), limits access through OneDrive or SharePoint permissions, and keeps a record of who changed what. Copilot can review the page for privacy risks and draft a leadership summary from the verified numbers.

Copilot does not give anyone access to data they could not already open - but anyone who can open the workbook sees everything in it, including hidden sheets and the raw table behind a dashboard. A privacy rule in .Rules and a check step in the skill make responsible sharing the default, and a human still approves every number before it leaves HR.

**How it works**

- Write the privacy rule into .Rules: no IDs, salaries or individual records; no group under 5.
- Ask Copilot which slicer combinations show fewer than 5 people, and suppress them.
- Share with specific people, review change history, and check every number in any summary.

**Worked example**

- Salary detail stays in the source file; the leadership copy shows only aggregates.
- Filtering IT + Director would show one person, so that combination is grouped into Other.
- Copilot drafts a 5-line summary using only dashboard numbers; HR checks each one before sending.

**Decision guide**

| Use when | Avoid when |
|---|---|
| The dashboard leaves the HR team or is used in a leadership meeting. | Individual-level records are visible to a broad audience. |
| The data includes pay, performance, health or exit details. | A small group's result could identify a person. |

**Practitioner quality lens**

- Minimise: Only the fields the audience needs.
- Protect: Small-group suppression and specific-people sharing.
- Audit: Change history and version history retained.

**Authoritative references**

- https://support.microsoft.com/en-us/excel/copilot/frequently-asked-questions-about-copilot-in-excel
- https://support.microsoft.com/en-us/excel/copilot/copilot-in-excel-change-history
- https://support.microsoft.com/en-us/excel/copilot/copilot-in-excel-rules

---


### Lab 8 · Dashboard Skill, Refresh & Share — Turn the Dashboard into a Reusable Skill, Refresh It and Share It Safely

Learning outcome: LO3: use Copilot to write a dashboard skill, run it on a workforce workbook with rules, prove the dashboard refreshes with new rows, and check it is safe to share.

Goal: Build BrightPath's workforce and attrition dashboard from a one-line skill call, prove it updates when new joiners arrive, and make sure no small group can be identified.

BrightPath's HR Director needs the same kind of dashboard every month. Instead of repeating a long brief, you will ask Copilot to write an hr-workforce-dashboard skill, save it to OneDrive and run it on the workforce workbook, whose .Rules sheet holds the period, the attrition definition and the privacy rule. Then you will append eight July joiners and refresh to prove the dashboard is live, check that no slicer combination exposes fewer than 5 people, and ask Copilot for a short leadership summary.

**What you'll build**

An hr-workforce-dashboard skill in OneDrive, and BrightPath_Workforce_Headcount.xlsx with a .Rules sheet, helper columns, a Checks sheet and an HR_Dashboard with six KPI cards, four charts and Location and Job_Level slicers - refresh-tested and privacy-checked   (Tools: Excel · Copilot (Chat only, Plan, Allow editing) · .Rules · custom skill · PivotTables, slicers, refresh.)

**Prerequisites**

- Lab 6 and Lab 7 completed: you can save a custom skill to OneDrive and build a dashboard with rules.
- This lab folder's data workbook, rules-template.txt and skills/hr-workforce-dashboard/SKILL.md (fallback).

**Step-by-step**

1. Go to onedrive.com (or your SharePoint site), open or create the folder work-c34 and select Upload > Files. Upload BrightPath_Workforce_Headcount.xlsx from this lab's data folder, then open it in Excel for the web or select Open in Desktop App. OneDrive gives you AutoSave and version history (recent Excel builds can also run Copilot on local files). Check that Formulas > Calculation Options is Automatic, then open Copilot from the Copilot button on the Home tab or the floating Copilot icon at the bottom right. Read the About_This_Data sheet before you prompt. Do not touch the New_Joiners_Jul2026 sheet until the refresh test.

   ```text
   Table : tblEmployees (420 employees hired since 2019; Status Active or Exited)
   Period: 2025-07-01 to 2026-06-30
   Attrition rate = exits in period / ((active on 2025-07-01 + active on 2026-06-30) / 2)
   ```

2. Create the workbook rules: + (Add work content) > Create workbook rules, then paste rules-template.txt into .Rules!A1 or send the prompt below. Note the privacy rule - the skill will follow it.

   ```text
   ALLOW EDITING:
   On the .Rules sheet, replace everything in column A with these rules, one per cell from A1 down.
   Keep each ## heading in its own cell. Do not change any other sheet.
   ## WORKFORCE DASHBOARD STANDARDS
   Never change or delete existing rows of tblEmployees.
   Never show Employee_ID, Monthly_Salary or any individual record on a dashboard.
   The 12-month period is 2025-07-01 to 2026-06-30. Active on a date = Hire_Date on or before it and Exit_Date blank or after it.
   Active headcount = employees with Status Active now, so new joiners count as soon as they are added to tblEmployees.
   Attrition rate = exits in the period / ((active on 2025-07-01 + active on 2026-06-30) / 2). New joiners after 2026-06-30 do not change it.
   Tenure_Years = years from Hire_Date to Exit_Date, or to 2026-06-30 if still active; use 0 for anyone hired after 2026-06-30.
   Do not show any group of fewer than 5 employees. Group it into Other or leave it out.
   ## LAYOUT
   Dashboard sheet HR_Dashboard: title in A1, "Data as at 2026-06-30" in A2, KPI cards in row 4, charts in a grid below, slicers on the right.
   Put every PivotTable on a sheet named Dash_Data. Never merge cells. Hide gridlines on HR_Dashboard.
   ## VISUAL STYLE
   Colours: navy #1F4E79 for headcount, teal #2A9D8F for hires, amber #F4A261 for exits and attrition.
   Every chart title states the finding. Rates as % with 1 decimal. Years with 1 decimal.
   ```

3. Let Copilot write the skill. In Chat only, describe the recipe you followed in Lab 7, made general. Read the reply and ask for changes until the steps are clear.

   ```text
   CHAT ONLY:
   Write a complete SKILL.md for a Copilot in Excel skill named hr-workforce-dashboard.
   Description: Use when I ask for a workforce, headcount or attrition dashboard from an employee table.
   Steps:
   1) Read the .Rules sheet first and follow it.
   2) Add helper columns Active_Start, Exit_In_Period and Tenure_Years using the period in the rules.
   3) Create a Checks sheet with independent COUNTIFS formulas for active headcount (Status Active), headcount at period start, headcount at period end, exits, voluntary exits, attrition rate and hires.
   4) Put all PivotTables on Dash_Data.
   5) Build HR_Dashboard: six KPI cards (Active headcount, Exits, Attrition rate, Voluntary attrition rate,
      Avg tenure of active staff, Hires), charts of headcount by Department, attrition rate by Department, hires vs exits by quarter and headcount by Location, and slicers for Location and Job_Level.
   6) Compare every KPI card with Checks and report any difference.
   7) List any slicer combination with fewer than 5 employees.
   Return only the SKILL.md text, starting with the --- front matter.
   ```

4. Save it as before: copy the reply into a plain-text file named SKILL.md, create the folder hr-workforce-dashboard in your OneDrive skills folder and upload the file (or upload the ready-made one from this lab folder). In Excel select Settings (...) > Manage plugins & skills > Custom skills > Refresh and turn the skill on.

   ```text
   OneDrive : .../skills/hr-workforce-dashboard/SKILL.md
   Check    : folder name = name: line = hr-workforce-dashboard
   ```

5. Run the skill in Plan mode: one line replaces the long brief, because the skill holds the steps and the rules hold the definitions. Review the plan, approve it, and select Done when Copilot finishes.

   ```text
   PLAN MODE:
   @hr-workforce-dashboard Build the workforce dashboard for the 12 months ending 2026-06-30.
   ```

6. Verify. Ask Copilot to compare the KPI cards with the Checks sheet the skill built, and to show the attrition formula. Compare both with the Test It values - a skill can build a wrong check too.

   ```text
   CHAT ONLY:
   Compare each KPI card on HR_Dashboard with the Checks sheet and tell me which, if any, do not match.
   Show the formula behind the Attrition rate card.

   ALLOW EDITING (if needed):
   The attrition rate must use average headcount = (active on 2025-07-01 + active on 2026-06-30) / 2, as the rules say.
   ```

7. Refresh test: ask Copilot to append the 8 rows from New_Joiners_Jul2026 to tblEmployees and refresh every PivotTable. Every KPI and chart that counts active staff must update.

   ```text
   ALLOW EDITING:
   Copy the 8 data rows in New_Joiners_Jul2026!A2:L9 to the first empty rows directly below tblEmployees so the table expands. Then refresh all PivotTables in the workbook and tell me the new Active headcount and the number of rows in tblEmployees.

   Expected after refresh:
   Active headcount : 237 -> 245
   tblEmployees rows: 420 -> 428
   New joiners by department: Recruitment +2, Sales & Client Services +2, IT +4
   ```

8. Check it is safe to share, then ask for a leadership summary that uses only dashboard numbers. Copilot drafts; you check every number before it leaves HR. Log your prompts.

   ```text
   CHAT ONLY:
   Review HR_Dashboard before I share it with the leadership team. Does any element expose individual employee data or a group smaller than 5? Is each KPI labelled with its period and definition?

   CHAT ONLY:
   Write a 5-line summary of HR_Dashboard for BrightPath's leadership team. Use only numbers shown on the dashboard, and label any explanation as a hypothesis to test.

   LOG (Allow editing):
   Prompt_Log already lists every prompt in this lab. For each one I ran, fill in What Copilot did and How I verified, and set Decision to Keep. I will tell you any row to change to Revise or Undo.
   ```


**Sample prompts**

| Purpose | Prompt |
|---|---|
| Write a skill | Write a SKILL.md named hr-workforce-dashboard that reads the .Rules sheet first and builds a verified workforce dashboard. |
| Run a skill | @hr-workforce-dashboard Build the workforce dashboard for the 12 months ending 2026-06-30. |
| Refresh | I added rows to tblEmployees. Refresh all PivotTables and tell me which KPI values changed. |
| Privacy | Which slicer combinations show a group of fewer than 5 employees? Suggest how to suppress them. |
| Explain | Explain how the Attrition rate card is calculated, step by step. |
| Improve the skill | Update the hr-workforce-dashboard SKILL.md so it also adds a timeline on Hire_Date. Return the full file. |
| Summary | Write a 5-line leadership summary using only the numbers on HR_Dashboard. |

**Test it**

| Check | Expected result |
|---|---|
| Active headcount | 237 before the refresh |
| Exits and attrition | 29 exits (25 voluntary); 13.0% attrition |
| Hires and tenure | 56 hires; average tenure about 3.1 years |
| Highest-attrition department | Recruitment: 21.8% (12 exits) |
| Refresh test | After adding the July joiners: headcount 245, 428 rows |
| Privacy | No Employee_ID, salary, individual record or group under 5 on the dashboard |

**Checkpoint and rejoin point**

You now have two custom skills in OneDrive. Run hr-workforce-dashboard again next month on new data, and improve the skill whenever you correct Copilot twice.

**Troubleshooting**

| If this happens | Fix |
|---|---|
| The new rows did not join the table and nothing changed on refresh. | Ask Copilot to resize tblEmployees to include the 8 new rows, then refresh all PivotTables again. |
| Attrition shows 12.2% (29 / 237) instead of 13.0%. | The card divides by closing headcount only. Ask Copilot to use the average of opening and closing headcount, then add that wording to the skill. |
| The skill ignores the .Rules sheet. | Check the sheet is named .Rules and is visible, and that step 1 of the skill says to read it. Start a New chat. |
| Helper columns show #VALUE! for blank Exit_Date. | Ask Copilot to treat blank Exit_Date explicitly, e.g. OR([@Exit_Date]="",[@Exit_Date]>DATE(2025,7,1)). |

**Challenge**

Improve the skill so it also adds a Timeline on Hire_Date and a chart of average Engagement_Score by Department for active staff. Save, refresh and rerun it, then check one department's average with AVERAGEIFS.

**Reflection**

What would happen to trust in this dashboard if one leader could filter it down to a single named team member?

> **Note:** The lab folder labs/lab-08-*/ holds the instruction PDF with sample prompts, sample-prompts.txt and the mock-data workbook (data/BrightPath_Workforce_Headcount.xlsx), plus rules-template.txt (the .Rules text to paste into column A), plus skills/hr-workforce-dashboard/SKILL.md (ready-made skill to upload if you run short of time). Use only the supplied fictional BrightPath data or information you are authorised to process. Do not paste credentials, confidential business material or real employees' personal data into Copilot, rules or skills.

---


### Topic 03 Recap - Mapped to Learning Outcomes

Use this checkpoint to explain what you can now do and identify any lab evidence you still need to repair.

| Learning outcome | Evidence from this topic |
|---|---|
| LO3 · Design | KPI definitions, layout and colours written in .Rules before anything is built. |
| LO3 · Build | Plan mode, then @chart-design and @theme-design; every KPI card matches the Checks sheet. |
| LO3 · Interact | Slicers, a timeline and GETPIVOTDATA cards tested against COUNTIFS. |
| LO3 · Reuse and share | A dashboard skill, a passing refresh test, and no group under 5 people on view. |

---


## Wrap-Up - The Complete C34 Workflow

Across eight labs you moved from single prompts to rules, skills and a reusable, interactive dashboard - always with a verification step between Copilot's output and a people decision.

**Excel fundamentals**

- Choose Chat only, Plan or Allow editing on purpose, and keep or undo every change.
- Prepare named tables, fix data quality, and describe formulas and XLOOKUPs in plain English.
- Build PivotTables, PivotCharts and charts, and tie each one to an independent total.

**Skills and rules**

- Personal preferences go in Personalization; one workbook's standards go in a visible .Rules sheet.
- Call built-in skills with @ - and hold them to your definitions.
- Let Copilot write a SKILL.md, save it in its OneDrive folder, refresh, and improve it over time.

**Interactive dashboards**

- Define KPIs, layout and colours in .Rules; plan, build and polish with Plan mode and skills.
- Wire slicers, a timeline and formula-driven KPI cards, and prove refresh with new rows.
- Save the recipe as a skill, and share only aggregated, permissioned views.

---


## Next Steps

- Re-run all eight labs from the original data workbooks without looking at your finished files.
- Add your three most common Copilot corrections to Personalization this week.
- Create a .Rules sheet for one important workbook in your role - it takes a few minutes.
- Turn one weekly or monthly report into a custom skill, test it on a copy, and improve it after each run.
- Check Microsoft Support regularly: Copilot in Excel's modes, skills and labels change frequently.


## Glossary

- **.Rules sheet** — A visible worksheet named .Rules whose column A holds rules Copilot follows in that workbook.
- **@mention** — Typing @ and a skill's name in the prompt to call that skill, e.g. @chart-design.
- **Allow editing** — The default Copilot mode: it plans and makes changes to the workbook in one go.
- **Attrition rate** — Exits in a period divided by average headcount in that period.
- **CAC** — Customer acquisition cost: sales and marketing spend divided by new clients won.
- **Chat only** — The Copilot mode that answers and analyses without changing the workbook.
- **Clean data with Copilot** — Automated cleanup (Data > Clean Data) that suggests fixes for spaces, capitals and similar issues.
- **Cost per hire** — Total recruiting cost in a period divided by hires in that period.
- **Custom skill** — A SKILL.md file you save in its own folder in your OneDrive skills folder.
- **Front matter** — The block between --- lines at the top of SKILL.md holding the skill's name and description.
- **Sheet reference** — Sheet!Cell, e.g. Control!B1 - a cell on another sheet; $ (e.g. $B$1) locks it when the formula is copied.
- **Structured reference** — A formula reference by table and column name; [@Column] means this row, Table[Column] the whole column.
- **GETPIVOTDATA** — An Excel function that reads a value from a PivotTable, so a KPI card follows its slicers.
- **KPI card** — A dashboard cell or shape that shows one headline measure, calculated by formula.
- **LTV** — Lifetime value: average revenue per client x gross margin / monthly churn.
- **Offer acceptance rate** — Accepted offers divided by offers with a decision (accepted + declined).
- **Personalization** — Your own Copilot preferences, stored with your account and applied in every workbook.
- **PivotChart** — A chart linked to a PivotTable that updates and filters with it.
- **Plan mode** — The Copilot mode that shows its planned steps and waits for your approval before changing anything.
- **Skill** — Saved instructions that tell Copilot how to do a repeatable task; built-in or custom.
- **Slicer** — An on-sheet filter button that can control several PivotTables and charts.
- **Timeline** — A slider filter for a date field that can control several PivotTables.
- **XLOOKUP** — An Excel function that finds a value in one range and returns the matching value from another.
