# Lab 5 — Build a Recruitment KPI Dashboard with Copilot

**Course:** Microsoft 365 Copilot for Excel  
**Course Code:** C34  
**Version:** v2.0 (8 October 2026)  
**Topic 3:** Use Copilot to Create Dashboards  
**Maps to:** LO3: define recruitment KPIs, use Copilot Plan and Allow editing modes to build a dashboard of KPI cards, PivotTables, charts and slicers, and verify every KPI against an independent check  
**Duration:** 45 minutes  
**Tools:** Excel · Copilot (Plan, Allow editing) · PivotTables, PivotCharts, slicers · BrightPath_Recruitment_Dashboard_Data.xlsx

> 📄 **Instruction PDF (with sample prompts):** [Lab-05-Instructions.pdf](Lab-05-Instructions.pdf)  
> 💬 **Copy-ready prompts:** [sample-prompts.txt](sample-prompts.txt)  
> 🌐 **Prompts web page (one-click copy):** [Lab 5 prompts](https://tertiarycourses.github.io/C34-Microsoft-365-Copilot-for-Excel/prompts/#lab-5)  
> 📊 **Mock data:** [data/BrightPath_Recruitment_Dashboard_Data.xlsx](data/BrightPath_Recruitment_Dashboard_Data.xlsx)  

---

## Goal

Deliver a one-page H1 2026 recruitment dashboard whose five KPI cards, three charts and two slicers all agree with independently calculated checks.

## What You Will Do

BrightPath's Head of Recruitment wants one page that shows H1 2026 applications, hires, offer acceptance rate, time to hire and cost per hire, by month, sector, source and recruiter. You will write the checks first, ask Copilot to plan the dashboard, review the plan against the KPI definitions, let Copilot build it in Allow editing mode, and then prove each KPI and slicer works.

## What You Will Build

BrightPath_Recruitment_Dashboard_Data.xlsx with a Checks sheet and a Dashboard sheet holding five formula-driven KPI cards, charts for hires vs target by month, hires by sector and applications by source, and Sector and Recruiter slicers connected to every PivotTable

## Prerequisites

- Topic 2 completed: you can build and verify PivotTables and charts with Copilot.
- This lab folder's data/BrightPath_Recruitment_Dashboard_Data.xlsx.

> **Data note.** Use only the supplied fictional BrightPath data or information you are authorised to process. Do not paste credentials, confidential business material or real employees' personal data into Copilot.

## Steps

### 1. Go to onedrive.com (or your SharePoint site), open or create the folder work-c34 and select Upload > Files. Upload BrightPath_Recruitment_Dashboard_Data.xlsx from this lab's data folder, then open it in Excel for the web or select Open in Desktop App. Confirm the title bar shows the cloud location and AutoSave is On, and that Formulas > Calculation Options is set to Automatic. Read the About_This_Data sheet before you prompt. Read the KPI definitions row on About_This_Data carefully - the dashboard must follow them exactly.

```text
Tables : tblApps (600 applications, Jan-Jun 2026) | tblTargets (monthly hire targets)
KPI definitions:
  Applications          = number of rows
  Hires                 = Furthest_Stage = "Hired"
  Offer acceptance rate = accepted / (accepted + declined)   (Pending excluded)
  Avg time to hire      = average Time_to_Hire_Days of hires
  Cost per hire         = total Sourcing_Cost / hires
```

### 2. Before Copilot builds the dashboard, choose Allow editing and ask Copilot to create a Checks sheet with these independent KPI formulas, exactly as written. These are the numbers the dashboard must match.

```text
Create a sheet named Checks. Enter these labels and formulas exactly as written, then tell me each result:
A1 Applications        B1 =ROWS(tblApps[Application_ID])
A2 Hires               B2 =COUNTIFS(tblApps[Furthest_Stage],"Hired")
A3 Offer acceptance    B3 =COUNTIFS(tblApps[Offer_Accepted],"Yes")/(COUNTIFS(tblApps[Offer_Accepted],"Yes")+COUNTIFS(tblApps[Offer_Accepted],"No"))
A4 Avg time to hire    B4 =AVERAGEIFS(tblApps[Time_to_Hire_Days],tblApps[Furthest_Stage],"Hired")
A5 Cost per hire       B5 =SUM(tblApps[Sourcing_Cost])/B2
```

### 3. Choose Plan. Paste the dashboard brief below. Read Copilot's plan: does it use the same KPI definitions (especially excluding Pending offers) and count hires by Hire_Date month? Reply with corrections until the plan matches. Do not approve yet.

```text
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

### 4. When the plan is right, switch to Allow editing and tell Copilot to carry it out. Watch the steps in the pane; you can select Stop if it goes off track. When it finishes, open Dash_Data and the Dashboard and review what was created.

```text
The plan is approved. Build the dashboard exactly as planned.
```

### 5. Ask Copilot to compare each KPI card with Checks!B1:B5 (with no slicer selected) and show each card's formula. Fix any difference with a targeted follow-up prompt - the most common error is including Pending offers.

```text
CHAT ONLY:
Compare each KPI card on the Dashboard with Checks!B1:B5, with no slicer selected. Show each card's
formula, its value, the Checks value and whether they match.

FOLLOW-UP EXAMPLE:
The Offer acceptance rate card must exclude Offer_Accepted = "Pending". Change its formula to
accepted / (accepted + declined) and show it as a percentage with 1 decimal.
```

### 6. Test interactivity. Ask Copilot to set the Sector slicer to Technology and report every KPI card, then to add the COUNTIFS checks for Technology and compare. Finally ask it to clear the slicer.

```text
ALLOW EDITING:
Set the Sector slicer on the Dashboard to Technology only, then tell me the value on every KPI card.

CHECK (Allow editing):
On the Checks sheet, add these formulas in D1:D2 with labels in C1:C2, and tell me if they match the cards:
=COUNTIFS(tblApps[Sector],"Technology")
=COUNTIFS(tblApps[Sector],"Technology",tblApps[Furthest_Stage],"Hired")

ALLOW EDITING:
Clear the Sector slicer so the dashboard shows all sectors.

If a KPI card does not change, ask: "Make the KPI cards respond to the Sector and Recruiter slicers."
```

### 7. Polish with one prompt per change, then ask Copilot (Chat only) to review the dashboard for clarity. Ask Copilot to log all prompts on Prompt_Log.

```text
Add a note under the title: "Source: tblApps; KPI definitions on About_This_Data".
Use one consistent colour scheme and align the charts in a grid.
CHAT ONLY: Review this dashboard for a Head of Recruitment. What is unclear or missing?

LOG (Allow editing):
Prompt_Log already lists every prompt in this lab. For each one I ran, fill in What Copilot did and
How I verified, and set Decision to Keep. I will tell you any row to change to Revise or Undo.
```

## Sample Prompts

| # | Purpose | Prompt |
|---|---|---|
| 1 | Plan | Plan a recruitment KPI dashboard on the Dashboard sheet from tblApps with KPI cards for applications, hires, offer acceptance rate, average time to hire and cost per hire. List the formulas before building. |
| 2 | KPI card | Create a KPI card that shows Cost per hire = total Sourcing_Cost / hires, formatted as S$ with 2 decimals. |
| 3 | Chart vs target | Create a column chart of hires by Hire_Date month with Hire_Target from tblTargets as a line. |
| 4 | Slicers | Add slicers for Sector and Recruiter and connect them to all PivotTables on Dash_Data. |
| 5 | Fix a KPI | The offer acceptance rate must exclude Pending offers. Update the formula and explain it. |
| 6 | Layout | Arrange the KPI cards in one row and the three charts in a grid below them with equal sizes. |
| 7 | Review | Review this dashboard for a Head of Recruitment. What is unclear, missing or potentially misleading? |
| 8 | Question | Which recruiter made the most hires in H1 2026, and what was their average time to hire? |

## Test It

With no slicer selected the KPI cards match Checks: Applications 600, Hires 41, Offer acceptance 78.8% (41 accepted of 52 decided; 10 pending excluded), Avg time to hire 27.8 days, Cost per hire S$348.20. Hires by month: 2026-02 = 9, 2026-03 = 10, 2026-04 = 9, 2026-05 = 7, 2026-06 = 6 (none in 2026-01). Manufacturing has the most hires (11). With Sector = Technology, Applications = 124, Hires = 5 and Avg time to hire = 42.0 days. Siti Hamid made the most hires (13). No KPI card contains a typed number.

## Checkpoint and Rejoin Point

Keep the workbook in work-c34. Lab 6 builds a second dashboard from the workforce workbook in its own folder.

## Troubleshooting

| If this happens | Fix |
|---|---|
| Copilot builds before you approve the plan. | You were in Allow editing mode. Undo (Ctrl+Z) or restore from version history, switch to Plan and resubmit. |
| KPI cards do not change when a slicer is used. | Cards built with COUNTIFS on tblApps ignore slicers. Ask Copilot to base them on the PivotTables (GETPIVOTDATA) or accept that cards show the H1 total and label them so. |
| Hires appear in January. | The chart is counting by Applied_Month. Ask Copilot to group hires by Hire_Date month as the KPI definition says. |
| Copilot stops part-way or times out. | Ask it to continue from the last completed step, or split the build into KPI cards, then charts, then slicers. |

## Challenge

Ask Copilot to add a sixth KPI card for 'Hires vs target %' (hires / sum of targets) and conditional formatting that turns it red below 90%. Verify it with your own formula.

## Reflection

Which KPI definition was easiest to get subtly wrong, and how would you stop that happening next month?

---

[← Lab 4](../lab-04-visualise-pay-and-placement-distributions-with-python-in-excel/README.md) · [Lab 6 →](../lab-06-build-an-interactive-workforce-and-attrition-dashboard/README.md)
