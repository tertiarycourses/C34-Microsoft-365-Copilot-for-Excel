# Microsoft 365 Copilot for Excel — Hands-On Labs

8 labs across 3 topics · 1 day · 7.5 instructional hours

Every lab has its own folder containing an **instruction PDF with sample prompts**, a copy-ready `sample-prompts.txt`, and a `data/` folder with the **mock HR / recruitment workbook**. All data belongs to *BrightPath Talent Solutions*, a fictional Singapore recruitment and staffing firm.

**Prompts web page:** every lab's prompts with one-click Copy buttons — [https://tertiarycourses.github.io/C34-Microsoft-365-Copilot-for-Excel/prompts/](https://tertiarycourses.github.io/C34-Microsoft-365-Copilot-for-Excel/prompts/)

Work through the labs in order. Each lab opens its own workbook, so you can rejoin at any lab.


## Topic 1 — Excel Fundamentals with Copilot

| # | Theme | Lab | Instructions | Mock data | You Build |
|---|-------|-----|--------------|-----------|-----------|
| 1 | Data Prep & Data Quality | [Explore and Question Recruitment Pipeline Data](lab-01-data-prep-and-data-quality/Lab-01-Instructions.pdf) | [PDF](lab-01-data-prep-and-data-quality/Lab-01-Instructions.pdf) · [prompts](lab-01-data-prep-and-data-quality/sample-prompts.txt) | [BrightPath_Candidate_Pipeline.xlsx](lab-01-data-prep-and-data-quality/data/BrightPath_Candidate_Pipeline.xlsx) | BrightPath_Candidate_Pipeline.xlsx in your work-c34 folder with table tblApplications, a Source_Clean column, duplicate IDs highlighted, Hired rows highlighted, a verified source-by-hires answer and a filled Prompt_Log |
| 2 | Formulas & XLOOKUP | [Generate and Verify Formula Columns and XLOOKUPs for Job Requisitions](lab-02-formulas-and-xlookup/Lab-02-Instructions.pdf) | [PDF](lab-02-formulas-and-xlookup/Lab-02-Instructions.pdf) · [prompts](lab-02-formulas-and-xlookup/sample-prompts.txt) | [BrightPath_Job_Requisitions.xlsx](lab-02-formulas-and-xlookup/data/BrightPath_Job_Requisitions.xlsx) | BrightPath_Job_Requisitions.xlsx with Recruiter_Name, Days_Open, Fill_Rate, Est_Fee and Ageing_Flag columns in tblRequisitions, overdue requisitions highlighted, and a Control sheet that reconciles the totals |
| 3 | PivotTables, PivotCharts & Visuals | [Build PivotTables, PivotCharts and Recruitment Charts with Copilot](lab-03-pivottables-pivotcharts-and-visuals/Lab-03-Instructions.pdf) | [PDF](lab-03-pivottables-pivotcharts-and-visuals/Lab-03-Instructions.pdf) · [prompts](lab-03-pivottables-pivotcharts-and-visuals/sample-prompts.txt) | [BrightPath_Recruitment_Metrics.xlsx](lab-03-pivottables-pivotcharts-and-visuals/data/BrightPath_Recruitment_Metrics.xlsx) | BrightPath_Recruitment_Metrics.xlsx with PivotTables and four refined charts on the Charts sheet: a monthly hires line PivotChart, a sorted 2026 hires-by-sector bar PivotChart, a 2026 recruitment funnel and a hire-rate vs cost-per-hire combo chart |

## Topic 2 — Skills and Rules

| # | Theme | Lab | Instructions | Mock data | You Build |
|---|-------|-----|--------------|-----------|-----------|
| 4 | Workbook Rules & Personalisation | [Clean Messy Client Orders with Cleanup, Personalisation and Rules](lab-04-workbook-rules-and-personalisation/Lab-04-Instructions.pdf) | [PDF](lab-04-workbook-rules-and-personalisation/Lab-04-Instructions.pdf) · [prompts](lab-04-workbook-rules-and-personalisation/sample-prompts.txt) | [BrightPath_Client_Orders_Raw.xlsx](lab-04-workbook-rules-and-personalisation/data/BrightPath_Client_Orders_Raw.xlsx)<br>[rules-template.txt](lab-04-workbook-rules-and-personalisation/rules-template.txt) | Your Personalization (saved to your account) and BrightPath_Client_Orders_Raw.xlsx with a .Rules sheet, a clean tblOrdersClean table (5 statuses, 5 regions, 12 clients), Rate, Delivered_Value, Budget_Used and Over_Budget columns, and a Checks sheet that proves the clean-up |
| 5 | Built-in Skills: Visual & Finance | [Analyse HR-Outsourcing Finances with Built-in Skills](lab-05-built-in-skills-visual-and-finance/Lab-05-Instructions.pdf) | [PDF](lab-05-built-in-skills-visual-and-finance/Lab-05-Instructions.pdf) · [prompts](lab-05-built-in-skills-visual-and-finance/sample-prompts.txt) | [BrightPath_HR_Outsourcing_Finance.xlsx](lab-05-built-in-skills-visual-and-finance/data/BrightPath_HR_Outsourcing_Finance.xlsx) | BrightPath_HR_Outsourcing_Finance.xlsx with a skill-designed MRR chart, a Unit_Economics sheet, a corrected Headcount_Budget, a Checks sheet that agrees with the skills, and a consistent theme |
| 6 | Custom Skills: Calendar | [Use Copilot to Build a Custom Interview-Calendar Skill](lab-06-custom-skills-calendar/Lab-06-Instructions.pdf) | [PDF](lab-06-custom-skills-calendar/Lab-06-Instructions.pdf) · [prompts](lab-06-custom-skills-calendar/sample-prompts.txt) | [BrightPath_Interview_Schedule.xlsx](lab-06-custom-skills-calendar/data/BrightPath_Interview_Schedule.xlsx)<br>[skills/interview-calendar/SKILL.md](lab-06-custom-skills-calendar/skills/interview-calendar/SKILL.md) | An interview-calendar skill in your OneDrive skills folder, and BrightPath_Interview_Schedule.xlsx with Cal_2026-07-06 and Cal_2026-07-13 calendar sheets, each with clashes flagged and a summary |

## Topic 3 — Interactive Dashboards with Copilot, Skills and Rules

| # | Theme | Lab | Instructions | Mock data | You Build |
|---|-------|-----|--------------|-----------|-----------|
| 7 | Interactive Dashboard | [Build an Interactive Recruitment Dashboard with Rules and Skills](lab-07-interactive-dashboard/Lab-07-Instructions.pdf) | [PDF](lab-07-interactive-dashboard/Lab-07-Instructions.pdf) · [prompts](lab-07-interactive-dashboard/sample-prompts.txt) | [BrightPath_Recruitment_Dashboard_Data.xlsx](lab-07-interactive-dashboard/data/BrightPath_Recruitment_Dashboard_Data.xlsx)<br>[rules-template.txt](lab-07-interactive-dashboard/rules-template.txt) | BrightPath_Recruitment_Dashboard_Data.xlsx with a .Rules sheet, a Checks sheet, PivotTables on Dash_Data and a Dashboard sheet with six formula-driven KPI cards (one with traffic-light colours), four PivotCharts, Sector, Source and Recruiter slicers and a Hire_Date timeline |
| 8 | Dashboard Skill, Refresh & Share | [Turn the Dashboard into a Reusable Skill, Refresh It and Share It Safely](lab-08-dashboard-skill-refresh-and-share/Lab-08-Instructions.pdf) | [PDF](lab-08-dashboard-skill-refresh-and-share/Lab-08-Instructions.pdf) · [prompts](lab-08-dashboard-skill-refresh-and-share/sample-prompts.txt) | [BrightPath_Workforce_Headcount.xlsx](lab-08-dashboard-skill-refresh-and-share/data/BrightPath_Workforce_Headcount.xlsx)<br>[rules-template.txt](lab-08-dashboard-skill-refresh-and-share/rules-template.txt)<br>[skills/hr-workforce-dashboard/SKILL.md](lab-08-dashboard-skill-refresh-and-share/skills/hr-workforce-dashboard/SKILL.md) | An hr-workforce-dashboard skill in OneDrive, and BrightPath_Workforce_Headcount.xlsx with a .Rules sheet, helper columns, a Checks sheet and an HR_Dashboard with six KPI cards, four charts and Location and Job_Level slicers - refresh-tested and privacy-checked |

---

> Use only the supplied fictional BrightPath data or information you are authorised to process. Do not paste credentials, confidential business material or real employees' personal data into Copilot, rules or skills.


_Tertiary Infotech Academy Pte Ltd · C34 · v3.0 (11 October 2026)_
