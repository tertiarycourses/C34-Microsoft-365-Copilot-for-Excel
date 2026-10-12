# Microsoft 365 Copilot for Excel

Use Copilot in Excel for everyday Excel work, make it consistent and repeatable with workbook rules and skills, and build highly visual, interactive dashboards, practised hands-on with realistic HR, recruitment and HR-outsourcing data.

| Course detail | Information |
|---|---|
| Course code | `C34` |
| Programme | Non-WSQ |
| Duration | 1 day · 7.5 hours |
| Format | Instructor-led, with 8 hands-on labs |
| Registration | [View course details and register](https://www.tertiarycourses.com.sg/microsoft-365-copilot-for-excel.html) |
| Lab prompts | [Copy-ready prompts page](https://tertiarycourses.github.io/C34-Microsoft-365-Copilot-for-Excel/prompts/) |
| Courseware version | v3.1 · 12 October 2026 |

---

## About the course

Copilot in Excel now does far more than answer questions. You choose how it works (**Chat only**, **Plan** or **Allow editing**), pick the AI model, bring in other files, clean messy data, and write formulas, PivotTables and charts from plain English. You can also teach it your standards: **Personalization** for your own preferences, a **`.Rules` sheet** for one workbook, and **skills** for repeatable tasks. Skills include Microsoft's built-in ones (such as `@chart-design` and `@theme-design`) and custom `SKILL.md` skills that Copilot can write for you, such as a Pareto-chart skill.

This course covers all of that in three topics, and ends with an interactive dashboard built from Copilot, rules and skills together.

Every lab uses mock data from **BrightPath Talent Solutions**, a fictional Singapore recruitment, flexible-staffing and HR-outsourcing firm. Each lab ends with a **Test It** step that lists the exact values you should get, so you can check Copilot's output against them.

The content follows Microsoft Support guidance for Copilot in Excel as of October 2026.

## Learning outcomes

By the end of the course, you will be able to:

1. **Excel fundamentals with Copilot.** Use the Chat only, Plan and Allow editing modes to prepare and clean data, write formulas and XLOOKUPs, and build PivotTables, PivotCharts and charts, verifying every result.
2. **Skills and rules.** Make Copilot's work consistent and repeatable with personalisation, workbook rules (`.Rules`), built-in skills and custom skills that you create with Copilot.
3. **Interactive dashboards.** Combine Copilot, rules and skills to build a highly visual, interactive and refreshable dashboard, and verify and share it responsibly.

## Topics covered

| Topic | Content |
|---|---|
| **1. Excel Fundamentals with Copilot** | Chat only, Plan and Allow editing; the Copilot pane, data sources and the model picker; data prep and automated cleanup; formulas and XLOOKUP; PivotTables, PivotCharts and charts |
| **2. Skills and Rules** | Personalisation; workbook rules on a `.Rules` sheet and how they work; Personalization vs `.Rules` vs skills compared; built-in chart and theme skills (`@chart-design`, `@theme-design`); custom `SKILL.md` skills; using Copilot to write, upload, install and improve a skill (a Pareto-chart skill); Pareto charts |
| **3. Interactive Dashboards** | What a skill is (reusable in any workbook, custom skills installed in your OneDrive skills folder); highly visual dashboard design; KPI definitions and style in `.Rules`; Plan mode plus `@chart-design` and `@theme-design`; slicers, a timeline and KPI cards; a reusable dashboard skill; refresh testing and responsible sharing |

## Labs

Each lab has its own folder containing:

- an **instruction PDF with sample prompts**;
- `sample-prompts.txt`, with prompts ready to copy;
- the **mock-data workbook** in `data/`, with a Prompt_Log sheet already listing the lab's prompts;
- for Labs 4, 7 and 8, a `rules-template.txt` to paste into the `.Rules` sheet;
- for Labs 6 and 8, a ready-made `skills/<name>/SKILL.md` to upload if you run short of time.

All lab prompts are also on one web page with one-click **Copy** buttons: [Lab Prompts](https://tertiarycourses.github.io/C34-Microsoft-365-Copilot-for-Excel/prompts/).

See the [labs index](labs/README.md).

| # | Lab | Theme | Instructions | Mock data |
|---:|---|---|---|---|
| 1 | [Explore and Question Recruitment Pipeline Data](labs/lab-01-data-prep-and-data-quality/Lab-01-Instructions.pdf) | Data Prep & Data Quality | [PDF](labs/lab-01-data-prep-and-data-quality/Lab-01-Instructions.pdf) | [Candidate pipeline](labs/lab-01-data-prep-and-data-quality/data/BrightPath_Candidate_Pipeline.xlsx) |
| 2 | [Generate and Verify Formula Columns and XLOOKUPs for Job Requisitions](labs/lab-02-formulas-and-xlookup/Lab-02-Instructions.pdf) | Formulas & XLOOKUP | [PDF](labs/lab-02-formulas-and-xlookup/Lab-02-Instructions.pdf) | [Job requisitions](labs/lab-02-formulas-and-xlookup/data/BrightPath_Job_Requisitions.xlsx) |
| 3 | [Build PivotTables, PivotCharts and Recruitment Charts with Copilot](labs/lab-03-pivottables-pivotcharts-and-visuals/Lab-03-Instructions.pdf) | PivotTables, PivotCharts & Visuals | [PDF](labs/lab-03-pivottables-pivotcharts-and-visuals/Lab-03-Instructions.pdf) | [Recruitment metrics](labs/lab-03-pivottables-pivotcharts-and-visuals/data/BrightPath_Recruitment_Metrics.xlsx) |
| 4 | [Clean Messy Client Orders with Cleanup, Personalisation and Rules](labs/lab-04-workbook-rules-and-personalisation/Lab-04-Instructions.pdf) | Workbook Rules & Personalisation | [PDF](labs/lab-04-workbook-rules-and-personalisation/Lab-04-Instructions.pdf) | [Client orders (raw)](labs/lab-04-workbook-rules-and-personalisation/data/BrightPath_Client_Orders_Raw.xlsx) |
| 5 | [Plot and Style Client Growth Charts with Built-in Visual Skills](labs/lab-05-built-in-skills-visual-design/Lab-05-Instructions.pdf) | Built-in Skills: Visual Design | [PDF](labs/lab-05-built-in-skills-visual-design/Lab-05-Instructions.pdf) | [HR-outsourcing client growth](labs/lab-05-built-in-skills-visual-design/data/BrightPath_HR_Outsourcing_Finance.xlsx) |
| 6 | [Use Copilot to Build a Custom Pareto-Chart Skill](labs/lab-06-custom-skills-pareto-chart/Lab-06-Instructions.pdf) | Custom Skills: Pareto Chart | [PDF](labs/lab-06-custom-skills-pareto-chart/Lab-06-Instructions.pdf) | [Job postings (500 rows)](labs/lab-06-custom-skills-pareto-chart/data/BrightPath_Job_Postings.xlsx) |
| 7 | [Build an Interactive Recruitment Dashboard with Rules and Skills](labs/lab-07-interactive-dashboard/Lab-07-Instructions.pdf) | Interactive Dashboard | [PDF](labs/lab-07-interactive-dashboard/Lab-07-Instructions.pdf) | [Dashboard data](labs/lab-07-interactive-dashboard/data/BrightPath_Recruitment_Dashboard_Data.xlsx) |
| 8 | [Turn the Dashboard into a Reusable Skill, Refresh It and Share It Safely](labs/lab-08-dashboard-skill-refresh-and-share/Lab-08-Instructions.pdf) | Dashboard Skill, Refresh & Share | [PDF](labs/lab-08-dashboard-skill-refresh-and-share/Lab-08-Instructions.pdf) | [Workforce headcount](labs/lab-08-dashboard-skill-refresh-and-share/data/BrightPath_Workforce_Headcount.xlsx) |

### Before you start

- You need a current Excel for Microsoft 365 desktop app (Windows or Mac) and a work account with a Microsoft 365 Copilot licence.
- Copy each lab workbook into a local `work-c34` folder and open it in Excel; no OneDrive upload is needed. OneDrive is used only for **custom skills** (Labs 6 and 8), which Excel reads from your OneDrive skills folder.
- In Lab 6, draft the Pareto-chart skill in [Microsoft 365 Copilot on the web](https://copilot.cloud.microsoft/), upload `SKILL.md` to your OneDrive skills folder, install it in Excel and run it. In Lab 8, draft the dashboard skill the same way. Each creation prompt includes the workbook context, so no upload is needed for drafting.
- Skills and workbook rules are fully supported only when Excel's display language is **English**.
- Copilot's menus and skills are still rolling out, so some labels may differ on your screen. Each lab includes troubleshooting tips.

## Courseware package

| Item | Files |
|---|---|
| Slide deck | [PowerPoint](<courseware/Microsoft 365 Copilot for Excel-v3.1.pptx>) · [PDF](<courseware/Microsoft 365 Copilot for Excel-v3.1.pdf>) |
| Learner Guide | [Word](<courseware/LG-Microsoft 365 Copilot for Excel.docx>) · [PDF](<courseware/LG-Microsoft 365 Copilot for Excel.pdf>) · [Markdown](<LG-Microsoft 365 Copilot for Excel.md>) |
| Lesson Plan | [Word](<courseware/LP-Microsoft 365 Copilot for Excel.docx>) · [PDF](<courseware/LP-Microsoft 365 Copilot for Excel.pdf>) |
| Labs | [labs/](labs/README.md): 8 lab folders with instruction PDFs, sample prompts, mock data, rules templates and skills |

The deck, Learner Guide, Lesson Plan and labs are all generated from one content source in `.agents/skills/non-wsq-courseware-build/build/`, so they stay aligned. To rebuild them, run `build_courseware.sh` in that folder.

## Distribution boundary

This repository publishes the current learner-facing courseware and labs. It does not publish superseded versions (`courseware/archive/`), source reference material or credentials. This is a non-WSQ commercial course, so it has no assessment.

## Data notice

All BrightPath organisations, people, IDs and values are **fictional training data**, generated for this course. Any resemblance to a real company or person is coincidental. Do not paste real employees' personal data into Copilot during practice.

## Provider

Developed and delivered by **Tertiary Infotech Academy Pte Ltd** (Tertiary Courses), Singapore · UEN 201200696W
Course page: <https://www.tertiarycourses.com.sg/microsoft-365-copilot-for-excel.html>
