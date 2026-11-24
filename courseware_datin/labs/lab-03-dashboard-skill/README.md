# Lab 3 — Install and Reuse the Dashboard Skill

[Instructions PDF](Lab-03-Instructions.pdf) · [Copy prompts](../../prompts/index.html#lab-3)

## Goal

Convert Prompt 2 into a reusable skill and use a short request to create a verified dashboard on fresh data.

## Prerequisites

- Have Prompt 2 and data/Datin_Sales_Finance_H2_GHS.xlsx open as a fresh OneDrive working copy.
- Use Excel with Manage plugins & skills available on your work account.

## Steps

### 1. In Copilot in Excel, use @skill-builder when available to draft a skill from Prompt 2. Review the name, trigger, theme and preservation checks. Check that the draft defines both paths: restyle an existing dashboard or build one if a create request finds none. Choose Save skill once the proposed files are correct.

```text
@skill-builder Turn the theme prompt below into an Excel skill named datin-dark-sales-dashboard. For an existing interactive sales dashboard, apply the theme while preserving its layout, source data, formulas, charts and working filters. If the request is to create a dashboard and none exists, first build one from the workbook sales table with GHS Net Sales, Gross Profit, Gross Margin and Order Count, monthly trend and regional charts, doughnut charts for product and channel mix, and Region, Channel and Product filters. Apply the theme and verify the filters. Show the proposed skill for approval. Theme prompt: Redesign the dashboard with a modern dark theme using navy blue backgrounds, gold accents, and high-contrast white text. Keep the existing layout, charts, metrics, and interactive filters fully functional. Make sure all labels and charts remain easy to read.
```

### 2. For the manual upload route, use the reviewed file or the provided skills/datin-dark-sales-dashboard/SKILL.md. In the Copilot pane choose … > Manage plugins & skills > Custom skills. Select Create OneDrive folder if needed, then Open skills folder. Upload the entire datin-dark-sales-dashboard folder, with SKILL.md inside it. If you saved through @skill-builder, confirm the file is in OneDrive and skip the manual upload.

### 3. Return to Custom skills and select Refresh. Find datin-dark-sales-dashboard under Manage plugins & skills and turn it on. Open the fresh workbook. Start a new Copilot chat, allow editing and send the short prompt.

```text
Create an interactive Datin sales dashboard.
```

### 4. Check the theme and controls. Record the skill name if Copilot shows which skill it used. For a controlled run, or if the short request misses the theme, start from a fresh workbook copy and select the skill from + > Choose plugins and skills or use this explicit request.

```text
@datin-dark-sales-dashboard Create an interactive Datin sales dashboard.
```

### 5. Check the July–December months and the three filter states below. Verify that the charts and KPI values use GHS, every chart follows the filters, and the source data is intact. Restore All and save the result.

## Test it

- **All regions and channels** · GHS 1,732,810 net sales, GHS 665,335 gross profit, 38.40% margin, 432 orders
- **Greater Accra plus Wholesale** · GHS 361,981 net sales, GHS 136,952 gross profit, 37.83% margin, 72 orders
- **Greater Accra plus Wholesale plus Rice** · GHS 42,525 net sales, GHS 15,390 gross profit, 36.19% margin, 12 orders
- **Skill discovery** · datin-dark-sales-dashboard appears after Refresh and is switched on.
- **Fresh dashboard** · July–December 2026 appears in the trend. All four charts react to filters and the theme is applied.

## Checkpoint

Keep the skill folder and the completed H2 working copy. Use the fresh starter again if you want to repeat the short-prompt test.

## Files

- [Datin_Sales_Finance_H2_GHS.xlsx](data/Datin_Sales_Finance_H2_GHS.xlsx)
- [Ready-To-Upload Skill](skills/datin-dark-sales-dashboard/SKILL.md)
- [Microsoft’s Excel skills guide](https://support.microsoft.com/en-gb/excel/copilot/copilot-in-excel-skills)
