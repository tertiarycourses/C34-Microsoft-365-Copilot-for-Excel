# Agentic Ai for Financial Data Analysis

**Datin edition v1.6 · C34 · 24 November 2026**

A one-day class for nine accounting professionals from West Africa. Work as the finance team of a fictional wholesale distributor in Tema, Ghana. Apply the Business Model Canvas to the case, use Meta Muse for the close brief, and build and repeat a verified sales dashboard in Excel. In Topic 4, review the published Bank of Ghana statements with Copilot, check the evidence in Word and apply five AI security controls.

**Class hours** 11:00 am–6:00 pm, with lunch from 12:30–1:30 pm. Afternoon tea breaks follow Topic 1 at 2:05–2:20 pm and Topic 2 at 3:45–4:00 pm. Instruction totals 5½ hours.

[Open the Datin prompt page on GitHub Pages](https://tertiarycourses.github.io/C34-Microsoft-365-Copilot-for-Excel/datin/prompts/) for all copy-ready prompts. An [offline copy](prompts/index.html) is included in this folder. Keep the folder structure when sharing the offline copy so it can open the lab files and workbooks.

## Course materials

| Material | Files |
|---|---|
| Trainer deck | [PowerPoint](<courseware/Agentic Ai for Financial Data Analysis-v1.6.pptx>) · [PDF](<courseware/Agentic Ai for Financial Data Analysis-v1.6.pdf>) |
| Interactive Topic 4 | [PowerPoint](<courseware/Topic 4 - Bank of Ghana Financial Report Review.pptx>) · Open in Slide Show to use the evidence routes and worked checks |
| Learner Guide | [Word](<courseware/LG-Agentic Ai for Financial Data Analysis.docx>) · [PDF](<courseware/LG-Agentic Ai for Financial Data Analysis.pdf>) |
| Lesson Plan | [Word](<courseware/LP-Agentic Ai for Financial Data Analysis.docx>) · [PDF](<courseware/LP-Agentic Ai for Financial Data Analysis.pdf>) |

## Labs

| Topic | Lab files |
|---|---|
| **1. Business Innovation with Agentic AI and AI Agents** | [Meta Muse PDF](labs/lab-01-meta-muse/Lab-01-Instructions.pdf) · [Copy prompts](prompts/index.html#lab-1) |
| **2. Create an Interactive Dashboard in Excel** | [Dashboard PDF](labs/lab-02-interactive-sales-dashboard/Lab-02-Instructions.pdf) · [H1 mock data](labs/lab-02-interactive-sales-dashboard/data/Datin_Sales_Finance_H1_GHS.xlsx) · [Working dark dashboard](labs/lab-02-interactive-sales-dashboard/examples/Datin_Sales_Dashboard_Example_GHS.xlsx) · [Copy prompts](prompts/index.html#lab-2) |
| **3. Skills** | [Skill PDF](labs/lab-03-dashboard-skill/Lab-03-Instructions.pdf) · [H2 mock data](labs/lab-03-dashboard-skill/data/Datin_Sales_Finance_H2_GHS.xlsx) · [SKILL.md](labs/lab-03-dashboard-skill/skills/datin-dark-sales-dashboard/SKILL.md) · [Copy prompts](prompts/index.html#lab-3) |
| **4. Analyse a Financial Report with Copilot** | [Review instructions](labs/lab-04-bog-financial-report/Lab-04-Instructions.pdf) · [Word workpaper with worked findings](<labs/lab-04-bog-financial-report/Bank of Ghana 2025 - Financial Report Review.docx>) · [Copy prompts and learner worksheet](https://tertiarycourses.github.io/C34-Microsoft-365-Copilot-for-Excel/datin/prompts/#lab-4) |

For Excel, use the Outlook training account assigned by the trainer. Check that Copilot editing, OneDrive and **Manage plugins & skills** are available. For Meta Muse, use the separate trainer-approved demo Gmail and Google Calendar.

The H1 and H2 starters each contain 432 fictional orders. The finished workbook uses Region, Channel and Product controls to drive the monthly line, regional bar, product doughnut and channel doughnut charts. Check All, Greater Accra and Wholesale, then add Rice. The Data Guide lists the expected results for all three selections.

## References

- Topic 1 reuses the native v57 teaching slides and applies their design throughout. Selected slides use finance examples and the Datin Meta Muse prompts. Reference [Business Innovation with Agentic AI and AI Agents](https://github.com/tertiarycourses/TGS-2023037472-Business-Innovation-with-Agentic-AI-and-AI-Agents).
- Topic 2 uses the create, check and restyle sequence from [Kevin Stratvert’s Excel dashboard demonstration](https://www.youtube.com/watch?v=YgcEesCDIkw).
- Topic 3 follows the custom skill folder, refresh, toggle and invocation guidance in [Microsoft Support](https://support.microsoft.com/en-gb/excel/copilot/copilot-in-excel-skills).
- Topic 4 uses [Bank of Ghana’s 2025 financial statements](https://www.bog.gov.gh/wp-content/uploads/2026/05/2025-Financial-Statements-1.pdf) and [Microsoft’s guide to attaching files in Copilot Chat](https://support.microsoft.com/en-au/microsoft-365-copilot/add-content-to-microsoft-copilot-chat-prompts).

## Rebuild

From the C34 project root, run `bash courseware_ghana/source/build.sh`. The Datin content and data are in `source/`. The build uses the reference v57 PowerPoint, its slide design components and the C34 document generators. If your reference repository checkout is in a different location, set `DATIN_REFERENCE_REPO` to its root. The build requires Python with python-pptx, python-docx, ReportLab, XlsxWriter, Pillow and PyMuPDF, plus LibreOffice’s `soffice`.
