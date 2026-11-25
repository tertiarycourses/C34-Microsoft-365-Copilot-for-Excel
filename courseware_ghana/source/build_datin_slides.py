"""Make the Datin deck with the published v57 slide parts and design components."""

import sys
from pathlib import Path

from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.util import Inches

import ccdeck as K
import course_data as C
from agent_quote import apply_agent_quote
from bog_slides import STANDALONE, add_topic, button, security_slide
from datin_poll import INSTRUCTIONS as POLL_INSTRUCTIONS, build_join_slide
from datasets import FIRST, FRESH
from labs import CREATE, SHORT, THEME, EXPLICIT, SKILL_NAME, VIDEO_URL, checks
from reference_topic import EDITS
from tema_case import COURSE_MAP, GENERAL, TOPIC_1_HANDOFF, TOPIC_2_HANDOFF, decision_slide, patch_reference, patch_text

REF_TITLE = "Business Innovation with Agentic AI and AI Agents"
REF_CODE = "TGS-2023037472"
SIGN_IN_CARDS = [
    ("Get your sign-in", "Ask the trainer for your assigned training account."),
    ("Check Copilot", "Open Excel and check that Copilot editing is available."),
]
SIGN_IN_TIP = "Use the access details provided by the trainer. Sign out after class."


def set_text(shape, value):
    runs = [run for paragraph in shape.text_frame.paragraphs for run in paragraph.runs]
    if runs[0].text in {"TIP  ", "WHY  "} and value.startswith(runs[0].text) and len(runs) >= 2:
        runs = runs[1:]
        value = value[5:]
    runs[0].text = value
    for run in runs[1:]:
        run.text = ""


class DatinDeck(K.Deck):
    def __init__(self, reference):
        super().__init__(C.TITLE, C.COURSE_CODE, C.SHORT_TITLE)
        self.prs = Presentation(str(reference))
        self.blank = self.prs.slide_layouts[6]
        self.slide_ids = self.prs.slides._sldIdLst
        self.reference = list(zip(self.prs.slides, list(self.slide_ids)))
        self.order = []

    def new(self, dark=False):
        slide = super().new(dark)
        self.order.append(self.slide_ids[-1])
        return slide

    def copy(self, number, course_map):
        slide, identifier = self.reference[number - 1]
        self.order.append(identifier)
        self.n += 1
        remaining = dict(EDITS.get(number, {}))
        for shape in slide.shapes:
            if not shape.has_text_frame or not shape.text.strip():
                continue
            old = shape.text
            if number == 31 and old.startswith("TIP  Then Settings › General › Usage › Redeem invite code:"):
                set_text(shape, "TIP  The Outlook training accounts are for Excel. Use the trainer’s approved demo Gmail and Calendar for Muse.")
            elif old in remaining:
                set_text(shape, remaining.pop(old))
            elif REF_CODE in old:
                set_text(shape, old.replace(REF_TITLE, C.SHORT_TITLE).replace(REF_CODE, C.COURSE_CODE))
            elif old == str(number) and shape.top >= Inches(7):
                set_text(shape, str(self.n))
        if remaining:
            raise ValueError(f"Reference v57 slide {number} did not match replacements: {list(remaining)}")

        for relation in list(slide.part.rels.values()):
            if relation.reltype == RT.SLIDE and relation.target_part == self.reference[11][0].part:
                new_id = slide.part.relate_to(course_map.part, RT.SLIDE)
                for element in slide.element.iter():
                    for key, value in list(element.attrib.items()):
                        if key.startswith("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}") and value == relation.rId:
                            element.set(key, new_id)
                slide.part.drop_rel(relation.rId)

        # Keep the learner-facing source slides and their links. Drop original
        # trainer notes for the slides where the instructions were rewritten.
        if number in EDITS:
            for relation in list(slide.part.rels.values()):
                if relation.reltype == RT.NOTES_SLIDE:
                    slide.part.drop_rel(relation.rId)
        return slide

    def finish(self, out):
        keep = {slide_id.rId for slide_id in self.order}
        for identifier in list(self.slide_ids):
            self.slide_ids.remove(identifier)
            if identifier.rId not in keep:
                self.prs.part.drop_rel(identifier.rId)
        for identifier in self.order:
            self.slide_ids.append(identifier)
        self.prs.core_properties.title = C.TITLE
        self.prs.core_properties.subject = "Datin edition • Finance and accounting professionals from West Africa"
        self.prs.core_properties.author = C.ORG
        self.save(out)


def photo(deck, kicker, title, image, points, *, dark=False):
    s = deck.new(dark=dark)
    K.header(deck, s, kicker, title)
    K.picture(s, str(image), 0.7, 1.8, maxw=7.5, maxh=4.95)
    for index, (heading, body, icon) in enumerate(points):
        y = 2.05 + index * 2.23
        K.card(s, 8.42, y, 4.27, 1.98, fill=K.PANEL if dark else K.WHITE, shadow=not dark)
        K.icon(s, icon, 8.67, y + 0.2, 0.52)
        K.text(s, 9.32, y + 0.2, 3.13, 0.42, heading, 19, K.WHITE if dark else K.INK,
               font=K.SERIF, bold=True, minpt=17)
        K.text(s, 8.68, y + 0.86, 3.71, 0.88, body, 17, K.SOFT if dark else K.INK, minpt=16)
    return s


def prompt(deck, kicker, title, copy, action, extra_url=None):
    s = deck.new()
    K.header(deck, s, kicker, title)
    K.card(s, 0.62, 2.0, 12.06, 2.7, fill=K.CHAR)
    K.text(s, 0.92, 2.28, 11.44, 2.12, copy, 22, K.WHITE, anchor="m", minpt=19)
    K.icon(s, "check", 0.77, 5.14, 0.57)
    K.text(s, 1.48, 5.12, 11.1, 1.0, action, 18, K.INK, minpt=17)
    if extra_url:
        shape = K.text(s, 1.48, 6.28, 11.1, 0.45, "Watch the Excel dashboard demonstration", 16, K.TEAL, bold=True)
        shape.text_frame.paragraphs[0].runs[0].hyperlink.address = extra_url


def agenda(deck):
    schedule = C.SCHEDULE(lambda _: "")[1][1]
    labels = [
        "Set up Excel and Muse", "Topic 1 · AI agents for finance", "Lunch",
        "Topic 1 · Tema canvas demo", "Lab 1 · Meta Muse", "Tea break after Topic 1",
        "Topic 2 · Excel dashboards", "Lab 2 · Build, style and verify", "Tea break after Topic 2",
        "Topic 3 · Reusable Excel skills", "Lab 3 · Install and reuse skill",
        "Topic 4 · Review a financial report", "Lab 4 · Review and document", "Recap and AI security",
    ]
    s = deck.new()
    K.header(deck, s, "Lesson plan", "11:00 am–6:00 pm", "5½ hours of instruction · 1 hour for lunch · two 15-minute afternoon tea breaks")
    for i, ((start, end, _, kind, _), label) in enumerate(zip(schedule, labels, strict=True)):
        col, row = divmod(i, 7)
        x, y = 0.65 + 6.15 * col, 2.40 + row * 0.61
        if kind in {"break", "lunch"}:
            K.card(s, x, y - 0.045, 5.98, 0.55, fill=K.CHAR, shadow=False)
        fg = K.WHITE if kind in {"break", "lunch"} else K.INK
        K.text(s, x + 0.12, y + 0.04, 1.53, 0.37, f"{start}–{end}", 14, K.ACC, bold=True, autofit=False)
        K.text(s, x + 1.7, y + 0.03, 4.1, 0.42, label, 16.5, fg, minpt=15.5)


def kpis(deck, period, title):
    for index, (label, result) in enumerate(checks(period)):
        amounts = result.split(", ")
        s = deck.new(dark=index == 1)
        K.header(deck, s, "Reconcile to Data Guide", title, label)
        items = [
            ("Net sales", amounts[0].removesuffix(" net sales"), "coins"),
            ("Gross profit", amounts[1].removesuffix(" gross profit"), "trend"),
            ("Gross margin", amounts[2].removesuffix(" margin"), "chart"),
            ("Orders", amounts[3].removesuffix(" orders"), "list"),
        ]
        for j, (metric, amount, icon) in enumerate(items):
            x = 0.63 + 3.07 * j
            K.card(s, x, 2.6, 2.88, 2.62, fill=K.PANEL if deck.dark else K.WHITE, shadow=not deck.dark)
            K.icon(s, icon, x + 0.2, 2.87, 0.59)
            K.text(s, x + 0.2, 3.68, 2.48, 0.52, metric, 17, K.SOFT if deck.dark else K.MUTED, autofit=False)
            K.text(s, x + 0.2, 4.25, 2.48, 0.57, amount, 24, K.WHITE if deck.dark else K.INK,
                   font=K.SERIF, bold=True, minpt=22)
        K.text(s, 0.68, 5.64, 12.0, 0.6, "Apply the same selection to the dashboard and check every KPI and chart.", 18,
               K.SOFT if deck.dark else K.MUTED)


def main():
    reference, logo, visuals, screenshot, out = map(Path, sys.argv[1:])
    D = DatinDeck(reference)
    K.cover(D, "Datin edition · C34", "Agentic Ai for\nFinancial Data Analysis", f"Version {C.VERSION[1:]}  ·  {C.VERSION_DATE}",
            f"{C.ORG}\n{C.PARTICIPANTS}",
            "www.tertiarycourses.com.sg   |   enquiry@tertiaryinfotech.com   |   +65 6100 0613",
            str(logo), badge=("AGENTIC AI", "DATIN · FINANCIAL ANALYSIS"))
    photo(D, "Datin finance team · fictional company", "Meet the Tema Wholesale Distributor",
          visuals / "tema-warehouse.png", [
              ("The business", "Sells food staples, household supplies and personal care products across Ghana.", "box"),
              ("Your role", "Review month-end exceptions and turn sales records into a reconciled GHS report.", "chart"),
          ])
    course_map = K.cards(D, "Course map", "Four Ways to Use AI in Finance", [
        ("Topic 1 · Agentic AI", "Identify agent roles, tool access and approval gates. Use Muse for the close brief.", "robot"),
        ("Topic 2 · Dashboard", "Build and restyle an Excel dashboard to explain revenue by region and product.", "chart"),
        ("Topic 3 · Skills", "Save and enable a Copilot skill. Reuse it to build on fresh H2 data.", "tool"),
        ("Topic 4 · Review a financial report", "Use Copilot to check the Bank of Ghana statements and record evidence in Word.", "file"),
    ], cols=2, desc_pt=16, bottom=6.34)
    patch_text(course_map, COURSE_MAP)
    agenda(D)
    K.cards(D, "Excel Copilot · sign in", "Use Your Assigned Outlook Training Account",
            SIGN_IN_CARDS, cols=2, desc_pt=21, bottom=5.35, tipt=SIGN_IN_TIP)
    K.flow(D, "Class materials", "How You'll Learn", [
        ("folder", "Open Datin", "Use courseware_ghana."),
        ("book", "Open your lab", "Read the lab PDF."),
        ("edit", "Copy a prompt", "Open prompts/index.html."),
        ("check", "Verify", "Check sources and totals."),
    ])

    # Reuse the published Topic 1 slide parts and their media in the Datin deck.
    for n in range(22, 31):
        slide = D.copy(n, course_map)
        patch_reference(slide, n)
        if n == 22:
            patch_text(slide, GENERAL[7])
        if n == 23:
            vote_hint = next(shape for shape in slide.shapes if shape.has_text_frame and
                             shape.text == "Raise 1, 2, 3 or 4 fingers, or type A, B, C or D in the chat.")
            set_text(vote_hint, POLL_INSTRUCTIONS)
            join_slide = D.new(dark=True)
            build_join_slide(join_slide, D.n, C.TITLE, C.COURSE_CODE,
                             image_path=visuals / "livepoll-230174.png")
    K.compare(D, "Topic 1 · tools and skills", "Tools Act. Skills Guide the Work.", (
        "Tool · a callable operation", "tool", [
            "Find a matching invoice in the authorised ledger.",
            "Read the demo calendar or insert a chart in Excel.",
        ]
    ), ("Skill · a reusable procedure", "book", [
        "Define how to review a duplicate invoice and route it.",
        "Define what the dashboard must show and how to verify it.",
    ]), tipt="Tool permissions still govern what Copilot or Muse can do.")
    for n in range(36, 57):
        slide = D.copy(n, course_map)
        if n == 36:
            apply_agent_quote(slide, visuals)
    K.statement(D, "Lunch · 12:30–1:30 pm", "Lunch Break", "Topic 1 resumes at 1:30 pm.")
    for n in (58, 57, 75, *range(59, 66), 72, 73):
        patch_reference(D.copy(n, course_map), n)
    decision_slide(D, D.new(), K)
    lab_one = photo(D, "Topic 1 · Lab 1", "Month-End Exceptions in Tema", visuals / "tema-dispatch.png", [
        ("Cut-off", "Check proof of delivery against a GHS 4,800 rice invoice.", "clock"),
        ("Approval queue", "Flag a possible supplier duplicate and an unsupported customer credit.", "shield"),
    ])
    patch_text(lab_one, GENERAL[59])
    for n in (31, 33, 34, 35):
        D.copy(n, course_map)
    K.statement(D, "Tea break 1 · 2:05–2:20 pm", "Tea Break", TOPIC_1_HANDOFF)

    K.section(D, "Topic 2", "Create an Interactive\nDashboard in Excel",
              ["Reconcile the source", "Create an interactive view", "Apply the dark theme", "Retest the controls"])
    finance_photo = photo(D, "Topic 2 · financial reporting", "What Changed in Revenue and Margin?",
          visuals / "ghana-finance-team.png", [
              ("Compare performance", "Show how net revenue moves over time, by region and by product.", "trend"),
              ("Explain margin", "Check the mix and COGS before drawing conclusions about profitability.", "coins"),
          ])
    patch_text(finance_photo, GENERAL[66])
    K.cards(D, "Topic 2 · accounting definition", "From Order Values to Gross Margin", [
        ("Net revenue", "60 units × GHS 100 less GHS 300 discount = GHS 5,700.", "coins"),
        ("Gross profit", "GHS 5,700 revenue less GHS 3,600 COGS = GHS 2,100.", "chart"),
        ("Gross margin", "GHS 2,100 ÷ GHS 5,700 = 36.84%. Divide the two totals.", "trend"),
    ], cols=3, desc_pt=20)
    K.flow(D, "Topic 2 · data flow", "One Set of Filters Drives the Dashboard", [
        ("folder", "Source table", "One row per order."),
        ("tool", "Filters", "Region, Channel and Product."),
        ("chart", "Summaries", "Four KPIs and four charts."),
        ("check", "Reconcile", "Control totals from Data Guide."),
    ], dark=True)
    K.cards(D, "Topic 2 · chart choice", "Four Views of Net Revenue", [
        ("Line · trend", "Compare monthly net revenue across Jan–Jun.", "trend"),
        ("Bars · region", "Compare Greater Accra, Ashanti and Western.", "chart"),
        ("Doughnut · product", "Show each product’s share of the selected revenue.", "coins"),
        ("Doughnut · channel", "Compare the Wholesale and Retail revenue mix.", "layers"),
    ], cols=2, desc_pt=18, bottom=6.68)
    K.image(D, "Topic 2 · working Excel example", "Datin Sales Performance Dashboard", str(screenshot))
    for candidate, title in [(screenshot.with_name("dashboard-top.png"), "Read the Trend and Product Mix"),
                             (screenshot.with_name("dashboard-bottom.png"), "Compare Regional and Channel Revenue")]:
        K.image(D, "Topic 2 · working Excel example", title, str(candidate))
    prompt(D, "Lab 2 · Prompt 1", "Build the Interactive Dashboard", CREATE,
           "Build on the H1 workbook. Run the lab follow-up to specify the trend, mix charts, region comparison and filters.", VIDEO_URL)
    prompt(D, "Lab 2 · Prompt 2", "Apply the Modern Dark Theme", THEME,
           "Save a pre-theme copy and compare identical Region, Channel and Product selections before and after.")
    kpis(D, FIRST, "Verify the H1 Dashboard")
    K.statement(D, "Tea break 2 · 3:45–4:00 pm", "Tea Break", TOPIC_2_HANDOFF)

    K.section(D, "Topic 3", "Make the Dashboard\nRepeatable with Skills",
              ["Turn Prompt 2 into instructions", "Save it in OneDrive", "Enable it in Excel", "Verify on H2 data"])
    K.compare(D, "Topic 3 · skill logic", "One Skill, Two Starting Points",
              ("Existing dashboard", "edit", ["Preserve layout and chart connections.", "Apply the colour and label standards.", "Recheck the current filter selection."]),
              ("Fresh workbook", "chart", ["Build the KPIs, filters and charts first.", "Apply the same design.", "Reconcile new figures for the second half."]))
    K.flow(D, "Topic 3 · installation", "Install the Skill Through Excel", [
        ("gear", "Manage skills", "Open Custom skills."),
        ("folder", "Open OneDrive", "Go to the Excel skills folder."),
        ("file", "Upload folder", "Keep SKILL.md inside."),
        ("loop", "Refresh", "Find the new skill."),
        ("check", "Enable", "Turn on automatic use."),
    ])
    prompt(D, "Lab 3 · reuse", "Build from a Fresh H2 Workbook", SHORT,
           f"If Copilot misses the skill, run: {EXPLICIT}")
    kpis(D, FRESH, "Verify the H2 Dashboard")
    review, memo = add_topic(D)
    button(course_map, 9.72, 6.49, 2.98, "Open Topic 4", review, accent=True)
    button(review, 0.65, 6.41, 2.30, "Course map", course_map)
    recap = K.cards(D, "Class recap", "What You Can Now Verify", [
        ("Muse", "Trace a close brief to the approved demo inbox and calendar.", "calendar"),
        ("Excel", "Reconcile net revenue and margin across region, product and channel.", "chart"),
        ("Skill", f"Enable {SKILL_NAME} and check results on a fresh workbook.", "check"),
        ("Review", "Trace Copilot findings to the published financial statements and verify them in Word.", "file"),
    ], cols=2, desc_pt=16)
    patch_text(recap, GENERAL[86])
    button(memo, 6.04, 6.24, 2.45, "Continue to recap", recap, accent=True)
    security_slide(D)
    K.statement(D, "Datin edition", "Thank You", "Take the same evidence and approval checks into your next close.")
    D.finish(out)

    topic_deck = K.Deck(C.TITLE, C.COURSE_CODE, C.SHORT_TITLE)
    review, memo = add_topic(topic_deck)
    controls = security_slide(topic_deck, return_to=review)
    button(memo, 6.04, 6.24, 3.08, "Review AI security", controls, accent=True)
    topic_deck.prs.core_properties.title = "Topic 4 · Bank of Ghana Financial Report Review"
    topic_deck.prs.core_properties.author = C.ORG
    topic_deck.save(out.with_name(STANDALONE))


if __name__ == "__main__":
    main()
