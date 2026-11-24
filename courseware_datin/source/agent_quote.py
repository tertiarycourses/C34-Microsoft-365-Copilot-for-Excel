"""Build the Datin leadership quotations slide and record its sources."""

from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR
from pptx.oxml.ns import qn
from pptx.util import Inches

from datin_poll import text

WHITE_HOUSE_MEETING = "https://commons.wikimedia.org/wiki/File:V20230504LJ-0199-2.jpg"

QUOTES = (
    (
        "Sundar Pichai", "CEO of Google and Alphabet",
        "The conversation has gone from ‘Can we build an agent?’ to ‘How do we manage thousands of them?’",
        "https://blog.google/innovation-and-ai/infrastructure-and-cloud/google-cloud/cloud-next-2026-sundar-pichai/",
        "pichai", "Lawrence Jackson / The White House", WHITE_HOUSE_MEETING,
    ),
    (
        "Satya Nadella", "CEO of Microsoft",
        "… agents now can reason over those live transactions along with petabytes of historical records in real time.",
        "https://news.microsoft.com/events/wp-content/uploads/2026/11/Nadella_TRANSCRIPT_Ignite-Keynote_Nov-2026.pdf",
        "nadella", "Lawrence Jackson / The White House", WHITE_HOUSE_MEETING,
    ),
    (
        "Jensen Huang", "Founder and CEO of NVIDIA",
        "The age of AI Agentics is here.",
        "https://fortune.com/2025/01/07/nvidias-jensen-huang-says-ai-agents-are-a-multi-trillion-dollar-opportunity/",
        "huang", "Daniel Torok / The White House",
        "https://commons.wikimedia.org/wiki/File:Jen-Hsun_Huang_2025.jpg",
    ),
    (
        "Bill Gates", "Co-founder and former CEO of Microsoft",
        "Agents are not only going to change how everyone interacts with computers. They’re also going to upend the software industry…",
        "https://www.gatesnotes.com/ai-agents",
        "gates", "Office of U.S. Senator Chris Coons",
        "https://commons.wikimedia.org/wiki/File:Bill_Gates_-_Nov._8,_2019.jpg",
    ),
)
MOLLICK_SOURCE = "https://www.oneusefulthing.org/p/a-guide-to-which-ai-to-use-in-the?action=share"
MOLLICK_QUOTE = "The shift from chatbot to agent is the most important change in how people use AI since ChatGPT launched."


def panel(slide, x, y, width, height):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(width), Inches(height)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor.from_string("2E2C27")
    shape.line.fill.background()


def apply_agent_quote(slide, visuals):
    for shape in list(slide.shapes):
        if shape.has_text_frame and shape.text and shape.top < Inches(7):
            shape.element.getparent().remove(shape.element)

    text(slide, 0.62, 0.43, 10, 0.33, "TOPIC 1 · ALWAYS-ON AI AGENTS", 14, "D97757", bold=True)
    text(slide, 0.62, 0.83, 10.3, 0.57, "What Leaders Say About AI Agents", 32, "FFFFFF",
         font="Cambria", bold=True)

    notes = ["Connect these perspectives to which actions Datin can delegate and which require review."]
    positions = [(0.65, 1.61), (6.73, 1.61), (0.65, 3.49), (6.73, 3.49)]
    for (speaker, role, quote, source, slug, credit, photo_source), (x, y) in zip(QUOTES, positions):
        panel(slide, x, y, 5.94, 1.74 if y < 3 else 2.03)
        photo = slide.shapes.add_picture(
            str(visuals / "portraits" / f"{slug}.jpg"), Inches(x + 0.12), Inches(y + 0.12),
            width=Inches(1.375), height=Inches(1.50),
        )
        photo.element.find(".//" + qn("p:cNvPr")).set("descr", f"Photograph of {speaker}")
        text(slide, x + 1.65, y + 0.12, 4.13, 0.29, speaker, 20, "F2A383", bold=True)
        text(slide, x + 1.65, y + 0.44, 4.13, 0.25, role, 12.5, "E8E5DE")
        quote_box = text(slide, x + 1.65, y + 0.78, 4.13, 1.16, f"“{quote}”", 16.5, "FFFFFF",
                         font="Cambria")
        quote_box.text_frame.vertical_anchor = MSO_ANCHOR.TOP
        notes.extend([
            f"{speaker} quotation source: {source}",
            f"{speaker} photograph: {credit}, cropped. U.S. public domain. {photo_source}",
        ])

    panel(slide, 0.65, 5.65, 12.02, 1.04)
    text(slide, 0.9, 5.83, 2.3, 0.33, "Ethan Mollick", 20, "F2A383", bold=True)
    text(slide, 0.9, 6.19, 2.3, 0.28, "Wharton professor", 13, "E8E5DE")
    text(slide, 3.15, 5.72, 9.14, 0.84, f"“{MOLLICK_QUOTE}”", 19, "FFFFFF", font="Cambria")
    notes.append(f"Ethan Mollick quotation source: {MOLLICK_SOURCE}")
    notes.append("Jensen Huang spoke at CES 2025. Keynote video: https://www.ces.tech/videos/nvidia-keynote/")
    slide.notes_slide.notes_text_frame.text = "\n".join(notes)
