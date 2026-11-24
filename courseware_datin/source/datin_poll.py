"""Audience voting and presenter results for the Datin opening poll."""

from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

CODE = "230174"
URL = f"https://alfredang.github.io/livepoll/?join={CODE}"
RESULTS_URL = f"https://alfredang.github.io/livepoll/?results={CODE}"
RESULTS_BUTTON = "Open the Poll Result"
RESULTS_INSTRUCTIONS = "Scan to vote. Open results to see the live tally."
QUESTION = "How Much of Your Team's Week Could an AI Agent Take On Today?"
INSTRUCTIONS = "Vote using the QR code on the next slide, or raise 1, 2, 3 or 4 fingers."


def text(slide, x, y, width, height, value, size, color, *, font="Calibri", bold=False, center=False):
    shape = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(width), Inches(height))
    write(shape, value, size, color, font=font, bold=bold, center=center)
    return shape


def write(shape, value, size, color, *, font="Calibri", bold=False, center=False):
    frame = shape.text_frame
    frame.clear()
    frame.word_wrap = True
    frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    frame.margin_left = frame.margin_right = Inches(0.05)
    frame.margin_top = frame.margin_bottom = Inches(0.02)
    paragraph = frame.paragraphs[0]
    paragraph.alignment = PP_ALIGN.CENTER if center else PP_ALIGN.LEFT
    paragraph.line_spacing = 1.05
    run = paragraph.add_run()
    run.text = value
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = RGBColor.from_string(color)


def link_results_button(button):
    button.fill.solid()
    button.fill.fore_color.rgb = RGBColor.from_string("E8E5DE")
    button.click_action.hyperlink.address = RESULTS_URL
    for paragraph in button.text_frame.paragraphs:
        for run in paragraph.runs:
            run.hyperlink.address = RESULTS_URL
            run.font.underline = False


def build_join_slide(slide, number, title, course_code, *, image_path=None, existing_picture=None):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = RGBColor.from_string("21201C")
    text(slide, 0.65, 0.46, 11.8, 0.36, "TOPIC 1 · LIVE POLL", 14, "D97757", bold=True)
    text(slide, 0.65, 1.26, 7.0, 1.90, QUESTION, 32, "FFFFFF", font="Cambria", bold=True)
    text(slide, 0.69, 3.35, 6.7, 0.55, RESULTS_INSTRUCTIONS, 20, "E8E5DE")
    button = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.71), Inches(4.18), Inches(3.05), Inches(0.71))
    button.line.fill.background()
    write(button, RESULTS_BUTTON, 20, "21201C", bold=True, center=True)
    link_results_button(button)
    text(slide, 0.69, 5.04, 6.7, 0.45, f"If prompted, enter {CODE}.", 17, "C9C4B9")
    text(slide, 8.16, 1.15, 4.4, 0.4, "SCAN TO VOTE", 16, "FFFFFF", bold=True, center=True)
    x, y, width = Inches(8.16), Inches(1.84), Inches(4.4)
    if existing_picture is None:
        picture = slide.shapes.add_picture(str(image_path), x, y, width=width)
    else:
        picture = existing_picture
        picture.height = round(picture.height * width / picture.width)
        picture.width, picture.left, picture.top = width, x, y
    picture.click_action.hyperlink.address = URL
    picture.element.find(".//" + qn("p:cNvPr")).set("descr", f"Scan to join the Datin live poll. Join code {CODE}.")
    text(slide, 0.5, 7.08, 6.6, 0.30, "© Tertiary Infotech Academy Pte Ltd", 9, "B9B4A9")
    text(slide, 7.1, 7.08, 5.2, 0.30, f"{title}  ·  {course_code}", 9, "B9B4A9", center=True)
    text(slide, 12.4, 7.08, 0.55, 0.30, str(number), 9, "B9B4A9", center=True)
