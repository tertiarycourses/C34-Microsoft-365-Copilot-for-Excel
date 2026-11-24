#!/usr/bin/env bash
set -euo pipefail
source_dir="$(cd "$(dirname "$0")" && pwd)"
edition="$(dirname "$source_dir")"
engine="$(dirname "$edition")/.agents/skills/non-wsq-courseware-build"
stage="$(mktemp -d)"
trap 'rm -rf "$stage"' EXIT
mkdir -p "$stage/build" "$stage/screenshots"
ln -s "$engine/assets" "$stage/assets"
for script in build_lesson_plan.py build_learner_guide.py prodoc.py inject_toc.py lab_prompts.py; do
  cp "$engine/build/$script" "$stage/build/"
done
cp "$source_dir/"*.py "$stage/build/"
export COURSE_REPO="$edition"
mkdir -p "$edition/courseware" "$edition/labs" "$edition/prompts"
reference_repo="${DATIN_REFERENCE_REPO:-$(dirname "$edition")/../../Tertiary WSQ Projects/TGS-2023037472-Business Innovation with Agentic AI and AI Agents/courseware}"
reference_commit=bda24288138f56a1e93c3185754f8c1926e407d3
reference_deck="courseware/Business Innovation with Agentic AI and AI Agents-v57.pptx"
reference_kit=".claude/skills/courseware-build"
mkdir -p "$stage/reference"
git -C "$reference_repo" archive "$reference_commit" "$reference_deck" "$reference_kit/assets/tertiary-infotech-logo.png" "$reference_kit/build/ccdeck.py" "$reference_kit/build/icons" "$reference_kit/build/render_icons.js" | tar -x -C "$stage/reference"
test "$(git hash-object "$stage/reference/$reference_deck")" = 1703a1c0a3865185565e2ef85c93e91f7ade904b
cp "$stage/reference/$reference_kit/build/ccdeck.py" "$stage/reference/$reference_kit/build/render_icons.js" "$stage/build/"
cp -R "$stage/reference/$reference_kit/build/icons" "$stage/build/icons"
python3 "$source_dir/build_labs.py"
soffice_bin="${SOFFICE:-soffice}"
profile="$(python3 -c 'import pathlib,sys; print(pathlib.Path(sys.argv[1]).as_uri())' "$stage/lo")"
example="$edition/labs/lab-02-interactive-sales-dashboard/examples/Datin_Sales_Dashboard_Example_GHS.xlsx"
"$soffice_bin" "-env:UserInstallation=$profile" --headless --convert-to pdf --outdir "$stage" "$example"
python3 - "$stage/Datin_Sales_Dashboard_Example_GHS.pdf" "$stage/screenshots/dashboard.png" <<'PY'
import sys
from pathlib import Path
import fitz
from PIL import Image, ImageChops

with fitz.open(sys.argv[1]) as pdf:
    page = next(page for page in pdf if "Sales performance" in page.get_text())
    pix = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0), alpha=False)
    image = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    white = Image.new("RGB", image.size, "white")
    border = ImageChops.difference(image, white).getbbox()
    if border is None:
        raise ValueError("Dashboard rendered as a blank page")
    image = image.crop(border)
    image.save(sys.argv[2])
    width, height = image.size
    folder = Path(sys.argv[2]).parent
    # Zoom each worksheet chart row. The chart grid is fixed in workbooks.py.
    image.crop((0, round(height * 0.315), width, round(height * 0.65))).save(folder / "dashboard-top.png")
    image.crop((0, round(height * 0.658), width, round(height * 0.955))).save(folder / "dashboard-bottom.png")
PY
title="$(cd "$stage/build" && python3 -c 'from course_data import SHORT_TITLE; print(SHORT_TITLE)')"
version="$(cd "$stage/build" && python3 -c 'from course_data import VERSION; print(VERSION)')"
out="$edition/courseware"
python3 "$stage/build/build_datin_slides.py" "$stage/reference/$reference_deck" "$stage/reference/$reference_kit/assets/tertiary-infotech-logo.png" "$source_dir/assets" "$stage/screenshots/dashboard.png" "$out/$title-$version.pptx"
python3 - "$stage/build/build_lesson_plan.py" <<'PY'
from pathlib import Path
import sys
path = Path(sys.argv[1])
text = path.read_text()
target = '("Training Provider",C.ORG'
assert text.count(target) == 1
path.write_text(text.replace(target, '("Participants",C.PARTICIPANTS),\n      ' + target))
PY
python3 "$stage/build/build_lesson_plan.py"
python3 "$stage/build/build_learner_guide.py"

render() {
  "$soffice_bin" "-env:UserInstallation=$profile" --headless --convert-to pdf --outdir "$out" "$@"
}
render "$out/$title-$version.pptx" "$out/LG-$title.docx" "$out/LP-$title.docx"
for kind in LG LP; do
  python3 "$stage/build/inject_toc.py" "$out/$kind-$title.docx" "$out/$kind-$title.pdf" 2
done
render "$out/LG-$title.docx" "$out/LP-$title.docx"
