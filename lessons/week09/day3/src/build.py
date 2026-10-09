"""Build Week 9 Day 3 materials: student packet PDF, teacher copy PDF, slides PPTX.

Usage:  python3 build.py            (from this folder; outputs land one level up)
Needs:  weasyprint, python-pptx, Pillow, poppler-utils (pdftoppm), TeX Gyre Adventor font.
"""
import json
import re
import subprocess
import sys
from pathlib import Path

from PIL import Image, ImageDraw
from weasyprint import HTML

SRC = Path(__file__).resolve().parent
OUT = SRC.parent
BUILD = SRC / "build"
PACKET_PDF = OUT / "Week9_Day3_Student_Packet.pdf"
TEACHER_PDF = OUT / "Week9_Day3_Teacher_Copy.pdf"
PX_PER_IN = 96
THUMB_DPI = 150


def walk(box):
    yield box
    for child in getattr(box, "children", None) or []:
        yield from walk(child)


def build_packet():
    doc = HTML(SRC / "student_packet.html").render()
    if len(doc.pages) != 5:
        sys.exit(f"Student packet must be 5 pages; got {len(doc.pages)}")
    regions = {}
    for page_no, page in enumerate(doc.pages, 1):
        for box in walk(page._page_box):
            el = getattr(box, "element", None)
            if el is None or not el.get("data-hl") or not hasattr(box, "border_width"):
                continue
            x, y = box.border_box_x(), box.border_box_y()
            w, h = box.border_width(), box.border_height()
            key = el.get("data-hl")
            if key in regions:  # union of split boxes
                px, py, pw, ph, _ = regions[key]
                x2, y2 = max(px + pw, x + w), max(py + ph, y + h)
                x, y = min(px, x), min(py, y)
                w, h = x2 - x, y2 - y
            regions[key] = (x, y, w, h, page_no)
    doc.write_pdf(PACKET_PDF)
    BUILD.mkdir(exist_ok=True)
    subprocess.run(["pdftoppm", "-r", str(THUMB_DPI), "-png", str(PACKET_PDF), str(BUILD / "packet")], check=True)
    (BUILD / "regions.json").write_text(json.dumps(regions, indent=1))
    return regions


def page_png(page_no):
    return BUILD / f"packet-{page_no}.png"


def make_thumb(key, regions):
    """Packet page with everything but the active box faded and the box outlined in red."""
    x, y, w, h, page_no = regions[key]
    img = Image.open(page_png(page_no)).convert("RGB")
    s = THUMB_DPI / PX_PER_IN
    pad = 4 * s
    rect = [x * s - pad, y * s - pad, (x + w) * s + pad, (y + h) * s + pad]
    fade = Image.new("RGB", img.size, "white")
    mask = Image.new("L", img.size, 150)
    ImageDraw.Draw(mask).rounded_rectangle(rect, radius=14 * s, fill=0)
    img = Image.composite(fade, img, mask)
    ImageDraw.Draw(img).rounded_rectangle(rect, radius=14 * s, outline=(223, 68, 42), width=int(3 * s))
    path = BUILD / f"thumb-{key}.png"
    img.save(path)
    return path


def build_teacher(regions):
    html = (SRC / "teacher_copy.html").read_text()

    def thumb(m):
        key = m.group(1)
        if key.startswith("page"):
            path = page_png(int(key[4:]))
        else:
            path = make_thumb(key, regions)
        return f'<img class="thumbimg" src="{path.relative_to(SRC)}">'

    html = re.sub(r"\{\{thumb:([\w-]+)\}\}", thumb, html)
    HTML(string=html, base_url=str(SRC)).write_pdf(TEACHER_PDF)


if __name__ == "__main__":
    regions = build_packet()
    print("packet ok:", ", ".join(sorted(regions)))
    if (SRC / "teacher_copy.html").exists():
        build_teacher(regions)
        print("teacher copy ok")
    if (SRC / "build_slides.py").exists():
        subprocess.run([sys.executable, str(SRC / "build_slides.py")], check=True)
