#!/usr/bin/env python3
"""Lyrics sheet (PDF) for the singers, read from a Sunday deck a human already reviewed.

It reads the finished PPTX, not the build spec, so the sheet matches what is on
screen. It refuses to run without --reviewed: the operator has to check the deck
first (see ../SKILL.md, step 8).

One song per page; a song that doesn't fit one column at ONE_COL_MIN_PT goes in
two columns on its page. Songs keep the deck's order and repeats. Hymn slides are
regrouped into whole stanzas using the himnario; each contemporary-song slide is
one block. Lowercase continuation lines (split only to fit the screen) are joined
back onto the previous line, and the font never shrinks so far that a line wraps.
"""
import argparse
import re
import sys
import unicodedata
from pathlib import Path

from pptx import Presentation
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.colors import HexColor
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (BaseDocTemplate, Frame, PageTemplate, Paragraph,
                                Spacer, KeepTogether, PageBreak,
                                NextPageTemplate, FrameBreak)

SKILL = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(SKILL / "scripts"))
from pptx_deck_builder import hymn_page_text, parse_hymn, hymn_sections  # noqa: E402

HIMNO_RE = re.compile(r"^himno\s+(\d+)$", re.I)
RANGE_RE = re.compile(r"^\d+:\d+")
TITLE_PT = 110
ONE_COL_MIN_PT = 12


def max_pt(box):
    sizes = [r.font.size.pt for p in box.text_frame.paragraphs for r in p.runs if r.font.size]
    return max(sizes or [0])


def norm(text):
    text = unicodedata.normalize("NFD", text.lower())
    return re.sub(r"[^a-z0-9]", "", text)


def join_continuations(lines):
    out = []
    for line in (l.strip() for l in lines):
        if out and line[:1].islower():
            out[-1] = f"{out[-1]} {line}"
        else:
            out.append(line)
    return out


def read_deck(path):
    songs, current = [], None
    for slide in Presentation(path).slides:
        boxes = [s for s in slide.shapes if s.has_text_frame and s.text_frame.text.strip()]
        if not boxes:
            continue
        texts = [b.text_frame.text.strip() for b in boxes]
        if len(boxes) >= 3 or any(RANGE_RE.match(t) for t in texts):
            current = None  # purpose statement, sermon, scripture
            continue
        if any(max_pt(b) >= TITLE_PT for b in boxes):
            hymn = next((int(m.group(1)) for t in texts if (m := HIMNO_RE.match(t))), None)
            if hymn:
                title = next(t for t in texts if not HIMNO_RE.match(t))
            else:
                title = " ".join(b.text_frame.text.strip() for b in sorted(boxes, key=lambda b: b.top))
            current = {"title": title, "hymn": hymn, "slides": []}
            songs.append(current)
            continue
        if current is None:
            continue
        lines = [p.text.strip() for p in boxes[0].text_frame.paragraphs if p.text.strip()]
        if lines and not re.match(r"^\d+$", lines[0]):
            current["slides"].append(lines)
    return songs


def hymn_blocks(song, hymnal):
    expected = [lines for _, lines, _ in hymn_sections(parse_hymn(hymn_page_text(hymnal, song["hymn"])), 99, 99)]
    blocks, acc, i = [], [], 0
    for lines in song["slides"]:
        acc += lines
        if i < len(expected) and len(norm(" ".join(acc))) >= len(norm(" ".join(expected[i]))):
            blocks.append(acc)
            acc, i = [], i + 1
    if acc:
        blocks.append(acc)
    return blocks


FONT_DIR = Path("/System/Library/Fonts/Supplemental")


def register_fonts():
    """Georgia (macOS) when present, otherwise reportlab's built-in Times."""
    if not (FONT_DIR / "Georgia.ttf").exists():
        return "Times-Roman", "Times-Bold"
    for name in ("Georgia", "Georgia Bold"):
        pdfmetrics.registerFont(TTFont(name, str(FONT_DIR / f"{name}.ttf")))
    return "Georgia", "Georgia Bold"


def render(songs, out, heading, hymnal):
    regular, bold = register_fonts()
    ink, soft = HexColor("#1a1a1a"), HexColor("#6b6b6b")
    title = ParagraphStyle("title", fontName=bold, fontSize=18, leading=22, textColor=ink, spaceAfter=14)
    num = ParagraphStyle("num", fontName=regular, fontSize=11, leading=14, textColor=soft)

    margin, gutter, head_h = 0.75 * inch, 0.4 * inch, 50
    w, h = letter
    full_w, body_h = w - 2 * margin, h - 2 * margin
    col_w, col_h = (full_w - gutter) / 2, body_h - head_h
    pad = dict(leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)

    def footer(canvas, doc):
        canvas.setFont(regular, 9)
        canvas.setFillColor(soft)
        canvas.drawRightString(w - margin, 0.4 * inch, f"{heading} · {doc.page}")

    # "one": a single column. "two": the title across the top, then two columns.
    one = PageTemplate(id="one", onPage=footer, frames=[Frame(margin, margin, full_w, body_h, **pad)])
    two = PageTemplate(id="two", onPage=footer, frames=[
        Frame(margin, h - margin - head_h, full_w, head_h, **pad),
        Frame(margin, margin, col_w, col_h, **pad),
        Frame(margin + col_w + gutter, margin, col_w, col_h, **pad)])
    doc = BaseDocTemplate(str(out), pagesize=letter, title=f"Letras · {heading}",
                          leftMargin=margin, rightMargin=margin, topMargin=margin, bottomMargin=margin)
    doc.addPageTemplates([one, two])

    def header(song):
        items = [Paragraph(f"Himno {song['hymn']}", num)] if song["hymn"] else []
        return items + [Paragraph(song["title"], title)]

    def block_parts(blocks, size):
        lyric = ParagraphStyle("lyric", fontName=regular, fontSize=size, leading=size * 1.3, textColor=ink)
        return [[Paragraph("<br/>".join(join_continuations(b)), lyric), Spacer(1, size * 0.85)] for b in blocks]

    def height(flowables, width):
        return sum(f.wrap(width, body_h)[1] + f.getSpaceAfter() for f in flowables)

    def no_wrap(blocks, size, width):
        # a wrapped line would read as two sung lines
        lines = [l for b in blocks for l in join_continuations(b)]
        return max(pdfmetrics.stringWidth(l, regular, size) for l in lines) <= width

    def fits_one(song, blocks, size):
        return no_wrap(blocks, size, full_w) and height(header(song) + [f for b in block_parts(blocks, size) for f in b], full_w) <= body_h

    def fits_two(blocks, size):
        # whole blocks only, filled top to bottom: left column first, then right
        if not no_wrap(blocks, size, col_w):
            return False
        cols, used = 1, 0
        for part in block_parts(blocks, size):
            bh = height(part, col_w)
            if used + bh > col_h:
                cols, used = cols + 1, 0
            used += bh
        return cols <= 2

    sizes = [s / 2 for s in range(28, 19, -1)]  # 14pt down to 10pt
    story = []
    for n, song in enumerate(songs):
        blocks = hymn_blocks(song, hymnal) if song["hymn"] else song["slides"]
        # one column if it fits at 12pt or more; otherwise two columns on the same page
        size = next((s for s in sizes if s >= ONE_COL_MIN_PT and fits_one(song, blocks, s)), None)
        layout = "one" if size else "two"
        if not size:
            size = next((s for s in sizes if fits_two(blocks, s)), sizes[-1])
        story.append(NextPageTemplate(layout))
        if n:
            story.append(PageBreak())
        story += header(song)
        if layout == "two":
            story.append(FrameBreak())
        story += [KeepTogether(part) for part in block_parts(blocks, size)]
    doc.build(story)


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[1])
    parser.add_argument("--deck", required=True, help="the reviewed Sunday PPTX")
    parser.add_argument("--output", help='default: "Letras - <deck name>.pdf" next to the deck')
    parser.add_argument("--heading", help='footer text, default: deck name ("OCTUBRE 3" -> "Octubre 3")')
    parser.add_argument("--hymnal", default=str(SKILL / "reference/hymnal.pdf"))
    parser.add_argument("--reviewed", action="store_true",
                        help="confirm a human already reviewed this deck (required)")
    args = parser.parse_args()

    if not args.reviewed:
        sys.exit("The lyrics PDF is made only from a deck a human already reviewed. "
                 "Ask the operator to check the deck, then rerun with --reviewed.")
    deck = Path(args.deck)
    out = Path(args.output) if args.output else deck.with_name(f"Letras - {deck.stem}.pdf")
    heading = args.heading or deck.stem.capitalize()

    songs = read_deck(deck)
    if not songs:
        sys.exit(f"No songs found in {deck}")
    render(songs, out, heading, args.hymnal)
    print(out)


if __name__ == "__main__":
    main()
