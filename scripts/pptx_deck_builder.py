#!/usr/bin/env python3
"""
worship-deck-builder core engine.

Rebuilds a church Sunday-service PPTX by:
  - cloning slide shapes (backgrounds, fonts, logo, positions) from a style
    template PPTX instead of rebuilding formatting by hand,
  - pulling hymn text out of a hymnal PDF and chunking it into slides,
  - pulling contemporary (non-hymnal) song text out of a local song-library
    JSON file that YOU maintain -- this tool never fetches lyrics from the
    open web.

See ../SKILL.md for the workflow this implements and ../README.md for
install/usage instructions. Run with --help for CLI flags.
"""
import argparse
import copy
import io
import json
import re
import sys
from pathlib import Path

from pptx import Presentation
from pptx.oxml.ns import qn
from pypdf import PdfReader

VERSE_RE = re.compile(r"^\s*(\d+)\s+(.*\S)\s*$")
CHORUS_RE = re.compile(r"^\s*coro[.:-]*\s*(.*)$", re.IGNORECASE)
CREDIT_RE = re.compile(r"^\s*-\s*\S")  # trailing "-Tr. X." / "-Ejemplo." credit line
PAGE_NUM_RE = re.compile(r"^\s*\d+\s*$")


# --------------------------------------------------------------------------
# Slide cloning: the mechanism that keeps output visually identical to the
# style template instead of re-deriving fonts/colors/positions by hand.
# --------------------------------------------------------------------------

def clone_slide(source_slide, dest_prs, layout_name="Blank"):
    layout = next((l for l in dest_prs.slide_layouts if l.name == layout_name), None)
    if layout is None:
        layout = dest_prs.slide_layouts[6]

    new_slide = dest_prs.slides.add_slide(layout)
    for shp in list(new_slide.shapes):
        shp._element.getparent().remove(shp._element)

    src_spTree = source_slide.shapes._spTree
    dst_spTree = new_slide.shapes._spTree

    rid_map = {}
    for rel_id, rel in source_slide.part.rels.items():
        if "image" in rel.reltype:
            image_part = rel.target_part
            _, new_rid = new_slide.part.get_or_add_image_part(io.BytesIO(image_part.blob))
            rid_map[rel_id] = new_rid

    skip_tags = {qn("p:nvGrpSpPr"), qn("p:grpSpPr")}
    for child in list(src_spTree):
        if child.tag in skip_tags:
            continue
        new_child = copy.deepcopy(child)
        for blip in new_child.iter(qn("a:blip")):
            old_rid = blip.get(qn("r:embed"))
            if old_rid and old_rid in rid_map:
                blip.set(qn("r:embed"), rid_map[old_rid])
        dst_spTree.append(new_child)

    return new_slide


def delete_all_slides(prs):
    xml_slides = prs.slides._sldIdLst
    for sld in list(xml_slides):
        prs.part.drop_rel(sld.get(qn("r:id")))
        xml_slides.remove(sld)


def get_textboxes(slide):
    return [s for s in slide.shapes if s.has_text_frame and s.text_frame.text.strip()]


def set_textbox_lines(shape, lines):
    """Replace a text box's paragraph text line-by-line, keeping each
    paragraph's existing run formatting (font/size/color/bold)."""
    tf = shape.text_frame
    txBody = tf._txBody
    while len(tf.paragraphs) < len(lines):
        txBody.append(copy.deepcopy(tf.paragraphs[-1]._p))
    while len(tf.paragraphs) > len(lines):
        p_el = tf.paragraphs[-1]._p
        p_el.getparent().remove(p_el)

    for p, line in zip(tf.paragraphs, lines):
        runs = p.runs
        if not runs:
            continue
        runs[0].text = line
        for extra in runs[1:]:
            extra._r.getparent().remove(extra._r)


def chunk_lines(lines, chunk_size):
    if chunk_size == "whole" or chunk_size is None:
        return [lines]
    return [lines[i:i + chunk_size] for i in range(0, len(lines), chunk_size)]


# --------------------------------------------------------------------------
# Hymnal PDF parsing
# --------------------------------------------------------------------------

def hymn_page_text(pdf_path, page_number):
    reader = PdfReader(pdf_path)
    return reader.pages[page_number - 1].extract_text()


def parse_hymn(raw_text):
    """Split a hymnal page's raw text into {"verses": [[line, ...], ...],
    "chorus": [line, ...]}. Tuned to a "Coro.-" / numbered-verse format
    (e.g. classic Spanish-language hymnals) -- proofread the result, this is
    a heuristic, not a guarantee, especially on irregularly-formatted hymns.
    """
    lines = [l.strip() for l in raw_text.splitlines()]
    lines = [l for l in lines if l and not PAGE_NUM_RE.match(l)]

    verses, chorus = [], []
    current, in_chorus = None, False
    started = False

    for line in lines:
        if not started:
            # skip title/subtitle/time-signature header lines until verse 1
            m = VERSE_RE.match(line)
            if m and m.group(1) == "1":
                started = True
            else:
                continue

        m = VERSE_RE.match(line)
        cm = CHORUS_RE.match(line)
        if m:
            current = []
            verses.append(current)
            in_chorus = False
            current.append(m.group(2))
        elif cm:
            in_chorus = True
            current = chorus
            if cm.group(1):
                current.append(cm.group(1))
        elif CREDIT_RE.match(line):
            continue
        else:
            if current is not None:
                current.append(line)

    return {"verses": verses, "chorus": chorus}


def hymn_sections(parsed, verse_chunk_size, chorus_chunk_size):
    """Yield (kind, lines) stanza-by-stanza in performance order: verse 1,
    chorus, verse 2, chorus, ... -- matching how these hymnals are actually
    sung, and chunk each stanza into slide-sized pieces."""
    sections = []
    for verse in parsed["verses"]:
        sections.append(("verse", verse, verse_chunk_size))
        if parsed["chorus"]:
            sections.append(("chorus", parsed["chorus"], chorus_chunk_size))
    return sections


# --------------------------------------------------------------------------
# Song library (contemporary / non-hymnal songs)
# --------------------------------------------------------------------------

def load_library(path):
    if not path:
        return {}
    return json.loads(Path(path).read_text(encoding="utf-8"))


def library_sections(entry):
    default_chunk = entry.get("chunk_size", 2)
    return [
        (sec.get("type", "verse"), sec["lines"], sec.get("chunk_size", default_chunk))
        for sec in entry["sections"]
    ]


# --------------------------------------------------------------------------
# Slide building
# --------------------------------------------------------------------------

def build_title_slide(target, title_template, title_lines, subtitle_lines):
    s = clone_slide(title_template, target)
    boxes = get_textboxes(s)
    if len(boxes) >= 2:
        set_textbox_lines(boxes[0], title_lines)
        set_textbox_lines(boxes[1], subtitle_lines)
    elif boxes:
        set_textbox_lines(boxes[0], title_lines + subtitle_lines)
    return s


def build_lyric_slides(target, lyric_template, sections):
    for _kind, lines, chunk_size in sections:
        for chunk in chunk_lines(lines, chunk_size):
            s = clone_slide(lyric_template, target)
            boxes = get_textboxes(s)
            if boxes:
                set_textbox_lines(boxes[0], chunk)


def detect_template_slides(prs, title_index=None, lyric_index=None):
    """Auto-detect a hymn-title slide (2 text boxes, second matches
    'Himno N') and a lyric slide (1 text box, 2+ lines) if not given
    explicitly. Explicit --title-slide-index/--lyric-slide-index (1-based)
    are more reliable -- use them when auto-detection picks the wrong
    slide."""
    slides = list(prs.slides)
    title_slide = slides[title_index - 1] if title_index else None
    lyric_slide = slides[lyric_index - 1] if lyric_index else None

    for s in slides:
        boxes = get_textboxes(s)
        if title_slide is None and len(boxes) == 2:
            second_text = boxes[1].text_frame.text.strip()
            if re.match(r"^himno\s+\d+$", second_text, re.IGNORECASE):
                title_slide = s
        if lyric_slide is None and len(boxes) == 1:
            if len(boxes[0].text_frame.paragraphs) >= 2:
                lyric_slide = s

    if title_slide is None or lyric_slide is None:
        raise SystemExit(
            "Could not auto-detect template slides. Pass --title-slide-index "
            "and --lyric-slide-index (1-based slide numbers in --template)."
        )
    return title_slide, lyric_slide


# --------------------------------------------------------------------------
# Build driver
# --------------------------------------------------------------------------

def build(spec_path):
    spec = json.loads(Path(spec_path).read_text(encoding="utf-8"))
    template = Presentation(spec["template"])
    title_slide, lyric_slide = detect_template_slides(
        template, spec.get("title_slide_index"), spec.get("lyric_slide_index")
    )

    target = Presentation(spec["template"])
    delete_all_slides(target)

    library = load_library(spec.get("library"))
    source_cache = {spec["template"]: template}

    def get_source(path):
        if path not in source_cache:
            source_cache[path] = Presentation(path)
        return source_cache[path]

    for item in spec["items"]:
        op = item["op"]
        if op == "clone_range":
            source = get_source(item.get("source", spec["template"]))
            src_slides = list(source.slides)
            for idx in range(item["start"] - 1, item["end"]):
                clone_slide(src_slides[idx], target)

        elif op == "hymn":
            page_no = item["himno"]
            raw = hymn_page_text(spec["hymnal"], page_no)
            parsed = parse_hymn(raw)
            sections = hymn_sections(
                parsed,
                item.get("verse_chunk_size", 2),
                item.get("chorus_chunk_size", 2),
            )
            build_title_slide(target, title_slide, [item["title"]], [f"Himno {page_no}"])
            build_lyric_slides(target, lyric_slide, sections)

        elif op == "library_song":
            entry = library[item["key"]]
            sections = library_sections(entry)
            build_title_slide(target, title_slide, entry["title_lines"], entry.get("subtitle_lines", []))
            build_lyric_slides(target, lyric_slide, sections)

        else:
            raise SystemExit(f"Unknown item op: {op!r}")

    target.save(spec["output"])
    print(f"wrote {spec['output']} ({len(list(target.slides))} slides)")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)

    b = sub.add_parser("build", help="Build a deck from a JSON spec file")
    b.add_argument("--spec", required=True, help="Path to songs.json (see examples/songs.example.json)")

    h = sub.add_parser("hymn", help="Look up and print one hymn's parsed text (debug helper)")
    h.add_argument("--hymnal", required=True)
    h.add_argument("--page", type=int, required=True, help="Page number (often == hymn number)")

    args = parser.parse_args()

    if args.command == "build":
        build(args.spec)
    elif args.command == "hymn":
        raw = hymn_page_text(args.hymnal, args.page)
        parsed = parse_hymn(raw)
        print(json.dumps(parsed, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    sys.exit(main())
