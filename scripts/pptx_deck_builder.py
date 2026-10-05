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
import datetime
import io
import json
import re
import sys
import unicodedata
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn
from pypdf import PdfReader

SKILL_DIR = Path(__file__).resolve().parent.parent
DEFAULT_TEMPLATE = SKILL_DIR / "reference" / "style-template.pptx"
DEFAULT_HYMNAL = SKILL_DIR / "reference" / "hymnal.pdf"
DEFAULT_LIBRARY = SKILL_DIR / "reference" / "song-library.json"

# A verse line starts with a 1-2 digit number (a verse number), optionally
# followed by "." or "-", then the verse text. The text may be glued straight
# onto the number ("3¡Que...") so the separator/space is optional. (?!\d) keeps
# longer numbers (tempo marks like "444", "1000 veces") from being read as
# verse 10/44; pure-number lines are already dropped by PAGE_NUM_RE. The text
# may start lowercase (hymnals have the odd typo, e.g. "4 la fe..."), so we
# don't constrain its first character.
VERSE_RE = re.compile(r"^\s*(\d{1,2})(?!\d)[.\-]?\s*(\S.*\S|\S)\s*$")
CHORUS_RE = re.compile(r"^\s*coro[.,:-]*\s*(.*)$", re.IGNORECASE)  # some pages print "Coro,-"
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

    _copy_slide_background(source_slide, new_slide)
    return new_slide


def _copy_slide_background(source_slide, new_slide):
    """Copy the slide-level <p:bg> element. clone_slide only copies the shape
    tree; <p:bg> lives directly under <p:cSld>, so without this the cloned
    slide loses its background fill. Many real decks put a solid dark <p:bg>
    behind a semi-transparent photo -- drop it and the slide renders
    washed-out. <p:bg> must be the first child of <p:cSld>."""
    src_cSld = source_slide._element.find(qn("p:cSld"))
    dst_cSld = new_slide._element.find(qn("p:cSld"))
    src_bg = src_cSld.find(qn("p:bg"))
    if src_bg is None:
        return
    old_bg = dst_cSld.find(qn("p:bg"))
    if old_bg is not None:
        dst_cSld.remove(old_bg)
    dst_cSld.insert(0, copy.deepcopy(src_bg))


def delete_all_slides(prs):
    xml_slides = prs.slides._sldIdLst
    for sld in list(xml_slides):
        prs.part.drop_rel(sld.get(qn("r:id")))
        xml_slides.remove(sld)


def get_textboxes(slide):
    return [s for s in slide.shapes if s.has_text_frame and s.text_frame.text.strip()]


def first_font_size(shape):
    """The first explicit run font size in a text box, as an int (0 when the
    box inherits its size). Used to tell a slide's big title box from the
    smaller ones around it."""
    for p in shape.text_frame.paragraphs:
        for r in p.runs:
            if r.font.size is not None:
                return int(r.font.size)
    return 0


def set_textbox_lines(shape, lines):
    """Replace a text box's paragraph text line-by-line, reusing the box's run
    formatting (font/size/color/bold).

    Every paragraph is rebuilt from a "donor" paragraph that actually has a run,
    so a line never lands on a run-less template paragraph and gets silently
    dropped. (The old implementation skipped run-less paragraphs -- e.g. the
    blank separator paragraph in a scripture text box -- which quietly ate a
    line of every verse that aligned with it.)"""
    if not lines:
        lines = [""]
    tf = shape.text_frame
    txBody = tf._txBody
    donor = next((p._p for p in tf.paragraphs if p.runs), None)
    if donor is None:
        # No run anywhere to clone formatting from: fall back to plain text.
        while len(tf.paragraphs) < len(lines):
            txBody.append(copy.deepcopy(tf.paragraphs[-1]._p))
        while len(tf.paragraphs) > len(lines):
            tf.paragraphs[-1]._p.getparent().remove(tf.paragraphs[-1]._p)
        for p, line in zip(tf.paragraphs, lines):
            if p.runs:
                p.runs[0].text = line
        return

    for p in list(tf.paragraphs):
        p._p.getparent().remove(p._p)
    for _ in lines:
        txBody.append(copy.deepcopy(donor))
    for p, line in zip(tf.paragraphs, lines):
        p.runs[0].text = line
        for extra in p.runs[1:]:
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
    def size(lines, requested):
        return auto_chunk_size(lines) if requested == "auto" else requested

    sections = []
    for verse in parsed["verses"]:
        sections.append(("verse", verse, size(verse, verse_chunk_size)))
        if parsed["chorus"]:
            sections.append(("chorus", parsed["chorus"], size(parsed["chorus"], chorus_chunk_size)))
    return sections


def hymn_title(raw_text):
    """Best-effort hymn title from a hymnal page: the ALL-CAPS header between
    the page number and verse 1. A long header wraps onto several lines
    ("BIENAVENTURADOS LOS DE LIMPIO" / "CORAZON"), so consecutive caps lines
    are joined. Returned in sentence case, the way the operator writes it
    ("Bienaventurados los de limpio corazon"). Some headers lack accents, so
    verify/override via the spec's "title". Returns None if no header is found."""
    parts = []
    for line in raw_text.splitlines():
        s = line.strip()
        if not s or PAGE_NUM_RE.match(s):
            continue
        if VERSE_RE.match(s) or CHORUS_RE.match(s):
            break
        letters = [c for c in s if c.isalpha()]
        if len(letters) >= 3 and all(c.isupper() for c in letters):
            parts.append(s)
        elif parts:
            break  # first non-caps line after the header (tempo/key line)
    if not parts:
        return None
    words = [PROPER_WORDS.get(w, w) for w in " ".join(parts).lower().split()]
    return capitalize_first(" ".join(words))


# Names kept capitalized when a caps header is turned into sentence case.
PROPER_WORDS = {w.lower(): w for w in [
    "Dios", "Cristo", "Jesús", "Jesus", "Jesucristo", "Señor", "Jehová", "Espíritu",
    "Emmanuel", "Sion", "Sión",
]}


# --------------------------------------------------------------------------
# Lyric text conventions (how the church's operator formats slides by hand)
# --------------------------------------------------------------------------

DEFAULT_FORMAT = {"capitalize_lines": True, "break_at": 28, "collapse_repeats": True}


def capitalize_first(line):
    for i, ch in enumerate(line):
        if ch.isalpha():
            return line[:i] + ch.upper() + line[i + 1:]
    return line


# Words a long line is split *before* when it has no comma near the middle
# ("Mi alma te anhela / y tiene sed", "Nada me falta / pues todo provees").
BREAK_BEFORE = {"y", "e", "o", "que", "pues", "de", "en", "a", "por", "para", "con", "sin"}
MIN_PIECE = 8  # never leave a stub like "Oh," or "a ti" on its own line


def break_long_line(line, max_len):
    """Split an over-long line at a natural phrase break near the middle, the
    way the operator does by hand: after a comma ("Renuevame, / Senor Jesus")
    or before a connector word ("Mi alma te anhela / y tiene sed"). Both
    pieces must be at least MIN_PIECE long. Repeat markers (//..//) are left
    untouched."""
    if not max_len or len(line) <= max_len:
        return [line]
    mid = len(line) / 2
    cuts = []  # (score, index the tail starts at)
    for i, ch in enumerate(line):
        if ch == ",":
            cuts.append((abs(i + 1 - mid) - 4, i + 1))  # commas win a near-tie
    for m in re.finditer(r"\s(\S+)", line):
        if m.group(1).lower() in BREAK_BEFORE:
            cuts.append((abs(m.start() - mid), m.start()))
    for _score, cut in sorted(cuts):
        head, tail = line[:cut].strip(), line[cut:].strip()
        if len(head) >= MIN_PIECE and len(tail) >= MIN_PIECE:
            return break_long_line(head, max_len) + break_long_line(tail, max_len)
    return [line]


def collapse_doubled_line(line):
    """'Oh, tu fidelidad, oh, tu fidelidad' -> '//Oh, tu fidelidad//': a line
    that is the same phrase twice is written once with the repeat marker."""
    if "/" in line:
        return line
    body = line.rstrip(",.")
    for i, ch in enumerate(body):
        if ch == ",":
            head, tail = body[:i].strip(), body[i + 1:].strip()
            if head and head.lower() == tail.lower():
                return f"//{head}//"
    return line


def apply_format(lines, fmt):
    """Operator conventions for lyric lines. Only the first piece of a split
    line is capitalized; the continuation stays lowercase, as she types it
    ("Hermoso eres Tú, / amado mío"). fmt "verbatim" leaves lines untouched."""
    if fmt == "verbatim":
        return list(lines)
    fmt = {**DEFAULT_FORMAT, **(fmt or {})}
    out = []
    for line in lines:
        if fmt.get("collapse_repeats"):
            line = collapse_doubled_line(line)
        pieces = break_long_line(line, fmt.get("break_at")) if fmt.get("break_at") else [line]
        for n, p in enumerate(pieces):
            out.append(capitalize_first(p) if fmt.get("capitalize_lines") and n == 0 else p)
    return out


# A hymn whose lines are all short fits 4 lines per slide; longer lines need 2.
AUTO_CHUNK_SHORT_LINE = 28


def auto_chunk_size(lines):
    return 4 if lines and max(len(l) for l in lines) <= AUTO_CHUNK_SHORT_LINE else 2


# --------------------------------------------------------------------------
# Song library (contemporary / non-hymnal songs)
# --------------------------------------------------------------------------

DEFAULT_LIBRARY_DIR = SKILL_DIR / "reference" / "song-library"


def slugify(title):
    s = unicodedata.normalize("NFKD", title or "").encode("ascii", "ignore").decode()
    s = re.sub(r"[^a-z0-9]+", "-", s.lower()).strip("-")
    return s or "cancion"


def load_library(path=None, folder=None):
    """Load the contemporary-song library, merging two sources so the library
    can grow one file at a time: the legacy flat song-library.json (key ->
    entry) if present, plus one <slug>.json per song in the song-library/
    folder. Folder entries win on key collision."""
    lib = {}
    if path and Path(path).exists():
        flat = json.loads(Path(path).read_text(encoding="utf-8"))
        lib.update({k: v for k, v in flat.items() if not k.startswith("_")})
    folder = Path(folder) if folder else DEFAULT_LIBRARY_DIR
    if folder.exists():
        for f in sorted(folder.glob("*.json")):
            if f.name.startswith("_"):  # _example.json etc.
                continue
            entry = json.loads(f.read_text(encoding="utf-8"))
            lib[entry.get("key", f.stem)] = entry
    return lib


def library_sections(entry):
    default_chunk = entry.get("chunk_size", 2)
    return [
        (sec.get("type", "verse"), sec["lines"], sec.get("chunk_size", default_chunk))
        for sec in entry["sections"]
    ]


def entry_title_lines(entry):
    """Two-tone title (white main, cream secondary) for a library entry,
    supporting both the new title_white/title_cream fields and the older
    title_lines/subtitle_lines shape."""
    if "title_white" in entry or "title_cream" in entry:
        white = entry.get("title_white", "")
        cream = entry.get("title_cream", "")
        return ([white] if white else []), ([cream] if cream else [])
    return entry.get("title_lines", []), entry.get("subtitle_lines", [])


# --------------------------------------------------------------------------
# Slide building
# --------------------------------------------------------------------------

# The template's lyric box is narrower than the slide's frame and has a fixed
# height. Its spAutoFit only resizes the box when PowerPoint re-lays it out
# during an edit -- not on open -- so long lines wrapped and the text sat low or
# spilled off the slide until the operator dragged every box wider and
# re-centered it. These constants let the builder do that up front.
LINE_HEIGHT = 1.06  # rendered line height / font size (99pt lines are 1333698 EMU)
CHAR_WIDTH = 0.45   # average glyph width / font size for the lyric font


def frame_bounds(slide, fallback):
    """(left, top, width, height) of the slide's decorative frame (the group
    shape the text sits on), or of `fallback` when there is none."""
    group = next((sh for sh in slide.shapes if sh.shape_type == MSO_SHAPE_TYPE.GROUP), None)
    src = group or fallback
    return src.left, src.top, src.width, src.height


def estimate_lines(lines, width, size_emu):
    """How many lines `lines` wrap to in a box `width` wide (greedy word wrap
    with an average glyph width -- close enough to center the box)."""
    cap = max(1, int(width / (size_emu * CHAR_WIDTH)))
    total = 0
    for line in lines:
        count, cur = 1, 0
        for word in line.split():
            add = len(word) + (1 if cur else 0)
            if cur and cur + add > cap:
                count, cur = count + 1, len(word)
            else:
                cur += add
        total += count
    return total


def fit_lyric_box(slide, box, slide_height, lines):
    """Widen a lyric box to the frame and center it vertically for its text."""
    size_emu = first_font_size(box)  # python-pptx font sizes are in EMU
    if not size_emu:
        return
    left, top, width, _h = frame_bounds(slide, box)
    height = int(estimate_lines(lines, width, size_emu) * size_emu * LINE_HEIGHT)
    box.left, box.width, box.height = left, width, height
    box.top = max(top, (slide_height - height) // 2)


def build_title_slide(target, title_template, title_lines, subtitle_lines):
    s = clone_slide(title_template, target)
    boxes = get_textboxes(s)
    if len(boxes) >= 2:
        set_textbox_lines(boxes[0], title_lines)
        set_textbox_lines(boxes[1], subtitle_lines)
    elif boxes:
        set_textbox_lines(boxes[0], title_lines + subtitle_lines)
    return s


def build_lyric_slides(target, lyric_template, sections, fmt=None):
    for _kind, lines, chunk_size in sections:
        for chunk in chunk_lines(lines, chunk_size):
            s = clone_slide(lyric_template, target)
            boxes = get_textboxes(s)
            if boxes:
                text = apply_format(chunk, fmt)
                set_textbox_lines(boxes[0], text)
                fit_lyric_box(s, boxes[0], target.slide_height, text)


def build_two_tone_title(target, title_template, white_lines, cream_lines):
    """Title slide with a main phrase + a secondary part, filling the template's
    two boxes (e.g. big white main / small cream secondary). Short titles pass
    cream_lines=[] and land entirely in the first box."""
    s = clone_slide(title_template, target)
    boxes = get_textboxes(s)
    if len(boxes) >= 2:
        white, cream = boxes[0], boxes[1]
        set_textbox_lines(white, white_lines)
        set_textbox_lines(cream, cream_lines or [""])
        # Widen the small phrase to the frame so it stays on one line
        # ("Temprano yo te"); the text is centered, so this doesn't move it.
        left, _t, width, _h = frame_bounds(s, white)
        white.left, white.width = left, width
        if not cream_lines:
            # One-word title ("Fidelidad"): set it at the big size, where the
            # cream word would have been.
            big = first_font_size(cream)
            for p in white.text_frame.paragraphs:
                for r in p.runs:
                    r.font.size = big
            white.top = int(cream.top - big * 0.18)
    elif boxes:
        set_textbox_lines(boxes[0], white_lines + cream_lines)
    return s


def sermon_boxes(boxes):
    """Map a sermon-title slide's boxes to (lead, title, reference).

    Templates don't agree on shape order, so go by layout instead: stack the
    boxes top to bottom, take the biggest-font one as the title, and the boxes
    directly above and below it as the lead-in phrase and the scripture
    reference. Anything further down (a logo) is left alone. Returns None if
    the slide isn't shaped like that."""
    ordered = sorted(boxes, key=lambda b: b.top or 0)
    if len(ordered) < 3:
        return None
    sizes = [first_font_size(b) for b in ordered]
    i = max(range(len(ordered)), key=lambda n: sizes[n])
    if i == 0 or i == len(ordered) - 1:
        return None
    return ordered[i - 1], ordered[i], ordered[i + 1]


def build_sermon_title(target, sermon_template, lead, title, reference):
    """The prédica (sermon) title slide: a small lead-in phrase, the big main
    word under it, and the passage reference below."""
    s = clone_slide(sermon_template, target)
    parts = sermon_boxes(get_textboxes(s))
    if parts is None:
        return s
    lead_box, title_box, ref_box = parts
    set_textbox_lines(lead_box, [lead] if lead else [""])
    set_textbox_lines(title_box, [title])
    set_textbox_lines(ref_box, [reference] if reference else [""])
    return s


def build_scripture(target, ref_template, text_template, book, rng, chunks):
    """A scripture segment: one reference slide (book + range) followed by one
    slide per chunk of numbered verse text."""
    s = clone_slide(ref_template, target)
    boxes = get_textboxes(s)
    if len(boxes) >= 2:
        set_textbox_lines(boxes[0], [book])
        set_textbox_lines(boxes[1], [rng])
    elif boxes:
        set_textbox_lines(boxes[0], [f"{book} {rng}"])
    for chunk in chunks:
        s = clone_slide(text_template, target)
        boxes = get_textboxes(s)
        if boxes:
            # Scripture text is set verbatim: keep the Bible wording,
            # verse numbers and blank separators exactly as given.
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


HIMNO_RE = re.compile(r"^himno\s+\d+$", re.IGNORECASE)
RANGE_RE = re.compile(r"^\d+\s*:\s*\d")  # verse range like "42:1-2" / "00:0-0"


def detect_song_title_slide(prs, index=None):
    """A contemporary-song title template: a 2-box slide whose second box is a
    plain title fragment -- NOT a 'Himno N' number and NOT a scripture range."""
    slides = list(prs.slides)
    if index:
        return slides[index - 1]
    for s in slides:
        boxes = get_textboxes(s)
        if len(boxes) == 2:
            second = boxes[1].text_frame.text.strip()
            if not HIMNO_RE.match(second) and not RANGE_RE.match(second):
                return s
    return None


SERMON_REF_RE = re.compile(r"\d")


def detect_sermon_title_slide(prs, index=None):
    """The prédica (sermon) title template: 3+ text boxes stacked as lead-in
    phrase / BIG title / passage reference (a logo box may sit lower down).

    The reference box has to carry digits (a chapter or chapter:verse), which
    is what separates this slide from the 'declaración de propósito' slide,
    whose bottom box is plain prose. Pass sermon_title_slide_index when a
    template needs it picked explicitly."""
    slides = list(prs.slides)
    if index:
        return slides[index - 1]
    for s in slides:
        boxes = get_textboxes(s)
        if len(boxes) < 3:
            continue
        parts = sermon_boxes(boxes)
        if parts is None:
            continue
        if SERMON_REF_RE.search(parts[2].text_frame.text):
            return s
    return None


def detect_scripture_slides(prs, ref_index=None, text_index=None):
    """Scripture reference template (2 boxes: book + range like '00:0-0') and
    scripture text template (1 box whose first paragraph is a bare number)."""
    slides = list(prs.slides)
    ref = slides[ref_index - 1] if ref_index else None
    text = slides[text_index - 1] if text_index else None
    for s in slides:
        boxes = get_textboxes(s)
        if ref is None and len(boxes) == 2 and RANGE_RE.match(boxes[1].text_frame.text.strip()):
            ref = s
        if text is None and len(boxes) == 1:
            first = boxes[0].text_frame.paragraphs[0].text.strip()
            if re.match(r"^\d+$", first):
                text = s
    return ref, text


# --------------------------------------------------------------------------
# Build driver
# --------------------------------------------------------------------------

def resolve_default(spec_value, default_path, label):
    if spec_value:
        return spec_value
    if not default_path.exists():
        raise SystemExit(
            f"No {label!r} given in the spec, and the default "
            f"{default_path} doesn't exist. This skill is set up for a "
            f"single church's own files -- put your real {label} there, "
            f"or pass an explicit {label!r} path in the spec."
        )
    return str(default_path)


def build(spec_path):
    spec = json.loads(Path(spec_path).read_text(encoding="utf-8"))
    template_path = resolve_default(spec.get("template"), DEFAULT_TEMPLATE, "template")
    hymnal_path = resolve_default(spec.get("hymnal"), DEFAULT_HYMNAL, "hymnal")
    library_path = spec.get("library") or (str(DEFAULT_LIBRARY) if DEFAULT_LIBRARY.exists() else None)

    fmt = spec.get("format")
    template = Presentation(template_path)
    title_slide, lyric_slide = detect_template_slides(
        template, spec.get("title_slide_index"), spec.get("lyric_slide_index")
    )
    song_title_slide = detect_song_title_slide(template, spec.get("song_title_slide_index")) or title_slide
    scr_ref_slide, scr_text_slide = detect_scripture_slides(
        template, spec.get("scripture_ref_slide_index"), spec.get("scripture_text_slide_index")
    )
    sermon_slide = detect_sermon_title_slide(template, spec.get("sermon_title_slide_index"))

    target = Presentation(template_path)
    delete_all_slides(target)

    library = load_library(library_path, spec.get("library_folder"))
    source_cache = {template_path: template}

    def get_source(path):
        if path not in source_cache:
            source_cache[path] = Presentation(path)
        return source_cache[path]

    # Hymn lines are never re-broken or collapsed: the widened box wraps them,
    # and the hymnal's wording is kept as printed.
    hymn_fmt = {**DEFAULT_FORMAT, **(fmt or {}), "break_at": None, "collapse_repeats": False}

    def render_song(white, cream, sections, verbatim=False):
        build_two_tone_title(target, song_title_slide, white, cream)
        build_lyric_slides(target, lyric_slide, sections, "verbatim" if verbatim else fmt)

    for item in spec["items"]:
        op = item["op"]
        if op == "clone_range":
            source = get_source(item.get("source", template_path))
            src_slides = list(source.slides)
            for idx in range(item["start"] - 1, item["end"]):
                clone_slide(src_slides[idx], target)

        elif op == "hymn":
            page_no = item["himno"]
            raw = hymn_page_text(hymnal_path, page_no)
            parsed = parse_hymn(raw)
            sections = hymn_sections(
                parsed,
                item.get("verse_chunk_size", "auto"),
                item.get("chorus_chunk_size", "auto"),
            )
            title = item.get("title") or hymn_title(raw) or f"Himno {page_no}"
            build_title_slide(target, title_slide, [title], [f"Himno {page_no}"])
            build_lyric_slides(target, lyric_slide, sections, hymn_fmt)

        elif op == "library_song":  # backward-compatible alias of "song" by key
            entry = library[item["key"]]
            white, cream = entry_title_lines(entry)
            render_song(white, cream, library_sections(entry), entry.get("verbatim", False))

        elif op == "song":
            if "key" in item:
                entry = library[item["key"]]
                white, cream = entry_title_lines(entry)
                sections = library_sections(entry)
                verbatim = entry.get("verbatim", False)
            else:
                white = [item["title_white"]] if item.get("title_white") else []
                cream = [item["title_cream"]] if item.get("title_cream") else []
                sections = [
                    (s.get("type", "verse"), s["lines"], s.get("chunk_size", item.get("chunk_size", 2)))
                    for s in item["sections"]
                ]
                verbatim = item.get("verbatim", False)
            render_song(white, cream, sections, verbatim)

        elif op == "sermon":
            if sermon_slide is None:
                raise SystemExit(
                    "sermon op needs a sermon-title template (a slide with a "
                    "lead-in phrase, a big title and a passage reference). "
                    "Pass sermon_title_slide_index in the spec."
                )
            build_sermon_title(target, sermon_slide, item.get("lead", ""),
                               item["title"], item.get("reference", ""))

        elif op == "scripture":
            if scr_ref_slide is None or scr_text_slide is None:
                raise SystemExit(
                    "scripture op needs scripture templates. Pass "
                    "scripture_ref_slide_index / scripture_text_slide_index in the spec."
                )
            build_scripture(target, scr_ref_slide, scr_text_slide,
                            item["book"], item["range"], item["chunks"])

        else:
            raise SystemExit(f"Unknown item op: {op!r}")

    target.save(spec["output"])
    print(f"wrote {spec['output']} ({len(list(target.slides))} slides)")


# --------------------------------------------------------------------------
# Song library CLI: grow a folder of unique contemporary songs so the builder
# has a local source to consult instead of a Drive folder / the open web.
# --------------------------------------------------------------------------

def insert_into_deck(spec_path):
    """Insert scripture slides into an EXISTING deck (e.g. one the operator
    already edited by hand) without touching any of its slides. Each item is a
    "scripture" op plus "before": the 1-indexed number of the deck's ORIGINAL
    slide it goes in front of (omit it to append at the end). The scripture
    reference/text slides are cloned from the deck itself."""
    spec = json.loads(Path(spec_path).read_text(encoding="utf-8"))
    prs = Presentation(spec["deck"])
    original = list(prs.slides)
    ref_t, text_t = detect_scripture_slides(
        prs, spec.get("scripture_ref_slide_index"), spec.get("scripture_text_slide_index")
    )
    if ref_t is None or text_t is None:
        raise SystemExit("No scripture slides found in the deck. Pass "
                         "scripture_ref_slide_index / scripture_text_slide_index in the spec.")
    id_list = prs.slides._sldIdLst
    for item in spec["items"]:
        if item.get("op", "scripture") != "scripture":
            raise SystemExit(f"insert only supports scripture items, got {item.get('op')!r}")
        n_before = len(id_list)
        build_scripture(prs, ref_t, text_t, item["book"], item["range"], item["chunks"])
        if item.get("before"):
            anchor_id = original[item["before"] - 1].slide_id
            anchor = next(e for e in id_list if int(e.get("id")) == anchor_id)
            for e in list(id_list)[n_before:]:
                anchor.addprevious(e)
    prs.save(spec.get("output", spec["deck"]))
    print(f"wrote {spec.get('output', spec['deck'])} ({len(prs.slides)} slides, "
          f"{len(prs.slides) - len(original)} inserted)")


def library_dir(folder=None):
    d = Path(folder) if folder else DEFAULT_LIBRARY_DIR
    d.mkdir(parents=True, exist_ok=True)
    return d


def save_song(entry, folder=None, overwrite=False):
    entry = dict(entry)
    entry.setdefault("key", slugify(entry.get("title") or entry.get("title_white", "")))
    entry.setdefault("date_added", datetime.date.today().isoformat())
    dest = library_dir(folder) / f"{entry['key']}.json"
    if dest.exists() and not overwrite:
        return None
    dest.write_text(json.dumps(entry, ensure_ascii=False, indent=2), encoding="utf-8")
    return dest


def song_list(folder=None):
    lib = load_library(str(DEFAULT_LIBRARY) if DEFAULT_LIBRARY.exists() else None, folder)
    for key in sorted(lib):
        e = lib[key]
        title = e.get("title") or " ".join(filter(None, [e.get("title_white"), e.get("title_cream")])) \
            or " ".join(e.get("title_lines", []))
        print(f"{key:30} {title}  {('- ' + e['artist']) if e.get('artist') else ''}")


def songs_from_deck(deck_path):
    """Extract unique contemporary songs from an existing deck: each 2-box
    title slide (not a hymn number, not a scripture range) starts a song, and
    the following single-box lyric slides become its sections (one slide each).
    Lets you backfill the library from decks you already made by hand."""
    prs = Presentation(deck_path)
    songs, current = [], None
    for s in prs.slides:
        boxes = get_textboxes(s)
        if len(boxes) == 2:
            first, second = (b.text_frame.text.strip() for b in boxes)
            # the hymn number sits in either box ("Titulo / Himno 116" or "HIMNO 138 / titulo")
            if HIMNO_RE.match(first) or HIMNO_RE.match(second) or RANGE_RE.match(second):
                current = None  # hymn or scripture -> not a library song
                continue
            current = {"title_white": boxes[0].text_frame.text.strip(),
                       "title_cream": second, "source": f"deck:{Path(deck_path).name}",
                       "sections": []}
            current["title"] = " ".join(filter(None, [current["title_white"], current["title_cream"]]))
            songs.append(current)
        elif len(boxes) == 1 and current is not None:
            lines = [p.text for p in boxes[0].text_frame.paragraphs if p.text.strip()]
            if lines and not re.match(r"^\d+$", lines[0].strip()):  # skip scripture text
                current["sections"].append({"type": "verse", "lines": lines, "chunk_size": "whole"})
        else:
            current = None
    return [s for s in songs if s["sections"]]


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)

    b = sub.add_parser("build", help="Build a deck from a JSON spec file")
    b.add_argument("--spec", required=True, help="Path to songs.json (see examples/songs.example.json)")

    ins = sub.add_parser("insert", help="Insert scripture slides into an existing (hand-edited) deck")
    ins.add_argument("--spec", required=True, help='JSON: {"deck", "output", "items": [{"op": "scripture", "before": N, ...}]}')

    h = sub.add_parser("hymn", help="Look up and print one hymn's parsed text (debug helper)")
    h.add_argument("--hymnal", default=None, help=f"Defaults to {DEFAULT_HYMNAL}")
    h.add_argument("--page", type=int, required=True, help="Page number (often == hymn number)")

    sg = sub.add_parser("song", help="Manage the local song library folder")
    sgsub = sg.add_subparsers(dest="song_cmd", required=True)
    sa = sgsub.add_parser("add", help="Add/normalize a song from a JSON file into the library")
    sa.add_argument("--file", required=True, help="JSON entry (title, title_white/cream, artist, sections)")
    sa.add_argument("--overwrite", action="store_true")
    sgsub.add_parser("list", help="List the unique songs in the library")
    si = sgsub.add_parser("import-deck", help="Extract songs from an existing deck into the library")
    si.add_argument("--deck", required=True, help="Path to a .pptx deck to mine for songs")
    si.add_argument("--overwrite", action="store_true")

    args = parser.parse_args()

    if args.command == "build":
        build(args.spec)
    elif args.command == "insert":
        insert_into_deck(args.spec)
    elif args.command == "hymn":
        hymnal_path = resolve_default(args.hymnal, DEFAULT_HYMNAL, "hymnal")
        raw = hymn_page_text(hymnal_path, args.page)
        parsed = parse_hymn(raw)
        print(json.dumps(parsed, ensure_ascii=False, indent=2))
    elif args.command == "song":
        if args.song_cmd == "add":
            entry = json.loads(Path(args.file).read_text(encoding="utf-8"))
            dest = save_song(entry, overwrite=args.overwrite)
            print(f"saved {dest}" if dest else "skipped (already exists; pass --overwrite)")
        elif args.song_cmd == "list":
            song_list()
        elif args.song_cmd == "import-deck":
            added = 0
            for entry in songs_from_deck(args.deck):
                dest = save_song(entry, overwrite=args.overwrite)
                if dest:
                    print(f"+ {dest.stem}  ({len(entry['sections'])} slides)")
                    added += 1
            print(f"imported {added} new song(s) into {DEFAULT_LIBRARY_DIR}")


if __name__ == "__main__":
    sys.exit(main())
