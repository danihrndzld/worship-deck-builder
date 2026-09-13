"""Tests for pptx_deck_builder. Run: `python3 tests/test_builder.py` (or pytest).

Uses only the committed synthetic reference files, never real church material.
"""
import io
import sys
import tempfile
import zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "scripts"))
import pptx_deck_builder as B  # noqa: E402
from pptx import Presentation  # noqa: E402
from pptx.oxml.ns import qn  # noqa: E402

REF = REPO / "reference"


def test_parse_hymn_handles_dot_glued_and_lowercase():
    # "1." separator, a verse number glued onto text ("3¡"), and a lowercase
    # verse start ("4 la") -- all three used to be mishandled.
    raw = "\n".join([
        "283", "BRILLA EN EL SITIO", "Key E Flat.",
        "1. Primera linea del verso uno,", "segunda linea del verso uno.",
        "Coro.- Linea del coro,", "otra linea del coro.",
        "2. Primera del dos,", "segunda del dos.",
        "3¡Glued al numero!", "segunda del tres.",
        "4 la fe empieza en minuscula,", "segunda del cuatro.",
    ])
    p = B.parse_hymn(raw)
    assert len(p["verses"]) == 4, p["verses"]
    assert p["verses"][2][0] == "¡Glued al numero!"
    assert p["verses"][3][0] == "la fe empieza en minuscula,"
    assert p["chorus"][0] == "Linea del coro,"


def test_hymn_title_from_caps_header():
    raw = "48\nHALLE UN BUEN AMIGO\nKey F.\n1 Halle un buen amigo,"
    assert B.hymn_title(raw) == "Halle Un Buen Amigo"


def test_capitalize_and_break_long_line():
    assert B.capitalize_first("no hay otro manantial") == "No hay otro manantial"
    assert B.break_long_line("Renuevame, Senor Jesus, pon en mi corazon", 20) == \
        ["Renuevame,", "Senor Jesus,", "pon en mi corazon"]
    assert B.break_long_line("linea corta", 30) == ["linea corta"]


def test_set_textbox_lines_never_drops_a_line():
    # Build a text box whose 3rd paragraph has NO run (blank separator), then
    # set 3 real lines: the old code dropped the line on the run-less paragraph.
    prs = Presentation(str(REF / "example-template.pptx"))
    box = B.get_textboxes(list(prs.slides)[9])[0]  # scripture text box "1..2"
    tf = box.text_frame
    # ensure there is a run-less paragraph in the middle
    from pptx.oxml.ns import qn as _qn
    blank = tf.paragraphs[0]._p.makeelement(_qn("a:p"), {})
    tf._txBody.insert(list(tf._txBody).index(tf.paragraphs[1]._p), blank)
    B.set_textbox_lines(box, ["uno", "dos", "tres"])
    assert [p.text for p in tf.paragraphs] == ["uno", "dos", "tres"]


def test_clone_slide_copies_background():
    prs = Presentation(str(REF / "example-template.pptx"))
    src = list(prs.slides)[0]
    # give the source slide a <p:bg>
    cSld = src._element.find(qn("p:cSld"))
    from lxml import etree
    bg = etree.SubElement(cSld, qn("p:bg"))
    bgpr = etree.SubElement(bg, qn("p:bgPr"))
    fill = etree.SubElement(bgpr, qn("a:solidFill"))
    clr = etree.SubElement(fill, qn("a:srgbClr")); clr.set("val", "000343")
    etree.SubElement(bgpr, qn("a:effectLst"))
    cSld.insert(0, bg)
    out = Presentation(str(REF / "example-template.pptx"))
    B.delete_all_slides(out)
    new = B.clone_slide(src, out)
    assert new._element.find(qn("p:cSld")).find(qn("p:bg")) is not None


def test_example_build_and_import_deck_roundtrip():
    with tempfile.TemporaryDirectory() as d:
        out = Path(d) / "out.pptx"
        spec = Path(d) / "spec.json"
        spec.write_text(f'''{{
          "template": "{REF}/example-template.pptx",
          "hymnal": "{REF}/example-hymnal.pdf",
          "library": "{REF}/song-library.example.json",
          "output": "{out}",
          "items": [
            {{"op":"clone_range","start":1,"end":3}},
            {{"op":"hymn","himno":1}},
            {{"op":"song","title_white":"Titulo","title_cream":"Cancion",
              "sections":[{{"type":"verse","lines":["Linea uno,","linea dos."]}}]}},
            {{"op":"scripture","book":"Libro","range":"1:1-2",
              "chunks":[["1","Texto uno.","","2","Texto dos."]]}}
          ]
        }}''')
        B.build(str(spec))
        assert zipfile.ZipFile(out).testzip() is None
        Presentation(str(out))  # reopens cleanly
        songs = B.songs_from_deck(str(out))
        titles = [s["title"] for s in songs]
        assert "Titulo Cancion" in titles  # the contemporary song was extracted
        assert not any("Himno" in t for t in titles)  # hymn/scripture skipped


def _run():
    fns = [v for k, v in sorted(globals().items()) if k.startswith("test_")]
    for fn in fns:
        fn()
        print(f"ok  {fn.__name__}")
    print(f"\n{len(fns)} passed")


if __name__ == "__main__":
    _run()
