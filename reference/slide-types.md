# Slide shapes

Reverse-engineered from a real church's week-over-week decks: the same 5 slide
shapes get reused every week, just with new text. `reference/example-template.pptx`
demonstrates each one (slide numbers below match that file).

| # | Shape | Text boxes | Notes |
|---|-------|-----------|-------|
| 1-2 | Intro / waiting slides | none | Background only, shown before the service starts. |
| 3 | Purpose statement | 2 header boxes + 1 paragraph box | Fixed org text, doesn't change week to week. |
| 4 | Hymn title | 2 boxes: song title, `Himno ##` | Second box is what `detect_template_slides()` matches on (`^himno\s+\d+$`). |
| 5, 7 | Lyric slide | 1 box, 2+ lines | One verse/chorus chunk per slide. Auto-fit shrinks text to fit as line count grows. |
| 6 | Contemporary song title | 2 boxes, both part of the title (no hymn number) | Used for songs not in the hymnal. |
| 8 | Scripture section title | topic + reference range | Opens a scripture-reading segment. |
| 9 | Scripture reference | book + verse range | One per reference if there's more than one. |
| 10 | Scripture verse text | numbered verses | Same shape as a lyric slide; numbers stay inline with the text. |

## What's NOT fixed

Line count per lyric slide is **not** a formula — it ranges 2-5 depending on the
stanza and looks tuned by a human for on-screen readability, not computed from a
rule. `pptx_deck_builder.py` exposes `verse_chunk_size` / `chorus_chunk_size` per
song (including a `"whole"` option to keep a short chorus on one slide) so you can
match the original operator's judgment call — but always render and eyeball the
result before Sunday.

## Reproducing this from your own deck

1. Open your church's most recent PPTX and list each slide's text boxes (`python-pptx`:
   `shape.text_frame.text` for each shape).
2. Confirm the same shapes above show up in the same order per song.
3. Note which slide index holds the hymn-title shape and which holds the lyric
   shape — pass them as `--title-slide-index` / `--lyric-slide-index` if
   auto-detection (see `SKILL.md`) picks the wrong one.
