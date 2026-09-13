---
name: ppt
description: Use when building or rebuilding this church's Sunday-service slide deck (PowerPoint) from a song list, when extracting hymn lyrics out of the church's himnario PDF by hymn number, or when a new week's PPTX needs the same background/fonts/logo as a previous week's deck.
---

# ppt (worship deck builder)

## Overview

Rebuilds this church's Sunday-service PPTX by **cloning real slide shapes**
(backgrounds, fonts, logo, positions) from a previous week's deck and swapping in
this week's text, instead of rebuilding formatting by hand or re-deriving fonts/
colors. This is a single-purpose, single-church skill: it always works from this
church's own reference files, not a generic path you pass in each time.

## When to use

- "build/update the slides for Sunday" with a song list for the week
- "pull the lyrics for Himno ## from the himnario"
- a deck needs the same background/fonts/logo as a previous week's PPTX, just new
  songs

## Reference files (this church's own, already in place)

- `reference/style-template.pptx` — the church's most recent real deck. Default
  style source for every build; only pass `"template"` in the spec to override it
  for a one-off.
- `reference/hymnal.pdf` — the church's real himnario PDF. Default hymn source;
  only pass `"hymnal"` in the spec to override it.
- `reference/song-library/` — this church's growing library of **unique
  contemporary songs**, one `<slug>.json` file per song (see
  `reference/song-library/_example.json` for the shape). This is the source the
  builder consults for non-hymn songs, so you don't have to dig through an old
  deck or the web each week. Add songs with `song add` / `song import-deck` (below).
  The legacy flat `reference/song-library.json` is still read if present.
- `reference/example-template.pptx` / `reference/example-hymnal.pdf` — synthetic
  placeholders used only by `examples/songs.example.json` to demo/test the tool.
  Not used by a real build.
- `reference/slide-types.md` — the 5 reusable slide shapes, in detail.

## Workflow

1. **Identify each song.** If it's a hymn, get its hymn number. In this himnario,
   **PDF page number equals hymn number** — try
   `scripts/pptx_deck_builder.py hymn --page <N>` first (it reads
   `reference/hymnal.pdf` by default); only text-search the PDF if that doesn't
   match.
2. **Parse the hymn page** into verses (numbered `1/2/3/4…`) and a chorus (`Coro.-`
   marker) — `parse_hymn()` in `scripts/pptx_deck_builder.py` does this. Treat the
   result as a draft, not gospel: some hymns won't fit this exact format.
3. **Chunk each stanza into slide-sized pieces.** There's no fixed formula for lines
   per slide (see `reference/slide-types.md`) — set `verse_chunk_size` /
   `chorus_chunk_size` per song in the build spec (a `"whole"` chorus is common for
   short ones), then look at the rendered output and adjust.
4. **Contemporary (non-hymnal) songs come from the local `reference/song-library/`
   folder, never from the open web.** Library-first: for each non-hymn song, look
   it up in the library and reference it by `key`; if it's missing, flag it for a
   human to add (typed in, or exported from a licensed source like CCLI SongSelect)
   — then save it with `song add` / `song import-deck` so it's there next time. A
   live web search is useful only to *identify* an unfamiliar song (title/artist);
   generic lyrics sites are unlicensed and often wrong (bad verse order, missing
   repeats, typos), so never source display text from them.
5. **Build the deck.** `scripts/pptx_deck_builder.py build --spec songs.json`
   clones each needed slide's real shape XML (and re-links its embedded images)
   from `reference/style-template.pptx`, then swaps in the new text — this is what
   keeps fonts/colors/backgrounds pixel-identical instead of reverse-engineered.
6. **Verify before Sunday.** Reopen the output with `python-pptx`, zip-integrity
   check it, and render to PDF/PNG (`soffice --headless --convert-to pdf`) to
   eyeball it. The chunking heuristic and any new song entries need a human look.

## Build spec (songs.json)

`template`, `hymnal`, and `library` are all optional — they default to
`reference/style-template.pptx`, `reference/hymnal.pdf`, and
`reference/song-library.json`. A normal week's spec only needs `output` and
`items`:

```json
{
  "output": "new-week.pptx",
  "format": { "capitalize_lines": true, "break_at": 30 },
  "items": [
    { "op": "clone_range", "start": 1, "end": 3 },
    { "op": "hymn", "himno": 64, "verse_chunk_size": 2, "chorus_chunk_size": 2 },
    { "op": "song", "title_white": "Hoy te Rindo", "title_cream": "mi Ser", "key": "hoy-te-rindo-mi-ser" },
    { "op": "scripture", "book": "San Mateo", "range": "5:14-16",
      "chunks": [ ["14", "Vosotros sois la luz del mundo;", "", "15", "Ni se enciende una luz..."] ] },
    { "op": "clone_range", "start": 42, "end": 66 }
  ]
}
```

Ops:
- `clone_range` — reuse slides verbatim (1-indexed, inclusive) from `template`
  (or another deck via `"source"`): intro slides, purpose statement, anything
  already correct in a past deck.
- `hymn` — pull `himno` N from the hymnal. `title` is optional; if omitted it's
  taken from the hymnal page's ALL-CAPS header (verify accents — the header has
  none, e.g. `Halle` for `Hallé`). **Never edit the hymnal's wording**; only the
  chunk sizes are yours to tune.
- `song` — a contemporary song. Either `"key"` (from `reference/song-library/`)
  or inline `title_white`/`title_cream` + `sections`. Renders a two-tone title
  (main phrase in white, secondary part in cream — keep the white part short so
  the big font stays on one line) plus lyric slides.
- `scripture` — a reference slide (`book` + `range`) then one slide per `chunks`
  entry of numbered verse text (blank line between verses). Set verbatim.
- `library_song` — legacy alias of `song` by key.

`format` (optional, church defaults shown): `capitalize_lines` capitalizes the
first letter of every lyric line; `break_at` splits any lyric line longer than N
chars at a comma into shorter centered lines. Repeat markers `//…//` / `///…///`
are shown literally (a cue to sing 2×/3×). Scripture text is never reformatted.

`title_slide_index` / `lyric_slide_index` / `song_title_slide_index` /
`scripture_ref_slide_index` / `scripture_text_slide_index` are optional overrides
when template auto-detection picks the wrong slide. Full reference:
`scripts/pptx_deck_builder.py build --help`.

## Growing the song library

```
python3 scripts/pptx_deck_builder.py song list
python3 scripts/pptx_deck_builder.py song add --file new-song.json
python3 scripts/pptx_deck_builder.py song import-deck --deck ../old-deck.pptx
```
`song import-deck` mines an existing deck's contemporary-song title + lyric
slides into the library (skipping hymns and scripture) — run it over your past
decks once to backfill, and it captures each new week's new songs going forward.

## Common mistakes

- Treating `verse_chunk_size`/`chorus_chunk_size` as exact — always render and
  spot-check. Two long lines on one slide overflow; give a long line its own slide.
- **Editing the hymnal's words** (e.g. "correcting" a spelling) — preserve the
  himnario text; only fix OCR glue (numbers stuck to text).
- Sourcing contemporary lyrics from a generic lyrics website instead of the
  `reference/song-library/` folder.
- Rebuilding fonts/positions by hand instead of cloning from the template.

## Script usage

```
uv pip install -r requirements.txt
python3 scripts/pptx_deck_builder.py build --spec songs.json
python3 scripts/pptx_deck_builder.py hymn --page 64   # debug helper
python3 tests/test_builder.py                         # run the tests
```
