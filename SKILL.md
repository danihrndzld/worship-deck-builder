---
name: worship-deck-builder
description: Use when building or rebuilding a church Sunday-service slide deck (PowerPoint) from a song list, when extracting hymn lyrics out of a hymnal/himnario PDF by hymn number, or when a new week's PPTX needs the same background/fonts/logo as a previous week's deck.
---

# Worship Deck Builder

## Overview

Rebuilds a Sunday-service PPTX by **cloning real slide shapes** (backgrounds, fonts,
logo, positions) from a previous week's deck and swapping in this week's text,
instead of rebuilding formatting by hand or re-deriving fonts/colors. Two source
documents drive it: a **style-template PPTX** (any recent deck from the church) and
a **hymnal PDF** (the church's own himnario) for songs that are public-domain or
already-licensed hymns.

## When to use

- "build/update the slides for Sunday" with a song list for the week
- "pull the lyrics for Himno ## from the himnario"
- a deck needs the same background/fonts/logo as a previous week's PPTX, just new
  songs

## Reference files

- `reference/example-template.pptx` / `reference/example-hymnal.pdf` — **synthetic**
  stand-ins (placeholder text, generic styling) that show the slide shapes and PDF
  format this tool expects. Point the real run at your own church's actual deck and
  hymnal instead — see README.md for why these two aren't the real thing.
- `reference/slide-types.md` — the 5 reusable slide shapes, in detail.
- `reference/song-library.example.json` — shape of the local song library described
  below.

## Workflow

1. **Identify each song.** If it's a hymn, get its hymn number. In at least one
   common hymnal, **PDF page number equals hymn number** — try
   `scripts/pptx_deck_builder.py hymn --hymnal <file> --page <N>` first; only
   text-search the PDF if that doesn't match.
2. **Parse the hymn page** into verses (numbered `1/2/3/4…`) and a chorus (`Coro.-`
   marker) — `parse_hymn()` in `scripts/pptx_deck_builder.py` does this. Treat the
   result as a draft, not gospel: some hymns won't fit this exact format.
3. **Chunk each stanza into slide-sized pieces.** There's no fixed formula for lines
   per slide (see `reference/slide-types.md`) — set `verse_chunk_size` /
   `chorus_chunk_size` per song in the build spec (a `"whole"` chorus is common for
   short ones), then look at the rendered output and adjust.
4. **Contemporary (non-hymnal) songs come from your own song library, never from
   the open web.** Maintain `reference/song-library.example.json`-shaped entries
   (verses/chorus you've typed in yourself, or exported from a licensed source like
   CCLI SongSelect) and reference them by key in the build spec. A live web search
   is useful only to *identify* an unfamiliar song (title/artist) — generic lyrics
   sites are unlicensed and often wrong (bad verse order, missing repeats, typos);
   never source display text from them. Anything not already in the library gets
   flagged for a human to type in, not fetched automatically.
5. **Build the deck.** `scripts/pptx_deck_builder.py build --spec songs.json` clones
   each needed slide's real shape XML (and re-links its embedded images) from the
   style template, then swaps in the new text — this is what keeps fonts/colors/
   backgrounds pixel-identical instead of reverse-engineered.
6. **Verify before Sunday.** Reopen the output with `python-pptx`, zip-integrity
   check it, and render to PDF/PNG (`soffice --headless --convert-to pdf`) to
   eyeball it. The chunking heuristic and any new song entries need a human look.

## Build spec (songs.json)

```json
{
  "template": "last-week.pptx",
  "hymnal": "himnario.pdf",
  "library": "song-library.json",
  "output": "new-week.pptx",
  "title_slide_index": 4,
  "lyric_slide_index": 5,
  "items": [
    { "op": "clone_range", "start": 1, "end": 3 },
    { "op": "hymn", "title": "...", "himno": 64, "verse_chunk_size": 2, "chorus_chunk_size": 2 },
    { "op": "library_song", "key": "song-key-in-library" },
    { "op": "clone_range", "start": 42, "end": 66 }
  ]
}
```

- `clone_range` reuses slides verbatim (1-indexed, inclusive) from `template`, or
  from another deck via an optional `"source"` field — use it for anything that
  didn't change (intro slides, purpose statement, a song already correct in a past
  deck, the scripture reading).
- `title_slide_index` / `lyric_slide_index` are optional — omit them to
  auto-detect (a slide with 2 text boxes where the second matches `Himno \d+` is
  the title template; the first 1-box, multi-line slide is the lyric template).
  Pass them explicitly if auto-detection picks the wrong slide.
- Full field reference: `scripts/pptx_deck_builder.py build --help`.

## Common mistakes

- Treating `verse_chunk_size`/`chorus_chunk_size` as exact — always render and
  spot-check.
- Sourcing contemporary lyrics from a generic lyrics website instead of the
  church's own library.
- Rebuilding fonts/positions by hand instead of `clone_range`/cloning from the
  style template.

## Script usage

```
pip install -r requirements.txt
python3 scripts/pptx_deck_builder.py build --spec songs.json
python3 scripts/pptx_deck_builder.py hymn --hymnal himnario.pdf --page 64   # debug helper
```
