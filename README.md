# worship-deck-builder

## Install

Paste this into a Claude Code chat:

```
Install the ppt skill:
1. Clone https://github.com/danihrndzld/worship-deck-builder into ~/.claude/skills/ppt
2. Ask me for the path to my church's hymnal PDF and my most recent service PPTX
3. Copy those two files into that skill's reference/ folder as reference/hymnal.pdf
   and reference/style-template.pptx (leave the example-*.pdf/pptx placeholders as they are)
4. Run `uv pip install -r requirements.txt` inside the installed skill folder
5. Confirm the skill is installed and tell me how to call it
```

Claude will clone the repo, copy your two real files into the *local* install
only (they're gitignored, so they never get committed or pushed anywhere),
install the Python dependencies, and confirm it's ready.

## How to call it

- Explicitly: `/ppt`
- Naturally: just ask, e.g. "build the Sunday slides", "update the service
  deck with this week's songs", or "pull the lyrics for Himno 64 from the
  himnario" — Claude matches these against the `description` in `SKILL.md`
  and loads the skill on its own.

## What it actually does

- Clones a slide's shape XML (background image, logo, text box position, font,
  color) from a template deck instead of rebuilding formatting by hand. The
  output is styled identically to the source because it *is* the source's
  shapes with new text swapped in.
- Reads a hymn's page out of your hymnal PDF and splits it into verses and a
  chorus. In at least one hymnal, page number equals hymn number, so this is
  often a direct lookup rather than a search.
- Chunks each verse/chorus into slide-sized pieces. There's no fixed number of
  lines per slide (it varies by stanza), so this is a tunable heuristic you
  check by rendering the result, not a black box.
- Keeps a local, church-owned library for contemporary (non-hymnal) songs, and
  never fetches lyrics from the open web to fill it in. More on why below.

## Why the `reference/` files are placeholders, not the real thing

`reference/example-template.pptx` and `reference/example-hymnal.pdf` exist so
you can see the expected slide shapes and PDF text format without this repo
shipping anyone's actual copyrighted material. Every hymn, verse, and chorus
line in them was written for this repo. None of it is transcribed from a
real hymnal or a real song.

The reason: a hymnal is a copyrighted compiled work, and most contemporary
worship songs are commercially released and copyrighted. A church typically
has the right to display these for its own congregation (a CCLI license or
similar usually covers that), but that license doesn't extend to publishing
the same PDF or the same lyrics-filled PPTX in a public GitHub repo for
anyone to download. So: point this tool at your own hymnal and your own
deck, locally, and keep those files out of version control (the `.gitignore`
here already excludes `*.pptx` / `*.pdf` other than the two placeholders).

The same reasoning is why step 4 above (the song library) is filled in by
you, from material your church already has the right to display, and not by
scraping a lyrics website. A live search is still useful for *identifying* an
unfamiliar song (checking the title and artist against what a search turns
up), but the sites that show up for a lyrics search are unlicensed
aggregators, and the text on them is frequently wrong: bad verse order,
missing repeats, typos. Neither the license question nor the accuracy
question gets better by automating the fetch.

## Other ways to install

**As a Claude Code skill (personal, all projects):**

```bash
git clone https://github.com/danihrndzld/worship-deck-builder ~/.claude/skills/ppt
```

Claude will pick it up next session. Ask it to "build the Sunday slides"
and it'll read `SKILL.md` and follow the workflow there.

**As a project skill (one repo only):**

```bash
git clone https://github.com/danihrndzld/worship-deck-builder .claude/skills/ppt
```

**As a standalone script (no Claude Code):**

```bash
git clone https://github.com/danihrndzld/worship-deck-builder
cd worship-deck-builder
uv pip install -r requirements.txt
```

## Use

1. `template` and `hymnal` don't need to be set — they default to
   `reference/style-template.pptx` and `reference/hymnal.pdf`, the two real files
   the install step copied in. Only set them in a spec if a particular build needs
   a different source.
2. If you have contemporary (non-hymnal) songs, add them to
   `reference/song-library.json` (copy `reference/song-library.example.json` for
   the shape) — this also defaults automatically once it exists.
3. Write this week's `songs.json`, listing songs under `items` — `hymn` for
   anything in the hymnal (by hymn number), `library_song` for anything in your
   library (by key), `clone_range` for anything you're reusing unchanged from a
   past deck (intro slides, the purpose statement, a scripture reading).
4. Build it:

   ```bash
   python3 scripts/pptx_deck_builder.py build --spec songs.json
   ```

5. Open the output (or render it with `soffice --headless --convert-to pdf
   your-deck.pptx` and check the PDF) before Sunday. The line-chunking
   heuristic and any brand-new song entries are worth a human glance.

Full spec format and CLI flags: `SKILL.md`, or `--help` on the script.

## Try it without any real files

```bash
cd examples
python3 ../scripts/pptx_deck_builder.py build --spec songs.example.json
```

builds `example-output.pptx` from the placeholder template, hymnal, and
library. Useful for confirming the tool works before pointing it at your
real material.

## License

MIT for the code — see `LICENSE`. The placeholder `reference/` files are
original and MIT-licensed too. Whatever real hymnal, deck, or song text *you*
point this at keeps whatever rights/license it already has; that's on you to
clear, same as it would be if you built the slides by hand.
