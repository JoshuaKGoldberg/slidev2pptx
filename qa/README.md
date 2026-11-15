# Round-Trip QA

This harness checks slidev2pptx against real talks by round-tripping PowerPoint decks:

1. `pptx2slidev/` converts a corpus `.pptx` into a Slidev deck, logging what it loses
2. Slidev exports the converted deck to PNGs in Chromium, one per click state
3. slidev2pptx builds the Slidev deck back into a `.pptx`, once with `--with-clicks` and once animated
4. PowerPoint renders slidev2pptx's `.pptx`, and the original `.pptx`, to PDF, which `pdftoppm` turns into PNGs
5. `compare/` scores each click state's Slidev rendering against slidev2pptx's with ImageMagick, as `pass`, `review`, `fail`, or `missing`

Each slide's last click state is also compared against the original deck's rendering, both for the converter alone and end to end.
PowerPoint for Mac now and then draws text without its font's kerning, so the harness has PowerPoint kern before it renders, and renders states that don't pass up to twice more, keeping the closest render.

## Requirements

- macOS with Microsoft PowerPoint, which renders decks by AppleScript, one at a time
- [poppler](https://poppler.freedesktop.org) for `pdftoppm`, [ImageMagick](https://imagemagick.org) for `magick`, and [fontconfig](https://www.freedesktop.org/wiki/Software/fontconfig) for `fc-list` and `fc-scan`, such as with `brew install poppler imagemagick fontconfig`
- Chromium for Slidev's export: `pnpm exec playwright install chromium`
- The corpus decks' fonts installed, or their text falls back and the converter logs them as missing

## Getting the Corpus

[`corpus/manifest.json`](./corpus/manifest.json) lists the decks: each one's id, title, events, and OneDrive or SharePoint share link.
The decks themselves aren't committed.
To get them:

1. Download the decks linked from the manifest's `url`s into one directory, naming each file by its deck's `number`, like `1.pptx` or `01-whats-new.pptx`
2. Run `node qa/corpus/import.ts path/to/downloaded/decks`

`import.ts` checks each file's text against its deck's title, printing `CHECK` for a doubtful match, then copies it to `corpus/decks/<id>.pptx` and records its SHA-256 and slide count in the manifest.
It lists the decks still missing.

`corpus/manifest.ts` rebuilds the manifest from a directory of speaking entries' Markdown, for when talks are added.

## Scripts

Run these from the repository root, in about this order:

1. `node qa/run.ts <deck id> [--from <step>]` round-trips one deck through the steps `convert`, `slidev`, `build`, `render`, and `compare`.
   A step whose output already exists is skipped, unless `--from` names it or an earlier step, so `--from build` rebuilds with the current source but keeps the converted deck and its Slidev rendering.
2. `node qa/batch.ts [--concurrency 3] [--from convert] [deck ids...]` runs `run.ts` for many decks, by default every imported one from `convert` on, three at a time.
   It retries each failure once, keeps the Mac and its display awake with `caffeinate`, and writes each deck's output to `runs/<id>.log`.
3. `node qa/check.ts <deck id> <output directory> [states...]` checks a fix quickly: it builds one deck from the working tree into the output directory, renders the named states (like `021-01`) or else every state that didn't pass, and prints each one's SSIM in the last run and now.
   It leaves `runs/` as it is, so it can run while a batch uses the runs.
4. `node qa/summarize.ts` collects every deck's `runs/<id>/report.json` into [`reports/summary.md`](./reports/summary.md).
5. `node qa/census/census.ts` counts which PowerPoint features the corpus decks use into [`reports/census.md`](./reports/census.md) and `reports/census.json`.

## Outputs

`corpus/decks/` and `runs/` are gitignored.
Each deck's run in `runs/<id>/` holds:

- `slidev/`: the converted deck, with `conversion-log.json` of what the converter lost
- `slidev-png/`: Slidev's rendering
- `pptx/`: slidev2pptx's `clicks.pptx` and `animated.pptx`, with the problems each build reported
- `pptx-png/` and `original-png/`: PowerPoint's renderings of slidev2pptx's deck and the original
- `diff/` and `sheets/`: difference images, and contact sheets of the states that didn't pass
- `report.json`: every state's scores and verdict, the build problems, and step timings

Everything else is committed, including [`reports/`](./reports/):

- `summary.md`: every deck's verdict counts and SSIMs, from `summarize.ts`
- `census.md` and `census.json`: the corpus's feature counts, from `census/census.ts`
- `review-round-*.md`: each round of reviewing the results, with its root causes, their status, and what its fixes changed
