# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```sh
npm run push   # local src/ -> the bound Apps Script project (overwrites remote files)
npm run pull   # remote -> src/
npm run open   # open the project in the Apps Script IDE
npx clasp show-file-status   # which files a push would upload
```

`push` replaces remote files wholesale, so pull first if anyone edited the script in the browser.
There is no build, no lint and no test runner; `typos` is the spell check.

## Style

Keep lines at 120 characters or shorter, in code and in Markdown alike.
Start every sentence on its own line, and let a sentence wrap only when it does not fit.

## Deployment model

The code runs as a container-bound Apps Script project inside the "NRC. Coffee" Google Sheet, and it can only be
exercised there.
Nothing in `src/` runs under Node: there are no modules, no `require`, no `import`.
Every file shares one global scope, which is why `Code.js` calls `madHeadsSlug` and `madHeadsItem` without importing
them.

`.clasp.json` sets `rootDir: "src"`, so clasp treats `src/` as the project root and uploads everything in it.
Do not add a `.claspignore` with paths written from the repo root; clasp resolves those patterns against `rootDir` and
they silently exclude all files.

`src/appsscript.json` is the manifest and must stay in `src/`.
OAuth scopes are granted interactively on first run, not declared there.

## Layering

`MadHeads.js` holds everything shop-specific and touches no Sheets API, so it can be exercised outside Apps Script.
`Code.js` holds everything sheet-specific and knows nothing about the shop's data shape.
Keep that split: a second shop (3champsroastery is expected) should arrive as its own file, not as branches inside these
two.

### MadHeads.js

Product data comes from the shop's SvelteKit endpoint, `/p/<slug>/__data.json`, never from scraped HTML.
That payload is flattened: every number in the array is an index into the same array, which is what `unflatten`
rebuilds.

A cart link identifies a *variant*, not a product, so the slug in the URL must be matched against `variants[].slug`.
Roast and grind are variant options keyed by `property_id` 27 and 31; a drip has no 31, which is how it gets `-`.
The price is always `price_wholesale`, never `price_retail`.

### Code.js

Columns are located by header text rather than by letter, because sheets differ between orders (Обсмажка in newer ones,
Обсмаження in older).
The header row is searched for in the first 10 rows.

Cells are written one at a time on purpose.
Writing whole columns with `setValues` would replace formulas in untouched rows with their computed values, and several
sheets in that spreadsheet hold formulas in these columns.

Rows are classified by their link, not by sheet name or column header: one existing sheet carries a MadHeads header over
3champsroastery links.

## Verifying changes

`MadHeads.js` can be checked from Node without mocking the shop: evaluate the file with a `UrlFetchApp` shim that shells
out to `curl`, then assert against known cart links.
Sheet code has no equivalent and has to be run from the spreadsheet's **Кава** menu, with errors read from the Apps
Script execution log.
