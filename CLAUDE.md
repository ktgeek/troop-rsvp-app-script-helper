# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A Google Apps Script (TypeScript) bound to the Troop 66 RSVP spreadsheet. It adds a "Troop 66" menu that turns the
Google Form responses into printable/working sheets: patrols, driving assignments, and a printable roster.
All logic lives in `Code.ts`.

## Commands

- `npm run build` — `tsc`, transpiles `Code.ts` to `dist/Code.js` (`module: none`, ES2017, strict)
- `npm run lint` — ESLint (typescript-eslint + prettier as an error, `no-console` as a warning)
- `npm run push` — build, then `clasp push` to deploy; `.clasp.json` has `rootDir: dist`, so pushing without
  building would deploy stale JS
- No test suite exists. The code needs the Apps Script runtime (`SpreadsheetApp`), so verify by pushing and running
  from the sheet's menu.

`npm run prettier-format` rewrites `Code.ts` in place. Prettier style: no semicolons, single quotes, no bracket
spacing, no arrow-paren, 100 columns.

## Architecture

- Menu entry points (`onOpen`, `createPatrolListing`, `createOutingListing`, `createDrivingAssignments`) are
  top-level functions because Apps Script resolves menu handlers by global name. Each carries an
  `eslint-disable no-unused-vars` comment, since nothing in the TS calls them. Add the same when adding a menu item,
  and register the item in `createMenu`.
- `module: none` is deliberate: no `import`/`export`, so everything stays in the global scope Apps Script expects.
- `responseData()` reads the `Form Responses 1` sheet (display values) and drops the header row. Every generator
  consumes that and reads fields **by hard-coded column index** (e.g. `row[2]` scout name, `row[3]` patrol,
  `row[15]` driver option, `row[17]` seat count). If the form changes its questions or order, update every index.
  Recent commits ("track rename of adult columns") show this being a recurring maintenance point.
- `createNamedSheet(name)` renames any existing sheet of that name to `"<name> (old)"` and inserts a fresh one.
  Running a generator twice stacks `(old)` sheets; it does not overwrite.
- Domain constants at the top: `ADULTS` (patrol values meaning an adult is camping, matched against the form's
  option text) and `DRIVERS` (maps form option text to a short label used in the driving sheet).
- Driving sheet: columns are camping adults plus parents who chose a driving option AND supplied a seat count;
  rows are scouts. Seat count is shown as `parseInt(count) + 1` (driver included).
