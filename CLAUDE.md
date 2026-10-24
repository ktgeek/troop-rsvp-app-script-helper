# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A Google Apps Script (TypeScript) bound to the Troop 66 RSVP spreadsheet. It adds a "Troop 66" menu that turns the
Google Form responses into printable/working sheets: patrols, driving assignments, and a printable roster.
Source lives in `src/`, one file per feature (see Architecture).

## Commands

- `npm run build` — clears `dist/`, then `tsc` transpiles `src/*.ts` to a flat `dist/*.js` (`module: none`, ES2017,
  strict). The clear keeps renamed or deleted sources from leaving stale files that `clasp push` would upload.
- `npm run lint` — ESLint (typescript-eslint + prettier as an error, `no-console` as a warning)
- `npm run push` — build, then `clasp push` to whatever script `.clasp.json` names; `.clasp.json` has
  `rootDir: dist`, so pushing without building would deploy stale JS
- `npm run deploy -- <sheet URL | sheet ID | saved name>` — build, then push to that sheet's bound script; see
  `README.md` for setup (`clasp login`) and usage. The sheet -> script ID map lives in gitignored `targets.json`,
  each run writes a gitignored `.clasp.deploy.json` (never `.clasp.json`, which `npm run push` uses), and
  `appsscript.json` is pushed too. Re-running it is how an existing
  deployment is updated. There is deliberately no auto-create for unknown sheets: `clasp create --type sheets`
  ignores `--parentId` and makes a new blank spreadsheet, so each sheet is recorded once with `--add`.
- No test suite exists. The code needs the Apps Script runtime (`SpreadsheetApp`), so verify by pushing and running
  from the sheet's menu.

`npm run prettier-format` rewrites `src/*.ts` in place. Prettier style: no semicolons, single quotes, no bracket
spacing, no arrow-paren, 100 columns.

## Architecture

- Files: `Code.ts` (`onOpen`, `createMenu`), `Responses.ts` (`ADULTS`, `FIELDS`, `responseData`, `createNamedSheet`),
  `Patrols.ts` (`createPatrolListing`, `sortPatrolSheet`), `PatrolRoster.ts`, `DrivingPrint.ts`, `PrintDialog.ts`,
  `Driving.ts` (with `DRIVERS`), `Roster.ts`.
- `Roster.ts` `rosterGroups()` turns responses into printable-roster groups (Adult Leaders, then patrols): it cleans
  names and phones, drops "none" answers, merges campout patrol/role/notes from the `Patrols` sheet, and attaches
  `issues` (the VERIFY reasons) to each person. Scouts are grouped by campout patrol once
  any scout has one other than the adults' `A` (`hasCampoutAssignments`); otherwise, or with no `Patrols` sheet, by the
  patrol they signed up under. Each group carries `roles`, its PoRs ranked, shown in the group band.
- The "print roster" menu item (`showRoster`) opens `src/RosterDialog.html` in a modal dialog. The page fetches
  `getRosterPage()` (groups plus saved settings) via `google.script.run`, renders the cards with `textContent`
  only, and its Print button calls `window.print()`; the user picks "Save as PDF". Print layout (page breaks, footer,
  page numbers via `@page` margin boxes) is all CSS. Title/date/lead/VERIFY-flag settings are stored by
  `saveRosterSettings` in document properties. The build copies every `src/*.html` into `dist/` (tsc doesn't), and
  `clasp push` uploads them from there. Apps Script needs every file's name unique regardless of type, so
  `Roster.html` could not sit beside `Roster.js`; keep base names distinct across `src/*.ts` and the HTML pages.
- `sortPatrolSheet()` sorts the `Patrols` data rows in JS (Campout Patrol as plain text, then Patrol, then last name
  via `compareByLastName`) and writes the values back; it runs at the end of `createPatrolListing` and from the menu.
  `lastNameKey` keeps particles like "Van" or "de" (`SURNAME_PARTICLES`) with the surname; add to that list if a
  troop name sorts wrong.
- `rosterPerson` spells known PoRs out for both rosters (`roleDisplayName`, from the `ROLES` table that also sets the
  rank order); unrecognized text prints as typed.
- The "print patrol roster" item (`showPatrolRoster`) is the scout-facing sibling of the roster: `PatrolRoster.ts`
  maps `rosterGroups()` to names, PoRs, allergies + dietary on one
  line, and only the Patrols-sheet notes (`patrolNotes`), and `PatrolRosterDialog.html` renders it. All three print
  dialogs keep their header fields through `readSettings`/`writeSettings` in `PrintDialog.ts` (`HeaderSettings` is the
  title/date pair). The dialogs share their styles and page script (`el`, `masthead`, `wireControls`, `loadDialog`)
  through `PrintShared.html`, pulled in by `include()` from a template (`showPrintDialog` evaluates them).
- The "print driving sheet" item (`showDrivingPrint`) prints the hand-marked `Driving` tabs, not form data: every tab
  whose name starts with `Driving` except `Driving (old...)` becomes a `DrivingPage`. Driver headers are parsed back
  out of the `Name (label: seats)` text `drivingText` wrote (`parseDriverHeader`); seats are the
  scout seats, which already count the driver's own scout (the form asks for seats not including them), so a car is over
  capacity above `seats`
  scouts. Any non-empty cell except a checkbox's "FALSE" counts as an assignment, and a tab
  with text in A1 is not treated as a driving sheet. `UNNAMED_DRIVER_MARK` (`!!`, set in `Driving.ts`) marks a column
  whose parent gave no driver name.
- Menu entry points (`onOpen`, `createPatrolListing`, `sortPatrolSheet`, `showRoster`, `showPatrolRoster`,
  `showDrivingPrint`, `createDrivingAssignments`) are top-level functions because Apps Script resolves menu handlers by
  global name. Register new ones in `createMenu`.
- `module: none` is deliberate: no `import`/`export`, so everything stays in the global scope Apps Script expects,
  and every pushed file shares it. Files may use each other's functions and constants, but no top-level statement
  may evaluate something defined in another file (push order isn't guaranteed); reference them only inside
  function bodies. ESLint is set to `sourceType: script` with `no-unused-vars` limited to local variables so
  cross-file globals aren't flagged.
- `responseData()` reads the `Form Responses 1` sheet (display values) and returns one object per response, keyed by
  the field names in `FIELDS` (e.g. `row.scout`, `row.patrol`, `row.seats`). Columns are found by matching row 1
  headers against each field's alias list (trimmed, case-insensitive), so column moves and added questions are
  harmless. When the form rewords a question, add the new wording as another alias; when a generator needs a new
  column, add a field. A missing header yields `''` and a `Logger.log` line rather than an error, and when a header
  appears twice the leftmost column wins. The exception is `row.signers`, which collects every "Permission" column
  (`SIGNATURE_HEADERS`) since that question can repeat.
- `createNamedSheet(name)` renames any existing sheet of that name to the next of `"<name> (old)"`,
  `"<name> (old 2)"`, ... after the highest in use, and inserts a fresh one. Running a generator repeatedly keeps every
  older copy (the higher the number, the newer, even if you delete one in between); it never overwrites or deletes.
- Domain constants at the top: `ADULTS` (patrol values meaning an adult is camping, matched against the form's
  option text) and `DRIVERS` (maps form option text to a short label used in the driving sheet).
- Driving sheet: columns are camping adults plus parents who chose a driving option AND supplied a seat count;
  rows are scouts. Seat count is shown as `parseInt(count) + 1`: the form asks for passenger seats not counting the
  driver's own scout.
