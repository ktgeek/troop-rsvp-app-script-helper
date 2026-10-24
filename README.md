# Troop RSVP app script helper

A Google Apps Script (TypeScript) that attaches to the Troop 66 RSVP spreadsheet and turns the Google Form responses
into working sheets. It adds a **Troop 66** menu:

| Menu item            | What it does                                                                           |
| -------------------- | -------------------------------------------------------------------------------------- |
| create patrol sheet  | Writes a `Patrols` sheet to fill in patrol assignments, roles, and notes               |
| sort patrol sheet    | Re-sorts `Patrols` by Campout Patrol, then Patrol, then Scout's last name              |
| create driving sheet | Writes a `Driving` sheet: camping adults and parent drivers by scout, with seats       |
| print roster         | Opens the printable roster, one card per person, to print or save as a PDF             |
| print driving sheet  | Opens the printable car assignments, one card per car, from your marked Driving sheets |
| print patrol roster  | Opens a scout-facing roll call sheet by patrol: names, roles, dietary needs, notes     |

`Patrols` is sorted when created and again whenever you pick **sort patrol sheet**. Campout patrols sort as plain text,
so blanks come first, then `1`, `2`, `3`, ..., and `A` (adults) last. Sorting rewrites the data rows' values, so
cell colors or other formatting you add by hand won't follow the rows.

The two sheet generators keep any existing sheet of the same name as `<name> (old)`, then `(old 2)`, `(old 3)`, ...
(higher is newer) instead of overwriting it. The new `Patrols` sheet starts blank, so copy the roles and notes you had
entered over from the old one before printing the roster, which reads them from `Patrols`.

## Setup (once per machine)

1. `npm install`
2. `npx clasp login` and sign in with the Google account that owns (or can edit) the spreadsheets.
3. Turn on the **Google Apps Script API** at <https://script.google.com/home/usersettings>.

## Deploying to a sheet

```
npm run deploy -- <sheet URL | sheet ID | saved name>
```

This builds, then pushes to that sheet's bound Apps Script project. **Updating a sheet you already deployed to is the
same command**: the push replaces the project's files in place, and the menu always runs the latest push.

A sheet has only one bound script, and Google has no way to look up its ID, so the script keeps a
`targets.json` (gitignored; shape in `targets.example.json`) mapping sheets to script IDs.

- **First deploy to a sheet:** record the sheet's script ID once, then deploy as usual. A sheet with no script gets
  one the first time you open Extensions > Apps Script.
  1. In the sheet, open Extensions > Apps Script > Project settings and copy the **Script ID**.
  2. `npm run deploy -- --add <sheet URL or ID> <script ID> --name <label>`
  3. `npm run deploy -- <label>` (or the sheet's URL or ID)
- `npm run deploy -- --list` shows saved targets; add `--dry-run` to any deploy to print the commands without running
  them.

Each deploy writes its own `.clasp.deploy.json` (leaving `.clasp.json`, which `npm run push` uses, alone) and pushes
`appsscript.json`, so the target's timezone and runtime settings become
this repo's. The first time a sheet runs the script, Google asks you to authorize it.

## Printing the roster

1. Open **Troop 66 > print roster**.
2. Fill in the outing title, date, and lead line (they're remembered for next time).
3. Click **Print / Save as PDF**, choose **Save as PDF** as the destination, and turn on **Background graphics**.

The roster groups scouts by the **Campout Patrol** you entered on the `Patrols` sheet (Adult Leaders first, numbered
patrols shown as "Patrol 2"), with each group's roles from the **PoResponsibility** column in its heading. If no
campout patrols are filled in beyond `A`, or there is no `Patrols` sheet, it groups by the patrol each scout signed up
under.

Leave **Show VERIFY flags** on for a working copy: cards with missing or suspicious contact and physician details get
a yellow strip. Turn it off for the copy that goes in the binder. The footer and page numbers appear in the printout,
not the preview, and need a recent Chrome.

## Printing the patrol roster

**Troop 66 > print patrol roster** is the scout-facing counterpart: no contact or medical details, just each patrol's
members (adults included, under Adult Leaders) with a roll call checkbox, their PoR, allergies and dietary
restrictions, and the notes from the `Patrols` sheet. It uses the same grouping as the leader roster. Notes are taken
from the `Patrols` sheet only, so the wording you edit there is what scouts see. Print it the same way as the roster.

## Printing the driving sheet

Generate the `Driving` sheet, then assign scouts by hand: put any mark in a scout's row under the driver they ride
with. **Troop 66 > print driving sheet** turns that into one card per car (driver, seats, to/from/towing label, and
the scouts in it), with cars ordered by the driver's last name. A red box at the top of each page lists scouts with
no car yet or marked in more than one car, and a car with more scouts than seats (minus the driver) is flagged red.

If the trip needs different cars each way, duplicate the tab and rename it, e.g. `Driving to camp` and
`Driving home`. Every tab whose name starts with `Driving` is printed, in tab order, each on its own page; the
`Driving (old)` copies the generator sets aside are skipped.

## Development

| Command                   | Purpose                                                      |
| ------------------------- | ------------------------------------------------------------ |
| `npm run build`           | Compile `src/*.ts` to `dist/` and copy the HTML and manifest |
| `npm run lint`            | ESLint with Prettier                                         |
| `npm run prettier-format` | Rewrite `src/*.ts` in Prettier style                         |
| `npm run push`            | Build, then `clasp push` to whatever `.clasp.json` names     |

There is no test suite; the code needs the Apps Script runtime, so check changes by deploying to a scratch copy of the
sheet. See `CLAUDE.md` for the architecture.

## Troubleshooting

- **"clasp is not logged in"** (or a permissions error on push): run `npx clasp login`, and make sure the Apps Script
  API is enabled (see Setup).
- **"A file with this name already exists in the current project"**: Apps Script needs every file name unique
  regardless of type, so `Roster.ts` and `Roster.html` can't coexist. The deploy script checks `src/` for this before
  pushing. If it still happens, the target script has a leftover file of that name from an earlier paste: delete it in
  Extensions > Apps Script and deploy again.
- **Duplicate menu items or old behavior after deploying to a hand-pasted script:** it may still have an old `Code.gs`
  next to the new `Code.js`. Delete the old file in the Apps Script editor.
- **"No script recorded for sheet ...":** the sheet isn't in `targets.json` yet. Record it with `--add` as above.
