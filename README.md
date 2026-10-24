# Troop RSVP app script helper

A Google Apps Script (TypeScript) that attaches to the Troop 66 RSVP spreadsheet and turns the Google Form responses
into working sheets. It adds a **Troop 66** menu:

| Menu item            | What it does                                                                     |
| -------------------- | -------------------------------------------------------------------------------- |
| create patrol sheet  | Writes a `Patrols` sheet to fill in patrol assignments, roles, and notes         |
| create driving sheet | Writes a `Driving` sheet: camping adults and parent drivers by scout, with seats |
| print roster         | Opens the printable roster, one card per person, to print or save as a PDF       |

The two sheet generators rename any existing sheet of the same name to `<name> (old)` instead of overwriting it.

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

Each deploy rewrites `.clasp.json` and pushes `appsscript.json`, so the target's timezone and runtime settings become
this repo's. The first time a sheet runs the script, Google asks you to authorize it.

## Printing the roster

1. Open **Troop 66 > print roster**.
2. Fill in the outing title, date, and lead line (they're remembered for next time).
3. Click **Print / Save as PDF**, choose **Save as PDF** as the destination, and turn on **Background graphics**.

Leave **Show VERIFY flags** on for a working copy: cards with missing or suspicious contact and physician details get
a yellow strip. Turn it off for the copy that goes in the binder. The footer and page numbers appear in the printout,
not the preview, and need a recent Chrome.

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
