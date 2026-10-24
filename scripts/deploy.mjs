#!/usr/bin/env node
// Builds and pushes the script to a Google Sheet's bound Apps Script project.
//
//   npm run deploy -- <sheet URL | sheet ID | saved name>
//   npm run deploy -- --add <sheet URL | sheet ID> <script ID> [--name <label>]
//   npm run deploy -- --list
//   (any of the above accepts --dry-run to print the commands instead of running them)
//
// A sheet can have only one bound script, and Google has no API to look up a sheet's existing script ID, so
// targets.json (gitignored) remembers the sheet -> script ID mapping. Record each sheet once with --add; opening
// Extensions > Apps Script in a sheet creates its bound script if it has none.
import {execFileSync} from 'node:child_process'
import {existsSync, readdirSync, readFileSync, writeFileSync} from 'node:fs'
import {dirname, join} from 'node:path'
import {fileURLToPath} from 'node:url'

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..')
const TARGETS_FILE = process.env.TARGETS_FILE || join(ROOT, 'targets.json')
// Separate from .clasp.json, which `npm run push` uses, so deploying to a sheet never retargets plain pushes.
const DEPLOY_CLASP_FILE = '.clasp.deploy.json'
const ID_PATTERN = /^[\w-]{20,}$/

const args = process.argv.slice(2)
const dryRun = take(args, '--dry-run')
const listOnly = take(args, '--list')
const addMode = take(args, '--add')
const name = takeValue(args, '--name')

function take(list, flag) {
  const i = list.indexOf(flag)
  if (i !== -1) list.splice(i, 1)
  return i !== -1
}

function takeValue(list, flag) {
  const i = list.indexOf(flag)
  if (i === -1) return undefined
  const [, value] = list.splice(i, 2)
  return value
}

function fail(message) {
  console.error(message)
  process.exit(1)
}

function loadTargets() {
  return existsSync(TARGETS_FILE) ? JSON.parse(readFileSync(TARGETS_FILE, 'utf8')) : {}
}

function saveTargets(targets) {
  writeFileSync(TARGETS_FILE, JSON.stringify(targets, null, 2) + '\n')
}

function sheetIdFrom(input) {
  const fromUrl = input.match(/\/spreadsheets\/d\/([\w-]+)/)
  if (fromUrl) return fromUrl[1]
  return ID_PATTERN.test(input) ? input : undefined
}

function run(command, commandArgs) {
  console.log(`$ ${command} ${commandArgs.join(' ')}`)
  if (!dryRun) execFileSync(command, commandArgs, {cwd: ROOT, stdio: 'inherit'})
}

function writeClaspFile(scriptId) {
  console.log(`writing ${DEPLOY_CLASP_FILE} for script ${scriptId}`)
  if (!dryRun)
    writeFileSync(
      join(ROOT, DEPLOY_CLASP_FILE),
      JSON.stringify({scriptId, rootDir: 'dist'}, null, 2) + '\n',
    )
}

// Saved names win over ID parsing, so a long name can't be mistaken for a sheet ID.
function resolveTarget(input, targets) {
  const named = Object.entries(targets).find(([, entry]) => entry.name === input)
  if (named) return {sheetId: named[0], entry: named[1]}

  const sheetId = sheetIdFrom(input)
  if (!sheetId)
    fail(`"${input}" is not a sheet URL/ID or a saved name. Saved: ${savedNames(targets)}`)
  return {sheetId, entry: targets[sheetId]}
}

function savedNames(targets) {
  return (
    Object.values(targets)
      .map(entry => entry.name)
      .join(', ') || '(none)'
  )
}

function add(input, scriptId, targets) {
  const sheetId = sheetIdFrom(input ?? '')
  if (!sheetId || !ID_PATTERN.test(scriptId ?? ''))
    fail('usage: --add <sheet URL|ID> <script ID> [--name <label>]')
  targets[sheetId] = {name: name || sheetId, scriptId}
  if (!dryRun) saveTargets(targets)
  console.log(`saved ${targets[sheetId].name} -> script ${scriptId}`)
}

// Apps Script keeps one namespace for file names regardless of type, so Roster.ts (-> Roster.js) and Roster.html
// can't be pushed together. Checking the sources catches that before any network call.
function assertUniqueFileNames() {
  const seen = new Map()
  for (const file of readdirSync(join(ROOT, 'src')).filter(file => /\.(ts|html)$/.test(file))) {
    const base = file.replace(/\.\w+$/, '')
    if (seen.has(base)) {
      fail(
        `Apps Script can't hold both ${seen.get(base)} and ${file}: rename one so the names differ.`,
      )
    }
    seen.set(base, file)
  }
}

function assertLoggedIn() {
  let report
  try {
    report = execFileSync('npx', ['clasp', 'show-authorized-user'], {cwd: ROOT, encoding: 'utf8'})
  } catch {
    report = 'not logged in'
  }
  if (/not logged in/i.test(report)) {
    fail(
      'clasp is not logged in. Run `npx clasp login`, and enable the Apps Script API at\n' +
        'https://script.google.com/home/usersettings, then retry.',
    )
  }
}

function deploy(input, targets) {
  const {sheetId, entry} = resolveTarget(input, targets)
  if (!entry) {
    fail(
      `No script recorded for sheet ${sheetId}. In the sheet, open Extensions > Apps Script (this creates the\n` +
        'bound script if there is none), copy the Script ID from Project settings, and run:\n' +
        '  npm run deploy -- --add <sheet URL or ID> <script ID> --name <label>',
    )
  }

  assertUniqueFileNames()
  if (!dryRun) assertLoggedIn()
  run('npm', ['run', 'build'])
  writeClaspFile(entry.scriptId)
  // -f because the pushed manifest (appsscript.json) should replace whatever the sheet's script currently has.
  run('npx', ['clasp', '-P', DEPLOY_CLASP_FILE, 'push', '-f'])
}

const targets = loadTargets()
if (listOnly) {
  for (const [sheetId, entry] of Object.entries(targets))
    console.log(`${entry.name}\t${sheetId}\t${entry.scriptId}`)
} else if (addMode) {
  add(args[0], args[1], targets)
} else if (args[0]) {
  deploy(args[0], targets)
} else {
  fail(
    'usage: npm run deploy -- <sheet URL | sheet ID | saved name>   (see the header of scripts/deploy.mjs)',
  )
}
