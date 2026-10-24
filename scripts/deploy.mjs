#!/usr/bin/env node
// Builds and pushes the script to a Google Sheet's bound Apps Script project.
//
//   npm run deploy -- <sheet URL | sheet ID | saved name>
//   npm run deploy -- --add <sheet URL | sheet ID> <script ID> [--name <label>]
//   npm run deploy -- --list
//   (any of the above accepts --dry-run to print the commands instead of running them)
//
// A sheet can have only one bound script, and Google has no API to look up a sheet's existing script ID, so
// targets.json (gitignored) remembers the sheet -> script ID mapping. Unknown sheets get a new bound script.
import {execFileSync} from 'node:child_process'
import {existsSync, readFileSync, writeFileSync} from 'node:fs'
import {dirname, join} from 'node:path'
import {fileURLToPath} from 'node:url'

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..')
const TARGETS_FILE = process.env.TARGETS_FILE || join(ROOT, 'targets.json')
const CLASP_FILE = join(ROOT, '.clasp.json')
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
  console.log(`writing .clasp.json for script ${scriptId}`)
  if (!dryRun)
    writeFileSync(CLASP_FILE, JSON.stringify({scriptId, rootDir: 'dist'}, null, 2) + '\n')
}

function resolveTarget(input, targets) {
  const sheetId = sheetIdFrom(input)
  if (sheetId) return {sheetId, entry: targets[sheetId]}

  const match = Object.entries(targets).find(([, entry]) => entry.name === input)
  if (!match)
    fail(`"${input}" is not a sheet URL/ID or a saved name. Saved: ${savedNames(targets)}`)
  return {sheetId: match[0], entry: match[1]}
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

function deploy(input, targets) {
  const {sheetId, entry} = resolveTarget(input, targets)
  run('npm', ['run', 'build'])

  if (entry) {
    writeClaspFile(entry.scriptId)
  } else {
    console.log(`No script recorded for sheet ${sheetId}; creating a bound script.`)
    console.log(
      'If the sheet already has one (pasted in by hand), creation fails: find its script ID under',
    )
    console.log(
      'Extensions > Apps Script > Project settings and run: npm run deploy -- --add <sheet> <script ID>',
    )
    run('npx', [
      'clasp',
      'create',
      '--type',
      'sheets',
      '--parentId',
      sheetId,
      '--title',
      name || 'Troop 66 RSVP helper',
      '--rootDir',
      'dist',
    ])
    if (!dryRun) {
      targets[sheetId] = {
        name: name || sheetId,
        scriptId: JSON.parse(readFileSync(CLASP_FILE, 'utf8')).scriptId,
      }
      saveTargets(targets)
    }
  }

  // -f because the pushed manifest (appsscript.json) should replace whatever the sheet's script currently has.
  run('npx', ['clasp', 'push', '-f'])
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
