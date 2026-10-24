// Roles come from the free-form PoResponsibility column of the Patrols sheet, so aliases are lower-case with dots
// removed; see findRole. The lowest rank prints first within a group, and the name is how the scout-facing roster
// spells the role out.
const ROLES: {aliases: string[]; rank: number; name: string}[] = [
  {aliases: ['scoutmaster'], rank: 0, name: 'Scoutmaster'},
  {aliases: ['senior patrol leader', 'spl'], rank: 0, name: 'Senior Patrol Leader'},
  {
    aliases: ['assistant senior patrol leader', 'aspl'],
    rank: 1,
    name: 'Assistant Senior Patrol Leader',
  },
  {aliases: ['patrol leader', 'pl'], rank: 2, name: 'Patrol Leader'},
  {aliases: ['assistant patrol leader', 'apl'], rank: 3, name: 'Assistant Patrol Leader'},
  {aliases: ['quartermaster', 'qm'], rank: 4, name: 'Quartermaster'},
]
const UNRANKED = 9

// Answers that mean "nothing to report" for the medical fields, compared lower-case with dots and spaces removed
// (so "N.A." and "n a" both match 'na').
const NONE_ANSWERS = ['', 'none', 'n/a', 'na', 'no', 'nope', 'nothing', '-']

type RosterPerson = {
  name: string
  isAdult: boolean
  // The patrol the scout signed up under, versus the campout patrol a leader assigned on the Patrols sheet.
  patrol: string
  campoutPatrol: string
  role: string
  primaryName: string
  primaryPhone: string
  secondaryName: string
  secondaryPhone: string
  physician: string
  physicianPhone: string
  allergies: string
  dietary: string
  medications: string
  notes: string
  // Just what a leader entered on the Patrols sheet, for printouts that leave out the raw form answers.
  patrolNotes: string
  // Things a human should double-check before the roster goes in the binder.
  issues: string[]
}

// roles lists the group's PoRs, highest rank first, e.g. "PL Sam Doe".
type RosterGroup = {title: string; roles: string[]; people: RosterPerson[]}

function cleanText(text: string | undefined): string {
  return (text ?? '').replace(/\s+/g, ' ').trim()
}

// Forms often get filled in with caps lock on; leave mixed-case names (de Vries, DeLuca) alone.
function niceName(text: string): string {
  const name = cleanText(text)
  const isAllCaps = name === name.toUpperCase() && name !== name.toLowerCase()
  // Capitalize whatever follows a separator rather than using \b\w, which treats accented letters as boundaries.
  return isAllCaps
    ? name
        .toLowerCase()
        .replace(/(^|[\s'-])(\S)/g, (_, separator, letter) => separator + letter.toUpperCase())
    : name
}

// Words that begin a multi-word surname ("Van Pelt", "de Vries"), compared lower-case.
const SURNAME_PARTICLES = [
  'van',
  'von',
  'de',
  'del',
  'della',
  'di',
  'da',
  'du',
  'la',
  'le',
  'st',
  'san',
  'ter',
  'ten',
  'bin',
]

// The final word is the surname, extended left over any particles ("Ben Van Pelt" -> "van pelt"). Other
// compound surnames ("Garcia Lopez") can't be told from a middle name, so they sort by the final word.
function lastNameKey(name: string): string {
  const words = cleanText(name).toLowerCase().split(' ')
  let start = words.length - 1
  while (start > 0 && SURNAME_PARTICLES.includes(words[start - 1])) {
    start--
  }
  return words.slice(start).join(' ')
}

// The full name breaks ties.
function compareByLastName(a: string, b: string): number {
  return lastNameKey(a).localeCompare(lastNameKey(b)) || a.localeCompare(b)
}

function lastTenDigits(text: string): string {
  return text.replace(/\D/g, '').slice(-10)
}

// Odd values are left alone rather than guessed at.
function formatPhone(text: string): string {
  const phone = cleanText(text)
  let digits = phone.replace(/\D/g, '')
  if (digits.length === 11 && digits.startsWith('1')) {
    digits = digits.slice(1)
  }
  return digits.length === 10
    ? `(${digits.slice(0, 3)}) ${digits.slice(3, 6)}-${digits.slice(6)}`
    : phone
}

function emptyIfNone(text: string | undefined): string {
  const answer = cleanText(text)
  return NONE_ANSWERS.includes(answer.toLowerCase().replace(/[.\s]/g, '')) ? '' : answer
}

function rosterIssues(person: RosterPerson): string[] {
  const issues: string[] = []
  const physicianName = person.physician.toLowerCase()
  const primaryName = person.primaryName.toLowerCase()

  if (!person.physician) {
    issues.push('no physician listed')
  }
  if (lastTenDigits(person.physicianPhone).length < 10) {
    issues.push('physician phone missing or invalid')
  }
  if (!person.primaryName || lastTenDigits(person.primaryPhone).length < 10) {
    issues.push('primary contact name or phone missing')
  }
  // "Dr" as a title, not the start of a name like Drew.
  if (/^dr(\.|\s|$)/.test(primaryName) || (physicianName && primaryName === physicianName)) {
    issues.push('primary contact looks like the physician')
  }
  const familyPhones = [person.primaryPhone, person.secondaryPhone].map(lastTenDigits)
  const physicianPhone = lastTenDigits(person.physicianPhone)
  if (physicianPhone && familyPhones.includes(physicianPhone)) {
    issues.push('physician phone matches a family contact')
  }

  return issues
}

type PatrolExtra = {patrol: string; campoutPatrol: string; role: string; notes: string}

// Role and notes a leader has entered on the Patrols sheet, keyed by lower-cased scout name. A name can appear more
// than once (a parent and child, say), so each key holds a list; see findPatrolExtra. The sheet is optional:
// without it the roster simply has no role badges.
function patrolExtras(): Map<string, PatrolExtra[]> {
  const extras = new Map<string, PatrolExtra[]>()
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Patrols')
  if (sheet == null) {
    return extras
  }

  const [headers, ...rows] = sheet.getDataRange().getDisplayValues()
  const normalizedHeaders = headers.map(normalizeHeader)
  const [scout, patrol, campoutPatrol, role, notes] = [
    'Scout',
    'Patrol',
    'Campout Patrol',
    'PoResponsibility',
    'Notes',
  ].map(header => firstColumn(normalizedHeaders, [header], `Patrols sheet: ${header}`))
  const cell = (row: string[], column: number) => (column === -1 ? '' : row[column])

  for (const row of rows) {
    const key = cleanText(cell(row, scout)).toLowerCase()
    extras.set(key, [
      ...(extras.get(key) ?? []),
      {
        patrol: cleanText(cell(row, patrol)).toLowerCase(),
        campoutPatrol: cleanText(cell(row, campoutPatrol)),
        role: cleanText(cell(row, role)),
        notes: emptyIfNone(cell(row, notes)),
      },
    ])
  }
  return extras
}

// When two people share a name, prefer the entry for the patrol they signed up under.
function findPatrolExtra(extras: Map<string, PatrolExtra[]>, row: ResponseRow): PatrolExtra {
  const candidates = extras.get(cleanText(row.scout).toLowerCase()) ?? []
  const patrol = cleanText(row.patrol).toLowerCase()
  return (
    candidates.find(candidate => candidate.patrol === patrol) ??
    candidates[0] ?? {patrol, campoutPatrol: '', role: '', notes: ''}
  )
}

// createPatrolListing copies the form's notes onto the Patrols sheet, where leaders may extend them. A sheet note
// that already contains the form's is the fuller copy, so print it alone rather than repeating the form text.
function combineNotes(formNotes: string, patrolNotes: string): string {
  if (!patrolNotes) return formNotes
  if (!formNotes || patrolNotes.toLowerCase().includes(formNotes.toLowerCase())) return patrolNotes
  return `${formNotes} ${patrolNotes}`
}

function rosterPerson(row: ResponseRow, extras: Map<string, PatrolExtra[]>): RosterPerson {
  const primaryName = niceName(row.primaryName)
  const secondaryName = niceName(row.secondaryName)
  const extra = findPatrolExtra(extras, row)
  const isAdult = ADULTS.includes(row.patrol)

  // Some families typed a parent's name (a signature) into Medications.
  const parentNames = [primaryName, secondaryName, ...row.signers.map(niceName)].map(name =>
    name.toLowerCase(),
  )
  let medications = emptyIfNone(row.medications)
  if (medications && parentNames.includes(niceName(medications).toLowerCase())) {
    medications = ''
  }

  const notes = combineNotes(emptyIfNone(row.notes), extra.notes)

  const person: RosterPerson = {
    name: niceName(row.scout),
    isAdult,
    patrol: isAdult ? '' : cleanText(row.patrol),
    campoutPatrol: extra.campoutPatrol,
    role: extra.role,
    primaryName,
    primaryPhone: formatPhone(row.primaryPhone),
    secondaryName,
    secondaryPhone: formatPhone(row.secondaryPhone),
    physician: cleanText(row.physician),
    physicianPhone: formatPhone(row.physicianPhone),
    allergies: emptyIfNone(row.allergies),
    dietary: emptyIfNone(row.dietary),
    medications,
    notes,
    patrolNotes: extra.notes,
    issues: [],
  }
  person.issues = rosterIssues(person)
  return person
}

function findRole(role: string) {
  const alias = role.toLowerCase().replace(/\./g, '')
  return ROLES.find(known => known.aliases.includes(alias))
}

function roleRank(role: string): number {
  return findRole(role)?.rank ?? UNRANKED
}

function byRoleThenLastName(a: RosterPerson, b: RosterPerson): number {
  return roleRank(a.role) - roleRank(b.role) || compareByLastName(a.name, b.name)
}

const ADULT_CAMPOUT_PATROL = 'a'

// A purely numeric campout patrol ("2") reads better as "Patrol 2"; names ("Eagles") are used as typed.
function campoutGroupTitle(campoutPatrol: string): string {
  if (campoutPatrol === '') return 'No campout patrol'
  return /^\d+$/.test(campoutPatrol) ? `Patrol ${campoutPatrol}` : campoutPatrol
}

// People are expected to be sorted already, so the roles come out highest rank first.
function rosterGroup(title: string, people: RosterPerson[]): RosterGroup {
  const roles = people.filter(person => person.role).map(person => `${person.role} ${person.name}`)
  return {title, roles, people}
}

// Campout patrols only apply once a leader has assigned some. While the Patrols sheet is missing or holds nothing
// but the adults' 'A', scouts stay in the patrols they signed up under.
function hasCampoutAssignments(scouts: RosterPerson[]): boolean {
  return scouts.some(
    scout =>
      scout.campoutPatrol !== '' && scout.campoutPatrol.toLowerCase() !== ADULT_CAMPOUT_PATROL,
  )
}

// Adult leaders first, then one group per campout patrol (or per signed-up patrol, see hasCampoutAssignments) in
// text order, ignoring case so "eagles" and "Eagles" are one group. Unlike the Patrols sheet, which puts unassigned
// scouts first to show the work left, a printout puts them last.
function rosterGroups(): RosterGroup[] {
  const extras = patrolExtras()
  const people = responseData()
    .filter(row => cleanText(row.scout) !== '')
    .map(row => rosterPerson(row, extras))

  const groups: RosterGroup[] = []
  const adults = people.filter(person => person.isAdult)
  if (adults.length > 0) {
    groups.push(rosterGroup('Adult Leaders', adults.sort(byRoleThenLastName)))
  }

  const scouts = people.filter(person => !person.isAdult)
  const byCampout = hasCampoutAssignments(scouts)
  const patrolOf = (scout: RosterPerson) => (byCampout ? scout.campoutPatrol : scout.patrol)
  const groupKey = (scout: RosterPerson) => patrolOf(scout).toLowerCase()
  const titleOf = (patrol: string) =>
    byCampout ? campoutGroupTitle(patrol) : patrol || 'No patrol'
  const keys = [...new Set(scouts.map(groupKey))].sort((a, b) =>
    a === '' || b === '' ? Number(a === '') - Number(b === '') : a.localeCompare(b),
  )
  for (const key of keys) {
    const members = scouts.filter(scout => groupKey(scout) === key)
    // The group is titled with the first spelling found.
    groups.push(rosterGroup(titleOf(patrolOf(members[0])), members.sort(byRoleThenLastName)))
  }

  return groups
}

const ROSTER_SETTINGS_KEY = 'rosterSettings'

type RosterSettings = HeaderSettings & {lead: string; showFlags: boolean}

// Spelled out rather than spread from DEFAULT_HEADER_SETTINGS: a top-level statement can't use another file's constant.
const DEFAULT_ROSTER_SETTINGS: RosterSettings = {
  title: '',
  subtitle: '',
  lead: '',
  showFlags: true,
}

function showRoster() {
  const dialog = HtmlService.createHtmlOutputFromFile('RosterDialog').setWidth(1000).setHeight(760)
  SpreadsheetApp.getUi().showModalDialog(dialog, 'Printable roster')
}

// Called from RosterDialog.html via google.script.run.
function getRosterPage(): {groups: RosterGroup[]; settings: RosterSettings} {
  return {
    groups: rosterGroups(),
    settings: readSettings(ROSTER_SETTINGS_KEY, DEFAULT_ROSTER_SETTINGS),
  }
}

// Called from RosterDialog.html via google.script.run.
function saveRosterSettings(settings: RosterSettings) {
  writeSettings(ROSTER_SETTINGS_KEY, settings)
}
