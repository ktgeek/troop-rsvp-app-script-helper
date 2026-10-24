// Lowest rank prints first within a group. Roles come from the free-form PoResponsibility column of the Patrols
// sheet, so keys are lower-case with dots removed; see roleRank.
const ROLE_RANK: {[role: string]: number} = {
  scoutmaster: 0,
  'senior patrol leader': 0,
  spl: 0,
  'assistant senior patrol leader': 1,
  aspl: 1,
  'patrol leader': 2,
  pl: 2,
  quartermaster: 3,
  qm: 3,
}
const UNRANKED = 9

// Answers that mean "nothing to report" for the medical fields, compared lower-case with dots and spaces removed
// (so "N.A." and "n a" both match 'na').
const NONE_ANSWERS = ['', 'none', 'n/a', 'na', 'no', 'nope', 'nothing', '-']

type RosterPerson = {
  name: string
  isAdult: boolean
  patrol: string
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
  // Things a human should double-check before the roster goes in the binder.
  issues: string[]
}

type RosterGroup = {title: string; people: RosterPerson[]}

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

type PatrolExtra = {patrol: string; role: string; notes: string}

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
  const [scout, patrol, role, notes] = ['Scout', 'Patrol', 'PoResponsibility', 'Notes'].map(
    header => firstColumn(normalizedHeaders, [header], `Patrols sheet: ${header}`),
  )
  const cell = (row: string[], column: number) => (column === -1 ? '' : row[column])

  for (const row of rows) {
    const key = cleanText(cell(row, scout)).toLowerCase()
    extras.set(key, [
      ...(extras.get(key) ?? []),
      {
        patrol: cleanText(cell(row, patrol)).toLowerCase(),
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
    candidates[0] ?? {patrol, role: '', notes: ''}
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
    issues: [],
  }
  person.issues = rosterIssues(person)
  return person
}

function roleRank(role: string): number {
  return ROLE_RANK[role.toLowerCase().replace(/\./g, '')] ?? UNRANKED
}

function byRoleThenLastName(a: RosterPerson, b: RosterPerson): number {
  const lastName = (person: RosterPerson) => person.name.split(' ').pop()?.toLowerCase() ?? ''
  return (
    roleRank(a.role) - roleRank(b.role) ||
    lastName(a).localeCompare(lastName(b)) ||
    a.name.localeCompare(b.name)
  )
}

// Adult leaders first, then one group per patrol, alphabetically.
function rosterGroups(): RosterGroup[] {
  const extras = patrolExtras()
  const people = responseData()
    .filter(row => cleanText(row.scout) !== '')
    .map(row => rosterPerson(row, extras))

  const groups: RosterGroup[] = []
  const adults = people.filter(person => person.isAdult)
  if (adults.length > 0) {
    groups.push({title: 'Adult Leaders', people: adults.sort(byRoleThenLastName)})
  }

  const scouts = people.filter(person => !person.isAdult)
  const patrols = [...new Set(scouts.map(scout => scout.patrol))].sort()
  for (const patrol of patrols) {
    const members = scouts.filter(scout => scout.patrol === patrol)
    groups.push({title: patrol || 'No patrol', people: members.sort(byRoleThenLastName)})
  }

  return groups
}

function createOutingListing() {
  const outputSheet = createNamedSheet('Printable Roster')

  responseData().forEach((row, i) => {
    const base = 13 * i + 1

    const rosterEntry: string[][] = [
      paddedArray(['Scout', row.scout, row.patrol]),
      paddedArray(['Allergies', row.allergies]),
      paddedArray(['Dietary Restrictions', row.dietary]),
      Array(4),
      paddedArray(['Primary Contact', row.primaryName, row.primaryPhone]),
      paddedArray(['Secondary Contact', row.secondaryName, row.secondaryPhone]),
      Array(4),
      paddedArray(["Physician's Name", row.physician, 'Phone', row.physicianPhone]),
      Array(4),
      paddedArray(['Additional Notes', row.notes]),
      Array(4),
      paddedArray(['--------------------']),
    ]

    outputSheet.getRange(`A${base}:D${base + 11}`).setValues(rosterEntry)
  })
}
