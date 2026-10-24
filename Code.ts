const ADULTS: string[] = ['Adult Leader (ASM or Committee)', 'Scouter Reserve (Adult)']

const DRIVERS: {[id: string]: string} = {
  'Drive both ways': 'both',
  'Drive to the campout only': 'to',
  'Drive from the campout only': 'from',
  'Tow both ways': 'tow',
  'Tow to the campout only': 'tow to',
  'Tow from the campout only': 'tow from',
}

// Form column headers each field may appear under. The form's wording has changed between outings, and
// columns have been added, removed, and moved, so we find columns by header rather than by position.
const FIELDS = {
  scout: [
    'Name of Scout or Adult Leader attending',
    'Name of Scout, Adult, or Adult Leader attending',
  ],
  patrol: ['Patrol Name'],
  primaryName: ['Primary Emergency Contact Name'],
  primaryPhone: ['Primary Emergency Contact Phone Number'],
  secondaryName: ['Secondary Emergency Contact Name'],
  secondaryPhone: ['Secondary Emergency Contact Phone Number'],
  physician: ["Physician's Name"],
  physicianPhone: ["Physician's Phone Number"],
  allergies: ['Allergies'],
  dietary: ['Dietary Restrictions'],
  driving: ['Driving'],
  driverName: ['Driver Name'],
  seats: ['Seats'],
  notes: ['Notes'],
}

type Field = keyof typeof FIELDS
type ResponseRow = Record<Field, string>

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function onOpen() {
  createMenu()
}

function createMenu() {
  SpreadsheetApp.getUi()
    .createMenu('Troop 66')
    .addItem('create patrol sheet', 'createPatrolListing')
    .addItem('create driving sheet', 'createDrivingAssignments')
    .addItem('create roster', 'createOutingListing')
    .addToUi()
}

function createNamedSheet(name: string): GoogleAppsScript.Spreadsheet.Sheet {
  const activeSheet = SpreadsheetApp.getActiveSpreadsheet()
  let output = activeSheet.getSheetByName(name)
  if (output != null) {
    output.setName(`${name} (old)`)
  }

  output = activeSheet.insertSheet()
  if (output != null) {
    output.setName(name)
  }

  return output
}

function normalizeHeader(header: string): string {
  return header.trim().replace(/\s+/g, ' ').toLowerCase()
}

// Maps each field to the index of its column, or -1 when the form has no matching header. When a header appears more
// than once (e.g. a question re-added to the form), the leftmost column wins.
function fieldColumns(headers: string[]): Record<Field, number> {
  const normalizedHeaders = headers.map(normalizeHeader)
  const columns = {} as Record<Field, number>

  for (const field of Object.keys(FIELDS) as Field[]) {
    const aliases = FIELDS[field].map(normalizeHeader)
    columns[field] = normalizedHeaders.findIndex(header => aliases.includes(header))
    if (columns[field] === -1) {
      Logger.log(`No column found for "${field}"; looked for: ${FIELDS[field].join(' | ')}`)
    }
  }

  return columns
}

function responseData(): ResponseRow[] {
  const formSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Responses 1')
  if (formSheet == null) {
    throw new Error('Form Responses 1 sheet not found')
  }

  const [headers, ...rows] = formSheet.getDataRange().getDisplayValues()
  const columns = fieldColumns(headers)
  const fields = Object.keys(FIELDS) as Field[]

  return rows.map(row => {
    const entry = {} as ResponseRow
    for (const field of fields) {
      entry[field] = columns[field] === -1 ? '' : row[columns[field]]
    }
    return entry
  })
}

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function createPatrolListing() {
  const outputSheet = createNamedSheet('Patrols')
  const rd = responseData()

  const listing: (string | null)[][] = [
    ['Scout', 'Patrol', 'Campout Patrol', 'PoResponsibility', 'Dietary Restrictions', 'Notes'],
    ...rd.map(row => {
      const campout_patrol = ADULTS.includes(row.patrol) ? 'A' : null
      return [row.scout, row.patrol, campout_patrol, null, row.dietary, row.notes]
    }),
  ]

  const size = rd.length + 1
  outputSheet.getRange(`A1:F${size}`).setValues(listing)
  outputSheet.getRange('A1:F1').setFontWeight('bold').setHorizontalAlignment('center')
  outputSheet.setFrozenRows(1)
  outputSheet.autoResizeColumns(1, 6)
  outputSheet.getRange(2, 3, outputSheet.getLastRow()).setHorizontalAlignment('right')
}

function paddedArray(array: string[], len: number = 4): string[] {
  return Array.from({...array, length: len})
}

// eslint-disable-next-line @typescript-eslint/no-unused-vars
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

function drivingText(a: ResponseRow, name: string): string {
  const how = DRIVERS[a.driving]
  const num = parseInt(a.seats) + 1

  return how == 'both' ? `${name} (${num})` : `${name} (${how}: ${num})`
}

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function createDrivingAssignments() {
  const outputSheet = createNamedSheet('Driving')
  const rd = responseData()

  const adultsCamping = rd.filter(a => ADULTS.includes(a.patrol))
  const scouts = rd.filter(a => !ADULTS.includes(a.patrol))
  const parentDrivers = scouts.filter(
    a => Object.keys(DRIVERS).includes(a.driving) && a.seats !== '',
  )

  const drivers: (string | null)[] = [
    null,
    ...adultsCamping.map(a => drivingText(a, a.scout)),
    ...parentDrivers.map(a => drivingText(a, a.driverName || `${a.scout}!!`)),
  ]

  const columns = drivers.length
  const drivingData: (string | null)[][] = [
    drivers,
    ...scouts.map(s => paddedArray([s.scout], drivers.length)),
  ]

  const rows = scouts.length + 1
  outputSheet.getRange(1, 1, rows, columns).setValues(drivingData)
  outputSheet.autoResizeColumns(1, columns)
  outputSheet.getRange(2, 2, rows - 1, columns - 1).setHorizontalAlignment('center')
}
