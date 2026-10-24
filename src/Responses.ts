const ADULTS: string[] = ['Adult Leader (ASM or Committee)', 'Scouter Reserve (Adult)']

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
  medications: ['Medications'],
}

// Parents sign the form in a "Permission" question that can appear in several columns, so unlike FIELDS every
// matching column is read.
const SIGNATURE_HEADERS = ['Permission']

type Field = keyof typeof FIELDS
type ResponseRow = Record<Field, string> & {signers: string[]}

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

function allColumns(normalizedHeaders: string[], aliases: string[]): number[] {
  const normalizedAliases = aliases.map(normalizeHeader)
  return normalizedHeaders.flatMap((header, index) =>
    normalizedAliases.includes(header) ? [index] : [],
  )
}

// Index of the leftmost column whose header matches an alias, or -1. A miss is logged so a reworded header gets noticed
// rather than silently producing blanks.
function firstColumn(normalizedHeaders: string[], aliases: string[], label: string): number {
  const index = allColumns(normalizedHeaders, aliases)[0] ?? -1
  if (index === -1) {
    Logger.log(`No column found for "${label}"; looked for: ${aliases.join(' | ')}`)
  }
  return index
}

// Maps each field to the index of its column, or -1 when the form has no matching header. When a header appears more
// than once (e.g. a question re-added to the form), the leftmost column wins.
function fieldColumns(headers: string[]): Record<Field, number> {
  const normalizedHeaders = headers.map(normalizeHeader)
  const columns = {} as Record<Field, number>

  for (const field of Object.keys(FIELDS) as Field[]) {
    columns[field] = firstColumn(normalizedHeaders, FIELDS[field], field)
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
  const signerColumns = allColumns(headers.map(normalizeHeader), SIGNATURE_HEADERS)
  const fields = Object.keys(FIELDS) as Field[]

  return rows.map(row => {
    const entry = {signers: signerColumns.map(column => row[column])} as ResponseRow
    for (const field of fields) {
      entry[field] = columns[field] === -1 ? '' : row[columns[field]]
    }
    return entry
  })
}
