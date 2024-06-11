const ADULTS: string[] = ['Adult Leader', 'Scouter Reserve']

const DRIVERS: {[id: string]: string} = {
  'Drive both ways': 'both',
  'Drive to the campout only': 'to',
  'Drive from the campout only': 'from',
  'Tow both ways': 'tow',
  'Tow to the campout only': 'tow to',
  'Tow from the campout only': 'tow from',
}

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

function responseData(): string[][] {
  const formSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Responses 1')
  if (formSheet == null) {
    throw new Error('Form Responses 1 sheet not found')
  }

  const rd = formSheet.getDataRange().getDisplayValues()
  // assumes row 1 contains column headers which our fucnctions don't need, and we'll ignore them by shifting
  rd.shift()

  return rd
}

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function createPatrolListing() {
  const outputSheet = createNamedSheet('Patrols')
  const rd = responseData()

  const listing: (string | null)[][] = [
    ['Scout', 'Patrol', 'Campout Patrol', 'PoResponsibility', 'Dietary Restrictions'],
    ...rd.map(row => {
      const campout_patrol = ADULTS.includes(row[3]) ? 'A' : null
      return [row[2], row[3], campout_patrol, null, row[13]]
    }),
  ]

  const size = rd.length + 1
  outputSheet.getRange(`A1:E${size}`).setValues(listing)
  outputSheet.getRange('A1:E1').setFontWeight('bold').setHorizontalAlignment('center')
  outputSheet.setFrozenRows(1)
  outputSheet.autoResizeColumns(1, 5)
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
      paddedArray(['Scout', row[2], row[3]]),
      paddedArray(['Allergies', row[12]]),
      paddedArray(['Dietary Restrictions', row[13]]),
      Array(4),
      paddedArray(['Primary Contact', row[6], row[7]]),
      paddedArray(['Secondary Contact', row[8], row[9]]),
      Array(4),
      paddedArray(["Physician's Name", row[10], 'Phone', row[11]]),
      Array(4),
      paddedArray(['Additional Notes', row[19]]),
      Array(4),
      paddedArray(['--------------------']),
    ]

    outputSheet.getRange(`A${base}:D${base + 11}`).setValues(rosterEntry)
  })
}

function drivingText(a: string[], name: string): string {
  const how = DRIVERS[a[15]]
  const num = parseInt(a[17]) + 1

  return how == 'both' ? `${name} (${num})` : `${name} (${how}: ${num})`
}

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function createDrivingAssignments() {
  const outputSheet = createNamedSheet('Driving')
  const rd = responseData()

  const adultsCamping = rd.filter(a => ADULTS.includes(a[3]))
  const scouts = rd.filter(a => !ADULTS.includes(a[3]))
  const parentDrivers = scouts.filter(a => Object.keys(DRIVERS).includes(a[15]))

  const drivers: (string | null)[] = [
    null,
    ...adultsCamping.map(a => drivingText(a, a[2])),
    ...parentDrivers.map(a => drivingText(a, a[16] || `${a[2]}!!`)),
  ]

  const columns = drivers.length
  const drivingData: (string | null)[][] = [
    drivers,
    ...scouts.map(s => paddedArray([s[2]], drivers.length)),
  ]

  const rows = scouts.length + 1
  outputSheet.getRange(1, 1, rows, columns).setValues(drivingData)
  outputSheet.autoResizeColumns(1, columns)
  outputSheet.getRange(2, 2, rows - 1, columns - 1).setHorizontalAlignment('center')
}
