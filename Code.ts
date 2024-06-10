const ADULTS = ['Adult Leader', 'Scouter Reserve']

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
    activeSheet.deleteSheet(output)
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
  const responseData = formSheet.getDataRange().getDisplayValues()
  // assumes row 1 contains column headers which our fucnctions don't need, and we'll ignore them by shifting
  responseData.shift()

  return responseData
}

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function createPatrolListing() {
  const outputSheet = createNamedSheet('Patrols')
  const rd = responseData()

  const listing: (string | null)[][] = [
    ['Scout', 'Patrol', 'Campout Patrol', 'PoResponsibility', 'Dietary Restrictions'],
  ]
  rd.forEach(row => {
    const campout_patrol = ADULTS.includes(row[3]) ? 'A' : null
    const patrolRow = [row[2], row[3], campout_patrol, null, row[13]]
    listing.push(patrolRow)
  })

  const size = responseData.length + 1
  outputSheet.getRange(`A1:E${size}`).setValues(listing)
  outputSheet.getRange('A1:E1').setFontWeight('bold').setHorizontalAlignment('center')
  outputSheet.setFrozenRows(1)
  outputSheet.autoResizeColumns(1, 5)
  outputSheet.getRange(2, 3, outputSheet.getLastRow()).setHorizontalAlignment('right')
}

// pA is "paddedArray"
function paddedArray(array: string[], len: number = 4): string[] {
  return Array.from({...array, length: len})
}

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function createOutingListing() {
  const outputSheet = createNamedSheet('Printable Roster')

  responseData().forEach((row, i) => {
    const base = 13 * i + 1

    const rosterEntry: string[][] = []
    rosterEntry.push(paddedArray(['Scout', row[2], row[3]]))
    rosterEntry.push(paddedArray(['Allergies', row[12]]))
    rosterEntry.push(paddedArray(['Dietary Restrictions', row[13]]))
    rosterEntry.push(Array(4))
    rosterEntry.push(paddedArray(['Primary Contact', row[6], row[7]]))
    rosterEntry.push(paddedArray(['Secondary Contact', row[8], row[9]]))
    rosterEntry.push(Array(4))
    rosterEntry.push(paddedArray(["Physician's Name", row[10], 'Phone', row[11]]))
    rosterEntry.push(Array(4))
    rosterEntry.push(paddedArray(['Additional Notes', row[19]]))
    rosterEntry.push(Array(4))
    rosterEntry.push(paddedArray(['--------------------']))

    outputSheet.getRange(`A${base}:D${base + 11}`).setValues(rosterEntry)
  })
}
function drivingText(a: string[], name: string): string {
  const how = DRIVERS[a[15]]
  const num = parseInt(a[17]) + 1
  if (how == 'both') {
    return `${name} (${num})`
  } else {
    return `${name} (${how}: ${num})`
  }
}

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function createDrivingAssignments() {
  const outputSheet = createNamedSheet('Driving')
  const rd = responseData()

  const scouts = rd.filter(a => !ADULTS.includes(a[3]))
  const adultsCamping = rd.filter(a => ADULTS.includes(a[3]))

  const potentialDrivers = rd.filter(a => Object.keys(DRIVERS).includes(a[16]))

  const driving: (string | null)[][] = []
  const drivers: (string | null)[] = adultsCamping
    .map(a => drivingText(a, a[2]))
    .concat(potentialDrivers.map(a => drivingText(a, a[16] || `${a[2]}!!`)))
  // add empty column at the start to reserve the first column for the scout names
  drivers.unshift(null)

  driving.push(drivers)
  const columns = drivers.length

  scouts.forEach(s => driving.push(paddedArray([s[2]], columns)))

  const rows = scouts.length + 1
  outputSheet.getRange(1, 1, rows, columns).setValues(driving)
  outputSheet.autoResizeColumns(1, columns)
  outputSheet.getRange(2, 2, rows - 1, columns - 1).setHorizontalAlignment('center')
}
