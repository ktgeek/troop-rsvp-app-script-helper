// "Driving (old)", "Driving (old 2)", ... are the copies createNamedSheet sets aside; they are never printed.
const OLD_DRIVING_SHEET = /^Driving \(old( \d+)?\)$/

// How a driver's label in the column header reads on the printout. Unknown labels print as typed.
const DRIVING_LABELS: {[label: string]: string} = {
  to: 'To the campout only',
  from: 'From the campout only',
  tow: 'Towing both ways',
  'tow to': 'Towing to the campout only',
  'tow from': 'Towing from the campout only',
}

const DRIVING_PRINT_SETTINGS_KEY = 'drivingPrintSettings'

type Car = {
  driver: string
  // The scout whose parent drives, when the form gave no driver name; empty otherwise.
  parentOf: string
  // Empty when the driver drives both ways.
  label: string
  // Scout seats, including the driver's own scout (the form's count plus one); null when the header has no number.
  seats: number | null
  passengers: string[]
  overBy: number
}

type DrivingPage = {
  name: string
  cars: Car[]
  unassigned: string[]
  // Scouts marked in more than one car, which is usually a slip while assigning.
  inSeveralCars: string[]
}

// Reads "Name (5)", "Name (to: 5)" or "Name (tow from: ?)", the headers drivingText writes. A header that was retyped
// by hand and no longer fits is kept whole as the driver, with unknown seats.
function parseDriverHeader(header: string): {driver: string; label: string; seats: number | null} {
  const match = cleanText(header).match(/^(.*?)\s*\((?:(.*?):\s*)?(\d+|\?)\)$/)
  if (match == null) {
    return {driver: cleanText(header), label: '', seats: null}
  }
  const [, driver, label = '', seats] = match
  return {
    driver,
    label: DRIVING_LABELS[label] ?? label,
    seats: seats === '?' ? null : parseInt(seats),
  }
}

// Checkbox cells read "FALSE" when unchecked; every other non-empty cell is a mark.
function isMarked(cell: string): boolean {
  const mark = cleanText(cell).toUpperCase()
  return mark !== '' && mark !== 'FALSE'
}

// The generated Driving sheet leaves A1 blank, which tells it apart from a spare tab like "Driving notes".
function isDrivingSheet(sheet: GoogleAppsScript.Spreadsheet.Sheet): boolean {
  const name = sheet.getName()
  return (
    name.startsWith('Driving') &&
    !OLD_DRIVING_SHEET.test(name) &&
    cleanText(sheet.getRange(1, 1).getDisplayValue()) === ''
  )
}

function drivingPage(sheet: GoogleAppsScript.Spreadsheet.Sheet): DrivingPage {
  const [headers, ...rows] = sheet.getDataRange().getDisplayValues()
  const scoutRows = rows.filter(row => cleanText(row[0]) !== '')

  const cars: Car[] = headers.flatMap((header, column) => {
    if (column === 0 || cleanText(header) === '') return []
    const passengers = scoutRows
      .filter(row => isMarked(row[column]))
      .map(row => cleanText(row[0]))
      .sort(compareByLastName)
    const {driver, label, seats} = parseDriverHeader(header)
    const unnamed = driver.endsWith(UNNAMED_DRIVER_MARK)
    return [
      {
        driver: unnamed ? 'Driver not named' : driver,
        parentOf: unnamed ? driver.slice(0, -UNNAMED_DRIVER_MARK.length).trim() : '',
        label,
        seats,
        passengers,
        overBy: seats == null ? 0 : Math.max(0, passengers.length - seats),
      },
    ]
  })
  // An unnamed driver sorts under the scout they are the parent of.
  cars.sort((a, b) => compareByLastName(a.parentOf || a.driver, b.parentOf || b.driver))

  const carsPerScout = (name: string) => cars.filter(car => car.passengers.includes(name)).length
  const names = scoutRows.map(row => cleanText(row[0])).sort(compareByLastName)
  return {
    name: sheet.getName(),
    cars,
    unassigned: names.filter(name => carsPerScout(name) === 0),
    inSeveralCars: names.filter(name => carsPerScout(name) > 1),
  }
}

function showDrivingPrint() {
  const dialog = HtmlService.createHtmlOutputFromFile('DrivingPrintDialog')
    .setWidth(1000)
    .setHeight(760)
  SpreadsheetApp.getUi().showModalDialog(dialog, 'Printable driving sheet')
}

// Called from DrivingPrintDialog.html via google.script.run.
function getDrivingPrintPage(): {pages: DrivingPage[]; settings: HeaderSettings} {
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets().filter(isDrivingSheet)
  if (sheets.length === 0) {
    throw new Error('No Driving sheet found. Run "create driving sheet" first.')
  }
  return {
    pages: sheets.map(drivingPage),
    settings: readSettings(DRIVING_PRINT_SETTINGS_KEY, DEFAULT_HEADER_SETTINGS),
  }
}

// Called from DrivingPrintDialog.html via google.script.run.
function saveDrivingPrintSettings(settings: HeaderSettings) {
  writeSettings(DRIVING_PRINT_SETTINGS_KEY, settings)
}
