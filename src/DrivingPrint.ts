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

type DrivingPrintSettings = {title: string; subtitle: string}

const DEFAULT_DRIVING_PRINT_SETTINGS: DrivingPrintSettings = {title: '', subtitle: ''}

type Car = {
  driver: string
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

function drivingPage(sheet: GoogleAppsScript.Spreadsheet.Sheet): DrivingPage {
  const [headers, ...rows] = sheet.getDataRange().getDisplayValues()
  const scoutRows = rows.filter(row => cleanText(row[0]) !== '')

  const cars: Car[] = headers.flatMap((header, column) => {
    if (column === 0 || cleanText(header) === '') return []
    const passengers = scoutRows
      .filter(row => cleanText(row[column]) !== '')
      .map(row => cleanText(row[0]))
      .sort(compareByLastName)
    const car = {...parseDriverHeader(header), passengers}
    return [{...car, overBy: car.seats == null ? 0 : Math.max(0, passengers.length - car.seats)}]
  })
  cars.sort((a, b) => compareByLastName(a.driver, b.driver))

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
function getDrivingPrintPage(): {pages: DrivingPage[]; settings: DrivingPrintSettings} {
  const sheets = SpreadsheetApp.getActiveSpreadsheet()
    .getSheets()
    .filter(
      sheet => sheet.getName().startsWith('Driving') && !OLD_DRIVING_SHEET.test(sheet.getName()),
    )
  if (sheets.length === 0) {
    throw new Error('No Driving sheet found. Run "create driving sheet" first.')
  }
  return {
    pages: sheets.map(drivingPage),
    settings: readSettings(DRIVING_PRINT_SETTINGS_KEY, DEFAULT_DRIVING_PRINT_SETTINGS),
  }
}

// Called from DrivingPrintDialog.html via google.script.run.
function saveDrivingPrintSettings(settings: DrivingPrintSettings) {
  writeSettings(DRIVING_PRINT_SETTINGS_KEY, settings)
}
