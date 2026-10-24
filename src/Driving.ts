const DRIVERS: {[id: string]: string} = {
  'Drive both ways': 'both',
  'Drive to the campout only': 'to',
  'Drive from the campout only': 'from',
  'Tow both ways': 'tow',
  'Tow to the campout only': 'tow to',
  'Tow from the campout only': 'tow from',
}

// Appended to a scout's name in the header when their parent drives but gave no driver name, so the column
// stands out for a person to fix; the driving printout reads it back.
const UNNAMED_DRIVER_MARK = '!!'

function paddedArray(array: string[], len: number = 4): string[] {
  return Array.from({...array, length: len})
}

function hasSeatCount(a: ResponseRow): boolean {
  return !isNaN(parseInt(a.seats))
}

// A "?" marks an answer that is missing or isn't a number (camping adults are listed even without one, and
// parents can type words into Seats), so a person notices instead of the driver silently vanishing.
function drivingText(a: ResponseRow, name: string): string {
  const how = DRIVERS[a.driving] ?? '?'
  const num = hasSeatCount(a) ? parseInt(a.seats) + 1 : '?'

  return how == 'both' ? `${name} (${num})` : `${name} (${how}: ${num})`
}

function createDrivingAssignments() {
  const outputSheet = createNamedSheet('Driving')
  const rd = responseData()

  const adultsCamping = rd.filter(a => ADULTS.includes(a.patrol))
  const scouts = rd.filter(a => !ADULTS.includes(a.patrol))
  const parentDrivers = scouts.filter(
    a => Object.keys(DRIVERS).includes(a.driving) && a.seats.trim() !== '',
  )

  const drivers: (string | null)[] = [
    null,
    ...adultsCamping.map(a => drivingText(a, a.scout)),
    ...parentDrivers.map(a => drivingText(a, a.driverName || `${a.scout}${UNNAMED_DRIVER_MARK}`)),
  ]

  const columns = drivers.length
  const drivingData: (string | null)[][] = [
    drivers,
    ...scouts.map(s => paddedArray([s.scout], drivers.length)),
  ]

  const rows = scouts.length + 1
  outputSheet.getRange(1, 1, rows, columns).setValues(drivingData)
  outputSheet.autoResizeColumns(1, columns)
  // A range needs at least one row and column, which is missing when there are no scouts or no drivers.
  if (rows > 1 && columns > 1) {
    outputSheet.getRange(2, 2, rows - 1, columns - 1).setHorizontalAlignment('center')
  }
}
