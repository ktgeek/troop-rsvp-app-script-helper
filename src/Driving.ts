const DRIVERS: {[id: string]: string} = {
  'Drive both ways': 'both',
  'Drive to the campout only': 'to',
  'Drive from the campout only': 'from',
  'Tow both ways': 'tow',
  'Tow to the campout only': 'tow to',
  'Tow from the campout only': 'tow from',
}

function paddedArray(array: string[], len: number = 4): string[] {
  return Array.from({...array, length: len})
}

function drivingText(a: ResponseRow, name: string): string {
  const how = DRIVERS[a.driving]
  const num = parseInt(a.seats) + 1

  return how == 'both' ? `${name} (${num})` : `${name} (${how}: ${num})`
}

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
