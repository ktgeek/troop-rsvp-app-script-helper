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
  // Skip the empty case, where a range needs at least one row.
  if (rd.length > 0) {
    outputSheet.getRange(2, 3, rd.length).setHorizontalAlignment('right')
  }
  // Adults only got their 'A' above, so sort once the listing is complete.
  sortPatrolSheet()
}

// Re-sorts the Patrols sheet's data rows by Campout Patrol, then Patrol, then Scout's last name. Campout patrols
// compare as plain text, so blanks come first and 'A' lands after numbered patrols. Used after hand-assigning
// campout patrols to check that experience, age, and home patrols are balanced.
function sortPatrolSheet() {
  const ui = SpreadsheetApp.getUi()
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Patrols')
  if (sheet == null) {
    ui.alert('There is no Patrols sheet to sort. Run "create patrol sheet" first.')
    return
  }
  if (sheet.getLastRow() < 2) return

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String)
  const normalizedHeaders = headers.map(normalizeHeader)
  const [scout, patrol, campout] = ['Scout', 'Patrol', 'Campout Patrol'].map(header =>
    firstColumn(normalizedHeaders, [header], `Patrols sheet: ${header}`),
  )
  if ([scout, patrol, campout].includes(-1)) {
    ui.alert('The Patrols sheet needs Scout, Patrol, and Campout Patrol columns to sort.')
    return
  }

  const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length)
  const text = (row: unknown[], column: number) => cleanText(String(row[column]))
  const sorted = range
    .getValues()
    .sort(
      (a, b) =>
        text(a, campout).localeCompare(text(b, campout)) ||
        text(a, patrol).localeCompare(text(b, patrol)) ||
        compareByLastName(text(a, scout), text(b, scout)),
    )
  range.setValues(sorted)
}
