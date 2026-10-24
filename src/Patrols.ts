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
