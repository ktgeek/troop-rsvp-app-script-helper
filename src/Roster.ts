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
