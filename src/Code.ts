function onOpen() {
  createMenu()
}

function createMenu() {
  SpreadsheetApp.getUi()
    .createMenu('Troop 66')
    .addItem('create patrol sheet', 'createPatrolListing')
    .addItem('sort patrol sheet', 'sortPatrolSheet')
    .addSeparator()
    .addItem('create driving sheet', 'createDrivingAssignments')
    .addSeparator()
    .addItem('print roster', 'showRoster')
    .addToUi()
}
