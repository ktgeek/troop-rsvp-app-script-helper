function onOpen() {
  createMenu()
}

function createMenu() {
  SpreadsheetApp.getUi()
    .createMenu('Troop 66')
    .addItem('create patrol sheet', 'createPatrolListing')
    .addItem('create driving sheet', 'createDrivingAssignments')
    .addItem('print roster', 'showRoster')
    .addToUi()
}
