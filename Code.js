function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Report')
    .addItem('Refresh Stores', 'refreshStores')
    .addItem('Add entry', 'refresh')
    .addItem('Sort', 'sortRange')
    .addToUi();
};
