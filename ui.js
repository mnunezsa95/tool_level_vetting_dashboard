// Updated on 11/13/24

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Vetting Tool").addItem("Vet Level", "vetLevels").addToUi();
}
