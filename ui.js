function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Vetting Tool").addItem("Vet Level", "vetLevels").addToUi();
}