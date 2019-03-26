function apikey_modal() {
  
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var result = ui.prompt(
      'Impostazioni',
      'Prego immettere la vostra APIKEY:',
      ui.ButtonSet.OK_CANCEL);

  // Process the user's response.
  var button = result.getSelectedButton();
  var text = result.getResponseText();
  
  if (button == ui.Button.OK) {
    var docProp = PropertiesService.getDocumentProperties()
    docProp.setProperty("APIKEY", text)
    update_ui();
  } else if (button == ui.Button.CANCEL) {
  } else if (button == ui.Button.CLOSE) {
  }
}