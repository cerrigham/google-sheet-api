function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var mainMenu = ui.createMenu("Cerved API");
  
  mainMenu.addItem("Cerca Selezionata", "entity_search_from_cell");
  mainMenu.addItem("Profilo Selezionata", "entity_profile_from_cell");
  mainMenu.addItem("Score CGS Selezionata", "score_cgs_from_cell");
  mainMenu.addItem("Real Estate Score Selezionata", "real_estate_score_from_cell");
  mainMenu.addSeparator();
  mainMenu.addItem("Arricchimento dati", "showSidebar");
  mainMenu.addSeparator();
  mainMenu.addItem("Azzera arricchimento", "clear_data");
  mainMenu.addItem("Impostazioni", "apikey_modal");
  mainMenu.addToUi();
  
  update_ui();
}

function property(name) {
   var docProp = PropertiesService.getDocumentProperties()
   return docProp.getProperty(name)
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('sidebar').setTitle('Scelta dei prodotti').setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

function formSubmit(form){
  var docProp = PropertiesService.getDocumentProperties()
  docProp.setProperty("entity_profile_enabled", form.entity_profile_enabled)
  docProp.setProperty("score_enabled", form.score_enabled)
  docProp.setProperty("real_estate_score_enabled", form.real_estate_score_enabled)
  entity_search_from_column();
}
