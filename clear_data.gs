function clear_data() {
  var ui = SpreadsheetApp.getUi();
  if (!SpreadsheetApp.getActiveSheet().getRange("B6").isBlank()) {
        result = ui.alert('Attenzione', 'Verranno eliminati i dati. Confermi?', ui.ButtonSet.YES_NO);
        if (result == ui.Button.NO) 
          return false;
      }
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CervedAPI");
  sheet.getRange(6, 2, 1000, 30).clear();
  
  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CervedProfile");
  sheet.getRange(6, 1, 1000, 30).clear();
  
  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CervedScore");
  sheet.getRange(6, 1, 1000, 30).clear();
  
  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("REScore");
  sheet.getRange(6, 1, 1000, 30).clear();
  
  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("REScore.Fabbricati");
  sheet.getRange(6, 1, 1000, 30).clear();
  
  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("REScore.Terreni");
  sheet.getRange(6, 1, 1000, 30).clear();
  
  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("REScore.Fabbricati.Possessi");
  sheet.getRange(6, 1, 1000, 30).clear();
  
  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("REScore.Terreni.Possessi");
  sheet.getRange(6, 1, 1000, 30).clear();
}
