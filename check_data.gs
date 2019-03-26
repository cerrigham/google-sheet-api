function check_data() {
  
  var ui = SpreadsheetApp.getUi();
  if (!SpreadsheetApp.getActiveSheet().getRange("B6").isBlank()) {
        result = ui.alert('Attenzione', 'Verranno sovrascritti i dati. Confermi?', ui.ButtonSet.YES_NO);
        if (result == ui.Button.NO) 
          return false;
      }
  return true;
}
