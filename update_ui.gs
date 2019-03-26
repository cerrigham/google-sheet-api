function update_ui() {
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CervedAPI");
  if (sheet != null) {
    var apikey = property("APIKEY");
    if (apikey!=null && apikey.length>0) {
      sheet.getRange("D2").setBackground("#08F");
      sheet.getRange("D2").setFontColor("#fff");
      sheet.getRange("D2").setValue("Sistema pronto");
    } else {
      sheet.getRange("D2").setBackground("#888");
      sheet.getRange("D2").setFontColor("#fff");
      sheet.getRange("D2").setValue("Prego inserire la vostra APIKEY");
    }
  }
}