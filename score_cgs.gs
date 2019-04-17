var start_cgs = 6;

function score_cgs_from_cell() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CervedAPI");
  if (sheet!=null)
  {
    var cell = sheet.getActiveCell().getColumn();
    if (cell==3)
    {
    var value = sheet.getActiveCell().getValue();
    var response = callScoreCGS(value);
    
     var html = HtmlService.createHtmlOutput("<pre>"+JSON.stringify(JSON.parse(response),null,2)+"<pre>").setTitle('Risultato dello Score CGS').setWidth(500);
     SpreadsheetApp.getUi().showSidebar(html);
    }
  }
}

function score_cgs(piva,idsoggetto,row) {
  var position = start_cgs+row;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CervedScore");
  var response = callScoreCGS(idsoggetto);
  var json = JSON.parse(response);
  
  var score = json.score != undefined ? json.scores[0] : undefined;
  
  if(piva != undefined) {
    sheet.getRange("A"+position).setValue(piva);
  }  else {
    sheet.getRange("A"+position).setValue("");
  }
  //sheet.getRange("A"+position).setValue(piva);
  sheet.getRange("B"+position).setValue(idsoggetto);
  sheet.getRange("C"+position).setValue(checkUndefined(json.denominazione));
  if(score != undefined) {
  	sheet.getRange("D"+position).setValue(checkUndefined(score.descrizione_score));
  	sheet.getRange("E"+position).setValue(checkUndefined(score.codice_score));
  	sheet.getRange("F"+position).setValue(checkUndefined(score.valore));
  	sheet.getRange("G"+position).setValue(checkUndefined(score.categoria_codice));
  	sheet.getRange("H"+position).setValue(checkUndefined(score.categoria_descrizione));
  	sheet.getRange("I"+position).setValue(checkUndefined(score.pd));
  	sheet.getRange("J"+position).setValue(checkUndefined(score.trend_codice));
  	sheet.getRange("K"+position).setValue(checkUndefined(score.trend_descrizione));
  }
}

function checkUndefined(p) {
  return p != undefined ? p : "n.a.";
}
