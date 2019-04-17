var start_search = 6;

function entity_search_from_cell() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CervedAPI");
  if (sheet!=null)
  {
    var cell = sheet.getActiveCell().getColumn();
    if (cell==1)
    {
    var value = sheet.getActiveCell().getValue();
    var response = callEntitySearch(value);
    
     var html = HtmlService.createHtmlOutput("<pre>"+JSON.stringify(JSON.parse(response),null,2)+"<pre>").setTitle('Risultato della ricerca').setWidth(500);
     SpreadsheetApp.getUi().showSidebar(html);
    }
  }
}

function entity_search_from_column() {
  if (!check_data()) 
    return;
  
  var ui = SpreadsheetApp.getUi();
  
  /* controllo presenza APIKEY */
  if(property('APIKEY')) {
    Logger.log("APIKEY inserita!");
    Logger.log(property('APIKEY'));
  } else {
    Logger.log("APIKEY non inserita!");
    ui.alert('Attenzione', 'Inserire APIKEY', ui.ButtonSet.OK);
    return;
  }

  // gestione modal durante esecuzione delle chiamate
  var output = HtmlService.createHtmlOutput('Si prega di non chiudere la finestra');
  ui.showModalDialog(output, 'Arricchimento...');
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CervedAPI");
  var end = false;
  var row = 0;
  var sourceCol = "A";
  if (sheet!=null)
  {
    while(!end) {
      var position = (row+start_search); 
      var source  = sourceCol + '' + position;
      if (sheet.getRange(source).isBlank()) {
        end=true;
      } else {
        var input = sheet.getRange(source).getValue();
        var response = callEntitySearch(input);
        flatten(sheet,row,response);
      }
      row = row+1;
    }
    
    // pulizia righe vuote
    deleteEmptyRows("CervedProfile");
    deleteEmptyRows("CervedScore");
    deleteEmptyRows("REScore");
    
    // gestione modal alla fine dell'esecuzione delle chiamate
    var output = HtmlService.createHtmlOutput('<script>google.script.host.close();</script>');
    
    ui.alert('Arricchimento completato');
  }
}

// pulizia righe vuote
// @see https://yagisanatode.com/2017/12/13/google-apps-script-iterating-through-ranges-in-sheets-the-right-and-wrong-way/
function deleteEmptyRows(sheetName) {
   var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    Logger.log("Pulizia di ");
    Logger.log(sheetName);
    var rangeData = s.getDataRange();
    var lastColumn = rangeData.getLastColumn();
    var lastRow = rangeData.getLastRow();
    for(var z = start_search; z < lastRow; z++) {
      //Logger.log(rangeData.getCell(z, 2).getValue());
      if(!rangeData.getCell(z, 2).getValue()) {
        s.deleteRow(z)
      }
    }
}

function flatten(sheet, row, response) {
  var json = JSON.parse(response);
  var match = json.peopleTotalNumber+json.companiesTotalNumber;
  var position = start_search+row;
  sheet.getRange('B'+position).setValue(match);
  if (match==1)
  {
    var collection = json.companies[0];
    if (json.peopleTotalNumber==1)
    {
      //collection = json.persons;
      collection = json.people[0];
    }
    var dati_anagrafici = collection.dati_anagrafici;
    var dati_attivita = collection.dati_attivita;
    sheet.getRange('C'+position).setValue(dati_anagrafici.id_soggetto);
    sheet.getRange('D'+position).setValue(dati_anagrafici.denominazione);
    sheet.getRange('E'+position).setValue(checkUndefined(dati_anagrafici.codice_fiscale));
    sheet.getRange('F'+position).setValue(checkUndefined(dati_anagrafici.partita_iva));
    if(dati_attivita != undefined) {
      sheet.getRange('G'+position).setValue(checkUndefined(dati_attivita.codice_ateco));
      sheet.getRange('H'+position).setValue(checkUndefined(dati_attivita.ateco));
      sheet.getRange('I'+position).setValue(checkUndefined(dati_attivita.codice_stato_attivita));
      sheet.getRange('J'+position).setValue(checkUndefined(dati_attivita.flag_operativa));
      sheet.getRange('K'+position).setValue(checkUndefined(dati_attivita.codice_rea));
    }
    if(dati_anagrafici.indirizzo != undefined) {
      sheet.getRange('L'+position).setValue(checkUndefined(dati_anagrafici.indirizzo.descrizione));
      sheet.getRange('M'+position).setValue(checkUndefined(dati_anagrafici.indirizzo.cap));
      sheet.getRange('N'+position).setValue(checkUndefined(dati_anagrafici.indirizzo.codice_comune));
      sheet.getRange('O'+position).setValue(checkUndefined(dati_anagrafici.indirizzo.codice_comune_istat));
      sheet.getRange('P'+position).setValue(checkUndefined(dati_anagrafici.indirizzo.provincia));
    }
    // Cascade on profile
    if (property('entity_profile_enabled')=='on')
    {
      entity_profile(dati_anagrafici.partita_iva,dati_anagrafici.id_soggetto,row);
    }
    
    // Cascade on profile
    if (property('score_enabled')=='on')
    {
      score_cgs(dati_anagrafici.partita_iva,dati_anagrafici.id_soggetto,row);
    }
    
     // Cascade on profile
    if (property('real_estate_score_enabled')=='on')
    {
      real_estate_score(dati_anagrafici.partita_iva,dati_anagrafici.id_soggetto,row);
    }
  }
}

function checkUndefined(p) {
  return p != undefined ? p : "n.a.";
}
