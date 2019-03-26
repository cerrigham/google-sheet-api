var start_profile = 6;

function entity_profile_from_cell() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CervedAPI");
  if (sheet!=null)
  {
    var cell = sheet.getActiveCell().getColumn();
    if (cell==3)
    {
    var value = sheet.getActiveCell().getValue();
    var response = callEntityProfile(value);
    
     var html = HtmlService.createHtmlOutput("<pre>"+JSON.stringify(JSON.parse(response),null,2)+"<pre>").setTitle('Risultato del Profilo').setWidth(500);
     SpreadsheetApp.getUi().showSidebar(html);
    }
  }
}
  
function entity_profile(piva,idsoggetto,row) {
    
  var position = start_profile+row;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CervedProfile");
  
  if(piva != undefined) {
    sheet.getRange("A"+position).setValue(piva);
  } 
  sheet.getRange("B"+position).setValue(idsoggetto);
  
  var response = callEntityProfile(idsoggetto);
  var json = JSON.parse(response);
  var dati_anagrafici = json.dati_anagrafici;
  var dati_attivita = json.dati_attivita;
  var dati_economici = json.dati_economici_dimensionali;
  if (dati_anagrafici.id_soggetto==idsoggetto)
  {
    if(piva == undefined) {
      sheet.getRange("A"+position).setValue(dati_anagrafici.codice_fiscale);
    }
    sheet.getRange("B"+position).setValue(idsoggetto);
    sheet.getRange("C"+position).setValue(checkUndefined(dati_anagrafici.telefono));
    sheet.getRange("D"+position).setValue(checkUndefined(dati_anagrafici.url_sito_web));
    if (dati_anagrafici.pec && dati_anagrafici.pec != undefined)
    {
      sheet.getRange("E"+position).setValue(checkUndefined(dati_anagrafici.pec.email[0]));
    }
    
    if(dati_attivita != undefined) {
      sheet.getRange("F"+position).setValue(checkUndefined(dati_attivita.data_costituzione));
      sheet.getRange("G"+position).setValue(checkUndefined(dati_attivita.data_inizio_attivita));
      sheet.getRange("H"+position).setValue(checkUndefined(dati_attivita.natura_giuridica));
      sheet.getRange("I"+position).setValue(checkUndefined(dati_attivita.data_iscrizione_rea));
      sheet.getRange("J"+position).setValue(checkUndefined(dati_attivita.codice_rea));
    }
    
    // Dati economico dimensionali
    if(dati_economici != undefined) {
      sheet.getRange("O"+position).setValue(checkUndefined(dati_economici.numero_dipendenti));
      sheet.getRange("p"+position).setValue(checkUndefined(dati_economici.numero_unita_locali));
      sheet.getRange("Q"+position).setValue(checkUndefined(dati_economici.anno_ultimo_bilancio));
      sheet.getRange("R"+position).setValue(checkUndefined(dati_economici.data_chiusura_ultimo_bilancio));
      sheet.getRange("S"+position).setValue(checkUndefined(dati_economici.fatturato));
      sheet.getRange("T"+position).setValue(checkUndefined(dati_economici.capitale_sociale));
      sheet.getRange("U"+position).setValue(checkUndefined(dati_economici.mol));
      sheet.getRange("V"+position).setValue(checkUndefined(dati_economici.attivo));
      sheet.getRange("W"+position).setValue(checkUndefined(dati_economici.patrimonio_netto));
    }
  }
  
  if(dati_attivita != undefined) {
    var ateco_info = dati_attivita.ateco_info
    if (ateco_info && ateco_info != undefined)
    {
      var codifica_ateco = ateco_info.codifica_ateco
      if (codifica_ateco && codifica_ateco != undefined)
      {
        sheet.getRange("K"+position).setValue(codifica_ateco.codice_ateco);    
        sheet.getRange("L"+position).setValue(codifica_ateco.ateco);    
        sheet.getRange("M"+position).setValue(codifica_ateco.macrosettore);
        sheet.getRange("N"+position).setValue(codifica_ateco.codice_macrosettore);    
      }
    }
  }
}

function checkUndefined(p) {
  return p != undefined ? p : "n.a.";
}