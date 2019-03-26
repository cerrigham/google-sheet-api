var scriptProperties = PropertiesService.getScriptProperties();

var start_real_estate_score = 6;
var start_position_fabbricati = 6;
var start_position_terreni = 6;
var start_position_fabbricati_possessi = 6;
var start_position_terreni_possessi = 6;

function real_estate_score_from_cell() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CervedAPI");
  if (sheet!=null)
  {
    var cell = sheet.getActiveCell().getColumn();
    if (cell==3)
    {
    var value = sheet.getActiveCell().getValue();
    var response = callRealEstateScore(value);
    
     var html = HtmlService.createHtmlOutput("<pre>"+JSON.stringify(JSON.parse(response),null,2)+"<pre>").setTitle('Risultato della Real Estate Score').setWidth(500);
     SpreadsheetApp.getUi().showSidebar(html);
    }
  }
}


function real_estate_score(piva,idsoggetto,row) {
  var position = start_real_estate_score+row;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("REScore");
  var response = callRealEstateScore(idsoggetto);
  var json = JSON.parse(response);
  
  // Info prese da DatiImmobiliariDTO.java e classi collegate
  if(piva != undefined) {
    sheet.getRange("A"+position).setValue(piva);
  } else {
    sheet.getRange("A"+position).setValue(json.codiceFiscale);
  }
  sheet.getRange("B"+position).setValue(idsoggetto);
  sheet.getRange("C"+position).setValue(json.dataRapporto);
  sheet.getRange("D"+position).setValue(json.numeroImmobili);
  sheet.getRange("E"+position).setValue(json.numeroFabbricati);
  sheet.getRange("F"+position).setValue(json.numeroTerreni);
  sheet.getRange("G"+position).setValue(json.scoreImmobiliare != undefined ? json.scoreImmobiliare.classe : "");
  
  // gestisco fabbricati
  if (json.fabbricati.length) {
    Logger.log("Ci sono Fabbricati");
    
    for (var i = 0; i < json.fabbricati.length; i++) {
      var fabbricato = json.fabbricati[i];
      var sheetFabbricati = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("REScore.Fabbricati");
      if (sheetFabbricati!=null) {
        if(piva != undefined) {
          sheetFabbricati.getRange("A"+start_position_fabbricati).setValue(piva);
        } else {
          sheetFabbricati.getRange("A"+start_position_fabbricati).setValue(json.codiceFiscale);
        }
        sheetFabbricati.getRange("B"+start_position_fabbricati).setValue(idsoggetto);
        sheetFabbricati.getRange("C"+start_position_fabbricati).setValue(fabbricato.idImmobile);
        sheetFabbricati.getRange("D"+start_position_fabbricati).setValue(checkUndefined(fabbricato.classe));
        sheetFabbricati.getRange("E"+start_position_fabbricati).setValue(fabbricato.codiceBelfiore);
        sheetFabbricati.getRange("F"+start_position_fabbricati).setValue(fabbricato.codiceComune);
        sheetFabbricati.getRange("G"+start_position_fabbricati).setValue(fabbricato.descrizioneComune);
        sheetFabbricati.getRange("H"+start_position_fabbricati).setValue(fabbricato.codiceProvincia);
        sheetFabbricati.getRange("I"+start_position_fabbricati).setValue(fabbricato.foglio);
        sheetFabbricati.getRange("J"+start_position_fabbricati).setValue(fabbricato.particella);
        sheetFabbricati.getRange("K"+start_position_fabbricati).setValue(checkUndefined(fabbricato.denominatoreParticella));
        sheetFabbricati.getRange("L"+start_position_fabbricati).setValue(checkUndefined(fabbricato.subalterno));
        sheetFabbricati.getRange("M"+start_position_fabbricati).setValue(checkUndefined(fabbricato.sezioneAmministrativa));
        sheetFabbricati.getRange("N"+start_position_fabbricati).setValue(checkUndefined(fabbricato.sezioneUrbana));
        sheetFabbricati.getRange("O"+start_position_fabbricati).setValue(fabbricato.indirizzo);
        sheetFabbricati.getRange("P"+start_position_fabbricati).setValue(fabbricato.piano);
        sheetFabbricati.getRange("Q"+start_position_fabbricati).setValue(checkUndefined(fabbricato.codiceCategoria));
        sheetFabbricati.getRange("R"+start_position_fabbricati).setValue(checkUndefined(fabbricato.descrizioneCategoria));
        sheetFabbricati.getRange("S"+start_position_fabbricati).setValue(checkUndefined(fabbricato.valoreConsistenza));
        sheetFabbricati.getRange("T"+start_position_fabbricati).setValue(checkUndefined(fabbricato.unitaMisuraConsistenza));
        sheetFabbricati.getRange("U"+start_position_fabbricati).setValue(checkUndefined(fabbricato.rendita));
        sheetFabbricati.getRange("V"+start_position_fabbricati).setValue(checkUndefined(fabbricato.superficieCatastale));
        sheetFabbricati.getRange("W"+start_position_fabbricati).setValue(checkUndefined(fabbricato.superficieCatastaleCoperta));
        if (fabbricato.stimaFabbricato != undefined) {
          sheetFabbricati.getRange("X"+start_position_fabbricati).setValue(checkUndefined(fabbricato.stimaFabbricato.valoreMinNormale));
          sheetFabbricati.getRange("Y"+start_position_fabbricati).setValue(checkUndefined(fabbricato.stimaFabbricato.valoreMaxNormale));
          sheetFabbricati.getRange("Z"+start_position_fabbricati).setValue(checkUndefined(fabbricato.stimaFabbricato.valoreMinOttimo));
          sheetFabbricati.getRange("AA"+start_position_fabbricati).setValue(checkUndefined(fabbricato.stimaFabbricato.valoreMaxOttimo));
          sheetFabbricati.getRange("AB"+start_position_fabbricati).setValue(checkUndefined(fabbricato.stimaFabbricato.valoreMinScadente));
          sheetFabbricati.getRange("AC"+start_position_fabbricati).setValue(checkUndefined(fabbricato.stimaFabbricato.valoreMaxScadente));
          sheetFabbricati.getRange("AD"+start_position_fabbricati).setValue(checkUndefined(fabbricato.stimaFabbricato.valorePuntuale));
          sheetFabbricati.getRange("AE"+start_position_fabbricati).setValue(checkUndefined(fabbricato.stimaFabbricato.livelloConfidenza));
        }
      }
      start_position_fabbricati = start_position_fabbricati + 1;
      
      // gestione possessi per fabbricati
      var sheetPossessoFabbricati = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("REScore.Fabbricati.Possessi");
      populate_possessi(sheetPossessoFabbricati, fabbricato, piva, idsoggetto, "F", start_position_fabbricati_possessi);
    }
    
    // gestione possessi per fabbricati
    //var sheetPossessoFabbricati = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("REScore.Fabbricati.Possessi");
    //populate_possessi(sheetPossessoFabbricati, fabbricato, piva, idsoggetto, "F", start_position_fabbricati_possessi);
  } else {
    Logger.log("Non ci sono Fabbricati");
  }
  
 // gestisco terreni
  if (json.terreni.length) {
     Logger.log("Ci sono Terreni");
    
    for (var i = 0; i < json.terreni.length; i++) {
      var terreno = json.terreni[i];
      var sheetTerreni = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("REScore.Terreni");
      if (sheetTerreni!=null) {
        if(piva != undefined) {
          sheetTerreni.getRange("A"+start_position_terreni).setValue(piva);
        } else {
         sheetTerreni.getRange("A"+start_position_terreni).setValue(json.codiceFiscale);
        }
        //sheetTerreni.getRange("A"+start_position_terreni).setValue(piva);
        sheetTerreni.getRange("B"+start_position_terreni).setValue(idsoggetto);
        sheetTerreni.getRange("C"+start_position_terreni).setValue(terreno.idImmobile);
        sheetTerreni.getRange("D"+start_position_terreni).setValue(terreno.classe);
        sheetTerreni.getRange("E"+start_position_terreni).setValue(terreno.codiceBelfiore);
        sheetTerreni.getRange("F"+start_position_terreni).setValue(terreno.codiceComune);
        sheetTerreni.getRange("G"+start_position_terreni).setValue(terreno.descrizioneComune);
        sheetTerreni.getRange("H"+start_position_terreni).setValue(terreno.codiceProvincia);
        sheetTerreni.getRange("I"+start_position_terreni).setValue(terreno.foglio);
        sheetTerreni.getRange("J"+start_position_terreni).setValue(terreno.particella);
        sheetTerreni.getRange("K"+start_position_terreni).setValue(checkUndefined(terreno.denominatoreParticella));
        sheetTerreni.getRange("L"+start_position_terreni).setValue(checkUndefined(terreno.subalterno));
        sheetTerreni.getRange("M"+start_position_terreni).setValue(checkUndefined(terreno.sezioneCensuaria))
        sheetTerreni.getRange("N"+start_position_terreni).setValue(terreno.codicePorzione);
        sheetTerreni.getRange("O"+start_position_terreni).setValue(terreno.descrizioneQualita);
        sheetTerreni.getRange("P"+start_position_terreni).setValue(terreno.superficieEttari);
        sheetTerreni.getRange("Q"+start_position_terreni).setValue(terreno.superficieAre);
        sheetTerreni.getRange("R"+start_position_terreni).setValue(terreno.superficieCentiare);
        sheetTerreni.getRange("S"+start_position_terreni).setValue(terreno.renditaDominicale);
        sheetTerreni.getRange("T"+start_position_terreni).setValue(terreno.renditaAgraria);
      }
      start_position_terreni = start_position_terreni + 1;
          
      // gestione possessi per terreni
      var sheetPossessoTerreni = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("REScore.Terreni.Possessi");
      populate_possessi(sheetPossessoTerreni, terreno, piva, idsoggetto, "T", start_position_terreni_possessi);
    }
    
    // gestione possessi per terreni
    //var sheetPossessoTerreni = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("REScore.Terreni.Possessi");
    //populate_possessi(sheetPossessoTerreni, terreno, piva, idsoggetto, "T", start_position_terreni_possessi);
  } else {
    Logger.log("Non ci sono Terreni");
  }
}

function populate_possessi(sheetPossessi, data, piva, idsoggetto, flag, start_position_possessi) {
  Logger.log("Populate_possessi");
  if(data.possessi.length) {
    Logger.log("Ci sono possessi");
    
    for (var j = 0; j < data.possessi.length; j++) {
      var possesso = data.possessi[j];
      if (sheetPossessi!=null) {
        Logger.log("Si sheet");
        //sheetPossessi.getRange("A"+start_position_possessi).setValue(piva);
        //sheetPossessi.getRange("B"+start_position_possessi).setValue(idsoggetto);
        sheetPossessi.getRange("A"+start_position_possessi).setValue(data.idImmobile);
        sheetPossessi.getRange("B"+start_position_possessi).setValue(checkUndefined(possesso.descrizioneTitolo));
        sheetPossessi.getRange("C"+start_position_possessi).setValue(checkUndefined(possesso.titolaritaOrig));
        sheetPossessi.getRange("D"+start_position_possessi).setValue(checkUndefined(possesso.descrizioneRegime));
        sheetPossessi.getRange("E"+start_position_possessi).setValue(checkUndefined(possesso.regimeOrig));
        sheetPossessi.getRange("F"+start_position_possessi).setValue(checkUndefined(possesso.quotaOrig));
        sheetPossessi.getRange("G"+start_position_possessi).setValue(checkUndefined(possesso.percentualeQuota));
      }
      if(flag == "F") {
        start_position_fabbricati_possessi = start_position_fabbricati_possessi + 1;
      } else {
        start_position_terreni_possessi = start_position_terreni_possessi + 1;
      }
      start_position_possessi = start_position_possessi + 1;
    }
  } else {
    Logger.log("Non ci sono possessi");
  }
}

function checkUndefined(p) {
  return p != undefined ? p : "";
}
