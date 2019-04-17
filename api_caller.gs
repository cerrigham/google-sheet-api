var options = {
    'muteHttpExceptions': true,
    'method': 'get'
  };

var baseUrl = 'https://api.cerved.com/cervedApi';
  
function callEntityProfile(id) {
  var response = UrlFetchApp.fetch(baseUrl+'/v1/entityProfile/live?id_soggetto='+id+'&apikey='+property('APIKEY'), options);
  var json = response.getContentText();
  return json;
}

function callEntitySearch(input) {
  var response = UrlFetchApp.fetch(baseUrl+'/v1/entitySearch/live?testoricerca='+input+'&apikey='+property('APIKEY'), options);
  var json = response.getContentText();
  return json;
}

function callScoreCGS(input) {
  // deprecated path
  // var response = UrlFetchApp.fetch(baseUrl+'score/impresa/cgs/corporate?subjectId='+input+'&apikey='+property('APIKEY'), options);
  // new path
  var response = UrlFetchApp.fetch(baseUrl+'/v1.1/score/impresa/corporate/C7_INTG?id_soggetto='+input+'&apikey='+property('APIKEY'), options);
  var json = response.getContentText();
  return json; 
}

function callRealEstateScore(input) {
  var response = UrlFetchApp.fetch(baseUrl+'/v1/realEstateData/score?idSoggetto='+input+'&apikey='+property('APIKEY'), options);
  var json = response.getContentText();
  return json; 
}
