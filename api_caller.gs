var options = {
    'muteHttpExceptions': true,
    'method': 'get'
  };

var baseUrl = 'https://api.cerved.com/cervedApi/v1/';
  
function callEntityProfile(id) {
  var response = UrlFetchApp.fetch(baseUrl+'entityProfile/live?id_soggetto='+id+'&apikey='+property('APIKEY'), options);
  var json = response.getContentText();
  return json;
}

function callEntitySearch(input) {
  var response = UrlFetchApp.fetch(baseUrl+'entitySearch/live?testoricerca='+input+'&apikey='+property('APIKEY'), options);
  var json = response.getContentText();
  return json;
}

function callScoreCGS(input) {
  var response = UrlFetchApp.fetch(baseUrl+'score/impresa/cgs/corporate?subjectId='+input+'&apikey='+property('APIKEY'), options);
  var json = response.getContentText();
  return json; 
}

function callRealEstateScore(input) {
  var response = UrlFetchApp.fetch(baseUrl+'realEstateData/score?idSoggetto='+input+'&apikey='+property('APIKEY'), options);
  var json = response.getContentText();
  return json; 
}