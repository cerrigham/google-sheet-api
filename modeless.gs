function modeless(title, message)
{
  var htmlOutput = HtmlService
    .createHtmlOutput('<font style="font-family: Helvetica;">'+message+'</font>')
    .setWidth(250)
    .setHeight(50);
  SpreadsheetApp.getUi().showModelessDialog(htmlOutput, title);
}

function modal(title, message)
{
  var htmlOutput = HtmlService
    .createHtmlOutput('<font style="font-family: Helvetica;">'+message+'</font>')
    .setWidth(250)
    .setHeight(50);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, title);
}
