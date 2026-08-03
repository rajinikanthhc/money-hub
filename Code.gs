function doGet() {
  const template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate().setTitle('Money Hub');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getTransactions() {
  const sheet = SpreadsheetApp
    .openById("1IK0fyzoe5GnaPcYQ92dTNChpDSwlN8i951scM9W92TM")
    .getSheetByName("Transactions");

  return sheet.getDataRange().getValues();
}