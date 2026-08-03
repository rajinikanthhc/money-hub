function doGet() {
  const template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate().setTitle('Money Hub');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getTransactions() {

  try {

    const ss = SpreadsheetApp.openById("1IK0fyzoe5GnaPcYQ92dTNChpDSwlN8i951scM9W92TM");

    const sheet = ss.getSheetByName("Transactions");

    if (!sheet) {
      throw new Error("Sheet 'Transactions' not found.");
    }

    return sheet.getDataRange().getDisplayValues();

  } catch (e) {

    throw new Error(e.message);

  }

}