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

function getDashboardData() {

  const sheet = SpreadsheetApp
    .openById("1IK0fyzoe5GnaPcYQ92dTNChpDSwlN8i951scM9W92TM")
    .getSheetByName("Transactions");

  const data = sheet.getDataRange().getDisplayValues();

  let income = 0;
  let expense = 0;

  for (let i = 1; i < data.length; i++) {

    const type = data[i][1];
    const amount = Number(data[i][4]);

    if (type === "Income")
      income += amount;

    if (type === "Expense")
      expense += amount;
  }

  return {

    income: income,

    expense: expense,

    savings: income - expense,

    balance: income - expense

  };

}

function getRecentTransactions() {

  const sheet = SpreadsheetApp
    .openById("1IK0fyzoe5GnaPcYQ92dTNChpDSwlN8i951scM9W92TM")
    .getSheetByName("Transactions");

  const data = sheet.getDataRange().getDisplayValues();

  data.shift();

  return data.reverse().slice(0,5);

}