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

    const type = data[i][2];          // Type
    const amount = Number(data[i][5]); // Amount

    if (type === "Income") {
      income += amount;
    } else if (type === "Expense") {
      expense += amount;
    }

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

function getPage(page) {

  return HtmlService
    .createHtmlOutputFromFile(page)
    .getContent();

}

function getAllTransactions() {

  const sheet = SpreadsheetApp
    .openById("1IK0fyzoe5GnaPcYQ92dTNChpDSwlN8i951scM9W92TM")
    .getSheetByName("Transactions");

  return sheet.getDataRange().getDisplayValues();

}

function saveTransaction(data) {

  const sheet = SpreadsheetApp
    .openById("1IK0fyzoe5GnaPcYQ92dTNChpDSwlN8i951scM9W92TM")
    .getSheetByName("Transactions");

  const id = sheet.getLastRow();

  sheet.appendRow([
    id,
    data.date,
    data.type,
    data.category,
    data.account,
    data.amount,
    data.payment,
    data.description
  ]);

  return true;

}