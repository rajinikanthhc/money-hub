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
    Utilities.formatDate(
    new Date(data.date),
    Session.getScriptTimeZone(),
    "dd-MMM-yyyy"
),
    data.type,
    data.category,
    data.account,
    data.amount,
    data.payment,
    data.description
  ]);

  return true;

}

function updateTransaction(data) {

  const sheet = SpreadsheetApp
    .openById("1IK0fyzoe5GnaPcYQ92dTNChpDSwlN8i951scM9W92TM")
    .getSheetByName("Transactions");

  sheet.getRange(data.row + 1, 2, 1, 7).setValues([[
    Utilities.formatDate(
      new Date(data.date),
      Session.getScriptTimeZone(),
      "dd-MMM-yyyy"
    ),
    data.type,
    data.category,
    data.account,
    data.amount,
    data.payment,
    data.description
  ]]);

  return true;

}

function deleteTransaction(row){

  const sheet = SpreadsheetApp
    .openById("1IK0fyzoe5GnaPcYQ92dTNChpDSwlN8i951scM9W92TM")
    .getSheetByName("Transactions");

  sheet.deleteRow(row + 1);

  return true;

}

function getAllCategories() {

  const sheet = SpreadsheetApp
    .openById("1IK0fyzoe5GnaPcYQ92dTNChpDSwlN8i951scM9W92TM")
    .getSheetByName("Categories");

  return sheet.getDataRange().getDisplayValues();

}

function saveCategory(data) {

  const sheet = SpreadsheetApp
    .openById("1IK0fyzoe5GnaPcYQ92dTNChpDSwlN8i951scM9W92TM")
    .getSheetByName("Categories");

  const id = sheet.getLastRow();

  sheet.appendRow([
    id,
    data.category,
    data.type
  ]);

  return true;

}

function updateCategory(data) {

  const sheet = SpreadsheetApp
    .openById("1IK0fyzoe5GnaPcYQ92dTNChpDSwlN8i951scM9W92TM")
    .getSheetByName("Categories");

  sheet.getRange(data.row + 1, 2, 1, 2).setValues([[
    data.category,
    data.type
  ]]);

  return true;

}

function deleteCategory(row) {

  const sheet = SpreadsheetApp
    .openById("1IK0fyzoe5GnaPcYQ92dTNChpDSwlN8i951scM9W92TM")
    .getSheetByName("Categories");

  sheet.deleteRow(row + 1);

  return true;

}

function getCategoriesByType(type) {

  const sheet = SpreadsheetApp
    .openById("1IK0fyzoe5GnaPcYQ92dTNChpDSwlN8i951scM9W92TM")
    .getSheetByName("Categories");

  const data = sheet.getDataRange().getDisplayValues();

  let list = [];

  for (let i = 1; i < data.length; i++) {

    if (data[i][2] == type) {

      list.push(data[i][1]);

    }

  }

  return list;

}

function getAllAccounts() {

  const sheet = SpreadsheetApp
    .openById("1IK0fyzoe5GnaPcYQ92dTNChpDSwlN8i951scM9W92TM")
    .getSheetByName("Accounts");

  return sheet.getDataRange().getDisplayValues();

}

function saveAccount(data) {

  const sheet = SpreadsheetApp
    .openById("1IK0fyzoe5GnaPcYQ92dTNChpDSwlN8i951scM9W92TM")
    .getSheetByName("Accounts");

  const id = sheet.getLastRow();

  sheet.appendRow([
    id,
    data.account,
    data.type,
    data.balance
  ]);

  return true;

}

function updateAccount(data) {

  const sheet = SpreadsheetApp
    .openById("1IK0fyzoe5GnaPcYQ92dTNChpDSwlN8i951scM9W92TM")
    .getSheetByName("Accounts");

  sheet.getRange(data.row + 1, 2, 1, 3).setValues([[
    data.account,
    data.type,
    data.balance
  ]]);

  return true;

}

function deleteAccount(row) {

  const sheet = SpreadsheetApp
    .openById("1IK0fyzoe5GnaPcYQ92dTNChpDSwlN8i951scM9W92TM")
    .getSheetByName("Accounts");

  sheet.deleteRow(row + 1);

  return true;

}