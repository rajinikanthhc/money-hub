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

  const ss = SpreadsheetApp.openById("1IK0fyzoe5GnaPcYQ92dTNChpDSwlN8i951scM9W92TM");

  // Transactions
  const transSheet = ss.getSheetByName("Transactions");
  const transData = transSheet.getDataRange().getDisplayValues();

  // Accounts
  const accountSheet = ss.getSheetByName("Accounts");
  const accountData = accountSheet.getDataRange().getDisplayValues();

  let income = 0;
  let expense = 0;
  let openingBalance = 0;

  // Calculate Income & Expense
  for (let i = 1; i < transData.length; i++) {

    const type = transData[i][2];
    const amount = Number(transData[i][5]);

    if (type === "Income") {
      income += amount;
    } else if (type === "Expense") {
      expense += amount;
    }

  }

  // Calculate Opening Balance
  for (let i = 1; i < accountData.length; i++) {

    openingBalance += Number(accountData[i][3]) || 0;

  }

  return {

    income: income,
    expense: expense,
    savings: income - expense,
    balance: openingBalance + income - expense

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

  const ss = SpreadsheetApp.openById("1IK0fyzoe5GnaPcYQ92dTNChpDSwlN8i951scM9W92TM");

  const accountSheet = ss.getSheetByName("Accounts");
  const transSheet = ss.getSheetByName("Transactions");

  const accounts = accountSheet.getDataRange().getDisplayValues();
  const transactions = transSheet.getDataRange().getDisplayValues();

  let result = [];

  for (let i = 1; i < accounts.length; i++) {

    let currentBalance = Number(accounts[i][3]) || 0;

    const accountName = accounts[i][1];

    for (let j = 1; j < transactions.length; j++) {

      if (transactions[j][4] == accountName) {

        const amount = Number(transactions[j][5]);

        if (transactions[j][2] == "Income") {

          currentBalance += amount;

        } else if (transactions[j][2] == "Expense") {

          currentBalance -= amount;

        }

      }

    }

    result.push([

      accounts[i][0],        // ID
      accountName,           // Account
      accounts[i][2],        // Type
      accounts[i][3],        // Opening Balance
      currentBalance         // Current Balance

    ]);

  }

  return result;

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

function getAllAccountNames() {

  const sheet = SpreadsheetApp
    .openById("1IK0fyzoe5GnaPcYQ92dTNChpDSwlN8i951scM9W92TM")
    .getSheetByName("Accounts");

  const data = sheet.getDataRange().getDisplayValues();

  let list = [];

  for (let i = 1; i < data.length; i++) {

    list.push(data[i][1]);

  }

  return list;

}

function getMonthlyBudget() {

  const ss = SpreadsheetApp.openById("1IK0fyzoe5GnaPcYQ92dTNChpDSwlN8i951scM9W92TM");

  const budgetSheet = ss.getSheetByName("Monthly Budget");
  const categorySheet = ss.getSheetByName("Categories");
  const transactionSheet = ss.getSheetByName("Transactions");

  const budgets = budgetSheet.getDataRange().getDisplayValues();
  const categories = categorySheet.getDataRange().getDisplayValues();
  const transactions = transactionSheet.getDataRange().getDisplayValues();

  // Category -> Budget Group
  let categoryMap = {};

  for (let i = 1; i < categories.length; i++) {

    categoryMap[categories[i][1]] = categories[i][3];

  }

  let result = [];

  for (let i = 1; i < budgets.length; i++) {

    const group = budgets[i][1];
    const budget = Number(budgets[i][2]);

    let actual = 0;

    for (let j = 1; j < transactions.length; j++) {

      const category = transactions[j][3];
      const amount = Number(transactions[j][5]);

      if (categoryMap[category] == group) {

        actual += amount;

      }

    }

    result.push({

      group: group,
      budget: budget,
      actual: actual,
      remaining: budget - actual

    });

  }

  return result;

}

function getReportData() {

  const ss = SpreadsheetApp.openById("1IK0fyzoe5GnaPcYQ92dTNChpDSwlN8i951scM9W92TM");

  const transSheet = ss.getSheetByName("Transactions");
  const catSheet = ss.getSheetByName("Categories");

  const transactions = transSheet.getDataRange().getDisplayValues();
  const categories = catSheet.getDataRange().getDisplayValues();

  // Category -> Budget Group
  let categoryMap = {};

  for (let i = 1; i < categories.length; i++) {
    categoryMap[categories[i][1]] = categories[i][3];
  }

  let income = 0;
  let expense = 0;

  let budgetSummary = {};
  let categorySummary = {};

  for (let i = 1; i < transactions.length; i++) {

    const type = transactions[i][2];
    const category = transactions[i][3];
    const amount = Number(transactions[i][5]);

    if (type == "Income") {

      income += amount;

    } else {

      expense += amount;

      // Category Summary
      categorySummary[category] =
        (categorySummary[category] || 0) + amount;

      // Budget Group Summary
      const group = categoryMap[category] || "Others";

      budgetSummary[group] =
        (budgetSummary[group] || 0) + amount;
    }

  }

  return {

    income: income,
    expense: expense,
    savings: income - expense,
    balance: income - expense,

    budgetSummary: budgetSummary,
    categorySummary: categorySummary

  };

}

function drawBudgetChart(data){

    const chartData = [["Budget Group","Amount"]];

    for(let group in data.budgetSummary){

        chartData.push([group,data.budgetSummary[group]]);

    }

    const chart = new google.visualization.PieChart(
        document.getElementById("budgetChart")
    );

    chart.draw(
        google.visualization.arrayToDataTable(chartData),
        {
            legend:{position:"right"},
            pieHole:0.4
        }
    );

}

function drawCategoryChart(data){

    const chartData = [["Category","Amount"]];

    for(let cat in data.categorySummary){

        chartData.push([cat,data.categorySummary[cat]]);

    }

    const chart = new google.visualization.PieChart(
        document.getElementById("categoryChart")
    );

    chart.draw(
        google.visualization.arrayToDataTable(chartData),
        {
            legend:{position:"right"}
        }
    );

}

function drawIncomeExpenseChart(data){

    const chart = new google.visualization.ColumnChart(
        document.getElementById("incomeExpenseChart")
    );

    chart.draw(

        google.visualization.arrayToDataTable([

            ["Type","Amount"],

            ["Income",data.income],

            ["Expense",data.expense]

        ]),

        {

            legend:"none"

        }

    );

}