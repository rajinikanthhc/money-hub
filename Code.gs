/* =========================================
   MONEY HUB - BACKEND
========================================= */

const SS = SpreadsheetApp.getActiveSpreadsheet();

const SHEETS = {
  ACCOUNTS: "Accounts",
  TRANSACTIONS: "Transactions",
  BUDGET: "Monthly Budget",
  CATEGORIES: "Categories"
};


/* =========================================
   WEB APP
========================================= */

function doGet() {

  return HtmlService
    .createTemplateFromFile("Index")
    .evaluate()
    .setTitle("Money Hub")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

}


/* =========================================
   INCLUDE FILES
========================================= */

function include(filename) {

  return HtmlService
    .createHtmlOutputFromFile(filename)
    .getContent();

}


/* =========================================
   LOAD ALL DATA
========================================= */

function getInitialData() {

  return {
    accounts: getAccounts(),
    transactions: getTransactions(),
    categories: getCategories(),
    budgets: getBudgets()
  };

}


/* =========================================
   ACCOUNTS
========================================= */

function getAccounts() {

  const sheet = SS.getSheetByName(SHEETS.ACCOUNTS);

  if (!sheet) {
    throw new Error("Accounts sheet not found.");
  }

  const values = sheet.getDataRange().getValues();

  if (values.length <= 1) {
    return [];
  }

  return values
    .slice(1)
    .filter(row => row[0] !== "")
    .map(row => ({
      id: row[0],
      name: row[1],
      type: row[2],
      openingBalance: Number(row[3]) || 0
    }));

}


/* =========================================
   CATEGORIES
========================================= */

function getCategories() {

  const sheet = SS.getSheetByName(SHEETS.CATEGORIES);

  if (!sheet) {
    throw new Error("Categories sheet not found.");
  }

  const values = sheet.getDataRange().getValues();

  if (values.length <= 1) {
    return [];
  }

  return values
    .slice(1)
    .filter(row => row[0] !== "")
    .map(row => ({
      id: row[0],
      name: row[1],
      type: row[2],
      budgetGroup: row[3]
    }));

}


/* =========================================
   BUDGETS
========================================= */

function getBudgets() {

  const sheet = SS.getSheetByName(SHEETS.BUDGET);

  if (!sheet) {
    throw new Error("Monthly Budget sheet not found.");
  }

  const values = sheet.getDataRange().getValues();

  if (values.length <= 1) {
    return [];
  }

  return values
    .slice(1)
    .filter(row => row[0] !== "")
    .map(row => ({
      id: row[0],
      category: row[1],
      monthlyBudget: Number(row[2]) || 0,
      remarks: row[3] || ""
    }));

}


/* =========================================
   TRANSACTIONS
========================================= */

function getTransactions() {

  const sheet = SS.getSheetByName(SHEETS.TRANSACTIONS);

  if (!sheet) {
    throw new Error("Transactions sheet not found.");
  }

  const values = sheet.getDataRange().getValues();

  if (values.length <= 1) {
    return [];
  }

  const result = values
    .slice(1)
    .filter(row => row[0] !== "")
    .map(row => {

      let date = row[1];

      if (date instanceof Date) {

        date = Utilities.formatDate(
          date,
          Session.getScriptTimeZone(),
          "yyyy-MM-dd"
        );

      } else {

        date = String(date);

      }

      return {

        id: row[0],

        date: date,

        type: row[2] || "",

        category: row[3] || "",

        fromAccount: row[4] || "",

        toAccount: row[5] || "",

        amount: Number(row[6]) || 0,

        description: row[7] || ""

      };

    });


  // Latest transaction first
  result.sort(function(a, b) {

    const dateA = new Date(a.date);
    const dateB = new Date(b.date);

    if (dateB - dateA !== 0) {
      return dateB - dateA;
    }

    // If dates are same, higher ID first
    return Number(b.id) - Number(a.id);

  });


  return result;

}

/* =========================================
   ADD TRANSACTION
========================================= */

function addTransaction(data) {

  validateTransaction(data);

  const sheet = SS.getSheetByName(SHEETS.TRANSACTIONS);

  if (!sheet) {
    throw new Error("Transactions sheet not found.");
  }

  const newId = getNextTransactionId(sheet);

  const date = parseDate(data.date);

  sheet.appendRow([

    newId,

    date,

    data.type,

    data.category || "",

    data.fromAccount || "",

    data.toAccount || "",

    Number(data.amount),

    data.description || ""

  ]);

  return {

    success: true,
    id: newId

  };

}


/* =========================================
   UPDATE TRANSACTION
========================================= */

function updateTransaction(data) {

  validateTransaction(data);

  const sheet = SS.getSheetByName(SHEETS.TRANSACTIONS);

  if (!sheet) {
    throw new Error("Transactions sheet not found.");
  }

  const values = sheet.getDataRange().getValues();

  let rowNumber = -1;

  for (let i = 1; i < values.length; i++) {

    if (String(values[i][0]) === String(data.id)) {

      rowNumber = i + 1;
      break;

    }

  }

  if (rowNumber === -1) {

    throw new Error("Transaction not found.");

  }

  sheet.getRange(rowNumber, 1, 1, 8).setValues([[

    data.id,

    parseDate(data.date),

    data.type,

    data.category || "",

    data.fromAccount || "",

    data.toAccount || "",

    Number(data.amount),

    data.description || ""

  ]]);

  return {

    success: true

  };

}


/* =========================================
   DELETE TRANSACTION
========================================= */

function deleteTransaction(id) {

  const sheet = SS.getSheetByName(SHEETS.TRANSACTIONS);

  if (!sheet) {
    throw new Error("Transactions sheet not found.");
  }

  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {

    if (String(values[i][0]) === String(id)) {

      sheet.deleteRow(i + 1);

      return {

        success: true

      };

    }

  }

  throw new Error("Transaction not found.");

}


/* =========================================
   NEXT TRANSACTION ID
========================================= */

function getNextTransactionId(sheet) {

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return 1;
  }

  const ids = sheet
    .getRange(2, 1, lastRow - 1, 1)
    .getValues()
    .flat()
    .map(Number)
    .filter(n => !isNaN(n));

  if (ids.length === 0) {
    return 1;
  }

  return Math.max.apply(null, ids) + 1;

}


/* =========================================
   VALIDATION
========================================= */

function validateTransaction(data) {

  if (!data.date) {

    throw new Error("Please select a date.");

  }

  if (!data.type) {

    throw new Error("Please select transaction type.");

  }

  if (!data.amount || Number(data.amount) <= 0) {

    throw new Error("Please enter a valid amount.");

  }

  if (data.type === "Expense") {

    if (!data.category) {
      throw new Error("Please select a category.");
    }

    if (!data.fromAccount) {
      throw new Error("Please select From Account.");
    }

  }

  if (data.type === "Income") {

    if (!data.category) {
      throw new Error("Please select a category.");
    }

    if (!data.toAccount) {
      throw new Error("Please select To Account.");
    }

  }

  if (data.type === "Transfer") {

    if (!data.fromAccount) {
      throw new Error("Please select From Account.");
    }

    if (!data.toAccount) {
      throw new Error("Please select To Account.");
    }

    if (data.fromAccount === data.toAccount) {

      throw new Error(
        "From Account and To Account cannot be the same."
      );

    }

  }

}


/* =========================================
   DATE
========================================= */

function parseDate(dateString) {

  if (dateString instanceof Date) {
    return dateString;
  }

  const parts = String(dateString).split("-");

  if (parts.length !== 3) {

    throw new Error("Invalid date.");

  }

  return new Date(
    Number(parts[0]),
    Number(parts[1]) - 1,
    Number(parts[2])
  );

}

/* =========================================
   ACCOUNTS - ADD / UPDATE / DELETE
========================================= */

function addAccount(data) {

  const sheet = SS.getSheetByName(SHEETS.ACCOUNTS);

  if (!sheet) {
    throw new Error("Accounts sheet not found.");
  }

  if (!data.name) {
    throw new Error("Please enter account name.");
  }

  if (!data.type) {
    throw new Error("Please select account type.");
  }

  const values = sheet.getDataRange().getValues();

  const duplicate = values.slice(1).some(function(row) {

    return String(row[1]).trim().toLowerCase() ===
           String(data.name).trim().toLowerCase();

  });

  if (duplicate) {
    throw new Error("Account already exists.");
  }

  const newId = getNextId(sheet);

  sheet.appendRow([
    newId,
    data.name.trim(),
    data.type,
    Number(data.openingBalance) || 0
  ]);

  return { success: true };
}


function updateAccount(data) {

  const sheet = SS.getSheetByName(SHEETS.ACCOUNTS);

  if (!sheet) {
    throw new Error("Accounts sheet not found.");
  }

  if (!data.name) {
    throw new Error("Please enter account name.");
  }

  if (!data.type) {
    throw new Error("Please select account type.");
  }

  const values = sheet.getDataRange().getValues();

  const duplicate = values.slice(1).some(function(row) {

    return String(row[0]) !== String(data.id) &&
           String(row[1]).trim().toLowerCase() ===
           String(data.name).trim().toLowerCase();

  });

  if (duplicate) {
    throw new Error("Account already exists.");
  }

  for (let i = 1; i < values.length; i++) {

    if (String(values[i][0]) === String(data.id)) {

      sheet.getRange(i + 1, 1, 1, 4).setValues([[
        data.id,
        data.name.trim(),
        data.type,
        Number(data.openingBalance) || 0
      ]]);

      return { success: true };
    }

  }

  throw new Error("Account not found.");
}


function deleteAccount(id) {

  const sheet = SS.getSheetByName(SHEETS.ACCOUNTS);

  if (!sheet) {
    throw new Error("Accounts sheet not found.");
  }

  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {

    if (String(values[i][0]) === String(id)) {

      sheet.deleteRow(i + 1);

      return { success: true };
    }

  }

  throw new Error("Account not found.");
}


/* =========================================
   CATEGORIES - ADD / UPDATE / DELETE
========================================= */

function addCategory(data) {

  const sheet = SS.getSheetByName(SHEETS.CATEGORIES);

  if (!sheet) {
    throw new Error("Categories sheet not found.");
  }

  if (!data.name) {
    throw new Error("Please enter category name.");
  }

  if (!data.type) {
    throw new Error("Please select category type.");
  }

  if (!data.budgetGroup) {
    throw new Error("Please enter budget group.");
  }

  const values = sheet.getDataRange().getValues();

  const duplicate = values.slice(1).some(function(row) {

    return String(row[1]).trim().toLowerCase() ===
           String(data.name).trim().toLowerCase();

  });

  if (duplicate) {
    throw new Error("Category already exists.");
  }

  const newId = getNextId(sheet);

  sheet.appendRow([
    newId,
    data.name.trim(),
    data.type,
    data.budgetGroup.trim()
  ]);

  return { success: true };
}


function updateCategory(data) {

  const sheet = SS.getSheetByName(SHEETS.CATEGORIES);

  if (!sheet) {
    throw new Error("Categories sheet not found.");
  }

  if (!data.name) {
    throw new Error("Please enter category name.");
  }

  if (!data.type) {
    throw new Error("Please select category type.");
  }

  if (!data.budgetGroup) {
    throw new Error("Please enter budget group.");
  }

  const values = sheet.getDataRange().getValues();

  const duplicate = values.slice(1).some(function(row) {

    return String(row[0]) !== String(data.id) &&
           String(row[1]).trim().toLowerCase() ===
           String(data.name).trim().toLowerCase();

  });

  if (duplicate) {
    throw new Error("Category already exists.");
  }

  for (let i = 1; i < values.length; i++) {

    if (String(values[i][0]) === String(data.id)) {

      sheet.getRange(i + 1, 1, 1, 4).setValues([[
        data.id,
        data.name.trim(),
        data.type,
        data.budgetGroup.trim()
      ]]);

      return { success: true };
    }

  }

  throw new Error("Category not found.");
}


function deleteCategory(id) {

  const sheet = SS.getSheetByName(SHEETS.CATEGORIES);

  if (!sheet) {
    throw new Error("Categories sheet not found.");
  }

  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {

    if (String(values[i][0]) === String(id)) {

      sheet.deleteRow(i + 1);

      return { success: true };
    }

  }

  throw new Error("Category not found.");
}


/* =========================================
   BUDGET - ADD / UPDATE / DELETE
========================================= */

function addBudget(data) {

  const sheet = SS.getSheetByName(SHEETS.BUDGET);

  if (!sheet) {
    throw new Error("Monthly Budget sheet not found.");
  }

  if (!data.category) {
    throw new Error("Please enter budget category.");
  }

  if (Number(data.monthlyBudget) < 0) {
    throw new Error("Please enter a valid budget.");
  }

  const values = sheet.getDataRange().getValues();

  const duplicate = values.slice(1).some(function(row) {

    return String(row[1]).trim().toLowerCase() ===
           String(data.category).trim().toLowerCase();

  });

  if (duplicate) {
    throw new Error("Budget category already exists.");
  }

  const newId = getNextId(sheet);

  sheet.appendRow([
    newId,
    data.category.trim(),
    Number(data.monthlyBudget) || 0,
    data.remarks ? data.remarks.trim() : ""
  ]);

  return { success: true };
}


function updateBudget(data) {

  const sheet = SS.getSheetByName(SHEETS.BUDGET);

  if (!sheet) {
    throw new Error("Monthly Budget sheet not found.");
  }

  if (!data.category) {
    throw new Error("Please enter budget category.");
  }

  if (Number(data.monthlyBudget) < 0) {
    throw new Error("Please enter a valid budget.");
  }

  const values = sheet.getDataRange().getValues();

  const duplicate = values.slice(1).some(function(row) {

    return String(row[0]) !== String(data.id) &&
           String(row[1]).trim().toLowerCase() ===
           String(data.category).trim().toLowerCase();

  });

  if (duplicate) {
    throw new Error("Budget category already exists.");
  }

  for (let i = 1; i < values.length; i++) {

    if (String(values[i][0]) === String(data.id)) {

      sheet.getRange(i + 1, 1, 1, 4).setValues([[
        data.id,
        data.category.trim(),
        Number(data.monthlyBudget) || 0,
        data.remarks ? data.remarks.trim() : ""
      ]]);

      return { success: true };
    }

  }

  throw new Error("Budget not found.");
}


function deleteBudget(id) {

  const sheet = SS.getSheetByName(SHEETS.BUDGET);

  if (!sheet) {
    throw new Error("Monthly Budget sheet not found.");
  }

  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {

    if (String(values[i][0]) === String(id)) {

      sheet.deleteRow(i + 1);

      return { success: true };
    }

  }

  throw new Error("Budget not found.");
}


/* =========================================
   GENERIC NEXT ID
========================================= */

function getNextId(sheet) {

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return 1;
  }

  const ids = sheet
    .getRange(2, 1, lastRow - 1, 1)
    .getValues()
    .flat()
    .map(Number)
    .filter(function(n) {
      return !isNaN(n);
    });

  if (ids.length === 0) {
    return 1;
  }

  return Math.max.apply(null, ids) + 1;
}