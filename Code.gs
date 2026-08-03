function doGet() {
  const template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate().setTitle('Money Hub');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getTransactions() {
  return [
    ["Date","Type","Category","Account","Amount","Payment Mode","Description"],
    ["03-Aug-2026","Expense","Grocery","SBI",850,"UPI","D-Mart"]
  ];
}