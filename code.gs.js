let ui = SpreadsheetApp.getUi();
let sheet = SpreadsheetApp.getActiveSheet();
let lastRow = sheet.getLastRow();
let lastColumn = sheet.getLastColumn();

let row;

function onInstall(e) {
  onOpen(e);
}

function onOpen(e) {
  ui.createAddonMenu().addItem("Predict", "showSidebar").addToUi();
}

function showSidebar() {
  const html = HtmlService.createTemplateFromFile("index").evaluate();
  html.setTitle("Quilt AI");
  ui.showSidebar(html);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getColumnHeader() {
  let data = SpreadsheetApp.getActiveSheet().getDataRange().getValues();
  const header = data[0].filter(Boolean);
  return header;
}

function getByName(columnName) {
  let columnValue = [];

  let data = SpreadsheetApp.getActiveSheet().getDataRange().getValues();
  let columnIndex = data[0].indexOf(columnName);

  for (row = 1; row < data.length; ++row) {
    columnValue.push(data[row][columnIndex]);
    // if (data[row][columnIndex] != "") {
    //   columnValue.push(data[row][columnIndex]);
    // }
  }
  return columnValue;
}

function writeHeader(last) {
  let newHeader = ["Label", "Score", "Theme"];
  sheet
    .getRange(1, last + 1, 1, newHeader.length)
    .setValues([newHeader])
    .setBackground("#9b4dca")
    .setFontColor("#ffffff");
}

function lastColumnValue() {
  let last = sheet.getLastColumn();
  return last;
}

function writeDataToRange(data, rowCount, counter, lastCol) {
  row = counter + 1;
  sheet
    .getRange(row, lastCol + 1, data.length, 3)
    .setValues(data)
    .setBackground("#f4f5f6");
}
