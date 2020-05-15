var budgetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('予算シート');
var usingColumn = 3;
var usingRows = 6;

function init() {
  var maxCol = budgetSheet.getRange(budgetSheet.getMaxRows(), 1).getNextDataCell(SpreadsheetApp.Direction.UP);
  this.lastCol = maxCol.getRow();
  this.lastMonth = new Date(Math.floor(maxCol.getValue() / 100), ("0" + maxCol.getValue() % 100).slice(-2), 1)
  this.nowMonth = getLatestYyMm();
}

function makeBudget(){
  init();
  copyLastToLatest();
}

function getLatestYyMm() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMM');
}

function copyLastToLatest() {
  var orgStartRow = this.lastCol - usingRows + 1;
  var tgtStartRow = this.lastCol + 1 ;
  
  var initRange = budgetSheet.getRange(tgtStartRow, 1, usingRows, 1);
  initRange.setValue(this.nowMonth);
  
  var orgRange = budgetSheet.getRange(orgStartRow, 2, usingRows, usingColumn);
  var targetRange = budgetSheet.getRange(tgtStartRow, 2, usingRows, usingColumn);
  orgRange.copyTo(targetRange);
}