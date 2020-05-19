var budgetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('予算シート');
const budgetUsingColumn = 3;
const budgetUsingRows = 6;

//ローカル変数作成
function init() {
  var maxCol = budgetSheet.getRange(budgetSheet.getMaxRows(), 1).getNextDataCell(SpreadsheetApp.Direction.UP);
  this.lastCol = maxCol.getRow();
  this.lastMonth = new Date(Math.floor(maxCol.getValue() / 100), ("0" + maxCol.getValue() % 100).slice(-2), 1)
  this.nowMonth = getLatestYyMm();
}

//main
function makeBudget(){
  init();
  copyLastToLatest();
}

//当月yyyyMM作成
function getLatestYyMm() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMM');
}

//先月予算をコピーして当月分とする
function copyLastToLatest() {
  var orgStartRow = this.lastCol - budgetUsingRows + 1;
  var tgtStartRow = this.lastCol + 1 ;
  
  var initRange = budgetSheet.getRange(tgtStartRow, 1, budgetUsingRows, 1);
  initRange.setValue(this.nowMonth);
  
  var orgRange = budgetSheet.getRange(orgStartRow, 2, budgetUsingRows, budgetUsingColumn);
  var targetRange = budgetSheet.getRange(tgtStartRow, 2, budgetUsingRows, budgetUsingColumn);
  orgRange.copyTo(targetRange);
}