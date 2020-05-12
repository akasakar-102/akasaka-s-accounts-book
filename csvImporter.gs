var spreadSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('data');
var usingColumn = 9;

function onOpen() {
  var ui = SpreadsheetApp.getUi();           // Uiクラスを取得する
  var menu = ui.createMenu('家計簿ツール');  // Uiクラスからメニューを作成する
  menu.addItem('月を指定してcsv取込', 'monthSelect');   // メニューにアイテムを追加する
  menu.addItem('最新月のcsv取込', 'autoImport');   // メニューにアイテムを追加する
  menu.addToUi();                            // メニューをUiクラスに追加する
}

//月指定ダイアログ
function monthSelect(){
  var ui = SpreadsheetApp.getUi();
  var message = '対象月を選択してください'
  var res = Browser.inputBox(message);
  doImport(res);
}

//最新月データの取り込み
function autoImport() {
  var yyyyMm = getLatestMonthFromDrive();
  doImport(yyyyMm);
}

//ドライブ所定のフォルダ以下の最新月データを定義
function getLatestMonthFromDrive() {
  //csv格納フォルダから最新月のものを取得 ファイル名は連携元アプリの仕様により「Export_yyyy_mm.csv」固定
  var folder = DriveApp.getFolderById("1Rp8DTrwZV78VlCMtvSBDxLLsd0woxCVU");
  var fileList = folder.getFiles();
  var latestMonth = 0;
  while (fileList.hasNext()) {
    var file = fileList.next();
    var fileName = file.getName().split(/[_.]/);
    if (fileName.length == 4 && fileName[0] == "Export" && fileName[3] == "csv"){
      var monthInName = parseInt(fileName[1] + fileName[2]);
      if (monthInName > latestMonth){
        latestMonth = monthInName;
      }
    }
  }
  return latestMonth.toString(10);
}

//インポートロジック本体
function doImport(yearMonth) {
  var startRow = 2;
  var lastRow = spreadSheet.getLastRow();
  var targetYear ="" + Math.floor(yearMonth / 100);
  var targetMonth =("0" + yearMonth % 100).slice(-2);
  var csvName = "Export_" + targetYear + "_" + targetMonth + ".csv";

  //csv格納フォルダ
  var folder = DriveApp.getFolderById("1Rp8DTrwZV78VlCMtvSBDxLLsd0woxCVU");
  var file = folder.getFilesByName(csvName);
  if (file.hasNext()){  
    var csvText = file.next().getBlob().getDataAsString("UTF-8"); 
    var csv = Utilities.parseCsv(csvText);
    //手入力されたデータは逃がして最後にマージ
    var manualInputData = getManualData(startRow, lastRow);
    var orgCsv = organizeCsv(csv);
    var importData = mergeData(orgCsv, manualInputData);

    //白紙化（Todo:書き足し対応）
    spreadSheet.getRange(startRow, 1, lastRow, importData[0].length).clear();

    spreadSheet.getRange(startRow,1,importData.length,importData[0].length).setValues(importData).sort([{column:2,ascending:true}]);
    return;
  }
}

//対象月に手入力データがある場合に取得してマージ
function getManualData(startRow, lastRow) {
  var mResult = [];
  var importFlgIndex = 8;
  var rec = spreadSheet.getRange(startRow, 1, lastRow, usingColumn).getValues();
  for(i=0; i < rec.length; i++){
    if(rec[i][importFlgIndex] != 1){
      mResult.push(rec[i]);
    }
  }
  　return mResult;
}

//元csvデータをgss表示用に再構築
function organizeCsv(csv){
  var dateColumnIndex = 0, categoryColumnIndex = 0, priceColumnIndex = 0, memoColumnIndex = 0, shopNameColumnIndex = 0;
  var date, category, price, memo, shopName, resultMonth, resultDate;
  var resultOneRecord = [];
  var result = [];
  for(i=0; csv.length > i; i++){
    var splitData  = csv[i];
    if(splitData[0] == 0){
      continue;
    }
    //ヘッダーを見て各Index決定
    if(i == 0){
      for(j=0; splitData.length > j; j++){
        var columnName = splitData[j];
        if(columnName == "DATE"){
          dateColumnIndex = j;
          continue;
        }
        else if(columnName == "CATEGORY"){
               categoryColumnIndex = j;
               continue;
        }
        else if(columnName == "PRICE"){
               priceColumnIndex = j;
               continue;
        }
        else if(columnName == "MEMO"){
               memoColumnIndex = j;
               continue;
        }
        else if(columnName == "SHOP NAME"){
               shopNameColumnIndex = j;
               continue;
        }
      }
      continue;
    }
    //Bodyから必要なデータのみ取得
    date = splitData[dateColumnIndex];
    category = splitData[categoryColumnIndex];
    price = splitData[priceColumnIndex];
    memo = splitData[memoColumnIndex];
    shopName = splitData[shopNameColumnIndex];
    var resultMonth = makeMonth(date);
    var resultDate = makeDate(date);
    var resultItem = makeItem(memo);
    var resultMemo = makeMemo(memo);
    var importFlg = 1;
    
    resultOneRecord = [resultMonth, resultDate, category, resultItem, price, shopName, resultMemo,"" ,importFlg];
    result.push(resultOneRecord);
  }
  return result;
}
  
//csvのdateからyyyymm作成
function makeMonth(date){
  var year = date.split('/')[0];
  var month = date.split('/')[1];
  return year + month;
}    

//csvのdateからmm/dd作成
function makeDate(date){
  var month = date.split('/')[1];
  var day = date.split('/')[2];
  console.log(month);
  console.log(day);
  return month + "/" + day;
}  
  
//csvのmemoから品名作成
function makeItem(memo){
  return memo.split('$')[0];
}  

//csvのmemoから備考欄作成
function makeMemo(memo){
  var memos = memo.split('$');
  if(memos.length != 1){
    return memo.split('$')[1];
  }
  return "";
}

//csv成形データと手入力データを結合
function mergeData(orgData, manualInputData){
  for(i =0; i < manualInputData.length; i++){
    orgData.push(manualInputData[i]);
  }
  return orgData;
}