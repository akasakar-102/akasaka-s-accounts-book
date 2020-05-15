//main
function mailScraping(){
  var afterDate = targetDateMaker();
  const SEARCH_TERM = 'subject:([おカネレコ]エクスポート) after:' + afterDate + ' has:attachment';
  fetchFile(SEARCH_TERM);
}

//gmail検索クエリ用当日日付
function targetDateMaker() {
  var date = new Date();
  var today = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd');
  return today;
}

//gmailの添付ファイルからExpotCSVを取得
function fetchFile(SEARCH_TERM){
  const threads = GmailApp.search(SEARCH_TERM, 0, 10);
  const messages = GmailApp.getMessagesForThreads(threads);
  for(const thread of messages){
    for(const message of thread){
      const attachments = message.getAttachments();
      for(const attachment of attachments){
        var attachName = (attachment.getName()).split(/[_.]/);
        if(attachName.length == 4 && attachName[0] == "Export" && attachName[3] == "csv"){
          addToFolder(attachment);
        }
      }
    }
  }
}

//Drive内に同名ファイルが存在したら Del&Ins
function addToFolder(attachment){
  const folder = DriveApp.getFolderById("1Rp8DTrwZV78VlCMtvSBDxLLsd0woxCVU");
  var filesList = folder.getFiles();
  while(filesList.hasNext()){
    var existFile = filesList.next();
    if(existFile.getName() == attachment.getName()){
      folder.removeFile(existFile);
    }
  }
  folder.createFile(attachment);
}