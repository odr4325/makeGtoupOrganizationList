/*
#####################################################
# ログ
# format: timestamp , priority , message
#####################################################
 */
function LogSheet(priority,message,flag) { //LOG関数
  log_(priority, message,flag);
}

function LogSheetdebug(message,flag) {
  log_('debug', message,flag);
}

function LogSheetinfo(message,flag) {
  log_('info', message,flag);
}

function LogSheetwarn(message,flag) {
  log_('warn', message,flag);
}

function LogSheeterror(message,flag) {
  log_('error', message,flag);
}

function LogSheetfatal(message,flag) {
  log_('fatal', message,flag);
}

function testlog(){
  LogSheet("INFO","testmessage");
}

function log_(priority, message, flag) {
  var sh = log_makesheet_();
  var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', "yyyy/MM/dd HH:mm:ss.SSS");
  var last_row = sh.getLastRow();
  sh.insertRowAfter(last_row).getRange(last_row+1, 1, 1, 3).setValues([[now, priority, message]]);
//  Browser.msgBox(sh);
  switch (flag){
    case 1:
      Logger.log("LogSheet: " + priority + ": " + message);
  }
  return sh;
}

function log_makesheet_() {
  var sheet_name = Cfg.logName;
//  var ss = SpreadsheetApp.openById(Cfg.sheetID);
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(sheet_name);
  
  if (sh == null) {
    var active_sh = ss.getActiveSheet(); // memorize current active sheet;
    var sheet_num = ss.getSheets().length;
    sh = ss.insertSheet(sheet_name, sheet_num);
//    sh.getRange('A1:C1').setValues([['Timestamp', 'priority', 'Message']]).setBackground('#cfe2f3').setFontWeight('bold');
    sh.getRange('A1:C1').setValues([['Timestamp', 'priority', 'Message']]);
    sh.getRange('A2:C2').setValues([[Utilities.formatDate(new Date(), 'Asia/Tokyo', "yyyy/MM/dd HH:mm:ss.SSS"), 'INFO', sheet_name + ' has been created.']]).clearFormat();

    // .insertSheet()を呼ぶと"log"シートがアクティブになるので、元々アクティブだったシートにフォーカスを戻す
    ss.setActiveSheet(active_sh);
  }
  return sh;
}

function LogSheetClear() { //LOGイニシャライズ
  var baseRows = 90;
  var leaveRows = 30;
  var sheet_name = 'log';
//  var ss = SpreadsheetApp.openById(Cfg.sheetID);
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(sheet_name);
  
  if (sh != null) {
    var lastR = sh.getLastRow();

    if ( baseRows <= lastR ) {
      sh.deleteRows(2, lastR - leaveRows);//　leave rows
    }
  }
//  if (sh != null) {
//    sh = ss.deleteSheet(sh);
//  }
  return sh;
}

/*
#####################################################
# FUNCTION
#####################################################
 */
function splitIndex(str, deli, end, checkStr) {
  if (!str) {
    return false;
  }
  
  //文字列をdelimiter毎に分けてend文字列以降のindexは削除する
  var retArr = new Array();
  var dataSet = false;
  var rowArr = str.split("\n");    //改行毎に配列へ

//  rowArr.splice(endIndex);                      //end以降の要素削除  arr[A,B,C,D]     arr[A/B/C,D,E/F]
  for(var i in rowArr){
    if ( rowArr[i].indexOf(checkStr) != -1) {
      retArr.push(rowArr[i].split(deli));       //デリミタ毎に分割 arr[A|B|C|D|end] arr[A/B/C|D|E/F|end]
      dataSet = true;
    }
  }
  
  if ( dataSet == true ) {
    return retArr;
  }else{
    return false;
  }
}

function split(str, deli, end) {
  var topPosition = 0;
  var characterOf = 0;
  var endstr = true;
  var index = 0;
  var arr = []
  var deliLength = deli.length;

  //strからdeliが検索されなくなるまでループする
  L1:while((characterOf = str.indexOf(deli, topPosition)) != -1) {
    if (str.substring(topPosition, characterOf) == end) {
      endstr = false;
      break L1;
    } 
    arr[index] = str.substring(topPosition, characterOf);
    topPosition = characterOf + deliLength;
    index++;
  }
  
  if (endstr) {
    arr[index] = str.substring(topPosition);
  }
  
  return arr;
}

