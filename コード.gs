/*
@author : kawagoe
@make   : 2017/10/12
@update : 2017/11/07 kawagoe 
*/

var ext = 5;
var Cfg = {
  logName : "log",
  getSheetName_s : "組織データ",
  getSheetName_j : "人事データ",
  inputCol_s : 6,   //入力列　組織ツリー用(列番号)
  inputCol_j : 7,   //入力列　アドレス用(列番号)
  
  childCol_s   : 0,  //自(子)コード列　（配列表記0から数えて)
  parentCol_s  : 1,  //親コード列　（配列表記0から数えて)
  addressCol_s : 2,  //メンバーメールアドレス列（配列表記0から数えて)
  
  parentCol_j  : 3,  //親コード列　（配列表記0から数えて)
  parentCnt_j  : ext - 3,  //親コードチェック列数（先頭3列より後の配列分チェック）
  addressCol_j : 0,  //メンバーメールアドレス列（配列表記0から数えて)
  
  stRow_s : 3,
  enCol_s : 4,
  stRow_j : 2,
  enCol_j : ext
};

function onOpen(){
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('CustomMenu')
  .addItem('子組織格納(アド/カレ用組織)', 'fn_storesSosiki')
  .addItem('メンバー格納(アド/カレ用組織)', 'fn_storesMember')
  .addToUi();
}

function fn_storesSosiki(){
  var startTime = new Date();
  LogSheet("INFO", Cfg.getSheetName_s + " : 組織追加処理開始")
  try{
    var member = "";
    var sheetName_s = Cfg.getSheetName_s;
    var ss = SpreadsheetApp.getActive().getSheetByName(sheetName_s);
    if ( ss == null ) {
      SpreadsheetApp.getActive().toast(sheetName_s + " シートが見つかりません","ERROR",5);
      return;
    }
    
    var enRow = ss.getLastRow() - Cfg.stRow_s + 1;
    if ( enRow < 1 ) {
      SpreadsheetApp.getActive().toast("データが見つかりません","ERROR",5);
      LogSheet("ERROR","データが見つかりません");
      return;
    }
    var datas = ss.getRange(Cfg.stRow_s, 1, enRow, Cfg.enCol_s).getValues();
    
    //親処理：一意のコード別に処理
    for ( var i in datas ) {
      if ( datas[i][Cfg.childCol_s] == "" ) {
        continue;
      }
      //子処理：親コード別に処理（該当コードが親コードと一致したら親コード行のメールを配列化）
      var memberCnt = 0;
      for ( var j in datas ) {
        if ( datas[i][Cfg.childCol_s]　== datas[j][Cfg.parentCol_s] ) {
          member = member + datas[j][Cfg.addressCol_s] + ",";
          memberCnt++;
        }
      }//for j
      //処理時間の計測
      var nowTime = new Date();
      var timeDiff = parseInt((nowTime.getTime() - startTime.getTime()) / (1000));
      Logger.log(Cfg.stRow_s + Number(i) + "行目(親組織), " + Cfg.inputCol_s + "メンバー[" + memberCnt + "], " + member);
      if( (Number(i) % 50) == 0 ){
        LogSheet("INFO",Number(i) + 1 + "行目(親組織), " + Cfg.inputCol_j + "| メンバー[" + memberCnt + "],  : " + timeDiff + " 秒 経過....");
      }
      ss.getRange(Cfg.stRow_s + Number(i), Cfg.inputCol_s).setValue(member);
      member = ""
    } //for i
    LogSheet("INFO",Number(i) + 1 + "行目(親組織), " + Cfg.inputCol_j + "| メンバー[" + memberCnt + "],  : " + timeDiff + " 秒 経過....");
    LogSheet("INFO", "処理終了")
  }catch(e){
    throw new Error( e.name + ", line:" + e.lineNumber + ", message:" + e.message );
    LogSheet("ERROR", e.name + ", line:" + e.lineNumber + ", message:" + e.message);
  }
}

function fn_storesMember(){
  var startTime = new Date();
  LogSheet("INFO", Cfg.getSheetName_j + " / " + Cfg.getSheetName_s + " : メンバー追加処理開始")
  try{
    var member = "";
    var sheetName_s = Cfg.getSheetName_s;
    var sheetName_j = Cfg.getSheetName_j;
    var ss = SpreadsheetApp.getActive().getSheetByName(sheetName_s);
    var ss2 = SpreadsheetApp.getActive().getSheetByName(sheetName_j);
    if ( ss == null ) {
      SpreadsheetApp.getActive().toast(sheetName_s + " シートが見つかりません","ERROR",5);
      return;
    }
    if ( ss2 == null ) {
      SpreadsheetApp.getActive().toast(sheetName_j + " シートが見つかりません","ERROR",5);
      return;
    }
    
    var enRow = ss.getLastRow() - Cfg.stRow_s + 1;
    if ( enRow < 1 ) {
      SpreadsheetApp.getActive().toast(sheetName_s + " データが見つかりません","ERROR",5);
      LogSheet("ERROR",sheetName_s + " データが見つかりません");
      return;
    }
    var datas = ss.getRange(Cfg.stRow_s, 1, enRow, Cfg.enCol_s).getValues();
    
    var enRow = ss2.getLastRow() - Cfg.stRow_j + 1;
    if ( enRow < 1 ) {
      SpreadsheetApp.getActive().toast(sheetName_j + " データが見つかりません","ERROR",5);
      LogSheet("ERROR",sheetName_j + " データが見つかりません");
      return;
    }
    var datas2 = ss2.getRange(Cfg.stRow_j, 1, enRow, Cfg.enCol_j).getValues();
    
    //親処理：一意のコード別に処理
    for ( var i in datas ) {
      //子処理：親コード別に処理（該当コードが親コードと一致したら親コード行のメールを配列化）
      if ( datas[i][Cfg.childCol_s] == "" ) {
        continue;
      }
      var tgtCol = 0;
      var memberCnt = 0;
      for ( var col = 0 ; col < Cfg.parentCnt_j ; col++ ) {
        Logger.log(col+"列目処理");
        tgtCol = Cfg.parentCol_j + col;
        for ( var j in datas2 ) {
          if ( datas[i][Cfg.childCol_s]　== datas2[j][tgtCol] ) {
            member = member + datas2[j][Cfg.addressCol_j] + ",";
            memberCnt++;
          }
        }//for j
      } //for col
      //処理時間の計測
      var nowTime = new Date();
      var timeDiff = parseInt((nowTime.getTime() - startTime.getTime()) / (1000));
      Logger.log(Cfg.stRow_s + Number(i) + "行目(組織), " + Cfg.inputCol_j + "| メンバー[" + memberCnt + "], " + member);
      if( (Number(i) % 50) == 0 ){
        LogSheet("INFO",Number(i) + 1 + "行目(組織), " + Cfg.inputCol_j + "| メンバー[" + memberCnt + "], : " + timeDiff + " 秒 経過....");
      }
      ss.getRange(Cfg.stRow_s + Number(i), Cfg.inputCol_j).setValue(member);
      member = ""
    } //for i
    LogSheet("INFO",Number(i) + 1 + "行目(組織), " + Cfg.inputCol_j + "| メンバー[" + memberCnt + "], : " + timeDiff + " 秒 経過....");
    LogSheet("INFO", "処理終了")
  }catch(e){
    throw new Error( e.name + ", line:" + e.lineNumber + ", message:" + e.message );
    LogSheet("ERROR", e.name + ", line:" + e.lineNumber + ", message:" + e.message);
  }
}