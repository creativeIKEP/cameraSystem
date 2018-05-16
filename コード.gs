
var kigen=7;//貸出期間。単位は「日」

//長期休暇での長期貸出などを行うときは以下のコメントアウトを解除

var SpecialKigenY=2018;     //特別貸出の返却期限の年
var SpecialKigenM=4;      //特別貸出の返却期限の月
var SpecialKigenD=10;     //特別貸出の返却期限の日
if((new Date(SpecialKigenY, SpecialKigenM-1,SpecialKigenD)).getTime()>=(new Date()).getTime()){
  var y=new Date().getYear();
  var m=new Date().getMonth()+1;
  var d=new Date().getDate();
  kigen=((new Date(SpecialKigenY, SpecialKigenM-1,SpecialKigenD)).getTime()-(new Date(y, m-1, d)).getTime())/(1000*60*60*24);
}


function doGet() {
  return HtmlService.createTemplateFromFile("cameraSystem").evaluate().setTitle('カメラ貸出システム');
}


function mainf(stno, camerano, lensno, sdcfno, cameracoment, lenscoment, sdcfcoment, key){
  //スプレッドシートの取得
 var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1 = spreadsheet.getSheetByName("シート1");
  
  //貸出、返却を行うカメラ番号、レンズ番号、sdcf番号の行を格納する変数
  var cameracell = 0;
  var lenscell = 0;
  var sdcfcell = 0;
  
  if(key=="貸出"){
    if(stno==""){
       return "エラー！学籍番号が入力されていません。";
    }
    else if(camerano=="" && lensno=="" && sdcfno==""){
     return "貸出物品が何も指定されていません。";
    }
    else{
      //入力異常がないかのチェック
      if(camerano != ""){
        if(camerano.substring(0,1) != "C"){
          //cameranoは大文字「C」から始まる番号
          return "エラー！カメラ番号の入力が間違っています。入力し直してください。";
        }
        cameracell=findRow2(sheet1, camerano, "B");//cameranoの行を探索し格納
        if(cameracell==0) {
          //スプレッドシートにない場合
          return "エラー！このカメラ番号は登録されていません。";
        }
        if(isdata(sheet1, "D", cameracell)){
          //物品の状態が空でない(=返却済みでない)時
          return "エラー！"+camerano+"は貸出中です";
        }
      }
      
      //cameraと同じ処理
      if(lensno != ""){
        if(lensno.substring(0,1) != "L"){
          return "エラー！レンズ番号の入力が間違っています。入力し直してください。";
        }
        lenscell=findRow2(sheet1, lensno, "B");
        if(lenscell==0) {
          return "エラー！このレンズ番号は登録されていません。";
        
        }
        if(isdata(sheet1, "D", lenscell)){
          return "エラー！"+lensno+"は貸出中です";
        }
      }
      
      //cameraと同じ処理
      if(sdcfno != ""){
        if(sdcfno.substring(0,1) != "S"){
          return "エラー！SD,CF番号の入力が間違っています。入力し直してください。";
        }
        sdcfcell=findRow2(sheet1, sdcfno, "B");
        if(sdcfcell==0) {
          return "エラー！このSD,CF番号は登録されていません。";
        }
        if(isdata(sheet1, "D", sdcfcell)){
          return "エラー！"+sdcfno+"は貸出中です";
        }
      }
      //ここまで異常がないかのチェック
      
      //以下、貸出処理
      var sheet2 = spreadsheet.getSheetByName("シート2");
      var stcell=findRow2(sheet2, stno, "A");//与えられた学生番号をスプレッドシートの「シート2」から探索
      if(stcell==0) {
        //学生番号が見つからなかった(=登録されていない)時
        return "エラー！この学籍番号は登録されていないか入力が間違っています。";
      }
      var name=sheet2.getRange("B"+stcell).getValue();//貸出者の名前の取得
      
      //貸出日と決められた貸出期間から返却期限日を計算
      var date = new Date();//今日の日にち(=貸出日)の取得
      var y=date.getYear();
      var m=date.getMonth()+1;
      var d=date.getDate()+kigen;
      var l=new Date(y, m, 0).getDate();
      while(d>l){
       d=d-l;
        m=m+1;
        if(m>12){m=1; y++;}
        l=new Date(y, m, 0).getDate();
      }
    
      
      //物品の状態の変更、貸出者、貸出日、備考をスプレッドシートに記述
      if(cameracell != 0){
        sheet1.getRange("D"+cameracell).setValue("貸出中");
        sheet1.getRange("E"+cameracell).setValue(name);
        sheet1.getRange("F"+cameracell).setValue(date);
        sheet1.getRange("G"+cameracell).setValue(new Date(y, m-1, d));
        sheet1.getRange("H"+cameracell).setValue(cameracoment);
      }
      if(lenscell != 0){
        sheet1.getRange("D"+lenscell).setValue("貸出中");
        sheet1.getRange("E"+lenscell).setValue(name);
        sheet1.getRange("F"+lenscell).setValue(date);
        sheet1.getRange("G"+lenscell).setValue(new Date(y, m-1, d));
        sheet1.getRange("H"+lenscell).setValue(lenscoment);
      }
      if(sdcfcell != 0){
        sheet1.getRange("D"+sdcfcell).setValue("貸出中");
        sheet1.getRange("E"+sdcfcell).setValue(name);
        sheet1.getRange("F"+sdcfcell).setValue(date);
        sheet1.getRange("G"+sdcfcell).setValue(new Date(y, m-1, d));
        sheet1.getRange("H"+sdcfcell).setValue(sdcfcoment);
      }
      //これまできたら貸出完了
      
      
      return "貸出を受け付けました!"+y+"年"+m+"月"+d+"日までに返却をお願いします。";//貸出完了通知
    }
  }
  
  //以下、返却の場合
  //貸出とプログラムはほぼ同じ
  else if(key=="返却"){
    if(stno==""){
      return "エラー！学籍番号が入力されていません。";
    }
    else if(camerano=="" && lensno=="" && sdcfno==""){
     return "貸出物品が何も指定されていません。";
    }
    else{
      var sheet2 = spreadsheet.getSheetByName("シート2");
      var stcell=findRow2(sheet2, stno, "A");
      if(stcell==0) {
        return "エラー！この学籍番号は登録されていないか入力が間違っています。";
      }
      var name=sheet2.getRange("B"+stcell).getValue();
      if(camerano != ""){
        if(camerano.substring(0,1) != "C"){
          return "エラー！カメラ番号の入力が間違っています。入力し直してください。";
        }
        cameracell=findRow2(sheet1, camerano, "B");
        if(cameracell==0) {
          return "エラー！このカメラ番号は登録されていません。";
        }
        var data=sheet1.getRange("D"+cameracell).getValue();
        var data2=sheet1.getRange("E"+cameracell).getValue();
        if(data==""){
          return "エラー！"+camerano+"は返却済みです。";
        }
        else if(data2 != name){
          return "エラー！貸出者本人が返却してください。";
        }
      }
      if(lensno != ""){
        if(lensno.substring(0,1) != "L"){
          return "エラー！レンズ番号の入力が間違っています。入力し直してください。";
        }
        lenscell=findRow2(sheet1, lensno, "B");
        if(lenscell==0) {
          return "エラー！このレンズ番号は登録されていません。";
        }
        var data=sheet1.getRange("D"+lenscell).getValue();
        var data2=sheet1.getRange("E"+lenscell).getValue();
        if(data==""){
          return "エラー！"+lensno+"は返却済みです。";
        }
        else if(data2 != name){
          return "エラー！貸出者本人が返却してください。";
        }
      }
      if(sdcfno != ""){
        if(sdcfno.substring(0,1) != "S"){
          return "エラー！SD,CF番号の入力が間違っています。入力し直してください。";
        }
        sdcfcell=findRow2(sheet1, sdcfno, "B");
        if(sdcfcell==0) {
          return "エラー！このSD,CF番号は登録されていません。";
        }
        var data=sheet1.getRange("D"+sdcfcell).getValue();
        var data2=sheet1.getRange("E"+sdcfcell).getValue();
        if(data==""){
          return "エラー！"+sdcfno+"は返却済みです。";
        }
        else if(data2 != name){
          return "エラー！貸出者本人が返却してください。";
        }
      }
    //ここまでエラーチェック
      
      //ここから返却手続き
      //貸出を消すー＞返却履歴を更新
      if(cameracell != 0){
        //シート1の貸出状態を消去
        sheet1.getRange("D"+cameracell).setValue("");
        sheet1.getRange("E"+cameracell).setValue("");
        sheet1.getRange("F"+cameracell).setValue("");
        sheet1.getRange("G"+cameracell).setValue("");
        sheet1.getRange("H"+cameracell).setValue("");
        sheet1.getRange("Z"+cameracell).setValue("");
        var sheet3 = spreadsheet.getSheetByName("シート3");//シート3、(返却履歴シート)の取得
        for(var i=24; i>=2; i-=2){
          //シート3上の過去の履歴を右にずらす
          var x=sheet3.getRange(cameracell, i).getValue();
          sheet3.getRange(cameracell, i+2).setValue(x);
          x=sheet3.getRange(cameracell, i+1).getValue();
           sheet3.getRange(cameracell, i+3).setValue(x);
        }
        //今回の返却の履歴を新たに記述
        var data = new Date();
        sheet3.getRange(cameracell, 2).setValue(name+"→");
        sheet3.getRange(cameracell, 3).setValue(data);
      }
      if(lenscell != 0){
        sheet1.getRange("D"+lenscell).setValue("");
        sheet1.getRange("E"+lenscell).setValue("");
        sheet1.getRange("F"+lenscell).setValue("");
        sheet1.getRange("G"+lenscell).setValue("");
        sheet1.getRange("H"+lenscell).setValue("");
        sheet1.getRange("Z"+lenscell).setValue("");
        var sheet3 = spreadsheet.getSheetByName("シート3");
        for(var i=24; i>=2; i-=2){
          var x=sheet3.getRange(lenscell, i).getValue();
          sheet3.getRange(lenscell, i+2).setValue(x);
          x=sheet3.getRange(lenscell, i+1).getValue();
           sheet3.getRange(lenscell, i+3).setValue(x);
        }
        var data = new Date();
        sheet3.getRange(lenscell, 2).setValue(name+"→");
        sheet3.getRange(lenscell, 3).setValue(data);
      }
      if(sdcfcell != 0){
        sheet1.getRange("D"+sdcfcell).setValue("");
        sheet1.getRange("E"+sdcfcell).setValue("");
        sheet1.getRange("F"+sdcfcell).setValue("");
        sheet1.getRange("G"+sdcfcell).setValue("");
        sheet1.getRange("H"+sdcfcell).setValue("");
        sheet1.getRange("Z"+sdcfcell).setValue("");
        var sheet3 = spreadsheet.getSheetByName("シート3");
        for(var i=24; i>=2; i-=2){
          var x=sheet3.getRange(sdcfcell, i).getValue();
          sheet3.getRange(sdcfcell, i+2).setValue(x);
          x=sheet3.getRange(sdcfcell, i+1).getValue();
           sheet3.getRange(sdcfcell, i+3).setValue(x);
        }
        var data = new Date();
        sheet3.getRange(sdcfcell, 2).setValue(name+"→");
        sheet3.getRange(sdcfcell, 3).setValue(data);
      }
      
      return "返却を受け付けました!";//返却完了通知
    }
  } 
}

//与えられたシート(sheet)上の列(col)にデータvalがあるかを探索する関数
//データが見つかったらそのデータが格納されているセルの行をreturn
function findRow2(sheet,val,col){
  var lastRow=sheet.getDataRange().getLastRow(); //対象となるシートの最終行を取得
  for(var i=1;i<=lastRow;i++){
    if(sheet.getRange(col+i).getValue() == val){
      return i;
    }
  }
  return 0;
}

//シートsheet上の行(row),列(col)に該当するセルにデータがあるかどうかを判定する関数
function isdata(sheet, col, row){
  var range = sheet.getRange(col+row);
  var data = range.getValue();
  if(data==""){return false;}
  else{return true;}
}



//貸出中で返却期限をすぎているものがないかをチェックする関数
//返却期限をすぎていればその物品の貸出者にメールを送信
//トリガー(Google Apps Scriptの機能)を使うと毎日自動実行されるため返却期限の確認ができる!!!!!
function check(){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("シート1");
  var today = new Date();
  var todayyear=today.getYear();
  var todaymonth=today.getMonth()+1;
  var todayday=today.getDate();
  var lastRow=sheet.getDataRange().getLastRow(); //対象となるシートの最終行を取得
  for(var i=2;i<=lastRow;i++){//すべての物品をチェック
    if(sheet.getRange("G"+i).getValue() != ""){
      var r=sheet.getRange("G"+i);
      var d=new Date(r.getValue());//返却日を取得
      var month=d.getMonth()+1;
      var day=d.getDate();
      var year=d.getYear();
      
      
      //返却期限をすぎていて、
      if((new Date(todayyear, todaymonth-1,todayday)).getTime()>(new Date(year, month-1, day)).getTime()){
        //貸出中なら
        if(sheet.getRange("D"+i).getValue()=="貸出中"){
          var name=sheet.getRange("E"+i).getValue();//貸出者の名前を取得
          //シート2から貸出者のメアドを取得
          var sheet2 = spreadsheet.getSheetByName("シート2");
          var x=findRow2(sheet2, name, "B");
          var meado=sheet2.getRange(x, 3).getValue();
          var num=sheet.getRange("B"+i).getValue();
          
          if(sheet.getRange("Z"+i).getValue()==""){
            Logger.log(num);//デバック用
            //メールの送信
            //MailApp.sendEmail(meado, "貸出中物品の返却について", "貸出物品番号「"+num+"」の返却期限が過ぎています。返却をお願いします。\nなお、今回のお知らせと返却が行き違いになった場合はご容赦ください。\n\n\n*このメールはコンピュータにより自動送信されています。このメールに返信しないでください。\n\nfrom:卒アル委員カメラ貸出システム(https://script.google.com/macros/s/AKfycbzHFuODYdk-53alyA6J7ZZYAwEdLvXq8-KDNu2EDA/dev)");
            sheet.getRange("Z"+i).setValue("M");
          }
        }
      }
    }
  }
}