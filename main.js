var sheet1Name = "貸出状況";
var sheet2Name = "ユーザ";
var sheet3Name = "返却履歴";
var sheet4Name = "設定";
var kigen = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet4Name).getRange(1, 2).getValue();
var specialKigen = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet4Name).getRange(2, 2).getValue();


function doGet() {
  return HtmlService.createTemplateFromFile("index").evaluate().setTitle('カメラ貸出システム');
}

function updateDue(){
  const today = Moment.moment().startOf('day');
  const specialDueDate = Moment.moment(specialKigen).startOf('day');
  if(specialDueDate >= today){
    kigen = specialDueDate.diff(today, 'days');
  }
}

function LendOrReturn(stno, camerano, lensno, sdcfno, cameracoment, lenscoment, sdcfcoment, key){
  updateDue();
  return fetchRental(stno, camerano, lensno, sdcfno, cameracoment, lenscoment, sdcfcoment, key);
}

//引数noの学籍番号から誰が何を借りていて、その返却期限がいつまでかを調べる関数
function CheckUserLendLimit(studentNo){
  updateDue();
  return userLentLimitList(studentNo);
}

//貸出中の物品一覧の取得
function LendData(){
  updateDue();
  return allLendData();
}

function getInfomation(){
  return infomationList().reverse();
}

//貸出中で返却期限をすぎているものがないかをチェックする関数
//返却期限をすぎていればその物品の貸出者にメールを送信
//トリガー(Google Apps Scriptの機能)を使うと毎日自動実行されるため返却期限の確認ができる!!!!!
function CheckEmail(){
  updateDue();
  sendReminderEmail(overDueCheck());
}
