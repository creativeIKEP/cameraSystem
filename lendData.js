function userLentLimitList(studentNo) {
  if(typeof studentNo !== "string"){
    throw new Error("エラー！学籍番号が有効なデータ型でありません。");
  }
  if(studentNo === ""){
    throw new Error("エラー！学籍番号が入力されていません。");
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet1 = spreadsheet.getSheetByName(sheet1Name);
  const sheet2 = spreadsheet.getSheetByName(sheet2Name);

  //学籍番号を探索
  const studentDataRow = FindRow(sheet2,studentNo, 1);//学籍番号を探索
  if(studentDataRow === null){
    throw new Error("エラー！この学籍番号は登録されていないか入力が間違っています。");
  }

  //学籍番号から名前を取得
  const userName=sheet2.getRange(studentDataRow, 2).getValue();
  //対象となるシートの最終行を取得
  const lastRow=sheet1.getDataRange().getLastRow();

  const textFinder = sheet1.getRange(2, 5, lastRow, 1).createTextFinder(userName);
  const findResults = textFinder.matchEntireCell(true).findAll();
  if(findResults.length === 0){
    return userName + "さんが貸出中の物品はありません。";
  }

  var resultStr = "";
  findResults.forEach(function(range){
    const rowNumber = range.getRow();
    const labelName = sheet1.getRange(rowNumber, 2).getValue();
    const objectName = sheet1.getRange(rowNumber, 3).getValue();
    const resutnDate = Moment.moment(sheet1.getRange(rowNumber, 7).getValue());
    const returnDateStr = resutnDate.format("YYYY/MM/DD");
    resultStr += labelName + "\t" + objectName + "\t" + returnDateStr + "まで\n";
  });
  return userName + "さんが貸出中の物品の返却期限\n\n" + resultStr;
}


function allLendData(){
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet1 = spreadsheet.getSheetByName(sheet1Name);
  const lastRow = sheet1.getDataRange().getLastRow();

  const textFinder = sheet1.getRange(2, 4, lastRow, 1).createTextFinder("貸出中");
  const findResults = textFinder.matchEntireCell(true).findAll();
  if(findResults.length === 0){
    return "貸出中の物品はありません。";
  }

  var resultStr = "";
  findResults.forEach(function(range){
    const rowNumber = range.getRow();
    const labelName = sheet1.getRange(rowNumber, 2).getValue();
    const objectName = sheet1.getRange(rowNumber, 3).getValue();
    const lendUserName = sheet1.getRange(rowNumber, 5).getValue();
    const resutnDate = Moment.moment(sheet1.getRange(rowNumber, 7).getValue());
    const returnDateStr = resutnDate.format("YYYY/MM/DD");
    const comment = sheet1.getRange(rowNumber, 8).getValue();
    resultStr += labelName + "\t" + objectName + "\t" + lendUserName + "\t" + returnDateStr + "\t" + comment + "\n";
  });
  return "貸出中の物品一覧\n\n" + resultStr;
}
