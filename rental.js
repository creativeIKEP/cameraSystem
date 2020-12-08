function fetchRental(studentno, camerano, lensno, sdcfno, cameracoment, lenscoment, sdcfcoment, key){
  rentalCheck(key, studentno, camerano, lensno, sdcfno);

  //スプレッドシートの取得
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet1 = spreadsheet.getSheetByName(sheet1Name);
  const sheet2 = spreadsheet.getSheetByName(sheet2Name);
  const sheet3 = spreadsheet.getSheetByName(sheet3Name);

  //貸出、返却を行うカメラ番号、レンズ番号、sdcf番号の行を格納する変数
  const cameraRow = FindRow(sheet1, camerano, 2);
  const lensRow = FindRow(sheet1, lensno, 2);
  const sdcfRow = FindRow(sheet1, sdcfno, 2);
  const studentNoRow = FindRow(sheet2, studentno, 1);//与えられた学生番号をスプレッドシートの「シート2」から探索
  const name = sheet2.getRange(studentNoRow, 2).getValue();

  if(key === "貸出"){
    //貸出日と決められた貸出期間から返却期限日を計算
    const today = Moment.moment();
    const returnDate = Moment.moment().add(kigen, "d");

    //物品の状態の変更、貸出者、貸出日、備考をスプレッドシートに記述
    if(cameraRow !== null){
      updateRentalStatus(cameraRow, "貸出中", name, today.format("YYYY/MM/DD"), returnDate.format("YYYY/MM/DD"), cameracoment);
    }
    if(lensRow !== null){
      updateRentalStatus(lensRow, "貸出中", name, today.format("YYYY/MM/DD"), returnDate.format("YYYY/MM/DD"), lenscoment);
    }
    if(sdcfRow !== null){
      updateRentalStatus(sdcfRow, "貸出中", name, today.format("YYYY/MM/DD"), returnDate.format("YYYY/MM/DD"), sdcfcoment);
    }
    //これまできたら貸出完了
    return "貸出を受け付けました!\n" + returnDate.format("YYYY年MM月DD日") + "までに返却をお願いします。";//貸出完了通知
  }

  //以下、返却の場合
  //貸出とプログラムはほぼ同じ
  //貸出を消すー＞返却履歴を更新
  if(cameraRow !== null){
    updateReturnData(cameraRow, name, camerano);
  }
  if(lensRow !== null){
    updateReturnData(lensRow, name, lensno);
  }
  if(sdcfRow !== null){
    updateReturnData(sdcfRow, name, sdcfno);
  }

  return "返却を受け付けました!";//返却完了通知
}

function rentalCheck(key, studentno, camerano, lensno, sdcfno){
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet1 = spreadsheet.getSheetByName(sheet1Name);
  const sheet2 = spreadsheet.getSheetByName(sheet2Name);

  if(typeof key !== "string" || typeof studentno !== "string" ||
  typeof camerano !== "string" || typeof lensno !== "string" || typeof sdcfno !== "string")
  {
    throw new Error("エラー！入力されたデータ型は無効です。");
  }
  if(key !== "貸出" && key !== "返却"){
    throw new Error("エラー！有効なリクエストは貸し出しか返却のみです。");
  }
  if(studentno === ""){
    throw new Error("エラー！学籍番号が入力されていません。");
  }
  if(camerano === "" && lensno === "" && sdcfno === ""){
    throw new Error("物品が何も指定されていません。");
  }

  const studentRow = FindRow(sheet2, studentno, 1);
  if(studentRow === null) {
    throw new Error("エラー！この学籍番号は登録されていないか入力が間違っています。");
  }

  const userName = sheet2.getRange(studentRow, 2).getValue();
  if(camerano !== ""){
    checkInput(key, camerano, userName);
  }
  if(lensno !== ""){
    checkInput(key, lensno, userName);
  }
  if(sdcfno !== ""){
    checkInput(key, sdcfno, userName);
  }
}


function checkInput(key, no, userName){
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet1 = spreadsheet.getSheetByName(sheet1Name);
  const sheet3 = spreadsheet.getSheetByName(sheet3Name);

  const sheet1RowNumber = FindRow(sheet1, no, 2);
  if(sheet1RowNumber === null) {
    //スプレッドシートにない場合
    throw new Error("エラー！" + no + "は、" + sheet1Name + "シートに登録されていません。");
  }

  const sheet3RowNumber = FindRow(sheet3, no, 1);
  if(sheet3RowNumber === null) {
    //スプレッドシートにない場合
    throw new Error("エラー！" + no + "は、" + sheet3Name + "シートに登録されていません。");
  }

  const status = sheet1.getRange(sheet1RowNumber, 4).getValue();
  if(key ==="貸出" && status !== ""){
    //物品の状態が空でない(=返却済みでない)時
    throw new Error("エラー！" + no + "は貸出中、もしくは貸出禁止状態です");
  }

  if(key === "返却"){
    if(status === ""){
      throw new Error("エラー！" + no + "は返却済みです。");
    }
    const LentUser = sheet1.getRange(sheet1RowNumber, 5).getValue();
    if(LentUser !== userName){
      throw new Error("エラー！貸出者本人が返却してください。");
    }
  }
}

function updateRentalStatus(rowNumber, status, name, rentalDate, returnDate, comment){
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet1 = spreadsheet.getSheetByName(sheet1Name);
  sheet1.getRange(rowNumber, 4).setValue(status);
  sheet1.getRange(rowNumber, 5).setValue(name);
  sheet1.getRange(rowNumber, 6).setValue(rentalDate);
  sheet1.getRange(rowNumber, 7).setValue(returnDate);
  sheet1.getRange(rowNumber, 8).setValue(comment);
}

function updateReturnData(sheet1RowNumber, userName, labelName){
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet1 = spreadsheet.getSheetByName(sheet1Name);
  const sheet3 = spreadsheet.getSheetByName(sheet3Name);
  const comment = sheet1.getRange(sheet1RowNumber, 8).getValue();
  Logger.log(comment);

  //シート1の貸出状態を消去
  updateRentalStatus(sheet1RowNumber, "", "", "", "", "");

  const numColmun = sheet3.getLastColumn();
  const sheet3RowNumber = FindRow(sheet3, labelName, 1);
  const moveTargetRange = sheet3.getRange(sheet3RowNumber, 4);
  sheet3.getRange(sheet3RowNumber, 2, 1, numColmun).moveTo(moveTargetRange);

  //今回の返却の履歴を新たに記述
  var whoAndCom = userName + "→";
  if(comment !== ""){
    whoAndCom += "\n" + comment;
  }

  sheet3.getRange(sheet3RowNumber, 2).setValue(whoAndCom);
  sheet3.getRange(sheet3RowNumber, 3).setValue(Moment.moment().format("YYYY/MM/DD"));
}

//与えられたシート(sheet)上の列(col)にデータvalがあるかを探索する関数
//データが見つかったらそのデータが格納されているセルの行をreturn
function FindRow(sheet, value, column){
  if(value === ""){return null;}

  const lastRow = sheet.getDataRange().getLastRow();
  const textFinder = sheet.getRange(1, column, lastRow, 1).createTextFinder(value);
  const findResults = textFinder.matchEntireCell(true).findAll();
  if(findResults.length > 1){
    throw new Error("エラー!" + value + "というデータが複数あります。");
  }
  if(findResults.length <= 0){
    return null;
  }
  return findResults[0].getRow();
}
