function infomationList() {
  const sheet4 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet4Name);
  const lastRow = sheet4.getDataRange().getLastRow();
  const infomationDatas = sheet4.getRange(11, 1, lastRow-10, 2).getValues();
  return infomationDatas.map(function(info){
    const infoDateStr = Moment.moment(info[0]).format("YYYY/MM/DD");
    return infoDateStr + ": " + info[1];
  });
}
