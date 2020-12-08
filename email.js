function overDueCheck(){
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet1 = spreadsheet.getSheetByName(sheet1Name);
  const sheet2 = spreadsheet.getSheetByName(sheet2Name);
  const sheet1LastRow = sheet1.getDataRange().getLastRow();
  const sheet2LastRow = sheet2.getDataRange().getLastRow();
  const lendDatas = sheet1.getRange(2, 1, sheet1LastRow, 8).getValues();
  const userDatas = sheet2.getRange(2, 1, sheet2LastRow, 3).getValues();

  return userDatas
  .map(extractLendData)
  .filter(function(data){
    return data.lendDatas.length > 0;
  });
}

function extractLendData(userData){
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet1 = spreadsheet.getSheetByName(sheet1Name);
  const sheet1LastRow = sheet1.getDataRange().getLastRow();
  const lendDatas = sheet1.getRange(2, 1, sheet1LastRow, 9).getValues();
  const userName = userData[1];
  const today = Moment.moment().startOf('day');

  var result = {
    emailAddress: userData[2],
    userName: userName,
    lendDatas: []
  };
  lendDatas.forEach(function(lendData, index){
    const returnDate = Moment.moment(lendData[6]).startOf('day');
    if(lendData[3] === "貸出中" && lendData[4] === userName && returnDate < today){
      const overWeekData = lendData[8];
      const overDueDays = today.diff(returnDate, "days");
      if(overWeekData === ""){
        sheet1.getRange(index + 2, 9).setValue(1);
        result.lendDatas.push([0, lendData]);
      }
      else if(overDueDays % 7 === 0){
        const overWeek = overDueDays / 7;
        sheet1.getRange(index + 2, 9).setValue(overWeek + 1);
        result.lendDatas.push([overWeek, lendData]);
      }
    }
  });
  return result;
}

function sendReminderEmail(overDueDatas){
  overDueDatas.forEach(function(overDueData){
    const address = overDueData.emailAddress;
    const subject = "貸出中物品の返却について";

    var body = overDueData.userName + "さん\n\
    以下の貸出物品の返却期限が過ぎています。返却をお願いします。\n\n\n";

    var html = overDueData.userName + "さん<br>\
    以下の貸出物品の返却期限が過ぎています。返却をお願いします。<br><br><br>";
    html += "<table style='border-collapse: collapse;'>\
    <tr>\
    <th style='border: solid 1px;'>物品番号</th>\
    <th style='border: solid 1px;'>詳細</th>\
    <th style='border: solid 1px;'>貸出日</th>\
    <th style='border: solid 1px;'>返却期限日</th>\
    </tr>";

    overDueData.lendDatas.forEach(function(lendData){
      const overWeek = lendData[0];
      var overWeekMessage = "";
      if(overWeek > 0){
        overWeekMessage = overWeek + "週間以上超過";
      }
      const rendatalDateStr = Moment.moment(lendData[1][5]).format("YYYY/MM/DD");
      const returnDateStr = Moment.moment(lendData[1][6]).format("YYYY/MM/DD");
      body += lendData[1][1] + " " + lendData[1][2] + " " + rendatalDateStr + " " + returnDateStr + " " + overWeekMessage + "\n";
      html += "<tr>\
      <td style='border: solid 1px;'>" + lendData[1][1] + "</td>\
      <td style='border: solid 1px;'>" + lendData[1][2] + "</td>\
      <td style='border: solid 1px;'>" + rendatalDateStr + "</td>\
      <td style='border: solid 1px;'>" + returnDateStr + "</td>\
      <td style='color: red;'>" + overWeekMessage + "</td>\
      </tr>";
    });

    body += "\n\nなお、今回のお知らせと返却が行き違いになった場合はご容赦ください。\n\n\
    *このメールはコンピュータにより自動送信されています。このメールに返信しないでください。\
    \nfrom:カメラ貸出システム\n";

    html += "</table><br><br>\
    なお、今回のお知らせと返却が行き違いになった場合はご容赦ください。<br><br>\
    *このメールはコンピュータにより自動送信されています。このメールに返信しないでください。<br>\
    From カメラ貸出システム";

    MailApp.sendEmail(address, subject, body, {htmlBody:html});
  });
}
