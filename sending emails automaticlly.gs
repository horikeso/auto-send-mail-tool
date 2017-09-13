function sendEmails() {

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = spreadsheet.getSheetByName("送信リスト");

  var startRow = 2;// 開始行
  var lastRow = mainSheet.getLastRow();// 最終行
  if (lastRow < 2) {
    return;
  }

  var startColumn = 1;// 開始列
  var lastColumn = mainSheet.getLastColumn();// 最終列

  var sendString = "送信済";

  // 最後の行が入ってしまうのでlastRow - 1としています。
  var mainDataRange = mainSheet.getRange(startRow, startColumn, lastRow - 1, lastColumn);
  var mainData = mainDataRange.getValues();

  var fromSheet = spreadsheet.getSheetByName("送信元");
  var fromDataRange = fromSheet.getRange(2, 1, 2, 2);
  var fromData = fromDataRange.getValues();
  var fromAddress = fromData[0][0];// 送信元
  var fromName = fromData[0][1];// 送信元名

  var subjectSheet = spreadsheet.getSheetByName("件名");
  var subjectDataRange = subjectSheet.getRange(2, 1);
  var subjectData = subjectDataRange.getValues();
  var subject = subjectData[0][0];// 件名

  var messageSheet = spreadsheet.getSheetByName("本文");
  var messageDataRange = messageSheet.getRange(2, 1);
  var messageData = messageDataRange.getValues();
  var message = messageData[0][0];// 本文

  var attachmentSheet = spreadsheet.getSheetByName("添付ファイル");
  var attachmentLastRow = attachmentSheet.getLastRow();

  if (attachmentLastRow >= 2) {
    // 最後の行が入ってしまうのでlastRow - 1としています。
    var attachmentDataRange = attachmentSheet.getRange(2, 1, attachmentLastRow - 1);
    var attachmentData = attachmentDataRange.getValues();
    Logger.log(attachmentData);

    // 添付ファイル
    var fileBlobList = [];
    for (var n = 0; n < attachmentData.length; ++n) {
      var attachmentString = attachmentData[n][0];
      var attachmentPathParts = attachmentString.split("/");
      Logger.log("attachmentPathParts : " + attachmentPathParts);

      var folder = null;
      var partsIndex = 0;
      while (partsIndex < attachmentPathParts.length - 1) {
        var folders = null;
 
        if (folder === null) {
          folders = DriveApp.getFoldersByName(attachmentPathParts[partsIndex]);
        } else {
          folders = folder.getFoldersByName(attachmentPathParts[partsIndex]);
        }

        while (folders.hasNext()) {
          folder = folders.next();
        }
        partsIndex++;
      }

      Logger.log("folder : " + folder);

      var files = null;
      if (folder === null) {
        files = DriveApp.getFilesByName(attachmentPathParts[partsIndex]);
      } else {
        files = folder.getFilesByName(attachmentPathParts[partsIndex]);
      }

      while (files.hasNext()) {
        fileBlobList.push(files.next().getBlob());// BlobSource[]
      }

      Logger.log("fileBlobList : " + fileBlobList);  
    }
  }

  for (var i = 0; i < mainData.length; ++i) {

    var row = mainData[i];

    var toAddress = row[0];// 宛先
    var variable1 = row[1];// 変数1
    var variable2 = row[2];// 変数2
    var variable3 = row[3];// 変数3
    var variable4 = row[4];// 変数4
    var variable5 = row[5];// 変数5
    var result = row[6];// 送信判定

    // 本文の変数置換
    var newMessage = message.replace(/{{変数1}}/g, variable1)
      .replace(/{{変数2}}/g, variable2)
      .replace(/{{変数3}}/g, variable3)
      .replace(/{{変数4}}/g, variable4)
      .replace(/{{変数5}}/g, variable5);

    var options = {
      attachments: fileBlobList
    };
    if (fromAddress != "") {
      options.from = fromAddress;
    }
    if (fromName != "") {
      options.name = fromName;
    }

    if (result != sendString) {
      GmailApp.sendEmail(toAddress, subject, newMessage, options);
      mainSheet.getRange(startRow + i, lastColumn).setValue(sendString);
      SpreadsheetApp.flush();
    }
  }
}

function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [];
  menuEntries.push({name: "メール一括送信", functionName: "sendEmails"});
  ss.addMenu("スクリプト", menuEntries);
}