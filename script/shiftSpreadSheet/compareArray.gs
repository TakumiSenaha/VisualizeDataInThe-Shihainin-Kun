function getSheetData() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('西暦と月の入力(yymm)', 'シート名(yymm):', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() === ui.Button.OK) {
    var sheetName = response.getResponseText();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    if (sheet == null) {
      return;
    }
    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();
    compareArrays(values, sheetName);
  }
  // キャンセルが選択された場合はnullを返す
  else {
    return null;
  }
}

function writeToSheet(sheetName, data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var numRows = data[0].length; // 転置後の行数
  var numCols = data.length;    // 転置後の列数
  //sheet.getRange(2, 2, numCols, numRows).setValues(data);
  console.log(numRows,numCols);
  sheet.getRange(16, 5, numCols, numRows).setValues(data);
}

function compareArrays(inputArray, sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var roomSheet = ss.getSheetByName("InstructionManual"); // 部屋名が記載されたシート名を指定
  var roomList = roomSheet.getRange(3, 3, roomSheet.getLastRow(), 1).getValues();

  // 部屋名リストの終わりを見つける
  var lastRoomIndex = -1;
  for (var i = 0; i < roomList.length; i++) {
    if (roomList[i][0] === "") {
      lastRoomIndex = i;
      break;
    }
  }

  // 空白セルが見つからなかった場合、リスト全体を使用
  if (lastRoomIndex === -1) {
    lastRoomIndex = roomList.length;
  }

  // 部屋名リストを取得
  var dictionary = roomList.slice(0, lastRoomIndex);
  console.log(dictionary);
  
  var result = [];
  for (var j = 0; j < dictionary.length; j++) {
    var innerResult = []; // 内部の結果配列を作成
    for (var i = 0; i < inputArray.length; i++) {
      var searchWord = dictionary[j];
      if (inputArray[i].join().indexOf(searchWord) !== -1) {
        innerResult.push("〇");
      } else {
        innerResult.push("");
      }
    }
    result.push(innerResult); // 内部の結果配列を result 配列に追加
  }
  console.log(result);
  writeToSheet("日ごとの入力", result);
  //writeToSheet("日ごとの入力", result);
}

function transposeArray(array) {
  return array[0].map(function(_, i) {
    return array.map(function(row) {
      return row[i];
    });
  });
}
