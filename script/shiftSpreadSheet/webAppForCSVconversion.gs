/**
 * V8トリガーバグ対策
*/
function doPost(e) {
  main();
}

function main() {
  var html = HtmlService.createHtmlOutputFromFile("index");
  SpreadsheetApp.getUi().showModalDialog(html, 'ローカルファイル読込');
}

function extractRoomName(input_string) {
  var delimiter = /[,、]/;  // 区切り文字としてカンマと読点を指定
  var room_name_list_ = input_string.split(delimiter).filter(function(name) {
    return name.trim() !== '';
  });
  return room_name_list_;
}

function isConsecutiveDate(date1, date2) {
  var oneDay = 24 * 60 * 60 * 1000; // 1日のミリ秒数
  var time1 = new Date(date1).getTime();
  var time2 = new Date(date2).getTime();
  return Math.abs(time1 - time2) === oneDay;
}

function countEmptyDates(date1, date2) {
  var oneDay = 24 * 60 * 60 * 1000; // 1日のミリ秒数
  var time1 = new Date(date1).getTime();
  var time2 = new Date(date2).getTime();
  var diffDays = Math.round(Math.abs((time2 - time1) / oneDay));
  return diffDays;
}


function processCSVFile(formObject) {
  // フォームで指定したテキストファイルを読み込む
  var fileBlob = formObject.myFile;
  
  // CSVファイルを取得
  var csv_content = fileBlob.getBlob().getDataAsString("Shift_JIS");

  // CSVデータの処理
  var room_name_list = [];
  var csv_data = Utilities.parseCsv(csv_content);
  var current_date = csv_data[1][0]; // 最初の日付を取得
  var room_name_list_tmp = [];

  // 最初の行を処理
  var firstDay = parseInt(current_date.split('/')[2]);
  if (firstDay != 1) {
    for (var j = 1; j < firstDay; j++) {
      room_name_list.push([]); // 空のリストを追加
      room_name_list.push([]); // 空のリストを追加
    }
  }
  var row = csv_data[1];
  var one_day_room_name_list = extractRoomName(row[2]);
  room_name_list_tmp = room_name_list_tmp.concat(one_day_room_name_list);
  var emptyRowCount = 0;
  // 2行目から処理を開始
  for (var i = 2; i < csv_data.length; i++) {
    row = csv_data[i];
    if (row[0] === "" || row[0] === null || row[0] === undefined) {
      emptyRowCount++;
      if (emptyRowCount > csv_data.length - 2) {
        throw new Error('エラー: 全ての行が空です。データを確認してください。');
      }
      continue; // Skip the current iteration and move to the next one
    }
    date = row[0];
    if (date != current_date) {
      room_name_list.push(room_name_list_tmp.slice());
      room_name_list_tmp = [];
      room_name_list.push([]); // 空のリストを追加
      var emptyDates = countEmptyDates(current_date, date);
      for (var j = 0; j < (emptyDates-1); j++) {
        room_name_list.push([]); // 空のリストを追加
        room_name_list.push([]); // 空のリストを追加
      }
    }
    one_day_room_name_list = extractRoomName(row[2]);
    room_name_list_tmp = room_name_list_tmp.concat(one_day_room_name_list);
    current_date = date;
  }

  // 最後の行の場合の処理
  if (room_name_list_tmp.length > 0) {
    room_name_list.push(room_name_list_tmp.slice());
  }

  // 出力データをCSVファイルに書き込む
  var csv_output = room_name_list.map(function(row) {
    return row.join(',');
  }).join('\n');

  // 現在のアクティブなスプレッドシートに新しいシートとして追加する
  var sheetName;
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('西暦と月を入力(yymm)', 'ex) 2023年6月→2306:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() === ui.Button.OK) {
    sheetName = response.getResponseText();
  }
  var active_sheet = SpreadsheetApp.getActiveSpreadsheet();
  // 既に存在するシートを取得する
  var existingSheet = active_sheet.getSheetByName(sheetName);
  if (existingSheet) {
    // シートが存在する場合は削除する
    active_sheet.deleteSheet(existingSheet);
    }
    // 新しいシートを追加する
    var new_sheet = active_sheet.insertSheet(sheetName);
    // CSVデータを処理してシートに書き込む
    var csv_content_array = Utilities.parseCsv(csv_output);
    new_sheet.getRange(1, 1, csv_content_array.length, csv_content_array[0].length).setValues(csv_content_array);
    // ファイルの処理が完了したことをユーザに通知する
    return 'success';
  }
