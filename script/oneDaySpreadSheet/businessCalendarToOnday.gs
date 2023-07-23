function main() {
  let activeSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  let ui = SpreadsheetApp.getUi();
  let instructionManual = activeSpreadSheet.getSheetByName("InstructionManual");
  //施設名コードを取得
  let facilityNameCode = instructionManual.getRange(2, 1).getValue();

  //営業カレンダー
  let reference_sheetName = facilityNameCode + "(利用者名)";
  //ondayシート名
  let input_sheetName;

  //営業カレンダー取得
  let reference_sheet = activeSpreadSheet.getSheetByName(reference_sheetName)
  if (reference_sheet == null) {
    ui.alert('sheetname : ' + reference_sheetName + 'not found');
  }

  let response = ui.prompt('西暦と月を入力(yymm)', 'ex) 2023年6月→2306:', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() === ui.Button.OK) {
    input_sheetName = response.getResponseText();
  }
  //ondayシート取得
  let input_sheet = activeSpreadSheet.getSheetByName(input_sheetName);
  if (input_sheet == null) {
  ui.alert('Sheet named ' + input_sheetName + ' not found');
  return; // ここで処理を終了
}

  //日付の取得
  let year_month = input_sheet.getRange(5, 1, 1, 2).getValues();
  let date = 1;
  console.log(year_month);

  //color templateを取得
  var venue;
  var color
  let colorTemplateSheet = activeSpreadSheet.getSheetByName("InstructionManual");
  let venuColorTemplate = colorTemplateSheet.getRange(2,2,10,2).getValues();
  // 辞書型（オブジェクト）を初期化
  let venueColors = {};
  for (let i = 0; venuColorTemplate[i][0] != ''; i++) {
    venue = venuColorTemplate[i][0]; // 会場名
    color = venuColorTemplate[i][1]; // 色(RGB値)
    venueColors[venue] = color; // 辞書型に追加
  }

  // ヘッダー行とタスク行の作成
  let headers = [new Date(year_month[0][0], year_month[0][1] - 1, date), "人数", "施設", "時間"];
  let tasks = ["タスク", "", "", ""];
  for (let i = 6; i <= 23; i++) {
    headers.push(i + ":00", "", "", "");
    tasks.push("", "", "", "");
  }
  headers.push("00:00", "", "", "", "01:00", "", "", "", "02:00", "", "", "", "03:00", "", "", "", "04:00", "", "", "", "05:00", "", "", "");
  tasks.push("", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");


  //営業カレンダー : reference_sheet, 書き込み先：input_sheet
  //テンプレの書き込みは(6,1), 団体名は(8,1) から
  //団体名，人数，会場，時間...

  let i;
  let j;
  let lastRow = reference_sheet.getLastRow();
  let lastColumn = reference_sheet.getLastColumn();

  let inputLine = 6;
  let headersLine = [];
  headersLine.push(6);
  let groupNum = 0;
  let infoOnday = new Array(100).fill('');
  let infoToOnday = [];
  infoToOnday.push(headers);
  infoToOnday.push(tasks);

  let time;
  let infoFromCalendar = reference_sheet.getRange(16, 1, lastRow - 15, lastColumn - 2).getValues();
  //input_sheet.getRange(16,1,lastRow-15,lastColumn-2).setValues(infoFromCalendar);
  //console.log(infoFromCalendar[0].length, infoFromCalendar.length);
  //console.log(infoFromCalendar);
  for (let j = 4; j < infoFromCalendar[0].length; j = j + 2) {
    for (let i = 0; i < infoFromCalendar.length; i = i + 4) {
      for (let k = 0; k < 2; k++) {
        time = getTimeValue(infoFromCalendar[i][j + k]);
        if (infoFromCalendar[i][j + k] == '' || time[0] === undefined || time[1] === undefined || isNaN(time[0]) || isNaN(time[1])) {
          continue;
        }
        infoOnday[0] = infoFromCalendar[i + 1][j + k];
        infoOnday[1] = infoFromCalendar[i + 3][j + k];
        let venueRow = 0;
        while (infoFromCalendar[i - venueRow][0] == ''){
          venueRow = venueRow + 4;
        }
        venue = infoFromCalendar[i - venueRow][0];
        infoOnday[2] = venue;
        infoOnday[3] = infoFromCalendar[i][j + k];
        infoOnday[time[0] - 1] = infoFromCalendar[i + 1][j + k] + ' : ' + infoFromCalendar[i + 3][j + k] + '人';

        if (venue in venueColors) {
          // venueがvenueColorsオブジェクトのキーとして存在する場合、その値を取得
          color = venueColors[venue];
        }
        else {
          // venueがvenueColorsオブジェクトのキーとして存在しない場合、デフォルトの色を設定
          color = "#FFFFE0";
        }
        input_sheet.getRange(inputLine + groupNum + date * 3 - 1, time[0], 1, Math.abs(time[1] - time[0])
        ).setBackground(color);
        infoOnday[time[1] - 2] = ';';
        groupNum += 1;

        infoToOnday.push(infoOnday);
        infoOnday = new Array(100).fill('');

      }
    }
    infoToOnday.push(new Array(100).fill(''));
    headersLine.push(inputLine + groupNum + date * 3);
    date = date + 1;
    infoToOnday.push(updateTemplateDate(headers.slice(), year_month[0][0], year_month[0][1], date));
    infoToOnday.push(tasks);
  }
  input_sheet.getRange(inputLine, 1, infoToOnday.length, 100).setValues(infoToOnday);
  //console.log(headersLine);

  for (let i = 0; i < headersLine.length; i++) {
    let row = headersLine[i];
    range = input_sheet.getRange(row, 1, 1, 98).setBackground("#ADFF2F"); // 背景色を黄色に設定
    input_sheet.getRange(row + 1, 1, 1, 98).setBackground("#FF6C5C"); // 明るいレッドベリー色に設定
  }
}

function updateTemplateDate(originalHeaders, year, month, date) {
  // ヘッダー行をコピー
  let newHeaders = originalHeaders.slice();
  // 日付を更新
  newHeaders[0] = new Date(year, month - 1, date);

  return newHeaders;
}


function getTimeValue(timeString) {
  // 辞書型の定義
  let timeDict = {
    "06:00": 5,
    "06:30": 7,
    "07:00": 9,
    "07:30": 11,
    "08:00": 13,
    "08:30": 15,
    "09:00": 17,
    "09:30": 19,
    "10:00": 21,
    "10:30": 23,
    "11:00": 25,
    "11:30": 27,
    "12:00": 29,
    "12:30": 31,
    "13:00": 33,
    "13:30": 35,
    "14:00": 37,
    "14:30": 39,
    "15:00": 41,
    "15:30": 43,
    "16:00": 45,
    "16:30": 47,
    "17:00": 49,
    "17:30": 51,
    "18:00": 53,
    "18:30": 55,
    "19:00": 57,
    "19:30": 59,
    "20:00": 61,
    "20:30": 63,
    "21:00": 65,
    "21:30": 67,
    "22:00": 69,
    "22:30": 71,
    "23:00": 73,
    "23:30": 75,
    "00:00": 77,
    "00:30": 79,
    "01:00": 81,
    "01:30": 83,
    "02:00": 85,
    "02:30": 87,
    "03:00": 89,
    "03:30": 91,
    "04:00": 93,
    "04:30": 95,
    "05:00": 97,
    "05:30": 99
  };


  // 文字列から時間を取り出す
  let times = timeString.split("/");

  // 辞書から対応する値を取得
  let value1 = timeDict[times[0]];
  let value2 = timeDict[times[1]];

  return [value1, value2];
}
