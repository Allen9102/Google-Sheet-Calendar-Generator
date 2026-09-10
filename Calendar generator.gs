function generateCalendar() {
  var ui = SpreadsheetApp.getUi();

  // 要求使用者輸入年份
  var yearInput = ui.prompt("行事曆生成", "請輸入要生成的年份 (西元年):", ui.ButtonSet.OK_CANCEL);
  if (yearInput.getSelectedButton() !== ui.Button.OK) return;
  var year = parseInt(yearInput.getResponseText());
  if (isNaN(year)) {
    ui.alert("請輸入有效的年份！");
    return;
  }

  // 要求使用者輸入月份
  var monthInput = ui.prompt("行事曆生成", "請輸入要生成的月份 (1-12):", ui.ButtonSet.OK_CANCEL);
  if (monthInput.getSelectedButton() !== ui.Button.OK) return;
  var month = parseInt(monthInput.getResponseText());
  if (isNaN(month) || month < 1 || month > 12) {
    ui.alert("請輸入有效的月份 (1-12)！");
    return;
  }

  // 創建新的工作表
  var sheetName = year + "年" + month + "月";
  var sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);

  // 設定標題
  sheet.getRange(1, 1, 1, 8).merge();
  sheet.getRange(1, 1).setValue(year + "年" + month + "月").setHorizontalAlignment("center").setFontSize(14).setFontWeight("bold");

  // 設定星期標題
  var daysOfWeek = ["星期", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"];
  sheet.getRange(2, 1, 1, 8).setValues([daysOfWeek]).setFontWeight("bold").setHorizontalAlignment("center");

  // 時間時段
  var timeSlots = [
    "8:00~9:00", "9:00~10:00", "10:00~11:00", "11:00~12:00", "12:00~13:00",
    "13:00~14:00", "14:00~15:00", "15:00~16:00", "16:00~17:00", "17:00~18:00",
    "18:00~19:00", "19:00~20:00", "20:00~21:00", "21:00~22:00", "22:00~23:00"
  ];

  // 取得該月份的第一天與最後一天
  var startDate = new Date(year, month - 1, 1);
  var lastDay = new Date(year, month, 0).getDate();
  var startWeekday = startDate.getDay();
  if (startWeekday === 0) startWeekday = 7; // 調整星期日為最後一天

  var rowOffset = 3;
  var currentDate = startDate;

  // 程式碼主體
  var rowOffset = 3;
  var currentDate = startDate;
  var nextMonth = month === 12 ? 1 : month + 1;

  while (true) {

    // 設定 "日期" 標題
    rowOffset++;
    sheet.getRange(rowOffset, 1)
      .setValue("日期")
      .setFontWeight("bold")
      .setFontSize(12);

    var stopAfterThisWeek = false; // ← 新增標記：本週填完後停止

    // 填入一整列的日期（星期一～星期日）
    for (var i = 1; i <= 7; i++) {

      // 若今天是下一月的星期日：仍要填，但記為最後一週
      if (currentDate.getMonth() + 1 === nextMonth && currentDate.getDay() === 0) {
        sheet.getRange(rowOffset, i + 1)
          .setValue(nextMonth + "月" + currentDate.getDate() + "日");
        
        stopAfterThisWeek = true; 
        currentDate.setDate(currentDate.getDate() + 1);
        break;
      }

      // 正常填日期（含跨月）
      if (currentDate.getDay() === i || (i === 7 && currentDate.getDay() === 0)) {
        var displayMonth = currentDate.getMonth() + 1;

        sheet.getRange(rowOffset, i + 1)
          .setValue(displayMonth + "月" + currentDate.getDate() + "日");

        currentDate.setDate(currentDate.getDate() + 1);
      }
    }

    // 一定要填時間欄（15 列）
    for (var j = 0; j < timeSlots.length; j++) {
      sheet.getRange(rowOffset + 1 + j, 1)
        .setValue(timeSlots[j])
        .setFontSize(10)
        .setBackground("#cccccc");
    }

    // 在 22:00~23:00 的下一列整列 A:H 上背景色#acb9ca
    var endRow = rowOffset + timeSlots.length + 1;
    sheet.getRange(endRow, 1, 1, 8).setBackground("#acb9ca");

    rowOffset = endRow;

    // 若這週是跨月的最後一週 → 填完時間就停止
    if (stopAfterThisWeek) break;
  }

  var lastRow = sheet.getLastRow();

  // 偵測 A 欄沒有文字的儲存格，將該列 A:H 顏色填滿 #acb9ca
  for (var row = 3; row <= lastRow; row++) {
    if (!sheet.getRange(row, 1).getValue()) {
      sheet.getRange(row, 1, 1, 8).setBackground("#acb9ca");
    }
  }

  // 偵測 A 欄為 "日期" 的列，檢查 B:H，合併沒有日期的儲存格
  for (var row = 3; row <= lastRow; row++) {
    if (sheet.getRange(row, 1).getValue() === "日期") {
      sheet.getRange(row, 1, 1, 8).setFontWeight("bold").setBackground("#cccccc");
      
      var emptyCells = [];
      for (var col = 2; col <= 8; col++) {
        var cell = sheet.getRange(row, col);
        if (!cell.getValue()) {
          emptyCells.push(cell);
        } else {
          if (emptyCells.length > 0) {
            mergeCellsAndFill(sheet, emptyCells);
          }
          emptyCells = [];
        }
      }
      if (emptyCells.length > 0) {
        mergeCellsAndFill(sheet, emptyCells);
      }
    }
  }

  // 設定B~H欄儲存格字體為 Arial，字體大小為 12，直到最後一列
  sheet.getRange(1, 2, lastRow, 9).setFontFamily("Arial").setFontSize(12);
  sheet.getRange(1, 1, 1, 8).setFontSize(14).setBackground("#8497b0"); // A1:H1 樣式
  sheet.getRange(2, 1, 1, 8).setFontSize(12).setBackground("#cccccc"); // A2:H2 樣式

  // 設定置中與框線
  sheet.getRange(2, 1, lastRow, 9).setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.getRange(2, 1, lastRow, 8).setBorder(true, true, true, true, true, true);
  
  // 凍結第2列
  sheet.setFrozenRows(2);

  // 將第一列到最後一列的列高設定為 25
  sheet.setRowHeights(1, sheet.getLastRow(), 25);
}

// 合併儲存格並填色
function mergeCellsAndFill(sheet, cells) {
  if (cells.length > 0) {
    var firstCell = cells[0];
    var lastCell = cells[cells.length - 1];

    var firstRow = firstCell.getRow();
    var lastRow = firstRow + 15;
    var firstCol = firstCell.getColumn();
    var lastCol = lastCell.getColumn();

    sheet.getRange(firstRow, firstCol, lastRow - firstRow + 1, lastCol - firstCol + 1)
      .merge()
      .setBackground("#cccccc");
  }
}
