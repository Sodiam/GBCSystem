/**
 * Author: Sodiam
 */

function day_string(d) {
  switch(d) {
    case 1:
    return "月";
    case 2:
    return "火";
    case 3:
    return "水";
    case 4:
    return "木";
    case 5:
    return "金";
    default:
    return "";
  }
}

function delete_trigger() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    var name = trigger.getHandlerFunction();
    if (name === "set_classtime" || name === "set_classtime2") {
      ScriptApp.deleteTrigger(trigger);
      Logger.log("Trigger \"" + name + "\" was successfully deleted.");
    }
  });
}

function make_trigger() {
  ScriptApp.newTrigger("set_classtime").forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet()).onFormSubmit().create();
}

function sc_test(dt) {
  try {
    Utilities.formatDate(dt, "Asia/Tokyo", "YYYY/MM/dd HH:mm:ss");
    return true;
  } catch(error) {
    return false;
  }
}

function timestamp(times) {
  var day_time = day_string(times.getDay());
  var hour = times.getHours();
  var minute = times.getMinutes();
  if ((hour == 10 && minute >=40) || hour == 11 || (hour == 12 && minute < 30)) {
    day_time += "2";
  } else if (hour == 12 || (hour == 13 && minute < 10)) {
    day_time += "L";
  } else if (hour == 13  || hour == 14) {
    day_time += "3";
  } else if (hour == 15 || (hour == 16 && minute < 50)) {
    day_time += "4";
  } else if (hour == 16 || hour == 17 || hour == 18) {
    day_time += "5";
  } else {
    return "";
  }
  return "(" + day_time + ")";
}

function set_classtime(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0]; // 最新のシートを先頭で扱うこと
  var lastRow = sheet.getLastRow();
  var tcell = sheet.getRange("A" + lastRow);
  var times = tcell.getValue(); // タイムスタンプ形式で出力される
  if (sc_test(times)) {
    tcell.setValue(Utilities.formatDate(times, "Asia/Tokyo", "YYYY/MM/dd HH:mm:ss") + timestamp(times));
    setReasonCellWrapping(sheet, lastRow);
  } else {
    for (var i = lastRow-1; i > 1; i--) {
      tcell = sheet.getRange("A" + i);
      times = tcell.getValue();
      if (sc_test(times)) {
        tcell.setValue(Utilities.formatDate(times, "Asia/Tokyo", "YYYY/MM/dd HH:mm:ss") + timestamp(times));
        setReasonCellWrapping(sheet, i);
        break;
      }
    }
  }
}

function setReasonCellWrapping(sheet, row) {
  var cell = sheet.getRange("F" + row);
  var cell2 = sheet.getRange("J" + row);
  cell.setWrap(true);
  cell2.setWrap(true);
}

function get_weekly_sheetname() {
  var startdate = new Date();
  var enddate = new Date();
  enddate.setDate(enddate.getDate()+4);
  var sheetname = (startdate.getMonth()+1) + "." + startdate.getDate() + "-" + (enddate.getMonth() + 1) + "." + enddate.getDate();
  return sheetname;
}

function weekly_newsheet() {
  var form = FormApp.openById("Form ID is private");  // 来訪者登録フォームは削除しないこと

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var startdate = new Date();
  var enddate = new Date();
  enddate.setDate(enddate.getDate() + 4);
  var month = startdate.getMonth();
  if (month != 1 && month != 2 && month != 7) {
    form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
    var sheetname = get_weekly_sheetname();
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      var name = sheets[i].getName();
      if (name.indexOf('フォームの回答') != -1 && name.indexOf(sheetname) == -1) {
        sheets[i].setName(sheetname);
      }
    }
    var sh = ss.getSheetByName(sheetname);
    sh.setColumnWidth(1, 160);
    sh.setColumnWidth(2, 35);
    sh.setColumnWidth(5, 180);
    sh.setColumnWidth(10, 200);
    //sh.setColumnWidth(11, 300);
    var i1 = sh.getRange('I1');
    var j1 = sh.getRange('J1');
    //var k1 = sh.getRange('K1');
    i1.setValue("対応者");
    j1.setValue("内容");
    //k1.setValue("(新SA向け)対応の様子");
  }
}

function delete_past_datas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sn = get_weekly_sheetname();
  var sh = ss.getSheetByName(sn);
  if (sh != null) {
    try {
      sh.deleteRows(2, sh.getLastRow()-1);
      Logger.log("Duplicated data are successfully deleted.");
    } catch (error) {
      Logger.log("Executed linear deletion.");
      for (var i = 2; i < sh.getLastRow(); i++) {
        sh.deleteRow(2);
      }
      Logger.log("Duplicated data are successfully deleted.");
    }
  }
}

function delete_empty_sheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var form = FormApp.openById("Form ID is Private");
  form.removeDestination();
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var lastRow = sheets[i].getLastRow();
    if (lastRow === 1) {
      ss.deleteSheet(sheets[i]);
      Logger.log("Deleted empty sheet.");
    }
  }
}

function printDateError(error, ts) {
  return "[error name]:" + error.name + "\n" +
  "[Received Value]:" + ts;
}
