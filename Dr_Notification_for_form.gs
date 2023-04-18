/**
 * Author: Sodiam
 */

const DEFAULT_MESSAGE = "private"

function timeTrigger() {
  // 土日・授業期間外以外の閉室日をDateオブジェクト形式で格納
  var closedDay = [];
  /*
    例：2021年11月23日→Date(2021, 10, 23) || Date("2021-11-23")
  */
  //トリガ削除
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    var name = trigger.getHandlerFunction();
    if (name == "openFormAcceptance" || name == "closeFormAcceptance") {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  //毎日分のトリガ設定
  var date = new Date().setHours(0, 0, 0, 0);
  var d = new Date();
  for (i = 0; i < closedDay.length; i++) {
    closedDay[i] = closedDay[i].setHours(0, 0, 0, 0);
  }
  var hasday = closedDay.indexOf(date) === -1 ? false : true;
  if (!hasday){
    customMessage = DEFAULT_MESSAGE;
    switch (d.getDay()) {
      case 5:
      d.setHours(12);
      d.setMinutes(30);
      ScriptApp.newTrigger("closeFormAcceptance").timeBased().at(d).create();
      switch(d.getMonth()) {
        case 3:
        case 4:
        case 5:
        case 6:
        d.setHours(13);
        d.setMinutes(10);
        ScriptApp.newTrigger("openFormAcceptance").timeBased().at(d).create();
        break;
        case 8:
        case 9:
        case 10:
        case 11:
        case 0:
        d.setHours(15);
        d.setMinutes(0);
        ScriptApp.newTrigger("openFormAcceptance").timeBased().at(d).create();
        break;
        default:
        return;
      }
      case 1:
      case 2:
      case 3:
      case 4:
      d.setHours(10);
      d.setMinutes(40);
      ScriptApp.newTrigger("openFormAcceptance").timeBased().at(d).create();
      d.setHours(18);
      ScriptApp.newTrigger("closeFormAcceptance").timeBased().at(d).create();
      break;
      default:
      return;
    }
  }
}

function pickUpFromSpreadsheet() {
  try {
    var spreadsheet = SpreadsheetApp.openById("ID is private"); // 時限名簿スプレッドシートのID
    var sheet = spreadsheet.getSheets()[0].getDataRange().getValues();
    var ids = spreadsheet.getSheets()[1].getDataRange().getValues();
    var day_time = timestamp();
    var staffs = [];
    var staff_id = createMemberMap(ids);
    var member = sheet.slice(0,sheet.length).filter(function(row) {
      if (row[0] === day_time) {
        for (var i = 1; i < row.length; i++) {
          if (row[i] == "") {
            break;
          }
          staffs.push(row[i]);
        }
        return staffs;
      }
    })[0];
    var jsonify = "";
    for (var i = 1; i < member.length; i++) {
      if (staff_id[member[i]] === undefined) {
        break;
      }
      jsonify += "<@" + staff_id[member[i]] + "> ";
    }
    jsonify += "\n";
    return jsonify;
  } catch(e) {
    return "";
  }
}

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

function timestamp() {
  var current_time = new Date();
  var day_time = day_string(current_time.getDay());
  var hour = current_time.getHours();
  var minute = current_time.getMinutes();
  if ((hour == 10 && minute >=40) || hour == 11 || (hour == 12 && minute < 30)) {
    day_time += "2";
  } else if (hour == 12 || (hour == 13 && minute < 10)) {
    day_time += "L";
  } else if (hour == 13 || hour == 14) {
    day_time += "3";
  } else if (hour == 15 || (hour == 16 && minute < 50)) {
    day_time += "4";
  } else if (hour == 16 || hour == 17 || hour == 18) {
    day_time += "5";
  } else {
    day_time = "";
  }
  return day_time;
}

function createMemberMap(ids) {
  var member_id = {};
  ids.slice(0,ids.length).filter(function(row) {
    member_id[row[0]] = row[1];
  });
  return member_id;
}

function sendMessageToSlack(e) {
  const itemResponses = e.response.getItemResponses();
  var url = "Slack channel url";
  var payload = {
    "blocks": [
      {
        "type": "section",
        "text": {
          "type": "mrkdwn",
          "text": pickUpFromSpreadsheet() + "フォームが送信されました。\n学籍番号：" + itemResponses[1].getResponse() + "\n学年　　：" + itemResponses[0].getResponse() + "\n氏名　　：" + itemResponses[2].getResponse() + "\n科目名　：" + itemResponses[3].getResponse() + "\n相談内容：" + itemResponses[4].getResponse()
        }
      }
    ]
  };
  var params = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload)
  };

  var message = UrlFetchApp.fetch(url, params);
}

function closeFormAcceptance() {
  var form = FormApp.getActiveForm();
  form.setAcceptingResponses(false);
  form.setCustomClosedFormMessage(DEFAULT_MESSAGE);
}

function openFormAcceptance() {
  var form = FormApp.getActiveForm();
  form.setAcceptingResponses(true);
}
