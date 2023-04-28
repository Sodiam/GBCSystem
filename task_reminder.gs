/**
 * Author: Sodiam
 */
function timeTrigger() {
  // トリガ削除
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    var name = trigger.getHandlerFunction();
    if (name == "main") {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // 毎日分のトリガ設定
  var date = new Date();
  date.setHours(9);
  date.setMinutes(0);
  ScriptApp.newTrigger("main").timeBased().at(date).create();
}

function main() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var today = new Date();
  var inCharges = [];
  var dueDates = [];
  var kinds = [];
  var message = "";
  var id = {"Member_name": "Member_ID"}; // slackで「メンバー→その他」からIDをコピー
  // Member's name and IDs are private
  
  //今日の日付と一致する期限日を取得
  var found = data.slice(1,data.length).filter(function(row) {
    var noticeDates = [];
    for (var i = 1; i <= 3; i++) {
      noticeDates.push(row[i]);
    }
    
    for (var i = 0; i < noticeDates.length; i++) {
      if(Utilities.formatDate(today, 'Asia/Tokyo', 'YYYY/MM/dd') === Utilities.formatDate(noticeDates[i], 'Asia/Tokyo', 'YYYY/MM/dd') && !row[5]) {
        dueDates.push(noticeDates[0]);
        inCharges.push(id[row[4]]);
        kinds.push(row[0]);
        return true;
      }
    }
    return false;
  })[0];
  
  if (typeof found === 'undefined') {
    return false;
  }
  
  for (var i = 0; i < dueDates.length; i++) {
    message = generateMessage(Utilities.formatDate(dueDates[i], 'Asia/Tokyo', 'YYYY/MM/dd'), kinds[i], inCharges[i]);
    sendMessageToSlack(message);
  }
}

function generateMessage(dueDate, kinds, inCharge) {
  return "<@" + inCharge + ">\n" + kinds + "の実行期限は" + dueDate + "までです。";
}

function generateLateMessage(kinds, inCharge) {
  return "<@" + inCharge + ">\n" + kinds + "の実行期限が過ぎているぞ！\nやり忘れはない？\nしっかりやったんであればチェックも忘れずに！";
}

function sendMessageToSlack(message) {
  var url = "Webhook URL is private";
  
  var payload = {
    "blocks": [
      {
        "type": "section",
        "text": {
          "type": "mrkdwn",
          "text": message
        }
      }
    ]
  };
  var params = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload)
  };
  
  UrlFetchApp.fetch(url, params);
}
