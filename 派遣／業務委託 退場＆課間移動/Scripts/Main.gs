/**
 * 通知設定クラス
 */
var NoticeSetting = function(triggerDay, toList, ccList, subject, body) {
  this.TriggerDay = triggerDay;
  this.ToList = toList;
  this.CcList = ccList;
  this.Subject= subject;
  this.Body= body;
};


/**
 * 通知アクションのチェック＆実行
 */
function triggerNoticeActionCheckAndExec() {
  // 通知設定の取得
  var noticeActArray = getNoticeActionArray();
  
  // 当日の営業日何日目かを取得
  var today = new Date();
  var busToday = getDaysOfBusinessDays(today);
  var isBusday = isBusinessDay(today);

  // 営業日のみ実行  
  if (isBusday) {
    // 実行する通知があれば実行
    var act;
    var nextMonth;
    var subject;
    for (var i = 0, iMax = noticeActArray.length; i < iMax; i++) {
      act = noticeActArray[i];
      if (act.TriggerDay == busToday) {
        // 翌月を取得
        nextMonth = addDateTime(today, 'month', 1).getMonth() + 1;
        // 件名
        subject = act.Subject.replace(/<%m%>/g, nextMonth);
        // 通知
        //sendNoticeMail(act.ToList, subject, act.Body, act.CcList);
        sendNoticeMail(null, subject, act.Body, null);
      }
    }
  }
}

/**
 * 月末繰り越しアクションのチェック＆実行
 */
function triggerIncrementActionCheckAndExec() {
  var today = new Date();
  var lastDay = new Date(today.getFullYear(), today.getMonth() + 1, 0).getDate(); // 月末日を取得
  // 月末日以外は処理しない
  if (today.getDate() != lastDay) return;
  
  var sheet = SpreadsheetApp.getActive().getSheetByName('記入欄');
  var data = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
  var isEnableData = false;
  var dt;
  var rowRange;
  for (var i = 0, iMax = data.length; i < iMax; i++) {
    if (data[i][0] == '時期') {
      isEnableData = true;
    } else if (isEnableData == true) {
      dt = data[i][0];
      if (dt != null && dt.getMonth() <= today.getMonth()) {
        rowRange = sheet.getRange(i + 1, 1, 1, sheet.getLastColumn());
        rowRange.setBackgroundRGB(200, 200, 200);
      }
    }
  }
}

/**
 * 通知設定を取得
 * 
 * @return {Array} 通知設定
 *                 { TriggerDay: 実行日(営業日○日目), ToList: TOリスト, CcList: CCリスト, Subject: 件名, Body: 本文 }
 */
function getNoticeActionArray() {
  var actArray = new Array();
  var sheet = SpreadsheetApp.getActive().getSheetByName('通知設定');
  var settingData = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
  // 通知先設定の取得
  var toList = new Array();
  var ccList = new Array();
  var isDstSetting = false;
  var dstColumnSet = { DstType: 0, Address: 0 };
  var mailAddr;
  for (var i = 0, iMax = settingData.length; i < iMax; i++) {
    if (settingData[i][0] == '宛先種別') {
      // タイトル行から列番号取得
      for (var j = 0, jMax = settingData[i].length; j < jMax; j++) {
        if (settingData[i][j] == '宛先種別') {
          dstColumnSet.DstType = j;
        } else if (settingData[i][j] == 'アドレス') {
          dstColumnSet.Address = j;
        }
      }
      isDstSetting = true;
      continue;
    } else if (settingData[i][0] == 'ここまで通知先設定') {
      break;
    }
    
    if (isDstSetting == true && dstColumnSet != null) {
      mailAddr = settingData[i][dstColumnSet.Address];
      if (mailAddr != null && mailAddr != '') {
        if (settingData[i][dstColumnSet.DstType] == 'TO') toList.push(mailAddr);
        else if (settingData[i][dstColumnSet.DstType] == 'CC') ccList.push(mailAddr);
      }
    }
  }
                      
  // 通知内容設定の取得
  var isContentSetting = false;
  var contentColumnSet = { TriggerDay: 0, Subject: 0, Body: 0 };
  var triggerDay;
  for (var i = 0, iMax = settingData.length; i < iMax; i++) {
    if (settingData[i][0] == '通知タイミング(営業日)') {
      // タイトル行から列番号取得
      for (var j = 0, jMax = settingData[i].length; j < jMax; j++) {
        if (settingData[i][j] == '通知タイミング(営業日)') {
          contentColumnSet.TriggerDay = j;
        } else if (settingData[i][j] == '件名') {
          contentColumnSet.Subject = j;
        } else if (settingData[i][j] == '本文') {
          contentColumnSet.Body = j;
        }
      }
      isContentSetting = true;
      continue;
    } else if (settingData[i][0] == 'ここまで通知内容設定') {
      break;
    }
    
    if (isContentSetting == true && contentColumnSet != null) {
      triggerDay = settingData[i][contentColumnSet.TriggerDay];
      if (triggerDay != null && triggerDay > 0) {
        actArray.push(new NoticeSetting(
          triggerDay, 
          toList, 
          ccList, 
          settingData[i][contentColumnSet.Subject],
          settingData[i][contentColumnSet.Body]
        ));
      }
    }
  }
  return actArray;
}
