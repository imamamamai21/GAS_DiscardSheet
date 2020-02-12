/* =============================================================
　　廃棄の依頼を受け・廃棄の確認をするためのシートです
============================================================= */

/**
 * 発注終了時にテンプレートをコピーします
 * @param テンプレートのシートどちらか request || conf
 * @return string シートのURL
 */
function createNewSheet(type) {
  var sheet = SpreadsheetApp.openById(MY_SHEET_ID);
  var templateSheet = sheet.getSheetByName(type === 'conf' ? SHEET_NAME_CONF_TEMPLATE : SHEET_NAME_REQUEST_TEMPLATE);
  var copySheet = templateSheet.copyTo(sheet);
  sheet.setActiveSheet(copySheet);
  sheet.moveActiveSheet(type === 'conf' ? 6 : 5);
  copySheet.setName(type === 'conf' ? SHEET_NAME_CONF : SHEET_NAME_REQUEST);
  return SHEET_URL_BASE + MY_SHEET_ID + "#gid=" + copySheet.getSheetId();
}

function showTitleError(key) {
  Browser.msgBox('データが見つかりません', '表のタイトル名を変えていませんか？ : ' + key, Browser.Buttons.OK);
}

function getFormatDate(date) {
 return Utilities.formatDate(date, 'JST', 'yyyy-MM-dd');
}

/**
 * 朝のチェック
 * 依頼表の受取り開始日・受け取り期限の日に動作する
 * トリガー登録：池田 10:00~11:00
 */
function morningCheck() {
  if (!requestSheet) return;
  if (!confSheet) createNewSheet('conf');
  
  // 受け取り開始日のメール
  if (requestSheet.checkReceptionStartDate()) requestSheet.sendMailReceptionStartDate();
  
  // 受け取り期限の翌日だったら新規シートを作る
  if (!requestSheet.checkReceptionEndDate()) return;
  var nowRequestSheet = requestSheet;
  
  // 依頼表のシート名変更 
  nowRequestSheet.sheet.setName(SHEET_NAME_REQUEST.replace('最新', Utilities.formatDate(new Date(), 'JST', 'yy.MM/dd')));
  // 確認シートにデータを移行
  confSheet.editRequestData(nowRequestSheet);

  // 締め切りを表示
  var lastRow = nowRequestSheet.sheet.getRange('G:G').getValues().filter(String).length + 2;
  nowRequestSheet.sheet.getRange('A' + lastRow).setValue('-------------------- 依頼は締め切りました --------------------');
  nowRequestSheet.sheet.getRange('A' + lastRow + ':R' + lastRow).setBackground('#666666').setFontColor('white');
  nowRequestSheet.sheet.deleteRows(lastRow + 1, nowRequestSheet.sheet.getMaxRows() - lastRow);
}
