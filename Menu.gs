/**
 * メニューを設定する
 * トリガー登録しています。(池田)
 */
function createMenu() {
  var ui = SpreadsheetApp.getUi();         // Uiクラスを取得する
  var menu = ui.createMenu('▼スクリプト');  // Uiクラスからメニューを作成する
  // メニューにアイテムを追加する
  menu.addItem('新しい依頼シートを作る', 'onClickCreateReqSheet');
  menu.addItem('新しい確認シートを作る', 'onClickCreateConfSheet');
  menu.addItem('業者確認メール本文を作成', 'onClickCreateConfMail');
  menu.addItem('台帳の廃棄データを確認シートに取り込む', 'onClickEditKintone');
  menu.addItem('廃棄済の台帳のステータスを更新', 'onClickUpdateKintoneStatus');
  menu.addToUi(); // メニューをUiクラスに追加する
}


function onClickCreateReqSheet() {
  // 最新のシートがあったら名前を書き換える
  if (requestSheet.sheet) requestSheet.sheet.setName(SHEET_NAME_REQUEST.replace('最新', Utilities.formatDate(new Date(), 'JST', 'yy.MM/dd')));
  createNewSheet('request');
  Browser.msgBox('新しい依頼用と確認用のシートを作りました。',
    '担当者は依頼表の実施月・受け取り開始日・期限などを埋めてください。'
    , Browser.Buttons.OK);
}

function onClickCreateConfSheet() {
  if (confSheet.sheet) {
    // 業者引渡し日にする
    var date = Utilities.formatDate(confSheet.getExecution(), 'JST', 'yy.MM/dd');
    confSheet.sheet.setName(SHEET_NAME_CONF.replace('最新', date));
  }
  createNewSheet('conf');
  Browser.msgBox('新しい依頼用と確認用のシートを作りました。',
    '担当者は依頼表の実施月・受け取り開始日・期限などを埋めてください。'
    , Browser.Buttons.OK);
}

function onClickCreateConfMail() {
  confSheet.updateMailText();
}

function onClickEditKintone() {
  if (!confSheet.sheet) Browser.msgBox('「確認(最新)」というシートが見つかりません。「新しいシートを作る」を実行してからやり直してください。');
  // 台帳から廃棄データを確認シートに移行
  confSheet.editKintoneData();
}

function onClickUpdateKintoneStatus() {
  var conf = Browser.msgBox('台帳のステータスを一括変更しますか？', 'この動作は取り消せません。間違いがないか最終確認をしてから実行してください。', Browser.Buttons.OK_CANCEL);
  if (conf === 'ok') confSheet.updateKintoneStatus();
}