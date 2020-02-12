/* =============================================================
　　廃棄・依頼受付用シート。
============================================================= */

var RequestSheet = function() {
  this.sheet = SpreadsheetApp.openById(MY_SHEET_ID).getSheetByName(SHEET_NAME_REQUEST);
  this.values = this.sheet ? this.sheet.getDataRange().getValues() : null;
  this.titleRow = 0;
  this.index = {};
  this.today = getFormatDate(new Date());
  
  this.createIndex = function() {
    const KEY_TEXT = '所属名';
    var me = this;
    var filterData = (function() {
      for(var i = 0; i < me.values.length; i++) {
        if (me.values[i].indexOf(KEY_TEXT) > -1) {
          me.titleRow = i + 1;
          return me.values[i];
        }
      }
    }());
    if(!filterData || filterData.length === 0) {
      showTitleError();
      return;
    }
    this.index = {
      belongs      : filterData.indexOf(KEY_TEXT),
      requesterName: filterData.indexOf('記入者'),
      date         : filterData.indexOf('記入日'),
      requesterMail: filterData.indexOf('メールアドレス'),
      pcNo         : filterData.indexOf('管理番号1'),
      pcNo2        : filterData.indexOf('管理番号2'),
      place        : filterData.indexOf('受渡拠点'),
      maker        : filterData.indexOf('メーカー'),
      product      : filterData.indexOf('製品名'),
      model        : filterData.indexOf('モデル'),
      type         : filterData.indexOf('種別'),
      serial       : filterData.indexOf('シリアル'),
      certificate  : filterData.indexOf('証明書追加発行'),
      ssd          : filterData.indexOf('ストレージ対応'),
      note         : filterData.indexOf('備考'),
      handOver     : filterData.indexOf('受け渡し方法'),
      received     : filterData.indexOf('受け取り状況')
    };
    return this.index;
  }
}
  
RequestSheet.prototype = {
  getRowKey: function(target) {
    var targetIndex = this.getIndex()[target];
    var alfabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('');
    var returnKey = (targetIndex > -1) ? alfabet[targetIndex] : '';
    if (!returnKey || returnKey === '') showTitleError(target);
    return returnKey;
  },
  getIndex: function() {
    return Object.keys(this.index).length ? this.index : this.createIndex();
  },
  /**
   * サービスデスク受け取り開始日か否かのチェック
   * @return boolean
   */
  checkReceptionStartDate: function() {
    var date = this.sheet.getRange('H1').getValue();
    if (date === '' || date === '未定') return false;
    else return this.today === getFormatDate(date);
  },
  /**
   * 受け取り期限日の翌日か否かのチェック
   * @return boolean
   */
  checkReceptionEndDate: function() {
    var endDate = this.getReceptionEndDate();
    if (endDate === '' || endDate === '未定') return false;
    endDate.setDate(endDate.getDate() + 1);
    return this.today === getFormatDate(endDate);
  },
  /**
   * 受け取り期限日を返す
   * @return string
   */
  getReceptionEndDate: function() {
    return this.sheet.getRange('J1').getValue();
  },
  /**
   * 廃棄実施月を返す
   * @return string 'XX月'
   */
  getExecution: function() {
    return this.sheet.getRange('F1').getValue();
  },
  /**
   * 受け取り開始メールを送る
   * @return boolean
   */
  sendMailReceptionStartDate: function() {
    var index = this.getIndex();
    var adresses = this.values.slice(this.titleRow)
      .map(function(value) { return value[index.requesterMail]; })
      .filter(function(value) { return value != ''; })
      .filter(function(x, i, self) { return self.indexOf(x) === i; }); // 重複を削除
      
    var template = mailSheet.getTemplate('startAcceptance');
    var text = template.text.replace('{廃棄実施月}', this.getExecution()).replace('{受取期限}', Utilities.formatDate(this.getReceptionEndDate(), 'JST', 'MM/dd'));
    var title = template.title.replace('{廃棄実施月}', this.getExecution());
    
    adresses.forEach(function(adress) {
      mailSheet.sendMail(adress, template.mail, title, text);
      var hoge = ''
    });
  }
};
var requestSheet = new RequestSheet();

function te() {
   var t = requestSheet.checkReceptionEndDate();
   
  requestSheet.sendMailReceptionStartDate();
  var isE = requestSheet.checkReceptionEndDate();
  var ex = requestSheet.getExecution();
  var hoge = ''
}