/**
 * メールテンプレートシート
 */
var mailSheet = function() {
  this.sheet = SpreadsheetApp.openById(MY_SHEET_ID).getSheetByName('メールテンプレート');
  this.values = this.sheet.getDataRange().getValues();
  this.titleRow = 1;
  this.index = {};
  
  this.createIndex = function() {
    var filterData = this.values[this.titleRow];
    this.index = {
      startAcceptance: filterData.indexOf('廃棄物受け取り開始'),
      confTrader     : filterData.indexOf('業者へ確定分送付')
    };
    return this.index;
  };
  
  this.getIndex = function() {
    return Object.keys(this.index).length ? this.index : this.createIndex();
  },

  /**
   * テンプレを返す
   * @param type{string} = startAcceptance, confTrader
   */
  this.getTemplate = function(type) {
    return {
      mail: this.values[this.titleRow + 2][this.getIndex()[type]],
      title: this.values[this.titleRow + 3][this.getIndex()[type]],
      text: this.values[this.titleRow + 4][this.getIndex()[type]]
    };
  };
  
  /**
   * メール送信
   * @param to string 送信先
   * @param from string 送信元
   * @param title string 件名
   * @param text string 本文
   */
  this.sendMail = function(to, from, title, text) {
    GmailApp.sendEmail(to, title, text, {
      from: from,
      replyTo: from, 
      name: '資産管理チーム'
    });
    return true;
  };
}

var mailSheet = new mailSheet();

function tes() {
  var m = mailSheet;
  var k1 = mailSheet.getTemplate('startAcceptance');
  var k2 = mailSheet.getTemplate('confTrader');
  Logger.log('hoge')
}