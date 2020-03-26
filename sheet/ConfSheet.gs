/* =============================================================
　　廃棄・確認用シート。
============================================================= */

var ConfSheet = function() {
  this.sheet = SpreadsheetApp.openById(MY_SHEET_ID).getSheetByName(SHEET_NAME_CONF);
  this.values = this.sheet ? this.sheet.getDataRange().getValues() : null;
  this.titleRow = 0;
  this.index = {};
  this.KINTONE_PREFIX = 'PC台帳: ';
  
  this.createIndex = function() {
    const KEY_TEXT = '依頼元所属名';
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
      pcNo         : filterData.indexOf('管理番号1'),
      pcNo2        : filterData.indexOf('管理番号2'),
      place        : filterData.indexOf('現物所在地'),
      maker        : filterData.indexOf('メーカー'),
      product      : filterData.indexOf('製品名'),
      model        : filterData.indexOf('モデル'),
      type         : filterData.indexOf('種別'),
      serial       : filterData.indexOf('シリアル'),
      certificate  : filterData.indexOf('証明書追加発行'),
      ssd          : filterData.indexOf('ストレージ対応'),
      note         : filterData.indexOf('備考'),
      conf         : filterData.indexOf('現物確認'),
      kintoneUpdate: filterData.indexOf('台帳更新'),
      row          : filterData.indexOf('行数'),
      mailText     : filterData.indexOf('▼業者確認用メール文([業者確認メール本文を作成]ボタン押下で更新)')
    };
    return this.index;
  }
}
  
ConfSheet.prototype = {
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
   * 台帳から「廃棄待ち」ステのデータを取得
   * jsonに基づき表に書き込む
   */
  editKintoneData: function(value, row) {
    var query = 'sub_status in ("廃棄待ち") and location in (\"SS\", \"ABTSD\") and ' + KintoneApi.QUERY_MY_OWNER;
    var fields = [KintoneApi.KEY_ID, 'pc_id', 'capc_id', 'location', 'pc_maker', 'pc_product', 'pc_model', 'pc_category', 'serial', 'appendix'];
    var response = KintoneApi.caApi.api.get(query, fields);
    var me = this;
    
    var data = response.map(function(value) {
      return [
        me.KINTONE_PREFIX + value[KintoneApi.KEY_ID].value,
        value.capc_id.value,
        value.pc_id.value,
        value.location.value === 'SS' ? 'スクランブルスクエア 21階' : value.location.value === 'ABTSD' ? 'AbemaTower 11階' : value.location.value,
        value.pc_maker.value,
        value.pc_product.value,
        value.pc_model.value,
        value.pc_category.value === 'N' ? 'ノートPC' : value.pc_category.value === 'D' ? 'デスクトップPC' : '',
        value.serial.value,
        '',
        '特別な対応不要',
        value.appendix.value,
        '=ROW()'
      ];
    });
    var lastRow = this.sheet.getRange('D:D').getValues().filter(String).length + 1;
    this.sheet
      .getRange(lastRow + 1, this.getIndex().belongs + 1, data.length, data[0].length)
      .setValues(data);
  },
  /**
   * 台帳のステータスを変更
   */
  updateKintoneStatus: function(value, row) {
    var index = this.getIndex();
    var today = Utilities.formatDate(new Date(), 'JST', 'yyyy-MM-dd');
    var KINTONE_PREFIX = this.KINTONE_PREFIX;
    
    var filteredValues = this.values.slice(this.titleRow).filter(function(value) {
      return value[index.conf] === '済' &&
        value[index.kintoneUpdate] != '済'&&
        value[index.belongs].indexOf(KINTONE_PREFIX) > -1;
    });
    var postValues = filteredValues.map(function(value) {
      return {
        id: value[index.belongs].replace(KINTONE_PREFIX, ''), 
        record: { 
          pc_status: { value: '廃止' }, sub_status: { value: '' }, location: { value: '社外' }, status_history: { value: today + ',廃棄シート' }, appendix: { value: value[index.note] + '\n' + today + ' 廃棄業者引き渡し済' }
        }
      };
    });
    KintoneApi.caApi.api.putRecords(postValues);
    
    var key = this.getRowKey('kintoneUpdate');
    var sheet = this.sheet;
    filteredValues.forEach(function(value) { // 台帳更新済にする
       sheet.getRange(key + value[index.row]).setValue('済');
    });
  },
  /**
   * 依頼用シートから、データを書き出す
   */
  editRequestData: function(baseSheet) {
    var reqIndex = baseSheet.getIndex();
    var index = this.getIndex();
    var baseData = baseSheet.values.slice(baseSheet.titleRow);
    var data = baseData
      .filter(function(value) { return value[reqIndex.received] === '受け取り済' })
      .map(function(value) {
        return [
          value[reqIndex.belongs],
          value[reqIndex.pcNo],
          value[reqIndex.pcNo2],
          value[reqIndex.place],
          value[reqIndex.maker],
          value[reqIndex.product],
          value[reqIndex.model],
          value[reqIndex.type],
          value[reqIndex.serial],
          value[reqIndex.certificate] === '不要' ? '' : value[reqIndex.certificate],
          value[reqIndex.ssd],
          value[reqIndex.note],
        ];
      });
    var lastRow = this.sheet.getRange('A:A').getValues().filter(String).length + 1;
    this.sheet
      .getRange(lastRow, reqIndex.belongs + 1, data.length, data[0].length)
      .setValues(data)
  },
  /**
   * メールの内容を更新する
   */
  updateMailText: function(value, row) {
    var name = Browser.inputBox('あなたの名前(名字のみ)を入力してください');
    
    var places = {};
    var index = this.getIndex();
    this.values.slice(this.titleRow).filter(function(value) {
      return value[index.conf] === '済';
    }).forEach(function(value) {
      if (value[index.place] === '') return;
      if (!places[value[index.place]]) places[value[index.place]] = [];
      places[value[index.place]].push(value);
    });
    
    var modelInfo = Object.keys(places).map(function (place) {
      var typeObj = {};
      places[place].forEach(function(value) {
        var type = (value[index.type] === '') ? 'その他' : value[index.type];
        if (!typeObj[type]) typeObj[type] = 0;
        typeObj[type] = (typeObj[type] + 1);
      })
      var type = Object.keys(typeObj).map(function (type) {
        return type + ' : ' + typeObj[type] + '台';
      });
      return '▼' + place + '▼\n' + type.join('\n');
    });
    
    var text = mailSheet.getTemplate('confTrader').text
      .replace('{name}', name)
      .replace('{list}', modelInfo. join('\n\n'));
      
    this.sheet.getRange(this.getRowKey('mailText') + '4').setValue(text)
  },
  /**
   * 業者引き渡し日を返す
   * @return date
   */
  getExecution: function() {
    return this.sheet.getRange('E1').getValue();
  }
};
var confSheet = new ConfSheet();

function tesConf() {
  //var h = confSheet.getRowKey('mailText');
  //confSheet.editKintoneData();
 // var sheet = SpreadsheetApp.openById(MY_SHEET_ID).getSheetByName('廃棄依頼用(2019.11)');
  confSheet.updateKintoneStatus();
}
