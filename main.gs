/**工作表Proccess */
class SpreadProcess {
  /**建構子 */
  constructor() {
    this.spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); //工作表
    this.activeSheet = this.spreadsheet.getActiveSheet();     //當前sheet
    this.monitorSheet = '';                                   //監測頁面
    this.usdtSheet = '';                                      //幣安API
    this.anchor = [0, 0];                                     //錨點
  }
  /**初始化 */
  init(monitorSheetName, usdtSheetName) {
    this.monitorSheet = monitorSheetName;
    this.usdtSheet = usdtSheetName;
    this.setAnchor();
    return false;
  }
  /**切換sheet */
  changeSheet(sheetType) {
    var sheet;
    if (sheetType == 'usdt') {
      sheet = this.spreadsheet.getSheetByName(this.usdtSheet);
    } else {
      sheet = this.spreadsheet.getSheetByName(this.monitorSheet);
    }
    this.activeSheet = this.spreadsheet.setActiveSheet(sheet);
    return false;
  }
  /**設定定位點 */
  setAnchor() {
    this.anchor = [0, 0];
    this.changeSheet('monitor');
    var values = this.getAllValues();
    
    //尋找定位點
    for (var row in values) {
      if (values[row][0] == "#") {
        this.anchor = [row, 0];
        break;
      }
    }
    return false;
  }
  /**取得資料 */
  getAllValues() {
    return this.activeSheet.getRange(1, 1, 20, 12).getValues();
  }
}

/**工作表實例 */
globalThis.spreadIntance = new SpreadProcess();
/**程式進入點 */
function main() {
  var monitorSheetName = 'LINE通知設定';
  var usdtSheetName = '幣安API';

  var gp = globalThis.spreadIntance = new SpreadProcess();
  gp.init(monitorSheetName, usdtSheetName);
  refreshData();
  interval();
}

/**循環執行 */
function interval() {
  var gp = globalThis.spreadIntance;
  gp.changeSheet('monitor');

  var limit = 10, count = 0;  //for Test
  while (true) {
    var isPolling = gp.activeSheet.getRange('C3').getValue();
    if (count > limit || !isPolling || isPolling == 'N') {
      break;
    }

    監測();
    Utilities.sleep(100);
    count++;                  //for Test
  }
}

/**刷新價格表 */
function refreshData() {
  var random = Math.random();
  var binanceUrl = 'https://www.binance.com/api/v3/ticker/price#' + random;
  
  var gp = globalThis.spreadIntance;
  gp.changeSheet('usdt');
  gp.activeSheet.getRange('A1').setValue(binanceUrl);
  Utilities.sleep(1000)
}

/**監測 */
function 監測() {
  refreshData();
  var gp = globalThis.spreadIntance;
  gp.changeSheet('monitor');
  
  var token = gp.activeSheet.getRange('D3').getValue();
  var values = gp.getAllValues();

  //循環檢查
  for (var n = gp.anchor[0]; n < 20; n++) {
    var 是否有交易對 = values[n][2];
    if (是否有交易對 != '') {
      var 現在價格 = values[n][8];
      var 通知價格 = values[n][6];
      if (通知價格 && 通知價格 <= 現在價格){
        var 交易對名稱 = values[n][4];
        var 前次價格 = values[n][9];
        var 振幅 = values[n][10];      
        var message = `\n交易對：${交易對名稱}\n` +
                      `現價：${現在價格}\n` +
                      `前價：${前次價格}\n` +
                      `通知：${通知價格}\n` +
                      `振幅：${振幅}\n` +
                      `時間：${new Date().toISOString()}`
        lineNotify(token, message);
      }
      gp.activeSheet.getRange('I' + n).setValue(現在價格);
    }
  }
}
