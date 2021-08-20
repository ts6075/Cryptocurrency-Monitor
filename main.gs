

function interval() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();  //工作表
  var usdtSheet = 'LINE通知設定';
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName(usdtSheet));
  var limit = 10, count = 0;
  while (true) {
    console.log('1>' + new Date().toISOString());
    var isPolling = spreadsheet.getActiveSheet().getRange('C3').getValue(); 
    if (count > limit || !isPolling || isPolling == 'N') {
      break;
    }
    console.log('2>' + new Date().toISOString());
    監測();
    console.log('3>' + new Date().toISOString());
    Utilities.sleep(100);
    console.log('4>' + new Date().toISOString());
    count = count + 1;
  }
}

function refreshData() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();  //工作表
  var usdtSheet = '幣安API';
  var random = Math.random();
  var binanceUrl = 'https://www.binance.com/api/v3/ticker/price#' + random
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName(usdtSheet)).getRange('A1').setValue(binanceUrl);
  //Utilities.sleep(1000)
}

function 監測() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();  //工作表
  var UsdtSheet = 'LINE通知設定';
  //refreshData();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName(UsdtSheet));
  var values = spreadsheet.getActiveSheet().getRange(1, 1, 20, 12).getValues();
  var anchor = [0, 0];
  //尋找定位點
  for (var row in values) {
    if (values[row][0] == "#") {
      anchor = [row, 0];
      break;
    }
  }
  //values[row][col]
  //循環檢查
  for (n = anchor[0]; n < 20; n++) {
    var 是否有交易對 = values[n][2];
      if (是否有交易對 != '') {
        var 通知價格 = values[n][6];
        var 實際價格 = values[n][8];
          if (通知價格 && 實際價格 >= 通知價格){
            var 交易對名稱 = spreadsheet.getActiveSheet().getRange('E' + n).getValue();
            var 現在價格 = spreadsheet.getActiveSheet().getRange('H' + n).getValue();
            var 前次價格 = spreadsheet.getActiveSheet().getRange('I' + n).getValue();
            var 漲跌幅 = spreadsheet.getActiveSheet().getRange('L' + n).getValue();          
            var 現在時間 = spreadsheet.getActiveSheet().getRange('L2').getValue();
            var token = spreadsheet.getActiveSheet().getRange('D3').getValue();
            var message = `\n交易對：${交易對名稱}\n` +
                          `現價：${現在價格}\n` +
                          `前價：${前次價格}\n` +
                          `通知：${通知價格}\n` +
                          `振幅：${漲跌幅}\n` +
                          `時間：${現在時間}`
            lineNotify(token, message);
          } 
          var 現在價格 = values[n][8];
          spreadsheet.getActiveSheet().getRange('I' + n).setValue(現在價格);
      }
    }
}

function lineNotify(token, message) {
  var options = {
    'method'  : 'post',
    'payload' : {'message' : message},
    'headers' : {'Authorization' : 'Bearer ' + token}
    };      
  UrlFetchApp.fetch('https://notify-api.line.me/api/notify',options);
}

