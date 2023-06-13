const StockData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Stock Data');

/*
function displayHoseData() {

  var cache = CacheService.getScriptCache();
  var hoseData = JSON.parse(cache.get("hoseData"));
  var valueI = StockData.getRange('I4').getValue();
  if(valueI != 'HOSE') {
    StockData.getRange('I4').setValue('HOSE')
  }
  var page = parseInt(StockData.getRange('E57').getValue()); // Lấy giá trị của ô 'page'

  var startIndex = (page - 1) * 50; // Vị trí bắt đầu
  var endIndex = Math.min(startIndex + 50, Object.keys(hoseData).length); // Vị trí kết thúc

  // Xóa dữ liệu cũ trên sheet
  StockData.getRange('A2:F51').clearContent();

  for (var i = startIndex; i < endIndex; i++) {
    var ticker = Object.keys(hoseData)[i];
    var data = hoseData[ticker];

    // Ghi thông tin của mã cổ phiếu lên sheet
    var rowData = [
      data.tickerName,
      data.companyName,
      data.last,
      data.chg,
      data.marketCap,
      data.vol
    ];
    StockData.getRange(i - startIndex + 2, 1, 1, rowData.length).setValues([rowData]);
  }
}

function displayHNXData() {

  var cache = CacheService.getScriptCache();
  var hnxData = JSON.parse(cache.get("hnxData"));
  var valueI = StockData.getRange('I4').getValue();
  if(valueI != 'HNX') {
    StockData.getRange('I4').setValue('HNX')
  }
  var page = parseInt(StockData.getRange('E57').getValue()); // Lấy giá trị của ô 'page'

  var startIndex = (page - 1) * 50; // Vị trí bắt đầu
  var endIndex = Math.min(startIndex + 50, Object.keys(hnxData).length); // Vị trí kết thúc

  // Xóa dữ liệu cũ trên sheet
  StockData.getRange('A2:F51').clearContent();

  for (var i = startIndex; i < endIndex; i++) {
    var ticker = Object.keys(hnxData)[i];
    var data = hnxData[ticker];

    // Ghi thông tin của mã cổ phiếu lên sheet
    var rowData = [
      data.tickerName,
      data.companyName,
      data.last,
      data.chg,
      data.marketCap,
      data.vol
    ];
    StockData.getRange(i - startIndex + 2, 1, 1, rowData.length).setValues([rowData]);
  }
}

function displayUpcomData() {

  var cache = CacheService.getScriptCache();
  var upcomData = JSON.parse(cache.get("upcomData"));
  var valueI = StockData.getRange('I4').getValue();
  if(valueI != 'UPCOM') {
    StockData.getRange('I4').setValue('UPCOM');
  }
  var page = parseInt(StockData.getRange('E57').getValue()); // Lấy giá trị của ô 'page'


  var startIndex = (page - 1) * 50; // Vị trí bắt đầu
  var endIndex = Math.min(startIndex + 50, Object.keys(upcomData).length); // Vị trí kết thúc

  // Xóa dữ liệu cũ trên sheet
  StockData.getRange('A2:F51').clearContent();

  for (var i = startIndex; i < endIndex; i++) {
    var ticker = Object.keys(upcomData)[i];
    var data = upcomData[ticker];

    // Ghi thông tin của mã cổ phiếu lên sheet
    var rowData = [
      data.tickerName,
      data.companyName,
      data.last,
      data.chg,
      data.marketCap,
      data.vol
    ];
    StockData.getRange(i - startIndex + 2, 1, 1, rowData.length).setValues([rowData]);
  }
}

function display() {
  var exchange = StockData.getRange('I4').getValue(); // Lấy tên sàn

  if (exchange == "HOSE") {
    displayHoseData();
  }
  if (exchange == "HNX") {
    displayHNXData();
  }
  if (exchange == "UPCOM") {
    displayUpcomData();
  }
}

function removeAllCache() {
  var cache = CacheService.getScriptCache();
  cache.removeAll(['hoseData', 'hnxData', 'upcomData']);
}

*/


//Hàm định dạng giá trị Market Cap
function formatMarketCap(marketCap) {
  var suffix = "";
  if (marketCap >= (1e12 / 10000)) {
    marketCap /= 1e12 / 10000;
    suffix = "T";
  } else if (marketCap >= (1e9 / 10000)) {
    marketCap /= 1e9 / 10000;
    suffix = "B";
  } else if (marketCap >= (1e6 / 10000)) {
    marketCap /= 1e6 / 10000;
    suffix = "M";
  } else if (marketCap >= (1e3 / 10000)) {
    marketCap /= 1e3 / 10000;
    suffix = "K";
  }
  return marketCap.toFixed(2).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + suffix;
}

// Hàm định dạng giá trị Volume
function formatVolume(volume) {
  var suffix = "";
  if (volume >= (1e9 / 10000)) {
    volume /= 1e9 / 10000;
    suffix = "B";
  } else if (volume >= (1e6 / 10000)) {
    volume /= 1e6 / 10000;
    suffix = "M";
  } else if (volume >= (1e3 / 10000)) {
    volume /= 1e3 / 10000;
    suffix = "K";
  }
  return volume.toFixed(2).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + suffix;
}

/*
function displayTickerAndCompanyName() {
  var exchange = StockData.getRange('I4').getValue(); // Lấy tên sàn
  var page = StockData.getRange('E57').getValue();
  var stockList = getTickerAndNameByEx(exchange);
  
  var startRow = 2 + (page - 1) * 50;
  var endRow = startRow + 49;
  
  var tickerRange = StockData.getRange('A' + startRow + ':A' + endRow);
  var nameRange = StockData.getRange('B' + startRow + ':B' + endRow);
  
  var tickers = stockList[0].slice(50 * (page - 1), 50 * page);
  var names = stockList[1].slice(50 * (page - 1), 50 * page);
  
  var tickerValues = [];
  var nameValues = [];
  
  for (var i = 0; i < tickers.length; i++) {
    tickerValues.push([tickers[i]]);
    nameValues.push([names[i]]);
  }
  
  tickerRange.setValues(tickerValues);
  nameRange.setValues(nameValues);
}
*/

function getTickers() {
  var startRow = 2;
  var endRow = startRow + 49;

  var tickerRange = StockData.getRange('A' + startRow + ':A' + endRow);
  var tickerValues = tickerRange.getValues();

  var tickers = tickerValues.map(function (row) {
    return row[0];
  });
  return tickers;
}


function displayTickerAndCompanyName() {
  var exchange = StockData.getRange('I4').getValue(); // Lấy tên sàn
  var page = StockData.getRange('E57').getValue();
  //StockData.getRange('A2:B').clearContent();
  //SpreadsheetApp.flush();
  var stockList = getTickerAndNameByEx(exchange);
  for (var i = 0; i < 50; i++) {
    StockData.getRange('A' + (i + 2)).setValue(stockList[0][50 * (page - 1) + i]);
    StockData.getRange('B' + (i + 2)).setValue(stockList[1][50 * (page - 1) + i]);
  }
}

// function getTickers() {
//   var tickers = new Array();
//   for(var i = 0; i < 50; i++) {
//     tickers[i] = StockData.getRange('A' + (i+2)).getValue();
//   }
//   return tickers;
// }

function displayInforTicker() {
  var tickers = getTickers();
  // var dataList = [];
  for (var i = 0; i < 50; i++) {
    try {
      var dataByTicker = getDataByTicker(tickers[i]);
      // Logger.log(dataByTicker);
      // dataList.push(dataByTicker);
      var inforCompany = getInforCompanyByTicker(tickers[i]);
      var tickerInSheet = StockData.getRange('A' + (i + 2)).getValue();
      if (tickerInSheet != tickers[i]) break;
      var last = dataByTicker.data[0].cp;
      var chg = dataByTicker.data[0].delta1y;
      // Định dạng giá trị Market Cap
      var marketCap = dataByTicker.data[0].cp * inforCompany.outstandingShare;
      marketCap = formatMarketCap(marketCap);
      // Định dạng giá trị Volume
      var vol = dataByTicker.data[0].av;
      vol = formatVolume(vol);
      StockData.getRange('C' + (i + 2)).setValue(last);
      StockData.getRange('D' + (i + 2)).setValue(chg.toFixed(2) + '%');
      StockData.getRange('E' + (i + 2)).setValue(marketCap);
      StockData.getRange('F' + (i + 2)).setValue(vol);
    } catch (error) {
      Logger.log('Error with ticker ' + tickers[i] + ': ' + error.message);
      continue;
    }
  }
}

function displayInforTickerOptimize() {
  var tickers = getTickers();
  var lock = LockService.getScriptLock();

  lock.waitLock(5000); // Chờ cho đến khi có khóa

  for (var i = 0; i < 50; i++) {
    var dataByTicker;
    var inforCompany;

    try {
      lock.waitLock(3000); // Chờ cho đến khi có khóa sau mỗi lần lấy dữ liệu

      dataByTicker = getDataByTicker(tickers[i]);
      inforCompany = getInforCompanyByTicker(tickers[i]);

      var tickerInSheet = StockData.getRange('A' + (i + 2)).getValue();
      if (tickerInSheet != tickers[i]) break;

      var last = dataByTicker.data[0].cp;
      var chg = dataByTicker.data[0].delta1y;
      var marketCap = last * inforCompany.outstandingShare;
      marketCap = formatMarketCap(marketCap);
      var vol = dataByTicker.data[0].av;
      vol = formatVolume(vol);

      StockData.getRange('C' + (i + 2)).setValue(last);
      StockData.getRange('D' + (i + 2)).setValue(chg.toFixed(2) + '%');
      StockData.getRange('E' + (i + 2)).setValue(marketCap);
      StockData.getRange('F' + (i + 2)).setValue(vol);
    } catch (error) {
      Logger.log('Error with ticker ' + tickers[i] + ': ' + error.message);
      continue;
    } finally {
      lock.releaseLock(); // Giải phóng khóa sau mỗi lần lấy và ghi dữ liệu
    }
  }
}
























