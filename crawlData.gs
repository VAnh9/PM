// link gốc: https://fiin-core.ssi.com.vn/Master/GetListOrganization?language=en
//lay du lieu danh sach ma chung khoan
function getListData() {
    var url = "https://raw.githubusercontent.com/VAnh9/PM/main/data.json";
    var response = UrlFetchApp.fetch(url);
    var data =  JSON.parse(response);
    Logger.log(data);
    return data;
}



//lay ma co phieu va ten cong ty theo san chung khoan
// lay ma co phieu va ten cong ty theo san chung khoan
function getTickerAndNameByEx(exchange) {
    var comGroupCode;
    if (exchange == "UPCOM") comGroupCode = "UpcomIndex";
    if (exchange == "HOSE") comGroupCode = "VNINDEX";
    if (exchange == "HNX") comGroupCode = "HNXIndex";

    var stockList = getListData();
    let result = new Array();
    let companyNameList = new Array();
    let tickerList = new Array();

    for (var i = 0; i < stockList.totalCount; i++) {
        if (stockList.items[i].comGroupCode == comGroupCode) {
            tickerList.push(stockList.items[i].ticker);
            companyNameList.push(stockList.items[i].organShortName);
        }
    }
    result.push(tickerList);
    result.push(companyNameList);
    return result;
}
/*
function getTickerAndNameByEx(exchange) {
  const comGroupCode = {
    UPCOM: "UpcomIndex",
    HOSE: "VNINDEX",
    HNX: "HNXIndex"
  };
  const stockList = getListData();
  const result = {
    tickerList: [],
    companyNameList: []
  };
  for (let i = 0; i < stockList.totalCount; i++) {
    const stockItem = stockList.items[i];
    if (stockItem.comGroupCode === comGroupCode[exchange]) {
      result.tickerList.push(stockItem.ticker);
      result.companyNameList.push(stockItem.organShortName);
    }
  }
  return result;
}
*/
function getDataByTicker(ticker) {
    var response = UrlFetchApp.fetch("https://apipubaws.tcbs.com.vn/stock-insight/v1/stock/second-tc-price?tickers=" + ticker);
    return JSON.parse(response);
}


function getHistoricalDataByTicker(ticker, timeStart, timeEnd) {
    var data = UrlFetchApp.fetch("https://apipubaws.tcbs.com.vn/stock-insight/v1/stock/bars-long-term?ticker=" + ticker +"&type=stock&resolution=D&from=" + timeStart + "&to=" + timeEnd);
    return JSON.parse(data);
}

function getInforCompanyByTicker(ticker) {
    var data = UrlFetchApp.fetch("https://apipubaws.tcbs.com.vn/tcanalysis/v1/ticker/" + ticker + "/overview");
    return JSON.parse(data);
}



// crawl HOSE data

async function crawlHoseData() {
  var hoseTickers = getTickerAndNameByEx('HOSE')[0];
  var hoseCompanyName = getTickerAndNameByEx('HOSE')[1];
  var hoseData = {}
  for(var i = 0; i < hoseTickers.length; i++) {
    try {
      var ticker = hoseTickers[i];
      var dataByTicker = await getDataByTicker(ticker);
      var inforCompany = await getInforCompanyByTicker(ticker);

      var tickerName = ticker;
      var companyName = hoseCompanyName[i];
      var last = dataByTicker.data[0].cp;
      var chg = dataByTicker.data[0].delta1y;
      // Định dạng giá trị Market Cap
      var marketCap = dataByTicker.data[0].cp * inforCompany.outstandingShare;
      marketCap = formatMarketCap(marketCap);
      // Định dạng giá trị Volume
      var vol = dataByTicker.data[0].av;
      vol = formatVolume(vol);

      hoseData[ticker] = {
        tickerName: tickerName,
        companyName: companyName,
        last: last,
        chg: chg,
        marketCap: marketCap,
        vol: vol
      };
    } catch(error) {
        Logger.log(error.message);
        continue;
    }
  }
  var cache = CacheService.getScriptCache();
  cache.put("hoseData", JSON.stringify(hoseData), 18000);
}

// crawl HNX data
function crawlHNXData() {
  var hnxTickers = getTickerAndNameByEx('HNX')[0];
  var hnxCompanyName = getTickerAndNameByEx('HNX')[1];
  var hnxData = {};

  for(var i = 0; i < hnxTickers.length; i++) {
    try {
      var ticker = hnxTickers[i];
      var dataByTicker = getDataByTicker(ticker);
      var inforCompany = getInforCompanyByTicker(ticker);
      
      var tickerName = ticker;
      var companyName = hnxCompanyName[i];
      var last = dataByTicker.data[0].cp;
      var chg = dataByTicker.data[0].delta1y;
      // Định dạng giá trị Market Cap
      var marketCap = dataByTicker.data[0].cp * inforCompany.outstandingShare;
      marketCap = formatMarketCap(marketCap);
      // Định dạng giá trị Volume
      var vol = dataByTicker.data[0].av;
      vol = formatVolume(vol);

      hnxData[ticker] = {
        tickerName: tickerName,
        companyName: companyName,
        last: last,
        chg: chg,
        marketCap: marketCap,
        vol: vol
      };
    } catch(error) {
        Logger.log(error.message);
        continue;
    }
  }
  var cache = CacheService.getScriptCache();
  cache.put("hnxData", JSON.stringify(hnxData), 18000);
}

// crawl UPCOM data
function crawlUpcomData() {
  var upcomTickers = getTickerAndNameByEx('UPCOM')[0];
  var upcomCompanyName = getTickerAndNameByEx('UPCOM')[1];
  var upcomData = {};

  for(var i = 0; i < 400; i++) {
    try {
      var ticker = upcomTickers[i];
      var dataByTicker = getDataByTicker(ticker);
      var inforCompany = getInforCompanyByTicker(ticker);
      
      var tickerName = ticker;
      var companyName = upcomCompanyName[i];
      var last = dataByTicker.data[0].cp;
      var chg = dataByTicker.data[0].delta1y;
      // Định dạng giá trị Market Cap
      var marketCap = dataByTicker.data[0].cp * inforCompany.outstandingShare;
      marketCap = formatMarketCap(marketCap);
      // Định dạng giá trị Volume
      var vol = dataByTicker.data[0].av;
      vol = formatVolume(vol);

      upcomData[ticker] = {
        tickerName: tickerName,
        companyName: companyName,
        last: last,
        chg: chg,
        marketCap: marketCap,
        vol: vol
      };
    } catch(error) {
        Logger.log(error.message);
        continue;
    }
  }
    var cache = CacheService.getScriptCache();
  cache.put("upcomData", JSON.stringify(upcomData), 18000);
}


