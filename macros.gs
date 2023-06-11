
function nextPage() {
  StockData.getRange('E57').setValue(StockData.getRange('E57').getValue() + 1);
  displayTickerAndCompanyName();
  displayInforTicker();
  //displayInforTickerOptimize();
}

function backPage() {
  var range = StockData.getRange('E57');
  var value = range.getValue();
  if (value != 1) {
    range.setValue(value - 1);
    displayTickerAndCompanyName();
    displayInforTicker();
  //displayInforTickerOptimize();
  }
}
/*
function onEdit(e) {
  var range = e.range;
  var sheet = range.getSheet();
  if (sheet.getName() == 'Stock Data' && range.getA1Notation() == 'I4') {
    StockData.getRange('E57').setValue(parseInt('1'));
    display();
  }
  if (sheet.getName() == 'Stock Data' && range.getA1Notation() == 'E57') {
    display();
  }
}
*/
function createTimeDrivenTriggers() {
  ScriptApp.newTrigger(crawlHoseData)
    .timeBased()
    .everyHours(5)
    .create();
  ScriptApp.newTrigger(crawlHNXData)
    .timeBased()
    .everyHours(5)
    .create();
  ScriptApp.newTrigger(crawlUpcomData)
    .timeBased()
    .everyHours(5)
    .create();
}



function createCustomOnEditTrigger(e) {
  var range = e.range;
  var sheet = range.getSheet();
  if (sheet.getName() == 'Stock Data' && range.getA1Notation() == 'I4') {
    StockData.getRange('E57').setValue(parseInt('1'));
    StockData.getRange('A2:F51').clearContent();
    SpreadsheetApp.flush();
    displayTickerAndCompanyName();
    displayInforTicker();
    //displayInforTickerOptimize();
  }
  if (sheet.getName() == 'Stock Data' && range.getA1Notation() == 'E57') {
    StockData.getRange('A2:F51').clearContent();
    SpreadsheetApp.flush();
    displayTickerAndCompanyName();
    displayInforTicker();
    //displayInforTickerOptimize();
  }
  if (sheet.getName() == 'Infor Tickets' && range.getA1Notation() == 'B4') {
    fetchOrganInfor();
    displayDetailTicker();
    displayRelevantPriceTicker();
  }
  if (sheet.getName() == 'Infor Tickets' && range.getA1Notation() == 'G17') {
    displayDetailTicker();
  }
}

function createCustomOnOpenTrigger() {
    fetchOrganInfor();
    displayRelevantPriceTicker();
    displayDetailTicker();
    StockData.getRange('I4').setValue('HOSE');
    StockData.getRange('E57').setValue(parseInt('1'));
    displayTickerAndCompanyName();
    displayInforTicker();
    //displayInforTickerOptimize();

}

function createSpreadSheetEditTrigger() {
  const ssE = SpreadsheetApp.getActive();
  ScriptApp.newTrigger('createCustomOnEditTrigger')
    .forSpreadsheet(ssE)
    .onEdit()
    .create();
}

function createSpreadSheetOpenTrigger() {
  const ssO = SpreadsheetApp.getActive();
  ScriptApp.newTrigger('createCustomOnOpenTrigger')
    .forSpreadsheet(ssO)
    .onOpen()
    .create();
}

function deleteAllTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
}

