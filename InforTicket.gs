const InforTicket = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Infor Tickets');



// Sử dụng hàm getOrganNameByCode
function fetchOrganInfor() {
  var OrganCode = InforTicket.getRange("B4").getValue();
  var OrganInfor = getInforCompanyByTicker(OrganCode);
  InforTicket.getRange("B7:G7").setValue(OrganInfor.shortName);
  InforTicket.getRange("B8").setValue(OrganInfor.exchange);
  InforTicket.getRange("B9").setValue(OrganInfor.stockRating);
  InforTicket.getRange("B10").setValue(OrganInfor.website);
  InforTicket.getRange("E8").setValue(OrganInfor.outstandingShare);
  InforTicket.getRange("E9").setValue(OrganInfor.industry);
  InforTicket.getRange("E10").setValue(OrganInfor.foreignPercent);
}

function displayRelevantPriceTicker() {
  var OrganCode = InforTicket.getRange("B4").getValue().toUpperCase();
  var OrganInfor = getDataByTicker(OrganCode);
  InforTicket.getRange("B12:C12").setValue(OrganInfor.data[0].cp);
  InforTicket.getRange("D12").setValue(OrganInfor.data[0].delta1y);
  InforTicket.getRange("E12").setValue((OrganInfor.data[0].delta1y / OrganInfor.data[0].cp) * 100);
  InforTicket.getRange("B16").setValue(OrganInfor.data[0].cp * OrganInfor.data[0].oscore);
  InforTicket.getRange("C16").setValue(OrganInfor.data[0].pe);
  InforTicket.getRange("D16").setValue(OrganInfor.data[0].vnipe); 
}


function displayDetailTicker() {
  var OrganCode = InforTicket.getRange("B4").getValue().toUpperCase();
  var timeChoose = InforTicket.getRange("G17").getValue();
  var d = new Date();
  var timeEnd = parseInt(d.getTime() / 1000);
  var timeStart = 0;
  InforTicket.getRange("B30:G").clearContent();
  SpreadsheetApp.flush();
  if(timeChoose == "1 Month") {
    timeStart = timeEnd - (24 * 60 * 60 * 30);
    var detailInforTicker = getHistoricalDataByTicker(OrganCode,timeStart,timeEnd);
    var dataLength = detailInforTicker.data.length - 1;
    var j = -1;
    for(var i = dataLength; i >= 0; i--) {
      j++;
      InforTicket.getRange("B" +(30 + j )).setValue(detailInforTicker.data[i].tradingDate.substring(0,10));
      InforTicket.getRange("C" +(30 + j)).setValue(detailInforTicker.data[i].close);
      InforTicket.getRange("D" +(30 + j)).setValue(detailInforTicker.data[i].high);
      InforTicket.getRange("E" +(30 + j)).setValue(detailInforTicker.data[i].low);
      InforTicket.getRange("F" +(30 + j)).setValue(detailInforTicker.data[i].open);
      InforTicket.getRange("G" +(30 + j)).setValue(detailInforTicker.data[i].volume);  
    }
  }
  if(timeChoose == "3 Months") {
    timeStart = timeEnd - (24 * 60 * 60 * 30 * 3);
    var detailInforTicker = getHistoricalDataByTicker(OrganCode,timeStart,timeEnd);
    var dataLength = detailInforTicker.data.length - 1;
    var j = -1;
    for(var i = dataLength; i >= 0; i--) {
      j++;
      InforTicket.getRange("B" +(30 + j )).setValue(detailInforTicker.data[i].tradingDate.substring(0,10));
      InforTicket.getRange("C" +(30 + j)).setValue(detailInforTicker.data[i].close);
      InforTicket.getRange("D" +(30 + j)).setValue(detailInforTicker.data[i].high);
      InforTicket.getRange("E" +(30 + j)).setValue(detailInforTicker.data[i].low);
      InforTicket.getRange("F" +(30 + j)).setValue(detailInforTicker.data[i].open);
      InforTicket.getRange("G" +(30 + j)).setValue(detailInforTicker.data[i].volume);  
    }
  }

  if(timeChoose == "6 Months") {
    timeStart = timeEnd - (24 * 60 * 60 * 30 * 6);
    var detailInforTicker = getHistoricalDataByTicker(OrganCode,timeStart,timeEnd);
    var dataLength = detailInforTicker.data.length - 1;
    var j = -1;
    for(var i = dataLength; i >= 0; i--) {
      j++;
      InforTicket.getRange("B" +(30 + j )).setValue(detailInforTicker.data[i].tradingDate.substring(0,10));
      InforTicket.getRange("C" +(30 + j)).setValue(detailInforTicker.data[i].close);
      InforTicket.getRange("D" +(30 + j)).setValue(detailInforTicker.data[i].high);
      InforTicket.getRange("E" +(30 + j)).setValue(detailInforTicker.data[i].low);
      InforTicket.getRange("F" +(30 + j)).setValue(detailInforTicker.data[i].open);
      InforTicket.getRange("G" +(30 + j)).setValue(detailInforTicker.data[i].volume);  
    }
  }

  if(timeChoose == "12 Months") {
    timeStart = timeEnd - (24 * 60 * 60 * 30 * 12);
    var detailInforTicker = getHistoricalDataByTicker(OrganCode,timeStart,timeEnd);
    var dataLength = detailInforTicker.data.length - 1;
    var j = -1;
    for(var i = dataLength; i >= 0; i--) {
      j++;
      InforTicket.getRange("B" +(30 + j )).setValue(detailInforTicker.data[i].tradingDate.substring(0,10));
      InforTicket.getRange("C" +(30 + j)).setValue(detailInforTicker.data[i].close);
      InforTicket.getRange("D" +(30 + j)).setValue(detailInforTicker.data[i].high);
      InforTicket.getRange("E" +(30 + j)).setValue(detailInforTicker.data[i].low);
      InforTicket.getRange("F" +(30 + j)).setValue(detailInforTicker.data[i].open);
      InforTicket.getRange("G" +(30 + j)).setValue(detailInforTicker.data[i].volume);  
    }
  }

  if(timeChoose == "36 Months") {
    timeStart = timeEnd - (24 * 60 * 60 * 30 * 36);
    var detailInforTicker = getHistoricalDataByTicker(OrganCode,timeStart,timeEnd);
    var dataLength = detailInforTicker.data.length - 1;
    var j = -1;
    for(var i = dataLength; i >= 0; i--) {
      j++;
      InforTicket.getRange("B" +(30 + j )).setValue(detailInforTicker.data[i].tradingDate.substring(0,10));
      InforTicket.getRange("C" +(30 + j)).setValue(detailInforTicker.data[i].close);
      InforTicket.getRange("D" +(30 + j)).setValue(detailInforTicker.data[i].high);
      InforTicket.getRange("E" +(30 + j)).setValue(detailInforTicker.data[i].low);
      InforTicket.getRange("F" +(30 + j)).setValue(detailInforTicker.data[i].open);
      InforTicket.getRange("G" +(30 + j)).setValue(detailInforTicker.data[i].volume);  
    }
  }
}












































