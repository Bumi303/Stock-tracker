function myPortfolio() {
  var apikey = 'DEHUNP5Q6X2OFW16'

  var ss= SpreadsheetApp.getActiveSpreadsheet();
  var stockSheet = ss.getSheetByName("Stocks");

  stockSheet.getRange('B12:E26').clear();

  var lastRow = stockSheet.getLastRow();

  for (var j = 12; j <= lastRow; j++)
  {
    var symbol = stockSheet.getRange(j, 1).getValue();

    var URL_STRING = "https://www.alphavantage.co/query?function=TIME_SERIES_DAILY&symbol=" + symbol + "&apikey=" + apikey;

    var response = UrlFetchApp.fetch(URL_STRING);
    var json = response.getContentText();
    var data = JSON.parse(json);

    for(const [key, value] of Object.entries(data)){
      for (const [key2, value2] of Object.entries(value)){
        var date = key2

        for (const [key3, value3] of Object.entries(value2)){
          if(key3 == '1. open')
          {
            var open = value3
          }
          if(key3 == '4. close')
          {
            var close = value3
          }
          if(key3 == '2. high')
          {
            var volume = value3
          }
        }
      break;
      }
    }
    stockSheet.getRange(j,2).setValue(date).setFontSize(10);
    stockSheet.getRange(j,3).setValue(open).setFontSize(10);
    stockSheet.getRange(j,4).setValue(close).setFontSize(10);
    stockSheet.getRange(j,5).setValue(volume).setFontSize(10);
  }
}

