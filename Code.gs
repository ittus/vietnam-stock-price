function getStockPrice(symbols) {
  var response = UrlFetchApp.fetch("https://priceservice.vndirect.com.vn/priceservice/secinfo/snapshot/q=codes:" + symbols)
  var json = response.getContentText()
  var data = JSON.parse(json)
  var res = {}
  for (var idx = 0; idx < symbols.length; idx++) {
    var value = data[idx]
    if (!value) { continue }
    const allValues = data[idx].split("|")
    if (allValues.length < 13){ continue }
    res[symbols[idx]] = data[idx].split("|")[13]
  }
  return res
}


function getAllPrices() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Portfolio")
  var symbols = sheet.getRange("A3:A100").getValues()
  symbols = symbols.filter(function(s) {
    return s[0].length > 0
  }).map(function(s) {
    return s[0]
  })

  Logger.log(symbols)

  var prices = getStockPrice(symbols)
  Logger.log(prices)
  var startRow = 3;
  for (var idx = 0; idx < symbols.length; idx++) {
    var row = startRow + idx;
    sheet.getRange(row, 6, 1, 1).setValue(prices[symbols[idx]] * 1000)
  }
  sheet.getRange(1, 1, 1, 1).setValue("Last update: " + (new Date()))
}

function onOpen() {
  getAllPrices();
  SpreadsheetApp.getUi()
  .createMenu('Stock Utils')
  .addItem('Refresh', 'getAllPrices')
  .addToUi();

}