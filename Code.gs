function sendEmail(){

  var emailAddress = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX";
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Watchlist");
  var tickerNumberRng = ss.getRange("TickerNumber");
  var totalTickers = ss.getRange("TotalTickers").getValue();
  var date = new Date();
  var day = date.getDay();
  var dateString = Utilities.formatDate(date, "GMT + 1", "dd/MM/YY EEEE"); // e.g. 04/03/22 Friday for Sydney, Australia timezone

  var subjectString = dateString + " [Watchlist Alert]"; // appends string to end of subject title which provides handle to filter alerts in Gmail
  var messageString = "";
   
  var sortRange = ss.getRange("B3:D100");

  // Sends email alerts only on weekdays
  if (day >= 1 && day <= 5){
   
    // Top 5 falls
    messageString += "Top 5 Daily Losers:\n";

    sortRange.sort({column: 4, ascending: true});

    for (var i = 1; i <= 5; i++){
    
      // Set scenario to current ticker value
      tickerNumberRng.setValue(i);

      var ticker = ss.getRange("Ticker").getValue();
      var dayChange = ss.getRange("DayChange").getValue();
      var dayChangeString = Utilities.formatString("%2.2f%", dayChange * 100);
      var emailLine = ticker + ": " + dayChangeString + "\n";

      messageString += emailLine;
    }

    // Top 5 rises
    messageString += "\nTop 5 Daily Risers:\n";

    sortRange.sort({column: 4, ascending: false});

    for (var i = 1; i <= 5; i++){
    
      // Set scenario to current ticker value
      tickerNumberRng.setValue(i);

      var ticker = ss.getRange("Ticker").getValue();
      var dayChange = ss.getRange("DayChange").getValue();
      var dayChangeString = Utilities.formatString("%2.2f%", dayChange * 100);
      var emailLine = ticker + ": " + dayChangeString + "\n";

      messageString += emailLine;
    }
    
    // Tickers that are < 10% from 52-week low
    
    messageString += "\n% From 52-Week Low:\n"
    
    for (var i = 1; i <= totalTickers; i++){
    
      // Set scenario to current ticker value
      tickerNumberRng.setValue(i);

      var pctLow = ss.getRange("PctLow").getValue();

      if (pctLow < 0.10){

        var ticker = ss.getRange("Ticker").getValue();
        var pctLowString = Utilities.formatString("%2.2f%", pctLow * 100);
        var emailLine = ticker + ": " + pctLowString + "\n";
        
        messageString += emailLine;
      }
    }
    // Email alert
    MailApp.sendEmail(emailAddress, subjectString, messageString);
  }
}