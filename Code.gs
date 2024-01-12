function sendEmail(){

  var emailAddress = 'alexktong.92@gmail.com'; // Input email address of recipient
  var timeZoneOffset = '+ 10' // Currently defined to align with timezone for Sydney, Australia
  
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Watchlist');
  var tickerNumberRng = ss.getRange('TickerNumber');
  var totalTickers = ss.getRange('TotalTickers').getValue();
  var date = new Date();
  var day = date.getDay();
  var dateString = Utilities.formatDate(date, 'GMT ' + timeZoneOffset, 'dd/MM/YY EEEE');

  var subjectString = '[Stock Watchlist Alert] ' + dateString; // appends string to end of subject title which provides handle to filter alerts in Gmail
  var messageString = '';
   
  var sortRange = ss.getRange('B3:D100');

  // Sends email alerts only on weekdays
  if (day >= 1 && day <= 5){
   
    // Top 5 falls
    messageString += 'Top 5 Daily Losers:\n';

    sortRange.sort({column: 4, ascending: true});

    for (var i = 1; i <= totalTickers; i++){
    
      // Set scenario to current ticker value
      tickerNumberRng.setValue(i);

      var dayChange = ss.getRange('DayChange').getValue();

      if (dayChange < 0) {
      
        var tickerDescription = ss.getRange('TickerDescription').getValue();
        var dayChangeString = Utilities.formatString('%2.2f%', dayChange * 100);
        var emailLine = tickerDescription + ': ' + dayChangeString + '\n';

        messageString += emailLine;
      }
    }

    // Top 5 rises
    messageString += '\nTop 5 Daily Risers:\n';

    sortRange.sort({column: 4, ascending: false});

    for (var i = 1; i <= totalTickers; i++){
    
      // Set scenario to current ticker value
      tickerNumberRng.setValue(i);

      var dayChange = ss.getRange('DayChange').getValue();

      if (dayChange > 0){
      
        var tickerDescription = ss.getRange('TickerDescription').getValue();
        var dayChangeString = Utilities.formatString('%2.2f%', dayChange * 100);
        var emailLine = tickerDescription + ': ' + dayChangeString + '\n';

        messageString += emailLine;
      }
    }
    
    // Tickers % from 52-week low
    messageString += '\n% From 52-Week Low:\n'
    
    for (var i = 1; i <= totalTickers; i++){
    
      // Set scenario to current ticker value
      tickerNumberRng.setValue(i);

      var pctLow = ss.getRange('PctLow').getValue();
      var tickerDescription = ss.getRange('TickerDescription').getValue();
      var pctLowString = Utilities.formatString('%2.2f%', pctLow * 100);
      var emailLine = tickerDescription + ': ' + pctLowString + '\n';
        
      messageString += emailLine;

    }
    // Email alert
    MailApp.sendEmail(emailAddress, subjectString, messageString);
  }
}