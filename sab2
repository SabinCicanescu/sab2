var REMINDER_SENT = "REMINDER_SENT";

function visitorsBadge() {

var currentDate = new Date();
var currentDay = currentDate.getDate();
var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Tracker"); //change with your Sheet name
  var startRow = 2;
  var dataRange = sheet.getRange(startRow, 1, sheet.getLastRow()-1, 9); // change with desired column
  var data = dataRange.getValues();
 
  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
    var emailAddress = row[1];
    var subject = "Reminder to return the visitor badge"; //change with your subject
    var message = "Please return the visitor badge by the end of the day."; //change with your message
    message = "\r\nHello," + "\r\n\r\n" + message + "\r\n\r\n" + "Best regards!"; //change with your message
    var row_input = row[7];
    var eventDateFormat = Utilities.formatDate(new Date(row[0]), "GMT+2", "MM/dd/yyyy");
    var eventDate = new Date(eventDateFormat);
    var dayOfEvent = eventDate.getDate();

    try {
      if (emailAddress != "" && row_input === "" && currentDay == dayOfEvent + 7) {

        GmailApp.sendEmail(emailAddress, subject, message);
        sheet.getRange(startRow + i, 9).setValue(REMINDER_SENT);
      }
      
    } catch(e) {
      console.log("ERROR :" + "Error message is " + e.message);
      
    }
  }
}

