/// Helpful: https://developers.google.com/apps-script/guides/support/troubleshooting
/// Seems to answer issue of Data Validation + getLast Row: https://stackoverflow.com/questions/10887461/issue-with-getlastrow

/* Goals of this Spreadsheet:
  - [Decided] start off simple
      - email auto reminders about moving up from Visitor Status to Exploratory Status
  - Extra functionality for down the line
      - email specificed 'buddy' at 11.5 month mark (if we want to personally ask 
        for an increase in membership status
      - [Possibly] Keep track of dues. 

  */
  
  /* TODO:
  - look up how to format emails sent by this project
  - Ask for input: Can someone write a draft of an email that would be sent at the intervals asking for 
    the Visitor to move up? 
  - Consider: ???? This on the same account as EMS? Could automate the input of data as folks fill out EMS ????
  - [NEEDS FIX or NOTING] If JOIN Date is edited, intervals will not be updated, b/c they are not blank
  */


/* 
Specs. It will... 
 * Populate the 4, 9, and 12 months dates IFF (Join Col. has a Date && Col's E, F, G Don't have data)
   - get data from the D col cell 
   - update E, F, G col's to be the appropriate dates
 * email the addresses in the email column
 * use the name fields to create a form email
 * it will calculate today's date
 * it will check daily to see if today's date === the 4, 9, 12 month date
 * it will email them only if so 
 * It will email Hope with a msg to remove them from the Visitor list (or
 else just a note that they signed up 12 months ago, and a question: are they EM's yet?
 
*/


function populateReminderDates(joinCellRow) {
 
  var joinCellCol = 4;
  
  var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var joined = activeSheet.getRange(joinCellRow, joinCellCol).getValue();
  var joinYear = joined.getYear();
  var joinMonth = joined.getMonth();
  var joinDay = joined.getDate();
  
  var afterFourMonths = new Date(joinYear, joinMonth + 4, joinDay);
  var afterNineMonths = new Date(joinYear, joinMonth + 9, joinDay);
  var afterElevenAndHalfMonths = new Date(joinYear, joinMonth + 11, joinDay + 14);
  
  activeSheet.getRange(joinCellRow, joinCellCol + 1).setValue(afterFourMonths);
  activeSheet.getRange(joinCellRow, joinCellCol + 2).setValue(afterNineMonths);
  activeSheet.getRange(joinCellRow, joinCellCol + 3).setValue(afterElevenAndHalfMonths);
  
};

function ifJoinedThenPopulateDates() {

  var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var JOINED_COL = 4;
  
  // loop through rows and check if JoinCol (D or 4) is blank
  var lastRow = activeSheet.getLastRow();

  for (var i = 3; i <= lastRow; ++i) {
    
    // if (JOIN Date is filled out, but not the other dates) {run populateReminderDates(where i = joineCellRow)}
    if(!(activeSheet.getRange(i, JOINED_COL).isBlank()) && (activeSheet.getRange(i, JOINED_COL + 1).isBlank())) {
      populateReminderDates(i);
    }
  }
  
};

function onEdit(e) {
  ifJoinedThenPopulateDates();
}

function sendEmail(activeRow) {
  var FIRST_NAME_COL = 1;
  var LAST_NAME_COL = 2;
  var EMAIL_COL = 3;
  
  var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  var firstName = activeSheet.getRange(activeRow, FIRST_NAME_COL).getValue();
  var lastName = activeSheet.getRange(activeRow, LAST_NAME_COL).getValue();
  var email = activeSheet.getRange(activeRow, EMAIL_COL).getValue();
  
  var subject = "Move Into Exploratory Status"
  var body = "Dear " + firstName + ", You've been a Visiting Member for so many months. Would you like to" +
    " become an Exploratory Member?"
  MailApp.sendEmail(email, subject, body)
}

function emailIfIntervalDateMatches() {
  var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = activeSheet.getLastRow();
  var today = new Date();
  var todayYear = today.getYear();
  var todayMonth = today.getMonth() + 1;
  var todayDay = today.getDate();
   
  // loop through rows
     // loop through each calculated date (columns)
  
  for (var row = 3; row <= lastRow; ++row) {
    if (shouldThePersonInThisRowBeEmailed(row)) {
      sendEmail(row)
    }
  }
}

function shouldThePersonInThisRowBeEmailed(row) {
  var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var today = new Date();
  var todayYear = today.getYear();
  var todayMonth = today.getMonth() + 1;
  var todayDay = today.getDate();
  
  for (var col = 5; col <= 7; ++col) {
    var dateInCell = activeSheet.getRange(row, col).getValue();
    var dateInCellYear = dateInCell.getYear();
    var dateInCellMonth = dateInCell.getMonth() + 1;
    var dateInCellDay = dateInCell.getDate();
    
    if ((todayYear === dateInCellYear) && (todayMonth === dateInCellMonth) && (todayDay === dateInCellDay)) {
      return true;
    }
  }
}




