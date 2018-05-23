function main() {

  var dateToday = new Date();
  var pastWeekendStartDate = new Date(dateToday.getFullYear(), dateToday.getMonth(), dateToday.getDate() - 3);
  var pastWeekendEndDate = new Date(dateToday.getFullYear(), dateToday.getMonth(), dateToday.getDate());
  
  var malvernStorePartiesCalendarID = "fizzkidz.com.au_j13ot3jarb1p9k70c302249j4g@group.calendar.google.com";
  var balwynStorePartiesCalendarID = "fizzkidz.com.au_7vor3m1efd3fqbr0ola2jvglf8@group.calendar.google.com";
  var travelPartiesCalendarID = "fizzkidz.com.au_b9aruprq8740cdamu63frgm0ck@group.calendar.google.com";
  
  var options = {timeMin : pastWeekendStartDate.toISOString(),
                 timeMax : pastWeekendEndDate.toISOString()};

  // malvern store parties
  var response = Calendar.Events.list(malvernStorePartiesCalendarID, options);
  for (var i = 0; i < response.items.length; i++) {
    var attachments = response.items[i].attachments;
    var bookingSheetID = attachments[0].fileId;
    sendFeedbackEmail(bookingSheetID);
  }
  
  // balwyn store parties
  var response = Calendar.Events.list(balwynStorePartiesCalendarID, options);
  for (var i = 0; i < response.items.length; i++) {
    var attachments = response.items[i].attachments;
    var bookingSheetID = attachments[0].fileId;
    sendFeedbackEmail(bookingSheetID);
  }
  
  // travel parties
  response = Calendar.Events.list(travelPartiesCalendarID, options);
  for (var i = 0; i < response.items.length; i++) {
    var attachments = response.items[i].attachments;
    var bookingSheetID = attachments[0].fileId;
    console.log(bookingSheetID);
    sendFeedbackEmail(bookingSheetID);
  }
}

function sendFeedbackEmail(bookingSheetID) {

  // Given that the spreadsheet was validated when it was created, validation does not need to take place.
  var sheet = SpreadsheetApp.openById(bookingSheetID).getActiveSheet();
  var parentName = sheet.getRange('B1').getDisplayValue();
  var emailAddress = sheet.getRange('B3').getDisplayValue();
  var childName = sheet.getRange('B4').getDisplayValue();

  var t = HtmlService.createTemplateFromFile('feedback_email_template');
  t.parentName = parentName;
  t.childName = childName;

  var body = t.evaluate().getContent();
  var subject = "We hope you enjoyed your party!";
  var signature = getGmailSignature();
  
  // Send the confirmation email
  GmailApp.sendEmail(emailAddress, subject, "", {htmlBody : body + signature, name : "Fizz Kidz"});
}

function getGmailSignature() {
  var draft = GmailApp.search("subject:signature label:draft", 0, 1);
  return draft[0].getMessages()[0].getBody();
}
