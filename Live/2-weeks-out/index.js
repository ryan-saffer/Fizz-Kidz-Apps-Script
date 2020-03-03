function test() {

  var parentName = "Ryan"
  var mobileNumber = "0413892120"
  var emailAddress = "ryansaffer@gmail.com"
  var childName = "Jimmy"
  var childAge = "5"
  var dateOfParty = new Date("10/11/2019")
  var timeOfParty = new Date(2019, 09, 10, 10, 00)
  var partyLength = "2"
  var partyType = "In-store"  
  var location = "Malvern"
  
  // Determine the start and end times of the party
  var startDate = new Date(
    dateOfParty.getFullYear(),
    dateOfParty.getMonth(),
    dateOfParty.getDate(),
    timeOfParty.getHours(),
    timeOfParty.getMinutes()
    );
  var endDate = determineEndDate(dateOfParty, timeOfParty, partyLength);
  
  // Using the HTML email template, inject the variables and get the content
  var t = HtmlService.createTemplateFromFile('email_template');
  t.parentName = parentName;
  t.childName = childName;
  t.childAge = childAge;
  t.startDate = buildFormattedStartDate(startDate)
  t.startTime = Utilities.formatDate(startDate, 'Australia/Sydney', 'hh:mm a');
  t.endTime = Utilities.formatDate(endDate, 'Australia/Sydney', 'hh:mm a');
  
  // determine location
  var updated_location = location;
  if (partyType == "In-store") {
    updated_location = (location == "Malvern") ? "our Malvern store" : "our Balwyn store";
  }
  t.location = updated_location;
  
  var body = t.evaluate().getContent();
  var subject = childName + "'s party is coming up - what you need to know"

  // determine the from email address
  var fromAddress = determineFromEmailAddress(location);

  var signature = getGmailSignature();
  
  // Send the confirmation email
  GmailApp.sendEmail(emailAddress, subject, "", {from: fromAddress, htmlBody : body + signature, name : "Fizz Kidz"});
}


/**
 * triggered weekly Thursday 9:34am
 * email parents with parties on following Friday+
 * update 'This Weekends Parties' folder in Drive
 */

function main() {
  
  var dateToday = new Date();
  var startDate  = new Date(
    dateToday.getFullYear(),
    dateToday.getMonth(),
    dateToday.getDate() + 15
    )
  var endDate = new Date(
    dateToday.getFullYear(),
    dateToday.getMonth(),
    dateToday.getDate() + 18
    )
  
  var malvernStorePartiesCalendarID = "fizzkidz.com.au_j13ot3jarb1p9k70c302249j4g@group.calendar.google.com";
  var balwynStorePartiesCalendarID = "fizzkidz.com.au_7vor3m1efd3fqbr0ola2jvglf8@group.calendar.google.com";
  var travelPartiesCalendarID = "fizzkidz.com.au_b9aruprq8740cdamu63frgm0ck@group.calendar.google.com";
  
  var options = {
    timeMin : startDate.toISOString(),
    timeMax : endDate.toISOString()
    };
  
  // malvern store parties
  var response = Calendar.Events.list(malvernStorePartiesCalendarID, options);
  for (var i = 0; i < response.items.length; i++) {
    var attachments = response.items[i].attachments;
    var bookingSheetID = attachments[0].fileId;
    sendEmail(bookingSheetID);
  }
  
  // balwyn store parties
  var response = Calendar.Events.list(balwynStorePartiesCalendarID, options);
  for (var i = 0; i < response.items.length; i++) {
    var attachments = response.items[i].attachments;
    var bookingSheetID = attachments[0].fileId;
    sendEmail(bookingSheetID);
  }
  
  // travel parties
  response = Calendar.Events.list(travelPartiesCalendarID, options);
  for (var i = 0; i < response.items.length; i++) {
    var attachments = response.items[i].attachments;
    var bookingSheetID = attachments[0].fileId;
    console.log(bookingSheetID);
    sendEmail(bookingSheetID);
  }
}

function sendEmail(bookingSheetID) {
  
  var sheet = SpreadsheetApp.openById(bookingSheetID).getActiveSheet();
  var parentName = sheet.getRange('B1').getDisplayValue();
  var mobileNumber = sheet.getRange('B2').getDisplayValue();
  var emailAddress = sheet.getRange('B3').getDisplayValue();  
  var childName = sheet.getRange('B4').getDisplayValue();  
  var childAge = sheet.getRange('B5').getDisplayValue();  
  var dateOfParty = sheet.getRange('B6').getValue();  
  var timeOfParty = sheet.getRange('B7').getValue(); 
  var partyLength = sheet.getRange('B8').getDisplayValue();
  var partyType = sheet.getRange('B10').getDisplayValue();  
  var location = sheet.getRange('B11').getDisplayValue();
  
  // Determine the start and end times of the party
  var startDate = new Date(
    dateOfParty.getFullYear(),
    dateOfParty.getMonth(),
    dateOfParty.getDate(),
    timeOfParty.getHours(),
    timeOfParty.getMinutes()
    );
  var endDate = determineEndDate(dateOfParty, timeOfParty, partyLength);
  
  // Using the HTML email template, inject the variables and get the content
  var t = HtmlService.createTemplateFromFile('email_template');
  t.parentName = parentName;
  t.childName = childName;
  t.childAge = childAge;
  t.startDate = buildFormattedStartDate(startDate)
  t.startTime = Utilities.formatDate(startDate, 'Australia/Sydney', 'hh:mm a');
  t.endTime = Utilities.formatDate(endDate, 'Australia/Sydney', 'hh:mm a');
  
  // determine location
  var updated_location = location;
  if (partyType == "In-store") {
    updated_location = (location == "Malvern") ? "our Malvern store" : "our Balwyn store";
  }
  t.location = updated_location;
  
  var body = t.evaluate().getContent();
  var subject = childName + "'s party is coming up - what you need to know"

  // determine the from email address
  var fromAddress = determineFromEmailAddress(location);

  var signature = getGmailSignature();
  
  // Send the confirmation email
  GmailApp.sendEmail(emailAddress, subject, "", {from: fromAddress, htmlBody : body + signature, name : "Fizz Kidz"});
}

function determineFromEmailAddress(location) {
  /**
   * If location is Malvern send from malvern@fizzkidz.com.au
   * If location is Balwyn send from info@fizzkidz.com.au
   * If location is neither (mobile) send from info@fizzkidz.com.au
   */
  
  if (location == "Malvern") {
    return "malvern@fizzkidz.com.au";
  }
  else if (location == "Balwyn") {
    return "info@fizzkidz.com.au";
  }
  else {
    return "info@fizzkidz.com.au";
  }
}

function determineEndDate(dateOfParty, timeOfParty, partyLength) {
  
  // determine when party ends
  var lengthHours = 0;
  var lengthMinutes = 0;
  switch (partyLength) {
    case "1":
      lengthHours = 1;
      break;
    case "1.5":
      lengthHours = 1;
      lengthMinutes = 30;
      break;
    case "2":
      lengthHours = 2;
      break;
    default:
      break;
  }
  
  var endDate = new Date(dateOfParty.getFullYear(), dateOfParty.getMonth(), dateOfParty.getDate(), timeOfParty.getHours() + lengthHours, timeOfParty.getMinutes() + lengthMinutes);
  
  return endDate;
}

function getGmailSignature() {
  var draft = GmailApp.search("subject:talia-signature label:draft", 0, 1);
  return draft[0].getMessages()[0].getBody();
}

/**
 * Returns the start date as a formatted string
 * 
 * @param {Date} date - the start date of the party
 * @returns {String} the date formatted as 'Friday 3rd June 2019'
 */
function buildFormattedStartDate(date) {

  var suffix = determineSuffix(date.getDate())
  var dayOfWeekAndMonth = Utilities.formatDate(date, 'Australia/Sydney', "EEEE d");
  var monthAndYear = Utilities.formatDate(date, 'Australia/Sydney',"MMMM y")
  return dayOfWeekAndMonth + suffix + ' ' + monthAndYear
}

/**
 * Returns correct suffix according to day of month ie '1st', '2nd', '3rd' or '4th'
 * 
 * @param {Int} day - the day of the month 1-31
 * @returns {String} the suffix of the day of the month
 */
function determineSuffix(day) {

  if (day >= 11 && day <= 13) {
    return 'th'
  }
  switch(day % 10) {
    case 1:
      return 'st'
    case 2:
      return 'nd'
    case 3:
      return 'rd'
    default:
      return 'th'
   }
}
