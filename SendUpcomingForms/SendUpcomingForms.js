function main() {
  
  var dateToday = new Date();
  var nextWeekendStartDate = new Date(dateToday.getFullYear(), dateToday.getMonth(), dateToday.getDate() + 6);
  var nextWeekendEndDate = new Date(dateToday.getFullYear(), dateToday.getMonth(), dateToday.getDate() + 9);
  
  var malvernStorePartiesCalendarID = "fizzkidz.com.au_j13ot3jarb1p9k70c302249j4g@group.calendar.google.com";
  var balwynStorePartiesCalendarID = "fizzkidz.com.au_7vor3m1efd3fqbr0ola2jvglf8@group.calendar.google.com";
  var travelPartiesCalendarID = "fizzkidz.com.au_b9aruprq8740cdamu63frgm0ck@group.calendar.google.com";
  
  var options = {timeMin : nextWeekendStartDate.toISOString(),
                 timeMax : nextWeekendEndDate.toISOString()};
  
  // malvern store parties
  var response = Calendar.Events.list(malvernStorePartiesCalendarID, options);
  for (var i = 0; i < response.items.length; i++) {
    var attachments = response.items[i].attachments;
    var bookingSheetID = attachments[0].fileId;
    sendPartyForm(bookingSheetID);
  }
  
  // balwyn store parties
  var response = Calendar.Events.list(balwynStorePartiesCalendarID, options);
  for (var i = 0; i < response.items.length; i++) {
    var attachments = response.items[i].attachments;
    var bookingSheetID = attachments[0].fileId;
    sendPartyForm(bookingSheetID);
  }
  
  // travel parties
  response = Calendar.Events.list(travelPartiesCalendarID, options);
  for (var i = 0; i < response.items.length; i++) {
    var attachments = response.items[i].attachments;
    var bookingSheetID = attachments[0].fileId;
    console.log(bookingSheetID);
    sendPartyForm(bookingSheetID);
  }
  
  // while we are here, lets also move the current weekends party sheets into the malvern stores folder
  shareCurrentWeekend();
}

function sendPartyForm(bookingSheetID) {
  
  // Given that the spreadsheet was validated when it was created, validation does not need to take place.
  
  var sheet = SpreadsheetApp.openById(bookingSheetID).getActiveSheet();
  var parentName = sheet.getRange('B1').getDisplayValue();
  var mobileNumber = sheet.getRange('B2').getDisplayValue();
  var emailAddress = sheet.getRange('B3').getDisplayValue();  
  var childName = sheet.getRange('B4').getDisplayValue();  
  var childAge = sheet.getRange('B5').getDisplayValue();  
  var dateOfParty = sheet.getRange('B6').getValue();  
  var timeOfParty = sheet.getRange('B7').getValue(); 
  var partyLength = sheet.getRange('B8').getDisplayValue();  
  var notes = sheet.getRange('B9').getDisplayValue();  
  var partyType = sheet.getRange('B10').getDisplayValue();  
  var location = sheet.getRange('B11').getDisplayValue();
  
  // Determine the start and end times of the party
  var startDate = new Date(dateOfParty.getFullYear(), dateOfParty.getMonth(), dateOfParty.getDate(), timeOfParty.getHours() - 1, timeOfParty.getMinutes());
  var endDate = determineEndDate(dateOfParty, timeOfParty, partyLength);
  
  // create a pre-filled form URL
  var preFilledURL = getPreFilledFormURL(emailAddress, startDate, parentName, mobileNumber, childName, childAge, partyLength, partyType, location, bookingSheetID);
  
  // Using the HTML email template, inject the variables and get the content
  var t = HtmlService.createTemplateFromFile('party_form_email_template');
  t.parentName = parentName;
  t.childName = childName;
  t.childAge = childAge;
  t.startDate = Utilities.formatDate(startDate, 'Australia/Sydney', 'EEEE d MMMM y');
  t.startTime = Utilities.formatDate(startDate, 'Australia/Sydney', 'hh:mm a');
  t.endTime = Utilities.formatDate(endDate, 'Australia/Sydney', 'hh:mm a');
  // determine location
  if (partyType == "In-store") {
    location = (location == "Malvern") ? "our Malvern store" : "our Balwyn store";
  }
  t.location = location;
  t.preFilledURL = preFilledURL;
  
  var body = t.evaluate().getContent();
  var subject = "Information Regarding Your Upcoming Party!";
  var signature = getGmailSignature();
  
  // Send the confirmation email
  GmailApp.sendEmail(emailAddress, subject, "", {htmlBody : body + signature, name : "Fizz Kidz"});
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
  
  var endDate = new Date(dateOfParty.getFullYear(), dateOfParty.getMonth(), dateOfParty.getDate(), timeOfParty.getHours() - 1 + lengthHours, timeOfParty.getMinutes() + lengthMinutes);
  
  return endDate;
}

function getPreFilledFormURL(emailAddress, dateOfParty, parentName, mobileNumber, childName, childAge, partyLength, partyType, location, bookingSheetID) {

  // form IDs
  var inStoreFormID = "1LH52NazS74FuIv1bisZ1kMQuEeC8l5RRT-5f7TzK1n4";
  var travelFormID = "14vQcTDdZSOniRaoOPdy-rnp8kyckWj44WjEbEnB3CE0";
  
  // open the correct form, create a response and get the items
  var formID = (partyType == "In-store") ? inStoreFormID : travelFormID;
  var form = FormApp.openById(formID);
  var formResponse = form.createResponse();
  var formItems = form.getItems();
  
  // first question - date and time
  var dateItem = formItems[1].asDateTimeItem();
  
  // due to strange time formatting behaviour, update the time to (time + 10) or (time + 11) depending on daylight savings
  var correctedPartyTime = 0;
  switch (dateOfParty.getTimezoneOffset()) {
    case -600: // GMT + 10
      correctedPartyTime = dateOfParty.getHours() + 10;
      break;
    case -660: // GMT + 11
      correctedPartyTime = dateOfParty.getHours() + 11;
      break;
    default:
        break;
  }
  dateOfParty = new Date(dateOfParty.getFullYear(), dateOfParty.getMonth(), dateOfParty.getDate(), correctedPartyTime, dateOfParty.getMinutes());
  var response = dateItem.createResponse(dateOfParty);
  formResponse.withItemResponse(response);
  
  // second question - parents name
  var parentNameItem = formItems[2].asTextItem();
  response = parentNameItem.createResponse(parentName);
  formResponse.withItemResponse(response);
  
  // third question - childs name
  var childNameItem = formItems[3].asTextItem();
  response = childNameItem.createResponse(childName);
  formResponse.withItemResponse(response);
  
  // fourth question - childs age
  var childAgeItem = formItems[4].asTextItem();
  response = childAgeItem.createResponse(childAge);
  formResponse.withItemResponse(response);

  // fifth question - location - only if in-store
  if (partyType == "In-store") {
    var locationItem = formItems[5].asMultipleChoiceItem();
    response = locationItem.createResponse(location);
    formResponse.withItemResponse(response);
  }
    
  return formResponse.toPrefilledUrl();
}

function shareCurrentWeekend() {
  
  var dateToday = new Date();
  var friday = new Date(dateToday.getFullYear(), dateToday.getMonth(), dateToday.getDate() + 1);
  var saturday = new Date(dateToday.getFullYear(), dateToday.getMonth(), dateToday.getDate() + 2);
  var sunday = new Date(dateToday.getFullYear(), dateToday.getMonth(), dateToday.getDate() + 3);
  var fridayFormatted = Utilities.formatDate(friday, 'Australia/Sydney', "MMM d y");
  var saturdayFormatted = Utilities.formatDate(saturday, 'Australia/Sydney', "MMM d y");
  var sundayFormatted = Utilities.formatDate(sunday, 'Australia/Sydney', "MMM d y");
  var dates = [fridayFormatted, saturdayFormatted, sundayFormatted];
  
  console.log(dates);
  
  var partySheetsFolder = DriveApp.getFolderById("1EoQxIm6wP8TCZR7EJboZrPygWcy2fb7z");
  var outputFolder = DriveApp.getFolderById("1n2CLk3uJLKx-Rn5HFR7CuEyW17nhO5f-");
  
  // clear the older party sheets first
  var existingFolders = outputFolder.getFolders();
  while (existingFolders.hasNext()) {
    var currentFolder = existingFolders.next();
    outputFolder.removeFolder(currentFolder);
  }
  
  // add the new ones
  for (var i = 0; i < dates.length; i++) {
    if (partySheetsFolder.getFoldersByName(dates[i]).hasNext()) {
      var currentFolder = partySheetsFolder.getFoldersByName(dates[i]).next();
      outputFolder.addFolder(currentFolder);
    }
  }
}

function getGmailSignature() {
  var draft = GmailApp.search("subject:signature label:draft", 0, 1);
  return draft[0].getMessages()[0].getBody();
}
