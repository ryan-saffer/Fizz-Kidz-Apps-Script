/**
 * triggered weekly Thursday 9:34am
 * email parents with parties on following Friday+
 * update 'This Weekends Parties' folder in Drive
 */

function main() {
  
  var dateToday = new Date();
  var nextWeekendStartDate  = new Date(
                                    dateToday.getFullYear(),
                                    dateToday.getMonth(),
                                    dateToday.getDate() + 8
                                  )
  var nextWeekendEndDate    = new Date(
                                    dateToday.getFullYear(),
                                    dateToday.getMonth(),
                                    dateToday.getDate() + 11
                                  )
  
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
  
  // generate a report of additional food and send off to managers
  generateReport();
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
  t.startDate = buildFormattedStartDate(startDate)
  t.startTime = Utilities.formatDate(startDate, 'Australia/Sydney', 'hh:mm a');
  t.endTime = Utilities.formatDate(endDate, 'Australia/Sydney', 'hh:mm a');
  
  // determine location
  var updated_location = location;
  if (partyType == "In-store") {
    updated_location = (location == "Malvern") ? "our Malvern store" : "our Balwyn store";
  }
  t.location = updated_location;
  t.preFilledURL = preFilledURL;
  
  var body = t.evaluate().getContent();
  var subject = "Information Regarding Your Upcoming Party!";

  // determine the from email address
  var fromAddress = determineFromEmailAddress(location);

  var signature = getGmailSignature(fromAddress);
  
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

function generateReport() {

  // 1. open form responses spreadsheet
  // 2. find range of rows within upcoming weekend
  // 3. search each row for additions

  var RESPONSES_FILE_ID = '1C2QOmdKoODDO0MOopJeSehTdeUvZxL4F9kDafCCfNgM'

  var responsesFile = SpreadsheetApp.openById(RESPONSES_FILE_ID),
      responsesSheet = responsesFile.getSheetByName("In-Store Responses"),
      malvernOrders = [0,0,0,0,0,0,0,0,0,0],
      balwynOrders = [0,0,0,0,0,0,0,0,0,0],
      dateToday = new Date(),
      dateFriday = new Date(dateToday.getFullYear(), dateToday.getMonth(), dateToday.getDate() + 1),
      dateSunday = new Date(dateToday.getFullYear(), dateToday.getMonth(), dateToday.getDate() + 3),
      dateMonday = new Date(dateToday.getFullYear(), dateToday.getMonth(), dateToday.getDate() + 4),
      startRow = 2,
      endRow = responsesSheet.getLastRow()
  
  for (var row = startRow; row <= endRow; row++) {
    var dateOfParty = responsesSheet.getRange('B'+row).getValue()

    if (dateFriday < dateOfParty && dateOfParty < dateMonday) {
      var location = responsesSheet.getRange('F'+row).getDisplayValue()
      var additionsRange = responsesSheet.getRange('J'+row+':R'+row)
      for (var col = 1; col <= 9; col++) {
        var val = additionsRange.getCell(1, col).getDisplayValue()
        if (location == "Malvern") {
          if (val == "One Serving") {
            malvernOrders[col-1] += 1
          }
          else if (val == "Two Servings") {
            malvernOrders[col-1] += 2
          }
        } else if (location == "Balwyn") {
          if (val == "One Serving") {
            balwynOrders[col-1] += 1
          }
          else if (val == "Two Servings") {
            balwynOrders[col-1] += 2
          }
        }
      }
      // also keep running total of lolly bags
      var lollyBagsVal = responsesSheet.getRange('S'+row).getDisplayValue()
      if (lollyBagsVal == "Yes") {
        var childrenCount = responsesSheet.getRange('G'+row).getDisplayValue()
        childrenCount = childrenCount.substring(5, 7)
        console.log("COUNT: " + childrenCount)
        if (location == "Malvern") {
          malvernOrders[9] += parseInt(childrenCount)
        }
        else if (location == "Balwyn") {
          balwynOrders[9] += parseInt(childrenCount)
        }
      }
    }
  }

  var reportTemplate = DriveApp.getFileById('1Syupwyg_tXjmTw4yN7Sysig_K3GJRwkM1IQwrBhnbZw')
  var report = reportTemplate.makeCopy()
  var newDoc = DocumentApp.openById(report.getId())
  var body = newDoc.getBody()

  var fridayFormatted = Utilities.formatDate(dateFriday, 'Australia/Sydney', 'd/MM/yy')
  var sundayFormatted = Utilities.formatDate(dateSunday, 'Australia/Sydney', 'd/MM/yy')

  body.replaceText('%DATERANGE%', fridayFormatted + ' - ' + sundayFormatted)
  body.replaceText('%MALV_FAIRYBREAD%',malvernOrders[0])
  body.replaceText('%MALV_FRANKFURTS%', malvernOrders[1])
  body.replaceText('%MALV_FRUITPLATTER%', malvernOrders[2])
  body.replaceText('%MALV_WATERMELON%', malvernOrders[3])
  body.replaceText('%MALV_SPRINGROLLS%', malvernOrders[4])
  body.replaceText('%MALV_WEDGES%', malvernOrders[5])
  body.replaceText('%MALV_VEGSAND%', malvernOrders[6])
  body.replaceText('%MALV_CHEESETOMSAND%', malvernOrders[7])
  body.replaceText('%MALV_COMBOSAND%', malvernOrders[8])
  body.replaceText('%MALV_LOLLYBAGS%', malvernOrders[9])
  body.replaceText('%BALWYN_FAIRYBREAD%',balwynOrders[0])
  body.replaceText('%BALWYN_FRANKFURTS%', balwynOrders[1])
  body.replaceText('%BALWYN_FRUITPLATTER%', balwynOrders[2])
  body.replaceText('%BALWYN_WATERMELON%', balwynOrders[3])
  body.replaceText('%BALWYN_SPRINGROLLS%', balwynOrders[4])
  body.replaceText('%BALWYN_WEDGES%', balwynOrders[5])
  body.replaceText('%BALWYN_VEGSAND%', balwynOrders[6])
  body.replaceText('%BALWYN_CHEESETOMSAND%', balwynOrders[7])
  body.replaceText('%BALWYN_COMBOSAND%', balwynOrders[8])
  body.replaceText('%BALWYN_LOLLYBAGS%', balwynOrders[9])

  newDoc.saveAndClose()

  var docblob = newDoc.getAs('application/pdf')
  var newFile = DriveApp.createFile(docblob)

  // email the reports
  var subject = "Additional Food Weekly Report"
  var attachments = [newFile]
  GmailApp.sendEmail('info@fizzkidz.com.au, malvern@fizzkidz.com.au', subject, "", {attachments: attachments})

  // delete the files
  DriveApp.removeFile(newFile)
  Drive.Files.remove(report.getId())

}

function getGmailSignature(fromAddress) {
  var draft;
  if (fromAddress == "info@fizzkidz.com.au") {
    draft = GmailApp.search("subject:talia-signature label:draft", 0, 1);
  }
  else if (fromAddress = "malvern@fizzkidz.com.au") {
    draft = GmailApp.search("subject:romy-signature label:draft", 0, 1);
  }
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
