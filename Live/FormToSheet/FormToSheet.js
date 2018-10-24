var additions_list = [
  "Fairy Bread",
  "Frankfurts",
  "Fruit Platter",
  "Bowl of Freshly Cut Watermelon",
  "Vegetarian Spring Rolls",
  "Wedges",
  "Vegemite Sandwiches",
  "Cheese & Tomato Sandwiches",
  "Combo Sandwich Platter",
  "Lolly Bags"  
]

var additions_prices =
{
  "Fairy Bread":{
     "One Serving":"$25",
     "Two Servings":"$40"
  },
  "Frankfurts":{
     "One Serving":"$20",
     "Two Servings":"$30"
  },
  "Fruit Platter":{
     "One Serving":"$40",
     "Two Servings":"$65"
  },
  "Bowl of Freshly Cut Watermelon":{
     "One Serving":"$20",
     "Two Servings":"$30"
  },
  "Vegetarian Spring Rolls":{
     "One Serving":"$20",
     "Two Servings":"$35"
  },
  "Wedges":{
     "One Serving":"$25",
     "Two Servings":"$40"
  },
  "Vegemite Sandwiches":{
     "One Serving":"$25",
     "Two Servings":"$40"
  },
  "Cheese & Tomato Sandwiches":{
     "One Serving":"$30",
     "Two Servings":"$50"
  },
  "Combo Sandwich Platter":{
    "One Serving":"$30",
    "Two Servings":"$50"
  },
  "Lolly Bags":{
     "One Serving":"$2.50 each",
     "Two Servings":"$2.50 each"
  }
}

function mergeCreations(string1, string2) {
  if (string2 != '') {
    if (string1 == '') {
      return string2;
    } else {
      string1 = string1 + ", " + string2;
      return string1;
    }
  } else {
    return string1;
  }
}

function onSubmit(e) {
  
  console.log(e.values);
  console.log(e.values.length);

  /**
   * Party type determined by number of questions
   * 25 questions in the In-store form, 11 in the Mobile form
   */
  var partyType = (e.values.length == 25) ? e.values[5] : "Mobile";
  console.log("Party type: " + partyType);
  
  var dateTime = e.values[1].split(" ");
  var date = dateTime[0].split("/");
  var day = date[0];
  var month = date[1];
  var year = date[2];
  var time = dateTime[1].split(':');
  var hours = time[0];
  var minutes = time[1];
  date = new Date(year, month - 1, day, hours, minutes);

  var formattedDate = Utilities.formatDate(date, 'Australia/Sydney', 'MMM d y');
  var formattedTime = Utilities.formatDate(date, 'Australia/Sydney', 'hh:mm a');
  
  var parentName = e.values[2];
  console.log("Parent Name: " + parentName);
  
  var childName = e.values[3];
  console.log("Child Name: " + childName);
  
  var childAge = e.values[4];
  console.log("Child age " + childAge);

  // forms diverge from here, so get respective answers
  if (partyType == "Malvern" || partyType == "Balwyn") {
    
    // question 5 is Malvern or Balwyn, so skip to 6
    var childrenCount = e.values[6];
    console.log("Children Count: " + childrenCount);
  
    var creations = mergeCreations(e.values[7], e.values[8]);
    console.log("Creations " + creations);

    var additions = '';
    for (var i = 0; i < additions_list.length; i++) { // each food option is own question
      var selected_serving = e.values[i+9];
      if (selected_serving != '') { // check if option had serving selected
        if (additions != '') { // for the case of the first item
          additions += ',\n';
        }
        var addition = additions_list[i];
        var price = additions_prices[addition][selected_serving];
        additions = additions + addition + " (" + selected_serving + ") - " + price;
      }
    }
    console.log("Additions: " + additions);
  
    var cakeRequired = e.values[19];
    console.log("Cake required: " + cakeRequired);
    
    var selectedCake = e.values[20];
    console.log("Selected Cake: " + selectedCake);
    
    var cakeFlavour = e.values[21];
    console.log("Cake Falvour: " + cakeFlavour);
    
    var extraInfo = e.values[22];
    console.log("Extra Info: " + extraInfo);
    
    var questions = e.values[23];
    console.log("Questions: " + questions);
    
    // if booking a cake, email Talia
    if (cakeRequired == "Yes please!") {
      sendCakeNotification(formattedDate, formattedTime, parentName, childName, childAge, selectedCake, cakeFlavour, partyType);
    }
    
    createPartySheet(formattedDate, formattedTime, parentName, childName, childAge, partyType, childrenCount, creations, additions, cakeRequired, selectedCake, cakeFlavour, extraInfo, questions);
  
  } else {
    var childrenCount = e.values[5];
    // since mobile only takes a number, add 'children
    childrenCount += " children"
    console.log("Children Count: " + childrenCount);
  
    var creations = mergeCreations(e.values[6], e.values[7]);
    console.log("Creations " + creations);

    var extraInfo = e.values[8];
    console.log("Extra Info: " + extraInfo);
  
    var questions = e.values[9];
    console.log("Questions: " + questions);
    
    createPartySheet(formattedDate, formattedTime, parentName, childName, childAge, partyType, childrenCount, creations, "", "", "", "", extraInfo, questions);
  }
}

function createPartySheet(date, time, parentName, childName, childAge, partyType, childrenCount, creations, additions, cakeRequired, selectedCake, cakeFlavour, extraInfo, questions) {
  
  var outputRootFolder = DriveApp.getFolderById("1EoQxIm6wP8TCZR7EJboZrPygWcy2fb7z");
  var template = DriveApp.getFileById("1zxcQlBSlhRYec9ZFanBcNGxqj8sQHdC9x92MBbpPxVE");
  
  // first ensure the filled in form matches a booking in the system, and get that ID
  var bookingSheetID = locateBooking(date, time, parentName, childName, childAge, partyType);
  if (bookingSheetID == null) {
    return;
  }
  // now get the details that weren't pre-filled
  var sheet = SpreadsheetApp.openById(bookingSheetID);
  var contactNumber = sheet.getRange('B2').getDisplayValue();
  var emailAddress = sheet.getRange('B3').getDisplayValue();
  var childAge = sheet.getRange('B5').getDisplayValue();
  var partyLength = sheet.getRange('B8').getDisplayValue();
  var notes = sheet.getRange('B9').getDisplayValue();
  var partyType = sheet.getRange('B10').getDisplayValue();
  var location = sheet.getRange('B11').getDisplayValue();
  
  // since we now have these details (such as email), send a notification email if there are any questions
  if (questions != "") {
    sendQuestionsNotification(date, time, location, parentName, emailAddress, childName, questions);
  }
  
  // search for existing folder of date, otherwise create a new one
  var dateFolderIter = outputRootFolder.getFoldersByName(date);
  var dateFolder;
  if(!dateFolderIter.hasNext()) { // no folder exists yet for that date
    dateFolder = outputRootFolder.createFolder(date);
  } else { // date folder exists
    dateFolder = dateFolderIter.next();
  }
  var outputFolder = getCorrectOutputFolder(dateFolder, partyType, location);
  var newFile = template.makeCopy(outputFolder);
  newFile.setName(time + ": " + parentName + " / " + childName + " " + childAge + "th");
  var newFileID = newFile.getId();
  
  // open the new document
  var document = DocumentApp.openById(newFileID);
  
  // get the table
  var table = document.getBody().getTables()[0];
  
  // set parents name
  var cell = table.getCell(0, 1);
  cell.setText(parentName);
  
  // set contact number
  cell = table.getCell(1,1);
  cell.setText(contactNumber);
  
  // set childs name and age
  cell = table.getCell(2,1);
  cell.setText(childName + " " + childAge);
  
  // set date of party
  cell = table.getCell(3,1);
  cell.setText(date);
  
  // set time of party
  cell = table.getCell(4,1);
  cell.setText(time);
  
  // set party length
  cell = table.getCell(5,1);
  cell.setText(partyLength + " hours");
  
  // set number of children
  cell = table.getCell(6,1);
  cell.setText(childrenCount);
  
  // set creations
  cell = table.getCell(7,1);
  // add each creation onto a new line
  creations = creations.split(',');
  var output = "";
  for (var i = 0; i < creations.length - 1; i++) {
    output  = output + creations[i] + '\n';
  }
  // dont need newline char on end of last line
  output = output + creations[creations.length - 1];
  cell.setText(output);
  
  // set additions
  cell = table.getCell(8,1);
  cell.setText(additions);
  
  // if cake required, display it; otherwise set it to 'no order'
  cell = table.getCell(9,1);
  if (cakeRequired == "Yes please!") {
    cell.setText(cakeFlavour + " " + selectedCake);
  }
  
  // set talias notes
  cell = table.getCell(10,1);
  cell.setText(notes);
  
  // set parents notes from form
  cell = table.getCell(11,1);
  cell.setText(extraInfo + "\n" + questions);
  
  // if a mobile party, add location on as a final row
  if (partyType == "Mobile") {
    var newRow = table.appendTableRow();
    var attributes = {}
    attributes[DocumentApp.Attribute.ITALIC] = true;
    newRow.appendTableCell("Location:").setAttributes(attributes);
    attributes[DocumentApp.Attribute.ITALIC] = false;
    newRow.appendTableCell(location).setAttributes(attributes);
  }
  
  // make the booking sheet obvious that it should not be edited
  // provide URL to the party sheet
  lockDownSheet(sheet, newFile);

  // finally, send a confirmation email
  // this is done inside this function, since we have already retrieved email address from the booking sheet
  sendThankYouEmail(emailAddress, parentName, childrenCount, creations, additions, cakeRequired, selectedCake, cakeFlavour, partyType, location, questions);
}

function sendThankYouEmail(emailAddress, parentName, childrenCount, creations, additions, cakeRequired, selectedCake, cakeFlavour, partyType, location, questions) {
  
  var t = HtmlService.createTemplateFromFile('form_completed_email_template');
  t.parentName = parentName;
  t.childrenCount = childrenCount;
  t.creations = creations;
  t.additions = additions;
  t.cakeRequired = cakeRequired;
  t.selectedCake = selectedCake;
  t.cakeFlavour = cakeFlavour;
  t.partyType = partyType;
  t.questions = questions;
  t.location = location;
  
  var body = t.evaluate().getContent();
  var subject = "Thank you";

  // determine from address
  var fromAddress = determineFromEmailAddress(location);

  var signature = getGmailSignature(fromAddress);
  
  // Send the confirmation email
  GmailApp.sendEmail(emailAddress, subject, "", {from: fromAddress, htmlBody: body + signature, name : "Fizz Kidz"});
}

function determineFromEmailAddress(location) {
  /**
   * If location is Malvern, send from malvern@fizzkidz.com.au
   * If location is Balwyn, send from info@fizzkidz.com.au
   * If location neither (mobile), send from info@fizzkidz.com.au
   */

   if (location == "Malvern") {
     return "malvern@fizzkidz.com.au";
   }
   else if (location == "Balwyn") {
     return "info@fizzkidz.com.au"
   }
   else {
     return "info@fizzkidz.com.au";
   }
}

function locateBooking(date, time, parentName, childName, childAge, partyType) {
  
  // find the booking sheet based on the form answers
  // this will be used to get the notes from the booking sheet
  // if the booking can't be found, email fizzkidz to alert them of error, and next steps
  
  // build the filename as it should be in the bookings folder
  console.log(partyType);
  var fileName = partyType + ": " + parentName + " / " + childName + " " + childAge + "th" + " : " + time;
  var matchingFiles = DriveApp.getFilesByName(fileName);
  if (matchingFiles.hasNext()) {
    return matchingFiles.next().getId();
  } else {
    // no matching booking was found. Notify Talia!
    sendErrorEmail(parentName, childName, childAge, date, time, partyType);
    return null;
  }
} 

function lockDownSheet(sheet, newFile) {
  var range = sheet.getRange('B1:B12');
  var helpText = "The party sheet has already been generated from this booking. Go there to make any changes.";
  for (var i = 1; i <= range.getHeight(); i++) {
    var currentCell = range.getCell(i,1);
    var rule = SpreadsheetApp.newDataValidation().requireTextEqualTo(currentCell.getDisplayValue()).setAllowInvalid(false).setHelpText(helpText).build();
    currentCell.setDataValidation(rule);
  }
  sheet.getRange('A13').setValue("Party sheet already generated. Go there to make any changes:").setFontSize(15).setFontColor('red').setWrap(true);
  sheet.getRange('B13').setValue(newFile.getUrl()).setFontSize(15);
}

function sendCakeNotification(date, time, parentName, childName, childAge, selectedCake, cakeFlavour, partyType) {
  
  // Using the HTML email template, inject the variables and get the content
  var t = HtmlService.createTemplateFromFile('cake_ordered_email_template');
  t.parentName = parentName;
  t.childName = childName;
  t.childAge = childAge;
  t.dateOfParty = date;
  t.timeOfParty = time;
  t.selectedCake = selectedCake;
  t.cakeFlavour = cakeFlavour;
  t.partyType = partyType;
  
  var body = t.evaluate().getContent();
  var subject = "Cake Order!";
  var fromAddress = determineFromEmailAddress(partyType);
  
  // Send the confirmation email
  GmailApp.sendEmail(fromAddress, subject, "", {from: fromAddress, htmlBody: body, name : "Fizz Kidz"});
}

function sendQuestionsNotification(date, time, location, parentName, emailAddress, childName, questions) {
  
  // Using the HTML email template, inject the variables and get the content
  var t = HtmlService.createTemplateFromFile('questions_email_template');
  t.parentName = parentName;
  t.childName = childName;
  t.dateOfParty = date;
  t.timeOfParty = time;
  t.location = location;
  t.questions = questions;
  t.emailAddress = emailAddress;
  
  var body = t.evaluate().getContent();
  var subject = "Questions asked in Party Form!";
  var fromAddress = determineFromEmailAddress(location);
  
  // Send the confirmation email
  GmailApp.sendEmail(fromAddress, subject, "", {from: fromAddress, htmlBody: body, name : "Fizz Kidz"});
}

function getCorrectOutputFolder(dateFolder, partyType, location) {

  if (partyType == "In-store") {
    if (dateFolder.getFoldersByName(location).hasNext()) {
      return dateFolder.getFoldersByName(location).next();
    } else { // folder not yet created, create now
      return dateFolder.createFolder(location);
    }
  } else {
    if (dateFolder.getFoldersByName("Mobile").hasNext()) {
      return dateFolder.getFoldersByName("Mobile").next();
    } else { // folder not yet created, create now
      return dateFolder.createFolder("Mobile")
    }
  }
}

function sendErrorEmail(parentName, childName, childAge, date, time, partyType) {
  
  // Using the HTML email template, inject the variables and get the content
  var t = HtmlService.createTemplateFromFile('error_finding_booking_template');
  t.parentName = parentName;
  t.childName = childName;
  t.childAge = childAge;
  t.dateOfParty = date;
  t.timeOfParty = time;
  t.partyType = partyType;
  
  var body = t.evaluate().getContent();
  var subject = "ERROR: Booking not found!";
  
  // Send the confirmation email
  GmailApp.sendEmail('info@fizzkidz.com.au', subject, "", {htmlBody: body, name : "Fizz Kidz"});
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