function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  var bookingFormID = "14pGuFT1Ru84XLiHhu9Hd43nl0o5dlgHlrxubU3F6yFo";
  var currentDocumentID = SpreadsheetApp.getActiveSpreadsheet().getId();
  
  // if in the booking form display options to book a party
  // otherwise, display options to update the booking of the current party
  if (bookingFormID == currentDocumentID) {
    ui.createMenu('FIZZ KIDZ OPTIONS')
    .addItem('Book in Party', 'showBookingConfirmationDialog')
    .addItem('Reset Sheet', 'resetSheet')
    .addToUi();
  } else {
    ui.createMenu('Edit / Delete Booking')
    .addItem('Enable Editing', 'showAuthorisationDialog')
    .addItem('Delete Booking', 'showDeleteConfirmationDialog')
    .addToUi();
    
    // we have loaded, so clear the loading cell
    SpreadsheetApp.getActive().getRange('C1').clear();
  }
}

function onEdit(e) {
  // reset the time format, since this breaks when a non-time value is entered
  var sheet = SpreadsheetApp.getActive();
  var timeCell = sheet.getRange('B7');
  timeCell.setNumberFormat('h:mm am/pm');
  
  // if this is a booking sheet for a booked in party, and the party details are changed, warn the user to update the party!
  var bookingFormID = "14pGuFT1Ru84XLiHhu9Hd43nl0o5dlgHlrxubU3F6yFo";
  var currentDocumentID = SpreadsheetApp.getActiveSpreadsheet().getId();
  
  if (currentDocumentID != bookingFormID) { // this is a booking sheet!
    
    var editRange = {
      top : 1,
      bottom : 11,
      left : 2,
      right : 2
    };
    
    // Exit if we're out of range
    var thisRow = e.range.getRow();
    if (thisRow < editRange.top || thisRow > editRange.bottom) return;
    
    var thisCol = e.range.getColumn();
    if (thisCol < editRange.left || thisCol > editRange.right) return;
    
    // We're in range; update the booking
    updateBooking(e);
      
  } else { // original booking form, validate store locations if in-store
    var partyType = sheet.getRange('B10').getDisplayValue();
    var locationCell = sheet.getRange('B11');
    if (e.value == "Mobile") { // changed to a mobile party
      locationCell.clearContent();
      locationCell.clearDataValidations();
    }
    else if (e.value == "In-store") { // changed to an in-store party
      validateLocationCell(locationCell);
    }
  }
}

function validateLocationCell(locationCell) {
  /*
   * This function applies a dataValidation to the locationCell
   * such that it must be one of the store locations.
   */
  locationCell.clearContent();
  var helpText = "In-store party location must be 'Malvern' or 'Balwyn'";
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(['Malvern', 'Balwyn']).setAllowInvalid(false).setHelpText(helpText).build();
  locationCell.setDataValidation(rule)
  locationCell.setValue("Malvern");
}

function showBookingConfirmationDialog() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(
    "Book in Party",
    "This will create a calendar event, attach the booking sheet, as well as send an email confirmation to the parent. \nEnsure all fields are filled in correctly.", 
    ui.ButtonSet.OK_CANCEL);
  
  if (result == ui.Button.OK) {
    beginWorkflow();
  }
}

function showAuthorisationDialog() {
  // restore validation which was removed during lockdown
  restoreValidation();
  
  // determine if we have any triggers installed.
  // if not,install the onEdit trigger
  var triggers = ScriptApp.getProjectTriggers();
  if (triggers.length == 0) {
    // add installable trigger for onEdit, which can run oAuth Services
    ScriptApp.newTrigger('onEdit').forSpreadsheet(SpreadsheetApp.getActive().getId()).onEdit().create();
  }
  
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(
    "Editing enabled",
    "You can now edit the booking!", 
    ui.ButtonSet.OK);
}

function restoreValidation() {
  var sheet = SpreadsheetApp.getActiveSheet();
  
  // first remove the old validation from the ones that don't need to be validated
  var range = sheet.getRange('B1:B12');
  for(var i = 1; i <= range.getHeight(); i++) {
    var currentCell = range.getCell(i, 1);
    currentCell.setDataValidation(null);
  }
  
  // then add the old validations back  
  currentCell = sheet.getRange('B6');
  helpText = "Party must have a valid date. Double-click on cell to display a date picker.";
  rule = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).setHelpText(helpText).build();
  currentCell.setDataValidation(rule);
  
  currentCell = sheet.getRange('B7');
  helpText = "Party time must be a valid time";
  rule = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).setHelpText(helpText).build();
  currentCell.setDataValidation(rule);
  
  var partyType = sheet.getRange('B10').getDisplayValue();
  currentCell = sheet.getRange('B8');
  if (partyType == "In-store") {
    helpText = "Party length must be either 1.5 or 2 hours";
    rule = SpreadsheetApp.newDataValidation().requireValueInList(['1.5','2']).setAllowInvalid(false).setHelpText(helpText).build();
  } else {
    helpText = "Party length must be either 1 or 1.5 hours";
    rule = SpreadsheetApp.newDataValidation().requireValueInList(['1','1.5']).setAllowInvalid(false).setHelpText(helpText).build();
  }
  currentCell.setDataValidation(rule);
  
  currentCell = sheet.getRange('B10');
  helpText = "Party type cannot be edited. To change store location or convert to a mobile party, you must delete this booking and create a new one.";
  rule = SpreadsheetApp.newDataValidation().requireTextEqualTo(currentCell.getValue()).setAllowInvalid(false).setHelpText(helpText).build();
  currentCell.setDataValidation(rule);

  if (partyType == "In-store") {
    currentCell = sheet.getRange('B11');
    helpText = "An In-store location cannot be changed. To move the booking to a different store, delete this booking and create a new one";
    rule = SpreadsheetApp.newDataValidation().requireTextEqualTo(currentCell.getDisplayValue()).setAllowInvalid(false).setHelpText(helpText).build();
    currentCell.setDataValidation(rule);
  }
  
  currentCell = sheet.getRange('B12');
  helpText = "Booking ID cannot be edited.";
  rule = SpreadsheetApp.newDataValidation().requireTextEqualTo(currentCell.getValue()).setAllowInvalid(false).setHelpText(helpText).build();
  currentCell.setDataValidation(rule);
}

function showDeleteConfirmationDialog() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(
    "Delete Party Booking",
    "This will delete the calendar event, as well as this booking sheet.", 
    ui.ButtonSet.OK_CANCEL);
  
  if (result == ui.Button.OK) {
    deleteBooking();
  }
}

function validateFields(parentName, mobileNumber, emailAddress, childName, childAge, dateOfParty, timeOfParty, partyLength, partyType, location, confirmationEmailRequired) {
  
  if(parentName == "") {
    Browser.msgBox("⚠️You must enter the parents name. Party not booked/updated. Try again.");
    throw new Error("You must enter the parents name. Operation cancelled.");
  }
  
  if(mobileNumber == "") {
    Browser.msgBox("⚠️You must enter the mobile number. Party not booked/updated. Try again.");
    throw new Error("You must enter the mobile number. Operation cancelled.");
  }
  if (mobileNumber.length != 10) {
    Browser.msgBox("⚠️Mobile number is not valid. Party not booked/updated. Try again.");
    throw new Error("Mobile number is not valid. Operation cancelled.");
  }
  
  if (emailAddress == "") {
    Browser.msgBox("⚠️You must enter the email address. Party not booked/updated. Try again.");
    throw new Error("You must enter the email address. Operation cancelled.");
  }
  if (!validateEmail(emailAddress)) {
    Browser.msgBox("⚠️Email address is not valid. Party not booked/updated. Try again.");
    throw new Error("Email address is not valid. Operation cancelled.");
  }
  
  if(childName == "") {
    Browser.msgBox("⚠️You must enter the childs name. Party not booked/updated. Try again.");
    throw new Error("You must enter the childs name. Operation cancelled.");
  }
  
  if(childAge == "") {
    Browser.msgBox("⚠️You must enter the childs age. Party not booked/updated. Try again.");
    throw new Error("You must enter the childs age. Operation cancelled.");
  }
  
  if(dateOfParty == "") {
    Browser.msgBox("⚠️You must enter the party date. Party not booked/updated. Try again.");
    throw new Error("You must enter the party date. Operation cancelled.");
  }
  if (!(dateOfParty instanceof Date)) {
    Browser.msgBox("⚠️Party date is invalid. Party not booked/updated. Try again.");
    throw new Error("Party date is invalid. Operation cancelled");
  }
  
  if(timeOfParty == "") {
    Browser.msgBox("⚠️You must enter the time of the party. Party not booked/updated. Try again.");
    throw new Error("You must enter the time of the party. Operation cancelled.");
  }
  if (!(timeOfParty instanceof Date)) {
    Browser.msgBox("⚠️Party time is invalid. Party not booked/updated. Try again.");
    throw new Error("Party time is invalid. Operation cancelled.");
  }
  if (timeOfParty.getFullYear() == 1900) {
    Browser.msgBox("⚠️Party time is invalid. Party not booked/updated. Try again.");
    throw new Error("Party time is invalid. Operation cancelled");
  }
 
  if(partyLength == "") {
    Browser.msgBox("⚠️You must enter the length of the party. Party not booked/updated. Try again.");
    throw new Error("You must enter the length of the party. Operation cancelled.");
  }
  
  if(partyType == "") {
    Browser.msgBox("⚠️You must enter the type of party as In-store or Mobile. Party not booked/updated. Try again.");
    throw new Error("You must enter the type of party as In-store or Mobile. Operation cancelled.");
  }
  // In-store must be 1.5 or 2 hours, Mobile must be 1 or 1.5 hours. Validate here
  if (partyType == "In-store") {
    if (partyLength == "1") {
      Browser.msgBox("⚠️An In-store party cannot have a party length of 1 hour. Party not booked/updated. Try again.");
      throw new Error("An In-store party cannot have a party length of 1 hour. Operation cancelled.");
    }
  }
  if (partyType == "Mobile") {
    if (partyLength == "2") {
      Browser.msgBox("⚠️A Mobile party cannot have a party length of 2 hours Party not booked/updated. Try again.");
      throw new Error("A Mobile party cannot have a party length of 2 hours. Operation cancelled.");
    }
  }
  
  // mobile party must have a location
  if (partyType == "Mobile") {
    if (location == "") {
      Browser.msgBox("⚠️Mobile party must have a location. Party not booked/updated. Try again.");
      throw new Error("Mobile party must have a location. Operation cancelled.");
    }
  }
  // in-store party cannot have a location
  if (partyType == "In-store") {
    if (location != "Malvern" && location != "Balwyn") {
      Browser.msgBox("⚠️An In-store party location must be 'Malvern' or 'Balwyn'. Party not booked/updated. Try again.");
      throw new Error("In-store party location must be Malern or Balwyn. Operation cancelled.");
    }
  }
  
  if(confirmationEmailRequired == "") {
    Browser.msgBox("⚠️You must enter if a confirmation email is required. Party not booked/updated. Try again.");
    throw new Error("You must enter if a confirmation email is required. Operation cancelled.");
  }
}

function validateEmail(email) {
  
  // Uses a regex to ensure the entered email address is valid
  var re = /^(([^<>()[\]\\.,;:\s@\"]+(\.[^<>()[\]\\.,;:\s@\"]+)*)|(\".+\"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
  return re.test(email);
}

function beginWorkflow() {
  var sheet = SpreadsheetApp.getActive();
  
  var parentName = sheet.getRange('B1').getDisplayValue();
  var mobileNumber = sheet.getRange('B2').getDisplayValue();
  var emailAddress = sheet.getRange('B3').getDisplayValue();
  var childName = sheet.getRange('B4').getDisplayValue();
  var childAge = sheet.getRange('B5').getDisplayValue();
  var dateOfParty = sheet.getRange('B6').getValue();
  var timeOfParty = sheet.getRange('B7').getValue();
  var partyLength = sheet.getRange('B8').getDisplayValue()
  var notes = sheet.getRange('B9').getDisplayValue();
  var partyType = sheet.getRange('B10').getDisplayValue();
  var location = sheet.getRange('B11').getDisplayValue();
  var confirmationEmailRequired = sheet.getRange('B12').getDisplayValue();
  
  // validate the data
  try {
    validateFields(parentName, mobileNumber, emailAddress, childName, childAge, dateOfParty, timeOfParty, partyLength, partyType, location, confirmationEmailRequired);
  } catch (err) {
    Logger.log(err);
    return;
  }
  
  // store party details in a new file
  var fileID = createCopyOfSheet(parentName, childName, childAge, dateOfParty, timeOfParty, partyType, location);
  
  // create the event
  createEvent(parentName, mobileNumber, childName, childAge, dateOfParty, timeOfParty, partyLength, partyType, location, fileID);
  
  // send a confirmation email to the parent, if selected
  if (confirmationEmailRequired == "YES") {
    sendConfirmationEmail(parentName, emailAddress, childName, childAge, dateOfParty, timeOfParty, partyLength, partyType, location);
  }
}

function createCopyOfSheet(parentName, childName, childAge, dateOfParty, timeOfParty, partyType, location) {
  // 
  // This function will create a copy of the booking under Party Bookings -> Date of Party -> "Parent / Child : Time"
  // It returns the ID of this document, in order to attach it to the calendar event
  //
  
  // Get the correct date
  var startDate = new Date(dateOfParty.getFullYear(), dateOfParty.getMonth(), dateOfParty.getDate(), timeOfParty.getHours() - 1, timeOfParty.getMinutes());
  var formattedTime = Utilities.formatDate(startDate, 'Australia/Sydney', 'hh:mm a');
  var formattedDate = Utilities.formatDate(startDate, 'Australia/Sydney', 'MMM d y');
  
  var outputRootFolder = DriveApp.getFolderById("1fxxEQzVjjhO0q1rmU8GzpXeWdNvl_hpy");
  var template = DriveApp.getFileById("14pGuFT1Ru84XLiHhu9Hd43nl0o5dlgHlrxubU3F6yFo");
  
  // create the filename. In-store use store name, Mobile use mobile
  partyType = (partyType == "In-store") ? location : partyType;
  var fileName = partyType + ": " + parentName + " / " + childName + " " + childAge + "th" + " : " + formattedTime;

  // search for existing folder of date, otherwise create a new one
  var dateFolder = outputRootFolder.getFoldersByName(formattedDate);
  var newFile = null;
  if(!dateFolder.hasNext()) { // no folder exists yet for that date, create one
    dateFolder = outputRootFolder.createFolder(formattedDate);
  } else {
    dateFolder = dateFolder.next();
  }
  // search for the party type within the new folder
  var partyTypeFolder = dateFolder.getFoldersByName(partyType);
  if (!partyTypeFolder.hasNext()) { // no folder exists yet for that date and that type, create one
    partyTypeFolder = dateFolder.createFolder(partyType);
    newFile = template.makeCopy(fileName, partyTypeFolder);
  } else {
    newFile = template.makeCopy(fileName, partyTypeFolder.next());
  }

  var newFileID = newFile.getId();
  
  // make required changes to this new file, such as removing confirmation email row, and validating store type only with chosen type
  var sheet = SpreadsheetApp.openById(newFileID).getActiveSheet();
  sheet.deleteRow(12);
  
  // set a cell to indicate loading - it will be removed in the onOpen trigger
  sheet.getRange('C1').setValue("LOADING FIZZ OPTIONS...").setFontSize(15).setFontColor('red');
  
  // lock down the cells, until they enable editing
  lockDownCells(sheet);
  
  return newFileID;
}

function createEvent(parentName, mobileNumber, childName, childAge, dateOfParty, timeOfParty, partyLength, partyType, location, fileID) {
  
  var eventName = parentName + " / " + childName + " " + childAge + "th " + mobileNumber;
  var startDate = new Date(dateOfParty.getFullYear(), dateOfParty.getMonth(), dateOfParty.getDate(), timeOfParty.getHours() - 1, timeOfParty.getMinutes());
  
  var endDate = determineEndDate(dateOfParty, timeOfParty, partyLength);
  
  var eventObj = { 
    summary: eventName,
    start: {dateTime: startDate.toISOString()},
    end: {dateTime: endDate.toISOString()},
    location: location,
    attachments: [{
      'fileUrl': 'https://drive.google.com/open?id=' + fileID,
      'title' : 'Booking Sheet'
    }]
  };
  
  // determine which calendar to use
  var calendarID = getCalendarID(partyType, location);

  var newEvent = Calendar.Events.insert(eventObj, calendarID, {'supportsAttachments': true});
  
  // now add this event ID to our booking sheet, in order to update/delete in the future
  var cell = SpreadsheetApp.openById(fileID).getActiveSheet().getRange('B12');
  cell.setValue(newEvent.id);
  
  // now lock down cell since this was left out earlier
  var helpText = "Booking cannot be edited until you select 'Edit / Delete Booking' -> 'Enable Editing', and follow the prompts";
  var rule = SpreadsheetApp.newDataValidation().requireTextEqualTo(cell.getDisplayValue()).setAllowInvalid(false).setHelpText(helpText).build();
  cell.setDataValidation(rule);
}

function sendConfirmationEmail(parentName, emailAddress, childName, childAge, dateOfParty, timeOfParty, partyLength, partyType, location) {
  
  // Determine the start and end times of the party
  var startDate = new Date(dateOfParty.getFullYear(), dateOfParty.getMonth(), dateOfParty.getDate(), timeOfParty.getHours() - 1, timeOfParty.getMinutes());
  var endDate = determineEndDate(dateOfParty, timeOfParty, partyLength);
  
  // Determine if making one or two creations
  var creationCount;
  if (partyType == "In-store") {
    switch (partyLength) {
      case "1.5":
        creationCount = "two";
        break;
      case "2":
        creationCount = "three";
        break;
      default:
        break;
    }
  } else if (partyType == "Mobile") {
    switch (partyLength) {
      case "1":
        creationCount = "two";
        break;
      case "1.5":
        creationCount = "three";
        break;
      default:
        break;
    }
  }
  
  // Using the HTML email template, inject the variables and get the content
  var t = HtmlService.createTemplateFromFile('booking_confirmation_email_template');
  t.parentName = parentName;
  t.childName = childName;
  t.childAge = childAge;
  t.startDate = Utilities.formatDate(startDate, 'Australia/Sydney', 'EEEE d MMMM y');
  t.startTime = Utilities.formatDate(startDate, 'Australia/Sydney', 'hh:mm a');
  t.endTime = Utilities.formatDate(endDate, 'Australia/Sydney', 'hh:mm a');
  var updated_location = location;
  if (partyType == "In-store") {
    updated_location = (location == "Malvern") ? "our Malvern store" : "our Balwyn store"
  }
  t.location = updated_location;
  t.creationCount = creationCount;
  
  // attach correct invitations
  var inStoreInvitations3pp = DriveApp.getFilesByName("Fizz Kidz Invitations - 3pp.pdf").next();
  var inStoreInvitationsLarge = DriveApp.getFilesByName("Fizz Kidz Invitations - Large.pdf").next();
  var balwynInvitations = DriveApp.getFilesByName("Fizz Kidz Invitations - Balwyn.pdf").next();
  var mobilePartyInvitations = DriveApp.getFilesByName("Fizz Kidz Mobile Party Invitations.jpg").next();
  
  var attachments = [];
  if (partyType == "In-store") {
    if (location == "Malvern") {
      attachments.push(inStoreInvitations3pp);
      attachments.push(inStoreInvitationsLarge);
    }
    else if (location == "Balwyn") {
      attachments.push(balwynInvitations);
    }
  } else {
    attachments.push(mobilePartyInvitations);
  }
  
  var body = t.evaluate().getContent();
  var subject = "Party Booking Confirmation";
  var signature = getGmailSignature();
  
  // Send the confirmation email
  GmailApp.sendEmail(emailAddress, subject, "", {htmlBody: body + signature, name : "Fizz Kidz", attachments : attachments});
}

function updateBooking(e) {
  var sheet = SpreadsheetApp.getActive();
  
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
  
  // unique to this function, so validate separately
  var eventID = sheet.getRange('B12').getDisplayValue();
  if (eventID == "") {
    Browser.msgBox("Booking ID field is empty. Cannot update the booking!");
    throw new Error("Error updating party. EventID was not found");
  }
  
  try {
    validateFields(parentName, mobileNumber, emailAddress, childName, childAge, dateOfParty, timeOfParty, partyLength, partyType, location, null);
  } catch (err) {
    Logger.log(err);
    return;
  }
  
  // get the start time and end time
  var eventName = parentName + " / " + childName + " " + childAge + "th " + mobileNumber;
  var startDate = new Date(dateOfParty.getFullYear(), dateOfParty.getMonth(), dateOfParty.getDate(), timeOfParty.getHours() - 1, timeOfParty.getMinutes());
  
  var endDate = determineEndDate(dateOfParty, timeOfParty, partyLength);
  
  // determine which calendar we should use
  var calendarID = getCalendarID(partyType, location);
  
  var event = CalendarApp.getCalendarById(calendarID).getEventById(eventID);
  
  // update the event
  event.setTitle(parentName + " / " + childName + " " + childAge + "th " + mobileNumber);
  event.setTime(startDate, endDate);
  event.setLocation(location);
  
  // move this booking sheet into the correct folder (if date has been changed)
  var date = Utilities.formatDate(startDate, 'Australia/Sydney', "MMM d y");
  var time = Utilities.formatDate(startDate, 'Australia/Sydney', "hh:mm a");
  var outputRootFolder = DriveApp.getFolderById("1fxxEQzVjjhO0q1rmU8GzpXeWdNvl_hpy");
  var currentFileID = SpreadsheetApp.getActiveSpreadsheet().getId();
  var currentFile = DriveApp.getFileById(currentFileID);
  var currentFolder = currentFile.getParents().next();
  var currentFolderParent = currentFolder.getParents().next();
  
  // update fileName
  partyType = (partyType == "In-store") ? location : partyType;
  currentFile.setName(partyType + ": " + parentName + " / " + childName + " " + childAge + "th" + " : " + time);
  
  // determine if the date was changed. If so, re-organise file in Drive. If not, we are done!
  var editedRow = e.range.getRow();
  var editedColumn = e.range.getColumn();
  if (editedRow != 6 || editedColumn != 2) { // not editing date cell, update is finished
    return;
  } else { // editing date cell
    if (parseInt(e.oldValue) == parseInt(e.value)) {
      // date has not changed, so update is finished
      return;
    }
  }
  
  // insert into new location
  var dateFolder = outputRootFolder.getFoldersByName(date);
  if(!dateFolder.hasNext()) { // no folder exists yet for that date, create one
    dateFolder = outputRootFolder.createFolder(date);
  } else {
    dateFolder = dateFolder.next();
  }
  var partyTypeFolder = dateFolder.getFoldersByName(partyType);
  if (!partyTypeFolder.hasNext()) { // no folder exists yet for that date and that party type, create one
    partyTypeFolder = dateFolder.createFolder(partyType);
    partyTypeFolder.addFile(currentFile);
  } else {
    partyTypeFolder.next().addFile(currentFile);
  }
  
  // finally, remove the file
  currentFolder.removeFile(currentFile);
  // if removing this file made that folder empty, delete the folder
  // if party type folder has no bookings, delete party type folder
  if (!currentFolder.getFiles().hasNext()) {
    Drive.Files.remove(currentFolder.getId());
  }
  // if date folder has no party type folders, delete the folder
  if (!currentFolderParent.getFolders().hasNext()) {
    Drive.Files.remove(currentFolderParent.getId());
  }
}

function deleteBooking() {
  var sheet = SpreadsheetApp.getActive();
  
  // since deleting, all we need is the event id and which calendar
  var eventID = sheet.getRange('B12').getDisplayValue();
  
  // determine which calendar we should use
  var partyType = sheet.getRange('B10').getDisplayValue();
  var location = sheet.getRange('B11').getDisplayValue();
  var calendarID = getCalendarID(partyType, location);
  
  var event = CalendarApp.getCalendarById(calendarID).getEventById(eventID);
  
  // delete
  event.deleteEvent();
  
  // delete booking sheet
  var outputRootFolder = DriveApp.getFolderById("1fxxEQzVjjhO0q1rmU8GzpXeWdNvl_hpy");
  var currentFileID = SpreadsheetApp.getActiveSpreadsheet().getId();
  var currentFile = DriveApp.getFileById(currentFileID);
  var currentFolder = currentFile.getParents().next();
  Drive.Files.remove(currentFileID); // use advanced Drive service to permanently delete, not just place in bin
  
  // if deleting the booking sheet leaves this folder empty, delete the folder
  // get folders parent folder (date folder)
  var dateFolder = currentFolder.getParents().next();
  if (!currentFolder.getFiles().hasNext()) {
    Drive.Files.remove(currentFolder.getId());
  }
  if (!dateFolder.getFolders().hasNext()) {
    Drive.Files.remove(dateFolder.getId());
  }
}

function getCalendarID(partyType, location) {

  // event IDs
  var malvernStorePartiesCalendarID = "fizzkidz.com.au_j13ot3jarb1p9k70c302249j4g@group.calendar.google.com";
  var balwynStorePartiesCalendarID = "fizzkidz.com.au_7vor3m1efd3fqbr0ola2jvglf8@group.calendar.google.com";
  var mobilePartiesCalendarID = "fizzkidz.com.au_b9aruprq8740cdamu63frgm0ck@group.calendar.google.com";

  var calendarID;
  if (partyType == "In-store") {
    if (location == "Malvern") {
      calendarID = malvernStorePartiesCalendarID;
    } else if (location == "Balwyn") {
      calendarID = balwynStorePartiesCalendarID;
    }
  } else {
    calendarID = mobilePartiesCalendarID;
  }

  return calendarID;
}

function lockDownCells(sheet) {
  var range = sheet.getRange('B1:B11'); // stop before booking ID, since this will be added when event is created
  var helpText = "Booking cannot be edited until you select 'Edit / Delete Booking' -> 'Enable Editing', and follow the prompts";
  
  for(i = 1; i <= range.getHeight(); i++) {
    var currentCell = range.getCell(i, 1);
    var rule = SpreadsheetApp.newDataValidation().requireTextEqualTo(currentCell.getDisplayValue()).setAllowInvalid(false).setHelpText(helpText).build();
    currentCell.setDataValidation(rule);
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

function getGmailSignature() {
  var draft = GmailApp.search("subject:signature label:draft", 0, 1);
  return draft[0].getMessages()[0].getBody();
}

function resetSheet() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getRange('B1:B9');
  range.clear({ contentsOnly : true });
  sheet.getRange('B10').setValue('In-store');
  validateLocationCell(sheet.getRange('B11'));
  sheet.getRange('B12').setValue('YES');
}