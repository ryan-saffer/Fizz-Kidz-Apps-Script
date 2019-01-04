//================================================================================
// Trigger Functions - see https://developers.google.com/apps-script/guides/triggers/
//================================================================================

/**
 * GAS Trigger Function - onOpen
 * Adds menu items to UI
 */
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
    SpreadsheetApp.getActiveSheet().getRange('C1').clear();
  }
}

/**
 * GAS Trigger Function - onEdit
 * @param {object} e the event object - see https://developers.google.com/apps-script/guides/triggers/events
 */
function onEdit(e) {

  var formatter = new Formatter(SpreadsheetApp.getActiveSheet())
  formatter.formatTimeCell()
  
  // only react to edits on existing bookings, not the booking form
  var bookingFormID = "14pGuFT1Ru84XLiHhu9Hd43nl0o5dlgHlrxubU3F6yFo";
  var currentDocumentID = SpreadsheetApp.getActiveSpreadsheet().getId();
  
  if (currentDocumentID != bookingFormID) { // this is an existing booking
    
    var editRange = {
      top : 1,
      bottom : 11,
      left : 2,
      right : 2
    };
    
    // return if the edited cell isn't within the edit range
    var thisRow = e.range.getRow();
    if (thisRow < editRange.top || thisRow > editRange.bottom) return;
    
    var thisCol = e.range.getColumn();
    if (thisCol < editRange.left || thisCol > editRange.right) return;
    
    // in range - update the booking
    var booking = new ExistingBooking(SpreadsheetApp.getActiveSheet())
    booking.updateBooking(e)
      
  } else { // original booking form, validate store locations if in-store
    if (e.value == "Mobile") { // changed to a mobile party
      formatter.clearLocationCell()
    }
    else if (e.value == "In-store") { // changed to an in-store party
      formatter.applyValidationToLocationCell()
    }
  }
}

//================================================================================
// Menu Item Functions
//================================================================================

/**
 * Menu item trigger function.
 * Confirms request to book in party
 */
function showBookingConfirmationDialog() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(
    "Book in Party",
    "This will create a calendar event, attach the booking sheet, as well as send an email confirmation to the parent. \nEnsure all fields are filled in correctly.", 
    ui.ButtonSet.OK_CANCEL);
  
  if (result == ui.Button.OK) {
    var party = new NewBooking(SpreadsheetApp.getActiveSheet())
    party.bookInParty();
  }
}

/**
 * Menu item trigger function.
 * Resets the booking including formatting
 */
function resetSheet() {
  var formatter = new Formatter(SpreadsheetApp.getActiveSheet())
  formatter.resetSheet()
}

/**
 * Menu item trigger function.
 * Re-enables editing of the booking which allows background scripts to run
 */
function showAuthorisationDialog() {
  // restore validation which was removed during lockdown
  var formatter = new Formatter(SpreadsheetApp.getActiveSheet())
  formatter.restoreValidation();
  
  // determine if we have any triggers installed.
  // if not,install the onEdit trigger
  var triggers = ScriptApp.getProjectTriggers();
  if (triggers.length == 0) {
    // add installable trigger for onEdit, which can run oAuth Services
    ScriptApp.newTrigger('onEdit').forSpreadsheet(SpreadsheetApp.getActive().getId()).onEdit().create();
  }
  
  var ui = SpreadsheetApp.getUi();
  ui.alert(
    "Editing enabled",
    "You can now edit the booking!", 
    ui.ButtonSet.OK);
}

/**
 * Menu item trigger function.
 * Confirms booking deletion request
 */
function showDeleteConfirmationDialog() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(
    "Delete Party Booking",
    "This will delete the calendar event, as well as this booking sheet.", 
    ui.ButtonSet.OK_CANCEL);
  
  if (result == ui.Button.OK) {
    var party = new ExistingBooking(SpreadsheetApp.getActiveSheet())
    party.deleteBooking();
  }
}