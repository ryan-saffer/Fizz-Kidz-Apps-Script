/**
 * A booking already existing in the system
 * 
 * @constructor
 * @param {Sheet} sheet the sheet of the booking - see https://developers.google.com/apps-script/reference/spreadsheet/sheet
 */
function ExistingBooking(sheet) {
    Booking.call(this, sheet) // extends 'Booking'

    //================================================================================
    // Properties
    //================================================================================

    this.eventID = sheet.getRange('B12').getDisplayValue()

    //================================================================================
    // Methods
    //================================================================================

    /**
     * Updates the booking in Drive/Calendar
     * 
     * @param {Event} e the event object given to onEdit - see https://developers.google.com/apps-script/guides/triggers/events
     */
    this.updateBooking = function(e) {

        // unique to this function, so validate separately
        if (this.eventID == "") {
            Browser.msgBox("Booking ID field is empty. Cannot update the booking!");
            throw new Error("Error updating party. EventID was not found");
        }

        try {
            this.validateFields();
        } catch (err) {
            Logger.log(err);
            return;
        }

        // get the start time and end time
        var eventName = this.parentName + " / " + this.childName + " " + this.childAge + "th " + this.mobileNumber;
        var startDate = new Date(this.dateOfParty.getFullYear(), this.dateOfParty.getMonth(), this.dateOfParty.getDate(), this.timeOfParty.getHours() - 1, this.timeOfParty.getMinutes());
        var endDate = this.determineEndDate();

        // determine which calendar we should use
        var calendarID = this.getCalendarID()

        var event = CalendarApp.getCalendarById(calendarID)
                                .getEventById(this.eventID);

        // update the event
        event.setTitle(eventName);
        event.setTime(startDate, endDate);
        event.setLocation(this.location);

        // move this booking sheet into the correct folder (if date has been changed)
        var date = Utilities.formatDate(startDate, 'Australia/Sydney', "MMM d y");
        var time = Utilities.formatDate(startDate, 'Australia/Sydney', "hh:mm a");
        var outputRootFolder = DriveApp.getFolderById("1fxxEQzVjjhO0q1rmU8GzpXeWdNvl_hpy");
        var currentFileID = SpreadsheetApp.getActiveSpreadsheet().getId();
        var currentFile = DriveApp.getFileById(currentFileID);
        var currentFolder = currentFile.getParents().next();
        var currentFolderParent = currentFolder.getParents().next();

        // update fileName
        this.partyType = (this.partyType == "In-store") ? this.location : this.partyType;
        var fileName = this.partyType + ": " + this.parentName + " / " + this.childName + " " + this.childAge + "th" + " : " + time
        currentFile.setName(fileName);

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
        var partyTypeFolder = dateFolder.getFoldersByName(this.partyType);
        if (!partyTypeFolder.hasNext()) { // no folder exists yet for that date and that party type, create one
            partyTypeFolder = dateFolder.createFolder(this.partyType);
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

    /**
     * Deletes the booking from Drive/Calendar
     */
    this.deleteBooking = function() {
        
        // determine which calendar we should use
        var calendarID = this.getCalendarID()
        
        var event = CalendarApp.getCalendarById(calendarID).getEventById(this.eventID);
        
        // delete
        event.deleteEvent();
        
        // delete booking sheet
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
}