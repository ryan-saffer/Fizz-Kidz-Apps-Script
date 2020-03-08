/**
 * A new booking created using 'Party Booking Form' - https://docs.google.com/spreadsheets/d/14pGuFT1Ru84XLiHhu9Hd43nl0o5dlgHlrxubU3F6yFo/edit?usp=sharing
 * 
 * @constructor
 * @param {Sheet} sheet the sheet of the booking - see https://developers.google.com/apps-script/reference/spreadsheet/sheet
 */
function NewBooking(sheet) {
    Booking.call(this, sheet) // extends 'Booking'

    //================================================================================
    // Properties
    //================================================================================

    this.confirmationEmailRequired = sheet.getRange('B12').getDisplayValue()

    //================================================================================
    // Methods
    //================================================================================

    /**
     * 1. Creates a party booking file
     * 2. Creates a calendar event
     * 3. Send confirmation email
     */
    this.bookInParty = function() {
        
        // validate the data
        try {
          this.validateFields();
        } catch (err) {
          Logger.log(err);
          return;
        }
        
        // 1.
        var fileID = this.createCopyOfSheet();
        
        // 2.
        this.createEvent(fileID);
        
        // 3.
        this.sendConfirmationEmail()
    }

    /**
     * Create a copy of the boooking in Drive at location:
     * 'Party Booking System -> Party Bookings -> Date of Paty -> Location -> "Location: Parent / Child : Time"'
     * 
     * @return {string} ID of newly created spreadsheet
     */
    this.createCopyOfSheet = function() {
    
        // Get the correct date
        var startDate = new Date(this.dateOfParty.getFullYear(), this.dateOfParty.getMonth(), this.dateOfParty.getDate(), this.timeOfParty.getHours(), this.timeOfParty.getMinutes());
        var formattedTime = Utilities.formatDate(startDate, 'Australia/Sydney', 'hh:mm a');
        var formattedDate = Utilities.formatDate(startDate, 'Australia/Sydney', 'MMM d y');
        
        var outputRootFolder = DriveApp.getFolderById("1fxxEQzVjjhO0q1rmU8GzpXeWdNvl_hpy");
        var template = DriveApp.getFileById("14pGuFT1Ru84XLiHhu9Hd43nl0o5dlgHlrxubU3F6yFo");
        
        // create the filename. In-store use store name, Mobile use mobile
        var partyType = (this.partyType == "In-store") ? this.location : this.partyType;
        var fileName = partyType + ": " + this.parentName + " / " + this.childName + " " + this.childAge + "th" + " : " + formattedTime;
    
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
        
        // lock down the cells, until editing is enabled
        var formatter = new Formatter(sheet)
        formatter.lockDownCells();
        
        return newFileID;
    }

    /**
     * Creates a calendar event with booking sheet attached
     * 
     * @param {string} fileID the ID of the booking spreadsheet
     */
    this.createEvent = function(fileID) {

        var eventName = this.parentName
            + " / " + this.childName
            + " " + this.childAge + "th "
            + this.mobileNumber;
        var startDate = new Date(
            this.dateOfParty.getFullYear(),
            this.dateOfParty.getMonth(),
            this.dateOfParty.getDate(),
            this.timeOfParty.getHours(),
            this.timeOfParty.getMinutes()
        );
        
        var endDate = this.determineEndDate();
        
        var eventObj = { 
            summary: eventName,
            start: {dateTime: startDate.toISOString()},
            end: {dateTime: endDate.toISOString()},
            location: this.location,
            attachments: [{
                'fileUrl': 'https://drive.google.com/open?id=' + fileID,
                'title': 'Booking Sheet'
            }]
        };
        
        // determine which calendar to use
        var calendarID = this.getCalendarID()
    
        var newEvent = Calendar.Events.insert(eventObj, calendarID, {'supportsAttachments': true});
        
        // now add this event ID to our booking sheet, in order to update/delete in the future
        var cell = SpreadsheetApp.openById(fileID).getActiveSheet().getRange('B12');
        cell.setValue(newEvent.id);
        
        // now lock down cell since this was left out earlier
        var helpText = "Booking cannot be edited until you select 'Edit / Delete Booking' -> 'Enable Editing', and follow the prompts";
        var rule = SpreadsheetApp.newDataValidation().requireTextEqualTo(cell.getDisplayValue()).setAllowInvalid(false).setHelpText(helpText).build();
        cell.setDataValidation(rule);
    }

    /**
     * Send booking confirmation email to parent, with invitations attached
     * Email is sent as html document, with personalised details injected as variables
     */
    this.sendConfirmationEmail = function() {
    
        if (this.confirmationEmailRequired == 'NO'){
            // no need to proceed further
            return
        }

        // Determine the start and end times of the party
        var startDate = new Date(
            this.dateOfParty.getFullYear(),
            this.dateOfParty.getMonth(),
            this.dateOfParty.getDate(),
            this.timeOfParty.getHours(),
            this.timeOfParty.getMinutes()
        );
        var endDate = this.determineEndDate();
        
        // Determine if making one or two creations
        var creationCount;
        if (this.partyType == "In-store") {
            switch (this.partyLength) {
                case "1.5":
                    creationCount = "two";
                    break;
                case "2":
                    creationCount = "three";
                    break;
                default:
                    break;
            }
        } else if (this.partyType == "Mobile") {
            switch (this.partyLength) {
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
        t.parentName = this.parentName;
        t.childName = this.childName;
        t.childAge = this.childAge;
        t.startDate = buildFormattedStartDate(startDate)
        t.startTime = Utilities.formatDate(startDate, 'Australia/Sydney', 'hh:mm a');
        t.endTime = Utilities.formatDate(endDate, 'Australia/Sydney', 'hh:mm a');
        var updated_location = this.location;
        if (this.partyType == "In-store") {
            updated_location = `our ${this.location} store`
        }
        t.partyType = this.partyType;
        t.location = updated_location;
        t.creationCount = creationCount;
        
        var body = t.evaluate().getContent();
        var subject = "Party Booking Confirmation";
    
        // determine which account to send from
        var fromAddress = determineFromEmailAddress(this.location);
    
        var signature = getGmailSignature();
        
        // Send the confirmation email
        GmailApp.sendEmail(this.emailAddress, subject, "", {from: fromAddress, htmlBody: body + signature, name : "Fizz Kidz"});
    }
}