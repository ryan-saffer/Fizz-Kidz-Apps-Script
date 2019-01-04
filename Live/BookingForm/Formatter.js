/**
 * A Formatter class used to keep the booking form values in order
 * 
 * @constructor
 * @param {Sheet} sheet the sheet of the form to format - see https://developers.google.com/apps-script/reference/spreadsheet/sheet
 */
var Formatter = function(sheet) {
    
    //================================================================================
    // Properties
    //================================================================================

    this.sheet = sheet

    //================================================================================
    // Methods
    //================================================================================

    /**
     * Reset the time format cell
     */
    this.formatTimeCell = function() {
        var timeCell = this.sheet.getRange('B7');
        timeCell.setNumberFormat('h:mm am/pm');
    }

    /**
     * Clears location cell
     */
    this.clearLocationCell = function() {
        var locationCell = this.sheet.getRange('B11')
        locationCell.clearContent()
        locationCell.clearDataValidations()
    }

    /*
    * Applies a dataValidation to the locationCell such that it must be one of the store locations.
    */
    this.applyValidationToLocationCell = function() {
        var locationCell = this.sheet.getRange('B11')
        locationCell.clearContent();
        var helpText = "In-store party location must be 'Malvern' or 'Balwyn'";
        var rule = SpreadsheetApp.newDataValidation()
                                    .requireValueInList(['Malvern', 'Balwyn'])
                                    .setAllowInvalid(false)
                                    .setHelpText(helpText)
                                    .build();
        locationCell.setDataValidation(rule)
        locationCell.setValue("Malvern");
    }

    /**
     * Re-enables editing and applies validation to cells
     */
    this.restoreValidation = function() {
        // first remove the old validation from the ones that don't need to be validated
        var range = this.sheet.getRange('B1:B12');
        for(var i = 1; i <= range.getHeight(); i++) {
            var currentCell = range.getCell(i, 1);
            currentCell.setDataValidation(null);
        }
        
        // then add the old validations back  
        currentCell = this.sheet.getRange('B6');
        helpText = "Party must have a valid date. Double-click on cell to display a date picker.";
        rule = SpreadsheetApp.newDataValidation()
                                .requireDate()
                                .setAllowInvalid(false)
                                .setHelpText(helpText)
                                .build();
        currentCell.setDataValidation(rule);
        
        currentCell = this.sheet.getRange('B7');
        helpText = "Party time must be a valid time";
        rule = SpreadsheetApp.newDataValidation()
                                .requireDate()
                                .setAllowInvalid(false)
                                .setHelpText(helpText)
                                .build();
        currentCell.setDataValidation(rule);
        
        var partyType = this.sheet.getRange('B10').getDisplayValue();
        currentCell = this.sheet.getRange('B8');
        if (partyType == "In-store") {
            helpText = "Party length must be either 1.5 or 2 hours";
            rule = SpreadsheetApp.newDataValidation()
                                .requireValueInList(['1.5','2'])
                                .setAllowInvalid(false)
                                .setHelpText(helpText)
                                .build();
        } else {
            helpText = "Party length must be either 1 or 1.5 hours";
            rule = SpreadsheetApp.newDataValidation()
                                .requireValueInList(['1','1.5'])
                                .setAllowInvalid(false)
                                .setHelpText(helpText)
                                .build();
        }
        currentCell.setDataValidation(rule);
        
        currentCell = this.sheet.getRange('B10');
        helpText = "Party type cannot be edited. To change store location or convert to a mobile party, you must delete this booking and create a new one.";
        rule = SpreadsheetApp.newDataValidation()
                                .requireTextEqualTo(currentCell.getValue())
                                .setAllowInvalid(false)
                                .setHelpText(helpText)
                                .build();
        currentCell.setDataValidation(rule);
        
        if (partyType == "In-store") {
            currentCell = this.sheet.getRange('B11');
            helpText = "An In-store location cannot be changed. To move the booking to a different store, delete this booking and create a new one";
            rule = SpreadsheetApp.newDataValidation()
                                .requireTextEqualTo(currentCell.getDisplayValue())
                                .setAllowInvalid(false)
                                .setHelpText(helpText)
                                .build();
            currentCell.setDataValidation(rule);
        }
        
        currentCell = this.sheet.getRange('B12');
        helpText = "Booking ID cannot be edited.";
        rule = SpreadsheetApp.newDataValidation()
                                .requireTextEqualTo(currentCell.getValue())
                                .setAllowInvalid(false)
                                .setHelpText(helpText)
                                .build();
        currentCell.setDataValidation(rule);
    }

    /**
     * Disables editing on all cells.
     * User must manually re-enable editing, which will authorise background scripts
     */
    this.lockDownCells = function() {
        var range = this.sheet.getRange('B1:B11'); // stop before booking ID, since this will be added when event is created
        var helpText = "Booking cannot be edited until you select 'Edit / Delete Booking' -> 'Enable Editing', and follow the prompts";
        
        for(i = 1; i <= range.getHeight(); i++) {
          var currentCell = range.getCell(i, 1);
          var rule = SpreadsheetApp.newDataValidation().requireTextEqualTo(currentCell.getDisplayValue()).setAllowInvalid(false).setHelpText(helpText).build();
          currentCell.setDataValidation(rule);
        }
    }

    /**
     * Clears the sheet and resets formatting
     */
    this.resetSheet = function() {
        
        var range = this.sheet.getRange('B1:B9');
        range.clear({ contentsOnly : true });
        this.sheet.getRange('B10').setValue('In-store');
        this.applyValidationToLocationCell();
        this.sheet.getRange('B12').setValue('YES');
        
        // reset the formatting of the cells
        range = this.sheet.getRange('B1:B9');
        range.setFontFamily("Arial");
        range.setFontSize(14);
        range.setFontColor("black");
        range.setHorizontalAlignment("right");
        range.setBorder(true, true, true, true, true, true);
        
        range = this.sheet.getRange('B10:B12');
        range.setFontFamily("Arial");
        range.setFontSize(14);
        range.setFontColor("black");
        range.setHorizontalAlignment("center");
        range.setBorder(true, true, true, true, true, true);
        
        range = this.sheet.getRange('B6');
        range.setNumberFormat("d mmmm yyy");
        
        range = this.sheet.getRange('B7');
        range.setNumberFormat("h:mm am/pm");
    }
}