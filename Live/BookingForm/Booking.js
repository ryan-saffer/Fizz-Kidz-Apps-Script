/**
 * {Abstract}
 * Booking object used to create and manage bookings
 * Subclass is either a NewBooking or ExistingBooking
 * 
 * @constructor
 * @param {Sheet} sheet the sheet of the booking - see https://developers.google.com/apps-script/reference/spreadsheet/sheet
 */
function Booking(sheet) {

  //================================================================================
  // Properties
  //================================================================================

  this.sheet = sheet

  this.parentName   = sheet.getRange('B1').getDisplayValue()
  this.mobileNumber = sheet.getRange('B2').getDisplayValue();
  this.emailAddress = sheet.getRange('B3').getDisplayValue();
  this.childName    = sheet.getRange('B4').getDisplayValue();
  this.childAge     = sheet.getRange('B5').getDisplayValue();
  this.dateOfParty  = sheet.getRange('B6').getValue();
  this.timeOfParty  = sheet.getRange('B7').getValue();
  this.partyLength  = sheet.getRange('B8').getDisplayValue()
  this.notes        = sheet.getRange('B9').getDisplayValue();
  this.partyType    = sheet.getRange('B10').getDisplayValue();
  this.location     = sheet.getRange('B11').getDisplayValue();

  //================================================================================
  // Methods
  //================================================================================

  /**
   * Validates the party values
   * 
   * @throws an error if any value invalid
   */
  this.validateFields = function() {

    if(this.parentName == "") {
        Browser.msgBox("⚠️You must enter the parents name. Party not booked/updated. Try again.");
        throw new Error("You must enter the parents name. Operation cancelled.");
    }
    
    if(this.mobileNumber == "") {
        Browser.msgBox("⚠️You must enter the mobile number. Party not booked/updated. Try again.");
        throw new Error("You must enter the mobile number. Operation cancelled.");
    }
    if (this.mobileNumber.length != 10) {
        Browser.msgBox("⚠️Mobile number is not valid. Party not booked/updated. Try again.");
        throw new Error("Mobile number is not valid. Operation cancelled.");
    }
    
    if (this.emailAddress == "") {
        Browser.msgBox("⚠️You must enter the email address. Party not booked/updated. Try again.");
        throw new Error("You must enter the email address. Operation cancelled.");
    }
    if (!this.validateEmail(this.emailAddress)) {
        Browser.msgBox("⚠️Email address is not valid. Party not booked/updated. Try again.");
        throw new Error("Email address is not valid. Operation cancelled.");
    }
    
    if(this.childName == "") {
        Browser.msgBox("⚠️You must enter the childs name. Party not booked/updated. Try again.");
        throw new Error("You must enter the childs name. Operation cancelled.");
    }
    
    if(this.childAge == "") {
        Browser.msgBox("⚠️You must enter the childs age. Party not booked/updated. Try again.");
        throw new Error("You must enter the childs age. Operation cancelled.");
    }
    
    if(this.dateOfParty == "") {
        Browser.msgBox("⚠️You must enter the party date. Party not booked/updated. Try again.");
        throw new Error("You must enter the party date. Operation cancelled.");
    }
    if (!(this.dateOfParty instanceof Date)) {
        Browser.msgBox("⚠️Party date is invalid. Party not booked/updated. Try again.");
        throw new Error("Party date is invalid. Operation cancelled");
    }
    
    if(this.timeOfParty == "") {
        Browser.msgBox("⚠️You must enter the time of the this. Party not booked/updated. Try again.");
        throw new Error("You must enter the time of the this. Operation cancelled.");
    }
    if (!(this.timeOfParty instanceof Date)) {
        Browser.msgBox("⚠️Party time is invalid. Party not booked/updated. Try again.");
        throw new Error("Party time is invalid. Operation cancelled.");
    }
    if (this.timeOfParty.getFullYear() == 1900) {
        Browser.msgBox("⚠️Party time is invalid. Party not booked/updated. Try again.");
        throw new Error("Party time is invalid. Operation cancelled");
    }
    
    if(this.partyLength == "") {
        Browser.msgBox("⚠️You must enter the length of the party. Party not booked/updated. Try again.");
        throw new Error("You must enter the length of the party. Operation cancelled.");
    }
    
    if(this.partyType == "") {
        Browser.msgBox("⚠️You must enter the type of party as In-store or Mobile. Party not booked/updated. Try again.");
        throw new Error("You must enter the type of party as In-store or Mobile. Operation cancelled.");
    }

    // In-store must be 1.5 or 2 hours, Mobile must be 1 or 1.5 hours
    if (this.partyType == "In-store") {
        if (this.partyLength == "1") {
        Browser.msgBox("⚠️An In-store party cannot have a party length of 1 hour. Party not booked/updated. Try again.");
        throw new Error("An In-store party cannot have a party length of 1 hour. Operation cancelled.");
        }
    }
    if (this.partyType == "Mobile") {
        if (this.partyLength == "2") {
        Browser.msgBox("⚠️A Mobile party cannot have a party length of 2 hours Party not booked/updated. Try again.");
        throw new Error("A Mobile party cannot have a party length of 2 hours. Operation cancelled.");
        }
    }
    
    // mobile party must have a location
    if (this.partyType == "Mobile") {
        if (this.location == "") {
        Browser.msgBox("⚠️Mobile party must have a location. Party not booked/updated. Try again.");
        throw new Error("Mobile party must have a location. Operation cancelled.");
        }
    }
    // in-store party cannot have a location
    if (this.partyType == "In-store") {
        if (this.location != "Malvern" && this.location != "Balwyn") {
        Browser.msgBox("⚠️An In-store party location must be 'Malvern' or 'Balwyn'. Party not booked/updated. Try again.");
        throw new Error("In-store party location must be Malern or Balwyn. Operation cancelled.");
        }
    }
    
    if(this.confirmationEmailRequired == "") {
        Browser.msgBox("⚠️You must enter if a confirmation email is required. Party not booked/updated. Try again.");
        throw new Error("You must enter if a confirmation email is required. Operation cancelled.");
    }
  }

  /**
   * Validates the email address
   * 
   * @param {string} email the email address
   * @returns {boolean} true if email is valid
   */
  this.validateEmail = function(email) {
  
    // Uses a regex to ensure the entered email address is valid
    var re = /^(([^<>()[\]\\.,;:\s@\"]+(\.[^<>()[\]\\.,;:\s@\"]+)*)|(\".+\"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
    return re.test(email);
  }

  /**
   * Gets the Google Calendar ID for this location
   * 
   * @returns {String} the ID of the correct Calendar
   */
  this.getCalendarID = function() {

    // event IDs
    var malvernStorePartiesCalendarID = "fizzkidz.com.au_j13ot3jarb1p9k70c302249j4g@group.calendar.google.com";
    var balwynStorePartiesCalendarID = "fizzkidz.com.au_7vor3m1efd3fqbr0ola2jvglf8@group.calendar.google.com";
    var mobilePartiesCalendarID = "fizzkidz.com.au_b9aruprq8740cdamu63frgm0ck@group.calendar.google.com";
  
    if (this.partyType == "In-store") {
      if (this.location == "Malvern") {
        return malvernStorePartiesCalendarID;
      } else if (this.location == "Balwyn") {
        return balwynStorePartiesCalendarID;
      }
    } else {
      return mobilePartiesCalendarID;
    }
  }

  /**
   * Determines the parties end date/time based on starting time and length
   * 
   * @returns {Date} the date and time the party ends
   */
  this.determineEndDate = function() {
  
    // determine when party ends
    var lengthHours = 0;
    var lengthMinutes = 0;
    switch (this.partyLength) {
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
    
    var endDate = new Date(this.dateOfParty.getFullYear(), this.dateOfParty.getMonth(), this.dateOfParty.getDate(), this.timeOfParty.getHours() - 1 + lengthHours, this.timeOfParty.getMinutes() + lengthMinutes);
    
    return endDate;
  }
}