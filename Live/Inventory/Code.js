//import 'google-apps-script'

var PRIVATE_KEY = "-----BEGIN PRIVATE KEY-----\nMIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQCtQFlPO9DmDk58\n4g8eLtX4h/HCxXsNo/OT7Dz9TMHQD3CvX7dBRHSbbyDtGcxTpq+KrN18kPyJmChS\nb3Tncptiq0hIUpy/+91+6qktTCEaXbZKVpZ0uSvzf8Mf1LhVMBRTkl5EcTAP0+cD\n6iPzZht2+P1VZf7rd4ba4qbPqM4curbTlBCihMXIqy22IIvGDj61d1SXCOnWHs4S\nWm1glaMdPB+vsaQLAbU1h9Er967kP5sCaIHwzY+Gzmv7EiFr3y7KlzOyLCPa8PNm\nGSl6FRppzVMv5lkPQsiUaSQK4j3Ol+3wWYqW/ECNfW+VuwYmI0nMSWU5wdo2KtVG\nNKKoC4RtAgMBAAECggEAAoOXiC3PBzeX7fn9zCtT0YpveKsS8Qy7AR+Bdw+BFHrU\n4MabyyeqJYNEUAx6yY/2piWCBUe5UmnR0/hoEt+334OqxdnlCmgmO6w+Djk3lcFc\nXtHI1yLEv4DQHQsiLaJH+Tp7gbS+xMwHYygno2WM6noMewvC2jnezBhT4VmKvCH3\nW1mMFv2MOxUuf6+gFkNYeSUwS35kJfGzG9So5IH1tm1COW1I1LDNLi/CiinMVrsW\nEH2tjRqui+lmaB3r0ZDOujvu2ImKcoW6sl7qf+Dhas7eQUTaqUTGJ5+nrjCf368y\n52zNZasrw0I7+K5tN7JSNosiNLhjjgJo/RrNvmFPAQKBgQDgbNG26ffH/Ak94Aou\n7t19U39js2cOkcnkI4zU1NwfJ2pdRujxobbTH6Ffy4QDAlCZxIV9M/Qcsc6XaN2o\nmIMoUuROtqoEEZ+CpAElEYWLWCP3L5FY5/5wHaz2YxQvy14+CgeUHKz5Y/z8IT4m\necxrFJ+wmqWB0z8QRMN3y+HcrQKBgQDFoGRI4ulDnn1wi3z8OzZ6uQINPVtszwFz\nWKhAvqQwrSVNZWquHGcuztmYbhbvTpJ1YGmaMmxAVWrh5DhCDIdJNy89JnmDo5V8\nk1vjWnXn/FGWHAd+GfWgxNn6Az9d4QTEdZvE0emtVrn+6xf0XiEOZosV1qxJvi+b\nJH4hHqN+wQKBgQDKXzZI4/gMvOg0hIeKRNkzfwy7gfYnfC167Ne8v+lyql9Ol3fN\nFE9BWB9zu5hiAj9eOYlKGoRBL9EkVWqz8jsrLHw1wp/TJXUaH/vsSj2LJsLfzmQZ\nsLGOtiPW1gdJBfEIrpCg7a7JAHILhYp+tYww7xsE7J7cT/ppGCjPKOmVzQKBgEyg\ncyH7sZyBYHv56d1XDDmrcIs3pjJbVWGnF537DWi+Sf9nemTGKI/yrlY3IXdqjMks\nN+YM9QJA3G938QRTHUWbOxrHx0fubrDd5jwSQDNSF0RP2+veHupWSXpyNeitrg6K\n13oKNkP6o6We/CvJL6IIypcOJMF3F7hc/vbSjWxBAoGBANZP88OVDo1TDLRtSU2J\nrmIFz7ELCqPxkGoMr0vkNur9OYMtgWBR9HihqHPfIJHUH60jLgCvm7R38ZkgUPnS\n4mj97TMOKjwL0QxK6ssttQBfhrBx2Sbm0nag7UbAq1kceRQomGAKfwKYsPlhjkV9\noCePA3Xee7JhlT74ma6ADomy\n-----END PRIVATE KEY-----\n"
var ACCOUNT_EMAIL = "apps-script@inventory-79ffa.iam.gserviceaccount.com"
var PROJECT_ID = "inventory-79ffa"

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu('FIZZ KIDZ OPTIONS')
  .addItem('Get latest values', 'showPullConfirmationDialogue')
  .addItem('Write to database', 'showPushConfirmationDialogue')
  .addToUi();
  
  pull()
}

function push() {
  writeToDB("MALVERN")
  writeToDB("WAREHOUSE")
  writeToDB("BALWYN")
}

function pull(){
  pullFromDB("MALVERN")
  pullFromDB("WAREHOUSE")
  pullFromDB("BALWYN")
}

function showPullConfirmationDialogue() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(
    "Refresh the spreadsheet",
    "This will replace everything within each location with the data from the database", 
    ui.ButtonSet.OK_CANCEL);
  
  if (result == ui.Button.OK) {
    pull()
  }
}

function showPushConfirmationDialogue() {
 
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(
    "Write to database",
    "WARNING: Quantities for all items in all locations will be replaced by the values currently shown.", 
    ui.ButtonSet.OK_CANCEL);
  
  if (result == ui.Button.OK) {
    push()
  }
}

function testPull() {

  var firestore = FirestoreApp.getFirestore(ACCOUNT_EMAIL, PRIVATE_KEY, PROJECT_ID)
  
  var warehouseItems = firestore.getDocuments("WAREHOUSE")
  var warehouseIds = firestore.getDocumentIds("WAREHOUSE")
  console.log(warehouseIds)
  for (var i = 0; i < warehouseItems.length; i++) {
    firestore.createDocument("MALVERN/"+warehouseIds[i],warehouseItems[i].fields)
    firestore.createDocument("BALWYN/"+warehouseIds[i],warehouseItems[i].fields)
  }
}

function pullFromDB(location) {
  
  var firestore = FirestoreApp.getFirestore(ACCOUNT_EMAIL, PRIVATE_KEY, PROJECT_ID)
  
  var sheet = SpreadsheetApp.getActive().getSheetByName(location)
  
  var documentIds = firestore.getDocumentIds(location)
  
  var row = 2
  for(var i = 0; i < documentIds.length; i++) {
    var document = firestore.getDocument(location + '/' + documentIds[i])
    sheet.getRange('A'+row).setValue(documentIds[i])
    sheet.getRange('B'+row).setValue(document.fields.DISP_NAME)
    sheet.getRange('C'+row).setValue(document.fields.QTY)
    sheet.getRange('D'+row).setValue(document.fields.UNIT)
    row++
  }
  
  var range = sheet.getRange('A2:D' + sheet.getLastRow())
  range.sort({column: 2, ascending: true})
}

function writeToDB(location) {
  
  var firestore = FirestoreApp.getFirestore(ACCOUNT_EMAIL, PRIVATE_KEY, PROJECT_ID)
  
  var sheet = SpreadsheetApp.getActive().getSheetByName(location)
  var cells = sheet.getRange('A2:' + sheet.getLastRow()).getValues()
  
  for (var i = 0; i < cells.length; i++) {
    var currentItem = cells[i]
    var qty = currentItem[2]
    var response = firestore.getDocument(location + '/' + currentItem[0])
    response.fields.QTY = qty
    firestore.updateDocument(location + '/' + currentItem[0],response.fields,false)
  }
}

function move() {
 
  var firestore = FirestoreApp.getFirestore(ACCOUNT_EMAIL, PRIVATE_KEY, PROJECT_ID)
  
  var moveSheet = SpreadsheetApp.getActive().getSheetByName("MOVE STOCK")
  var range = moveSheet.getRange('A2:D2').getValues()[0]
  var itemID = range[0]
  var from = range[1]
  var to = range[2]
  var qty = range[3]
  
  var fromSheet = SpreadsheetApp.getActive().getSheetByName(from)
  var toSheet = SpreadsheetApp.getActive().getSheetByName(to)
  
  // remove where its coming from
  var fromResponse = firestore.getDocument(from + '/' + itemID)
  fromResponse.fields.QTY -= qty
  firestore.updateDocument(from + '/' + itemID, fromResponse.fields, false)
  
  // add to where its going
  var toResponse = firestore.getDocument(to + '/' + itemID)
  toResponse.fields.QTY += qty
  firestore.updateDocument(to + '/' + itemID, toResponse.fields, false)
  
  pull()
}

function automateStock() {
  
  // 1. Get all parties within past week
  // 2. for each store, for each creation, remove ingredients * num of kids

  var firestore = FirestoreApp.getFirestore(ACCOUNT_EMAIL, PRIVATE_KEY, PROJECT_ID)
  var RESPONSES_FILE_ID = '1C2QOmdKoODDO0MOopJeSehTdeUvZxL4F9kDafCCfNgM'

  var responsesFile = SpreadsheetApp.openById(RESPONSES_FILE_ID),
      inStoreResponsesSheet = responsesFile.getSheetByName("In-Store Responses"),
      mobileResponsesSheet = responsesFile.getSheetByName("Mobile Responses"),
      today = new Date(),
      dateToday = new Date(today.getFullYear(), today.getMonth(), today.getDate() - 2),
      dateMonday = new Date(dateToday.getFullYear(), dateToday.getMonth(), dateToday.getDate() - 7)
  
  console.log(dateToday, dateMonday)
  
  // start with In-Store
  automateInStore(inStoreResponsesSheet, dateToday, dateMonday, firestore)
  
  // then do mobile
  automateMobile(mobileResponsesSheet, dateToday, dateMonday, firestore)
}

function automateMobile(mobileResponsesSheet, dateToday, dateMonday, firestore) {
  
  var endRow = mobileResponsesSheet.getLastRow()
  for (var row = 2; row <= endRow; row++) {
    var dateOfParty = mobileResponsesSheet.getRange('B'+row).getValue()
    
    // ensure we are within range, or break
    if (dateOfParty < dateMonday || dateOfParty > dateToday) { continue }
    
    // within range of past weekend
    var creations = mobileResponsesSheet.getRange('G'+row).getValue().split(', ')
    creations = creations.concat(mobileResponsesSheet.getRange('H'+row).getValue().split(', '))
    var childrenCount = mobileResponsesSheet.getRange('F'+row).getDisplayValue()
    
    for (var i = 0; i < creations.length; i++) {
      adjustCreationStock(creations[i], "BALWYN", childrenCount, firestore)
    }
    
    // adjust general mobile items
    adjustItem("FIZZ_BAGS","BALWYN", 1, firestore)
    adjustItem("TABLE_CLOTH", "BALWYN", 1, firestore)
  }
  
}

function automateInStore(inStoreResponsesSheet, dateToday, dateMonday, firestore) {
  
  var endRow = inStoreResponsesSheet.getLastRow()
  for (var row = 2; row <= endRow; row++) {
    var dateOfParty = inStoreResponsesSheet.getRange('B'+row).getValue()
    
    // ensure we are within range, or break
    if (dateOfParty < dateMonday || dateOfParty > dateToday) { continue }
 
    // within range of past weekend
    var creations = inStoreResponsesSheet.getRange('H'+row).getValue().split(', ')
    creations = creations.concat(inStoreResponsesSheet.getRange('I'+row).getValue().split(', '))
    var childrenCount = inStoreResponsesSheet.getRange('G'+row).getDisplayValue().substring(5,7)
    var location = inStoreResponsesSheet.getRange('F'+row).getDisplayValue().toUpperCase()
    console.log("ENTERED - LOCATION: " + location)
    
    for (var i = 0; i < creations.length; i++) {
      adjustCreationStock(creations[i], location, childrenCount, firestore)
    }
    
    // finally, adjust food and general items
    adjustFoodStock(location, childrenCount, firestore)
    adjustGeneralStock(location, childrenCount, firestore)
  }
}

function adjustCreationStock(ingredient, location, count, firestore) {

  if (ingredient == "") { return }
  
  var response = firestore.getDocument('USAGE/' + ingredient)
  
  for(var key in response.fields) {
  
    var amountUsed = response.fields[key]
    var current = firestore.getDocument(location + '/' + key)
    current.fields.QTY -= amountUsed * count
    if (current.fields.QTY < 0) {
      current.fields.QTY = 0
    }
    firestore.updateDocument(location + '/' + key, current.fields, false)
  }
}

function adjustFoodStock(location, count, firestore) {
  
  // Numbers taken from 'Flexible Party Food Menu'
  if (count <= 15) {
    adjustItem("CORN_CHIPS", location, 1, firestore)
    adjustItem("POTATO_CHIPS", location, 1, firestore)
    adjustItem("WAFERS", location, 1, firestore)
    adjustItem("POPCORN", location, 1, firestore)
    adjustItem("SAUSAGE_ROLLS", location, (+count + 3) * (1/24), firestore)
    adjustItem("PARTY_PIES", location, (+count + 3) * (1/24), firestore)
  } else if (count <= 20) {
    adjustItem("CORN_CHIPS", location, 1, firestore)
    adjustItem("POTATO_CHIPS", location, 1, firestore)
    adjustItem("WAFERS", location, 2, firestore)
    adjustItem("POPCORN", location, 1, firestore)
    adjustItem("SAUSAGE_ROLLS", location, (+count + 5) * (1/24), firestore)
    adjustItem("PARTY_PIES", location, (+count + 5) * (1/24), firestore)
  } else if (count <= 28) {
    adjustItem("CORN_CHIPS", location, 2, firestore)
    adjustItem("POTATO_CHIPS", location, 2, firestore)
    adjustItem("WAFERS", location, 2, firestore)
    adjustItem("POPCORN", location, 1, firestore)
    adjustItem("SAUSAGE_ROLLS", location, (+count + 5) * (1/24), firestore)
    adjustItem("PARTY_PIES", location, (+count + 5) * (1/24), firestore)
  }
}

/*
 * Adjust an particular items qty by a certain amount
 */
function adjustItem(item, location, amount, firestore) {
  var itemResponse = firestore.getDocument(location + '/' + item)
  itemResponse.fields.QTY -= amount
  if (itemResponse.fields.QTY < 0) {
    itemResponse.fields.QTY = 0
  }
  firestore.updateDocument(location + '/' + item, itemResponse.fields, false)
}
    

function adjustGeneralStock(location, count, firestore) {
 
  var response = firestore.getDocument("USAGE/GENERAL_" + location)
  for(var key in response.fields) {
    var amountUsed = response.fields[key]
    var current = firestore.getDocument(location + '/' + key)
    current.fields.QTY -= amountUsed * count
    if (current.fields.QTY < 0) {
      current.fields.QTY = 0
    }
    firestore.updateDocument(location + '/' + key, current.fields, false)
  }
}