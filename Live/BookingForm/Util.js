/**
 * Returns the email address that the email should be sent from based on party location
 * If Malvern, send from "malvern@fizzkidz.com.au"
 * If Balwyn or Mobile, send from "info@fizzkidz.com.au"
 * 
 * @param {String} location the location of the store
 * @returns {String} email address to send from
 */
function determineFromEmailAddress(location) {

  if(location == "Malvern") {
    // send from malvern@fizzkidz.com.au
    
    return "malvern@fizzkidz.com.au";
  }
  else if(location == "Balwyn") {
    // send from info@fizzkidz.com.au

    return "info@fizzkidz.com.au";
  }
  else { // mobile party
    // send from info@fizzkidz.com.au
    return "info@fizzkidz.com.au";
  }
}

/**
 * Gets the correct managers signature depending on who email is being sent from
 * 
 * @param {String} fromAddress the email address sending the email
 * @returns {String} the signature
 */
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