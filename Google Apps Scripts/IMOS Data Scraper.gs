// Define the needed URL and payload to post
const churchUserNameURL = "https://id.churchofjesuschrist.org/api/v1/authn";
const churchPasswordURL = "https://id.churchofjesuschrist.org/api/v1/authn/factors/password/verify?rememberDevice=false"

//Define authentication credentials
const IMOS_USERNAME = "YOUR IMOS USERNAME HERE";
const IMOS_PASSWORD = "YOUR IMOS PASSWORD HERE";

function authn(){
  const userNamePayload = {
    "username": IMOS_USERNAME,
    "options":{"warnBeforePasswordExpired":true,"multiOptionalFactorEnroll":true}
  };
  
  const userNameOptions = {
    "method" : "post",
    "contentType" : "application/json",
    "payload" : JSON.stringify(userNamePayload)
  };
  
  let sessionData = JSON.parse(
    UrlFetchApp.fetch(
      churchUserNameURL,
      userNameOptions
    )
  );    
  
  //Logger.log(sessionData.stateToken);

  const passwordPayload = {
    "password": IMOS_PASSWORD,
    "stateToken": sessionData.stateToken
  };
  
  const passwordOptions = {
    "method" : "post",
    "contentType" : "application/json",
    "payload" : JSON.stringify(passwordPayload),
    "muteHttpExceptions" : true
  };
  
  Logger.log(
    UrlFetchApp.fetch(
      churchPasswordURL,
      passwordOptions,
    )
  );//*/
};

