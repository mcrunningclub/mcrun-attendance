const CLIENT_ID = '***REMOVED***';
const CLIENT_SECRET = '***REMOVED***';


function test2() {
  // Make a POST request with form data.
  var resumeBlob = Utilities.newBlob('Hire me!', 'text/plain', 'resume.txt');
  var formData = {
    'name': 'Bob Smith',
    'email': 'bob@example.com',
    'resume': resumeBlob
  };
  // Because payload is a JavaScript object, it is interpreted as
  // as form data. (No need to specify contentType; it automatically
  // defaults to either 'application/x-www-form-urlencoded'
  // or 'multipart/form-data')
  var options = {
    'method' : 'post',
    'payload' : formData
  };
  var response = UrlFetchApp.fetch('https://httpbin.org/post', options);
  var result = JSON.parse(response.getContentText());
  Logger.log(response.getResponseCode());
  Logger.log(JSON.stringify(result, null, 2));
}

/**
 * Authorizes and makes a request to the Strava API.
 */
function run() {
  var service = getService_();
  if (service.hasAccess()) {
    var url = 'https://www.strava.com/api/v3/activities';
    var options = {
      headers: {
        Authorization: 'Bearer ' + service.getAccessToken()
      },
      muteHttpExceptions: true
    };

    var response = UrlFetchApp.fetch(url, options);
    var result = JSON.parse(response.getContentText());
    Logger.log(JSON.stringify(result, null, 2));
  } else {
    var authorizationUrl = service.getAuthorizationUrl();
    Logger.log('Open the following URL and re-run the script: %s',
        authorizationUrl);
  }
}

/**
 * Reset the authorization state, so that it can be re-tested.
 */
function reset() {
  var service = getService_();
  service.reset();
}

/**
 * Configures the service.
 * Three required and optional parameters are not specified
 * because the library creates the authorization URL with them
 * automatically: `redirect_url`, `response_type`, and
 * `state`.
 */
function getService_() {
  return OAuth2.createService('Strava')
    // Set the endpoint URLs.
    .setAuthorizationBaseUrl('https://www.strava.com/oauth/authorize')
    .setTokenUrl('https://www.strava.com/oauth/token')

    // Set the client ID and secret.
    .setClientId(CLIENT_ID)
    .setClientSecret(CLIENT_SECRET)

    // Set the name of the callback function that should be invoked to
    // complete the OAuth flow.
    .setCallbackFunction('authCallback_')

    // Set the property store where authorized tokens should be persisted.
    .setPropertyStore(PropertiesService.getUserProperties())
    .setScope('read,read_all,profile:read_all,profile:write,activity:read,activity:read_all,activity:write');


    // .setRedirectUri(ScriptApp.getService().getUrl())
}

/**
 * Handles the OAuth callback.
 */
function authCallback_(request) {
  var service = getService_();
  var authorized = service.handleCallback(request);
  if (authorized) {
    return HtmlService.createHtmlOutput('Success!');
  } else {
    return HtmlService.createHtmlOutput('Denied.');
  }
}

/**
 * Logs the redict URI to register.
 */
function logRedirectUri() {
  Logger.log(OAuth2.getRedirectUri());
}