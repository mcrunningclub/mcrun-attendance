// NOTE TO SELF: TAKE THIS OUT OF OUR CODE AND HIDE IT SOMEWHERE
 
/**
 * Configure the service using the OAuth2 library: https://github.com/googleworkspace/apps-script-oauth2.
 * 
 * 
 * @trigger Form Submission.
 */
function getStravaService() {
  // Create a new service called "Strava"
  return OAuth2.createService('Strava')
  // Set the endpoint URL for Strava auth
    .setAuthorizationBaseUrl('https://www.strava.com/oauth/authorize')
    .setTokenUrl('https://www.strava.com/oauth/token')
    // Set the client ID and secret
    .setClientId(CLIENT_ID)
    .setClientSecret(CLIENT_SECRET)
      // Set the name of the callback function in the script referenced
      // above that should be invoked to complete the OAuth flow.
      // (see the authCallback function below)
    .setCallbackFunction('authCallback')
    // Set the property store where authorized tokens should be persisted.
    .setPropertyStore(PropertiesService.getUserProperties())
    // Set the scopes to request (space-separated for Google services).
    .setScope('activity:read_all');
}
 
// handle the callback
function authCallback(request) {
  var stravaService = getStravaService();
  var isAuthorized = stravaService.handleCallback(request);
  if (isAuthorized) {
    return HtmlService.createHtmlOutput('Success! You can close this tab.');
  } else {
    return HtmlService.createHtmlOutput('Denied. You can close this tab');
  }
}