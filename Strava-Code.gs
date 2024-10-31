// const CLIENT_ID = '***REMOVED***';
// const CLIENT_SECRET = '***REMOVED***';

function getOAuthService() {
  return OAuth2.createService('Strava')
    .setAuthorizationBaseUrl('https://www.strava.com/oauth/authorize')
    .setTokenUrl('https://www.strava.com/oauth/token')
    .setClientId(CLIENT_ID)
    .setClientSecret(CLIENT_SECRET)
    .setRedirectUri(ScriptApp.getService().getUrl())
    .setPropertyStore(PropertiesService.getUserProperties())
    .setScope('activity:read_all'); // Or other scopes needed
}


function authorize() {
  const service = getOAuthService();
  if (!service.hasAccess()) {
      const authorizationUrl = service.getAuthorizationUrl();
      Logger.log('Authorize the script by visiting this URL: %s', authorizationUrl);
  } else {
      Logger.log('Authorization successful!');
  }
}


function getActivities() {
  const service = getOAuthService();
  if (service.hasAccess()) {
      const url = 'https://www.strava.com/api/v3/athlete/activities';
      const response = UrlFetchApp.fetch(url, {
          headers: {
              Authorization: `Bearer ${service.getAccessToken()}`
          }
      });
      Logger.log(response.getContentText());
  } else {
      Logger.log('No access yet. Run authorize() to authenticate.');
  }
}



function test(){
  var StravaApiV3 = require('strava_api_v3');
  var defaultClient = StravaApiV3.ApiClient.instance;

  // Configure OAuth2 access token for authorization: strava_oauth
  var strava_oauth = defaultClient.authentications['strava_oauth'];
  strava_oauth.accessToken = "***REMOVED***";

  var api = new StravaApiV3.ClubsApi();

  var id = ***REMOVED***; // {Long} The identifier of the club.

  var opts = { 
    'page': 1, // {Integer} Page number. Defaults to 1.
    'perPage': 56 // {Integer} Number of items per page. Defaults to 30.
  };

  var callback = function(error, data, response) {
    if (error) {
      console.error(error);
    } else {
      console.log('API called successfully. Returned data: ' + data);
    }
  };
  var res = api.getClubActivitiesById(id, opts, callback);

  Logger.log(res);

}
