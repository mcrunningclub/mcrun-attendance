const STRAVA_BASE_URL = 'https://www.strava.com/api/v3/'
const ACTIVITIES_ENDPOINT = 'athlete/activities'
const MAPS_FOLDER = 'run_maps'

/**
 * Maps an Object containing param, value pairs to a query string.
 * Ex: {"param1": val1, "param2": val2} -> "?param1=val1&param2=val2"
 *
 * @author [Jikael Gagnon](<jikael.gagnon@mail.mcgill.ca>)
 * @date  Nov 7, 2024
 * @update  Nov 7, 2024
 */

function query_object_to_string(query_object) {
  if (query_object.length === 0) {
    return ''
  }

  var param_value_list = Object.entries(query_object);
  var param_strings = param_value_list.map(([param, value]) => `${param}=${value}`);
  var query_string = param_strings.join('&');
  return '?' + query_string;
}


/**
 * Makes an API request to the given endpoint with the given query
 *  Ex: 'clubs/693906/activities', {"param1": val1, "param2": val2} -> API response
 *
 * @author [Jikael Gagnon](<jikael.gagnon@mail.mcgill.ca>)
 * @date  Nov 7, 2024
 * @update  Nov 7, 2024
 */

function callStravaAPI(endpoint, query_object) {
  // Set up the service
  var service = getStravaService();

  if (service.hasAccess()) {
    Logger.log('App has access.');

    // API Endpoint
    var endpoint = STRAVA_BASE_URL + endpoint;
    // Get string in for "?param1=val1&param2=val2&...&paramN=valN"
    var query_string = query_object_to_string(query_object);

    var headers = {
      Authorization: 'Bearer ' + service.getAccessToken()
    };

    var options = {
      headers: headers,
      method: 'GET',
      muteHttpExceptions: false,
    };

    // Get response from API
    const urlString = endpoint + query_string;
    var response = JSON.parse(UrlFetchApp.fetch(urlString, options));
    return response;
  }
  else {
    Logger.log("App has no access yet.");

    // Open this url to gain authorization from Strava
    var authorizationUrl = service.getAuthorizationUrl();
    Logger.log("Open the following URL and re-run the script: %s", authorizationUrl);
  }
}


function strava_main() {
  /** Activity example */
  // var endpoint = 'activities/13889807290';
  // var query_object = {"include_all_efforts" : true};
  // var response = callStravaAPI(endpoint, query_object);

  /** Club activites example */
  // var endpoint = '/clubs/693906/activities'
  // var endpoint = '/athletes/61928790';
  // var response = callStravaAPI(endpoint, {})
  // saveMapToFile(response, 'example.png')

  /** Individual athlete example */
  var endpoint = 'athlete/activities'
  var response = callStravaAPI(endpoint, {});

  const stats = getRunStats(response[0]);
  console.log(stats);

  //saveMapToFile(activity0, 'test-march-22.png');
}


/**
 * Extract target run stats from Strava activity.
 * 
 * @see 'https://developers.strava.com/docs/reference/#api-models-SummaryActivity'
 * 
 * @param {object} activity  A Strava object `SummaryActivity`.
 * @return {object}  Extracted stats from `activity`.
 * 
 * @author [Andrey S Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Mar 22, 2025
 * @update  Mar 22, 2025
 */

function getRunStats(activity) {
  const targetStats = [
    'name',
    'distance',
    'moving_time',
    'elapsed_time',
    'total_elevation_gain',
    'average_speed',
    'max_speed',
    'map',
  ];

  const extracted = targetStats.reduce((acc, aStat) => {
    acc[aStat] = activity[aStat];
    return acc;
  }, {});

  return extracted;
}


/**
 * Takes a response for a given activity from the Strava API and saves an image of the map to the
 * desired location.
 * 
 * @author [Jikael Gagnon](<jikael.gagnon@mail.mcgill.ca>)
 * @date  Dec 1, 2024
 * @update  Dec 1, 2024
 */

function saveMapToFile(api_response, filename) {
  var polyline = api_response['map']['summary_polyline']
  var map = Maps.newStaticMap();
  map.addPath(polyline)
  DriveApp.createFile(Utilities.newBlob(map.getMapImage(), 'image/png', filename));
}

/**
 * Finds the most recent head run submission and returns the timestamp as a Date object
 * @author [Jikael Gagnon](<jikael.gagnon@mail.mcgill.ca>)
 * @date  Dec 1, 2024
 * @update  Dec 1, 2024
 */

function getLatestSubmissionTimestamp() {
  const sheet = ATTENDANCE_SHEET;
  const lastRow = sheet.getLastRow();
  var timestamp = sheet.getRange(lastRow, TIMESTAMP_COL).getValue();
  return new Date(timestamp);
}


/**
 * Converts a Date timestamp to a Unix Epoch timestamp
 * (the number of seconds that have elapsed since January 1, 1970)
 * @author [Jikael Gagnon](<jikael.gagnon@mail.mcgill.ca>)
 * @date  Dec 1, 2024
 * @update  Dec 1, 2024
 */

function getUnixEpochTimestamp(timestamp) {
  return Math.floor(timestamp.getTime() / 1000);
}


/**
 * Saves file to MAPS_FOLDER/<Unix Epoch timestamp of submisstion>.png
 * (the number of seconds that have elapsed since January 1, 1970)
 * @author [Jikael Gagnon](<jikael.gagnon@mail.mcgill.ca>)
 * @date  Dec 1, 2024
 * @update  Dec 1, 2024
 */

function getSaveLocation(submissionTime) {
  return MAPS_FOLDER + '/' + submissionTime.toString() + '.png'
}


/**
 * Gets the most recent head run submission and saves the map
 * of the corresponding Strava activity to MAPS_FOLDER/<Unix Epoch timestamp of submisstion>.png
 *
 * @author [Jikael Gagnon](<jikael.gagnon@mail.mcgill.ca>)
 * @date  Dec 1, 2024
 * @update  Dec 1, 2024
 */

function getMapForLatestRun() {
  const sheet = ATTENDANCE_SHEET;
  var submissionTimestamp = getLatestSubmissionTimestamp();
  var now = new Date();
  var subEpochTime = getUnixEpochTimestamp(submissionTimestamp);
  var nowEpochTime = getUnixEpochTimestamp(now);
  var query_object = { 'after': subEpochTime, 'before': nowEpochTime }
  const endpoint = ACTIVITIES_ENDPOINT
  var response = callStravaAPI(endpoint, query_object)

  if (response.length == 0) {
    // Create an instance of ExecutionError with a custom message
    var errorMessage = "No Strava activity has been found for the run that occured on " + submissionTimestamp.toString();
    throw new Error(errorMessage); // Throw the ExecutionError
  }

  var activity = response[0]
  var saveLocation = getSaveLocation(subEpochTime)
  saveMapToFile(activity, saveLocation)
}


/**
 * Configure the service using the OAuth2 library: https://github.com/googleworkspace/apps-script-oauth2.
 *
 * @trigger Form Submission.
 */

function getStravaService() {
  // Get CLIENT_ID & CLIENT_SECRET from Script Properties
  const scriptProperties = PropertiesService.getScriptProperties();
  const CLIENT_ID = scriptProperties.getProperty(SCRIPT_PROPERTY.clientID);
  const CLIENT_SECRET = scriptProperties.getProperty(SCRIPT_PROPERTY.clientSecret);

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
    .setCallbackFunction('authCallback_')
    // Set the property store where authorized tokens should be persisted.
    .setPropertyStore(PropertiesService.getUserProperties())
    // Set the scopes to request (space-separated for Google services).
    .setScope('activity:read_all,profile:read_all')
    ;

  // Handle the callback with helper
  function authCallback_(request) {
    var stravaService = getStravaService();
    var isAuthorized = stravaService.handleCallback(request);
    if (isAuthorized) {
      return HtmlService.createHtmlOutput('Success! You can close this tab.');
    } else {
      return HtmlService.createHtmlOutput('Denied. You can close this tab');
    }
  };
}

