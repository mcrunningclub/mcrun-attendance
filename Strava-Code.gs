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
  if (Object.keys(query_object).length === 0) {
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

  // set up the service
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
      muteHttpExceptions: true
    };

    // Get response from API
    var response = JSON.parse(UrlFetchApp.fetch(endpoint + query_string, options));

    return response;

  }
  else {
    Logger.log("App has no access yet.");

    // open this url to gain authorization from github
    var authorizationUrl = service.getAuthorizationUrl();

    Logger.log("Open the following URL and re-run the script: %s",
      authorizationUrl);
  }
}

/**
 * Takes a response for a given activity from the Strava API and saves an image of the map to the
 * desired location
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

function strava_main() {

  // Club activites example
  // var endpoint = '/clubs/693906/activities'
  // var query_object = {}
  // var response = callStravaAPI(endpoint, {})
  // console.log(response)
  // // saveMapToFile(response, 'example.png')

  // Individual athlete example

  var endpoint = 'athlete/activities'
  var query_object = {}
  var response = callStravaAPI(endpoint, {})
  console.log(response)
}


