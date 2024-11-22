const STRAVA_BASE_URL = 'https://www.strava.com/api/v3/'

/**

 */

/**
 * Maps an Object containing param, value pairs to a query string.
 * Ex: {"param1": val1, "param2": val2} -> "?param1=val1&param2=val2"
 * 
 * @author [Jikael Gagnon](<jikael.gagnon@mail.mcgill.ca>)
 * @date  Nov 7, 2024
 * @update  Nov 7, 2024
 */

function query_object_to_string(query_object){
  if (Object.keys({}).length === 0)
  {
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
      method : 'GET',
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

function polyline_to_map(api_response, filename)
{
  var polyline = api_response['map']['polyline']
  var map = Maps.newStaticMap();
  map.addPath(polyline)
  DriveApp.createFile(Utilities.newBlob(map.getMapImage(), 'image/png', filename));
}

function getLatestRunData() {
  const sheet = ATTENDANCE_SHEET;
  const lastRow = sheet.getLastRow();
  var time = sheet.getRange(lastRow, TIMESTAMP_COL).getValue();
  var headrunners = sheet.getRange(lastRow, HEADRUNNERS).getValue();
  console.log(time);
  console.log(headrunners.split(' , '));
}

function strava_main()
{

  // Club activites example
  // var endpoint = '/clubs/693906/activities'
  // var query_object = {}
  // var response = callStravaAPI(endpoint, {})
  // console.log(response)
  // // polyline_to_map(response, 'example.png')

  // Individual athlete example

  var endpoint = '/activities/12832996323'
  var query_object = {}
  var response = callStravaAPI(endpoint, {})
  console.log(response)
  polyline_to_map(response, 'example.png')
}


