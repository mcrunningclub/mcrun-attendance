const STRAVA_BASE_URL = 'https://www.strava.com/api/v3/'

/**
 * Function mapping an Object which maps params to values to a query string.
 * Ex: {"param1": val1, "param2": val2} -> "?param1=val1&param2=val2"
 */

function query_object_to_string(query_object){
  var param_value_list = Object.entries(query_object);
  var param_strings = param_value_list.map(([param, value]) => `${param}=${value}`);
  var query_string = param_strings.join('&');
  return '?' + query_string;
}

// call the Strava API
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

function strava_main()
{
  var endpoint = 'clubs/693906/activities'
  var query_object = {"per_page":1};
  var response = callStravaAPI(endpoint, query_object)
  console.log(response)
}


