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

// TODO: Take the logic from here it put it into a function...

function strava_main()
{
  // var endpoint = 'activities/12832996323'
  // var query_object = {}
  // var response = callStravaAPI(endpoint, {})
  // console.log(response)
  var map = Maps.newStaticMap();
  var polyline = 'orwtGj}a`MG?GEOSOIY_@OMe@ISF_@BMBMASBY?OG_@g@KEa@[QWKGKS]Qa@e@i@e@o@c@EOIAWa@KGY[ECG?i@]IMYQKSMGAGUMQW?F\\RPRJHDHPJCCWMGKMMQKGM]_@}@k@EESOsAuAEA[SQWc@UPFEEECSBIEEGUSIMa@WCEGCCGQOY_@QIEG}@k@IOWQYOQMKQSOGAGIMGe@g@WMOOCIGACGYKk@i@e@]WYMGOQ_@WIKMGcAgAj@b@ZPWWQIOQUCYUCAGBKh@Sf@GTCDG@GCGM]Wi@k@c@[KOcAu@KKAGMIOSSKQMQYe@KUOMU_@UCIY[e@o@e@c@w@e@QQWe@]YUIi@o@OG@L^JMOGCE?m@o@EAIKTYJg@r@gBD[DCDQLYJa@h@sA\\kAASKUEEMAKKMEMWm@_@IMSOUWSIOMQGG?ECE?i@MSAYKE@e@OaAOa@KI?SM]AE@g@Qa@GKE[Cg@Mm@KWMMAa@KQA_@Qa@AEEk@KEEE@[KQAEGG?EEoAWYIXDFDN?v@NPFFJL`@?PCXI\\g@rAUr@KTGX[t@CBM`@KP}@pCUd@c@pAKJEZ]r@APQTEXWl@GVGDEXWv@ORSd@G\\KPEPOZ?HUPEX@LD@FFXd@\\TRZt@f@JPVLJJ@FTTHLTNHJLHFN\\RPXJJDBNPTHT\\NHX^r@j@LRNFDJv@j@Z\\VP?FLLTJN\\LFX^LFNPJTJFFRN~@`@j@j@d@DFb@Z\\`@LHNRLBVXXL^d@XL^d@\\RFLXRR\\d@Xf@h@JFJDd@`@NTD@^VBHPNRZLDDCF@HR`@\\Zb@VNVZZPXVHNJBT\\dAbAl@ZJLJDTZVNZ\\\\TNRNLBFPH`@d@RHFPh@h@FBDJPDJJTLRZRNPZRLX\\VPPXv@b@BHTNBFNJHPN@f@h@FBPPRJ`@h@\\PXTFNPHZX`@Rn@j@HDJLPJr@n@Gb@An@@Ld@l@ZVD@PCPGFUDKNBVTRHJGBYZk@v@{BHMBQ^y@hA_Df@eBNQDSHMZ{@HWf@qAjAgDXs@Pm@R]PIb@MRCLIZItAmAT?RCT@FCTBL?LGVENTT@LRPLTDFFFNJD'
  map.addPath(polyline)
  DriveApp.createFile(Utilities.newBlob(map.getMapImage(), 'image/png', 'map.png'));
}


