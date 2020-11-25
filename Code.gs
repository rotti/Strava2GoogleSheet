// https://www.benlcollins.com/spreadsheets/strava-api-with-google-sheets/
// custom menu
function onOpen() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('Strava App')
    .addItem('Get data', 'getStravaActivityData')
    .addToUi();
}

// Get athlete activity data
function getStravaActivityData() {
  
  // get the sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Sheet1');
  
  // get the last activity time and convert to epoch
  var lastRow = sheet.getLastRow();
  var dateColumn = 5; // activity.start_date is inside the fifth column
  var lastActivityDate = SpreadsheetApp.getActiveSheet().getRange(lastRow, dateColumn).getValue();
  var lastActivityEpoch = new Date(lastActivityDate).getTime();
  var lastActivityName = SpreadsheetApp.getActiveSheet().getRange(lastRow, 2).getValue(); // activity.name is in the 2nd column
  
  //Logger.log('Last Activity Date:', lastActivityDate); 
  //Logger.log('Last Activity Epoch:', lastActivityEpoch);
  
  var collectAfter = "";
  // check if sheet has activities (more than the header) and get position to search for new activities afterwards if so
  // set your hereStartsData to last line number with headers etc.
  var hereStartsData = 2;
  if (lastRow > hereStartsData) {
  //  activities collected. start at last activity
    collectAfter = lastActivityEpoch;
    collectAfter = (collectAfter-(collectAfter%1000))/1000; //remove last 3 digits from epoch for StravaAPI
    Logger.log('Found data in sheet. Last activity is from:', lastActivityDate, 'with name:', lastActivityName );
  } else {
  // empty sheet
    Logger.log('did not found data in sheet. Collect all activities from 2010-01-01 on and create from scratch.');
    collectAfter = "1262304001"; // 1262304001 =  Friday, January 1, 2010 0:00:01                     
  }
      
  // call the Strava API to retrieve data
  var page_number = 0;
  while (true) {
    page_number = page_number + 1;  
    var data = callStravaAPI(collectAfter, page_number);
    
    if (data === undefined || data.length == 0) {
      Logger.log("No further activities to collect.");
      break;
    }
  
    // empty array to hold activity data
    var stravaData = [];
    
    // loop over activity data and add to stravaData array for Sheet
    data.forEach(function(activity) {
      var arr = [];
      arr.push(
        activity.id,
        activity.name,
        activity.type,
        activity.distance,
        activity.start_date, 
        // dont add stuff before
        new Date(activity.start_date).getFullYear(),
        new Date(activity.start_date).toLocaleDateString('de-DE', { month: 'long' }),
        new Date(activity.start_date).toLocaleDateString('de-DE', { weekday: 'long' }),
        activity.total_elevation_gain,
        activity.moving_time,
        activity.elapsed_time,
        activity.elapsed_time - activity.moving_time,
        activity.average_speed,
        activity.gear_id,
        activity.achievement_count,
        activity.kudos_count,
        activity.athlete_count,
        activity.suffer_score,
        activity.pr_count,
        "https://strava.com/activities/" + activity.id
      );
      stravaData.push(arr);
    });
  
    // paste the values into the Sheet
    sheet.getRange(sheet.getLastRow() + 1, 1, stravaData.length, stravaData[0].length).setValues(stravaData);
  }    
}

// call the Strava API
function callStravaAPI(collectAfter, page_number) {
  
  // set up the service
  var service = getStravaService();
  
  if (service.hasAccess()) {
    //while (true) {
      Logger.log('App has access.'); 
      
      var endpoint = 'https://www.strava.com/api/v3/athlete/activities';
      
      //page_number = 1;
      //var params = '?after=' + collectAfter + '&per_page=5&page=' + page_number;
      var params = '?after=' + collectAfter + '&per_page=200&page=' + page_number;
      //Logger.log('Params:', params);
            
      var headers = {
        Authorization: 'Bearer ' + service.getAccessToken()
      };
    
      var options = {
        headers: headers,
        method : 'GET',
        muteHttpExceptions: true
      };
    
      var response = JSON.parse(UrlFetchApp.fetch(endpoint + params, options));  
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
