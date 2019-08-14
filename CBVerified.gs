// Sheet vars
var ui = SpreadsheetApp.getUi();
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getActiveSheet();
var rangeData = sheet.getDataRange();
var lastColumn = rangeData.getLastColumn();
var lastRow = rangeData.getLastRow();
var searchRange = sheet.getRange(1,10, lastRow, 1);
var domainCheckValues = searchRange.getValues()
var lastRowNonEmpty = domainCheckValues.filter(String).length;
var domainQueried = "Domain Queried";

// Crunchbase API key
var USER_KEY = 'YOUR_USER_KEY_HERE';

// Menu
function onOpen() {
  ui.createMenu('Domain Checker')
  .addItem('Run check', 'domainLoop')
  .addItem('Are we done yet?', 'areWeDoneYet')
  .addToUi();
};

// Google script timeout function
function isTimeUp_(start) {
  var now = new Date();
  return now.getTime() - start.getTime() > 270000; // 4m 30s
}

function domainLoop() {
  // Set start time
  var start = new Date();

  // Get array of values in the search Range
  var rangeValues = searchRange.getValues();

  // Get array of Domain checked status values
  var domainQueriedRange = sheet.getRange(1,lastColumn, lastRowNonEmpty, 1);
  var domainQueriedValues = domainQueriedRange.getValues();

  // Loop through array
  for ( i = 2; i < lastRowNonEmpty + 1; i++){
    // Check if time is up
    if (isTimeUp_(start)) {
      Logger.log("Time up");
      break;
    };
    // Run loop
    var domain = rangeValues[i-1];
    var query = [i, domain];

    var domainChecked = String(domainQueriedValues[i-1]);

    if (domainChecked !== domainQueried) {

      getCrunchbaseOrgs(query);

    };
  };
  areWeDoneYet();
};

// Check last record and run loop again if not finished
function areWeDoneYet() {
  var domainQueriedRange = sheet.getRange(2,lastColumn, lastRowNonEmpty, 1);
  var domainQueriedValues = domainQueriedRange.getValues();
  var loopBroken = false;
  for ( i = 0; i < lastRowNonEmpty; i++) {
   var domainChecked = String(domainQueriedValues[i]);
    if (domainChecked !== domainQueried) {
     domainLoop();
     loopBroken = true;
     break
    }
  }
  if (!loopBroken) {
   ui.alert("All domains queried");
  }
};

// query to call Crunchbase API
function getCrunchbaseData(url,query) {

  try {
    var response = UrlFetchApp.fetch(url);
    var responseCode = response.getResponseCode();
    var responseData = response.getContentText();
    var json = JSON.parse(responseData);
    return [responseCode, json];
  }
  catch (e) {
    Logger.log(e);
    return ["Error:", e];
  }
}

// function to retrive organizations data
function getCrunchbaseOrgs(query) {

  // URL and params for the Crunchbase API
  var url = 'https://api.crunchbase.com/v3.1/odm-organizations?domain_name=' + encodeURI(query[1]) + '&user_key=' + USER_KEY;

  var json = getCrunchbaseData(url,query[1]);

  if (json[1] === "Error:") {
    // deal with error with fetch operation
    sheet.getRange(query[0],12,1,2).clearContent();
    sheet.getRange(query[0],12,1,1).setValues([json]);
  }
  else {
    if (json[0] !== 200) {
      // deal with error from api
      sheet.getRange(i,12,1,2).clearContent();
      sheet.getRange(i,12,1,2).setValues([["Error, server returned code:",json[0]]]);
    }
    else {
      // correct data comes back, check array exists
      if (!json[1].data.items || !json[1].data.items.length) {
       // no results from api
       sheet.getRange(i,12,1,2).clearContent();
       sheet.getRange(i,12,1,1).setValues([["No results"]]);
      }

      else {
      // present result
      var data = json[1].data.items[0].properties

      // parse into array for Google Sheet
      var outputData = [[
        "Yes",
        data.name,
        data.country_code
      ]];

      // clear any old data
      sheet.getRange(query[0],11,1,3).clearContent();

      // insert new data
      sheet.getRange(query[0],11,1,3).setValues(outputData);
      }
    }
  }
  sheet.getRange(i, lastColumn).setValue(domainQueried); // Update the last column with "Queried"
  SpreadsheetApp.flush(); // Make sure the last cell is updated right away
};
