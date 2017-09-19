// Copyright 2017, Epiphany Solutions All Rights Reserved.
//
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
//     http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.

/**
* @name AdWords and Analytics Daily Performance Report
*
* @overview The AdWords Performance script pulls data from Google AdWords/Analytics andthen exports the data to a Google Sheet. https://docs.google.com/spreadsheets/d/1l8S00Wzg-T_auMZT-kOC8LKPiDcFnEEJMi3VZxQtWdE/edit#gid=933359712
*
* @author Nathan Jackson [nathan.jackson@epiphanysolutions.co.uk]
*
* @version 1.0
*
* @changelog
* - version 1.0
*   - Released initial version.
*/
var REPORTING_OPTIONS = {
  apiVersion: 'v201705'
};
SPREADSHEET = SpreadsheetApp.openById('1l8S00Wzg-T_auMZT-kOC8LKPiDcFnEEJMi3VZxQtWdE');
SPREADSHEET_CONFIG = SPREADSHEET.getSheetByName('Config');
SCRIPT_SETTINGS = SPREADSHEET_CONFIG.getRange("A3:C8").getValues();
SPREADSHEET_ADWORDS = SPREADSHEET.getSheetByName('AdWords Data');
SPREADSHEET_ANALYTICS = SPREADSHEET.getSheetByName('Analytics Data');
SPREADSHEET_LOG = SPREADSHEET.getSheetByName('Script Logs');
DATATIMEZONE = "GMT"

function main() {
  //If the Script Status is true, run the rest of the script

  if(SCRIPT_SETTINGS[0][2] === true) {
    //Get the to and from dates in the right formats. This also checks for if it's the script's first run.
    var adwordsDateTo = manipulateDateByDays(-1, "yyyyMMdd")
    var analyticsDateTo = manipulateDateByDays(-1, "yyyy-MM-dd")
    var adwordsDateFrom = manipulateDateByDays((SCRIPT_SETTINGS[1][2] === true) ? -SCRIPT_SETTINGS[5][2] : -1, "yyyyMMdd");
    var analyticsDateFrom = manipulateDateByDays((SCRIPT_SETTINGS[1][2] === true) ? -SCRIPT_SETTINGS[5][2] : -1, "yyyy-MM-dd");
    
    //Create the headings of the AdWords/Analytics sheets if it's the first script run
    if(SCRIPT_SETTINGS[1][2] === true) {
      var adwordsSheetHeaders = SCRIPT_SETTINGS[2][2].match(/SELECT (.+) FROM/)[1].split(', ');
      var adwordsSheetTopRow = SPREADSHEET_ADWORDS.getRange(1, 1, 1, adwordsSheetHeaders.length);
      adwordsSheetTopRow.setValues([adwordsSheetHeaders]);
      
      var analyticsQuery = JSON.parse(SCRIPT_SETTINGS[4][2]);
      var analyticsDimensions = analyticsQuery.arguments.dimensions.split(',');
      var analyticsMetrics = analyticsQuery.metrics.split(',');
      var analyticsSheetTopRow = SPREADSHEET_ANALYTICS.getRange(1, 1, 1, analyticsDimensions.length+analyticsMetrics.length);
      analyticsSheetTopRow.setValues([analyticsDimensions.concat(analyticsMetrics)]);
      
      //Set the value of first script run to 
      SPREADSHEET_CONFIG.getRange('B4').setValue('No');
    }
    
    //Check if the AdWords script query starts with a SELECT, which should help catch if the Google Sheet has a formula error
    if(SCRIPT_SETTINGS[2][2].substring(0,7) == 'SELECT ') {
      //Pull the AdWords report
      var adwordsData = [];
      var report = AdWordsApp.report(SCRIPT_SETTINGS[2][2] + ' DURING ' + adwordsDateFrom + ',' + adwordsDateTo, REPORTING_OPTIONS);
      var rows = report.rows();
      var adwordsSheetHeaders = SCRIPT_SETTINGS[2][2].match(/SELECT (.+) FROM/)[1].split(', ');
      
      //Loop through the rows
      while (rows.hasNext()) {
        var row = rows.next();
        //Create a temporary array for us to put the AdWords data (as the rows is an object of varying headings depending on the AWQL query)
        var tmpArray = [];
        
        //Fix the date, assuming it's always the first element in the array, as the AWQL query dictates
        var dateString = row["Date"].split('-');
        tmpArray.push(dateString[2] + "/" + dateString[1] + "/" + dateString[0]);
        
        //Push the rest of the values onto the temporary array
        for(var x = 1; x < adwordsSheetHeaders.length; x++) {
          tmpArray.push(row[adwordsSheetHeaders[x]]);
        }
        adwordsData.push(tmpArray);
      }
      Logger.log(adwordsData);
      //insert more rows so we can paste in the performance report in bulk
      SPREADSHEET_ADWORDS.insertRowsAfter(SPREADSHEET_ADWORDS.getMaxRows(), adwordsData.length);
      //And finally push it in bulk into the spreadsheet
      SPREADSHEET_ADWORDS.getRange(SPREADSHEET_ADWORDS.getLastRow()+1, 1, adwordsData.length, adwordsData[0].length).setValues(adwordsData);
    } else {
      addToLog('High','AdWords data not pulled as the query doesn\'t start with "SELECT "- '+SCRIPT_SETTINGS[2][2]);
    }

    //Check if there's a profile ID and main KPI selected
    if(SCRIPT_SETTINGS[3][2] > 0 && SCRIPT_SETTINGS[4][2].length > 0) {
      var analyticsData = []; 
      //Pull out the metrics in its own variable, and delete it from the options variable
      var gaJson = JSON.parse(SCRIPT_SETTINGS[4][2]);
      var options = gaJson.arguments
      var metrics = gaJson.metrics;
      
      //Request the report
      var analyticsData = Analytics.Data.Ga.get('ga:' + SCRIPT_SETTINGS[3][2], analyticsDateFrom, analyticsDateTo, metrics, options).getRows();
      Logger.log(analyticsData);
      
      //Check if there are any results
      if(analyticsData.length > 0) {
        //If so, loop and modify the date so it's usable
        for(var i=0; i < analyticsData.length; i++) {
          analyticsData[i][0] = analyticsData[i][0].substring(6,8) + "/" + analyticsData[i][0].substring(4,6) + "/" + analyticsData[i][0].substring(0,4);
        }
        
        SPREADSHEET_ANALYTICS.insertRowsAfter(SPREADSHEET_ANALYTICS.getMaxRows(), analyticsData.length);
        SPREADSHEET_ANALYTICS.getRange(SPREADSHEET_ANALYTICS.getLastRow()+1, 1, analyticsData.length, analyticsData[0].length).setValues(analyticsData);
      } else {
        //Add to the log that there's no data, as that's odd
        addToLog('Critical','Analytics data attempted to be pulled, but no results. Either a bad query, or there was no data for the previous day!');
      }
    } else { 
      addToLog('High','Analytics data not pulled as one or more of the required fields (in the spreadsheet) aren\'t filled in');
    }
  }
  else
  {
    addToLog('Low','AdWords script not run due to settings');
  }
}

/**
* Post a message to the script log in the Google sheet
*
* @param {String} severity Suggested values are "Notice", "Warning", "Error", "Critical"
* @param {String} message More details on the error
*
* @return {Boolean} True if function was passed required parameters, false if not.
*/
function addToLog(severity, message) {
  if(message.length > 0 && severity.length > 0) {
    SPREADSHEET_LOG.appendRow([new Date(),'Daily Performance Export',severity,message]);
    return true;
  }
  return false;
}

/**
* Get a date in the past of future in a specified format
*
* @param {Integer} dayOffset The number of days you want to go forwards to or back until
* @param {String} format The format of the date you want returning (e.g. yyyyMMdd)
*
* @return {String} A formatted date
*/
function manipulateDateByDays(dayOffset,format) {
  if(typeof dayOffset !== "undefined" && format.length > 0) {
    var x = new Date();
    x.setDate(x.getDate() + dayOffset);
    return Utilities.formatDate(x, DATATIMEZONE, format);
  }
  return false;
}