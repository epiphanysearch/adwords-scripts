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
* @name AdWords Performance Anomaly Detector
*
* @overview Runs through campaign performance over a week of day average for the previous week and emails results. https://docs.google.com/spreadsheets/d/1l8S00Wzg-T_auMZT-kOC8LKPiDcFnEEJMi3VZxQtWdE/edit#gid=933359712
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
SCRIPT_SETTINGS = SPREADSHEET_CONFIG.getRange("E3:G10").getValues();
SPREADSHEET_LOG = SPREADSHEET.getSheetByName('Script Logs');
DATATIMEZONE = "GMT"

function main() {
  var weeksToDo = SCRIPT_SETTINGS[0][2];
  var weeksDone = 0;
  var campaignHistory = {};
  var alerts = [];
  
  //Loop through how many weeks we want to get an average of. For more alerts you'd typically use less weeks.
  for(var i = 1; i < (weeksToDo+1); i++)
  {
    //Pull an AdWords reports for all campaigns wit impressions
    var dateRange = manipulateDateByDays(-((i*7)+1),'yyyyMMdd');
    var report = AdWordsApp.report('SELECT CampaignName, Impressions, Clicks, Cost, Conversions FROM CAMPAIGN_PERFORMANCE_REPORT WHERE Impressions > 0 DURING '+dateRange+','+dateRange,REPORTING_OPTIONS);
    
    //Loop through the report rows
    var rows = report.rows();
    while (rows.hasNext()) {
      var row = rows.next();
      //Add the campaign to our history object, if we haven't before
      if(typeof campaignHistory[row['CampaignName']] === 'undefined') campaignHistory[row['CampaignName']] = {'impressions':0,'clicks':0,'cost':0,'conversions':0,'weeksOfData':0,'averages':{},'yesterday':{},'flags':[],'newCampaign':false};
      
      //Add the results of the current date to the history
      campaignHistory[row['CampaignName']].impressions = campaignHistory[row['CampaignName']].impressions + parseInt(row['Impressions'],10);
      campaignHistory[row['CampaignName']].clicks = campaignHistory[row['CampaignName']].clicks + parseInt(row['Clicks'],10);
      campaignHistory[row['CampaignName']].cost = campaignHistory[row['CampaignName']].cost + parseFloat(row['Cost']);
      campaignHistory[row['CampaignName']].conversions = campaignHistory[row['CampaignName']].conversions + parseFloat(row['Conversions']);
      
      //If the campaign had impressions on this date, count it.
      if(parseInt(row["Impressions"]) > 0) campaignHistory[row['CampaignName']].weeksOfData++;
    }
  }
  
  //Create a report for yesterday's figures of all paused and enabled campaigns
  var dateRange = manipulateDateByDays(-1,'yyyyMMdd');
  var report = AdWordsApp.report('SELECT CampaignName, Impressions, Clicks, Cost, Conversions FROM CAMPAIGN_PERFORMANCE_REPORT WHERE CampaignStatus In [PAUSED,ENABLED] DURING '+dateRange+','+dateRange,REPORTING_OPTIONS);
  var rows = report.rows();
  while (rows.hasNext()) {
    var row = rows.next();
    
    //If the campaign history doesn't exist for this campaign, add it and mark it as a new campaign. Please note, it may not always be a new campaign, but campaigns that had activity yesterday but not in the past weeks
    if(typeof campaignHistory[row['CampaignName']] === 'undefined') campaignHistory[row['CampaignName']] = {'impressions':0,'clicks':0,'cost':0,'conversions':0,'weeksOfData':0,'averages':{},'yesterday':{},'flags':[],'newCampaign':true};
    
    //Assign yesterday's data to the campaign object
    campaignHistory[row['CampaignName']].yesterday.impressions = parseInt(row['Impressions'],10);
    campaignHistory[row['CampaignName']].yesterday.clicks = parseInt(row['Clicks'],10);
    campaignHistory[row['CampaignName']].yesterday.cost = parseFloat(row['Cost']);
    campaignHistory[row['CampaignName']].yesterday.conversions = parseFloat(row['Conversions']);
    campaignHistory[row['CampaignName']].yesterday.ctr = (row['Clicks'] / row['Impressions']).toFixed(2);
    campaignHistory[row['CampaignName']].yesterday.cpc = (row['Cost'] / row['Clicks']).toFixed(2);
    campaignHistory[row['CampaignName']].yesterday.convRate = (row['Conversions'] / row['Clicks']).toFixed(2);
    
    //Average the performance stats
    campaignHistory[row['CampaignName']].averages = {'impressions':0,'clicks':0,'cost':0,'conversions':0,'cpc':0.00,'convRate':0.00,'ctr':0.00}
    campaignHistory[row['CampaignName']].averages.impressions = campaignHistory[row['CampaignName']].impressions / campaignHistory[row['CampaignName']].weeksOfData;
    campaignHistory[row['CampaignName']].averages.clicks = campaignHistory[row['CampaignName']].clicks / campaignHistory[row['CampaignName']].weeksOfData;
    campaignHistory[row['CampaignName']].averages.cost = campaignHistory[row['CampaignName']].cost / campaignHistory[row['CampaignName']].weeksOfData;
    campaignHistory[row['CampaignName']].averages.conversions = row['Conversions'] / campaignHistory[row['CampaignName']].weeksOfData;
    campaignHistory[row['CampaignName']].averages.cpc = campaignHistory[row['CampaignName']].cost / campaignHistory[row['CampaignName']].clicks;
    campaignHistory[row['CampaignName']].averages.ctr = campaignHistory[row['CampaignName']].averages.clicks / campaignHistory[row['CampaignName']].impressions;
    campaignHistory[row['CampaignName']].averages.convRate = campaignHistory[row['CampaignName']].averages.conversions / campaignHistory[row['CampaignName']].clicks;
    
    //Compare the performance stats
    campaignHistory[row['CampaignName']].change = {'impressions':0,'ctr':0.00,'cost':0,'cpc':0.00,'convRate':0.00}
    campaignHistory[row['CampaignName']].change.impressions = (row['Impressions'] - campaignHistory[row['CampaignName']].averages.impressions) / campaignHistory[row['CampaignName']].averages.impressions;
    campaignHistory[row['CampaignName']].change.ctr = (campaignHistory[row['CampaignName']].yesterday.ctr - campaignHistory[row['CampaignName']].averages.ctr) / campaignHistory[row['CampaignName']].averages.ctr;
    campaignHistory[row['CampaignName']].change.cost = (row['Cost'] - campaignHistory[row['CampaignName']].averages.cost) / campaignHistory[row['CampaignName']].averages.cost;
    campaignHistory[row['CampaignName']].change.cpc = (campaignHistory[row['CampaignName']].yesterday.cpc - campaignHistory[row['CampaignName']].averages.cpc) / campaignHistory[row['CampaignName']].averages.cpc;
    campaignHistory[row['CampaignName']].change.convRate =  (campaignHistory[row['CampaignName']].yesterday.convRate - campaignHistory[row['CampaignName']].averages.convRate) / campaignHistory[row['CampaignName']].averages.convRate;
    
    //Check whether the performance stats are different enough to alert. If one of the stats is out of the limits set in the spreadsheet, flag it
    if(campaignHistory[row['CampaignName']].change.impressions > SCRIPT_SETTINGS[3][2] || campaignHistory[row['CampaignName']].change.impressions < -SCRIPT_SETTINGS[3][2]) campaignHistory[row['CampaignName']].flags.push('Impression');
    if(campaignHistory[row['CampaignName']].change.impressions > SCRIPT_SETTINGS[4][2] || campaignHistory[row['CampaignName']].change.impressions < -SCRIPT_SETTINGS[4][2]) campaignHistory[row['CampaignName']].flags.push('CTR');
    if(campaignHistory[row['CampaignName']].change.cost > SCRIPT_SETTINGS[5][2] || campaignHistory[row['CampaignName']].change.cost < -SCRIPT_SETTINGS[5][2]) campaignHistory[row['CampaignName']].flags.push('Cost');
    if(campaignHistory[row['CampaignName']].change.cpc > SCRIPT_SETTINGS[6][2] || campaignHistory[row['CampaignName']].change.cpc < -SCRIPT_SETTINGS[6][2]) campaignHistory[row['CampaignName']].flags.push('Avg CPC');
    if(campaignHistory[row['CampaignName']].change.convRate > SCRIPT_SETTINGS[7][2] || campaignHistory[row['CampaignName']].change.convRate < -SCRIPT_SETTINGS[7][2]) campaignHistory[row['CampaignName']].flags.push('Conv. Rate');
    
    //Check if the campaign should be added to the email alert (i.e. has a flag)
    if(campaignHistory[row['CampaignName']].flags.length > 0) alerts.push([
      row['CampaignName'], campaignHistory[row['CampaignName']].flags.toString(),
      campaignHistory[row['CampaignName']].yesterday.impressions, campaignHistory[row['CampaignName']].yesterday.clicks, campaignHistory[row['CampaignName']].yesterday.ctr, campaignHistory[row['CampaignName']].yesterday.cpc, campaignHistory[row['CampaignName']].yesterday.cost, campaignHistory[row['CampaignName']].yesterday.conversions, campaignHistory[row['CampaignName']].yesterday.convRate,
      campaignHistory[row['CampaignName']].averages.impressions, campaignHistory[row['CampaignName']].averages.clicks, campaignHistory[row['CampaignName']].averages.ctr, campaignHistory[row['CampaignName']].averages.cpc, campaignHistory[row['CampaignName']].averages.cost, campaignHistory[row['CampaignName']].averages.conversions, campaignHistory[row['CampaignName']].averages.convRate,
      campaignHistory[row['CampaignName']].change.impressions, campaignHistory[row['CampaignName']].change.ctr, campaignHistory[row['CampaignName']].change.cpc, campaignHistory[row['CampaignName']].change.cost, campaignHistory[row['CampaignName']].change.convRate
    ]);
  }
  SPREADSHEET_LOG.appendRow([new Date(),'Anomaly Detector','Notice','Script ran - '+ alerts.length +' campaigns triggered alerts']);
  //Check if we need to alert anyone
  if(alerts.length > 0)
  {
    if(SCRIPT_SETTINGS[1][2] == true) {
      //Sort the array on the average cost high to low
      alerts.sort(function(a,b) {
        return b[13] - a[13];
      })
      
      //Write the first lines of the email
      var html_email_text = 'Hi<br><br>There are ' + alerts.length + ' campaigns which have triggered a performance alert.<br><br><table border="1" cellpadding="5" cellspacing="0" height="100%" width="100%"><tr><th bgcolor="#f89728" color="#ffffff" colspan="2"></th><th bgcolor="#f89728" color="#ffffff" colspan="7">Yesterday</th><th bgcolor="#f89728" color="#ffffff" colspan="7">' + weeksToDo + ' Week Average</th><th bgcolor="#f89728" color="#ffffff" colspan="5">% Change From Average</th></tr><tr><th bgcolor="#f89728" color="#ffffff">Campaign</th><th bgcolor="#f89728" color="#ffffff">Alerts</th><th bgcolor="#f89728" color="#ffffff">Impressions</th><th bgcolor="#f89728" color="#ffffff">Clicks</th><th bgcolor="#f89728" color="#ffffff">CTR</th><th bgcolor="#f89728" color="#ffffff">CPC</th><th bgcolor="#f89728" color="#ffffff">Cost</th><th bgcolor="#f89728" color="#ffffff">Conversions</th><th bgcolor="#f89728" color="#ffffff">Conv. Rate</th><th bgcolor="#f89728" color="#ffffff">Impressions</th><th bgcolor="#f89728" color="#ffffff">Clicks</th><th bgcolor="#f89728" color="#ffffff">CTR</th><th bgcolor="#f89728" color="#ffffff">CPC</th><th bgcolor="#f89728" color="#ffffff">Cost</th><th bgcolor="#f89728" color="#ffffff">Conversions</th><th bgcolor="#f89728" color="#ffffff">Conv. Rate</th><th bgcolor="#f89728" color="#ffffff">Impressions</th><th bgcolor="#f89728" color="#ffffff">CTR</th><th bgcolor="#f89728" color="#ffffff">CPC</th><th bgcolor="#f89728" color="#ffffff">Cost</th><th bgcolor="#f89728" color="#ffffff">Conv. Rate</th></tr>';
      var email_text = "Hi\n\nThere are " + alerts.length + " campaigns which have triggered a performance alert";
      //Loop through the rows of alerts
      for(var j = 0; j < alerts.length; j++) {
        //For the plain email text, just mention what flags there are
        email_text += alerts[j][0] + " has differences in " + alerts[j][1] + "\n";
        
        //But for the html email content, create a row and loop through the columns
        html_email_text += '<tr>';
        for(var k = 0; k < alerts[j].length; k++) {
          //Format the numbers & text based on the column
          var valueWithFormatting = ""
          if([4,8,11,15,16,17,18,19,20].indexOf(k) != -1) valueWithFormatting = (alerts[j][k]*100).toFixed(2) + '%'; //Format as percentage
          else if([5,6,12,13].indexOf(k) > -1) valueWithFormatting = "£" + parseFloat(alerts[j][k]).toFixed(2) //Format as currency
          else if([0,1].indexOf(k) > -1) valueWithFormatting = alerts[j][k];
          else valueWithFormatting = (alerts[j][k]).toFixed(0); //Leave as normal format
          
          if(valueWithFormatting !== "") html_email_text += '<td>' + valueWithFormatting + '</td>';
        }
        html_email_text += '</tr>';
      }
      //Remove any NaN (not a number) that occurs on math errors
      html_email_text = html_email_text.replace(/(£?NaN%?)/gm,"-");
      
      //Finish the email
      email_text += "\n\nThanks\nEpiphany AdWords Script";
      html_email_text += '</table><br><br>Thanks<br>Epiphany AdWords Script';
      
      //And send the email to the email addresses in the Google sheet
      MailApp.sendEmail(SCRIPT_SETTINGS[2][2],"Epiphany AdWords Scripts - PPC Performance Anomaly",email_text, {htmlBody:html_email_text});
    } else {
      SPREADSHEET_LOG.appendRow([new Date(),'Anomaly Detector','Notice','Alert Emails turned off, but there were '+ alerts.length +' campaigns that triggered alerts']); 
    }
  }
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
