/********************************************************************************
 *
 *  Seznam Webmaster API Google Apps Script v 0.1 - Get urls data into Google Sheets
 *  Copyright (C) 2017  Marek Čech
 *
 *   This program is free software: you can redistribute it and/or modify
 *   it under the terms of the GNU General Public License as published by
 *   the Free Software Foundation, either version 3 of the License, or
 *   (at your option) any later version.
 * 
 *   This program is distributed in the hope that it will be useful,
 *   but WITHOUT ANY WARRANTY; without even the implied warranty of
 *   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 *   GNU General Public License for more details.
 * 
 *   You should have received a copy of the GNU General Public License
 *   along with this program.  If not, see <http://www.gnu.org/licenses/>.
 *     
 *   Contact us and leave feedback: 
 *   http://www.digitalniarchitekti.cz/
 *   https://www.facebook.com/digitalniarchitekti
 * 
 */

// Your API key you get on https://reporter.seznam.cz/wm/
var API_KEY = 'INSERT YOUR API';

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  // custom menu for user friendlines
  ui.createMenu('Seznam Webmasters')
      .addItem('History Data', 'seznamHistory' )
      .addSeparator()
      .addItem('Error Urls', 'seznamErrorUrls' )
      .addItem('Content Urls', 'seznamContentUrls' )
      .addItem('Index Urls', 'seznamIndexUrls' )
      .addItem('Redirect Urls', 'seznamRedirectUrls' )
      .addSeparator()
      .addItem('Reindex Selected', 'seznamReindex' )
      .addItem('Details about Selected', 'seznamDetails' )
  //    .addSubMenu(ui.createMenu('Sub-menu')
  //        .addItem('Second item', 'Test Submenu'))
      .addToUi();
}



/********************************************************************************
 * Call the API and get history data
 * 
 */

function seznamHistory() {
  
  // URL and params for the Seznam Webmaster API
  var root = 'https://reporter.seznam.cz/wm-api/';
  var endpoint = '/web?key='+ API_KEY;
  
  // parameters for url fetch
  var params = {
    'method': 'GET',
    'muteHttpExceptions': true,
    'headers': {
      'Authorization': 'key ' + API_KEY
    }
  };
  
  try {
    // call the Seznam Webmaster API
    var response = UrlFetchApp.fetch(root+endpoint, params);
    var data = response.getContentText();
    var json = JSON.parse(data);
    
    // get just history data
    var history = json['history'];
    
    // Log the campaignData array
    Logger.log(history);
    
    // blank array to hold the history data for Sheet
    var historyData = [];
  
    // Add the history data to the array
    for (var i = 0; i < history.length; i++) {
      
      // put the history data into array for Google Sheets
   
        historyData.push([
          history[i]["date"],
          history[i]["counts"]["downloaded"],
          history[i]["counts"]["error"],
          history[i]["counts"]["indexed"],
          history[i]["counts"]["redirected"],
        ]);
    
   }
    
    // Log the historyData array
    Logger.log(historyData);
    
    // select the history output sheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Index History');
    
    // calculate the number of rows and columns needed
    var numRows = historyData.length;
    var numCols = historyData[0].length;
    
    // output the numbers to the sheet
    sheet.getRange(4,1,numRows,numCols).setValues(historyData);
    
    // adds formulas to calculate change for error and indexed urls
    for (var i = 1; i < numRows; i++) {
      sheet.getRange(4+i,6).setFormulaR1C1('=iferror(R[0]C[-3]-R[-1]C[-3];"N/a")').setNumberFormat("0");
      sheet.getRange(4+i,7).setFormulaR1C1('=iferror(R[0]C[-3]-R[-1]C[-3];"N/a")').setNumberFormat("0");
    }
    
  }
  catch (error) {
    // deal with any errors
    Logger.log(error);
  };
  
}

/********************************************************************************
 * 
 * After selection run this from menu item and Reindex the pages
 * 
 */

function seznamReindex() {
  
  // URL and params for the Seznam Webmaster API
  var root = 'https://reporter.seznam.cz/wm-api/';
  var endpoint = '/web/document/reindex?key='+ API_KEY +'&url=';
  
  // parameters, POST method must be used
  var params = {
    'method': 'POST',
    'muteHttpExceptions': true,
    'headers': {
      'Authorization': 'key ' + API_KEY
    }
  };
  
  try {
    
  // find what is selected
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var data = sheet.getActiveRange().getValues();
    
  var date = Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yyyy")

    
  // go thorugh all selected URLs
    
  for (var i = 0; i < data.length; i++) {
    
    Logger.log(data[i]);
    var send = UrlFetchApp.fetch(root+endpoint+data[i], params);
    Logger.log(send);
    
    
  }
    // provide feedback for user
    sheet.getActiveRange().offset(0, 1, i).setValue("requested "+date);
  
  }
   catch (error) {
    // deal with any errors
    Logger.log(error);
  };
}

  
/********************************************************************************
 * 
 * After selection run this from menu item and get details for selected Urls
 * 
 */

function seznamDetails() {
  
  // URL and params for the Seznam Webmaster API
  var root = 'https://reporter.seznam.cz/wm-api/';
  var endpoint = '/web/document?key='+ API_KEY +'&url=';
  // parameters
  var params = {
    'method': 'GET',
    'muteHttpExceptions': true,
    'headers': {
      'Authorization': 'key ' + API_KEY
    }
  };
  
  try {
  // find what is selected 
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var data = sheet.getActiveRange().getValues();
    
  var date = Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yyyy")

  // every Url selected will be processed separately 
  
  for (var i = 0; i < data.length; i++) {
    
    Logger.log("Requested URL:" + data[i]);
    var response = UrlFetchApp.fetch(root+endpoint+data[i], params);
    Logger.log("API Response:" +response);
    
    var datajson = response.getContentText();
    var json = JSON.parse(datajson);
    
    // reset the array
    var detailsData = [];
    
    // Log the json array
    Logger.log("Parsed response:" +json);
    
   
      
    // put the data into array for Google Sheets
  
     detailsData.push([
          json["title"],
          json["url"],
          json["meta"]["desc"],
          json["meta"]["keywords"],
          json["meta"]["author"],
                  
        ]);
       
     // cycle trough all OG info and append them to the array
     for (var a = 0; a < json.openGraphData.length; a++) {
         
          detailsData[0].push(
            json.openGraphData[a]["name"],
            json.openGraphData[a]["content"]
                              );
    }
                      
    // Log the detailsData array
    Logger.log("Output after PUSH:" +detailsData);    
    
    // select the approriate sheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Url Details');
    
    // get length
    var numRows = detailsData.length;
    var numCols = detailsData[0].length;
    
    // output the info to the sheet, offset it for new line based on number of cycle
    sheet.getRange(4,1,numRows,numCols).offset(i, 0).setValues(detailsData);                   
   }
    
  // give feedback to user that everything is ok
  sheet.getActiveRange().offset(0, 1).setValue("details requsted "+date);
 
  
  }
   catch (error) {
    // deal with any errors
    Logger.log(error);
  };
}

  
/********************************************************************************
 * 
 * Retrives Error URL list and populates a Google Sheet
 * 
 */
function seznamErrorUrls() {
  
  // URL and params for the Seznam Webmaster Api
  var root = 'https://reporter.seznam.cz/wm-api/';
  var endpoint = '/web/documents?key='+ API_KEY;
  
  var params = {
    'method': 'GET',
    'muteHttpExceptions': true,
    'headers': {
      'Authorization': 'key ' + API_KEY
    }
  };
  try {
    // call the API
    var response = UrlFetchApp.fetch(root+endpoint, params);
    var data = response.getContentText();
    var json = JSON.parse(data);
    
    // get just error URLs
    var Urls = json.error['urls'];
    
    // separate the total count of urls
    var count = json.error['count'];
    
    // blank array to hold the data for Sheet
    var readyUrls = [];
    
    // Add the Urls data to the array, empty "" is there just for simple cleaning the reindex column
    var numbering = 0;
    Urls.forEach(function(el) {
     numbering++;
     readyUrls.push([numbering, el, ""]);
    });
    
    // Log the Urls array
    Logger.log(readyUrls);
    
    // select the output shee
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Error Urls');
    
    // calculate the number of rows and columns needed
    var numRows = readyUrls.length;
    var numCols = readyUrls[0].length;
    
    // output the Urls to the sheet
    sheet.getRange(4,1,numRows,numCols).setValues(readyUrls);
    sheet.getRange(1,4).setValue(count);
    
  }
  catch (error) {
    // deal with any errors
    Logger.log(error);
  };
  
}
  
/********************************************************************************
 * 
 * Retrives Content URL list and populates a Google Sheet
 * 
 */
  
function seznamContentUrls() {
  
   // URL and params for the Seznam Webmaster Api
  var root = 'https://reporter.seznam.cz/wm-api/';
  var endpoint = '/web/documents?key='+ API_KEY;
  
  var params = {
    'method': 'GET',
    'muteHttpExceptions': true,
    'headers': {
      'Authorization': 'key ' + API_KEY
    }
  };
  try {
    // call the API
    var response = UrlFetchApp.fetch(root+endpoint, params);
    var data = response.getContentText();
    var json = JSON.parse(data);
    
    // get just content URLs
    var Urls = json.content['urls'];
    
    // separate the total count of urls
    var count = json.content['count'];
    
    // blank array to hold the data for Sheet
    var readyUrls = [];
    
    // Add the Urls data to the array, empty "" is there just for simple cleaning the reindex column
    var numbering = 0;
    Urls.forEach(function(el) {
     numbering++;
     readyUrls.push([numbering, el, ""]);
    });
    
    // Log the Urls array
    Logger.log(readyUrls);
    
    // select the output sheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Content Urls');
    
    // calculate the number of rows and columns needed
    var numRows = readyUrls.length;
    var numCols = readyUrls[0].length;
    
    // output the Urls to the sheet
    sheet.getRange(4,1,numRows,numCols).setValues(readyUrls);
    sheet.getRange(1,4).setValue(count);
    
  }
  catch (error) {
    // deal with any errors
    Logger.log(error);
  };
  
}
  
/********************************************************************************
 * 
 * Retrives index URL list and populates a Google Sheet
 * 
 */

function seznamIndexUrls() {
  
   // URL and params for the Seznam Webmaster Api
  var root = 'https://reporter.seznam.cz/wm-api/';
  var endpoint = '/web/documents?key='+ API_KEY;
  
  var params = {
    'method': 'GET',
    'muteHttpExceptions': true,
    'headers': {
      'Authorization': 'key ' + API_KEY
    }
  };
  try {
    // call the API
    var response = UrlFetchApp.fetch(root+endpoint, params);
    var data = response.getContentText();
    var json = JSON.parse(data);
    
    // get just index URLs
    var Urls = json.index['urls'];
    
    // separate the total count of urls
    var count = json.index['count'];
    
    // blank array to hold the data for Sheet
    var readyUrls = [];
    
    // Add the Urls data to the array, empty "" is there just for simple cleaning the reindex column
    var numbering = 0;
    Urls.forEach(function(el) {
     numbering++;
     readyUrls.push([numbering, el, ""]);
    });
    
    // Log the Urls array
    Logger.log(readyUrls);
    
    // select the output shee
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Index Urls');
    
    // calculate the number of rows and columns needed
    var numRows = readyUrls.length;
    var numCols = readyUrls[0].length;
    
    // output the Urls to the sheet
    sheet.getRange(4,1,numRows,numCols).setValues(readyUrls);
    sheet.getRange(1,4).setValue(count);
    
  }
  catch (error) {
    // deal with any errors
    Logger.log(error);
  };
  
}
  
/********************************************************************************
 * 
 * Retrives redirect URL list and populates a Google Sheet
 * 
 */
  
function seznamRedirectUrls() {
  
   // URL and params for the Seznam Webmaster Api
  var root = 'https://reporter.seznam.cz/wm-api/';
  var endpoint = '/web/documents?key='+ API_KEY;
  
  var params = {
    'method': 'GET',
    'muteHttpExceptions': true,
    'headers': {
      'Authorization': 'key ' + API_KEY
    }
  };
  try {
    // call the API
    var response = UrlFetchApp.fetch(root+endpoint, params);
    var data = response.getContentText();
    var json = JSON.parse(data);
    
    // get just redirect URLs
    var Urls = json.redirect['urls'];
    
    // separate the total count of urls
    var count = json.redirect['count'];
    
    // blank array to hold the data for Sheet
    var readyUrls = [];
    
    // Add the Urls data to the array, empty "" is there just for simple cleaning the reindex column
    var numbering = 0;
    Urls.forEach(function(el) {
     numbering++;
     readyUrls.push([numbering, el, ""]);
    });
    
    // Log the Urls array
    Logger.log(readyUrls);
    
    // select the output shee
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Redirect Urls');
    
    // calculate the number of rows and columns needed
    var numRows = readyUrls.length;
    var numCols = readyUrls[0].length;
    
    // output the Urls to the sheet
    sheet.getRange(4,1,numRows,numCols).setValues(readyUrls);
    sheet.getRange(1,4).setValue(count);
    
  }
  catch (error) {
    // deal with any errors
    Logger.log(error);
  };
  
}


