/*************************************************************************************************************************************
                                                           ACTION ITEM LOG APP
                                                         Google App Script (GAS)
                                                        Created by Scott Lawrence
                                             Subject Matter Expert - Oracle Project Finance
                                                   Contact: scott.lawrence@aholdusa.com
 
                                                   Scheduling GAS - time trigger setup  
                                                     Function - sendMasterActionEmail
                                 Timing: Weekday time triggers run the function Monday- Friday at 4pm - 5pm.
                                                       Function - sendUpdateRequest
                                 Timing: Weekday time triggers run the function Monday - Friday at noon - 1pm.
 
 
 ======================================================================================================================================
 changelog - 07/29/2012 - Version 3.0 - Tasks Completed: Sort and group requested action items for single email, Forms - update form - and writing too master action item updates.
             Action Item Log app can now collect action items that have not been updated in 2 or 3 business weeks and send email requests for updates to be submitted by filling out
             form that reports to a new sheet in the workbook.  

             06/22/2012 - Version 3.0 beta TBD release - Form used to receive updates.
 
             06/07/2012 - Version 3.0 - Alpha NOT RELEASED (Commented out)- Added feature to send emails requesting updates for action items 
             that have not been update in last 10 business days or 15 business days.  Logic for needing update is within spreadsheet, 
             so accessed as a property of the object.
             
             07/19/2012 - Version 2.2 - CSS styling added to action item email, framework code for displaying action.
              
             06/07/2012 - Version 2.1 - Refactored code and added feature to update approved reports.

             05/31/2012 - Version 2.0 - Sendemail function now generates a summary of the daily action item activity including new,updates,
             and closed action items. 

             05/23/2012 - Version 1.0 - This script is designed to send an email w/attachment of the spreadsheet the script is embeded 
             into.  Resource triggers are used to schedule the deliverly of the email.  The email is design as noreply html message. 
 
 =======================================================================================================================================
 tasks for next release
   01.) refactor code 
   02.) reduce hardcoding - emails, spreadsheet ssID, message
   03.) make the code more OOP, where possible.
   04.) comment code
   05.) programatically schedule delivery of email, trigger occurs all different times 
   06.) If 5, then create delivery calendar with method that boolean return that could be used with logic to send today or not
   07.) Currently used depercated methods  -- getEndColumn, replaced with ...
   08.) 
   09.) Personalize Master action item email to display assigned action items in table
   10.)   
   11.) Provide Statistics of action items  (nmbOfNewAIs,nmbOfClosedAIs,break-down of modules,Owner stats)
   12.) Personilze the emails of owners, 

========================================================================================================================================
Details for each version release
  
  version 3.0 7/29/12
  sl notes:   Writing to spreadsheet:  https://developers.google.com/apps-script/articles/writing_spreadsheet_data
              Spreadsheets and Forms:  https://developers.google.com/apps-script/articles/expense_report_approval
           Searching array for value:  https://developer.mozilla.org/en/JavaScript/Reference/Global_Objects/Array/indexOf
                     CSS Font-family:  http://www.typechart.com/
Dymanically create object properties:  http://www.sitepoint.com/forums/showthread.php?571957-quot-dynamic-quot-i-variable-names-in-javascript
                               Array:  https://developer.mozilla.org/en/JavaScript/Reference/Global_Objects/Array
                              splice:  https://developer.mozilla.org/en/JavaScript/Reference/Global_Objects/Array/splice
                                       http://www.w3schools.com/jsref/jsref_splice.asp
  
  version 3.0b -  06/22/12
  sl notes:  Using Continue looping method:  http://www.w3schools.com/js/js_break.asp
                         Finding Undefined:  http://stackoverflow.com/questions/27509/detecting-an-undefined-object-property-in-javascript
                         mailto parameters:  http://www.echoecho.com/htmllinks11.htm
            
  version 3.0a -  06/07/12 - 
  sl notes:  Just framework of code.
  
  version 2.2  -  07/19/12 
  sl notes: CSS styling:  http://border-radius.com/ 
                          http://www.colorpicker.com/
                          http://sixrevisions.com/web_design/creating-html-emails/
             HTML Table:  http://coding.smashingmagazine.com/2008/08/13/top-10-css-table-designs/
                          http://msdn.microsoft.com/en-us/library/ms532998(v=vs.85).aspx
  version 2.1  -  06/07/12 - No references
  
  version 2.0  -  05/31/12
  sl notes:  Reading data from spreadsheet: https://developers.google.com/apps-script/articles/reading_spreadsheet_data
                                            functions getRowsData....
                  javascript array methods: http://www.bennadel.com/blog/1796-Javascript-Array-Methods-Unshift-Shift-Push-And-Pop-.htm
                                            http://www.w3schools.com/jsref/jsref_push.asp
 version 1.0 -   
 sl notes: app script template: http://productforums.google.com/forum/#!category-topic/apps-script/services/ZhtBhCrg3u4
     javascript date resources: http://www.webreference.com/js/scripts/basic_date/index.html
                                http://www.w3schools.com/js/js_obj_date.asp
              string resources: http://www.developfortheweb.com/2009/03/multi-line-strings-in-javascript/
                                http://joemaller.com/js-mailer.shtml
         html message resource: markdown http://daringfireball.net/projects/markdown/dingus
               google api docs: https://developers.google.com/apps-script/class_spreadsheet#getId
                                https://developers.google.com/apps-script/articles/sending_emails
                                https://developers.google.com/apps-script/class_mailapp
                                https://developers.google.com/apps-script/service_spreadsheet
            codetemplate notes: - sending email
                                "This is a good tutorial:
                                http://code.google.com/googleapps/appsscript/articles/twitter_tutorial.html
                                The example code above was written based on other forum topics:
                                http://www.google.com/support/forum/p/apps-script/thread?tid=780918cacbcb92a5&hl=en
                                and
                                http://www.google.com/support/forum/p/apps-script/thread?tid=45dfe4b5213bb2a3&hl=en
                                Another useful post:
                                http://www.google.com/support/forum/p/apps-script/thread?tid=23a0e4a8ab42f104&hl=en
                                Yes, it's possible.  It is important to build the export URL for your spreadsheet:
                                You can put the exportUrl in a cell, in a UI App, in an e-mail, etc.
                                For other formats for your document read the Google Documents List Data API v3.0(Labs):
                                http://code.google.com/apis/documents/docs/3.0/developers_guide_protocol.html#DownloadingSpreadsheets"

**************************************************************************************************************************************/




/*
GGG              l       t        t                 l   ff                 t                  
G                 l       t        t          ii     l   f                  t  ii              
G  GG ooo ooo ggg l eee  ttt u  u ttt ooo rrr     aa l  fff u  u nnn   ccc ttt    ooo nnn   ss 
G   G o o o o g g l e e   t  u  u  t  o o r   ii a a l   f  u  u n  n c     t  ii o o n  n  s  
 GGG  ooo ooo ggg l ee    tt  uuu  tt ooo r   ii aaa l   f   uuu n  n  ccc  tt ii ooo n  n ss  
                g                                                                              
              ggg                         

*/

// getRowsData iterates row by row in the input range and returns an array of objects.
// Each object contains all the data for a given row, indexed by its normalized column name.
// Arguments:
//   - sheet: the sheet object that contains the data to be processed
//   - range: the exact range of cells where the data is stored
//   - columnHeadersRowIndex: specifies the row number where the column names are stored.
//       This argument is optional and it defaults to the row immediately above range; 
// Returns an Array of objects.
function getRowsData(sheet, range, columnHeadersRowIndex) {
  columnHeadersRowIndex = columnHeadersRowIndex || range.getRowIndex() - 1;
  var numColumns = range.getEndColumn() - range.getColumn() + 1;
  var headersRange = sheet.getRange(columnHeadersRowIndex, range.getColumn(), 1, numColumns);
  var headers = headersRange.getValues()[0];
  return getObjects(range.getValues(), normalizeHeaders(headers));
}

// For every row of data in data, generates an object that contains the data. Names of
// object fields are defined in keys.
// Arguments:
//   - data: JavaScript 2d array
//   - keys: Array of Strings that define the property names for the objects to create
function getObjects(data, keys) {
  var objects = [];
  for (var i = 0; i < data.length; ++i) {
    var object = {};
    var hasData = false;
    for (var j = 0; j < data[i].length; ++j) {
      var cellData = data[i][j];
      if (isCellEmpty(cellData)) {
        continue;
      }
      object[keys[j]] = cellData;
      hasData = true;
    }
    if (hasData) {
      objects.push(object);
    }
  }
  return objects;
}

// Returns an Array of normalized Strings.
// Arguments:
//   - headers: Array of Strings to normalize
function normalizeHeaders(headers) {
  var keys = [];
  for (var i = 0; i < headers.length; ++i) {
    var key = normalizeHeader(headers[i]);
    if (key.length > 0) {
      keys.push(key);
    }
  }
  return keys;
}

// Normalizes a string, by removing all alphanumeric characters and using mixed case
// to separate words. The output will always start with a lower case letter.
// This function is designed to produce JavaScript object property names.
// Arguments:
//   - header: string to normalize
// Examples:
//   "First Name" -> "firstName"
//   "Market Cap (millions) -> "marketCapMillions
//   "1 number at the beginning is ignored" -> "numberAtTheBeginningIsIgnored"
function normalizeHeader(header) {
  var key = "";
  var upperCase = false;
  for (var i = 0; i < header.length; ++i) {
    var letter = header[i];
    if (letter == " " && key.length > 0) {
      upperCase = true;
      continue;
    }
    if (!isAlnum(letter)) {
      continue;
    }
    if (key.length == 0 && isDigit(letter)) {
      continue; // first character must be a letter
    }
    if (upperCase) {
      upperCase = false;
      key += letter.toUpperCase();
    } else {
      key += letter.toLowerCase();
    }
  }
  return key;
}

// Returns true if the cell where cellData was read from is empty.
// Arguments:
//   - cellData: string
function isCellEmpty(cellData) {
  return typeof(cellData) == "string" && cellData == "";
}

// Returns true if the character char is alphabetical, false otherwise.
function isAlnum(char) {
  return char >= 'A' && char <= 'Z' ||
    char >= 'a' && char <= 'z' ||
    isDigit(char);
}

// Returns true if the character char is a digit, false otherwise.
function isDigit(char) {
  return char >= '0' && char <= '9';
}

//------------------------------------------------------------------------------------------------------------------------------

// Generate an array of objects from two arrays.  
// Array values are used to create an object with two properties: address, body
// @param address array
// @param body array
// @param counter interger
function getObjectsFromTwoArrays(address, body, counter) {
   var objects =[]; 
  
  for (var i = 0; i < counter; i++) {
    var object = {};
    object.body = body[i];
    object.address = address[i];
    objects.push(object);
  }  
  return objects;
}

//Detail for requestupdates
function getRequestedActionItems () {
  var actionItemObjects = actionItem.getActionItemObjects();
  var requestUpdatesbody = [];
  var requestUpdatesAddresses = [];
  var counter = 0;

  // Loop thru the array of objects building string blocks from objects 
  for(var i =0;actionItemObjects.length > i; i++) {
    //build string message
    /*
    var requestUpdate = "<p><strong>Action Item #: </strong>" + actionItemObjects[i].referencenumber + "<br>";
    requestUpdate += "<strong>Action/ Issue Decription: </strong>" + actionItemObjects[i].description + "<br>";
    requestUpdate += "<strong>Owner: </strong>" + actionItemObjects[i].owner + "<br>";
    requestUpdate += "<strong>Status: </strong>" + actionItemObjects[i].status + "<br>";
    requestUpdate += "<strong>Updates: </strong>" + actionItemObjects[i].updates + "<br></p>";
    */
    
    var requestUpdate = (<r><![CDATA[ <tr><td style="padding:2px"> ]]></r>).toString() + actionItemObjects[i].referencenumber + "</td>";
    requestUpdate += (<r><![CDATA[ <td style="padding:2px"> ]]></r>).toString() + actionItemObjects[i].description + "</td>";
    requestUpdate += (<r><![CDATA[ <td style="padding:2px"> ]]></r>).toString() + actionItemObjects[i].owner + "</td>";
    requestUpdate += (<r><![CDATA[ <td style="padding:2px"> ]]></r>).toString() + actionItemObjects[i].status + "</td></tr>";
    
    var address = actionItemObjects[i].updateremail;
    // populate arrays with variable string from sheet objects
    if(actionItemObjects[i].needsupdating === true) {
      requestUpdatesbody.push(requestUpdate);
      requestUpdatesAddresses.push(address);
      counter += 1;
    }
  }
  //get an array of object from the arrays
  var requestUpdates = getObjectsFromTwoArrays(requestUpdatesAddresses, requestUpdatesbody, counter);
  var requestUpdatesGrouped = groupRequestUpdates (requestUpdates);
 
  // for (var i = 0; i < requestUpdatesGrouped.length; i++){
  //  Logger.log("Email: " + requestUpdatesGrouped[i].address + "\n\r" + " Body_0: " + requestUpdatesGrouped[i].body_0 + "\n\r" + " Body_1: " + requestUpdatesGrouped[i].body_1 + "\n\r"+ " Body_2: " + requestUpdatesGrouped[i].body_2 + "\n\r");
  // } 
   
  // for (var i = 0; i < requestUpdates.length; i++){
  //  Logger.log("Email: " + requestUpdates[i].address + " Body: " + requestUpdates[i].body + "/n/r")
  // }

  // for (var i = 0; i < requestUpdatesGrouped.length; i++){
  //  Logger.log("Email: " + requestUpdatesGrouped[i].address + "\n\r" + " Body: " + requestUpdatesGrouped[i].body + "\n\r")
  // }

  //old code 2.3
  //return requestUpdates;
  return requestUpdatesGrouped;
}

//Create an array of objects from an array of objects
//@params array of objects
//If any object's address property is equal to another object's address property, 
//create a new object with the first objects address property and all the body properties as body_x ... (x+1) property
//return array of objects
function groupRequestUpdates ( requestUpdates ) {
  var objects = [];  //Return all objects in array of objects
  var emailaddress = [];  //Record each loop address to ensure only one pass for each unique address property from requestUpdates
  
  //Loop through each requestUpdates returning a new object with two properties:address,body
  for ( var i = 0; i < requestUpdates.length; i++) {
    var object = {};  //New object 
    var bodies = [];  //array to hold all the individual bodies for requestedUpdates grouping based on same address from requestedUpdates 
    var address = requestUpdates[i].address; //To test if the address has already been loop through
    
    if ( emailaddress.indexOf(address) != -1) { //Test to make sure address has not already been loop through 
         continue;  //If already looped move to next requestedUpdates
    }
    
    //Loop through all requstUpdates if they have the same address, add the body to the bodies array
    //Build the new object define address and body property. Populate body property with all bodies from similar address
    for ( var j = 0; j < requestUpdates.length; j++) {
      if ( requestUpdates[i].address === requestUpdates[j].address) {
           bodies.push( requestUpdates[j].body);
      }
    }
    
    object.address = requestUpdates[i].address; // define property
    var allBodies = "";
    for ( var x = 0; x < bodies.length; x++) {  //put all the bodies together
          allBodies +=bodies[x];
       
      //var property = "body_" + (x);
      //object[property] = bodies[x];
    }
    object.body = allBodies; // define property
    objects.push( object ); //Add the array of objects
    emailaddress.push( requestUpdates[i].address ); //add address to processed emailaddresses
  }
  return objects;
}

/*------------------------------------------------------------------------------------------------------------------------------------------------------------------
ACTION ITEM FRAMEWORK








--------------------------------------------------------------------------------------------------------------------------------------------------------------------*/

var actionItem = {
  
  getActionItemObjects: function () {
    var ss = SpreadsheetApp.getActiveSpreadsheet();  //Make Master Action Item Log active sheet
    var sheet = ss.getSheets()[0];  // Get the range of cells that store action item data.
    var actionItemDataRange = ss.getRangeByName("actionItemLog");
    var actionItemObjects = getRowsData(sheet, actionItemDataRange);  // For every row of action item data, generate an action Item object.
    return actionItemObjects;
  },
  
  // Create a array of objects using getRowsData for the Action Item Log  **copy
  // Using three arrays populate each new, updated, or closed action items in a html message
  // Interate through each array and populate into string
  // return the Action Item Log section of the email
  getActionItemMessage: function (frequency) {
    var actionItemObjects = actionItem.getActionItemObjects();
    var frequency = frequency;
    var newActions = [];
    var updatedActions = [];
    var closedActions = [];
    
  //New,Updated,Closed tables for diplaying action items  
    var noUpdate = (<r><![CDATA[<p>Nothing to report today.<br></p>]]></r>).toString();
    var noUpdateAll = (<r><![CDATA[<p><Strong>Action Item Updates</strong></p><br><p>No new activity.<br></p>]]></r>).toString();
    
    var newActionTitleString = (<r><![CDATA[ <h4>New Action Items</h4></br>
    <table style="width:100%;background-color:white;-webkit-border-radius: 5px;border-radius: 5px;background-color:#F7F7DF;">
      <thead style="padding:10px;-webkit-border-radius: 5px;border-radius: 5px;margin-bottom:1px;background-color:#DFDFF7;">
        <tr>
          <td style="padding:2px;"><h4>Action Item</h4></td><td style="padding:2px;"><h4>Decription</h4></td><td style="padding:2px;"><h4>Owner</h4></td><td><h4>Action Needed</h4></td><td style="padding:2px;"><h4>Module</h4></td>
        </tr>
      </thead>
     <tbody>]]></r>).toString();
    
    var newActionFooterString = (<r><![CDATA[</tbody></table>]]></r>).toString();
    var newActionsListString = "";
    
    var updatedActionTitleString = (<r><![CDATA[ <h4>Updated Action Items</h4></br>
    <table style="width:100%;background-color:white;-webkit-border-radius: 5px;border-radius: 5px;background-color:#F7F7DF;">
      <thead style="padding:10px;-webkit-border-radius: 5px;border-radius: 5px;margin-bottom:1px;background-color:#DFDFF7;">
        <tr>
          <td style="padding:2px;"><h4>Action Item</h4></td><td style="padding:2px;"><h4>Decription</h4></td><td style="padding:2px;"><h4>Owner</h4></td><td><h4>Status</h4></td><td style="padding:2px;"><h4>Updates</h4></td>
        </tr>
      </thead>
      <tbody>]]></r>).toString();
    
    var updatedActionFooterString = (<r><![CDATA[</tbody></table>]]></r>).toString();
    var updatedActionsListString = ""; 
    
    var closedActionTitleString = (<r><![CDATA[ <h4>Closed Action Items</h4></br>
    <table style="width:100%;background-color:white;-webkit-border-radius: 5px;border-radius: 5px;background-color:#F7F7DF;">
      <thead style="padding:10px;-webkit-border-radius: 5px;border-radius: 5px;margin-bottom:1px;background-color:#DFDFF7;">
        <tr>
          <td style="padding:2px;"><h4>Action Item</h4></td><td style="padding:2px;"><h4>Decription</h4></td><td style="padding:2px;"><h4>Owner</h4></td>
        </tr>
      </thead>
      <tbody>]]></r>).toString();
    
    var closedActionFooterString = (<r><![CDATA[</tbody></table>]]></r>).toString();                           
    var closedActionsListString = "";
    
  //end of tables  
                                     
  //
    for(var i =0;actionItemObjects.length > i; i++) {
      
      var newAction = (<r><![CDATA[ <tr><td style="padding:2px"> ]]></r>).toString() + actionItemObjects[i].referencenumber + "</td>";
      newAction += (<r><![CDATA[ <td style="padding:2px;min-width:75px;"> ]]></r>).toString() + actionItemObjects[i].description + "</td>";
      newAction += (<r><![CDATA[ <td style="padding:2px"> ]]></r>).toString() + actionItemObjects[i].owner + "</td>";
      newAction += (<r><![CDATA[ <td style="padding:2px"> ]]></r>).toString() + actionItemObjects[i].actionneeded + "</td>";
      newAction += (<r><![CDATA[ <td style="padding:2px"> ]]></r>).toString() + actionItemObjects[i].module + "</td></tr>";
      
      var updatedAction = (<r><![CDATA[ <tr><td style="padding:2px; vertical-align:text-top;"> ]]></r>).toString() + actionItemObjects[i].referencenumber + "";
      updatedAction += (<r><![CDATA[ <td style="padding:2px; min-width:200px; vertical-align:text-top;"> ]]></r>).toString() + actionItemObjects[i].description + "</td>";
      updatedAction += (<r><![CDATA[ <td style="padding:2px; vertical-align:text-top;"> ]]></r>).toString() + actionItemObjects[i].owner + "</td>";
      updatedAction += (<r><![CDATA[ <td style="padding:2px; vertical-align:text-top;"> ]]></r>).toString() + actionItemObjects[i].status + "</td>";
      updatedAction += (<r><![CDATA[ <td style="padding:2px; min-width:450px; vertical-align:text-top;"> ]]></r>).toString() + actionItemObjects[i].updates + "</td></tr>";
      
      var closedAction = (<r><![CDATA[ <tr><td style="padding:2px"> ]]></r>).toString() + actionItemObjects[i].referencenumber + "</td>";
      closedAction += (<r><![CDATA[ <td style="padding:2px;min-width:75px;"> ]]></r>).toString() + actionItemObjects[i].description + "</td>";
      closedAction += (<r><![CDATA[ <td style="padding:2px"> ]]></r>).toString() + actionItemObjects[i].owner + "</td></tr>";
      
      
      if(frequency === "daily") {
        if(actionItemObjects[i].isupdated === true) {
          updatedActions.push(updatedAction);
        }
        else if(actionItemObjects[i].isnew === true) {
          newActions.push(newAction);
        }
        else if(actionItemObjects[i].isclosed === true) {
          closedActions.push(closedAction);
        }
       }
      else if(frequency === "weekly") {
        if(actionItemObjects[i].isupdatedweek === true) {
          updatedActions.push(updatedAction);
        }
        else if(actionItemObjects[i].isnewweek === true) {
          newActions.push(newAction);
        }
        else if(actionItemObjects[i].isclosedweek === true) {
          closedActions.push(closedAction);
        }
       }
    }
  
    if( newActions.length > 0) {                           
      for( var i = 0;newActions.length > i;i++) {
            newActionsListString += newActions[i].toString();
      }
    } 
    else {
      newActionsListString = noUpdate;
    }
                                
    if( updatedActions.length > 0) {
      for( var i = 0;updatedActions.length > i;i++) {
           updatedActionsListString += updatedActions[i].toString();
      }
    }
    else {
      updatedActionsListString = noUpdate;
    }
  
    if(closedActions.length > 0) {                               
      for(var i = 0;closedActions.length > i;i++) {
       closedActionsListString += closedActions[i].toString();
      }
    }
    else {
      closedActionsListString = noUpdate;
    } 
   
   var newActionItemSection = newActionTitleString + newActionsListString + newActionFooterString;
   var updatedActionItemSection = updatedActionTitleString + updatedActionsListString + updatedActionFooterString;
   var closedActionItemSection = closedActionTitleString + closedActionsListString + closedActionFooterString;   
   var actionItemUpdates =  newActionItemSection + updatedActionItemSection + closedActionItemSection;
    
    //To preview Action Item Log Message
    //Logger.log(newActionItemSection + "\n" + updatedActionItemSection + "\n" + closedActionItemSection);
    
    return actionItemUpdates;
  },
  //returns array of approved report names
  getApprovedReports: function (frequency) {
      //Go to sheet
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheets()[5];
      //Range
      var rangename = "approvedreports";
      var approvedReportRange = ss.getRangeByName(rangename);
      //Object and array for strong
      var approvedReports = [];
      var approvedReportObjects = getRowsData(sheet, approvedReportRange);
      var length = approvedReportObjects.length;
      
      //for each object if approved today create string and add to array 
      for(var i = 0;i < length;i++){
        var approvedReport = "<p><strong>Report Name: </strong>" + approvedReportObjects[i].reportname + "<br></p>";
       
        if(frequency === "daily") {
          if(approvedReportObjects[i].isapproved === true) {
            approvedReports.push(approvedReport);
          }
        }
        else if(frequency === "weekly") {
            if(approvedReportObjects[i].isapprovedweekly === true) {
            approvedReports.push(approvedReport);
          }
        }
      }
      //Approved Reports
      var approvedReportTitleString = (<r><![CDATA[ <h4>Approved Reports</h4></br> ]]></r>).toString(); 
      var approvedReportsString = "";
      var reportLength = approvedReports.length;
      var noUpdate = (<r><![CDATA[<p>Nothing to report today.<br></p>]]></r>).toString();
      
      if(reportLength > 0) {
        for(var i = 0;reportLength > i; i++){
        approvedReportsString += approvedReports[i].toString() + "\n";
        }
      }
      else {
        approvedReportsString = noUpdate;
      }
      return approvedReportTitleString + approvedReportsString;
  }
};

/*----------------------------------------------------------------------------------------------------------------------------------------------------------------
EMAIL FRAMEWORK






------------------------------------------------------------------------------------------------------------------------------------------------------------------*/



var ActionEmail = {

  getFile: function () {
      //If you want to send by e-mail the exported document, from script, use the following code:
      // get authorization
      var oauthConfig = UrlFetchApp.addOAuthService("google");
      oauthConfig.setAccessTokenUrl("https://www.google.com/accounts/OAuthGetAccessToken");
      oauthConfig.setRequestTokenUrl("https://www.google.com/accounts/OAuthGetRequestToken?scope=https://spreadsheets.google.com/feeds/");
      oauthConfig.setAuthorizationUrl("https://www.google.com/accounts/OAuthAuthorizeToken");
      oauthConfig.setConsumerKey("anonymous");
      oauthConfig.setConsumerSecret("anonymous");

      //supposedly even better code, dont quite understand it though
      //oauthConfig.setConsumerKey(ScriptProperties.getProperty("consumerKey"));
      //oauthConfig.setConsumerSecret(ScriptProperties.getProperty("consumerSecret"));
      
      // We make this request to ensure that we are authorized prior to
      // actually using the script. Otherwise, the UrlFetchApp.fetch() would serve no purpose 
      var requestData = {
      "method": "GET",
      "oAuthServiceName": "google",
      "oAuthUseToken": "always"};
  
      // the exportUrl also sets the export format of the document
      var exportUrl = "https://spreadsheets.google.com/feeds/download/spreadsheets/Export?key=" + "tr12fyjeVyz8HFsFmzMzBxA" + "&exportFormat=" + "xls";
      var file = UrlFetchApp.fetch(exportUrl, requestData);  // return type: HTTPResponse - the http response data
      
      return file;
  
  },
  // Timestamp for file name and subject
  // Return today's date formated as 'MM.DD.YYYY'
  getDateStamp: function()  {
      //today's date
      var todaydate = new Date();
      var day = ((todaydate.getDate()<10) ? "0" : "")+ todaydate.getDate();
      var month = (((todaydate.getMonth() + 1) <10) ? "0" : "") + (todaydate.getMonth()+ 1);
      var year = todaydate.getYear ();
      var datestamp = month + "." + day + "." + year;
      return datestamp;
  },
  //if the receipt is able to receive hmtl messages 
  getMasterWeeklyHeader: function () {
    
  var htmlmessageHeader = (<r><![CDATA[ 
                    <table style="background-color:WhiteSmoke;-webkit-border-radius: 5px;border-radius: 5px;width:100%;border:0px;">
                    <thead style="-webkit-border-radius: 5px;border-radius: 5px;text-align:center">
                    <tr>
                    <td>
                    <h4><FONT COLOR="808080">This is an automatically generated email, please do not reply.</FONT></h4>
                    <h5><FONT COLOR="808080">Please find the attached Master Action Item Log.</FONT></h5>
                    <hr />
                    </td>
                    </tr>
                    </thead>
                    <tbody>
                    
                    <p><div style="font-weight:bold">Attention:</div> Distribution List  </p>
                    <p><div style="font-weight:bold">Subject:</div><em>Oracle Project: Finance Stream Weekly Digest</em></p>
                    
                    <div style="font-family: “Lucida Sans Unicode”, sans-serif;font-size: 15px;font-style: normal;font-weight: normal;text-transform: normal;letter-spacing: -0.6px;line-height: 1.5em;">
                    <p><em>Enclosed find this week's updates on action items and approved reports.</em></p>
                    
                    <p>The Master Action Item Log is a collection of tasks requiring measurable resources and attention.  The log is exclusive 
                    of open design issues, which are logged by Oracle project management and maintained by David Vincent.     </p>
                    
                    <p>All updates are to be submitted to <a href=mailto:stella.fisher@aholdusa.com title="Update Action Item Log">Stella Fisher</a>, 
                    unless otherwise advised.     </p>
                    
                    <p>The approved report list logs recent Oracle application related reports that have been approved by the Reporting Stream.  
                    Once the report stream approves a report, the approved report is ready to be developed.       </p>
                    
                    <p>The action item log and report status will be reviewed by the Oracle Finance team weekly as part of our team meeting.  </p>
                    
                    <p>Below is a summary of this week's activity,
                    </p>
                    
   ]]></r>).toString();  
   
   return htmlmessageHeader;
   },
   //if the receipt is able to receive hmtl messages   
   getMasterDailyHeader: function () {
     
        var htmlmessageHeader = (<r><![CDATA[ 
        
                          <table style="background-color:WhiteSmoke;-webkit-border-radius: 5px;border-radius: 5px;width:100%;">
                          <thead style="-webkit-border-radius: 5px;border-radius: 5px;text-align:center">
                          <tr>
                          <td>
                          <h4><FONT COLOR="808080">This is an automatically generated email, please do not reply.</FONT></h4>
                          <h5><FONT COLOR="808080">Please find the attached Master Action Item Log.</FONT></h5>
                          <hr />
                          </td>
                          </tr>
                          </thead>
                          <tbody>
                          
                          <p><div style="font-weight:bold">Attention:</div> Distribution List  </p>
                          <p><div style="font-weight:bold">Subject:</div><em>Oracle Project: Finance Stream Daily Digest</em></p>
                          
                          <div style="font-family: “Lucida Sans Unicode”, sans-serif;font-size: 15px;font-style: normal;font-weight: normal;text-transform: normal;letter-spacing: -0.6px;line-height: 1.5em;">
                          <p><em>Enclosed find today’s updates on action items and approved reports.</em></p>
                          
                          <p><strong>Update:</strong>If you would prefer to only receive this digest once a week, 
                          sign-up <a href="https://docs.google.com/a/ahold.com/spreadsheet/viewform?formkey=dEZZNUZpQTFIMXQ2X1BnX0lOVWVPcmc6MQ#gid=0" title="Sign-Up">here</a>, and you will stop receiving this 
                          daily digest.  The Weekly digest will be distributed every Friday by the end of day.</p>
                          
                          <p>The Master Action Item Log is a collection of tasks requiring measurable resources and attention.  The log is exclusive 
                          of open design issues, which are logged by Oracle project management and maintained by David Vincent.     </p>
                          
                          <p>All updates are to be submitted to <a href=mailto:stella.fisher@aholdusa.com title="Update Action Item Log">Stella Fisher</a>, 
                          unless otherwise advised.     </p>
                          
                          <p>The approved report list logs recent Oracle application related reports that have been approved by the Reporting Stream.  
                          Once the report stream approves a report, the approved report is ready to be developed.       </p>
                          
                          <p>The action item log and report status will be reviewed by the Oracle Finance team weekly as part of our team meeting.  </p>
                          
                          <p>Below is a summary of today's activity,
                          </p>

                          
         ]]></r>).toString();
         

   
    return htmlmessageHeader;
    },
    ////Email body + footer
    getMasterFooter: function () {
       var htmlmessageFooter =  (<r><![CDATA[ 
                    </div>
                    </tbody>
                    <tfooter>
                    <div style="background-color:#F0EABD;padding:20px;margin-bottom:0px;margin-top:20px;-webkit-border-radius: 5px;border-radius: 5px;border:0px">
                    <hr />
                    <p style="font-weight:bold;font-size:large"><FONT COLOR="000080">Ahold</FONT><FONT COLOR="6495ED">USA</FONT><FONT COLOR="A9A9A9">SIMPLIFY FOR SUCCESS <em>powered by,</FONT> <FONT COLOR="B22222">ORACLE</FONT></em></p>
                    <hr />
                    <p><em>You are receiving this message because you have been granted access to the <a href="https://sites.google.com/a/ahold.com/oracle-finance-team/" title="Oracle Finance Team">Oracle Finance Team Google site</a>.</em> <br />
                    <em>If you would like to be removed from the distribution send a message <a href=mailto:scott.lawrence@aholdusa.com title="Remove me">here</a></em>.</p>
                    </div>
                    </tfooter>
                    </table>
                                           
  ]]></r>).toString();
  
  return htmlmessageFooter;
  }
    
    
};
/*-------------------------------------------------------------------------------------------------------------------------------------------------------------------
EMAILS




--------------------------------------------------------------------------------------------------------------------------------------------------------------------*/

//
function sendUpdateRequest () {
  //
  var requestedActionItems = getRequestedActionItems();
  var recipient = "";
  var header = (<r><![CDATA[ 
  
              <table style="font-family:Helvetica;-webkit-border-radius: 5px;border-radius: 5px;">
              <thead style="-webkit-border-radius: 5px;border-radius: 5px;text-align:center">
              </thead>
              <tbody style="font-family: “Lucida Grande”, sans-serif;font-size: 11.67px;font-style: normal;font-weight: normal;text-transform: normal;letter-spacing: normal;line-height: 1.4em;">
              <p>Dear Recipient,</p>

              <p>The following Oracle Project: Finance Stream action item(s) has not been updated in 
              a while, could you provide us with an update on current progress, status, and estimated date of completion.</p>

              <p>Please provide all updates to <a href="https://docs.google.com/a/ahold.com/spreadsheet/viewform?formkey=dHIxMmZ5amVWeXo4SEZzRm16TXpCeEE6MQ#gid=12">here</a>
              </p><br>
                  
             <table style="width:100%;background-color:white;-webkit-border-radius: 5px;border-radius: 5px;background-color:#F7F7DF;">
              <thead style="padding:10px;-webkit-border-radius: 5px;border-radius: 5px;margin-bottom:1px;background-color:#DFDFF7;">
                <tr>
                <td style="padding:2px;"><h4>Action Item</h4></td><td style="padding:2px;"><h4>Decription</h4></td><td style="padding:2px;"><h4>Owner</h4></td><td style="padding:2px;"><h4>Status</h4></td>
                </tr>
              </thead>
                <tbody>

  ]]></r>).toString();

  var footer =  (<r><![CDATA[              
              </tbody></table>
              <br><p>Thank You,
              Oracle Project: Finance Stream</p>
              </tbody>
              </table>
              
]]></r>).toString();

  var body = "";
  var htmlbody = "";
  var cc = "scott.lawrence@aholdusa.com";
  var subject = "Oracle Finance Stream: Request For Action Item Update";
  var name = "Oracle Project Finance";
  
  for(var i = 0;i < requestedActionItems.length; i++){
    if(typeof requestedActionItems[i].address === "undefined"){
    continue;
    }

    recipient = requestedActionItems[i].address;
    body = requestedActionItems[i].body;
    htmlbody = header + body + footer
    MailApp.sendEmail(recipient, subject, body, {name:name, cc:cc, noReply:true, htmlBody: htmlbody});
  }

}



// Master Action Item Log Email
// Send a end of day 
function sendMasterActionEmailDaily () {
  var file = ActionEmail.getFile();
  var content = file.getContent();  // return type: Byte - Gets the content of an HTTP response
  var datestamp = ActionEmail.getDateStamp();
  var actionItemLogmessage = actionItem.getActionItemMessage("daily");    //Action Item Log Message
  var approvedReportsString = actionItem.getApprovedReports("daily");
  var htmlmessageHeader = ActionEmail.getMasterDailyHeader();
  var htmlmessageFooter = ActionEmail.getMasterFooter();
  var recipients = "";
  var bcc = "scott.lawrence@aholdusa.com";
             
             //ltowsen@aholdusa.com,mark.walker@aholdusa.com,allen.huang@aholdusa.com,carla.johnson@aholdusa.com,clau@aholdusa.com,lkinsey@aholdusa.com,\
             //dhooper@aholdusa.com          
             
  //email contents - non-html 
  var message = "This is an automatically generated email, please do not reply.\r\n\r\n" +
      "Please find the attached Master Action Item Log for " + datestamp + "\r\n\r\n\r\n\r\n" + 
      "If you would like to be removed from the distribution send a message to Scott Lawrence, scott.lawrence@aholdusa.com\r\n";
 
  var htmlmessage = htmlmessageHeader + actionItemLogmessage + approvedReportsString + htmlmessageFooter;                          
  var myFiles = [{fileName:"Master Action Item Log " + datestamp + ".xlsx", content:content, mimeType:"application//xlsx"}];
  
  MailApp.sendEmail(recipients, "Oracle Project: Finance Master Action Item Log " + datestamp, message, {name: "Oracle Project Finance", attachments: myFiles, bcc: bcc, noReply: true, htmlBody: htmlmessage});
}

// Master Action Item Log Email
// Send a end of day 
function sendMasterActionEmailWeekly () {
  var file = ActionEmail.getFile();
  var content = file.getContent();  // return type: Byte - Gets the content of an HTTP response
  var datestamp = ActionEmail.getDateStamp();
  var actionItemLogmessage = actionItem.getActionItemMessage("weekly");    //Action Item Log Message
  var approvedReportsString = actionItem.getApprovedReports("weekly");
  var htmlmessageHeader = ActionEmail.getMasterWeeklyHeader();
  var htmlmessageFooter = ActionEmail.getMasterFooter();
  var recipients = "";
  var bcc = "scott.lawrence@aholdusa.com, wnace@aholdusa.com, lchruney@aholdusa.com, jwhitten@aholdusa.com, kathleen.Connors@aholdusa.com,\
             mheffern@aholdusa.com, jmerriman@aholdusa.com, dan.rosenberry@aholdusa.com, leesnyder@ahold.com, stella.fisher@aholdusa.com,\
             rute.silva@retail-consult.com, vitor.oliveira@retail-consult.com,\
             brandy.mccabe@aholdusa.com, scott.lawrence@aholdusa.com, sfranklin@aholdusa.com, mlkarns@ahold.com,\
             tom.donnelly@ahold.com, dawn.bishop@ahold.com, dave.burger@aholdusa.com";
  
  //email contents - non-html 
  var message = "This is an automatically generated email, please do not reply.\r\n\r\n" +
      "Please find the attached Master Action Item Log for " + datestamp + "\r\n\r\n\r\n\r\n" + 
      "If you would like to be removed from the distribution send a message to Scott Lawrence, scott.lawrence@aholdusa.com\r\n";
    
  var htmlmessage = htmlmessageHeader + actionItemLogmessage + approvedReportsString + htmlmessageFooter;                          
  var myFiles = [{fileName:"Master Action Item Log " + datestamp + ".xlsx", content:content, mimeType:"application//xlsx"}];
  
  MailApp.sendEmail(recipients, "Oracle Project: Finance Master Action Item Log " + datestamp, message, {name: "Oracle Project Finance", attachments: myFiles, bcc: bcc, noReply: true, htmlBody: htmlmessage});
}



