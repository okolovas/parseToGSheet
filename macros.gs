/**
 * Copyright Google LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     https://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
/**
 * Function is executed when the Spreadsheet is opened or reloaded.
 * onOpen() is used to add UI menu
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Sandbox')
    .addItem('Load task list', 'loadTaskList')
    .addToUi();
}
/** 
 * Main function
 * Creates 2-level list:
 * project
 *   - task
 *   - task
 *   - task
 * project
 *   - task
 * project
 *   - task
 *   - task
 * Get common data from 'Sandbox' sheet
 * and load task links and tester names from YouTrack
 * using token
 */
function loadTaskList() {
  // Get an active Spreadsheet
  var sheet = SpreadsheetApp.getActiveSheet();

  // Get spreadsheet, data range and values for the sandbox task list
  var sandbox = SpreadsheetApp.getActive().getSheetByName("Sandbox");
  var taskRange = sandbox.getDataRange();
  var taskValues = taskRange.getRichTextValues();
  var taskValuesCount = taskValues.length;

  // Get projects list from the corresponding spreadsheet
  var projects = SpreadsheetApp.getActive().getSheetByName("Projects").getDataRange().getValues();
  var projectsCount = projects.length;


  // Iterate the list of all projects and determiane which of them do have tasks
  var currentPos = sheet.getActiveCell().getRow();
  var yt = false;

  var takeThisProject = false;
  for ( var i=0; i<projectsCount; i++ ) {
    takeThisProject = false; // ?
    for ( var j=0; j<taskValuesCount; j++ ) {
      var runs = taskValues[j][0].getRuns();
      var runsLength = runs.length;

      if ( runsLength==3 ) {
        if ( runs[2].getText().trim() == projects[i] ) {
          takeThisProject = true;
          sheet.getRange(currentPos, 2, 1, 1).setValue(projects[i]).setFontWeight("bold");
          currentPos ++;
        }
        else {
          takeThisProject = false;
        }
      }
      else if( takeThisProject ) {
        sheet.getRange(currentPos, 3, 1, 1).setValue( taskValues[j][0].getText() );
        if ( runsLength==7 ) {
          var ytURL = runs[3].getLinkUrl();
          var ytURLParsed = ytURL.split('/');
          var ytURLParsedLenth = ytURLParsed.length;
          var issueId = '-';
          if ( ytURLParsed[ytURLParsedLenth-2]=='issue' )
            issueId = ytURLParsed[ytURLParsedLenth-1];
          sheet.getRange(currentPos, 2, 1, 1).setValue(ytURL);
          if ( !yt && issueId!='-' ) {
            sheet.getRange(currentPos, 5, 1, 1).setValue(getTester(issueId));
//            yt = true;
          }
        }
        currentPos ++;
      }
    }
  }
}

// Get a tester name from YouTrack using token auth
function getTester(issueId) {
    var token = '<ADD YOUR TOKEN HERE>';
    var issueResponse = ytIssueCustomFields('youtrack.YOUR-DOMAIN', token, issueId);

    if (issueResponse.getResponseCode() == 200) {
      customFields = JSON.parse(issueResponse.getContentText()).customFields;

      for (i=0; i<customFields.length; i++) {
        var name = customFields[i].name;
        if ( name=='Тестировщик' && customFields[i].value.length != 0 )
          return result = customFields[i].value[0].name;
      }
    } else {
      return errorHandler(JSON.parse(response.getContentText()));
    }
}

// Gets YouTrack issue custom fields information
function ytIssueCustomFields(domain, token, issueId) {
  head = {
    'Authorization':"Bearer " + token,
    'Content-Type': 'application/json'
  }
  params = {
    headers:  head,
    method : "get",
    muteHttpExceptions: true
  }

  filter = "?fields=id,customFields(id,name,value(name))";
  
  return UrlFetchApp.fetch("https://" + domain + "/api/issues/" + issueId + filter, params);
}

// errorHandler
function errorHandler(error) {
  var errorDesc = ""
  
  if (error.error_children != undefined) {
    for (i=0;i<error.error_children.length;i++) {
      errorDesc = errorDesc + error.error_children[i].error + " ";
    }
  }
  return errorDesc;
}

