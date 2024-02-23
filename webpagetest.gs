/*
 * App Script to submit tests to WebPageTest and retrieve results
 *
 * Will only work in the context of the matching Google Spreadsheet.
 * https://docs.google.com/spreadsheets/d/1Hz_8griZtkDhCVqSCmeyxuHHZGbzm885nt7ws65GvIM
 *
 * This is currently a work-in-progress and the code has too many 'magic numbers' for my liking e.g. row and column offsets
 */

/*
 * License
 *
 * Copyright (c) 2013-2020 Andy Davies, @andydavies, http://andydavies.me
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this software and
 * associated documentation files (the "Software"), to deal in the Software without restriction,
 * including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense,
 * and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so,
 * subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all copies or substantial
 * portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED
 * TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL
 * THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
 * TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
 */

/**
 * Globals
 */

// Named tabs

var TESTS_TAB = "Tests";
var SCENARIOS_TAB = "Scenarios"

// Named ranges on settings tabs

var SERVER_URL = "ServerURL";
var API_KEY = "APIKey";
var NORMALIZE_KEYS = "NormalizeKeys";

var PARAMETERS_MAP = "ParametersMap";
var RESULTS_MAP = "ResultsMap";

/**
 * Adds WebPageTest menu, with actions to submit tests, check their progress and clear results
 */
function onOpen() {
    
    var spreadsheet = SpreadsheetApp.getActive();
    
    var entries = [{name: "Run Tests", functionName: "submitTests"},
                   {name: "Get Results", functionName: "getResults"},
                   null,
                   {name: "Update Scenario Columns", functionName: "updateScenarioColumns"},
                   {name: "Update Test Columns", functionName: "updateTestColumns"}];
    
    spreadsheet.addMenu("WebPageTest", entries);
};


/**
 * Extracts parameters from spreadsheet and submits tests to WPT
 */

function submitTests() {
    
    var spreadsheet = SpreadsheetApp.getActive();
    var sheet = spreadsheet.getSheetByName(TESTS_TAB);
    
    spreadsheet.toast('Submitting tests…', 'Status', 5);
    
    var range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4);
    
    var server = getServerURL();
    var APIKey = getAPIKey();
    
    var testScenarios = getTestScenarios();
    
    var submitted = 0; // Track how many tests were submitted
    
    for(n = 0; n < range.getNumRows(); n++) {
        var cells = range.offset(n, 0, 1, 4).getValues();
        
        var pageURL = cells[0][0];
        var scenario = testScenarios[cells[0][1]];
        var testURL = cells[0][2];
        var testStatus = cells[0][3];
        
        // If there's no URL for test then it's not been submitted (TODO: what about submission failures i.e. statusCode 400)
        if(testURL == "" && scenario != undefined) {
            
            var params = [{
                    param: "url",
          value: pageURL,
                },
                {
                    param: "f",
                    value: "json"
                }];
            
            params = params.concat(scenario); // TODO: what happens if scenerio doesn't exist?
            
            var querystring = buildQueryString(params);
            
            // Submit tests via POST to allow URLs that exceed 2K
            var wptAPI = server + "/runtest.php";

            var options = {
        method: "post",
        payload: querystring,
        headers: {
          "X-WPT-API-KEY": APIKey,
        },
            };
            
            var response = UrlFetchApp.fetch(wptAPI, options);
            var result = JSON.parse(response.getContentText());
            
            // get a new offset for result cells
            var responseCells = range.offset(n, 2, 1, 2); // TODO: Why not just do this earlier and have two ranges?
            
            if(result.statusCode == 200) {
                responseCells.setValues([[result.data.userUrl, ""]]);
                responseCells.clearNote();
                submitted++;
            }
            else {
                responseCells.setValues([["", result.statusCode]]);
                responseCells.setNote(response);
            }
        }
    }
    
    // If any tests submitted, get a first pass of results and start trigger to poll for results
    if(submitted > 0) {
        
        getResults(); // get result of test submission
        
        var pollingInterval = getPollingInterval(submitted);
        
        spreadsheet.toast('Polling for results until all tests complete…', 'Status', 60);
        
        startTrigger(pollingInterval);
    }
}


/**
 * Checks the status of any uncompleted tests, retrieves the results and inserts into sheet
 */

function getResults() {
    
    var spreadsheet = SpreadsheetApp.getActive();
    var sheet = spreadsheet.getSheetByName(TESTS_TAB);

    // Build querystring, allowing the WPT fields to be normalised (remove - and .) or not
    var normalizeKeys = getNormalizeKeys();

    var params = [{
            param: "f",
            value: "json"
        },
        {
            param: "normalizekeys",
            value: normalizeKeys
        }];

    var querystring = buildQueryString(params);
    
    var range = sheet.getRange(2, 3, sheet.getLastRow() - 1, 2); // Just get URL for test, and status columns
    
    var urls_array = range.getValues();
    
    var resultsMap = getResultsMap();
    
    var outstandingResults = 0; // track how many tests yet to complete
    
    for (var i = 0; i < urls_array.length; i++) {
        
        var url = urls_array[i][0];
        var status = urls_array[i][1];
        
        if (url && status < 200) {
            
            // WebPageTest
            var wptAPI = url + "?" + querystring;
            
            var response = UrlFetchApp.fetch(wptAPI);
            var result = JSON.parse(response.getContentText());
            
            e = sheet.setActiveCell("D" + (2 + i));
            e.setValue(result.statusCode);
            
            if(result.statusCode < 200) {
                outstandingResults++;
            }
            else if(result.statusCode == 200) {
                
                for(var column in resultsMap) {
                    cell = sheet.setActiveCell(column + (2 + i));
                    
                    try {
                        var value = eval("result." + resultsMap[column].value);  // TODO: remove eval
                        
                        // some results field may not exist in some tests e.g. SpeedIndex relies on video capture
                        if(value != undefined) {
                            cell.setValue(eval("result." + resultsMap[column].value));
                        }
                    }
                    catch(e) {
                        // do nothing
                    }
                }
            }
        }
    }
    
    // If all tests have completed cancel the trigger
    if(outstandingResults == 0) {
        cancelTrigger()
    }
}


/**
 * Retrieves WPT server URL from Settings tab
 *
 * @return {string} server URL
 */

function getServerURL() {
    var spreadsheet = SpreadsheetApp.getActive();
    var range = spreadsheet.getRange(SERVER_URL);
    
    return range.getValue(); // TODO check for trailing / and add if necessary
}


/**
 * Retrieves WPT API key from Settings tab
 *
 * @return {string} API key
 */

function getAPIKey() {
    
    var spreadsheet = SpreadsheetApp.getActive();
    var range = spreadsheet.getRange(API_KEY);
    
    return range.getValue();
}

/**
 * Retrieves normalizeKeys parameter from Settings tab
 *
 * @return {boolean} normalizeKey
 */

function getNormalizeKeys() {
    
    var spreadsheet = SpreadsheetApp.getActive();
    var range = spreadsheet.getRange(NORMALIZE_KEYS);
    
    return range.getValue();
}

/**
 * Builds a querystring
 *
 * @param {Array.<{param: string, value: string}>} key/value pairs of URL parameters
 *
 * @return {string} querystring
 */

function buildQueryString(params) {
    
    var querystring = params.reduce(function(a, b) {
        return a.concat(encodeURIComponent(b.param) + "=" + encodeURIComponent(b.value));
    }, []);
    
    return querystring.join("&");
}


/**
 * get the parameters map
 */

function getParametersMap() {
    
    return getMap(PARAMETERS_MAP);
}


/**
 * get the results map
 */

function getResultsMap() {
    
    return getMap(RESULTS_MAP);
}


/**
 * Retrieves map of column name, title and API param or results value the column is mapped to
 *
 * @param {string} rangeName - named range within Spreadsheet
 *
 * @return {dictionary} Object.<string, {name: string, value: string}>
 *
 * TODO: check range has 3 columns
 */

function getMap(rangeName) {
    
    var spreadsheet = SpreadsheetApp.getActive();
    var range = spreadsheet.getRange(rangeName);
    var values = range.getValues();
    
    var map = {};
    
    for(n = 0; n < values.length; n++) {
        
        map[values[n][0]] = {
        name: values[n][1],
        value: values[n][2]
        }
    }
    
    return map;
}


/**
 * Sets column headers on Scenarios tab
 */

function updateScenarioColumns() {
    
    var map = getParametersMap();
    
    var spreadsheet = SpreadsheetApp.getActive();
    for(var column in map) {
        cell = spreadsheet.getRange(SCENARIOS_TAB + "!" + column + "1");
        cell.setValue(map[column].name);
    }
}


/**
 * Sets column headers on Tests tab
 */

function updateTestColumns() {
    
    var map = getResultsMap();
    
    var spreadsheet = SpreadsheetApp.getActive();
    for(var column in map) {
        cell = spreadsheet.getRange(TESTS_TAB + "!" + column + "1");
        cell.setValue(map[column].name);
    }
}


/**
 * Builds dictionary of test parameters from Scenarios tab
 *
 * @return {dictionary} Object.<string, {param: string, value: string}>
 */

function getTestScenarios() {
    
    var spreadsheet = SpreadsheetApp.getActive();
    var sheet = spreadsheet.getSheetByName(SCENARIOS_TAB);
    
    var range = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
    
    var map = getParametersMap();
    
    var scenarios = {};
    
    for (y = 1; y < sheet.getLastRow(); y++) {
        
        var scenario = [];
        
        var cell = range.getCell(y, 1);
        var name = cell.getValue();
        
        for (x = 2; x <= sheet.getLastColumn(); x++) {
            
            var cell = range.getCell(y, x);
            
            if(!cell.isBlank()) {
                
                var cellName = cell.getA1Notation();
                var column = cellName.match("[A-Z]*")[0]; // TODO: Urgh
                
                scenario.push({
                param: map[column].value,
                value: cell.getValue()
                });
            }
        }
        
        scenarios[name] = scenario;
    }
    
    return scenarios;
}


/**
 * Starts a trigger to call getResults at a defined interval
 *
 * @param {int} number of minutes between each check (must be 1, 5, 10, 15, 30)
 */

function startTrigger(interval) {
    
    // Check for existing trigger, if it doesn't exist create a new one
    var spreadsheet = SpreadsheetApp.getActive();
    var triggerId =  ScriptProperties.getProperty(spreadsheet.getId());
    
    if(!triggerId) {
        var trigger = ScriptApp.newTrigger("getResults").timeBased().everyMinutes(interval).create();
        
        ScriptProperties.setProperty(spreadsheet.getId() , trigger.getUniqueId());
    }
}


/**
 * Cancels trigger for onResults
 */

function cancelTrigger() {
    
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
    var triggerId =  ScriptProperties.getProperty(spreadsheet.getId());
    
    ScriptProperties.deleteProperty(spreadsheet.getId());
    
    // Locate a trigger by unique ID
    var allTriggers = ScriptApp.getProjectTriggers();
    
    // Loop over all triggers
    for (var i = 0; i < allTriggers.length; i++) {
        if (allTriggers[i].getUniqueId() == triggerId) {
            // Found the trigger so now delete it
            ScriptApp.deleteTrigger(allTriggers[i]);
            break;
        }
    }
}

/**
 * Determine polling interval for checking test results
 *
 * @param {int} number of tests submitted
 *
 * @return {int} interval between check for test status
 *
 * Appscript supports polling intervals of 1, 5, 10, 15, 30 minutes
 *
 * Need to vary polling interval as can exceed urlfetch quota in large test runs
 */

function getPollingInterval(tests) {
    
    var pollingInterval;
    
    if(tests <= 5) {
        pollingInterval = 1;
    } else if (tests <= 10) {
        pollingInterval = 5;
    } else {
        pollingInterval = 30;
    }
    
    return pollingInterval;
}
