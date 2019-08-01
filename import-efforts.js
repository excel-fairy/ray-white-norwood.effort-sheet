
/**
 * Called by custom menu
 */
function openEffortsPopup() {
    var htmlTemplate = HtmlService.createTemplateFromFile('importefforts');
    htmlTemplate.data = {
        agents: getAgents(),
    };
    var htmlOutput = htmlTemplate.evaluate()
        .setTitle('Efforts input')
        .setWidth(1200)
        .setHeight(600);
    SpreadsheetApp.getUi().showDialog(htmlOutput);
}


/**
 * Map
 * Keys: column
 * Values: Corresponding efforts
 */
var tableColumns = {
    a: 'Connects',
    b: 'Voicemail',
    c: 'Dead End',
    d: 'App Booked',
    e: 'App Completed',
    f: 'CMA',
    g: 'Hot LIst',
    h: 'Pipeline',
    i: 'PM Leads',
    j: 'Loan Leads'
};

/**
 * Map
 * Keys: row
 * Values: Corresponding task
 */
var tableRows = {
    1: 'Open Callbacks',
    2: 'Cold Calls',
    3: 'Past Open',
    4: 'Sold Calls',
    5: 'Email Enquiries',
    6: 'Incoming Calls / VM Call Backs',
    7: 'The Bucket',
    8: 'Door Knocks',
    9: 'Thankyou Cards Sent',
    10: 'Little Erics',
    11: 'Letters Sent',
    12: 'Appraisal Complete',
    13: 'Database Outgoing/Incoming Calls',
    14: 'Pipe Line Calls',
    15: 'Hot List Calls',
    16: 'Homeless Calls',
    17: 'Door Knocks (On DB)',
    18: 'Referral',
    19: 'Appraisal Completed (Was ON DB)',
    20: 'Thankyou Cards Sent',
    21: 'CMA\'s/TP\'s/DB Items Sent'
}

/**
 * Main function
 * Called by HTML button in popup
 */
function importEfforts(data) {
    var rowsToInsert = createEffortsRowsEfforts(data);
    getEffortsToImportRange(rowsToInsert.length).setValues(rowsToInsert);
}

/**
 * Create efforts to insert
 * @param data form data
 */
function createEffortsRowsEfforts(data) {
    var agent = data.agent;
    var date = data.date;
    console.log("Importing efforts for agent '" + agent + "' and date '" + date + "'");
    return createEffortsAndTasksMatrix(agent, date, data);
}

/**
 * Create efforts to insert from input matrix
 * @param agent
 * @param date
 * @param data form data
 */
function createEffortsAndTasksMatrix(agent, date, data) {
    var retVal = [];
    for (var property in data) {
        if (data.hasOwnProperty(property)) {
            if(property.indexOf('-') !== -1) {
                // Table data
                if(data[property]){
                    // Cell has content
                    var splittedProp = property.split('-');
                    var task = tableRows[splittedProp[1]];
                    var effort = tableColumns[splittedProp[0]];
                    var number = data[property];
                    console.log("Effort '" + effort + "' and task '" + task + "' have value: " + number);
                    for (var i = 0; i < number; i++) {
                        retVal.push([agent, date, task, effort]);
                    }
                }
            }
        }
    }
    return retVal;
}

/**
 * Get the list of agents
 * @returns {*} The list of agents
 */
function getAgents() {
    return EFFORT_SPREADSHEET.dataValidSheet.sheet.getRange(
        EFFORT_SPREADSHEET.dataValidSheet.agentsFirstRow,
        EFFORT_SPREADSHEET.dataValidSheet.agentsCol,
        EFFORT_SPREADSHEET.dataValidSheet.sheet.getLastRow() - EFFORT_SPREADSHEET.dataValidSheet.agentsFirstRow + 1,
        1)
        .getValues()
        .map(function (el) {
            return el[0];
        })
        .filter(function (el) {
            return !!el;
        });
}

/**
 * Return the range to import efforts to
 * @param nbEffortsToImport Number of efforts to import
 * @returns {*} The range
 */
function getEffortsToImportRange(nbEffortsToImport) {
    return EFFORT_SPREADSHEET.databaseSheet.sheet.getRange(
        EFFORT_SPREADSHEET.databaseSheet.sheet.getLastRow() + 1,
        EFFORT_SPREADSHEET.databaseSheet.effortsFirstCol,
        nbEffortsToImport,
        EFFORT_SPREADSHEET.databaseSheet.effortsLastCol - EFFORT_SPREADSHEET.databaseSheet.effortsFirstCol + 1);
}
