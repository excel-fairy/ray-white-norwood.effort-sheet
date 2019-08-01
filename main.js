var EFFORT_SPREADSHEET = {
    databaseSheet: {
        sheet: SpreadsheetApp.getActiveSpreadsheet().getSheetByName("database"),
        effortsFirstCol: ColumnNames.letterToColumn('A'),
        effortsLastCol: ColumnNames.letterToColumn('D'),
        effortsFirstRow: '3'
    },
    dataValidSheet: {
        sheet: SpreadsheetApp.getActiveSpreadsheet().getSheetByName('data valid'),
        agentsCol: ColumnNames.letterToColumn('A'),
        agentsFirstRow: '3'
    }
}
