
/**
 * Called by custom menu
 */
function openEffortsPopup() {
    var htmlTemplate = HtmlService.createTemplateFromFile('importefforts');
    htmlTemplate.data = {
        agents: ['e', 'a'],
        borrowers: ['Antra Group', 'Ray Petty', 'Fundsquire Pty Ltd']
    };
    var htmlOutput = htmlTemplate.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME)
        .setTitle('Import efforts')
        .setWidth(1500)
        .setHeight(700);
    SpreadsheetApp.getUi().showDialog(htmlOutput);
}


/**
 * Main function
 * Called by HTML button in popup
 */
function importEfforts(data) {
    SpreadsheetApp.getUi().alert ('Loan is being imported. It will appear in the "Loans" tab shortly');
    insertLoanInLoansSheet(data);
}

function insertLoanInLoansSheet(data){
    // Override loanReference (autocomputed) only if the entity is none of the below
    if(data.entityName !== 'Dacosi Investments Pty Ltd (Derek Goh)' && data.entityName !== 'Dacosi ST Pty Ltd (Derek Goh)')
        data.loanReference =  getIncrementedLoanReference(getLastLoanReferenceOfEntity(data.entityName));
    var rowToInsert = buildLoanToInsert(data);
    var loansOriginalSheet = INTEREST_STATEMENT_SPREADSHEET.loansSheet.sheet;
    var lastEntityRow = getLastLoanOfEntityRow(data.entityName);
    var rangeRowToSet = loansOriginalSheet.getRange(lastEntityRow + 1,
        ColumnNames.letterToColumn(INTEREST_STATEMENT_SPREADSHEET.loansSheet.firstLoansColumn),
        1,
        ColumnNames.letterToColumn(INTEREST_STATEMENT_SPREADSHEET.loansSheet.lastLoansColumn) + 1
        - ColumnNames.letterToColumn(INTEREST_STATEMENT_SPREADSHEET.loansSheet.firstLoansColumn));

    duplicateLastEntityRow(lastEntityRow);
    rangeRowToSet.setValues([rowToInsert]);
}

// Duplicate row to get all the data that won't be overwritten
function duplicateLastEntityRow(lastEntityRow){
    var loansOriginalSheet = INTEREST_STATEMENT_SPREADSHEET.loansSheet.sheet;
    loansOriginalSheet.insertRowAfter(lastEntityRow);
    var lastRangeRowOfEntity = loansOriginalSheet.getRange(lastEntityRow,
        ColumnNames.letterToColumn(INTEREST_STATEMENT_SPREADSHEET.loansSheet.firstLoansColumn),
        1,
        loansOriginalSheet.getLastColumn()
        - ColumnNames.letterToColumn(INTEREST_STATEMENT_SPREADSHEET.loansSheet.firstLoansColumn) + 1);
    var rangeRowToCopyDestination = loansOriginalSheet.getRange(lastEntityRow + 1,
        ColumnNames.letterToColumn(INTEREST_STATEMENT_SPREADSHEET.loansSheet.firstLoansColumn),
        1,
        loansOriginalSheet.getLastColumn()
        - ColumnNames.letterToColumn(INTEREST_STATEMENT_SPREADSHEET.loansSheet.firstLoansColumn) + 1);
    lastRangeRowOfEntity.copyTo(rangeRowToCopyDestination);
}

function buildEffortToInsert(data) {
    //TODO
    var row = [];
    var interestRatePercent = data.interestRate / 100;
    row[ColumnNames.letterToColumnStart0('A')] = data.loanReference;
    row[ColumnNames.letterToColumnStart0('B')] = '';
    row[ColumnNames.letterToColumnStart0('C')] = data.entityName;
    row[ColumnNames.letterToColumnStart0('D')] = data.amountBorrowed;
    row[ColumnNames.letterToColumnStart0('E')] = data.dateBorrowed;
    row[ColumnNames.letterToColumnStart0('F')] = '';
    row[ColumnNames.letterToColumnStart0('G')] = data.dueDate;
    row[ColumnNames.letterToColumnStart0('H')] = interestRatePercent;
    row[ColumnNames.letterToColumnStart0('I')] = interestRatePercent * data.amountBorrowed;
    row[ColumnNames.letterToColumnStart0('J')] = 'No';
    row[ColumnNames.letterToColumnStart0('K')] = data.ballooninvestment;
    row[ColumnNames.letterToColumnStart0('L')] = '';
    row[ColumnNames.letterToColumnStart0('M')] = data.borrowerEntity;
    return row;
}


function getLastEffortRow(entityName) {
//TODO
    var lastRow = -1;
    var allLoans = getAllLoansFirstThreeColumns();
    var loanReference = getLastLoanReferenceOfEntity(entityName);
    if(loanReference !== null) { // A loan of this entity has already been imported
        for(var i=0; i < allLoans.length; i++){
            var currentLoanReference = allLoans[i][ColumnNames.letterToColumnStart0(INTEREST_STATEMENT_SPREADSHEET.loansSheet.loanReferenceColumn)];
            if( currentLoanReference === loanReference)
                lastRow = i + INTEREST_STATEMENT_SPREADSHEET.loansSheet.firstLoanRow;
        }
        return lastRow;
    }
    else { // First loan of this entity to be imported
        var beforeEntityLoan = getLastLoanOfEntityBeforeThisEntity(entityName);
        if(beforeEntityLoan !== null) {
            var beforeEntityName = beforeEntityLoan[ColumnNames.letterToColumnStart0(INTEREST_STATEMENT_SPREADSHEET.loansSheet.entityNameColumn)];
            for (var i = 0; i < allLoans.length; i++) {
                var currentLoanEntityName = allLoans[i][ColumnNames.letterToColumnStart0(INTEREST_STATEMENT_SPREADSHEET.loansSheet.entityNameColumn)];
                if (currentLoanEntityName === beforeEntityName)
                    lastRow = i + INTEREST_STATEMENT_SPREADSHEET.loansSheet.firstLoanRow;
            }
            return lastRow;
        }
        else // First loan of this entity to be imported and no entity with a name before this one in the list of loans
            return INTEREST_STATEMENT_SPREADSHEET.loansSheet.firstLoanRow - 1;
    }
}
