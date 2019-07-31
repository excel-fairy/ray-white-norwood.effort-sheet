function onOpen() {
    SpreadsheetApp.getUi()
        .createMenu('Run scripts')
        .addItem('Import efforts', 'openEffortsPopup')
        .addToUi();
}