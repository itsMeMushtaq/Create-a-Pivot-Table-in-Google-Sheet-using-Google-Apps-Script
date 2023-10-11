function onOpen() {
    let ui = SpreadsheetApp.getUi();
    ui.createMenu('Pivot Table').addItem('Sales Report', 'filterData').addToUi();
} // end of onOpen()

function filterData() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName("Sales");
    const sourceData = sheet.getRange('A:D');

    // const pivotSheet = spreadsheet.insertSheet(spreadsheet.getActiveSheet().getIndex() + 1); // Adds a New Sheet
    const pivotSheet = spreadsheet.getSheetByName("PivotSheet");
    pivotSheet.setHiddenGridlines(true); // Hides Gridlines

    // Setting Location for PivotTable
    const pivotRange = pivotSheet.getRange('A3');

    // Initiate PivotTable at pivotRange
    let pivotTable = pivotRange.createPivotTable(sourceData);
    pivotTable = pivotRange.createPivotTable(sourceData);

    // Let's create PivotGroup
    let pivotGroup = pivotTable.addRowGroup(1);
    pivotTable = pivotRange.createPivotTable(sourceData);
    pivotGroup = pivotTable.addRowGroup(1);
    pivotGroup = pivotTable.addColumnGroup(3);

    pivotTable = pivotRange.createPivotTable(sourceData);
    let pivotValue = pivotTable.addPivotValue(4, SpreadsheetApp.PivotTableSummarizeFunction.SUM);
    pivotGroup = pivotTable.addRowGroup(1);
    pivotGroup = pivotTable.addColumnGroup(3);

    const criteria = SpreadsheetApp.newFilterCriteria().whenNumberGreaterThan(0).build();
    pivotTable.addFilter(4, criteria);
} // end of filterData()
