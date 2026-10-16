function onOpen()
{
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('Calendar')
        .addItem('Generate Calendar', 'runCalendar')
        .addItem('Update Sheet Names', runUpdateSheetNames)
        .addToUi();
}

function runCalendar()
{
    MightyfistCalendar.createCalendar();
}

function runUpdateSheetNames()
{
    MightyfistCalendar.updateSheetNames();
}