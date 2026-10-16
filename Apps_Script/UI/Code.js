function onOpen()
{
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('Calendar')
        .addItem('Generate Calendar', 'runCalendar')
        .addToUi();
}

function runCalendar()
{
    MightyfistCalendar.createCalendar();
}