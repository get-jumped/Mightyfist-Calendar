function onOpen()
{
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('Calendar')
      .addItem('Generate Calendar', 'runCreateCalendar')
      .addItem('Update Sheet Names', 'runUpdateSheetNames')
      .addToUi();
}

function runCreateCalendar() {
  MightyfistCalendar.createCalendar();
}

function runUpdateSheetNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get its file name
  const name = ss.getName();
  var sheetName = SpreadsheetApp.getActiveSheet().getName();
  
  if(name.indexOf("Instructor") != -1)
  {
      MightyfistCalendar.updateSheetNames("Instructor", sheetName);
  }
  else if(name.indexOf("Trainee") != -1)
  {
      MightyfistCalendar.updateSheetNames("Trainee", sheetName);
  }
  else
  {
    throw new Error('Unexpected file name "' + name + '": expected it to contain "Instructor" or "Trainee".');
  }
}