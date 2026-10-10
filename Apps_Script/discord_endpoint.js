function doPost(e){
    try {
        // Parse the incoming JSON data from the Python bot
        var data = JSON.parse(e.postData.contents);

        // Example: Log the data or write it to a Google Sheet
        var command = data.command;
        var username = data.username;
        var message = data.message;

        if(command == 'unavailable')
        {
            unavailable(username, message);
        }

        // Return a success response back to the Python bot
        return ContentService.createTextOutput(JSON.stringify({ "status": "success", "username": username, "message": message }))
                                .setMimeType(ContentService.MimeType.JSON);
                            
    } catch (error) {
        return ContentService.createTextOutput(JSON.stringify({ "status": "error", "message": error.toString() }))
                                .setMimeType(ContentService.MimeType.JSON);
    }
}

/**
 * 
 * @param {string} username 
 * @param {string} msg 
 */
function unavailable(username, msg)
{
    const calendarSheet = SpreadsheetApp.openById('1E_nqdi1ZDexe-v64TkuQq_Th8ScHAq6-KZlFiyxEKjI').getSheetByName("Sheet1");
    const searchCalendar = calendarSheet.getDataRange().createTextFinder('10');

    // Custom configurations (Optional)
    searchCalendar.matchEntireCell(true); // Matches whole cell content only
    searchCalendar.matchCase(true);       // Case-sensitive search

    // Find the first occurrence
    const foundRange = searchCalendar.findNext();

    const cellLength = getCalendarWeekHeight();

    if (foundRange) {
        const row = foundRange.getRow();
        const column = foundRange.getColumn();
        console.log(`Found value at Row: ${row}, Column: ${column}, with height of ${cellLength}`);
    } else {
        console.log("Value not found.");
    }
}



/** --------------------------- HELPER FUNCTIONS --------------------------- */
/**
 * @return {int} Number of cells taken up by a week
 */
function getCalendarWeekHeight()
{
    const calendarSheet = SpreadsheetApp.openById('1E_nqdi1ZDexe-v64TkuQq_Th8ScHAq6-KZlFiyxEKjI').getSheetByName("Sheet1");
    const numRows = calendarSheet.getLastRow();
    var height = 1;

    // Starts at 1 to skip the header row but we still need to include it in height
    for(let i = 2; i < numRows; i++)
    {
        let cell = calendarSheet.getRange(i, 2);
        if(cell.getBackground() == '#000000')
        {
            break;
        }
        height++;
    }

    return height;
}








