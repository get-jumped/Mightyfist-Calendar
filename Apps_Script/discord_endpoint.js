function doPost(e){
    try {
        // Parse the incoming JSON data from the Python bot
        var data = JSON.parse(e.postData.contents);

        // Example: Log the data or write it to a Google Sheet
        var command = data.command;
        var username = data.username;

        if(command == 'unavailable')
        {
            var days = data.days;
            unavailable(username, days);
        }
        else if (command == 'confirm')
        {
            var channel = data.channel;
            confirm(username);
        }
        
        // Return a success response back to the Python bot
        return ContentService.createTextOutput(JSON.stringify({ "status": "success", "username": username}))
                                .setMimeType(ContentService.MimeType.JSON);
                            
    } catch (error) {
        return ContentService.createTextOutput(JSON.stringify({ "status": "error", "message": error.toString() }))
                                .setMimeType(ContentService.MimeType.JSON);
    }
}

/**
 * 
 * @param {string} username 
 * @param {Array} days 
 */
function unavailable(username, days)
{
    const calendarSheet = SpreadsheetApp.openById('1E_nqdi1ZDexe-v64TkuQq_Th8ScHAq6-KZlFiyxEKjI').getSheetByName("Sheet1");
    const cellLength = getCalendarWeekHeight();

    for(let i = 0; i < days.length; i++)
    {
        const searchCalendar = calendarSheet.getDataRange().createTextFinder(days[i]);

        // Custom configurations (Optional)
        searchCalendar.matchEntireCell(true); // Matches whole cell content only
        searchCalendar.matchCase(true);       // Case-sensitive search

        // Find the first occurrence
        const foundRange = searchCalendar.findNext();

        if (foundRange) {
            const row = foundRange.getRow();
            const column = foundRange.getColumn();

            const searchDay = calendarSheet.getRange(row, column, cellLength, 1).createTextFinder(username).matchEntireCell(false);
            const unavailableEntry = searchDay.findNext();

            if(unavailableEntry)
            {
                const userRow = unavailableEntry.getRow();
                const userCol = unavailableEntry.getColumn();

                calendarSheet.getRange(userRow, userCol).setFontLine("line-through");
            }
            else{
              console.log(`${username} not found on ${days[i]}`);
            }
        } else {
            console.log("Value not found.");
        }
    }
}

/**
 * 
 * @param {String} username 
 * @param {String} channel 
 */
function confirm(username, channel)
{

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






