function doPost(e){
    try {
        // Parse the incoming JSON data from the Python bot
        var data = JSON.parse(e.postData.contents);

        // Example: Log the data or write it to a Google Sheet
        var command = data.command;
        var username = data.username;
        var result;

        if(command == 'unavailable')
        {
            var days = data.days;
            result = unavailable(username, days);
        }
        else if (command == 'confirmSchedule')
        {
            var channel = data.channel;
            result = confirmSchedule(username, channel);
        }
        else if (command == 'confirmCalendar')
        {
            var channel = data.channel;
            result = confirmCalendar(username, channel);
        }
        
        // Return a success response back to the Python bot
        return ContentService.createTextOutput(JSON.stringify({ "status": "success", "username": username, "message": result}))
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
    const calendarSheet = SpreadsheetApp.openById(scriptProps.getProperty('calendarSheet'))
      .getSheetByName(scriptProps.getProperty('calendarName'));
    const cellLength = getCalendarWeekHeight();
    var messages = [];

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
                messages.push(`Successfully removed from the ${days[i]}`);
            }
            else{
              messages.push(`Could not find ${username} on the ${days[i]}.`);
            }
        } else {
            messages.push(`Could not find ${days[i]} in the calendar.`);
        }
    }

    return messages;
}

/**
 * 
 * @param {String} username 
 * @param {String} channel 
 */
function confirmSchedule(username, channel)
{
    var id = null;
    var name = null;
    if(channel == "instructor")
    {
        id = scriptProps.getProperty('instructorSheet');
        name = scriptProps.getProperty('instructorSchedule');
    }
    else if(channel == "trainee")
    {
        id = scriptProps.getProperty('traineeSheet');
        name = scriptProps.getProperty('traineeSchedule');
    }

    const sheet = SpreadsheetApp.openById(id).getSheetByName(name);
    const finder = sheet.getDataRange().createTextFinder(username)
        .matchEntireCell(false)
        .matchCase(false);
    
    const foundCells = finder.findAll();

    if(foundCells.length == 0)
    {
        return `Could not find ${username} in the schedule.`;
    }

    foundCells.forEach(cell => {
        var changeCell = sheet.getRange(cell.getA1Notation());

        var data = changeCell.getValue();

        if(data.indexOf("*") == -1)
        {
            changeCell.setValue("**" + data);
        }

    });

    return `Successfully confirmed schedule for ${username}.`;
}

/**
 * NEEDS WORK
 * @param {String} username 
 * @param {String} channel 
 */
function confirmCalendar(username, channel)
{
    var id = scriptProps.getProperty('calendarSheet');
    var name = scriptProps.getProperty('calendarName');
    var color = null;

    if(channel == 'instructor')
    {
        const colorSheet = SpreadsheetApp.openById(scriptProps.getProperty('colorSheet'))
            .getSheetByName(scriptProps.getProperty('colorName'));
        const colorFinder = colorSheet.getDataRange().createTextFinder(username)
            .matchEntireCell(false)
            .matchCase(false);
        
        const colorCells = colorFinder.findAll();

        colorCells.forEach(cell => {
            var colorCell = colorSheet.getRange(cell.getA1Notation());

            color = colorCell.getBackground();
        });
    }
    else
    {
        color = '#f9cb9c';
    }

    const sheet = SpreadsheetApp.openById(id).getSheetByName(name);
    const finder = sheet.getDataRange().createTextFinder(username)
        .matchEntireCell(false)
        .matchCase(false);
    
    const foundCells = finder.findAll();

    if(foundCells.length == 0)
    {
        return `Could not find ${username} in the schedule.`;
    }

    foundCells.forEach(cell => {
        var changeCell = sheet.getRange(cell.getA1Notation());

        var data = changeCell.getValue();

        changeCell.setBackground(color);
    });

    return `Successfully confirmed calendar for ${username}.`;
}

/** --------------------------- HELPER FUNCTIONS --------------------------- */
/**
 * @return {int} Number of cells taken up by a week
 */
function getCalendarWeekHeight()
{
    const calendarSheet = SpreadsheetApp.openById(scriptProps.getProperty('calendarSheet'))
        .getSheetByName(scriptProps.getProperty('calendarName'));
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






