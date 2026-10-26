const monthNames = [
        "January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December"
    ];

const scriptProps = PropertiesService.getScriptProperties();

function createCalendar() {
    // Consider giving the sheet that lists the availability a better name
    const sourceSheet = SpreadsheetApp.openById(scriptProps.getProperty('instructorSheet'))
        .getSheetByName(scriptProps.getProperty('instructorSchedule'));
    var instructorDimensions = getTableDimensions(scriptProps.getProperty('instructorSheet'), 
        scriptProps.getProperty('instructorSchedule'));
    let instructorRange = sourceSheet.getRange(instructorDimensions[0],
        instructorDimensions[1],
        instructorDimensions[2],
        instructorDimensions[3]
    );
    var instructorData = instructorRange.getValues();

    const traineeSheet = SpreadsheetApp.openById(scriptProps.getProperty('traineeSheet')).
        getSheetByName(scriptProps.getProperty('traineeSchedule'));
    var traineeDimensions = getTableDimensions(scriptProps.getProperty('traineeSheet'), 
        scriptProps.getProperty('traineeSchedule'));
    let traineeRange = traineeSheet.getRange(traineeDimensions[0],
        traineeDimensions[1],
        traineeDimensions[2],
        traineeDimensions[3]
    );
    var traineeData = traineeRange.getValues();

    // console.log(traineeMerges.map(r => r.getA1Notation()));

    schedule = fillMerges(instructorData);
    // console.log(schedule);

    traineeSchedule = fillMerges(traineeData);
    // console.log(traineeSchedule);

    const date = new Date();
    const year = date.getFullYear();
    const month = date.getMonth();

    let calendar = buildCalendar(year, month);
    // console.log(calendar);

    fillCalendar(calendar, schedule, traineeSchedule);
}

/**
 * Fills in missing times due to merged cells (only fills in missing times if 
 * someone signed up for the timeslot. Reason to avoid confusion with cells that
 * have no timeslots)
 * 
 * @param {2D Array} data a 2D array that represents the schedule
 * @returns newSchedule - a 2D array that is in the format we want to put into the calender spreadsheet
 */
function fillMerges(data) {
    let newSchedule =[];

    for(let i = 1; i < data.length; i++)
    {
        let temp = [];
        for(let j = 1; j < data[i].length; j += 2)
        {
            let name = transformName(removeStars(data[i][j]).trim());
            let time = checkTime(data[i][j]);

            if(time == null)
            {
                // IF there is a name but no time with it
                if(data[i][j] != '' && data[i][j - 1] == '')
                {
                    var back = 1;
                    while(data[i - back][j - 1] == '')
                    {
                        back++;
                    }
                    time = data[i - back][j - 1];
                }
                // IF there is a time but no name
                else if(data[i][j] == '' && data[i][j - 1] != '')
                {
                    name = '';
                    time = '';
                }
                // IF there is a name and time
                else
                {
                    time = data[i][j - 1];
                }

                temp.push(name + " " + transformTime(time));
            }
            else
            {
                temp.push(name + " " + time);
            }
        }
        newSchedule.push(temp);
    }

    return newSchedule
}

// Returns the calendar for the given year and month
function buildCalendar(year, month) {
    let calendar = [];

    //Gets what day (MON, TUES, WED, ...) the first day of the month falls on
    const firstDay = new Date(year, month, 1).getDay();
    
    //Gets the total number of days/last day of the month
    const numDays = new Date(year, month + 1, 0).getDate();

    //Fills the array with null for the days from the previous month
    let week = [];
    for(let i = 0; i < firstDay; i++)
    {
        week.push(null);
    }

    //Creates the calendar for the given month
    for(let j = 1; j <= numDays; j++)
    {
        week.push(j);

        if(week.length == 7)
        {
            calendar.push(week);
            week = []
        }
    }

    if(week.length > 0)
    {
        while(week.length < 7)
        {
            week.push(null);
        }
        calendar.push(week);
    }

    return calendar;
}


/***
 * IMPORTANT PLAN TO MAKE CHANGES WHERE YOU PUT IN EITHER AN ARRAY OR NUM FOR THE 
 * NUMBER OF SCHEDULES, HEADERS, CONSTANTS YOU HAVE IN ONE LINE OF THE CALENDAR
 */
/**
 * Function that fills the calendar with the given month and given schedule
 * @param {2D array} calendar represents the current month
 * @param {2D array} schedule represents the schedule of instructors
 */
function fillCalendar(calendar, schedule, traineeSchedule) {
    const calendarSheet = SpreadsheetApp.openById(scriptProps.getProperty('calendarSheet'))
        .getSheetByName(scriptProps.getProperty('calendarName'));
    const BUFFER = 20;
    const TRANIEES = [Array(5).fill('TRAINEES')];
    //Only need plus 1 because schedules include headers for the days which covers the two headers in the loop
    //Then only 1 is needed to push you past the last row
    const daySize = 1 + schedule.length + traineeSchedule.length;

    for(let i = 1; i <= calendar.length; i++)
    {
        let startIndex = i + (daySize * (i - 1));
        
        fillHeaders(calendarSheet, [calendar[i - 1]], 'black', 'white', startIndex, 1);

        let index = 1;
        fillWeekdaySchedule(calendarSheet, schedule, startIndex + index);
        index += schedule.length;

        fillHeaders(calendarSheet, TRANIEES, 'light_blue_2', 'black', startIndex + index, 2);
        index++;

        fillWeekdaySchedule(calendarSheet, traineeSchedule, startIndex + index);
    }

    calendarSheet.getDataRange().setHorizontalAlignment('center').setVerticalAlignment('middle');
    calendarSheet.autoResizeColumns(1, 7);

    let maxColSize = -1;
    for(let j = 1; j <= calendar.length; j++)
    {
        if(calendarSheet.getColumnWidth(j) > maxColSize)
        {
            maxColSize = calendarSheet.getColumnWidth(j);
        }
    }

    calendarSheet.setColumnWidths(1, 7, maxColSize + BUFFER);
}

/**
 * 
 * @param {*} sheet  
 * @param {*} header 
 * @param {*} background_color 
 * @param {*} text_color 
 * @param {*} start 
 * @param {*} height 
 * @param {*} width 
 */
function fillHeaders(sheet, header, background_color, text_color, row, col)
{
    range = sheet.getRange(row, col, header.length, header[0].length);
    range.setValues(header);
    range.setBackground(colorNameToHex(background_color));
    range.setFontColor(colorNameToHex(text_color));
}

function fillWeekdaySchedule(sheet, schedule, start)
{
    range = sheet.getRange(start, 2, schedule.length, schedule[0].length);
    range.setValues(schedule);
}

/**
 * Works under the assumption that the table starts with Mon
 * 
 * @param {String} id 
 * @return {Array} row, columnm, number of rows followed by number of columns in the table
 */
function getTableDimensions(id, name)
{
  console.log(id, " ", name)
    const sheet = SpreadsheetApp.openById(id).getSheetByName(name);
    const dataRange = sheet.getDataRange()
    const find = dataRange.createTextFinder('Mon')
        .matchCase(false)
        .matchEntireCell(true);

    const foundCell = find.findNext();

    if(!foundCell)
    {
        return [-1, -1];
    }

    var row = foundCell.getRow();
    var col = foundCell.getColumn();
    var maxRow = sheet.getMaxRows();
    var maxCol = sheet.getMaxColumns();

    const merged = dataRange.getMergedRanges();

    var rowLength = 0;
    var colLength = 0;

    while((row + rowLength <= maxRow && 
        sheet.getRange(row + rowLength, col).getBackground() != '#ffffff') || 
        mergedAt(merged, row + rowLength, col))
    {
        rowLength += 1;
    }

    while((col + colLength <= maxCol && 
        sheet.getRange(row, col + colLength).getBackground() != '#ffffff') || 
        mergedAt(merged, row, col + colLength))
    {
        colLength += 1;
    }

    return [row, col, rowLength, colLength];
}

function updateSheetNames(chanel, name) {
    if(chanel == "Instructor")
    {
        scriptProps.setProperties({
          'instructorSchedule': name
        });
    }
    else if (chanel == "Trainee")
    {
        scriptProps.setProperties({
          'traineeSchedule': name
        });
    }
    else
    {
        throw new Error('Unexpected file name "' + name + '": expected it to contain "Instructor" or "Trainee".');
    }
    
}

// ---------------------------- HELPER FUNCTIONS ---------------------------- 

/**
 * Function takes a name and truncates the first name to fit in the schedule
 * @param {string} name - string of an instructors name
 * @return {string} Truncated name in the form of first inital lastname (J Doe)
 */
function transformName(name)
{
    if(name.length == 0)
    {
        return '';
    }

    var end = name.indexOf(' ', name.indexOf(' ') + 1);
    if(end == -1)
    {
        end = name.length;
    }

    return name.slice(0, end);
}

/**
 * Function takes a time and truncates it by removing the AM/PM and 'to'
 * @param {string} time - string of a time slot
 * @return {string} Truncated time in the form of XX:XX-XX:XX
 */
function transformTime(time)
{
    if(time == undefined)
    {
        return '';
    }

    if(time.length == 0)
    {
        return '';
    }

    let startTime = time.slice(0, time.indexOf(' '));
    let endTime = "";
    
    for(let i = time.indexOf(' ') + 1; i < time.length; i++)
    {
        if(/^\d$/.test(time[i]))
        {
            endTime = time.slice(i, time.indexOf(' ', i));
            break;
        }
    }

    return startTime + "-" + endTime;
}

/**
 * 
 * @param {Range} merged 
 * @param {int} row 
 * @param {int} col 
 * @returns {boolean}
 */
function mergedAt(merged, row, col)
{
    var found = merged.find(range => 
        row >= range.getRow() && row < range.getRow() + range.getNumRows() &&
        col >= range.getColumn() && col < range.getColumn() + range.getNumColumns());
    
    return found !== undefined;
}

/**
 * 
 * @param {*} name 
 * @returns 
 */
function removeStars(name)
{
    while(name.indexOf('*') != -1)
    {
        name = name.slice(name.indexOf('*') + 1);
    }

    return name;
}

/**
 * 
 * @param {*} data 
 * @returns 
 */
function checkTime(data)
{
    if(data.indexOf('(') != -1)
    {
        var time = data.slice(data.indexOf('(') + 1, data.indexOf(')'));

        return time;
    }
    else
    {
        return null;
    }
}

/**
 * GENERATED WITH GEMINI
 * White is the default color
 * Converts a standard Google color name to its corresponding Hex code.
 * @param {string} colorName The name of the color (case-insensitive).
 * @return {string} The hex color code or a fallback if not found.
 */
function colorNameToHex(colorName) {
  // Normalize the input name
  const name = colorName.toLowerCase().replace(/\s+/g, '_');

  // Google Workspace Standard Color Map
  const colorMap = {
    // Grayscale
    'black': '#000000',
    'white': '#ffffff',
    'dark_gray_4': '#434343',
    'dark_gray_3': '#666666',
    'dark_gray_2': '#999999',
    'dark_gray_1': '#b7b7b7',
    'gray': '#cccccc',
    'light_gray_1': '#d9d9d9',
    'light_gray_2': '#efefef',
    'light_gray_3': '#f3f3f3',

    // Standard Chromatic Tones
    'red': '#ff0000',
    'orange': '#ff9900',
    'yellow': '#ffff00',
    'green': '#00ff00',
    'cyan': '#00ffff',
    'cornflower_blue': '#4a86e8',
    'blue': '#0000ff',
    'purple': '#9900ff',
    'magenta': '#ff00ff',
    'red_berry': '#990000',

    // Common Light/Dark Variants
    'light_red_3': '#f4cccc',
    'light_red_2': '#ea9999',
    'light_red_1': '#e06666',
    'dark_red_1': '#cc0000',
    'dark_red_2': '#990000',
    'dark_red_3': '#660000',

    'light_green_3': '#d9ead3',
    'light_green_2': '#b6d7a8',
    'light_green_1': '#93c47d',
    'dark_green_1': '#6aa84f',
    'dark_green_2': '#38761d',
    'dark_green_3': '#274e13',

    'light_blue_3': '#c9daf8',
    'light_blue_2': '#a4c2f4',
    'light_blue_1': '#6d9eeb',
    'dark_blue_1': '#3c78d8',
    'dark_blue_2': '#1155cc',
    'dark_blue_3': '#1c4587'
  };

  // Return the hex code if found, otherwise return a default fallback (e.g., black)
  return colorMap[name] || '#000000'; 
}

// console.log(buildCalendar(2026, 8));
// console.log(transformName("John Doe"));
// console.log(transformTime('3:45 PM to 6:00 PM'));

// Structure DATA
// [ [ 'Mon', '', 'Tue', '', 'Wed', '', 'Thu', '', 'Fri', '' ],
//   [ '3:45 PM to 6:00 PM','Instructor One','3:45 PM to 6:00 PM','Instructor Five','3:45 PM to 6:00 PM','Instructor Four','3:45 PM to 6:00 PM','Instructor Eight','4:30 PM to 6:45 PM','' ],
//   [ '4:30 PM to 6:45 PM','Instructor Two','4:30 PM to 6:45 PM','Instructor Six','4:30 PM to 6:45 PM','Instructor Seven','','','4:30 PM to 7:30 PM','Instructor Three' ],
//   [ '5:15 PM to 7:30 PM','Instructor Three','5:15 PM to 7:30 PM','Instructor Three','5:15 PM to 7:30 PM','Instructor Five','5:15 PM to 7:30 PM','Instructor Six','5:15 PM to 7:30 PM','Instructor One' ],
//   [ '','Instructor Four','','Instructor two','','','','Instructor Seven','','Instructor Eight' ],
//   [ '', '', '', '', '', '', '', '', '', '' ] ]

// Structure of Schedule