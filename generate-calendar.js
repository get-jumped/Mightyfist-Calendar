function pullData() {
    const monthNames = [
        "January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December"
    ];

    // Consider giving the sheet that lists the availability a better name
    const sourceSheet = SpreadsheetApp.openById('1tlSdMngNsnmXRxxgbEIFBlbQ6UphVphRQHQEW2wtpIs').getSheetByName("Sheet1");
    let data = sourceSheet.getRange(9, 2, 6, 10).getValues();

    schedule = fillMerges(data);
    console.log(schedule);

    const date = new Date();
    const year = date.getFullYear();
    const month = date.getMonth();

    let calendar = buildCalendar(year, month);
    console.log(calendar);

    fillSpreadSheet(calendar, schedule);
}

// Fills in missing times due to merged cells (only fills in missing times if 
// someone signed up for the timeslot. Reason to avoid confusion with cells that
// have no timeslots)
// INPUT: data - a 2D array that represents the schedule
// OUTPUT: newSchedule - a 2D array that is in the format we want to put into the calender spreadsheet
function fillMerges(data) {
    let newSchedule =[];

    for(let i = 1; i < data.length; i++)
    {
        let temp = [];
        for(let j = 1; j < data[i].length; j += 2)
        {
            let name = data[i][j];
            let time = "";

            // IF there is a name but no time with it
            if(data[i][j] != '' && data[i][j - 1] == '')
            {
                time = data[i - 1][j - 1];
            }
            // IF there is a name and time
            else
            {
                time = data[i][j - 1];
            }

            temp.push(transformName(name) + " " + transformTime(time));
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

// INPUT: calendar - a 2D array that represents the current month
//        schedule - a 2D array that represents the schedule of instructors  
function fillSpreadSheet(calendar, schedule) {
    const calendarSheet = SpreadsheetApp.openById('1E_nqdi1ZDexe-v64TkuQq_Th8ScHAq6-KZlFiyxEKjI').getSheetByName("Sheet1");

    for(let i = 1; i <= calendar.length; i++)
    {
        let startIndex = i + (6 * (i - 1));
        calendarSheet.getRange(startIndex, 1, 1, 7).setValues([calendar[i - 1]]);
        calendarSheet.getRange(startIndex + 1, 2, 5, 5).setValues(schedule);
    }

    calendarSheet.getDataRange().setHorizontalAlignment('center').setVerticalAlignment('middle');
}



// ---------------------------- HELPER FUNCTIONS ---------------------------- 

// Function takes a name and truncates the first name to fit in the schedule
// INPUT: name - string of an instructors name
// OUTPUT: truncated name in the form of first inital lastname (J Doe)
function transformName(name)
{
    if(name.length == 0)
    {
        return '';
    }

    return name[0] + " " + name.slice(name.indexOf(' ') + 1);
}

// Function takes a time and truncates it by removing the AM/PM and 'to'
// INPUT: time - string of a time slot
// OUTPUT: truncated time in the form of XX:XX-XX:XX
function transformTime(time)
{
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