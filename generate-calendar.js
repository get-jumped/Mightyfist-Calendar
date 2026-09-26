function pullData() {
    // Consider giving the sheet that lists the availability a better name
    const sourceSheet = SpreadsheetApp.openById('1tlSdMngNsnmXRxxgbEIFBlbQ6UphVphRQHQEW2wtpIs').getSheetByName("Sheet1");
    let data = sourceSheet.getRange(9, 2, 6, 10).getValues();

    data = fillMerges(data);
    console.log(data);
}

// Fills in missing times due to merged cells (only fills in missing times if 
// someone signed up for the timeslot. Reason to avoid confusion with cells that
// have no timeslots)
// INPUT: schedule is a 2D array that represents the schedule
function fillMerges(schedule) {

    for(let i = 1; i < schedule.length; i++)
    {
        for(let j = 1; j < schedule[i].length; j += 2)
        {
            if(schedule[i][j] != '' && schedule[i][j - 1] == '')
            {
                schedule[i][j - 1] = schedule[i - 1][j - 1];
            }
        }
    }

    return schedule
}

pullData();

// Structure of recieved data
// [ [ 'Mon', '', 'Tue', '', 'Wed', '', 'Thu', '', 'Fri', '' ],
//   [ '3:45 PM to 6:00 PM','Instructor One','3:45 PM to 6:00 PM','Instructor Five','3:45 PM to 6:00 PM','Instructor Four','3:45 PM to 6:00 PM','Instructor Eight','4:30 PM to 6:45 PM','' ],
//   [ '4:30 PM to 6:45 PM','Instructor Two','4:30 PM to 6:45 PM','Instructor Six','4:30 PM to 6:45 PM','Instructor Seven','','','4:30 PM to 7:30 PM','Instructor Three' ],
//   [ '5:15 PM to 7:30 PM','Instructor Three','5:15 PM to 7:30 PM','Instructor Three','5:15 PM to 7:30 PM','Instructor Five','5:15 PM to 7:30 PM','Instructor Six','5:15 PM to 7:30 PM','Instructor One' ],
//   [ '','Instructor Four','','Instructor two','','','','Instructor Seven','','Instructor Eight' ],
//   [ '', '', '', '', '', '', '', '', '', '' ] ]