function pullData() {
    // Consider giving the sheet that lists the availability a better name
    const sourceSheet = SpreadsheetApp.openById('1tlSdMngNsnmXRxxgbEIFBlbQ6UphVphRQHQEW2wtpIs').getSheetByName("Sheet1");
    const data = sourceSheet.getDataRange().getValues();

    console.log(data);
}

pullData();