function doPost(e){
    try {
        // Parse the incoming JSON data from the Python bot
        var data = JSON.parse(e.postData.contents);

        // Example: Log the data or write it to a Google Sheet
        var username = data.username;
        var message = data.message;

        // Return a success response back to the Python bot
        return ContentService.createTextOutput(JSON.stringify({ "status": "success", "username": username, "message": message }))
                                .setMimeType(ContentService.MimeType.JSON);
                            
    } catch (error) {
        return ContentService.createTextOutput(JSON.stringify({ "status": "error", "message": error.toString() }))
                                .setMimeType(ContentService.MimeType.JSON);
    }
}