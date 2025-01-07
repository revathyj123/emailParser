/*function getEmailData() {
  // Search for emails with the specific subject and sender
  var threads = GmailApp.search('subject:"Thirukkural Mutrodhal 2025" from:info@signupgenius.com');

  if (threads.length === 0) {
    Logger.log('No matching threads found.');
    return;
  }

  Logger.log("Number of threads: " + threads.length);

  // Get the Google Sheet
  var sheet = getSheet();
  var existingData = getExistingSheetData(sheet);

  // Loop through threads
  threads.forEach(function (thread) {
    var messages = thread.getMessages();

    messages.forEach(function (message) {
      var dateReceived = message.getDate();
      var body = message.getBody();
      Logger.log("Processing email received on: " + dateReceived);

      // Parse Adhigaram details and Comments
      var adhigaramDetails = parseAllAdhigaramDetails(body);
      var comments = parseAllMyComments(body);

      // Ensure each Adhigaram is associated with a Comment
      adhigaramDetails.forEach(function (adhigaram, index) {
        var comment = comments[index] || { name: "No name found", phoneNumber: "No number found" };

        // Clean up Contact Name
        comment.name = cleanContactName(comment.name);

        // Create a unique key for the row
        var rowKey = `${dateReceived}-${adhigaram.number}-${adhigaram.name}-${comment.name}`;

        // Add to sheet if not already present
        if (!existingData.has(rowKey)) {
          Logger.log("Adding new row: " + rowKey);
          sheet.appendRow([dateReceived, adhigaram.number, adhigaram.name, comment.name, comment.phoneNumber]);
          existingData.add(rowKey);
        } else {
          Logger.log("Duplicate entry found, skipping: " + rowKey);
        }
      });
    });
  });
}*/

function getEmailData() {
    // Search for emails with the specific subject and sender
    var threads = GmailApp.search('subject:"Thirukkural Mutrodhal 2025" from:info@signupgenius.com');
  
    if (threads.length === 0) {
      Logger.log('No matching threads found.');
      return;
    }
  
    Logger.log("Number of threads: " + threads.length);
  
    // Get the Google Sheet
    var sheet = getSheet();
    var existingData = getExistingSheetData(sheet);
  
    // Loop through threads
    threads.forEach(function (thread) {
      var messages = thread.getMessages();
  
      messages.forEach(function (message) {
        var dateReceived = message.getDate();
        var body = message.getBody();
        Logger.log("Processing email received on: " + dateReceived);
  
        // Parse Adhigaram details and Comments
        var adhigaramDetails = parseAllAdhigaramDetails(body);
        var comments = parseAllMyComments(body);
  
        // Ensure each Adhigaram is associated with a Comment
        adhigaramDetails.forEach(function (adhigaram, index) {
          var comment = comments[index] || { name: "No name found", phoneNumber: "No number found" };
  
          // Clean up Contact Name
          comment.name = cleanContactName(comment.name);
  
          // Create a unique key for the row
          var rowKey = `${adhigaram.number}-${adhigaram.name}`;  // Only use Adhigaram number and name for uniqueness
  
          // Add to sheet if not already present
          if (!existingData.has(rowKey)) {
            Logger.log("Adding new row: " + rowKey);
            sheet.appendRow([dateReceived, adhigaram.number, adhigaram.name, comment.name, comment.phoneNumber]);
            existingData.add(rowKey);
          } else {
            Logger.log("Duplicate entry found, skipping: " + rowKey);
          }
        });
      });
    });
  }
  
  
  // Function to get or create the 'Signups' Google Sheet
  function getSheet() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName('Signups');
  
    if (!sheet) {
      sheet = spreadsheet.insertSheet('Signups');
      sheet.appendRow(['Date Received', 'Adhigaram Number', 'Adhigaram Name', 'Contact Name', 'Contact Number']);
    }
  
    return sheet;
  }
  
  // Function to get all existing data from the sheet as a Set
  function getExistingSheetData(sheet) {
    var data = sheet.getDataRange().getValues();
    var existingData = new Set();
  
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var rowKey = `${row[0]}-${row[1]}-${row[2]}-${row[3]}`;
      existingData.add(rowKey);
    }
  
    return existingData;
  }
  
  // Function to parse all Adhigaram details
  function parseAllAdhigaramDetails(body) {
    var adhigaramRegex = /(\d+)\)\s([A-Za-z\s\-]+)/g;
    var matches = [];
    var match;
  
    while ((match = adhigaramRegex.exec(body)) !== null) {
      matches.push({
        number: match[1],
        name: match[2].trim()
      });
    }
  
    return matches;
  }
  /*
  function parseAllMyComments(body) {
    var comments = [];
    
    // Updated regex to handle space between "My Comment:" and the name, plus other possible formats
    var regex = /My\s*Comment[:\s]*([A-Za-z\s]+)[\s\/\-\*,]*([\+?\d{1,3}[\s\-]?\(?\d{1,4}\)?[\s\-]?\d{1,3}[\s\-]?\d{1,3}|\d{10,15}])?/gi
  
    var match;
    while ((match = regex.exec(body)) !== null) {
      var contactName = match[1].trim();  // Trim any extra spaces from the name
      var contactNumber = match[2] || 'No number found';  // Default if no number found
      
      // Add to the list of comments
      comments.push({
        name: contactName,
        phoneNumber: contactNumber
      });
    }
  
    // If no matches are found, add a default entry for no name and no number
    if (comments.length === 0) {
      comments.push({
        name: 'No name found',
        phoneNumber: 'No number found'
      });
    }
  
    return comments;
  }*/
  
  function parseAllMyComments(body) {
    var comments = [];
    
    // Updated regex to handle space between "My Comment:" and the name, plus other possible formats
    var regex = /My\s*Comment[:\s]*([A-Za-z\s]+)[\s\/\-\*,]*([\+?\d{1,3}[\s\-]?\(?\d{1,4}\)?[\s\-]?\d{1,3}[\s\-]?\d{1,3}|\d{10,15}])?/gi;
  
    var match;
    while ((match = regex.exec(body)) !== null) {
      var contactName = match[1].trim();  // Trim any extra spaces from the name
      var contactNumber = match[2] || 'No number found';  // Default if no number found
      
      // Add to the list of comments
      comments.push({
        name: contactName,
        phoneNumber: contactNumber
      });
    }
  
    // If no matches are found, add a default entry for no name and no number
    if (comments.length === 0) {
      comments.push({
        name: 'No name found',
        phoneNumber: 'No number found'
      });
    }
  
    return comments;
  }
  
  
  
  // Function to validate and clean phone numbers
  function validatePhoneNumber(phoneNumber) {
    // Remove non-numeric characters
    var cleanedNumber = phoneNumber.replace(/\D/g, "");
  
    // Ensure number length is between 10 and 15 digits
    return cleanedNumber.length >= 10 && cleanedNumber.length <= 15 ? cleanedNumber : "No number found";
  }
  
  // Function to clean contact name
  function cleanContactName(contactName) {
    // Remove trailing hyphens, unwanted terms, and extra spaces
    var cleanedName = contactName
      .replace(/\s*-\s*$/, "") // Remove trailing hyphens
      .replace(/\b(Son|Mob No|WhatsApp|Daughter)\b/gi, "") // Remove unwanted terms
      .trim();
  
    return cleanedName || "No name found"; // Return default if cleaned name is empty
  }