//Getting Data from Emails using the Subject and Sender
 function getEmailData() {
  var threads = GmailApp.search('subject:"Thirukkural Mutrodhal 2025 " from:info@signupgenius.com');

  if (threads.length === 0) {
    Logger.log('No matching threads found.');
    return;
  }

  // Get the Google Sheet
  var sheet = getSheet();

  threads.forEach(function (thread) {
    var messages = thread.getMessages();

    messages.forEach(function (message) {
      var dateReceived = message.getDate();
      var body = message.getBody();
      // Parse Adhigaram details and Comments
      var adhigaramDetails = parseAllAdhigaramDetails(body);
      var comments = parseAllMyComments(body);

      // Call the parseAndUpdateSheet function to process data
      parseAndUpdateSheet(sheet, dateReceived, adhigaramDetails, comments);
    });
  });
}
//Parsing the Emails to get Date, Adhigaram No, Name, Contact Name and Number
function parseAndUpdateSheet(sheet, dateReceived, adhigaramDetails, comments) {
    var existingData = getExistingSheetData(sheet);
    Logger.log("existingData: " + existingData);

    adhigaramDetails.forEach(function (adhigaram, index) {
    var comment = comments[index] || { name: "No name found", phoneNumber: "No number found" };

    // Clean up contact name and phone number
    comment.name = cleanContactName(comment.name, adhigaram.name);
         Logger.log("Clean Comment name " + comment.name);

    // Create a unique key for the row
    var rowKey = `${adhigaram.number}`;
    if (!existingData.has(rowKey)) {

      sheet.appendRow([adhigaram.number, adhigaram.name, comment.name, comment.phoneNumber, dateReceived]);
      existingData.add(rowKey);
    // Sort the sheet based on "Adhigaram number" (assuming it's in column A, adjust if needed)
    // Get the range starting from the second row to avoid the header row
    var range = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());  // Start from row 2, column 1
    range.sort({column: 1, ascending: true}); // Sort by column B (2), change if necessary

    } else {
      Logger.log("Duplicate entry found, skipping: " + rowKey);
    }
  });
}

// Function to get or create the 'Signups' Google Sheet
function getSheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('Signups');

  if (!sheet) {
    sheet = spreadsheet.insertSheet('Signups');
    sheet.appendRow(['Adhigaram Number', 'Adhigaram Name', 'Contact Name', 'Contact Number', 'Date Received']);
  }

  return sheet;
}

// Function to get all existing data from the sheet as a Set
function getExistingSheetData(sheet) {
  var data = sheet.getDataRange().getValues();
  var existingData = new Set();

  for (var i = 1; i < data.length; i++) {
    var row = data[i];

    var rowKey = `${row[0]}`;
    Logger.log("rowKey"+rowKey);
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

//Parse My Comment to retrieve the Contact name and Contact number
function parseAllMyComments(body) {
  var comments = [];

  // Updated regex to capture name, email, and phone numbers
  var regex = /My\s*Comment[:\s]*(?:(?:Name:\s*)?([A-Za-z]+(?:\s+[A-Za-z]+)*))?[\s,;:\/-]*(?:(?:email[:\s]*)?([\w.\-]+@[a-zA-Z_]+?\.[a-zA-Z]{2,6}))?[\s,\/-]*(?:(?:phone|number|contact|[:\s]*)?((?:\+?\d{1,4}[\s-]?)?(?:\d{3}[-\s]?\d{3}[-\s]?\d{4})(?:[\s,]+)?(?:\+?\d{1,4}[\s-]?\d{3}[-\s]?\d{3}[-\s]?\d{4})?))?/g;

  var match;

  while ((match = regex.exec(body)) !== null) {
    Logger.log("Inside while loop");

    // Extract name, email, and phone
    var contactName = match[1] ? match[1].trim() : (match[5] ? match[5].split(',')[0].trim() : 'No name found'); // Group 1: Name or from group 5
    var contactEmail = match[2] ? match[2].trim() : null;           // Group 2: Email
    var phoneNumbers = match[3] ? match[3].trim() : null;           // Group 3: Phone (multiple possible)

    Logger.log("contactName = " + contactName + ", contactEmail = " + contactEmail + ", phoneNumbers = " + phoneNumbers);

    // If no name found, it should be defaulted
    if (!contactName) {
      contactName = 'No name found';
    }

    // Store multiple phone numbers if found
    var contactNumbers = [];
    
    if (phoneNumbers) {
      // Split multiple phone numbers (if any) and clean them
      var phoneArray = phoneNumbers.split(/[\s,]+/); // Split by space or comma
      phoneArray.forEach(function(number) {
        // Clean each number (remove unwanted characters)
        var cleanedNumber = validatePhoneNumber(number.trim());
        if (cleanedNumber) {
          contactNumbers.push(cleanedNumber); // Add to the contact numbers list
        }
      });
    }

    // If no phone number is found, use email as contact number
    if (contactNumbers.length === 0 && contactEmail) {
      contactNumbers.push(contactEmail);
    }

    // If no phone numbers or email are found, set "No number found"
    if (contactNumbers.length === 0) {
      contactNumbers.push('No number found');
    }

    // Add parsed comment details to the list
    comments.push({
      name: contactName,
      phoneNumber: contactNumbers.join(', ') // Join multiple numbers if present
    });
  }

  // Default entry if no matches found
  if (comments.length === 0) {
    comments.push({
      name: 'No name found',
      phoneNumber: 'No number found'
    });
  }

  return comments;
}

// Helper function to validate phone numbers
function validatePhoneNumber(phoneNumber) {
  // Remove any non-digit characters (for simplicity, you can extend validation for country codes, etc.)
  var cleanedNumber = phoneNumber.replace(/\D/g, "");
  if (cleanedNumber.length >= 10) {  // Basic check for length
    return cleanedNumber; 
  }
  return null;  // Return null if the number is not valid
}

function cleanContactName(contactName, adhigaramName) {
  // Split both Adhigaram name and contact name into parts using spaces, hyphens, etc.
  var adhigaramParts = adhigaramName.split(/[\s\-]+/); // Splits by space or hyphen
  var contactParts = contactName.split(/[\s\-]+/); // Splits by space or hyphen

  // Refined logic to avoid matching partial words in the name itself
  var isSubstringMatch = false;
  for (var i = 0; i < adhigaramParts.length; i++) {
    for (var j = 0; j < contactParts.length; j++) {
      Logger.log("adhigaramParts["+i+"]"+adhigaramParts[i]);
      Logger.log("contactParts["+j+"]"+contactParts[j]);

      if (adhigaramParts[i].includes(contactParts[j]) || contactParts[j].includes(adhigaramParts[i])) {
        // If exact match, we can flag it and return "No name found"
        isSubstringMatch = true;
        break;
      }
    }
    if (isSubstringMatch) break;
  }
  // If there's an exact match in the name parts, return "No name found"
  if (isSubstringMatch) {
    return "No name found";
  }
  Logger.log("Name in cleanContactName " + contactName);

  // Clean the contact name by removing unwanted terms and trimming it
  var cleanedName = contactName
    .replace(/\s*-\s*$/, "") // Remove trailing hyphens
    .replace(/\b(Son|Mob No|Daughter|I will recite|Adhigaram|Will take this Adhikaram|will join)\b/g, "") // Remove unwanted terms
    .trim();
  
  Logger.log("Name after cleaning in cleanContactName " + cleanedName);

  // Return cleaned name, or default if it's empty
  return cleanedName || "No name found";
}