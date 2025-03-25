/**
 * Bubble RSVP Evite System - Complete Backend
 * A fully rewritten backend that handles data across multiple sheets with
 * proper validation, formatting, and row structure.
 */

// ==========================================================================
// CONFIGURATION - CUSTOMIZE THESE VALUES
// ==========================================================================
const CONFIG = {
  // Sheet names
  SHEET_NAME: "RSVPs",
  GUEST_LIST_SHEET_NAME: "Guest List",
  DIETARY_SHEET_NAME: "Dietary Information",
  
  // Visual elements
  APP_NAME: "Bubble RSVP Evite",
  BACKGROUND_IMAGE: "https://i.imgur.com/4pCJnLn.jpeg",
  
  // Dropdown values
  ATTENDANCE_VALUES: ["Y", "N", "Maybe"],
  STATUS_VALUES: ["Pending", "Confirmed", "Cancelled"],
  
  // UI elements
  HEADER_COLOR: "#4285F4",
  TITLE_ROW_COLOR: "#E8F5FE",
  CONFIRMED_COLOR: "#e6ffe6",  // Only used for RSVP sheet status column
  CANCELLED_COLOR: "#ffecec",  // Only used for RSVP sheet status column
  
  // Row structure
  HEADER_ROW: 1,
  TITLE_ROW: 2,
  DATA_START_ROW: 3,
  
  // Fixed text
  RSVP_TITLE_TEXT: "RSVP Submissions - Newest First",
  GUEST_LIST_TITLE_TEXT: "Guest List - Confirmed & Pending RSVPs",
  DIETARY_TITLE_TEXT: "Dietary Restrictions and Special Requests",
  
  // Event Details (customize these for your event)
  EVENT: {
    TITLE: "Callie's Birthday Celebration",
    DATE: "Saturday, April 28, 2025",
    TIME: "12:00 PM - 2:00 PM",
    LOCATION: "Kenwood Baptist Church - Pavillion",
    LOCATION_LINK: "https://maps.app.goo.gl/eqZJb2o5WgAAMCZb7",
    DESCRIPTION: "We are parked towards the back of the building (outside) in the pavillion area!",
    GIFT_INFO: "No gifts are necessary, just your presence is required! If you insist on bringing Callie a gift, clothes are preferred because God has blessed her with plenty of toys already!",
    DRESS_CODE: "Casual outdoor attire",
    ADDITIONAL_INFO: "We'll have a thematic race and other games so bring your game faces!"
  }
};

// ==========================================================================
// INITIALIZATION AND SETUP
// ==========================================================================

/**
 * Add menu item when spreadsheet is opened
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('RSVP Tools')
    .addItem('Confirm Selected RSVPs', 'confirmSelectedRsvps')
    .addItem('Create RSVP Manager Panel', 'showRsvpManager')
    .addSeparator()
    .addItem('Refresh RSVP Status', 'refreshRsvpStatus')
    .addItem('Add Form Validation', 'addFormValidation')
    .addItem('Ensure All Sheets Exist', 'ensureAllSheetsExist')
    .addSeparator()
    .addItem('Regenerate Dietary Information', 'regenerateDietaryInformation')
    .addItem('Regenerate Guest List', 'regenerateGuestList')
    .addItem('Regenerate All Sheets', 'regenerateAllSheets')
    .addToUi();
  
  // Automatically ensure all sheets exist when the spreadsheet is opened
  ensureAllSheetsExist();
  
  // Add form validation automatically
  addFormValidation();
}

/**
 * Ensures all required sheets exist and are properly formatted
 */
function ensureAllSheetsExist() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Create or get the main RSVP sheet
    let rsvpSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (!rsvpSheet) {
      console.log(`Creating main RSVP sheet: ${CONFIG.SHEET_NAME}`);
      rsvpSheet = ss.insertSheet(CONFIG.SHEET_NAME);
    }
    formatRsvpSheet(rsvpSheet);
    
    // Create or get the Guest List sheet
    let guestListSheet = ss.getSheetByName(CONFIG.GUEST_LIST_SHEET_NAME);
    if (!guestListSheet) {
      console.log(`Creating guest list sheet: ${CONFIG.GUEST_LIST_SHEET_NAME}`);
      guestListSheet = ss.insertSheet(CONFIG.GUEST_LIST_SHEET_NAME);
    }
    formatGuestListSheet(guestListSheet);
    
    // Create or get the Dietary Information sheet
    let dietarySheet = ss.getSheetByName(CONFIG.DIETARY_SHEET_NAME);
    if (!dietarySheet) {
      // Try common variations
      const variations = [
        "Dietary Information", "Dietary Info", "Dietary Restrictions", 
        "Dietary", "Diet Info", "Diet Restrictions"
      ];
      
      for (const variation of variations) {
        if (variation !== CONFIG.DIETARY_SHEET_NAME) {
          console.log(`Trying alternative name: "${variation}"`);
          dietarySheet = ss.getSheetByName(variation);
          if (dietarySheet) {
            console.log(`Found dietary sheet with name: "${variation}"`);
            break;
          }
        }
      }
      
      // If still not found, create it
      if (!dietarySheet) {
        console.log(`Creating dietary information sheet: ${CONFIG.DIETARY_SHEET_NAME}`);
        dietarySheet = ss.insertSheet(CONFIG.DIETARY_SHEET_NAME);
      }
    }
    formatDietarySheet(dietarySheet);
    
    return true;
  } catch (error) {
    console.error("Error ensuring sheets exist: " + error.toString());
    SpreadsheetApp.getUi().alert("Error ensuring sheets exist: " + error.toString());
    return false;
  }
}

// ==========================================================================
// SHEET FORMATTING FUNCTIONS
// ==========================================================================

/**
 * Format the main RSVP sheet with headers and structure
 */
function formatRsvpSheet(sheet) {
  try {
    // Define the headers for the RSVP sheet
    const headers = [
      "Timestamp",
      "Name",
      "Email", 
      "Phone",
      "Attending",
      "Number of Guests",
      "Guest Names",
      "Dietary Restrictions",
      "Comments",
      "RSVP Date",
      "Follow-up Sent",
      "Status"
    ];
    
    // Get existing headers if any
    let existingHeaders = [];
    if (sheet.getLastRow() > 0) {
      existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    }
    
    // Check if we need to add/update headers
    const needsHeaders = sheet.getLastRow() === 0 || existingHeaders.length < headers.length;
    
    if (needsHeaders) {
      // Add header row
      sheet.getRange(CONFIG.HEADER_ROW, 1, 1, headers.length).setValues([headers]);
      formatHeaderRow(sheet, headers.length);
      
      // Add title row if it doesn't exist or if we're recreating headers
      if (sheet.getLastRow() < CONFIG.TITLE_ROW) {
        sheet.insertRowAfter(CONFIG.HEADER_ROW);
      }
      formatTitleRow(sheet, CONFIG.RSVP_TITLE_TEXT, headers.length);
      
      // Set column widths
      sheet.setColumnWidth(1, 180);  // Timestamp
      sheet.setColumnWidth(2, 150);  // Name
      sheet.setColumnWidth(3, 200);  // Email
      sheet.setColumnWidth(4, 120);  // Phone
      sheet.setColumnWidth(5, 100);  // Attending
      sheet.setColumnWidth(6, 100);  // Number of Guests
      sheet.setColumnWidth(7, 200);  // Guest Names
      sheet.setColumnWidth(8, 200);  // Dietary Restrictions
      sheet.setColumnWidth(9, 300);  // Comments
      sheet.setColumnWidth(10, 180); // RSVP Date
      sheet.setColumnWidth(11, 120); // Follow-up Sent
      sheet.setColumnWidth(12, 120); // Status
    }
    
    // Ensure we have proper freezing
    sheet.setFrozenRows(CONFIG.TITLE_ROW);
    
    return sheet;
  } catch (error) {
    console.error("Error formatting RSVP sheet: " + error.toString());
    return sheet;
  }
}

/**
 * Format the Guest List sheet with headers and structure
 */
function formatGuestListSheet(sheet) {
  try {
    // Define the headers for the Guest List sheet
    const headers = [
      "Name",
      "Email",
      "Phone",
      "Attending",
      "Number of Guests",
      "Guest Names",
      "Status",
      "RSVP Date"
    ];
    
    // Check if we need to add headers
    if (sheet.getLastRow() === 0) {
      // Add header row
      sheet.getRange(CONFIG.HEADER_ROW, 1, 1, headers.length).setValues([headers]);
      formatHeaderRow(sheet, headers.length);
      
      // Add title row
      sheet.insertRowAfter(CONFIG.HEADER_ROW);
      formatTitleRow(sheet, CONFIG.GUEST_LIST_TITLE_TEXT, headers.length);
      
      // Set column widths
      sheet.setColumnWidth(1, 150);  // Name
      sheet.setColumnWidth(2, 200);  // Email
      sheet.setColumnWidth(3, 120);  // Phone
      sheet.setColumnWidth(4, 100);  // Attending
      sheet.setColumnWidth(5, 100);  // Number of Guests
      sheet.setColumnWidth(6, 200);  // Guest Names
      sheet.setColumnWidth(7, 120);  // Status
      sheet.setColumnWidth(8, 180);  // RSVP Date
    }
    
    // Ensure we have proper freezing
    sheet.setFrozenRows(CONFIG.TITLE_ROW);
    
    return sheet;
  } catch (error) {
    console.error("Error formatting Guest List sheet: " + error.toString());
    return sheet;
  }
}

/**
 * Format the Dietary Information sheet with headers and structure
 */
function formatDietarySheet(sheet) {
  try {
    // Define the headers for the Dietary Sheet
    const headers = [
      "Name",
      "Dietary Restrictions",
      "Number of Guests",
      "Last Updated"
    ];
    
    // Check if we need to add headers
    if (sheet.getLastRow() === 0) {
      // Add header row with main title
      sheet.getRange(CONFIG.HEADER_ROW, 1, 1, headers.length).setValues([[CONFIG.DIETARY_TITLE_TEXT, "", "", ""]]);
      sheet.getRange(CONFIG.HEADER_ROW, 1, 1, headers.length).merge();
      sheet.getRange(CONFIG.HEADER_ROW, 1).setBackground(CONFIG.HEADER_COLOR)
                                          .setFontColor("#FFFFFF")
                                          .setFontWeight("bold")
                                          .setHorizontalAlignment("center");
      
      // Add column headers in the second row
      sheet.getRange(CONFIG.TITLE_ROW, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(CONFIG.TITLE_ROW, 1, 1, headers.length).setBackground(CONFIG.TITLE_ROW_COLOR)
                                                           .setFontWeight("bold");
      
      // Set column widths
      sheet.setColumnWidth(1, 150);  // Name
      sheet.setColumnWidth(2, 300);  // Dietary Restrictions
      sheet.setColumnWidth(3, 100);  // Number of Guests
      sheet.setColumnWidth(4, 180);  // Last Updated
    }
    
    // Ensure we have proper freezing
    sheet.setFrozenRows(CONFIG.TITLE_ROW);
    
    return sheet;
  } catch (error) {
    console.error("Error formatting Dietary sheet: " + error.toString());
    return sheet;
  }
}

/**
 * Format a standard header row
 */
function formatHeaderRow(sheet, columnCount) {
  const headerRange = sheet.getRange(CONFIG.HEADER_ROW, 1, 1, columnCount);
  headerRange.setBackground(CONFIG.HEADER_COLOR)
             .setFontColor("#FFFFFF")
             .setFontWeight("bold");
}

/**
 * Format a standard title row
 */
function formatTitleRow(sheet, titleText, columnCount) {
  // Add title text
  sheet.getRange(CONFIG.TITLE_ROW, 1, 1, columnCount).setValues([[titleText].concat(Array(columnCount-1).fill(""))]);
  
  // Merge cells and format
  sheet.getRange(CONFIG.TITLE_ROW, 1, 1, columnCount).merge();
  sheet.getRange(CONFIG.TITLE_ROW, 1).setBackground(CONFIG.TITLE_ROW_COLOR)
                                    .setFontWeight("bold")
                                    .setHorizontalAlignment("center");
}

// ==========================================================================
// DATA VALIDATION FUNCTIONS
// ==========================================================================

/**
 * Add data validation (dropdowns) to all sheets
 */
function addFormValidation() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Add validation to RSVP sheet
    const rsvpSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (rsvpSheet) {
      addRsvpSheetValidation(rsvpSheet);
    }
    
    // Add validation to Guest List sheet
    const guestListSheet = ss.getSheetByName(CONFIG.GUEST_LIST_SHEET_NAME);
    if (guestListSheet) {
      addGuestListSheetValidation(guestListSheet);
    }
    
    console.log("Form validation has been added successfully");
  } catch (error) {
    console.error("Error adding form validation:", error);
    SpreadsheetApp.getUi().alert("Error adding form validation: " + error.toString());
  }
}

/**
 * Add validation to the RSVP sheet
 */
function addRsvpSheetValidation(sheet) {
  try {
    // Find the column indices
    const headers = sheet.getRange(CONFIG.HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
    const attendingColIndex = headers.indexOf("Attending") + 1;
    const statusColIndex = headers.indexOf("Status") + 1;
    
    if (attendingColIndex === 0 || statusColIndex === 0) {
      console.error("Could not find required columns in RSVP sheet");
      return;
    }
    
    // Create rules
    const attendingRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(CONFIG.ATTENDANCE_VALUES, true)
      .setAllowInvalid(false)
      .build();
    
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(CONFIG.STATUS_VALUES, true)
      .setAllowInvalid(false)
      .build();
    
    // Calculate range (from data start row to end of sheet)
    const totalRows = Math.max(sheet.getMaxRows() - CONFIG.DATA_START_ROW + 1, 1);
    
    // Apply validation to columns
    sheet.getRange(CONFIG.DATA_START_ROW, attendingColIndex, totalRows, 1).setDataValidation(attendingRule);
    sheet.getRange(CONFIG.DATA_START_ROW, statusColIndex, totalRows, 1).setDataValidation(statusRule);
    
  } catch (error) {
    console.error("Error adding validation to RSVP sheet:", error);
  }
}

/**
 * Add validation to the Guest List sheet
 */
function addGuestListSheetValidation(sheet) {
  try {
    // Find the column indices
    const headers = sheet.getRange(CONFIG.HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
    const attendingColIndex = headers.indexOf("Attending") + 1;
    const statusColIndex = headers.indexOf("Status") + 1;
    
    if (attendingColIndex === 0 || statusColIndex === 0) {
      console.error("Could not find required columns in Guest List sheet");
      return;
    }
    
    // Create rules
    const attendingRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(CONFIG.ATTENDANCE_VALUES, true)
      .setAllowInvalid(false)
      .build();
    
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(CONFIG.STATUS_VALUES, true)
      .setAllowInvalid(false)
      .build();
    
    // Calculate range (from data start row to end of sheet)
    const totalRows = Math.max(sheet.getMaxRows() - CONFIG.DATA_START_ROW + 1, 1);
    
    // Apply validation to columns
    sheet.getRange(CONFIG.DATA_START_ROW, attendingColIndex, totalRows, 1).setDataValidation(attendingRule);
    sheet.getRange(CONFIG.DATA_START_ROW, statusColIndex, totalRows, 1).setDataValidation(statusRule);
    
  } catch (error) {
    console.error("Error adding validation to Guest List sheet:", error);
  }
}

/**
 * Add validation to a single row in the RSVP sheet
 */
function addValidationToRow(sheet, rowIndex) {
  try {
    // Find the column indices
    const headers = sheet.getRange(CONFIG.HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
    const attendingColIndex = headers.indexOf("Attending") + 1;
    const statusColIndex = headers.indexOf("Status") + 1;
    
    if (attendingColIndex === 0 || statusColIndex === 0) {
      console.error("Could not find required columns for row validation");
      return;
    }
    
    // Create rules
    const attendingRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(CONFIG.ATTENDANCE_VALUES, true)
      .setAllowInvalid(false)
      .build();
    
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(CONFIG.STATUS_VALUES, true)
      .setAllowInvalid(false)
      .build();
    
    // Apply validation to the row
    sheet.getRange(rowIndex, attendingColIndex).setDataValidation(attendingRule);
    sheet.getRange(rowIndex, statusColIndex).setDataValidation(statusRule);
    
  } catch (error) {
    console.error("Error adding validation to row:", error);
  }
}

// ==========================================================================
// RSVP DATA HANDLING FUNCTIONS
// ==========================================================================

/**
 * Process a new RSVP submission or update an existing one
 */
function saveRSVP(data) {
  try {
    console.log("Received RSVP data:", JSON.stringify(data));
    
    // Validate required fields
    if (!data || !data.name || data.name.trim() === "") {
      throw new Error("Name is required");
    }
    
    // Get the spreadsheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Ensure all required sheets exist
    ensureAllSheetsExist();
    
    // Get the sheets
    const rsvpSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    const guestListSheet = ss.getSheetByName(CONFIG.GUEST_LIST_SHEET_NAME);
    const dietarySheet = ss.getSheetByName(CONFIG.DIETARY_SHEET_NAME);
    
    // Create timestamp for this submission
    const timestamp = new Date();
    
    // Handle the RSVP sheet update
    updateRsvpSheet(rsvpSheet, data, timestamp);
    
    // Update the Guest List sheet
    updateGuestListSheet(guestListSheet, data, timestamp);
    
    // Update the Dietary Information sheet if needed
    if (data.dietary && data.dietary.trim() !== "") {
      updateDietarySheet(dietarySheet, data, timestamp);
    }
    
    return {
      status: "success",
      message: "Thank you! Your RSVP has been recorded."
    };
  } catch (error) {
    console.error("Error saving RSVP:", error);
    return {
      status: "error",
      message: "There was a problem saving your RSVP: " + error.toString()
    };
  }
}

/**
 * Update the main RSVP sheet with the submission data
 */
function updateRsvpSheet(sheet, data, timestamp) {
  try {
    // Check if this person already has an RSVP
    const existingRowIndex = findExistingEntry(sheet, data.name);
    
    // Determine appropriate status
    let status;
    if (existingRowIndex > CONFIG.TITLE_ROW) {
      // Get existing status
      const headers = sheet.getRange(CONFIG.HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
      const statusColIndex = headers.indexOf("Status") + 1;
      const existingStatus = sheet.getRange(existingRowIndex, statusColIndex).getValue();
      
      // Update status based on attendance change
      if (existingStatus === "Confirmed" && data.attending === "N") {
        status = "Cancelled"; // Previously confirmed, now not attending
      } else if (existingStatus !== "Confirmed") {
        status = (data.attending === "Y" || data.attending === "Maybe") ? "Pending" : "Cancelled";
      } else {
        status = existingStatus; // Keep confirmed status
      }
    } else {
      // New entry - set initial status
      status = (data.attending === "Y" || data.attending === "Maybe") ? "Pending" : "Cancelled";
    }
    
    // Prepare row data
    const rowData = [
      timestamp,                     // Timestamp
      data.name || "",               // Name
      data.email || "",              // Email
      data.phone || "",              // Phone
      data.attending || "No Response", // Attending
      parseInt(data.guests || "1", 10), // Number of Guests
      data.guestNames || "",         // Guest Names
      data.dietary || "",            // Dietary Restrictions
      data.comments || "",           // Comments
      timestamp,                     // RSVP Date
      false,                         // Follow-up Sent
      status                         // Status
    ];
    
    // Update or insert the row
    if (existingRowIndex > CONFIG.TITLE_ROW) {
      // Update existing row
      sheet.getRange(existingRowIndex, 1, 1, rowData.length).setValues([rowData]);
      
      // Re-apply formatting
      formatRsvpRow(sheet, existingRowIndex, status);
    } else {
      // Insert new row at the data start position
      sheet.insertRowAfter(CONFIG.TITLE_ROW);
      
      // Add data to the new row
      sheet.getRange(CONFIG.DATA_START_ROW, 1, 1, rowData.length).setValues([rowData]);
      
      // Apply formatting
      formatRsvpRow(sheet, CONFIG.DATA_START_ROW, status);
      
      // Ensure validation on the new row
      addValidationToRow(sheet, CONFIG.DATA_START_ROW);
    }
    
    return true;
  } catch (error) {
    console.error("Error updating RSVP sheet:", error);
    throw error;
  }
}

/**
 * Apply formatting to a row in the RSVP sheet
 * Modified to only style the status column and maintain plain formatting elsewhere
 */
function formatRsvpRow(sheet, rowIndex, status) {
  try {
    // Format timestamp columns
    sheet.getRange(rowIndex, 1).setNumberFormat("MMM dd, yyyy HH:mm:ss"); // Timestamp
    sheet.getRange(rowIndex, 10).setNumberFormat("MMM dd, yyyy HH:mm:ss"); // RSVP Date
    
    // Add checkbox for follow-up column
    sheet.getRange(rowIndex, 11).insertCheckboxes();
    
    // Make sure text is black for readability and clear any existing formatting
    const columnCount = sheet.getLastColumn();
    sheet.getRange(rowIndex, 1, 1, columnCount).setFontColor("#000000")
                                               .setBackground(null)
                                               .setFontWeight(null);
    
    // Only apply color to the status column (column 12)
    if (status === "Confirmed") {
      sheet.getRange(rowIndex, 12).setBackground(CONFIG.CONFIRMED_COLOR);
    } else if (status === "Cancelled") {
      sheet.getRange(rowIndex, 12).setBackground(CONFIG.CANCELLED_COLOR);
    }
  } catch (error) {
    console.error("Error formatting RSVP row:", error);
  }
}

/**
 * Update the Guest List sheet with RSVP data
 */
function updateGuestListSheet(sheet, data, timestamp) {
  try {
    // Check if this person already exists in the guest list
    const existingRowIndex = findExistingEntry(sheet, data.name);
    
    // Handle non-attending guests differently
    if (data.attending !== "Y" && data.attending !== "Maybe") {
      if (existingRowIndex > CONFIG.TITLE_ROW) {
        // If previously in the list, update with Cancelled status
        const headers = sheet.getRange(CONFIG.HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
        const statusColIndex = headers.indexOf("Status") + 1;
        const attendingColIndex = headers.indexOf("Attending") + 1;
        
        // Get current status
        const currentStatus = sheet.getRange(existingRowIndex, statusColIndex).getValue();
        
        // Update status based on previous value
        const newStatus = (currentStatus === "Confirmed") ? 
          "Cancelled (Previously Confirmed)" : "Cancelled";
        
        // Update attending and status
        sheet.getRange(existingRowIndex, attendingColIndex).setValue(data.attending);
        sheet.getRange(existingRowIndex, statusColIndex).setValue(newStatus);
        
        // Update timestamp
        const dateColIndex = headers.indexOf("RSVP Date") + 1;
        if (dateColIndex > 0) {
          sheet.getRange(existingRowIndex, dateColIndex).setValue(timestamp);
          sheet.getRange(existingRowIndex, dateColIndex).setNumberFormat("MMM dd, yyyy HH:mm:ss");
        }
      }
      return;
    }
    
    // Determine status to use
    let status = "Pending";
    if (existingRowIndex > CONFIG.TITLE_ROW) {
      // Check if already confirmed
      const headers = sheet.getRange(CONFIG.HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
      const statusColIndex = headers.indexOf("Status") + 1;
      const currentStatus = sheet.getRange(existingRowIndex, statusColIndex).getValue();
      
      // Preserve Confirmed status
      if (currentStatus === "Confirmed") {
        status = "Confirmed";
      }
    }
    
    // Prepare guest list data
    const guestListData = [
      data.name || "",                // Name
      data.email || "",               // Email
      data.phone || "",               // Phone
      data.attending || "No Response", // Attending
      parseInt(data.guests || "1", 10), // Number of Guests
      data.guestNames || "",          // Guest Names
      status,                         // Status
      timestamp                       // RSVP Date
    ];
    
    // Update or insert the row
    if (existingRowIndex > CONFIG.TITLE_ROW) {
      // Update existing row
      sheet.getRange(existingRowIndex, 1, 1, guestListData.length).setValues([guestListData]);
      
      // Re-apply formatting
      formatGuestListRow(sheet, existingRowIndex);
    } else {
      // Insert new row at the data start position
      sheet.insertRowAfter(CONFIG.TITLE_ROW);
      
      // Add data to the new row
      sheet.getRange(CONFIG.DATA_START_ROW, 1, 1, guestListData.length).setValues([guestListData]);
      
      // Apply formatting
      formatGuestListRow(sheet, CONFIG.DATA_START_ROW);
      
      // Ensure validation on the new row
      addValidationToRow(sheet, CONFIG.DATA_START_ROW);
    }
    
    return true;
  } catch (error) {
    console.error("Error updating Guest List sheet:", error);
    // Continue execution even if there's an error
    return false;
  }
}

/**
 * Apply formatting to a row in the Guest List sheet
 * Modified to use plain formatting without backgrounds or special styling
 */
function formatGuestListRow(sheet, rowIndex) {
  try {
    // Format timestamp column
    sheet.getRange(rowIndex, 8).setNumberFormat("MMM dd, yyyy HH:mm:ss");
    
    // Make sure text is black for readability and clear any existing formatting
    const columnCount = sheet.getLastColumn();
    sheet.getRange(rowIndex, 1, 1, columnCount).setFontColor("#000000")
                                               .setBackground(null)
                                               .setFontWeight(null);
  } catch (error) {
    console.error("Error formatting Guest List row:", error);
  }
}

/**
 * Update the Dietary Information sheet with RSVP data
 */
function updateDietarySheet(sheet, data, timestamp) {
  try {
    // Only include guests who are attending or confirmed
    if (data.attending !== "Y" && data.attending !== "Maybe" && data.status !== "Confirmed") {
      return;
    }
    
    // Skip if no dietary info
    if (!data.dietary || data.dietary.trim() === "") {
      return;
    }
    
    // Check if this person already exists in the dietary list
    const existingRowIndex = findExistingDietaryEntry(sheet, data.name);
    
    // Prepare dietary data
    const dietaryData = [
      data.name || "",                   // Name
      data.dietary || "",                // Dietary Restrictions
      parseInt(data.guests || "1", 10),  // Number of Guests
      timestamp                          // Last Updated
    ];
    
    // Update or insert the row
    if (existingRowIndex > CONFIG.TITLE_ROW) {
      // Update existing row
      sheet.getRange(existingRowIndex, 1, 1, dietaryData.length).setValues([dietaryData]);
      
      // Re-apply formatting
      formatDietaryRow(sheet, existingRowIndex);
    } else {
      // Find the correct position to insert the new row
      let insertRowIndex = CONFIG.DATA_START_ROW;
      if (sheet.getLastRow() >= CONFIG.DATA_START_ROW) {
        insertRowIndex = sheet.getLastRow() + 1;
        sheet.insertRowAfter(sheet.getLastRow());
      } else {
        // Make sure we have the right structure
        while (sheet.getLastRow() < CONFIG.TITLE_ROW) {
          sheet.insertRowAfter(sheet.getLastRow());
        }
        sheet.insertRowAfter(CONFIG.TITLE_ROW);
      }
      
      // Add data to the new row
      sheet.getRange(insertRowIndex, 1, 1, dietaryData.length).setValues([dietaryData]);
      
      // Apply formatting
      formatDietaryRow(sheet, insertRowIndex);
    }
    
    return true;
  } catch (error) {
    console.error("Error updating Dietary sheet:", error);
    // Continue execution even if there's an error
    return false;
  }
}

/**
 * Apply formatting to a row in the Dietary Information sheet
 * Modified to use plain formatting without backgrounds or special styling
 */
function formatDietaryRow(sheet, rowIndex) {
  try {
    // Format timestamp column
    sheet.getRange(rowIndex, 4).setNumberFormat("MMM dd, yyyy HH:mm:ss");
    
    // Make sure text is black for readability and clear any existing formatting
    const columnCount = sheet.getLastColumn();
    sheet.getRange(rowIndex, 1, 1, columnCount).setFontColor("#000000")
                                               .setBackground(null)
                                               .setFontWeight(null);
  } catch (error) {
    console.error("Error formatting Dietary row:", error);
  }
}

// ==========================================================================
// DATA REGENERATION FUNCTIONS
// ==========================================================================

/**
 * Regenerate all secondary sheets from the RSVP data
 */
function regenerateAllSheets() {
  try {
    regenerateGuestList();
    regenerateDietaryInformation();
    SpreadsheetApp.getUi().alert("All sheets have been regenerated successfully!");
    return true;
  } catch (error) {
    console.error("Error regenerating all sheets:", error);
    SpreadsheetApp.getUi().alert("Error: " + error.toString());
    return false;
  }
}

/**
 * Regenerate the Guest List sheet from the RSVP data
 */
function regenerateGuestList() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Get the RSVP sheet as the source of truth
    const rsvpSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (!rsvpSheet) {
      SpreadsheetApp.getUi().alert("RSVP sheet not found");
      return false;
    }
    
    // Get or create the Guest List sheet
    let guestListSheet = ss.getSheetByName(CONFIG.GUEST_LIST_SHEET_NAME);
    if (!guestListSheet) {
      guestListSheet = ss.insertSheet(CONFIG.GUEST_LIST_SHEET_NAME);
      formatGuestListSheet(guestListSheet);
    }
    
    // Clear existing content from data rows
    const lastRow = guestListSheet.getLastRow();
    if (lastRow >= CONFIG.DATA_START_ROW) {
      guestListSheet.getRange(CONFIG.DATA_START_ROW, 1, lastRow - CONFIG.DATA_START_ROW + 1, 
                             guestListSheet.getLastColumn()).clearContent();
    }
    
    // Ensure proper headers and structure
    formatGuestListSheet(guestListSheet);
    
    // Get all data from the RSVP sheet
    const rsvpData = rsvpSheet.getDataRange().getValues();
    
    // Find column indices
    const headers = rsvpData[0];
    const nameColIndex = headers.indexOf("Name");
    const emailColIndex = headers.indexOf("Email");
    const phoneColIndex = headers.indexOf("Phone");
    const attendingColIndex = headers.indexOf("Attending");
    const guestsColIndex = headers.indexOf("Number of Guests");
    const guestNamesColIndex = headers.indexOf("Guest Names");
    const statusColIndex = headers.indexOf("Status");
    
    if (nameColIndex === -1 || attendingColIndex === -1 || statusColIndex === -1) {
      SpreadsheetApp.getUi().alert("Required columns not found in RSVP sheet");
      return false;
    }
    
    // Filter and extract guest information
    const guestEntries = [];
    
    // Skip header rows
    for (let i = CONFIG.TITLE_ROW; i < rsvpData.length; i++) {
      const row = rsvpData[i];
      const name = row[nameColIndex];
      const email = row[emailColIndex];
      const phone = row[phoneColIndex];
      const attending = row[attendingColIndex];
      const guests = row[guestsColIndex] || 1;
      const guestNames = row[guestNamesColIndex] || "";
      const status = row[statusColIndex] || "Pending";
      
      // Include relevant entries
      if (name && (attending === "Y" || attending === "Maybe" || 
                   status === "Confirmed" || status === "Cancelled")) {
        guestEntries.push([
          name,           // Name
          email,          // Email
          phone,          // Phone
          attending,      // Attending
          guests,         // Number of Guests
          guestNames,     // Guest Names
          status,         // Status
          new Date()      // RSVP Date
        ]);
      }
    }
    
    // If we have entries, add them to the sheet
    if (guestEntries.length > 0) {
      // Insert data starting at the data start row
      guestListSheet.getRange(CONFIG.DATA_START_ROW, 1, guestEntries.length, 8).setValues(guestEntries);
      
      // Format all rows with plain formatting
      for (let i = 0; i < guestEntries.length; i++) {
        formatGuestListRow(guestListSheet, CONFIG.DATA_START_ROW + i);
      }
      
      // Add validation to the columns
      addGuestListSheetValidation(guestListSheet);
    }
    
    SpreadsheetApp.getUi().alert(`Guest List updated with ${guestEntries.length} entries`);
    return true;
  } catch (error) {
    console.error("Error regenerating guest list:", error);
    SpreadsheetApp.getUi().alert("Error: " + error.toString());
    return false;
  }
}

/**
 * Regenerate the Dietary Information sheet from the RSVP data
 */
function regenerateDietaryInformation() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Get the RSVP sheet as the source of truth
    const rsvpSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (!rsvpSheet) {
      SpreadsheetApp.getUi().alert("RSVP sheet not found");
      return false;
    }
    
    // Get or create the Dietary Information sheet
    let dietarySheet = ss.getSheetByName(CONFIG.DIETARY_SHEET_NAME);
    if (!dietarySheet) {
      dietarySheet = ss.insertSheet(CONFIG.DIETARY_SHEET_NAME);
      formatDietarySheet(dietarySheet);
    }
    
    // Clear existing content from data rows
    const lastRow = dietarySheet.getLastRow();
    if (lastRow >= CONFIG.DATA_START_ROW) {
      dietarySheet.getRange(CONFIG.DATA_START_ROW, 1, lastRow - CONFIG.DATA_START_ROW + 1, 
                           dietarySheet.getLastColumn()).clearContent();
    }
    
    // Ensure proper headers and structure
    formatDietarySheet(dietarySheet);
    
    // Get all data from the RSVP sheet
    const rsvpData = rsvpSheet.getDataRange().getValues();
    
    // Find column indices
    const headers = rsvpData[0];
    const nameColIndex = headers.indexOf("Name");
    const dietaryColIndex = headers.indexOf("Dietary Restrictions");
    const guestsColIndex = headers.indexOf("Number of Guests");
    const attendingColIndex = headers.indexOf("Attending");
    const statusColIndex = headers.indexOf("Status");
    
    if (nameColIndex === -1 || dietaryColIndex === -1) {
      SpreadsheetApp.getUi().alert("Required columns not found in RSVP sheet");
      return false;
    }
    
    // Filter and extract dietary information
    const dietaryEntries = [];
    
    // Skip header rows
    for (let i = CONFIG.TITLE_ROW; i < rsvpData.length; i++) {
      const row = rsvpData[i];
      const name = row[nameColIndex];
      const dietary = row[dietaryColIndex];
      const guests = row[guestsColIndex] || 1;
      const attending = row[attendingColIndex];
      const status = row[statusColIndex];
      
      // Only include valid entries with dietary information
      if (name && dietary && dietary.toString().trim() !== "") {
        if (attending === "Y" || attending === "Maybe" || status === "Confirmed") {
          dietaryEntries.push([
            name,           // Name
            dietary,        // Dietary Restrictions
            guests,         // Number of Guests
            new Date()      // Last Updated
          ]);
        }
      }
    }
    
    // If we have entries, add them to the sheet
    if (dietaryEntries.length > 0) {
      // Insert data starting at the data start row
      dietarySheet.getRange(CONFIG.DATA_START_ROW, 1, dietaryEntries.length, 4).setValues(dietaryEntries);
      
      // Format all rows with plain formatting
      for (let i = 0; i < dietaryEntries.length; i++) {
        formatDietaryRow(dietarySheet, CONFIG.DATA_START_ROW + i);
      }
    }
    
    SpreadsheetApp.getUi().alert(`Dietary Information updated with ${dietaryEntries.length} entries`);
    return true;
  } catch (error) {
    console.error("Error regenerating dietary information:", error);
    SpreadsheetApp.getUi().alert("Error: " + error.toString());
    return false;
  }
}

// ==========================================================================
// ADMIN FUNCTIONS
// ==========================================================================

/**
 * Refresh all sheets and status
 */
function refreshRsvpStatus() {
  try {
    ensureAllSheetsExist();
    regenerateAllSheets();
    addFormValidation();
    SpreadsheetApp.getUi().alert("RSVP status refreshed and all sheets regenerated!");
  } catch (error) {
    console.error("Error refreshing RSVP status:", error);
    SpreadsheetApp.getUi().alert("Error: " + error.toString());
  }
}

/**
 * Confirm selected RSVPs in the active sheet
 */
function confirmSelectedRsvps() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    
    // Check that we're on a relevant sheet
    const isRsvpSheet = sheet.getName() === CONFIG.SHEET_NAME;
    const isGuestListSheet = sheet.getName() === CONFIG.GUEST_LIST_SHEET_NAME;
    
    if (!isRsvpSheet && !isGuestListSheet) {
      SpreadsheetApp.getUi().alert(`Please select RSVPs on either the '${CONFIG.SHEET_NAME}' or '${CONFIG.GUEST_LIST_SHEET_NAME}' sheet.`);
      return;
    }
    
    // Get selected ranges
    const selectedRanges = sheet.getActiveRangeList().getRanges();
    if (!selectedRanges || selectedRanges.length === 0) {
      SpreadsheetApp.getUi().alert("Please select some rows to confirm.");
      return;
    }
    
    // Find important column indices
    const headers = sheet.getRange(CONFIG.HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
    const nameColIndex = headers.indexOf("Name") + 1;
    const attendingColIndex = headers.indexOf("Attending") + 1;
    const statusColIndex = headers.indexOf("Status") + 1;
    
    if (nameColIndex === 0 || attendingColIndex === 0 || statusColIndex === 0) {
      SpreadsheetApp.getUi().alert("Could not find required columns in the sheet.");
      return;
    }
    
    // Process each selected range
    let confirmedCount = 0;
    const confirmedNames = [];
    
    for (const range of selectedRanges) {
      // Skip header rows
      const startRow = Math.max(range.getRow(), CONFIG.DATA_START_ROW);
      
      // Skip if range is entirely in header rows
      if (startRow > range.getLastRow()) continue;
      
      // Calculate number of data rows in this range
      const firstRowOffset = (startRow > range.getRow()) ? startRow - range.getRow() : 0;
      const numRows = range.getNumRows() - firstRowOffset;
      
      if (numRows <= 0) continue;
      
      // Get data for the selected rows
      const nameData = sheet.getRange(startRow, nameColIndex, numRows, 1).getValues();
      const attendingData = sheet.getRange(startRow, attendingColIndex, numRows, 1).getValues();
      
      // Process each row
      for (let i = 0; i < numRows; i++) {
        const name = nameData[i][0];
        const attending = attendingData[i][0];
        
        // Skip empty names
        if (!name) continue;
        
        // Only confirm if attending is "Y" or "Maybe"
        if (attending === "Y" || attending === "Maybe") {
          // Update status in this sheet
          sheet.getRange(startRow + i, statusColIndex).setValue("Confirmed");
          
          // Only add background color if this is the RSVP sheet
          if (isRsvpSheet) {
            sheet.getRange(startRow + i, statusColIndex).setBackground(CONFIG.CONFIRMED_COLOR);
          }
          
          confirmedCount++;
          confirmedNames.push(name);
        }
      }
    }
    
    // If confirmed any entries, update the other sheet too
    if (confirmedNames.length > 0) {
      if (isRsvpSheet) {
        updateGuestListConfirmations(confirmedNames);
      } else {
        updateRsvpConfirmations(confirmedNames);
      }
    }
    
    SpreadsheetApp.getUi().alert(`Successfully confirmed ${confirmedCount} RSVPs.`);
  } catch (error) {
    console.error("Error confirming RSVPs:", error);
    SpreadsheetApp.getUi().alert("Error: " + error.toString());
  }
}

/**
 * Update confirmations in the Guest List based on names
 * Modified to not apply background colors
 */
function updateGuestListConfirmations(names) {
  try {
    if (!names || names.length === 0) return;
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const guestListSheet = ss.getSheetByName(CONFIG.GUEST_LIST_SHEET_NAME);
    
    if (!guestListSheet) return;
    
    // Find name and status column indices
    const headers = guestListSheet.getRange(CONFIG.HEADER_ROW, 1, 1, guestListSheet.getLastColumn()).getValues()[0];
    const nameColIndex = headers.indexOf("Name") + 1;
    const statusColIndex = headers.indexOf("Status") + 1;
    
    if (nameColIndex === 0 || statusColIndex === 0) return;
    
    // Get all data to find the rows to update
    const dataRange = guestListSheet.getDataRange();
    const data = dataRange.getValues();
    
    // Start from the data start row (skip headers)
    for (let i = CONFIG.TITLE_ROW; i < data.length; i++) {
      const rowName = data[i][nameColIndex - 1]; // -1 because arrays are 0-indexed
      
      if (names.includes(rowName)) {
        // Update this row - just change the status value, not the formatting
        guestListSheet.getRange(i + 1, statusColIndex).setValue("Confirmed");
      }
    }
  } catch (error) {
    console.error("Error updating Guest List confirmations:", error);
  }
}

/**
 * Update confirmations in the RSVP sheet based on names
 * Status column gets color for visibility
 */
function updateRsvpConfirmations(names) {
  try {
    if (!names || names.length === 0) return;
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const rsvpSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    
    if (!rsvpSheet) return;
    
    // Find name and status column indices
    const headers = rsvpSheet.getRange(CONFIG.HEADER_ROW, 1, 1, rsvpSheet.getLastColumn()).getValues()[0];
    const nameColIndex = headers.indexOf("Name") + 1;
    const statusColIndex = headers.indexOf("Status") + 1;
    
    if (nameColIndex === 0 || statusColIndex === 0) return;
    
    // Get all data to find the rows to update
    const dataRange = rsvpSheet.getDataRange();
    const data = dataRange.getValues();
    
    // Start from the data start row (skip headers)
    for (let i = CONFIG.TITLE_ROW; i < data.length; i++) {
      const rowName = data[i][nameColIndex - 1]; // -1 because arrays are 0-indexed
      
      if (names.includes(rowName)) {
        // Update this row
        rsvpSheet.getRange(i + 1, statusColIndex).setValue("Confirmed");
        // Only apply background to status column
        rsvpSheet.getRange(i + 1, statusColIndex).setBackground(CONFIG.CONFIRMED_COLOR);
      }
    }
  } catch (error) {
    console.error("Error updating RSVP confirmations:", error);
  }
}

/**
 * Confirm a specific RSVP by row
 */
function confirmRsvpByRow(rowIndex) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const rsvpSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    
    if (!rsvpSheet) {
      throw new Error("RSVP sheet not found");
    }
    
    // Find column indices
    const headers = rsvpSheet.getRange(CONFIG.HEADER_ROW, 1, 1, rsvpSheet.getLastColumn()).getValues()[0];
    const nameColIndex = headers.indexOf("Name") + 1;
    const statusColIndex = headers.indexOf("Status") + 1;
    
    if (nameColIndex === 0 || statusColIndex === 0) {
      throw new Error("Required columns not found");
    }
    
    // Get the name to update in Guest List too
    const name = rsvpSheet.getRange(rowIndex, nameColIndex).getValue();
    
    // Update status in RSVP sheet
    rsvpSheet.getRange(rowIndex, statusColIndex).setValue("Confirmed");
    rsvpSheet.getRange(rowIndex, statusColIndex).setBackground(CONFIG.CONFIRMED_COLOR);
    
    // Also update in Guest List
    updateGuestListConfirmations([name]);
    
    return true;
  } catch (error) {
    console.error("Error confirming RSVP by row:", error);
    throw error;
  }
}

// ==========================================================================
// HELPER FUNCTIONS
// ==========================================================================

/**
 * Find an existing entry by name in any sheet
 */
function findExistingEntry(sheet, name) {
  if (!sheet || !name) return 0;
  
  try {
    // Get all data from the sheet
    const data = sheet.getDataRange().getValues();
    
    // Determine the column index for name (usually the first or second column)
    let nameColIndex = 1; // Default to second column (index 1)
    
    if (data.length > 0 && data[0].includes("Name")) {
      nameColIndex = data[0].indexOf("Name");
    }
    
    // Skip header rows, start checking from the data start row
    for (let i = CONFIG.TITLE_ROW; i < data.length; i++) {
      if (data[i][nameColIndex] && 
          data[i][nameColIndex].toString().toLowerCase() === name.toLowerCase()) {
        return i + 1; // +1 because arrays are 0-indexed, but sheet rows are 1-indexed
      }
    }
    
    return 0; // Not found
  } catch (error) {
    console.error("Error finding existing entry:", error);
    return 0;
  }
}

/**
 * Find an existing entry in the Dietary Information sheet
 */
function findExistingDietaryEntry(sheet, name) {
  if (!sheet || !name) return 0;
  
  try {
    // Get all data from the sheet
    const data = sheet.getDataRange().getValues();
    
    // Determine the column index for name (usually the first column)
    let nameColIndex = 0; // Default to first column (index 0)
    
    // Skip header rows, start checking from the data start row
    for (let i = CONFIG.TITLE_ROW; i < data.length; i++) {
      if (data[i][nameColIndex] && 
          data[i][nameColIndex].toString().toLowerCase() === name.toLowerCase()) {
        return i + 1; // +1 because arrays are 0-indexed, but sheet rows are 1-indexed
      }
    }
    
    return 0; // Not found
  } catch (error) {
    console.error("Error finding existing dietary entry:", error);
    return 0;
  }
}

// ==========================================================================
// RSVP MANAGER FUNCTIONS
// ==========================================================================

/**
 * Show RSVP Manager sidebar panel
 */
function showRsvpManager() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; margin: 10px; }
      h2 { color: #4285F4; }
      .rsvp-item { 
        padding: 8px; 
        margin-bottom: 8px; 
        border-radius: 4px;
        border-left: 4px solid #4285F4;
      }
      .rsvp-confirmed { background-color: ${CONFIG.CONFIRMED_COLOR}; }
      .rsvp-pending { background-color: #ffffff; }
      .rsvp-cancelled { background-color: ${CONFIG.CANCELLED_COLOR}; }
      .rsvp-name { font-weight: bold; }
      .rsvp-status { color: #666; font-size: 0.8em; }
      .button-row { margin-top: 5px; }
      button {
        background-color: #4285F4;
        color: white;
        border: none;
        padding: 5px 10px;
        border-radius: 4px;
        cursor: pointer;
        font-size: 0.9em;
      }
      button:hover { background-color: #3367d6; }
      .refresh-btn {
        display: block;
        margin: 15px auto;
        padding: 8px 15px;
      }
      .stats {
        background-color: #f8f9fa;
        padding: 10px;
        border-radius: 4px;
        margin-bottom: 15px;
      }
      .status-filter {
        margin-bottom: 10px;
      }
    </style>
    
    <h2>RSVP Manager</h2>
    
    <div class="stats" id="rsvp-stats">Loading stats...</div>
    
    <div class="status-filter">
      <label>Filter: </label>
      <select id="status-filter" onchange="loadRsvps()">
        <option value="all">All RSVPs</option>
        <option value="pending" selected>Pending</option>
        <option value="confirmed">Confirmed</option>
        <option value="cancelled">Cancelled</option>
        <option value="maybe">Maybe</option>
      </select>
    </div>
    
    <div id="rsvp-list">Loading RSVPs...</div>
    <button class="refresh-btn" onclick="loadRsvps()">Refresh List</button>
    
    <script>
      // Load RSVPs on page load
      window.onload = function() {
        loadStats();
        loadRsvps();
      };
      
      // Load RSVP stats
      function loadStats() {
        google.script.run
          .withSuccessHandler(function(stats) {
            displayStats(stats);
          })
          .withFailureHandler(function(error) {
            document.getElementById('rsvp-stats').innerHTML = 
              "Error loading stats: " + error;
          })
          .getRsvpStats();
      }
      
      // Display the stats
      function displayStats(stats) {
        let html = "<strong>RSVP Summary:</strong><br>";
        html += "Total: " + stats.total + "<br>";
        html += "Confirmed: " + stats.confirmed + "<br>";
        html += "Pending: " + stats.pending + "<br>";
        html += "Maybe: " + stats.maybe + "<br>";
        html += "Cancelled: " + stats.cancelled + "<br>";
        html += "Total Guests: " + stats.totalGuests;
        
        document.getElementById('rsvp-stats').innerHTML = html;
      }
      
      // Load the RSVP list
      function loadRsvps() {
        document.getElementById('rsvp-list').innerHTML = "Loading...";
        
        const filter = document.getElementById('status-filter').value;
        
        google.script.run
          .withSuccessHandler(function(rsvps) {
            displayRsvps(rsvps);
          })
          .withFailureHandler(function(error) {
            document.getElementById('rsvp-list').innerHTML = 
              "Error loading RSVPs: " + error;
          })
          .getRsvpsByStatus(filter);
      }
      
      // Display the RSVPs
      function displayRsvps(rsvps) {
        const container = document.getElementById('rsvp-list');
        
        if (!rsvps || rsvps.length === 0) {
          container.innerHTML = "<p>No RSVPs found matching the filter.</p>";
          return;
        }
        
        let html = "";
        
        rsvps.forEach(function(rsvp) {
          let statusClass = "rsvp-pending";
          if (rsvp.status === "Confirmed") {
            statusClass = "rsvp-confirmed";
          } else if (rsvp.status === "Cancelled" || rsvp.status.includes("Cancelled")) {
            statusClass = "rsvp-cancelled";
          }
          
          html += '<div class="rsvp-item ' + statusClass + '">';
          html += '<div class="rsvp-name">' + rsvp.name + '</div>';
          html += '<div class="rsvp-status">' + 
                  'Status: ' + rsvp.status + ' | ' +
                  'Attending: ' + rsvp.attending + ' | ' +
                  'Guests: ' + rsvp.guests + 
                  '</div>';
          
          if (rsvp.status !== "Confirmed" && (rsvp.attending === "Y" || rsvp.attending === "Maybe")) {
            html += '<div class="button-row">';
            html += '<button onclick="confirmRsvp(' + rsvp.row + ')">Confirm RSVP</button>';
            html += '</div>';
          }
          
          html += '</div>';
        });
        
        container.innerHTML = html;
      }
      
      // Confirm an RSVP
      function confirmRsvp(rowIndex) {
        google.script.run
          .withSuccessHandler(function(result) {
            alert("RSVP confirmed successfully!");
            loadRsvps(); // Refresh the list
            loadStats(); // Refresh the stats
          })
          .withFailureHandler(function(error) {
            alert("Error confirming RSVP: " + error);
          })
          .confirmRsvpByRow(rowIndex);
      }
    </script>
  `)
  .setTitle('RSVP Manager')
  .setWidth(300);
  
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Get RSVP statistics for the manager panel
 */
function getRsvpStats() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    
    if (!sheet) {
      return { 
        total: 0, 
        confirmed: 0, 
        pending: 0, 
        maybe: 0, 
        cancelled: 0, 
        totalGuests: 0 
      };
    }
    
    // Find column indices
    const headers = sheet.getRange(CONFIG.HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
    const attendingColIndex = headers.indexOf("Attending");
    const statusColIndex = headers.indexOf("Status");
    const guestsColIndex = headers.indexOf("Number of Guests");
    
    if (attendingColIndex === -1 || statusColIndex === -1 || guestsColIndex === -1) {
      return { 
        total: 0, 
        confirmed: 0, 
        pending: 0, 
        maybe: 0, 
        cancelled: 0, 
        totalGuests: 0 
      };
    }
    
    // Get all data rows (skip header rows)
    const data = sheet.getDataRange().getValues();
    
    // Initialize counters
    let total = 0;
    let confirmed = 0;
    let pending = 0;
    let maybe = 0;
    let cancelled = 0;
    let totalGuests = 0;
    
    // Skip header rows
    for (let i = CONFIG.TITLE_ROW; i < data.length; i++) {
      const row = data[i];
      
      // Skip empty rows
      if (!row[headers.indexOf("Name")]) {
        continue;
      }
      
      total++;
      
      // Count by status
      if (row[statusColIndex] === "Confirmed") {
        confirmed++;
        // Add guests for confirmed RSVPs
        totalGuests += parseInt(row[guestsColIndex] || 1, 10);
      }
      else if (row[statusColIndex] === "Pending") {
        pending++;
      }
      else if (row[statusColIndex] === "Cancelled" || 
               (row[statusColIndex] && row[statusColIndex].includes("Cancelled"))) {
        cancelled++;
      }
      
      // Count maybe responses
      if (row[attendingColIndex] === "Maybe") {
        maybe++;
      }
    }
    
    return {
      total: total,
      confirmed: confirmed,
      pending: pending,
      maybe: maybe,
      cancelled: cancelled,
      totalGuests: totalGuests
    };
  } catch (error) {
    console.error("Error getting RSVP stats:", error);
    return { 
      error: error.toString(),
      total: 0, 
      confirmed: 0, 
      pending: 0, 
      maybe: 0, 
      cancelled: 0, 
      totalGuests: 0 
    };
  }
}

/**
 * Get RSVPs filtered by status
 */
function getRsvpsByStatus(statusFilter) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    
    if (!sheet) {
      return [];
    }
    
    // Find column indices
    const headers = sheet.getRange(CONFIG.HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
    const nameColIndex = headers.indexOf("Name");
    const attendingColIndex = headers.indexOf("Attending");
    const statusColIndex = headers.indexOf("Status");
    const guestsColIndex = headers.indexOf("Number of Guests");
    
    if (nameColIndex === -1 || attendingColIndex === -1 || statusColIndex === -1) {
      return [];
    }
    
    // Get all data rows (skip header rows)
    const data = sheet.getDataRange().getValues();
    
    // Filter RSVPs based on the status filter
    const rsvps = [];
    
    // Skip header rows
    for (let i = CONFIG.TITLE_ROW; i < data.length; i++) {
      const row = data[i];
      const name = row[nameColIndex];
      
      // Skip empty rows
      if (!name) {
        continue;
      }
      
      const attending = row[attendingColIndex];
      const status = row[statusColIndex];
      const guests = row[guestsColIndex] || 1;
      
      // Apply filter
      if (statusFilter === "all" ||
          (statusFilter === "pending" && status === "Pending") ||
          (statusFilter === "confirmed" && status === "Confirmed") ||
          (statusFilter === "cancelled" && (status === "Cancelled" || (status && status.includes("Cancelled")))) ||
          (statusFilter === "maybe" && attending === "Maybe")) {
        
        rsvps.push({
          name: name,
          attending: attending,
          status: status,
          guests: guests,
          row: i + 1 // +1 because arrays are 0-indexed but sheet rows are 1-indexed
        });
      }
    }
    
    return rsvps;
  } catch (error) {
    console.error("Error getting RSVPs by status:", error);
    return [];
  }
}

// ==========================================================================
// EVENT DETAILS FUNCTIONS
// ==========================================================================

/**
 * Get event details for the web app
 */
function getEventDetails() {
  return CONFIG.EVENT;
}

// ==========================================================================
// FRONTEND INTEGRATION FUNCTIONS
// ==========================================================================

/**
 * Get existing RSVPs for the frontend bubbles
 */
function getExistingRsvps() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    
    if (!sheet) {
      return [];
    }
    
    // Find column indices
    const headers = sheet.getRange(CONFIG.HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
    const nameColIndex = headers.indexOf("Name");
    const attendingColIndex = headers.indexOf("Attending");
    const statusColIndex = headers.indexOf("Status");
    
    if (nameColIndex === -1 || attendingColIndex === -1 || statusColIndex === -1) {
      return [];
    }
    
    // Get all data rows (skip header rows)
    const data = sheet.getDataRange().getValues();
    
    // Build the list of RSVPs
    const rsvps = [];
    
    // Skip header rows
    for (let i = CONFIG.TITLE_ROW; i < data.length; i++) {
      const row = data[i];
      const name = row[nameColIndex];
      
      // Skip empty rows
      if (!name) {
        continue;
      }
      
      rsvps.push({
        name: name,
        attending: row[attendingColIndex],
        status: row[statusColIndex]
      });
    }
    
    return rsvps;
  } catch (error) {
    console.error("Error getting existing RSVPs:", error);
    return [];
  }
}

/**
 * Get confirmed RSVPs for frontend bubble display
 */
function getConfirmedRsvps() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    
    if (!sheet) {
      return [];
    }
    
    // Find column indices
    const headers = sheet.getRange(CONFIG.HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
    const nameColIndex = headers.indexOf("Name");
    const statusColIndex = headers.indexOf("Status");
    
    if (nameColIndex === -1 || statusColIndex === -1) {
      return [];
    }
    
    // Get all data rows (skip header rows)
    const data = sheet.getDataRange().getValues();
    
    // Build the list of confirmed RSVPs
    const confirmedNames = [];
    
    // Skip header rows
    for (let i = CONFIG.TITLE_ROW; i < data.length; i++) {
      const row = data[i];
      const name = row[nameColIndex];
      const status = row[statusColIndex];
      
      // Skip empty rows
      if (!name) {
        continue;
      }
      
      // Only include confirmed RSVPs
      if (status === "Confirmed") {
        confirmedNames.push(name);
      }
    }
    
    console.log(`Found ${confirmedNames.length} confirmed RSVPs for bubble display`);
    return confirmedNames;
  } catch (error) {
    console.error("Error getting confirmed RSVPs:", error);
    return [];
  }
}

/**
 * Get the background image URL for the frontend
 */
function getBackgroundImageUrl() {
  return CONFIG.BACKGROUND_IMAGE;
}

/**
 * Handler for web application requests
 */
function doGet(e) {
  try {
    return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle(CONFIG.APP_NAME)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    console.error("Error serving app:", error);
    return HtmlService.createHtmlOutput(
      "<h1>Error</h1><p>There was a problem loading the application. Please try again later.</p>" +
      "<p>Error details: " + error.toString() + "</p>"
    );
  }
}

/**
 * Include HTML files for templating
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}