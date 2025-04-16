/**
 * @OnlyCurrentDoc
 * Automates scheduling of Personal Voice slots from a Google Form into a
 * Google Sheet, sends emails, and handles rescheduling.
 * Uses a combined Time Slot column (HHMM-HHMM) and reads form data via e.namedValues.
 * Version: 2.6 
 */

// --- Configuration ---
const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const SCHEDULE_SHEET_NAME = "Assembly Schedule"; // Name of the sheet with the schedule
const ADMIN_EMAIL = "INSERT EMAIL HERE"; // <--- CHANGE THIS EMAIL
// --- Configuration (Additions for Populate Schedule) ---
const NUMBER_OF_WEEKS_TO_ADD = 4; // How many weeks forward to populate each time you run it
const LIGHT_COLOR = '#ffffff'; // White
const DARK_COLOR = '#bfbdbd';  // Dark Grey 



// Column indices (A=1, B=2, etc.) - ADJUSTED FOR NEW LAYOUT!
const COL_DATE = 1;
const COL_TIMESLOT = 2; // New combined column
const COL_DURATION = 3; // Shifted
const COL_ACTIVITY_TYPE = 4; // Shifted
const COL_STUDENT_NAME = 5; // Shifted
const COL_CLASS = 6; // Shifted
const COL_PHONE = 7; // Shifted
const COL_SUBJECT = 8; // Shifted
const COL_SLIDES = 9; // Shifted
const COL_EMAIL = 10; // Shifted
const COL_TIME_REQUESTED = 11; // Shifted
const COL_CONFIRMATION_SENT = 12; // Shifted
const COL_REMINDER_SENT = 13; // Shifted
const COL_ASSIGNED_SLOT_ID = 14; // Shifted
const COL_FORM_TIMESTAMP = 15; // Shifted

// Activity Type Constants
const TYPE_EMPTY = "Empty";
const TYPE_PV = "Personal Voice";
const TYPE_ANNOUNCEMENT = "Announcement"; // Example other types
const TYPE_BLOCKED = "Blocked"; // Example other types

// Assembly Time Rules (Singapore Timezone assumed based on your location)
const SG_TIMEZONE = Session.getScriptTimeZone(); // Or specify "Asia/Singapore"
const MONDAY_START_HOUR = 9;
const MONDAY_START_MINUTE = 5;
const MONDAY_END_HOUR = 9;
const MONDAY_END_MINUTE = 40;
const TUE_FRI_START_HOUR = 8;
const TUE_FRI_START_MINUTE = 5;
const TUE_FRI_END_HOUR = 8;
const TUE_FRI_END_MINUTE = 20;
const SLOT_DURATION = 5; // Minutes per standard row/block

// --- Function to Populate Schedule Manually (with Formatting) ---
function populateSchedule() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SCHEDULE_SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert(`Error: Sheet "${SCHEDULE_SHEET_NAME}" not found.`);
    Logger.log(`Error in populateSchedule: Sheet "${SCHEDULE_SHEET_NAME}" not found.`);
    return;
  }

  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Populate & Format Schedule', // Updated title
    `This will add assembly slots for approximately ${NUMBER_OF_WEEKS_TO_ADD} future week(s) (starting after the last date found) AND apply alternating day background colors to the entire schedule.\n\nEnsure Column B ('Time Slot') is formatted as PLAIN TEXT before proceeding.\n\nContinue?`,
    ui.ButtonSet.YES_NO);

  if (response !== ui.Button.YES) {
    Logger.log("User cancelled schedule population and formatting.");
    return;
  }

  Logger.log(`Starting schedule population for ${NUMBER_OF_WEEKS_TO_ADD} week(s).`);

  // --- PART 1: Populate New Data ---
  let lastRowBeforeAdding = sheet.getLastRow();
  let valuesBeforeAdding = sheet.getDataRange().getValues(); // Get current values
  let lastDate = null;

  // Find the last date already in the sheet
  if (valuesBeforeAdding.length > 1) {
     for (let i = valuesBeforeAdding.length - 1; i >= 1; i--) {
         if (valuesBeforeAdding[i][COL_DATE - 1] instanceof Date) {
             lastDate = new Date(valuesBeforeAdding[i][COL_DATE - 1]);
             lastDate.setHours(0, 0, 0, 0);
             Logger.log(`Last existing date found: ${lastDate.toDateString()}`);
             break;
         }
     }
  }

  let startDate = new Date();
  startDate.setHours(0, 0, 0, 0);

  if (lastDate) {
     startDate = new Date(lastDate);
     startDate.setDate(startDate.getDate() + 1);
     Logger.log(`Starting population from: ${startDate.toDateString()}`);
  } else {
      Logger.log("No existing dates found or sheet empty. Starting population from today.");
  }

  const endDateLimit = new Date(startDate);
  endDateLimit.setDate(endDateLimit.getDate() + NUMBER_OF_WEEKS_TO_ADD * 7);

  let newData = []; // Array to hold new rows
  let currentDate = new Date(startDate);

  while (currentDate <= endDateLimit) {
    const dayOfWeek = currentDate.getDay();
    if (dayOfWeek === 0 || dayOfWeek === 6) {
      currentDate.setDate(currentDate.getDate() + 1);
      continue;
    }

    let startTime = new Date(currentDate);
    let endTimeLimit = new Date(currentDate);

    if (dayOfWeek === 1) { // Monday
      startTime.setHours(MONDAY_START_HOUR, MONDAY_START_MINUTE, 0, 0);
      endTimeLimit.setHours(MONDAY_END_HOUR, MONDAY_END_MINUTE, 0, 0);
    } else { // Tuesday - Friday
      startTime.setHours(TUE_FRI_START_HOUR, TUE_FRI_START_MINUTE, 0, 0);
      endTimeLimit.setHours(TUE_FRI_END_HOUR, TUE_FRI_END_MINUTE, 0, 0);
    }

    let currentTimeSlotStart = new Date(startTime);
    while (currentTimeSlotStart < endTimeLimit) {
      let currentTimeSlotEnd = new Date(currentTimeSlotStart);
      currentTimeSlotEnd.setMinutes(currentTimeSlotStart.getMinutes() + SLOT_DURATION);
      const timeSlotString = `${formatTime(currentTimeSlotStart)}-${formatTime(currentTimeSlotEnd)}`;
      newData.push([new Date(currentDate), timeSlotString, SLOT_DURATION, TYPE_EMPTY]);
      currentTimeSlotStart.setMinutes(currentTimeSlotStart.getMinutes() + SLOT_DURATION);
    }
    currentDate.setDate(currentDate.getDate() + 1);
  }

  // Add the new rows to the sheet if any were generated
  if (newData.length > 0) {
      const startRowForNewData = lastRowBeforeAdding + 1;
      const numColsToPopulate = 4; // Date, Time Slot, Duration, Activity Type
      sheet.getRange(startRowForNewData, 1, newData.length, numColsToPopulate).setValues(newData);
      Logger.log(`Successfully added ${newData.length} new schedule slots.`);
      SpreadsheetApp.flush(); // Ensure data is written before formatting
  } else {
      Logger.log("No new data generated.");
  }

  // --- PART 2: Apply Alternating Background Colors ---
  Logger.log("Applying alternating background colors to the schedule...");

  const dataRange = sheet.getDataRange(); // Get range *after* adding new data
  const values = dataRange.getValues();   // Get values *after* adding new data
  const numRows = dataRange.getNumRows();
  const numCols = sheet.getLastColumn(); // Format all columns up to the last one with any data

  let previousDateString = null;
  let useLightColor = true; // Start with light color

  // Loop through all data rows (skip header row 1, index 0)
  for (let i = 1; i < numRows; i++) {
    const currentRow = i + 1; // 1-based row index for getRange
    const dateValue = values[i][COL_DATE - 1]; // Get date from Column A

    if (dateValue instanceof Date) {
        const currentDateString = dateValue.toDateString(); // Compare based on date string "Tue Apr 16 2024"

        // Check if the day has changed compared to the previous row
        if (previousDateString !== null && currentDateString !== previousDateString) {
            useLightColor = !useLightColor; // Flip the color
        }

        // Determine the color for the current row
        const bgColor = useLightColor ? LIGHT_COLOR : DARK_COLOR;

        // Apply the background color to the entire row (up to the last column with data)
        try {
            sheet.getRange(currentRow, 1, 1, numCols).setBackground(bgColor);
        } catch (formatError) {
            // Log error but continue trying to format other rows
            Logger.log(`Error setting background for row ${currentRow}: ${formatError}`);
        }


        // Update the previous date string for the next iteration
        previousDateString = currentDateString;

    } else {
         // Optional: Handle rows where Column A is not a valid date (e.g., set a default color or skip)
         // sheet.getRange(currentRow, 1, 1, numCols).setBackground(null); // Clear background if date is invalid? Or just leave it.
         // Logger.log(`Skipping background format for row ${currentRow} as Column A is not a valid date.`);
    }
  }

   SpreadsheetApp.flush(); // Ensure formatting is applied

  // Final confirmation message
  if (newData.length > 0) {
      ui.alert(`Success! Added ${newData.length} new schedule slots and updated background colors.`);
      Logger.log(`Formatting complete.`);
  } else {
      ui.alert("No new slots were added, but existing row background colors have been updated.");
      Logger.log(`Formatting complete (no new rows added).`);
  }
} // End of populateSchedule function


// --- Helper Function: Format Time for HHMM ---
function formatTime(dateObj) {
    const hours = dateObj.getHours().toString().padStart(2, '0');
    const minutes = dateObj.getMinutes().toString().padStart(2, '0');
    return hours + minutes;
}

// --- Function to Add Custom Menu ---
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Assembly Scheduler')
      .addItem('Populate & Format Schedule', 'populateSchedule') // Updated menu item text
      .addToUi();
}




// --- Main Function: Triggered on Form Submission ---

function onFormSubmit(e) {
  Logger.log("onFormSubmit triggered (v2.1 - using namedValues).");
  try {
    // Log the event object structure for reference if needed in future debugging
    // Logger.log("Event object (e): " + JSON.stringify(e, null, 2));

    // --- 1. Check for necessary event data ---
    if (!e || !e.namedValues) {
      Logger.log("Stopping execution because 'e' or 'e.namedValues' is missing.");
      MailApp.sendEmail(ADMIN_EMAIL, "PV Script Error", "onFormSubmit trigger fired, but the event object (e) or e.namedValues was missing. Check trigger setup. See logs.");
      return; // Stop processing
    }

    Logger.log("e.namedValues found. Processing...");

    // --- 2. Get Form Data using e.namedValues ---
    const namedValues = e.namedValues;

    // **CRITICAL: Ensure these keys EXACTLY match your Google Form question titles**
    const formKeys = {
      email: 'Email Address', // Key used when "Collect email addresses" is enabled
      timestamp: 'Timestamp', // Automatically added by Sheets link
      name: 'Name',
      class: 'Class',
      phone: 'Phone Number',
      subject: 'Subject of Personal Voice Sharing',
      slides: 'Slides Link',
      timeRequired: 'Time Required (5min blocks)'
    };

    // Helper to safely get values from namedValues
    function getNamedValue(key) {
      return (namedValues[key] && namedValues[key][0]) ? namedValues[key][0] : null;
    }

    // Extract data
    const submitterEmail = getNamedValue(formKeys.email);
    const timestampString = getNamedValue(formKeys.timestamp);
    const formTimestamp = timestampString ? new Date(timestampString) : new Date(); // Use current time as fallback if timestamp missing

    const studentName = getNamedValue(formKeys.name) || "";
    const studentClass = getNamedValue(formKeys.class) || "";
    const studentPhone = getNamedValue(formKeys.phone) || "";
    const sharingSubject = getNamedValue(formKeys.subject) || "";
    const slidesLink = getNamedValue(formKeys.slides) || "";
    const timeReqString = getNamedValue(formKeys.timeRequired);

    // Parse time requested
    let timeRequested = SLOT_DURATION; // Default
    if (timeReqString) {
        const match = timeReqString.toString().match(/\d+/); // Find digits
        if (match) {
            timeRequested = parseInt(match[0], 10);
        }
    }

    // Populate the studentData object
    let studentData = {
      email: submitterEmail,
      timestamp: formTimestamp,
      name: studentName,
      class: studentClass,
      phone: studentPhone,
      subject: sharingSubject,
      slides: slidesLink,
      timeRequested: timeRequested,
    };

    // Validate mandatory data (e.g., email)
    if (!studentData.email) {
       Logger.log("Error: Submitter email not found in namedValues. Check Form's 'Collect email addresses' setting is enabled and the key '" + formKeys.email + "' is correct.");
       MailApp.sendEmail(ADMIN_EMAIL, "PV Script Error", "Form submitted but submitter email was missing in e.namedValues. Cannot proceed with scheduling for this submission. Check form settings and script's formKeys.");
       return; // Stop if email is mandatory
    }

    Logger.log(`Processing submission from: ${studentData.email}, Requesting: ${studentData.timeRequested} min`);

    // --- 3. Find and Assign Slot ---
    const assignmentResult = findAndAssignSlot(studentData);

    // --- 4. Send Confirmation Email ---
    if (assignmentResult.success) {
      Logger.log(`Slot assigned (New ID: ${assignmentResult.assignedSlotId}). Sending confirmation to ${studentData.email}`);
      const slotInfo = assignmentResult.slotInfo;
      const success = sendConfirmationEmail(
        studentData.email,
        studentData.name,
        slotInfo.date, // Date object
        slotInfo.startTime, // Date object (start of first block)
        slotInfo.endTime, // Date object (end of last block)
        studentData.subject
      );
      updateStatus(assignmentResult.assignedSlotId, COL_CONFIRMATION_SENT, success ? "Yes" : "Error");
      Logger.log(`Confirmation email status updated for ${assignmentResult.assignedSlotId}: ${success ? 'Yes' : 'Error'}`);

    } else {
      Logger.log(`No suitable slot found for submission from ${studentData.email}. Notifying student and admin.`);
      MailApp.sendEmail(
          studentData.email + "," + ADMIN_EMAIL,
          "Unable to Schedule Your Personal Voice Sharing",
          `Hi ${studentData.name || 'Student'},\n\nWe received your request to share on "${studentData.subject || 'Your Topic'}", but unfortunately, we couldn't automatically find a suitable ${studentData.timeRequested}-minute slot in the upcoming Assembly schedule.\n\nPlease check with the teacher-in-charge for manual scheduling options.\n\nThank you.`
      );
    }

  } catch (error) {
    Logger.log(`Error in onFormSubmit: ${error}\nStack: ${error.stack}`);
    // Try to log the event object again in case of error during processing
    try { Logger.log("Event object (e) at time of error: " + JSON.stringify(e, null, 2)); } catch (e) {}
    MailApp.sendEmail(ADMIN_EMAIL, "PV Script Error in onFormSubmit", `An error occurred: ${error}\n\nCheck script logs for details.\n\nStack Trace:\n${error.stack}`);
  }
} // End of onFormSubmit function


// --- Helper Function: Parse "HHMM-HHMM" Time Slot String ---
function parseTimeSlot(timeSlotString, baseDate) {
  if (!timeSlotString || typeof timeSlotString !== 'string' || !(baseDate instanceof Date)) {
    Logger.log(`Invalid input to parseTimeSlot: timeSlotString=${timeSlotString}, baseDate=${baseDate}`);
    return { start: null, end: null };
  }
  // Updated regex to be more robust
  const parts = timeSlotString.trim().match(/^(\d{2})(\d{2})-(\d{2})(\d{2})$/);
  if (!parts) {
    Logger.log(`Could not parse time slot string format: "${timeSlotString}"`);
    return { start: null, end: null };
  }
  try {
    const startHour = parseInt(parts[1], 10);
    const startMinute = parseInt(parts[2], 10);
    const endHour = parseInt(parts[3], 10);
    const endMinute = parseInt(parts[4], 10);

    // Basic sanity check for hour/minute values
    if (startHour > 23 || startMinute > 59 || endHour > 23 || endMinute > 59) {
         Logger.log(`Invalid hour/minute value in time slot: "${timeSlotString}"`);
         return { start: null, end: null };
    }

    const startDate = new Date(baseDate);
    startDate.setHours(startHour, startMinute, 0, 0);

    const endDate = new Date(baseDate);
    endDate.setHours(endHour, endMinute, 0, 0);

    // Handle overnight case (e.g. 2300-0100) - less likely for assembly but good practice
    if (endDate <= startDate && !(endDate.getHours() === 0 && endDate.getMinutes() === 0)) { // Allow midnight end
       // Check if it looks like an overnight slot
       if (endHour * 60 + endMinute < startHour * 60 + startMinute) {
          endDate.setDate(endDate.getDate() + 1); // Set end date to the next day
          Logger.log(`Detected potential overnight slot "${timeSlotString}", adjusted end date.`);
       } else {
           // End time is same or earlier on the same day - invalid
           Logger.log(`End time is not after start time in slot: "${timeSlotString}"`);
           return { start: null, end: null };
       }
    }

    return { start: startDate, end: endDate };
  } catch (e) {
      Logger.log(`Error parsing numbers in time slot string "${timeSlotString}": ${e}`);
      return { start: null, end: null };
  }
}

/**
 * Helper function to find an available slot specifically on a given date.
 * Very similar to findAndAssignSlot, but restricted to one day.
 *
 * @param {object} studentData The data object for the student needing a slot.
 * @param {Date} targetDate The specific date to search on.
 * @return {object} Result object: { success: boolean, assignedSlotId?: string, slotInfo?: object }
 */
function findSlotOnDate(studentData, targetDate) {
  if (!(targetDate instanceof Date)) {
    Logger.log(`findSlotOnDate called with invalid targetDate: ${targetDate}`);
    return { success: false };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SCHEDULE_SHEET_NAME);
  if (!sheet) return { success: false }; // Error logged elsewhere if sheet missing

  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  const numRows = dataRange.getNumRows();
  const blocksNeeded = Math.ceil(studentData.timeRequested / SLOT_DURATION);
  const targetDateString = targetDate.toDateString(); // For comparison

  Logger.log(`Searching specifically on date ${targetDateString} for ${blocksNeeded} blocks for ${studentData.email}`);

  const assignedSlotId = `PV_${studentData.email}_${new Date().getTime()}_RE`; // Add suffix for reschedule ID

  for (let i = 1; i < numRows; i++) { // Start row 2 (index 1)
    const rowDateVal = values[i][COL_DATE - 1];
    const timeSlotString = values[i][COL_TIMESLOT - 1];
    const activityType = values[i][COL_ACTIVITY_TYPE - 1];

    // Skip if not a date or doesn't match the target date
    if (!(rowDateVal instanceof Date) || rowDateVal.toDateString() !== targetDateString) {
      continue;
    }

    // --- Rest of the logic is almost identical to findAndAssignSlot's loop ---
    // --- (Validation, time parsing, assembly time check, consecutive empty check) ---

    if (!timeSlotString || activityType === undefined || activityType === null) continue;

    const parsedCurrentTime = parseTimeSlot(timeSlotString, rowDateVal);
    if (!parsedCurrentTime.start) continue;

    if (!isWithinAssemblyTime(rowDateVal, parsedCurrentTime.start)) continue;

    if (activityType === TYPE_EMPTY) {
      let consecutiveEmpty = 0;
      let potentialSlotRows = [];
      let firstSlotStartTime = null;
      let lastSlotEndTime = null;
      let previousSlotEndTime = null;

      for (let j = 0; j < blocksNeeded; j++) {
        const checkRowIndex = i + j;
        if (checkRowIndex >= numRows) break;

        const checkActivity = values[checkRowIndex][COL_ACTIVITY_TYPE - 1];
        const checkDateVal = values[checkRowIndex][COL_DATE - 1];
        const checkTimeSlotString = values[checkRowIndex][COL_TIMESLOT - 1];

        if (!(checkDateVal instanceof Date) || !checkTimeSlotString || checkActivity === undefined || checkActivity === null) break;

        // Ensure still on the target date
        if (checkDateVal.toDateString() !== targetDateString) break;

        const parsedCheckTime = parseTimeSlot(checkTimeSlotString, checkDateVal);
        if (!parsedCheckTime.start) break;

        if (checkActivity === TYPE_EMPTY && isWithinAssemblyTime(checkDateVal, parsedCheckTime.start)) {
           if (j > 0 && previousSlotEndTime && parsedCheckTime.start.getTime() !== previousSlotEndTime.getTime()) break; // Check consecutiveness

           consecutiveEmpty++;
           potentialSlotRows.push(checkRowIndex + 1);
           if (j === 0) firstSlotStartTime = parsedCheckTime.start;
           lastSlotEndTime = parsedCheckTime.end;
           previousSlotEndTime = parsedCheckTime.end;
        } else {
           break;
        }
      }

      // Found suitable block on this specific date!
      if (consecutiveEmpty >= blocksNeeded) {
        Logger.log(`Found suitable slot ON TARGET DATE ${targetDateString} starting at row ${i + 1}.`);

        // Update the sheet (same update logic as findAndAssignSlot)
        for (let k = 0; k < blocksNeeded; k++) {
            const sheetRowIndex = potentialSlotRows[k];
            let updateData = [];
            updateData[COL_ACTIVITY_TYPE - 1] = TYPE_PV;
            updateData[COL_STUDENT_NAME - 1] = studentData.name;
            updateData[COL_CLASS - 1] = studentData.class;
            updateData[COL_PHONE - 1] = studentData.phone;
            updateData[COL_SUBJECT - 1] = studentData.subject;
            updateData[COL_SLIDES - 1] = studentData.slides;
            updateData[COL_EMAIL - 1] = studentData.email;
            updateData[COL_TIME_REQUESTED - 1] = studentData.timeRequested;
            updateData[COL_CONFIRMATION_SENT - 1] = "No"; // Rescheduled, needs confirmation
            updateData[COL_REMINDER_SENT - 1] = "No";
            updateData[COL_ASSIGNED_SLOT_ID - 1] = assignedSlotId;
            updateData[COL_FORM_TIMESTAMP - 1] = studentData.formTimestamp; // Keep original timestamp

            let updateValuesRow = [];
            for(let colIdx = COL_ACTIVITY_TYPE; colIdx <= COL_FORM_TIMESTAMP; colIdx++) {
                updateValuesRow.push(updateData[colIdx - 1]);
            }
            const targetRange = sheet.getRange(sheetRowIndex, COL_ACTIVITY_TYPE, 1, COL_FORM_TIMESTAMP - COL_ACTIVITY_TYPE + 1);
            targetRange.setValues([updateValuesRow]);
        }
        SpreadsheetApp.flush();

        return {
          success: true,
          assignedSlotId: assignedSlotId,
          slotInfo: {
            date: rowDateVal, // The target date
            startTime: firstSlotStartTime,
            endTime: lastSlotEndTime,
            rows: potentialSlotRows
          }
        };
      }
      // Advance 'i' if consecutive blocks were checked
      i += (consecutiveEmpty > 0 ? consecutiveEmpty - 1 : 0);
    } // End if activityType is Empty
  } // End FOR loop through rows

  // No suitable slot found on this specific date
  Logger.log(`No suitable ${blocksNeeded}-block slot found specifically on date ${targetDateString}.`);
  return { success: false };
}

// --- Helper Function: Find and Assign Slot ---
function findAndAssignSlot(studentData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SCHEDULE_SHEET_NAME);
  if (!sheet) {
    Logger.log(`Error in findAndAssignSlot: Sheet "${SCHEDULE_SHEET_NAME}" not found.`);
    MailApp.sendEmail(ADMIN_EMAIL, "PV Script Error", `Sheet "${SCHEDULE_SHEET_NAME}" not found.`);
    return { success: false };
  }
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues(); // Read all data at once
  const numRows = dataRange.getNumRows();
  const blocksNeeded = Math.ceil(studentData.timeRequested / SLOT_DURATION);
  const today = new Date();
  today.setHours(0, 0, 0, 0); // Start checking from today

  const assignedSlotId = `PV_${studentData.email}_${new Date().getTime()}`; // Unique ID

  Logger.log(`Searching for ${blocksNeeded} consecutive blocks for ${studentData.email}`);

  for (let i = 1; i < numRows; i++) { // Start from row 2 (index 1) to skip header
    const rowDateVal = values[i][COL_DATE - 1];
    const timeSlotString = values[i][COL_TIMESLOT - 1];
    const activityType = values[i][COL_ACTIVITY_TYPE - 1];

    // Basic validation
    if (!(rowDateVal instanceof Date) || !timeSlotString || activityType === undefined || activityType === null) {
      continue; // Skip rows with missing date, timeslot, or activity type
    }

    const rowDate = new Date(rowDateVal); // Copy date to avoid modification issues
    rowDate.setHours(0, 0, 0, 0);

    if (rowDate < today) continue; // Skip past dates

    // Parse the time slot for the current row
    const parsedCurrentTime = parseTimeSlot(timeSlotString, rowDateVal);
    if (!parsedCurrentTime.start) {
        // Logger.log(`Skipping row ${i + 1} due to unparseable time slot: "${timeSlotString}"`);
        continue; // Skip if time slot is invalid
    }

    // Check if the slot start time is within allowed assembly times
    if (!isWithinAssemblyTime(rowDateVal, parsedCurrentTime.start)) {
      // Logger.log(`Row ${i+1} (${timeSlotString}) is outside assembly time.`);
      continue;
    }

    // Check if it's an empty slot and if there are enough consecutive slots
    if (activityType === TYPE_EMPTY) {
      let consecutiveEmpty = 0;
      let potentialSlotRows = []; // Store row indices (1-based)
      let firstSlotStartTime = null;
      let lastSlotEndTime = null;
      let previousSlotEndTime = null; // To check time consecutiveness

      for (let j = 0; j < blocksNeeded; j++) {
        const checkRowIndex = i + j;
        if (checkRowIndex >= numRows) {
             // Logger.log(`Consecutive check stopped at row ${checkRowIndex+1}: Reached end of sheet.`);
             break;
        }

        const checkActivity = values[checkRowIndex][COL_ACTIVITY_TYPE - 1];
        const checkDateVal = values[checkRowIndex][COL_DATE - 1];
        const checkTimeSlotString = values[checkRowIndex][COL_TIMESLOT - 1];

        // Ensure the potential next slot is valid
         if (!(checkDateVal instanceof Date) || !checkTimeSlotString || checkActivity === undefined || checkActivity === null) {
            // Logger.log(`Consecutive check stopped at row ${checkRowIndex+1}: Invalid data.`);
            break;
         }

         // Check date is the same as the starting slot's date
         if (new Date(checkDateVal).toDateString() !== new Date(rowDateVal).toDateString()) {
             // Logger.log(`Consecutive check stopped at row ${checkRowIndex+1}: Date mismatch.`);
             break; // Slot spans across midnight/different days, not allowed
         }

        // Parse the potential next slot's time
        const parsedCheckTime = parseTimeSlot(checkTimeSlotString, checkDateVal);
         if (!parsedCheckTime.start) {
             Logger.log(`Stopping consecutive check at row ${checkRowIndex + 1}, invalid time slot: "${checkTimeSlotString}"`);
             break; // Invalid time slot in sequence
         }

        // Ensure it's empty and within assembly time
        if (checkActivity === TYPE_EMPTY && isWithinAssemblyTime(checkDateVal, parsedCheckTime.start))
        {
           // Check if the start time of this slot matches the end time of the previous slot in the sequence
           if (j > 0 && previousSlotEndTime && parsedCheckTime.start.getTime() !== previousSlotEndTime.getTime()) {
              Logger.log(`Non-consecutive time gap detected between row ${checkRowIndex} and ${checkRowIndex + 1}. Expected start: ${Utilities.formatDate(previousSlotEndTime, SG_TIMEZONE, "HH:mm")}, Actual start: ${Utilities.formatDate(parsedCheckTime.start, SG_TIMEZONE, "HH:mm")}`);
              break; // Times are not contiguous
           }

           consecutiveEmpty++;
           potentialSlotRows.push(checkRowIndex + 1); // Store 1-based index
           if (j === 0) {
               firstSlotStartTime = parsedCheckTime.start; // Store start time of the first block
           }
           lastSlotEndTime = parsedCheckTime.end; // Store end time of the *current* block in sequence
           previousSlotEndTime = parsedCheckTime.end; // Update for next iteration's check

        } else {
           // Logger.log(`Consecutive check stopped at row ${checkRowIndex+1}: Not empty or outside assembly time.`);
           break; // Not consecutive, not empty, or outside allowed time
        }
      }

      // Found a suitable block!
      if (consecutiveEmpty >= blocksNeeded) {
        Logger.log(`Found suitable slot starting at row ${i + 1} for ${blocksNeeded} blocks.`);

        // Update the sheet for all allocated rows
        for (let k = 0; k < blocksNeeded; k++) {
          const sheetRowIndex = potentialSlotRows[k]; // Get the 1-based row index

          // Prepare only the columns we need to update
          let updateData = [];
          updateData[COL_ACTIVITY_TYPE - 1] = TYPE_PV; // Index adjusted for 0-based array
          updateData[COL_STUDENT_NAME - 1] = studentData.name;
          updateData[COL_CLASS - 1] = studentData.class;
          updateData[COL_PHONE - 1] = studentData.phone;
          updateData[COL_SUBJECT - 1] = studentData.subject;
          updateData[COL_SLIDES - 1] = studentData.slides;
          updateData[COL_EMAIL - 1] = studentData.email;
          updateData[COL_TIME_REQUESTED - 1] = studentData.timeRequested;
          updateData[COL_CONFIRMATION_SENT - 1] = "No";
          updateData[COL_REMINDER_SENT - 1] = "No";
          updateData[COL_ASSIGNED_SLOT_ID - 1] = assignedSlotId;
          updateData[COL_FORM_TIMESTAMP - 1] = studentData.timestamp;

           // Create a 2D array for setValues, matching the target range width
           // We need values for *all* columns in the target range
          let updateValuesRow = [];
          for(let colIdx = COL_ACTIVITY_TYPE; colIdx <= COL_FORM_TIMESTAMP; colIdx++) {
               updateValuesRow.push(updateData[colIdx - 1]); // Add prepared data
          }

          // Get range for the single row to update (only PV-specific columns)
          const targetRange = sheet.getRange(sheetRowIndex, COL_ACTIVITY_TYPE, 1, COL_FORM_TIMESTAMP - COL_ACTIVITY_TYPE + 1);
          targetRange.setValues([updateValuesRow]); // setValues expects a 2D array
          Logger.log(`Updated row ${sheetRowIndex} for AssignedSlotID: ${assignedSlotId}`);
        }

        SpreadsheetApp.flush(); // Ensure changes are written

        return {
          success: true,
          assignedSlotId: assignedSlotId,
          slotInfo: {
            date: rowDateVal, // The date of the first slot
            startTime: firstSlotStartTime, // Parsed start time Date object of the first slot
            endTime: lastSlotEndTime, // Parsed end time Date object of the *last* slot
            rows: potentialSlotRows // 1-based row indices
          }
        };
      }
       // If enough empty slots weren't found consecutively, the outer loop continues
       // Advance 'i' to skip rows already checked in the inner loop to avoid redundant checks
       i += (consecutiveEmpty > 0 ? consecutiveEmpty - 1 : 0);
    }
  }

  // No suitable slot found after checking all rows
  Logger.log(`Search complete. No suitable ${blocksNeeded}-block slot found for ${studentData.email}.`);
  return { success: false };
}


// --- Helper Function: Check if START time is within Assembly Time ---
// Accepts the base date and the *parsed start time* Date object
function isWithinAssemblyTime(dateVal, startTimeDateObj) {
  if (!(dateVal instanceof Date) || !(startTimeDateObj instanceof Date)) {
    return false;
  }

  const dayOfWeek = dateVal.getDay(); // 0=Sun, 1=Mon, ..., 6=Sat
  const hour = startTimeDateObj.getHours();
  const minute = startTimeDateObj.getMinutes();

  if (dayOfWeek === 1) { // Monday
    const startMinutes = MONDAY_START_HOUR * 60 + MONDAY_START_MINUTE;
    // Slot must START before assembly ENDS
    const endMinutesBoundary = MONDAY_END_HOUR * 60 + MONDAY_END_MINUTE;
    const currentMinutes = hour * 60 + minute;
    return currentMinutes >= startMinutes && currentMinutes < endMinutesBoundary;
  } else if (dayOfWeek >= 2 && dayOfWeek <= 5) { // Tuesday to Friday
    const startMinutes = TUE_FRI_START_HOUR * 60 + TUE_FRI_START_MINUTE;
    const endMinutesBoundary = TUE_FRI_END_HOUR * 60 + TUE_FRI_END_MINUTE;
    const currentMinutes = hour * 60 + minute;
    return currentMinutes >= startMinutes && currentMinutes < endMinutesBoundary;
  } else {
    return false; // Weekend
  }
}


// --- Helper Function: Send Confirmation Email ---
// Accepts parsed Date objects for startTime and endTime
function sendConfirmationEmail(email, name, date, startTime, endTime, subject) {
  if (!email || !(date instanceof Date) || !(startTime instanceof Date) || !(endTime instanceof Date)) {
      Logger.log(`Cannot send confirmation - missing data: email=${email}, date=${date}, startTime=${startTime}, endTime=${endTime}`);
      return false;
  }

  const formattedDate = Utilities.formatDate(date, SG_TIMEZONE, "EEEE, MMMM dd, yyyy");
  const formattedStartTime = Utilities.formatDate(startTime, SG_TIMEZONE, "hh:mm a");
  const formattedEndTime = Utilities.formatDate(endTime, SG_TIMEZONE, "hh:mm a"); // Use end time from last block

  const emailSubject = "Confirmation: Your Personal Voice Sharing Slot";
  const body = `
    <p>Hi ${name || 'Student'},</p>
    <p>Your Personal Voice sharing slot for "<strong>${subject || 'Your Topic'}</strong>" has been scheduled!</p>
    <p><strong>Date:</strong> ${formattedDate}</p>
    <p><strong>Time:</strong> ${formattedStartTime} - ${formattedEndTime}</p>
    <p>Please prepare your sharing and ensure any slides are ready and accessible via the link you provided.</p>
    <p>Further details regarding reporting time will be sent in a reminder email the day before your scheduled slot.</p>
    <p>Thank you!</p>
    <br>
    <p><em>(This is an automated message.)</em></p>
  `;

  try {
    MailApp.sendEmail({
      to: email,
      subject: emailSubject,
      htmlBody: body,
      // bcc: ADMIN_EMAIL // Optional: Blind copy admin for confirmations
      });
    Logger.log(`Confirmation email sent to ${email}`);
    return true;
  } catch (error) {
    Logger.log(`Error sending confirmation email to ${email}: ${error}`);
    return false;
  }
}

// --- Helper Function: Send Reminder Email ---
// Accepts parsed Date object for startTime
function sendReminderEmail(email, name, date, startTime) {
   if (!email || !(date instanceof Date) || !(startTime instanceof Date)) {
       Logger.log(`Cannot send reminder - missing data: email=${email}, date=${date}, startTime=${startTime}`);
       return false;
   }

  const formattedDate = Utilities.formatDate(date, SG_TIMEZONE, "EEEE, MMMM dd, yyyy");
  const formattedStartTime = Utilities.formatDate(startTime, SG_TIMEZONE, "hh:mm a");
  const dayOfWeek = date.getDay();

  let reportingTime = "";
  let reportingLocation = "Hall"; // Or be more specific if needed

  if (dayOfWeek === 1) { // Monday
    reportingTime = "8:45 AM"; // e.g., 20 mins before 9:05 AM
  } else if (dayOfWeek >= 2 && dayOfWeek <= 5) { // Tuesday to Friday
    reportingTime = "7:45 AM"; // e.g., 20 mins before 8:05 AM
  } else {
      Logger.log(`Attempted to send reminder for a weekend date (${formattedDate}) for ${email}. Skipping.`);
      return false; // Don't send for weekends
  }

  const emailSubject = `Reminder: Your Personal Voice Sharing Tomorrow (${formattedDate})`;
  const body = `
    <p>Hi ${name || 'Student'},</p>
    <p>This is a reminder about your Personal Voice sharing session tomorrow, <strong>${formattedDate}</strong>, starting at <strong>${formattedStartTime}</strong>.</p>
    <p><strong>Please report to ${reportingLocation} by ${reportingTime}</strong> to set up and prepare.</p>
    <p>Ensure your slides (if any) are accessible via the link you provided earlier.</p>
    <p>We look forward to your sharing!</p>
    <br>
    <p><em>(This is an automated message, please do not reply to this email.)</em></p>
  `;

  try {
    MailApp.sendEmail({
      to: email,
      subject: emailSubject,
      htmlBody: body,
       // bcc: ADMIN_EMAIL // Optional: Blind copy admin for reminders
      });
    Logger.log(`Reminder email sent to ${email} for slot on ${formattedDate}`);
     return true;
  } catch (error) {
    Logger.log(`Error sending reminder email to ${email}: ${error}`);
     return false;
  }
}

// --- Helper Function: Send Reschedule Notification Email ---
// Accepts Date objects for oldStartTime, newStartTime, newEndTime
function sendRescheduleEmail(email, name, oldDate, oldStartTime, newDate, newStartTime, newEndTime, subject) {
   if (!email || !(oldDate instanceof Date) || !(oldStartTime instanceof Date) || !(newDate instanceof Date) || !(newStartTime instanceof Date) || !(newEndTime instanceof Date) ) {
       Logger.log(`Cannot send reschedule email - missing or invalid date/time objects.`);
       return false;
   }

  const formattedOldDate = Utilities.formatDate(oldDate, SG_TIMEZONE, "EEEE, MMMM dd, yyyy");
  const formattedOldStartTime = Utilities.formatDate(oldStartTime, SG_TIMEZONE, "hh:mm a");
  const formattedNewDate = Utilities.formatDate(newDate, SG_TIMEZONE, "EEEE, MMMM dd, yyyy");
  const formattedNewStartTime = Utilities.formatDate(newStartTime, SG_TIMEZONE, "hh:mm a");
  const formattedNewEndTime = Utilities.formatDate(newEndTime, SG_TIMEZONE, "hh:mm a"); // Use end time of last block


  const emailSubject = "Important Update: Your Personal Voice Sharing Slot Rescheduled";
  const body = `
    <p>Hi ${name || 'Student'},</p>
    <p>Please note that your previously scheduled Personal Voice sharing slot for "<strong>${subject || 'Your Topic'}</strong>" on <strong>${formattedOldDate} at ${formattedOldStartTime}</strong> has been rescheduled due to changes in the assembly programme.</p>
    <p>Your new slot is:</p>
    <p><strong>Date:</strong> ${formattedNewDate}</p>
    <p><strong>Time:</strong> ${formattedNewStartTime} - ${formattedNewEndTime}</p>
    <p>We apologize for any inconvenience this may cause. Please prepare for the new date and time.</p>
    <p>A reminder email with reporting details will be sent the day before your new scheduled slot.</p>
    <p>Thank you for your understanding.</p>
    <br>
    <p><em>(This is an automated message)</em></p>
  `;

  try {
    MailApp.sendEmail({
      to: email,
      subject: emailSubject,
      htmlBody: body,
      bcc: ADMIN_EMAIL // Definitely notify admin of rescheduling
      });
    Logger.log(`Reschedule notification sent to ${email}. Old: ${formattedOldDate}, New: ${formattedNewDate}`);
     return true;
  } catch (error) {
    Logger.log(`Error sending reschedule email to ${email}: ${error}`);
     return false;
  }
}


// --- Helper Function: Update Status in Sheet ---
function updateStatus(assignedSlotId, columnToUpdate, status) {
  if (!assignedSlotId) {
      Logger.log("updateStatus called with null/empty assignedSlotId.");
      return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SCHEDULE_SHEET_NAME);
   if (!sheet) {
    Logger.log(`Error in updateStatus: Sheet "${SCHEDULE_SHEET_NAME}" not found.`);
    return;
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return; // No data rows to check

  // Get only the Assigned Slot ID column data
  const dataRange = sheet.getRange(2, COL_ASSIGNED_SLOT_ID, lastRow - 1, 1);
  const slotIds = dataRange.getValues();

  let updated = false;
  try {
      for (let i = 0; i < slotIds.length; i++) {
        if (slotIds[i][0] && slotIds[i][0].toString() === assignedSlotId.toString()) { // Check if ID matches
          const rowIndex = i + 2; // 1-based index for sheet rows (data starts row 2)
          sheet.getRange(rowIndex, columnToUpdate).setValue(status);
          updated = true;
          // Logger.log(`Updated row ${rowIndex}, column ${columnToUpdate} to "${status}" for Slot ID ${assignedSlotId}`);
          // Continue checking in case slot spanned multiple rows (though assigned ID should be same)
        }
      }
      if(updated) {
          SpreadsheetApp.flush(); // Ensure the change is saved immediately if any updates were made
      } else {
           Logger.log(`Did not find any rows with AssignedSlotID "${assignedSlotId}" to update column ${columnToUpdate}.`);
      }
  } catch (error) {
      Logger.log(`Error updating status for ${assignedSlotId}, column ${columnToUpdate} to ${status}: ${error}`)
      // Consider adding error status update? e.g., sheet.getRange(rowIndex, columnToUpdate).setValue("Update Error");
  }
}

// --- Daily Function: Triggered Daily (REVISED v6) ---

function dailyCheck() {
  Logger.log("Starting dailyCheck function (v6 - Corrected Clearing & Reschedule Debug).");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SCHEDULE_SHEET_NAME);
   if (!sheet) {
    Logger.log(`Error in dailyCheck: Sheet "${SCHEDULE_SHEET_NAME}" not found.`);
     MailApp.sendEmail(ADMIN_EMAIL, "PV Script Error", `Sheet "${SCHEDULE_SHEET_NAME}" not found during daily check.`);
    return;
  }
  const dataRange = sheet.getDataRange();
  const displayValues = dataRange.getDisplayValues();
  const realValues = dataRange.getValues();
  const numRows = dataRange.getNumRows();
  const lastCol = sheet.getLastColumn();

  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  tomorrow.setHours(0, 0, 0, 0);
  const tomorrowDateString = tomorrow.toDateString();

  let processedSlotIdsForReminders = new Set(); // Stores {id, status, data}
  let displacedSlotsData = new Map(); // <assignedSlotId, studentData> - Stores data ONCE
  let displacedIds = new Set();       // Stores IDs identified as displaced

  // --- Pass 1: Detect Reminders and Identify Displaced Booking IDs ---
  Logger.log("Daily Check - Pass 1: Identifying reminders and displaced booking IDs.");
  for (let i = 1; i < numRows; i++) { // Start row 2 (index 1)
    const assignedSlotId = realValues[i][COL_ASSIGNED_SLOT_ID - 1];
    if (!assignedSlotId) continue;

    const activityType = realValues[i][COL_ACTIVITY_TYPE - 1];
    const dateVal = realValues[i][COL_DATE - 1];
    const timeSlotString = realValues[i][COL_TIMESLOT - 1];

    if (!(dateVal instanceof Date) || !timeSlotString || !activityType) {
        Logger.log(`Row ${i + 1} has AssignedSlotID "${assignedSlotId}" but missing critical data (Date/Time/Activity). Skipping processing for this row.`);
        continue;
    }

    const parsedTime = parseTimeSlot(timeSlotString, dateVal);
    if (!parsedTime.start) {
        Logger.log(`Row ${i + 1} has AssignedSlotID "${assignedSlotId}" but unparseable time slot: "${timeSlotString}". Cannot process for reminder/displacement.`);
        continue;
    }

    // A. Reminder Check
    if (activityType === TYPE_PV) {
      const reminderSent = displayValues[i][COL_REMINDER_SENT - 1];
      const slotDate = new Date(dateVal);
      slotDate.setHours(0, 0, 0, 0);
      if (slotDate.toDateString() === tomorrowDateString && reminderSent !== "Yes" && reminderSent !== "Error") {
          let found = false;
          processedSlotIdsForReminders.forEach(item => { if (item.id === assignedSlotId) found = true; });
          if (!found) {
             const studentEmail = realValues[i][COL_EMAIL - 1];
             const studentName = realValues[i][COL_STUDENT_NAME - 1];
             Logger.log(`Queueing reminder for: ID ${assignedSlotId}, Row ${i+1}, Email: ${studentEmail}`);
             if (studentEmail && studentName) {
                 processedSlotIdsForReminders.add({id: assignedSlotId, status: null, data: {email: studentEmail, name: studentName, date: dateVal, startTime: parsedTime.start}});
             } else {
                 Logger.log(`Cannot queue reminder for ID ${assignedSlotId} - missing email or name.`);
                 processedSlotIdsForReminders.add({id: assignedSlotId, status: "Error", data: null});
             }
          }
      }
    }

    // B. Displacement Identification: If ID exists and Activity is NOT PV
    if (activityType !== TYPE_PV) {
       if (!displacedIds.has(assignedSlotId)) {
           Logger.log(`Detected displacement trigger: Row ${i + 1}, Slot ID ${assignedSlotId}, Activity Type is now "${activityType}"`);
           displacedIds.add(assignedSlotId);
       }
       if (!displacedSlotsData.has(assignedSlotId)) {
         const timeRequested = realValues[i][COL_TIME_REQUESTED - 1];
         displacedSlotsData.set(assignedSlotId, {
            email: realValues[i][COL_EMAIL - 1],
            name: realValues[i][COL_STUDENT_NAME - 1],
            class: realValues[i][COL_CLASS - 1],
            phone: realValues[i][COL_PHONE - 1],
            subject: realValues[i][COL_SUBJECT - 1],
            slides: realValues[i][COL_SLIDES - 1],
            timeRequested: (typeof timeRequested === 'number' && timeRequested > 0) ? timeRequested : SLOT_DURATION,
            formTimestamp: realValues[i][COL_FORM_TIMESTAMP - 1],
            originalDate: dateVal,
            originalStartTime: parsedTime.start
         });
          Logger.log(`Stored displacement data for Slot ID ${assignedSlotId} from row ${i+1}`);
       }
    }
  } // End Pass 1 FOR loop

  // --- Send Reminders (after iterating) ---
  if (processedSlotIdsForReminders.size > 0) {
      Logger.log(`Sending reminders for ${processedSlotIdsForReminders.size} unique slot IDs.`);
      processedSlotIdsForReminders.forEach(item => {
          if (item.data && !item.status) {
             const success = sendReminderEmail(item.data.email, item.data.name, item.data.date, item.data.startTime);
             item.status = success ? "Yes" : "Error";
          }
          if (item.status) {
             updateStatus(item.id, COL_REMINDER_SENT, item.status);
          }
      });
      SpreadsheetApp.flush();
  }


  // --- Clear Displaced Rows (Conditional Revert) ---
  let rowsToClearMap = new Map(); // <assignedSlotId, array of row indices>
  if (displacedIds.size > 0) {
      Logger.log(`Finding all rows associated with ${displacedIds.size} displaced booking ID(s).`);
      const allValues = sheet.getDataRange().getValues(); // Get current data again
      const idColumnValues = sheet.getRange(2, COL_ASSIGNED_SLOT_ID, sheet.getLastRow() > 1 ? sheet.getLastRow() - 1 : 1, 1).getValues();

      displacedIds.forEach(idToClear => {
          let indicesForId = [];
          for (let k = 0; k < idColumnValues.length; k++) {
              if (idColumnValues[k][0] && idColumnValues[k][0].toString() === idToClear.toString()) {
                  indicesForId.push(k + 2); // Store 1-based row index
              }
          }
          if (indicesForId.length > 0) {
               rowsToClearMap.set(idToClear, indicesForId);
               Logger.log(`ID ${idToClear} found in rows: ${indicesForId.join(', ')}`);
          } else {
              Logger.log(`Warning: Displaced ID ${idToClear} detected but no rows found with this ID during clearing phase.`);
          }
      });

      // Now perform the clearing - conditionally reverting Activity Type
      if (rowsToClearMap.size > 0) {
          Logger.log(`Clearing data for displaced bookings (preserving non-PV activities).`);
          rowsToClearMap.forEach((rowIndices, slotId) => {
              Logger.log(`Processing rows for Slot ID ${slotId}: ${rowIndices.join(', ')}`);
              rowIndices.forEach(rowIndex => {
                  try {
                      // 1. Clear PV-specific columns (Name to Timestamp)
                      const startClearCol = COL_STUDENT_NAME;
                      const endClearCol = COL_FORM_TIMESTAMP;
                      const numColsToClear = endClearCol - startClearCol + 1;

                      if(startClearCol <= endClearCol) {
                          sheet.getRange(rowIndex, startClearCol, 1, numColsToClear).clearContent();
                      }

                      // 2. Clear the Assigned Slot ID
                      sheet.getRange(rowIndex, COL_ASSIGNED_SLOT_ID).clearContent();

                      // 3. **Conditionally set Activity Type back to Empty**
                      const activityCell = sheet.getRange(rowIndex, COL_ACTIVITY_TYPE);
                      const currentActivityType = activityCell.getValue();
                      if (currentActivityType === TYPE_PV) {
                          activityCell.setValue(TYPE_EMPTY);
                          Logger.log(`Row ${rowIndex}: Set Activity Type from PV to Empty.`);
                      } else {
                          // Leave the non-PV activity type (e.g., "Announcement") as is
                          Logger.log(`Row ${rowIndex}: Activity Type is "${currentActivityType}", leaving unchanged.`);
                      }

                  } catch (clearError) {
                      Logger.log(`Error clearing/resetting row ${rowIndex} for ID ${slotId}: ${clearError}`);
                  }
              });
          });
          SpreadsheetApp.flush(); // IMPORTANT: Ensure sheet updates before finding new slots
          Logger.log("Finished clearing displaced data (preserving non-PV activity types).");
      }
  } // End Clearing section


  // --- Pass 2: Process Rescheduling Attempts ---
   Logger.log(`Daily Check - Pass 2: Processing ${displacedSlotsData.size} unique displaced bookings.`);
  if (displacedSlotsData.size > 0) {

    for (const [assignedSlotId, studentData] of displacedSlotsData.entries()) {
        Logger.log(`Attempting to reschedule original Slot ID: ${assignedSlotId} for ${studentData.email}`);

        if(!studentData.email || !studentData.name || !(studentData.originalDate instanceof Date) || !(studentData.originalStartTime instanceof Date)){
            Logger.log(`Cannot reschedule Slot ID ${assignedSlotId} - missing critical data captured. Notifying admin.`);
            const originalTimeStr = studentData.originalStartTime ? Utilities.formatDate(studentData.originalStartTime, SG_TIMEZONE, "HH:mm") : 'N/A';
            const originalDateStr = studentData.originalDate ? Utilities.formatDate(studentData.originalDate, SG_TIMEZONE, "yyyy-MM-dd") : 'N/A';
            MailApp.sendEmail(ADMIN_EMAIL, "PV Script Reschedule Failed", `Failed to reschedule Slot ID ${assignedSlotId}. Original Slot: ${originalDateStr} ${originalTimeStr}. Reason: Missing critical student or original slot data captured during displacement.`);
            continue;
        }

        // ** TWO-PASS SEARCH **
        let assignmentResult = null;

        // Pass 2a: Try finding a slot on the ORIGINAL date first
        Logger.log(`Reschedule Pass 2a: Searching for ${studentData.timeRequested}min slot on original date (${studentData.originalDate.toDateString()}) for ${studentData.email}`);
        assignmentResult = findSlotOnDate(studentData, studentData.originalDate);

        // Pass 2b: If no slot found on original date, search entire schedule
        if (!assignmentResult.success) {
            Logger.log(`Reschedule Pass 2b: No slot found on original date. Performing full schedule search for ${studentData.email}`);
            assignmentResult = findAndAssignSlot(studentData);
        }

        // Process the result
        if (assignmentResult.success) {
           Logger.log(`Reschedule SUCCESS: Found new slot for original ID ${assignedSlotId}. New Assigned ID: ${assignmentResult.assignedSlotId}. Location: ${assignmentResult.slotInfo.date.toDateString()}. Sending notification.`);
           const newSlotInfo = assignmentResult.slotInfo;
           const notifySuccess = sendRescheduleEmail(
              studentData.email,
              studentData.name,
              studentData.originalDate,
              studentData.originalStartTime,
              newSlotInfo.date,
              newSlotInfo.startTime,
              newSlotInfo.endTime,
              studentData.subject
          );
           updateStatus(assignmentResult.assignedSlotId, COL_CONFIRMATION_SENT, notifySuccess ? "Yes" : "Error");
            Logger.log(`Reschedule email status updated for new Slot ID ${assignmentResult.assignedSlotId}: ${notifySuccess ? "Yes" : "Error"}`);
        } else {
            Logger.log(`Reschedule FAILED: Could not find ANY new slot for original ID ${assignedSlotId} (${studentData.email}). Notifying student and admin.`);
            const formattedOldDate = Utilities.formatDate(studentData.originalDate, SG_TIMEZONE, "EEEE, MMMM dd, yyyy");
            const formattedOldStartTime = Utilities.formatDate(studentData.originalStartTime, SG_TIMEZONE, "hh:mm a");
            MailApp.sendEmail(
                studentData.email + "," + ADMIN_EMAIL,
                "Important Update: Your Personal Voice Sharing Slot [Cancelled]",
                `Hi ${studentData.name || 'Student'},\n\nDue to changes in the Assembly programme, your Personal Voice sharing slot for "${studentData.subject || 'Your Topic'}" (originally scheduled for ${formattedOldDate} starting around ${formattedOldStartTime}) had to be removed.\n\nUnfortunately, the system could not automatically find a replacement slot at this time.\n\nPlease contact the teacher-in-charge to discuss manual rescheduling options if you still wish to present.\n\nWe apologize for the inconvenience.\n\nThank you.`
            );
        }
    } // End FOR loop for rescheduling
  } else {
      Logger.log("Daily Check - Pass 2: No displaced bookings needed rescheduling today.");
  }

  Logger.log("dailyCheck function finished.");
} // End dailyCheck function

