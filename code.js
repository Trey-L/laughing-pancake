/**
 * @OnlyCurrentDoc
 * Automates scheduling of Personal Voice slots from a Google Form into a
 * Google Sheet, sends emails, and handles rescheduling.
 * Uses a combined Time Slot column (HHMM-HHMM) and reads form data via e.namedValues.
 * Version: 2.1
 * Made by Trey Leong
 */

// --- Configuration ---
const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const SCHEDULE_SHEET_NAME = "Assembly Schedule"; // Name of the sheet with the schedule
const ADMIN_EMAIL = "INSERT YOUR OWN EMAIL HERE"; // <--- CHANGE THIS EMAIL

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
    <p>Thank you for volunteering!</p>
    <br>
    <p><em>(This is an automated message)</em></p>
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
  let reportingLocation = "the School Hall"; // Or be more specific if needed

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
    <p><em>(This is an automated message)</em></p>
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

// --- Daily Function: Triggered Daily ---

function dailyCheck() {
  Logger.log("Starting dailyCheck function.");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SCHEDULE_SHEET_NAME);
   if (!sheet) {
    Logger.log(`Error in dailyCheck: Sheet "${SCHEDULE_SHEET_NAME}" not found.`);
     MailApp.sendEmail(ADMIN_EMAIL, "PV Script Error", `Sheet "${SCHEDULE_SHEET_NAME}" not found during daily check.`);
    return;
  }
  const dataRange = sheet.getDataRange();
  // Get values once - use real values for dates/logic, display values for simple string checks like "Yes"
  const displayValues = dataRange.getDisplayValues();
  const realValues = dataRange.getValues();
  const numRows = dataRange.getNumRows();

  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  tomorrow.setHours(0, 0, 0, 0);
  // Use a reliable comparison format - ISO string parts or comparing date objects
  const tomorrowDateString = tomorrow.toDateString(); // e.g., "Tue Apr 16 2024"

  let processedSlotIdsForReminders = new Set(); // Avoid multiple reminders for multi-block slots
  let displacedSlots = new Map(); // Store info about slots that need rescheduling <assignedSlotId, studentData>

  // --- Pass 1: Send Reminders and Detect Displaced Slots ---
  Logger.log("Daily Check - Pass 1: Checking for reminders and displacements.");
  for (let i = 1; i < numRows; i++) { // Start row 2 (index 1)
    const assignedSlotId = realValues[i][COL_ASSIGNED_SLOT_ID - 1];
    const activityType = realValues[i][COL_ACTIVITY_TYPE - 1];
    const dateVal = realValues[i][COL_DATE - 1];
    const timeSlotString = realValues[i][COL_TIMESLOT - 1]; // Get the time slot string

    // Skip if essential data is missing for processing this row
    if (!(dateVal instanceof Date) || !timeSlotString || !activityType) continue;

    // Parse time for logic below - needed for both reminders and displacement original time
    const parsedTime = parseTimeSlot(timeSlotString, dateVal);
    if (!parsedTime.start) {
         Logger.log(`Skipping row ${i + 1} in daily check due to unparseable time slot: "${timeSlotString}"`);
         continue; // Cannot process without valid time
     }

    // A. Reminder Check
    if (activityType === TYPE_PV && assignedSlotId) { // Only check PV slots with an ID
      const reminderSent = displayValues[i][COL_REMINDER_SENT - 1]; // Use display value for "Yes"/"No" check
      const slotDate = new Date(dateVal);
      slotDate.setHours(0, 0, 0, 0);

      // Compare dates reliably
      if (slotDate.toDateString() === tomorrowDateString && reminderSent !== "Yes" && reminderSent !== "Error") {
          if (!processedSlotIdsForReminders.has(assignedSlotId)) {
            const studentEmail = realValues[i][COL_EMAIL - 1];
            const studentName = realValues[i][COL_STUDENT_NAME - 1];

             Logger.log(`Found slot for reminder: ID ${assignedSlotId}, Email: ${studentEmail}, Date: ${slotDate.toDateString()}`);

            if (studentEmail && studentName) {
              // Pass the parsed start time to the reminder function
              const success = sendReminderEmail(studentEmail, studentName, dateVal, parsedTime.start);
              updateStatus(assignedSlotId, COL_REMINDER_SENT, success ? "Yes" : "Error");
              processedSlotIdsForReminders.add(assignedSlotId); // Mark as processed
            } else {
                 Logger.log(`Skipping reminder for row ${i+1} (ID: ${assignedSlotId}) - missing email or name.`);
                 updateStatus(assignedSlotId, COL_REMINDER_SENT, "Error"); // Mark as error if data missing
                 processedSlotIdsForReminders.add(assignedSlotId);
            }
          } else {
              // Logger.log(`Reminder already processed for Slot ID ${assignedSlotId} (Row ${i+1})`);
          }
      }
    }

    // B. Displacement Check
    // If a row HAS an Assigned Slot ID but its Activity Type is NO LONGER "Personal Voice", it's been displaced.
    if (assignedSlotId && activityType !== TYPE_PV) {
       Logger.log(`Detected potential displacement: Row ${i + 1}, Slot ID ${assignedSlotId}, Activity Type is now "${activityType}"`);
      if (!displacedSlots.has(assignedSlotId)) {
         // Store the necessary data to reschedule this student later.
         // We only need one row's worth of data per displaced slot ID.
         const timeRequested = realValues[i][COL_TIME_REQUESTED - 1];
         displacedSlots.set(assignedSlotId, {
            email: realValues[i][COL_EMAIL - 1],
            name: realValues[i][COL_STUDENT_NAME - 1],
            class: realValues[i][COL_CLASS - 1],
            phone: realValues[i][COL_PHONE - 1],
            subject: realValues[i][COL_SUBJECT - 1],
            slides: realValues[i][COL_SLIDES - 1],
            // Ensure timeRequested is a valid number, default to SLOT_DURATION
            timeRequested: (typeof timeRequested === 'number' && timeRequested > 0) ? timeRequested : SLOT_DURATION,
            formTimestamp: realValues[i][COL_FORM_TIMESTAMP - 1],
            originalDate: dateVal, // Store the original date/time for the notification email
            originalStartTime: parsedTime.start // Store the original parsed start time
         });
          Logger.log(`Stored displacement data for Slot ID ${assignedSlotId}`);
      }
      // Important: Clear the PV-specific data and Assigned Slot ID from the manually changed row
      // to prevent re-detecting it tomorrow. Clears Activity Type through Form Timestamp.
      sheet.getRange(i + 1, COL_ACTIVITY_TYPE, 1, COL_FORM_TIMESTAMP - COL_ACTIVITY_TYPE + 1).clearContent();
      Logger.log(`Cleared PV data from manually changed row ${i + 1} (originally Slot ID ${assignedSlotId})`);
    }
  }
   SpreadsheetApp.flush(); // Make sure clearing is done before Pass 2

  // --- Pass 2: Process Rescheduling ---
   Logger.log(`Daily Check - Pass 2: Processing ${displacedSlots.size} displaced slots.`);
  if (displacedSlots.size > 0) {
    for (const [assignedSlotId, studentData] of displacedSlots.entries()) {
        Logger.log(`Attempting to reschedule Slot ID: ${assignedSlotId} for ${studentData.email}`);

        // Check if critical data was captured before attempting reschedule
        if(!studentData.email || !studentData.name){
            Logger.log(`Cannot reschedule Slot ID ${assignedSlotId} - missing critical data (email/name) captured from sheet. Notifying admin.`);
            const originalTimeStr = studentData.originalStartTime ? Utilities.formatDate(studentData.originalStartTime, SG_TIMEZONE, "HH:mm") : 'N/A';
            const originalDateStr = studentData.originalDate ? Utilities.formatDate(studentData.originalDate, SG_TIMEZONE, "yyyy-MM-dd") : 'N/A';
            MailApp.sendEmail(ADMIN_EMAIL, "PV Script Reschedule Failed", `Failed to reschedule Slot ID ${assignedSlotId}. Original Slot: ${originalDateStr} ${originalTimeStr}. Reason: Missing student name or email in the sheet row when displacement was detected.`);
            continue; // Skip to the next displaced slot
        }

        // Reuse the findAndAssignSlot logic
        const assignmentResult = findAndAssignSlot(studentData);

        if (assignmentResult.success) {
           Logger.log(`Successfully rescheduled Slot ID ${assignedSlotId} (New ID: ${assignmentResult.assignedSlotId}) for ${studentData.email}. Sending notification.`);
           const newSlotInfo = assignmentResult.slotInfo; // Contains parsed dates/times
           // Send reschedule notification using original and new slot info
          const notifySuccess = sendRescheduleEmail(
              studentData.email,
              studentData.name,
              studentData.originalDate,     // Original Date object
              studentData.originalStartTime, // Original parsed start time Date object
              newSlotInfo.date,              // New Date object
              newSlotInfo.startTime,         // New parsed start time Date object
              newSlotInfo.endTime,           // New parsed end time Date object
              studentData.subject
          );
           // Update confirmation status for the *new* assignment ID
           updateStatus(assignmentResult.assignedSlotId, COL_CONFIRMATION_SENT, notifySuccess ? "Yes" : "Error");
            Logger.log(`Reschedule email status updated for new Slot ID ${assignmentResult.assignedSlotId}: ${notifySuccess ? "Yes" : "Error"}`);

        } else {
            // Could not find a new slot automatically
            Logger.log(`Failed to find a new slot for displaced Slot ID ${assignedSlotId} (${studentData.email}). Notifying student and admin.`);
             // Format original date/time for the notification
             const formattedOldDate = Utilities.formatDate(studentData.originalDate, SG_TIMEZONE, "EEEE, MMMM dd, yyyy");
             const formattedOldStartTime = studentData.originalStartTime ? Utilities.formatDate(studentData.originalStartTime, SG_TIMEZONE, "hh:mm a") : "[Original Time Unknown]";
             // Notify student and admin about the cancellation
             MailApp.sendEmail(
                studentData.email + "," + ADMIN_EMAIL,
                "Important Update: Your Personal Voice Sharing Slot Cancelled",
                `Hi ${studentData.name || 'Student'},\n\nDue to changes in the Assembly programme, your Personal Voice sharing slot for "${studentData.subject || 'Your Topic'}" (originally scheduled for ${formattedOldDate} starting around ${formattedOldStartTime}) had to be removed.\n\nUnfortunately, the system could not automatically find a replacement slot at this time.\n\nPlease contact the teacher-in-charge to discuss manual rescheduling options if you still wish to present.\n\nWe apologize for the inconvenience.\n\nThank you.`
            );
        }
    }
  } else {
      Logger.log("Daily Check - Pass 2: No displaced slots detected today.");
  }

  Logger.log("dailyCheck function finished.");
}
