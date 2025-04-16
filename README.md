# Google Apps Script: School Assembly Slot Scheduler

This Google Apps Script automates the process of scheduling student "Personal Voice" sharing slots during morning assembly based on submissions from a Google Form. It updates a central Google Sheet schedule, sends confirmation and reminder emails, and handles rescheduling when slots are manually overwritten, prioritizing same-day rescheduling.

## Overview

This script addresses the manual work involved in:
1.  Receiving student volunteer sign-ups via Google Forms.
2.  Finding available 5-minute time slots in a pre-defined assembly schedule (stored in a Google Sheet).
3.  Updating the Google Sheet with the student's details for their assigned slot.
4.  Notifying the student via email about their confirmed slot.
5.  Sending a reminder email the day before their scheduled slot.
6.  Detecting when a scheduled slot is manually overridden (e.g., for a priority announcement).
7.  Attempting to automatically reschedule the displaced student, prioritizing finding a new slot on the *same day* before searching other dates.
8.  Preserving manual overrides (like "Announcement") while clearing other data from displaced slots before rescheduling.

## Features

*   **Form Processing:** Automatically triggered when a linked Google Form is submitted, using `e.namedValues` for robust data retrieval.
*   **Slot Allocation:** Finds the next available sequence of 5-minute "Empty" slots matching the student's requested time.
*   **Sheet Updates:** Populates the designated rows in the Google Sheet schedule with student details and marks the slot as "Personal Voice".
*   **Confirmation Emails:** Sends an immediate confirmation email to the student with their assigned date and time range.
*   **Reminder Emails:** Sends a reminder email the day before the scheduled slot, including reporting time and location.
*   **Intelligent Rescheduling:**
    *   Detects manual overrides of scheduled PV slots.
    *   Clears original student data but *preserves* the manual override activity type (e.g., "Announcement").
    *   Attempts to find a new slot first *on the original date*.
    *   If no same-day slot is found, searches the entire schedule.
    *   Notifies the student of the reschedule or cancellation.
*   **Status Tracking:** Updates columns in the sheet to indicate whether confirmation and reminder emails have been sent ("Yes", "No", "Error").
*   **Automatic Schedule Population:** Includes a function (`populateSchedule`) accessible via a custom menu to automatically add future dates and time slots to the schedule sheet, applying alternating row colors for readability.
*   **Error Notifications:** Sends basic error notifications to a designated admin email address.

## Prerequisites

*   A Google Account (e.g., your school's G Suite account).
*   A Google Form for student sign-ups.
*   A Google Sheet to manage the assembly schedule.

## Setup Instructions

Follow these steps carefully to implement the script:

### 1. Google Sheet Setup

1.  Create a new Google Sheet or use an existing one. Name it descriptively (e.g., "Assembly Programme").
2.  Rename the first sheet tab to **`Assembly Schedule`**. The script specifically looks for this name (case-sensitive).
3.  Set up the following columns in the **first row (Row 1)**. The **exact order and names** are important for the script.

    | Column | Header                  | Format         | Notes                                           |
    | :----- | :---------------------- | :------------- | :---------------------------------------------- |
    | A      | `Date`                  | Date           | Date of the assembly slot                       |
    | B      | `Time Slot`             | **Plain Text** | Format: `HHMM-HHMM` (e.g., `0905-0910`)        |
    | C      | `Duration (min)`        | Number         | Typically `5` for standard slots                |
    | D      | `Activity Type`         | Plain Text     | e.g., `Empty`, `Personal Voice`, `Announcement` |
    | E      | `Student Name`          | Plain Text     | Populated by script                             |
    | F      | `Class`                 | Plain Text     | Populated by script                             |
    | G      | `Phone Number`          | Plain Text     | Populated by script                             |
    | H      | `Subject`               | Plain Text     | Populated by script                             |
    | I      | `Slides Link`           | Plain Text     | Populated by script                             |
    | J      | `Email Address`         | Plain Text     | Populated by script                             |
    | K      | `Time Requested (min)`  | Number         | Populated by script                             |
    | L      | `Confirmation Sent`     | Plain Text     | Updated by script (`Yes`/`No`/`Error`)         |
    | M      | `Reminder Sent`         | Plain Text     | Updated by script (`Yes`/`No`/`Error`)         |
    | N      | `Assigned Slot ID`      | Plain Text     | Unique ID generated by script                   |
    | O      | `Form Timestamp`        | Date Time      | Populated by script                             |

4.  **Format Column B:** Ensure the `Time Slot` column (B) is formatted as **Plain Text** *before* running the schedule population function.
5.  **Populate Initial Schedule (Optional but Recommended):** You can manually fill in the first few days/weeks or use the script's population function (see Step 6). Ensure `Activity Type` is `Empty` for available slots.

### 2. Google Form Setup

1.  Create your Google Form with questions corresponding to the data needed (Name, Class, Phone, Subject, Slides Link, Time Required).
    *   For **"Time Required"**, use a format the script can understand (e.g., Dropdown or Short Answer returning "5 min", "10", etc.). The script extracts the number.
2.  **Crucially:** In the Form editor, go to **Settings** > **Responses**. Enable **"Collect email addresses"** (set to "Verified" recommended). The script relies on the field title `Email Address` being present in the response data.
3.  Link the Form to your Google Sheet:
    *   Go to the **Responses** tab in the Form editor.
    *   Click the Google Sheets icon ("Create spreadsheet").
    *   Select **"Select existing spreadsheet"** and choose the Google Sheet you set up in Step 1.
    *   This creates a *new tab* (e.g., "Form Responses 1"). **Leave this tab as is.** The script interacts primarily with the `Assembly Schedule` tab.

### 3. Google Apps Script Setup

1.  Open your Google Sheet (`Assembly Programme`).
2.  Go to **Extensions** > **Apps Script**. This opens the script editor.
3.  Delete any placeholder code.
4.  Copy the **entire** code from the `Code.gs` (or equivalent) file provided (ensure it includes `onOpen`, `populateSchedule`, `onFormSubmit`, `findAndAssignSlot`, `findSlotOnDate`, `dailyCheck`, email functions, helpers, and all constants).
5.  Paste the code into the script editor.
6.  **Configure the Script:**
    *   Find the `// --- Configuration ---` section near the top.
    *   Change `const ADMIN_EMAIL = "your_admin_email@yourschool.edu.sg";` to a real email address for error notifications.
    *   Verify/adjust constants like `NUMBER_OF_WEEKS_TO_ADD`, assembly start/end times (`MONDAY_...`, `TUE_FRI_...`), and colors (`LIGHT_COLOR`, `DARK_COLOR`) if needed.
    *   **Verify `formKeys`:** Locate the `formKeys` object within the `onFormSubmit` function. Ensure the values (e.g., `'Name'`, `'Class'`, `'Time Required (5min blocks)'`) **exactly match** the question titles in *your* Google Form (case-sensitive). The keys `Email Address` and `Timestamp` correspond to standard fields when collecting emails and linking to Sheets.
7.  **Save the script:** Click the floppy disk icon (Save project). Give it a name (e.g., "Assembly Scheduler").

### 4. Authorization

The script needs your permission to access your Sheet, send emails, and run automatically.

1.  You will likely be prompted when saving, running a function manually for the first time, or setting up triggers.
2.  A dialog "Authorization required" will appear. Click **Review permissions**.
3.  Choose your Google account.
4.  You may see a "Google hasn't verified this app" screen. Click **Advanced**, then click **"Go to [Your Script Name] (unsafe)"**. *(This is standard for custom scripts)*.
5.  Review the permissions the script needs (manage spreadsheets, send email as you, run when you're not present, display UI elements). Click **Allow**.

### 5. Trigger Setup

Triggers tell the script when to run automatically.

1.  In the script editor, click the **Triggers** icon (looks like a clock) on the left sidebar.
2.  Click the **+ Add Trigger** button (bottom right).
3.  **Trigger 1: Form Submission**
    *   Choose which function to run: `onFormSubmit`
    *   Choose which deployment should run: `Head`
    *   Select event source: `From spreadsheet`
    *   Select event type: `On form submit`
    *   Failure notification settings: `Notify me immediately` (Recommended)
    *   Click **Save**.
4.  Click **+ Add Trigger** again.
5.  **Trigger 2: Daily Check**
    *   Choose which function to run: `dailyCheck`
    *   Choose which deployment should run: `Head`
    *   Select event source: `Time-driven`
    *   Select type of time based trigger: `Daily timer`
    *   Select time of day: Choose a time unlikely to conflict with manual edits (e.g., `1am - 2am` or `4am - 5am`). This must run *before* the assembly day to send reminders.
    *   Failure notification settings: `Notify me immediately` (Recommended)
    *   Click **Save**.

### 6. Populate Initial Schedule (Using Script)

1.  **Reload** your Google Sheet after saving the script.
2.  A new menu item **"Assembly Scheduler"** should appear.
3.  Ensure Column B (`Time Slot`) is formatted as **Plain Text**.
4.  Click **Assembly Scheduler** > **Populate & Format Schedule**.
5.  Confirm in the dialog box.
6.  The script will add future slots (defined by `NUMBER_OF_WEEKS_TO_ADD`) and apply alternating background colors to the entire schedule. Run this whenever you need to extend the schedule further.

## How it Works

1.  **Form Submission (`onFormSubmit`):** Triggered on form submission, extracts data via `e.namedValues`.
2.  **Slot Finding (`findAndAssignSlot`):** Searches the `Assembly Schedule` sheet for consecutive `Empty` slots matching the required duration, starting from the current date.
3.  **Sheet Update:** If a slot is found, updates relevant rows with student data, sets `Activity Type` to `Personal Voice`, adds a unique `Assigned Slot ID`.
4.  **Confirmation Email:** Sends confirmation to the student.
5.  **Daily Check (`dailyCheck` - Time Trigger):**
    *   **Reminders:** Sends reminder emails for PV slots scheduled the next day.
    *   **Displacement Detection:** Identifies bookings where at least one associated row is no longer `Personal Voice`.
    *   **Data Capture:** Stores essential data for the displaced student.
    *   **Clearing:** Finds *all* rows belonging to the displaced booking. Clears student data and the ID from all these rows. Sets `Activity Type` back to `Empty` *only* if it was originally `Personal Voice` (preserves manual overrides like "Announcement").
    *   **Two-Pass Rescheduling:**
        *   Calls `findSlotOnDate` to search for a new slot *only on the original date*.
        *   If unsuccessful, calls `findAndAssignSlot` to search the entire schedule.
    *   **Reschedule/Cancellation Email:** Notifies the student of the new slot or cancellation.
6.  **Schedule Population (`populateSchedule` - Manual Menu):** Adds future date/time slots and applies alternating row formatting when run by the user.

## Important Notes

*   **Timezone:** Uses the Google Sheet's timezone setting. Ensure it's correct for your location.
*   **Sheet/Column Names:** The script strictly depends on the sheet name `Assembly Schedule` and the exact column headers/order defined in Step 1.
*   **Form Question Titles:** The script relies on exact matches between your form questions and the keys defined in the `formKeys` object within `onFormSubmit`.
*   **Manual Overrides:** Manually changing an `Activity Type` for a scheduled PV slot *will* trigger the rescheduling process on the next daily run. The manual activity type (e.g., "Announcement") will be preserved on that specific row.
*   **Testing:** **Thoroughly test** all functions (form submission, rescheduling, daily checks, schedule population) before relying on the script. Use the script editor's Logs (`View` > `Logs` or "Executions" tab) to debug issues.
*   **Permissions:** The script requires permissions to manage your Sheet, send emails as you, and run automatically.

## Customization

*   **Admin Email:** Change `ADMIN_EMAIL` constant.
*   **Form Questions:** Modify the `formKeys` object in `onFormSubmit`.
*   **Assembly Times/Weeks to Add:** Adjust constants like `MONDAY_START_HOUR`, `TUE_FRI_END_MINUTE`, `NUMBER_OF_WEEKS_TO_ADD`.
*   **Colors:** Modify `LIGHT_COLOR`, `DARK_COLOR` constants.
*   **Email Content:** Edit the `send...Email` functions.
