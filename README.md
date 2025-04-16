# Google Apps Script: Automated Student Sharing (Personal Voice Sharing) Scheduler

This Google Apps Script automates the process of scheduling student sharing slots during morning assembly based on submissions from a Google Form. It updates a central Google Sheet schedule, sends confirmation and reminder emails, and handles basic rescheduling if slots are manually overwritten.

## Overview

This script addresses the manual work involved in:
1.  Receiving student volunteer sign-ups via Google Forms.
2.  Finding available 5-minute time slots in a pre-defined assembly schedule (stored in a Google Sheet).
3.  Updating the Google Sheet with the student's details for their assigned slot.
4.  Notifying the student via email about their confirmed slot.
5.  Sending a reminder email the day before their scheduled slot.
6.  Attempting to automatically reschedule a student if their slot is manually overridden in the Google Sheet for priority events.

## Features

*   **Form Processing:** Automatically triggered when a linked Google Form is submitted.
*   **Slot Allocation:** Finds the next available sequence of 5-minute "Empty" slots matching the student's requested time.
*   **Sheet Updates:** Populates the designated rows in the Google Sheet schedule with student details (Name, Class, Subject, Links, Email, etc.) and marks the slot as "Personal Voice".
*   **Confirmation Emails:** Sends an immediate confirmation email to the student with their assigned date and time.
*   **Reminder Emails:** Sends a reminder email the day before the scheduled slot, including reporting time and location.
*   **Rescheduling:** Detects if a scheduled "Personal Voice" slot is manually changed to something else in the Sheet. It attempts to find a new slot for the displaced student and notifies them of the change (or cancellation if no new slot is found).
*   **Status Tracking:** Updates columns in the sheet to indicate whether confirmation and reminder emails have been sent ("Yes", "No", "Error").
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

4.  **Populate Initial Schedule:** Manually fill in rows for upcoming assembly days and their 5-minute slots. Enter the `Date`, `Time Slot` (as text `HHMM-HHMM`), and `Duration (min)`. Set `Activity Type` to `Empty` for slots available for Personal Voice. Mark any known announcements or blocked times accordingly.

### 2. Google Form Setup

1.  Create your Google Form with questions corresponding to the data needed (Name, Class, Phone, Subject, Slides Link, Time Required).
    *   For **"Time Required"**, use a format the script can understand (e.g., Dropdown or Short Answer returning "5 min", "10 min", etc.). The script extracts the number.
2.  **Crucially:** In the Form editor, go to **Settings** > **Responses**. Enable **"Collect email addresses"** (set to "Verified" recommended). This adds the `Email Address` field automatically.
3.  Link the Form to your Google Sheet:
    *   Go to the **Responses** tab in the Form editor.
    *   Click the Google Sheets icon ("Create spreadsheet").
    *   Select **"Select existing spreadsheet"** and choose the Google Sheet you set up in Step 1.
    *   This will create a *new tab* in your Sheet (e.g., "Form Responses 1"). **Leave this tab as is.** The script works primarily with the `Assembly Schedule` tab you prepared.

### 3. Google Apps Script Setup

1.  Open your Google Sheet (`Assembly Programme`).
2.  Go to **Extensions** > **Apps Script**. This opens the script editor.
3.  Delete any placeholder code (like `function myFunction() {}`).
4.  Copy the **entire** code from the `Code.gs` (or equivalent) file in this repository.
5.  Paste the code into the script editor.
6.  **Configure the Script:**
    *   Find the `// --- Configuration ---` section near the top.
    *   Change `const ADMIN_EMAIL = "your_admin_email@yourschool.edu.sg";` to a real email address that should receive error notifications.
    *   **Verify `formKeys`:** Locate the `formKeys` object within the `onFormSubmit` function. Ensure the values (e.g., `'Name'`, `'Class'`, `'Time Required (5min blocks)'`) **exactly match** the question titles in *your* Google Form (case-sensitive). The keys `Email Address` and `Timestamp` should generally not be changed if using standard Form/Sheet linking.
7.  **Save the script:** Click the floppy disk icon (Save project). Give it a name (e.g., "Assembly Scheduler"). The file will likely be named `Code.gs` by default.

### 4. Authorization

The script needs your permission to access your Sheet, send emails, and run automatically.

1.  You might be prompted when saving, or when setting up triggers (next step). If not, select any function (like `dailyCheck`) from the dropdown menu next to the Debug (bug) icon and click **Run**.
2.  A dialog "Authorization required" will appear. Click **Review permissions**.
3.  Choose your Google account.
4.  You might see a "Google hasn't verified this app" screen. Click **Advanced**, then click **"Go to [Your Script Name] (unsafe)"**.
5.  Review the permissions the script needs (manage spreadsheets, send email as you, run when you're not present, etc.). Click **Allow**.

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
    *   Click **Save**. You might need to authorize again.
4.  Click **+ Add Trigger** again.
5.  **Trigger 2: Daily Check**
    *   Choose which function to run: `dailyCheck`
    *   Choose which deployment should run: `Head`
    *   Select event source: `Time-driven`
    *   Select type of time based trigger: `Daily timer`
    *   Select time of day: Choose a time unlikely to conflict with manual edits (e.g., `1am - 2am` or `4am - 5am`). This needs to run *before* the assembly day to send reminders.
    *   Failure notification settings: `Notify me immediately` (Recommended)
    *   Click **Save**.

## How it Works

1.  **Form Submission:** When a student submits the linked Google Form, the `onFormSubmit` trigger fires the corresponding function.
2.  **Data Extraction:** The script reads the submitted data (including the automatically collected email) using `e.namedValues`.
3.  **Slot Finding:** `findAndAssignSlot` searches the `Assembly Schedule` sheet for the required number of consecutive 5-minute rows marked as `Empty`, starting from the current date and within valid assembly times.
4.  **Sheet Update:** If suitable slots are found, the script updates those rows with the student's details, changes `Activity Type` to `Personal Voice`, sets email statuses to `No`, and adds a unique `Assigned Slot ID`.
5.  **Confirmation Email:** `sendConfirmationEmail` sends an email to the student with their assigned date and time range. The `Confirmation Sent` status is updated.
6.  **Daily Check:** The `dailyCheck` function runs once per day via its time trigger.
    *   It scans the schedule for `Personal Voice` slots scheduled for *tomorrow*. If found and a reminder hasn't been sent, `sendReminderEmail` sends the reminder and updates the `Reminder Sent` status.
    *   It also scans for rows that *have* an `Assigned Slot ID` but are *no longer* marked as `Personal Voice`. This indicates a manual override. The script stores the student's data, clears the PV info from the overridden row, and attempts to find a *new* slot using `findAndAssignSlot`. If successful, `sendRescheduleEmail` notifies the student; otherwise, a cancellation email is sent.

## Important Notes

*   **Timezone:** The script uses the Google Sheet's timezone setting (`Session.getScriptTimeZone()`) for date/time comparisons and email formatting. Ensure this is set correctly for your location (Singapore).
*   **Sheet Structure:** The script is highly dependent on the exact sheet name (`Assembly Schedule`) and the column order/headers defined above. Changes will break the script unless the `COL_` constants and code logic are updated accordingly.
*   **Form Questions:** The script relies on the exact text of your form questions matching the keys defined in the `formKeys` object within `onFormSubmit`.
*   **Initial Data:** Ensure your `Assembly Schedule` sheet has rows representing future available "Empty" slots for the script to find.
*   **Manual Overrides:** If you manually edit a scheduled student slot (e.g., change `Activity Type` to `Announcement`), the script will detect this on its next daily run and attempt to reschedule the student. The original row's PV data will be cleared.
*   **Testing:** Test thoroughly with dummy form submissions and by manually changing schedule entries *before* relying on the script for live scheduling. Check the script logs (`View` > `Logs` or "Executions" tab) for errors.
*   **Permissions:** The script requires significant permissions. Understand what you are allowing when authorizing.

## Customization

*   **Admin Email:** Change `ADMIN_EMAIL` for error notifications.
*   **Form Questions:** Modify the `formKeys` object in `onFormSubmit` if your form questions differ.
*   **Assembly Times:** Adjust the `MONDAY_START/END_HOUR/MINUTE` and `TUE_FRI_START/END_HOUR/MINUTE` constants if your assembly schedule changes.
*   **Email Content:** Modify the `sendConfirmationEmail`, `sendReminderEmail`, and `sendRescheduleEmail` functions to change the wording or formatting of the emails.
