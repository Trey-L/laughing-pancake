# laughing-pancake
> A Google Apps Script for automating the process of allocating students available time slots during morning assembly for sharings.
> 
1. Students enter their details into a Google Form
2. The Script automatically triggers onFormSubmission to allocate the next available time slot in a Google Sheet to the students, also sending a Confirmation Email.

Ensure to do the following:

1. **Configure `ADMIN_EMAIL`:** Change the placeholder email in the Configuration section to a real address.
2. **Verify `formKeys`:** Double-check that the keys listed in the `formKeys` object inside `onFormSubmit` **exactly** match your Google Form question titles (case-sensitive).
3. **Triggers:** Ensure you have two triggers set up:
    - `onFormSubmit` running "From spreadsheet" on "On form submit".
    - `dailyCheck` running "Time-driven" on a "Daily timer" (e.g., early morning).
