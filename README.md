# McRUN Attendance Code

## Project Description
The McRUN Attendance Code is a Google Apps Script-based solution designed to streamline attendance tracking for the McRUN club. It integrates with Google Sheets, Google Forms, and external tools to automate attendance submissions, formatting, and reporting. The system also includes features for email notifications, data validation, and integration with a Points Ledger for tracking participation.

## Key Features
- **Attendance Management**: Automates the processing of attendance submissions from Google Forms and the McRUN app.
- **Data Formatting**: Ensures uniform formatting of names, headruns, and other attendance details.
- **Email Notifications**: Sends reminders and attendance copies to headrunners and club executives.
- **Integration with Points Ledger**: Transfers attendance data to a Points Ledger for tracking participation.
- **Custom Menus**: Provides a user-friendly interface in Google Sheets for executing scripts.
- **Triggers and Scheduling**: Automates tasks like checking for missing attendance and updating calendar-based triggers.

## Folder Structure
```
appsscript.json
Attendance-Variables.gs
Copy-Email.html
Formatting.gs
HeadRun-Attendance.gs
HeadRun-Info.gs
Import.gs
LICENSE
Points-Ledger.gs
README.md
Reminder-Email.html
Triggers.gs
Unregistered.gs
User-Menu.gs
```

## Example Usage

### Automating Attendance Processing
```javascript
onFormSubmission(); // Automatically processes a new Google Form submission.
```

### Formatting Data
```javascript
cleanSheetData(); // Formats all rows in the attendance sheet.
```

### Sending Notifications
```javascript
sendEmailReminder_({
  emailsByLevel: { beginner: ["headrunner@mail.com"] },
  headrunTitle: "Monday 6pm"
});
```

### Managing Triggers
```javascript
updateWeeklyCalendarTriggers(); // Updates triggers for the upcoming week.
```

## License
This project is licensed under the MIT License. See the LICENSE file for details.