# Google Apps Script Lab

This repository contains a collection of Google Apps Scripts (GAS) that automate various processes in Google Workspace, such as Google Calendar automation, delegation, Google Sheets synchronization, PTO tracking, and account reactivation management.

## Table of Contents
- [Google Apps Script Lab](#google-apps-script-lab)
  - [Table of Contents](#table-of-contents)
  - [Scripts Overview](#scripts-overview)
  - [Requirements](#requirements)
  - [Installation](#installation)
  - [Usage](#usage)
  - [Contributing](#contributing)
  - [License](#license)

## Scripts Overview

### 1. [gas_cab_form_calendar_automation.js](gas_cab_form_calendar_automation.js)
**Description**: This script automates the creation of Google Calendar events based on form submissions. It integrates with Google Forms and Google Calendar to schedule events based on the data submitted in the form. It is designed for use with the CAB (Change Advisory Board) to streamline event creation and scheduling.

### 2. [gas_gcal_delegation.js](gas_cab_form_calendar_automation.js)
**Description**: This script handles Google Calendar delegation. It automatically assigns delegation access to specified users. It can be used to manage calendar permissions at scale, especially useful for handling user transitions in organizations.

### 3. [gas_gcal_gsheet_sync.js](gas_gcal_gsheet_sync.js)
**Description**: This script syncs data between Google Calendar and Google Sheets. It is used to pull calendar events and update a Google Sheet with the event details, enabling reporting or further processing based on calendar activity.

### 4. [gas_gcal_pto_sync.js](gas_gcal_gsheet_sync.js)
**Description**: This script automates PTO (Paid Time Off) tracking by syncing employee PTO events from a shared Google Calendar to a Google Sheet. It can generate reports on PTO usage and ensure that records are kept up to date across multiple systems.

### 5. [gas_gw_account_reactivation_manager.js](gas_gw_account_reactivation_manager.js)
**Description**: This script manages Google Workspace account reactivations. It automates the process of monitoring suspended accounts and reactivating them based on certain criteria, such as form submissions or administrative actions.

### 6. [gas_lh_intake.js](gas_lh_intake.js)
**Description**: This script processes emails with specific criteria and extracts attachments to be stored in Google Drive. It's designed for processing legal hold reports and saving them automatically into a designated Google Drive folder.

## Requirements
- **Google Apps Script API Access**: Ensure that your Google Apps Script project has the necessary API access, including permissions for Google Calendar, Google Sheets, Gmail, and Google Drive, depending on the script's purpose.
- **Google Workspace Admin Permissions**: Some scripts may require admin-level access to manage user accounts and calendar permissions.

## Installation
1. Open the Google Apps Script editor within your Google Workspace.
2. Copy the content of the desired script into a new Google Apps Script project.
3. Set up the required triggers and permissions for the script, as outlined in each script's comments.
4. Save and run the script.

## Usage
1. **Calendar Automation**:
   - Automate calendar event creation via Google Forms submissions.
   - Run `gas_cab_form_calendar_automation.js` to set up events.

2. **Calendar Delegation**:
   - Run `gas_gcal_delegation.js` to assign or remove calendar delegation rights.

3. **Google Calendar to Sheets Sync**:
   - Use `gas_gcal_gsheet_sync.js` to pull event data from a Google Calendar into a Google Sheet.

4. **PTO Sync**:
   - Sync PTO events from a shared calendar using `gas_gcal_pto_sync.js`.

5. **Account Reactivation**:
   - Automate Google Workspace account reactivations using `gas_gw_account_reactivation_manager.js`.

6. **Legal Hold Intake**:
   - Process and save email attachments automatically with `gas_lh_intake.js`.

## Contributing
Feel free to submit issues or pull requests if you would like to contribute to the scripts or improve automation workflows.

## License
This project is licensed under the MIT License.
