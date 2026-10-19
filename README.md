# IT Support Request Tracker & Analytics Dashboard

A Google Sheets-based IT support request management solution designed to organize incoming requests, automate follow-ups, and support operational reporting.

## Project Overview

This project combines Google Forms, Google Sheets, and Google Apps Script to streamline the IT support request workflow, from submission and department assignment to status tracking and reporting.

The goal is to reduce repetitive administrative tasks, improve request visibility, and make IT support operations easier to monitor through structured data and reporting.

## Key Features

- **Automated Request Tracking:** Generate request IDs and record submission dates and times.
- **Email Notifications:** Route requests to the appropriate IT department using a configurable email mapping sheet.
- **Status Management:** Track requests through Pending, In Progress, and Resolved statuses.
- **Update Tracking:** Record the latest update date and time when request status or notes change.
- **Business Hours Calculation:** Calculate request age and resolution time based on defined working hours, excluding weekends and configured holidays.
- **Daily Reporting:** Generate an Excel report containing requests updated during the day.
- **Analytics Dashboard:** Organize request data to monitor workload, request status, department distribution, and resolution performance.
- **Configurable Settings:** Maintain department email mappings, SLA settings, and holiday information separately.

## Tools & Technologies

- Google Forms
- Google Sheets
- Google Apps Script (JavaScript)
- Spreadsheet Formulas and Data Validation
- Data Cleaning and Structuring
- Operational Reporting and Data Visualization

## Workflow

1. An employee submits an IT support request through Google Forms.
2. The form response is processed and transferred into the structured request tracker.
3. The system generates a request ID and assigns the request to the appropriate department.
4. An email notification is sent to the responsible team based on the department email mapping.
5. IT staff update the request status and add notes as the issue progresses.
6. The tracker records update dates and times and calculates request age and resolution time using business-hours logic.
7. Daily reports and the dashboard help monitor request activity and operational performance.

## Business Hours Logic

The project uses business-hours calculations to provide more meaningful measurements of request age and resolution time.

The defined working schedule is:

- **Working Days:** Sunday to Thursday
- **Working Hours:** 9:00 AM to 5:00 PM
- **Weekends:** Friday and Saturday are excluded.
- **Holidays:** Configured holidays are excluded from working-time calculations.

This approach is intended to measure working time rather than simply counting all elapsed hours, helping make request tracking and resolution reporting more relevant to business operations.

## AI-Assisted Development

AI tools assisted me in writing and refining parts of the Google Apps Script (JavaScript) code.

I used this assistance as part of the development process while working on the request workflow, automation logic, and reporting features. Testing and refinement are important parts of validating the implementation and ensuring the expected behavior.

## Sample Data & Privacy

This project uses **sample data for demonstration and testing purposes only**. The sample records do not represent actual employees or real IT support requests.

Any email addresses, employee details, and other potentially sensitive information should be replaced with fictional examples before publishing the project. The repository is intended to demonstrate the workflow and its analytical value without exposing private information.

## Business Value

This project demonstrates how Google Workspace tools and lightweight automation can help organizations:

- Reduce repetitive administrative work.
- Standardize IT support request tracking.
- Improve visibility into request status and department workload.
- Monitor request age and resolution performance.
- Organize operational data for reporting and analysis.
- Support more informed decisions through a structured dashboard.

## Project Status

Developed as a practical portfolio project, with testing and refinement of the automation, business-hours calculations, and reporting workflow.

## Future Improvements

- Enhance SLA monitoring and identify requests that exceed their targets.
- Expand dashboard metrics for workload and resolution performance.
- Improve validation and error handling across the automated workflow.

