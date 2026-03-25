# hr-performance-review-consolidator-Automation-
Automating the consolidation of Employee Performance Reviews from Excel into a centralised Google Sheets master database 

_An automation that consolidate employe performance review data from signed Excel forms into a centralised Google Sheets master database, eliminating manual data collection while preserving the physical signature process that HR and complaiance depends on._

🧩 The Problem
In many organisations, HR teams conduct full-year performance reviews using structured Excel forms. While digital-first tools exist, Excel remains the preferred format for one critical reason:

_Both the Manager and Employee are required to physically sign the completed review document._

This makes a fully digital workflow (e.g. Google Forms, Typeform) impractical those tools cannot capture wet signatures or produce a printable, signable record that satisfies audit and compliance requirements.

As a result, HR team are left with a painful manual process every review cycle:
- Collecting dozens (or hundreds) of completed .xlsx files from employees.
- Opening each file one by one to read through responses.
- Manually copying data (names, ratings, comments, goals) into a master tracker, creating the Master data file.
- Hours of repetitive work prone to copy and paste errors and missed entries.
- No single source of truth for leadership to view team performance at a glance.

💡 The Solution
This project automates the extraction and consolidation of performance review data from completed Excel forms into a live Google Sheets master database, while keeping the Excel-based signing workflow fully intact.

HR simply drops completed .xlsx review files into a designated Google Drive folder. The automation handles the rest.

Employee fills & signs Excel form
        ↓
Manager signs Excel form
        ↓
File dropped into Google Drive folder  ←── HR's only manual step
        ↓
Apps Script extracts all fields automatically
        ↓
Master Google Sheet updated with new row (able to set timer)

⚙️ How It Works
Built entirely with Google Apps Script (no external dependencies, no servers, no cost).

1. Source Folder Monitoring: The scripts scans a designated Drive folder (Performance Reviews - Inbox) for new .xlsx files

2. Temporary Coversion: Each .xlsx is temporarily converted to a Google Sheet in memory using the Drive API, allowing cell-level data extraction.

3. Field Mapping: Specific cells are read based on the known structure of the Full Year Review template (employee details, all narrative sections, satisfaction scores and signature dates)

4. Master Sheet Update: Extracted data is appended as a new row in the master Google Sheet with a timestamp.

5. Duplicate Prevention: A processing log tav tracks file IDs so the name file is never processed twice.

6. Cleanup: The temporary Google Sheet copy is deleted immediately after extraction.

⏳ Scheduling
A time-based trigger runs extractAllReviews() automatically every night at midnight — no manual action needed after initial setup.

🛠️ Tech Stack
Tool                                    Role
Google Apps Script                      Core automation runtime
Google Drive API (Advanced Service)     .xlsx → Google Sheets conversion
Google Sheets                           Master database and processing log
Google Drive                            File storage and folder-based input trigger
Microsoft Excel (.xlsx)                 Original review form format (preserved for                                         signatures)

📁 Project Structure
├── PerformanceReviewAutomation.gs   # Main Apps Script file
└── README.md                        # This file

Setup Guide
**Prerequisites**

- A Google account with Google Drive access
- The Full Year Performance Review .xlsx template

Steps
1. Create the Apps Script project

Go to script.google.com → New project
Paste the contents of PerformanceReviewAutomation.gs
Save and name the project PR Automation

2. Enable the Drive Advanced Service

In the editor, click Services (+) in the left sidebar
Search for Drive API → click Add

3. Run initial setup

Select setupMasterSheet from the function dropdown
Click ▶ Run and grant permissions when prompted
This creates:

Master Performance Reviews Google Sheet in your Drive
Performance Reviews - Inbox folder in your Drive



4. Set up the nightly trigger (optional but recommended)

Select setupTrigger from the function dropdown
Click ▶ Run
New files dropped into the inbox folder will now be processed automatically every night

5. Process files manually anytime

Drop .xlsx review files into the Performance Reviews - Inbox folder
Run extractAllReviews() or wait for the nightly trigger

🔑 Key Design Decisions
Why keep Excel instead of migrating to Google Forms?
The organisation requires physical signatures from both the employee and the manager on the completed review document for compliance and record-keeping purposes. Excel allows the form to be printed, signed, scanned, and re-uploaded something a purely digital form cannot replicate.

Why Google Apps Script over Python/other tools?
Zero infrastructure. No servers, no credentials to manage, no third-party services. Everything runs inside the Google ecosystem that HR already uses daily. Any non-technical HR staff member can trigger the process with one click.

Why duplicate detection via file ID?
File names can change (e.g. an employee re-uploads a corrected version with a different name). Google Drive file IDs are permanent and unique, making them a more reliable deduplication key.

🙋 Author
Built as part of a data analytics portfolio to demonstrate practical automation solutions for real HR operational pain points.
