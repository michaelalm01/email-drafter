# Design Plan: PowerShell + Outlook Email Draft Automation

## 1. Objective

Build a Windows-based PowerShell automation script that reads recipient and document information from a CSV file, generates personalized Outlook email drafts using a reusable email template, attaches the appropriate PDF files, and saves the drafts in Outlook for manual review.

**The script must never send emails automatically.** All generated messages must remain in Drafts until a user manually reviews and sends them.

## 2. Technology Stack

- Operating System: Windows 10/11
- Automation: PowerShell
- Email client: Classic Microsoft Outlook for Windows
- Email template: Outlook template (`.oft`)
- Input data: CSV file (`.csv`)
- Attachments: PDF files stored in a designated local folder
- Integration: Outlook COM automation

Note: Outlook COM automation requires Classic Outlook. New Outlook does not support this integration method.

## 3. Project Structure

Create the following folder structure:

```text
EmailAutomation/
├── scripts/
│   └── Create-EmailDrafts.ps1
├── templates/
│   └── EmailTemplate.oft
├── input/
│   ├── Recipients.csv
│   └── settings.json
├── attachments/
│   ├── Collection Reminder  - Northwind Traders, Inc..pdf
│   └── Collection Reminder  - Contoso Supplies.pdf
├── logs/
│   └── execution.log
└── README.md
```

## 4. Input CSV

The CSV file must contain one row per store. The store name is the anchor for the draft. A header row names the columns. Values that contain commas or line breaks must be quoted.

Required columns:

| Column | Description | Example |
|---|---|---|
| StoreName | Store the draft is for | Northwind Traders, Inc. |
| Recipients | Recipient email addresses, separated by semicolons | alex@example.com; billing@example.com |

Attachment filenames follow `Collection Reminder  - {Store Name}.pdf`. The script attaches every PDF in the attachments directory whose file name contains the store name.

Each row produces one email draft. Every address in `Recipients` is placed on that draft.

`input/settings.json` supplies the shared draft values: `cc`, `senderNames`, and `settlementDate`. Every address in `cc` is added to the CC field of every draft. `senderNames` are joined with ` / ` and, with `settlementDate`, replace the matching placeholders in the template.

## 5. Email Template Design

Create the template using Classic Outlook's email editor and save it as `templates/EmailTemplate.oft`.

### Subject

```text
Collection Reminder - {Store Name}
```

### Email Body

Dear Sir/Madam,

Please see attached Collection Reminder regarding your outstanding account.

We request your immediate attention and settlement on or before {Settlement Date}.

Thank You.

Kind Regards,
{Sender Names}

### Template Requirements

- Professional business tone.
- Consistent font, spacing, and formatting.
- Support HTML formatting, hyperlinks, and company branding.
- Preserve the original template without modifying it during processing.
- Replace placeholders with values from the corresponding CSV row.
- Retain the standard email signature where required.

Implementation note: Outlook `.oft` files do not automatically replace arbitrary placeholders. The PowerShell script must explicitly replace each placeholder in the generated message.

## 6. Processing Workflow

1. Load and validate the CSV file.
2. Confirm that all required columns exist.
3. Validate recipient email addresses and required values.
4. Load the Outlook template.
5. Process each CSV row independently.
6. Create a new Outlook message from the template.
7. Populate the recipient, CC list, subject, and email body.
8. Replace all supported placeholders with row-specific values.
9. Find PDF files in the attachments directory whose names contain the store name.
10. Attach each matching PDF when it passes validation.
11. Save the message to Outlook Drafts.
12. Log the result for that row.
13. Display a final summary showing successful drafts and errors.

## 7. Attachment Handling

- Search only the configured attachments directory.
- Attach every PDF whose file name contains the store name.
- Each draft requires at least one matching PDF.
- Reject the row when no matching PDF exists rather than creating an email without the collection reminder.
- Attach only files that are already in the attachments directory.

## 8. Safety and Error Handling

The automation must:

- Never invoke Outlook's Send operation.
- Save every successfully generated email using the Drafts folder.
- Leave the original template unchanged.
- Skip invalid rows and continue processing other valid rows.
- Record errors, missing attachments, and successful draft creation.
- Avoid logging unnecessary personal information or full email bodies.
- Warn about duplicate store names.
- Require confirmation before processing a large batch.
- Provide a dry-run mode that validates the input without creating drafts.

## 9. Logging and Reporting

Write execution details to `logs/execution.log`.

Record:

- Execution timestamp.
- Number of CSV rows processed.
- Number of drafts created.
- Number of rows skipped.
- Missing attachment errors.
- Validation failures.
- Unexpected Outlook or PowerShell exceptions.

Display a summary at the end of each execution.

## 10. Implementation Phases

### Phase 1 — Environment Setup
- Verify that Classic Outlook is installed and configured.
- Create the project directory structure.
- Prepare the CSV file and attachment folder.
- Create and save the Outlook template.

### Phase 2 — Basic Draft Creation
- Connect to Outlook using PowerShell COM automation.
- Load the template.
- Create a single test draft.
- Populate recipient, subject, and message body.
- Save the draft without sending.

### Phase 3 — CSV Integration
- Read the CSV input.
- Map CSV columns to template placeholders.
- Generate a separate draft for each valid row.

### Phase 4 — Attachment Integration
- Match each store row to PDF files whose names contain the store name.
- Validate that each match is a PDF in the attachments directory.
- Attach every matching PDF to that store's draft.

### Phase 5 — Reliability
- Implement validation, logging, exception handling, duplicate detection, and dry-run mode.
- Test missing files and malformed CSV rows.
- Confirm that no execution path sends an email.

### Phase 6 — Acceptance Testing
- Test with two or three sample recipients.
- Verify recipient, subject, body formatting, and attachment contents.
- Confirm that the drafts appear in Outlook's Drafts folder.
- Manually inspect the messages before sending.
- Confirm that the original template remains unchanged.

## 11. Acceptance Criteria

The project is complete when:

1. The script can read a valid CSV file.
2. Every valid row creates exactly one draft.
3. Each draft contains the correct recipient and subject.
4. All supported placeholders are replaced.
5. Every PDF whose file name contains the store name is attached to that store's draft.
6. Invalid rows and missing attachments are reported.
7. The original template is preserved.
8. Drafts are saved in Outlook and are available for manual review.
9. No emails are sent automatically.
10. A clear execution summary and log are produced.

## 12. Future Enhancements

Potential extensions include:

- Multiple templates selected by `EmailType`.
- Additional attachment types beyond the collection reminder.
- PDF text extraction to identify document references automatically.
- Automatic matching between PDF filenames and CSV records.
- A simple Windows graphical interface for selecting files and starting the process.
- Preview mode showing the generated email before saving.
- Duplicate-draft detection across repeated runs.
- Optional integration with Microsoft Graph for environments using New Outlook.

## 13. Initial Scope

The first implementation should focus on Classic Outlook, one `.oft` template, one CSV file with one row per store, the recipients and attachment filenames for that store, and draft creation only.

Keep the script modular so CSV parsing, template replacement, attachment validation, and Outlook integration can be tested independently.
