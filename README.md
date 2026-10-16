# Email Drafter

PowerShell script that creates Classic Outlook drafts from a CSV file. Each row is one store. `Recipients` lists the email addresses, separated by semicolons. PDFs whose file names contain the store name are attached.

The script never sends email. Review each message in Drafts before you send it.

## Run a dry run

```powershell
.\scripts\Create-EmailDrafts.ps1 -DryRun
```

A dry run validates the CSV, template, and attachment files. It does not open Outlook and does not create drafts.

## Create drafts

```powershell
.\scripts\Create-EmailDrafts.ps1
```

Classic Outlook is required. Batches of 25 or more drafts ask for confirmation. Pass `-Force` to skip that prompt.

## Sample files

- `input/Recipients.csv` contains two fictional stores.
- `input/settings.json` contains the CC addresses, sender names, and settlement date used in every draft.
- `attachments/` contains matching test PDFs. PDF files are gitignored.
- `templates/EmailTemplate.oft` is the Outlook template.

Regenerate the test PDFs with `.\scripts\New-SamplePdf.ps1`. Regenerate the template with `.\scripts\New-EmailTemplate.ps1`.
