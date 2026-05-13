# Troop 93 Eventa Plan Apps Script

## Using clasp

Get the Apps Script Script ID
Open your sheet.
Extensions -> Apps Script.
In Apps Script, open Project Settings.
Copy Script ID.
Clone the bound script into this workspace
In your workspace folder:
cd /Users/julianje/src/troop_events_plan
Clone:
clasp clone YOUR_SCRIPT_ID
This creates local project files such as Code.gs and appsscript.json.
Pull and push workflow
Pull latest from Google before edits:
clasp pull
After edits, push back:
clasp push
If prompted about overwriting, pull first, then re-apply your edits, then push.
Optional but useful local structure
Keep all Apps Script files in this folder.
Add a README only if you want team notes (not required for functionality).
Keep credentials and API keys out of committed files.
Test in Google Sheet
After push, return to Sheet and run a function once from Apps Script editor to authorize.
Recalculate formulas in the sheet and confirm Day/Sunset behavior.
Common issues and quick fixes:

Wrong Google account
Run clasp login again in the account that owns the sheet/script.
Permission denied on clone/push
Confirm that account has edit access to the bound Apps Script project.
Conflict errors
Run clasp pull first, resolve differences locally, then clasp push.
If you want, I can guide you interactively through each command one by one and troubleshoot any error output you get.