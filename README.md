# Scout Troop 93 Events Plan Apps Script

To facilitate Scout Troop annual planning while working with our PLC (Patrol Leaders Council), I found it helpful to rapidly create and maintain a Google Sheet that automatically determines the day(s) of the week based on a date or date-range and automatically looks up the time of sunset to help us plan events that require daylight or darkness. This lookup is cached and only runs on date-changes for speed compared to a basic formula that constantly re-runs when opening a sheet. The lookup is performed based on our Lat/Long, but that could be enhanced to get calculated from a zipcode or similar in the future.

You are welcome to take this and use it. If you develop enhancements, I will consider incorporating those if you provide a PR.

I make no guarantee or warrantee on the accuracy or performance. This is supported best-effort in my hour-a-week ;-) with Scouting.

This repository keeps the Google Sheet Apps Script source under version control while leaving the local clasp project config untracked so that your worksheet ID is not shared. `Clasp` is a tool that will pull or push Apps Script code from/to your configured Google sheet to enable advanced functions and calculations beyond what native formulas provide.

## Repo layout

- `src/Code.js` and `src/appsscript.json` are committed source files.
- `.clasp.json` is local-only and ignored by git.
- `.clasp.json.example` is a committed template you can copy to `.clasp.json` after cloning.
- `.env` is reserved for any local-only secrets or overrides if you ever add them.

## Setup

1. Clone the repo.
2. Copy `.clasp.json.example` to `.clasp.json`.
3. Replace `YOUR_SCRIPT_ID_HERE` with your bound Apps Script project ID.
4. Confirm `.clasp.json` has `"rootDir": "src"`.
5. Run `clasp login` if needed.
6. Run `clasp pull` or `clasp push` from the repository root.

## Working with clasp

Because `rootDir` points to `src`, clasp reads from and writes to the `src` folder automatically.

Use `clasp pull` before editing if the sheet project may have changed.
Use `clasp push` after making local changes.
If you see a conflict, pull first, reconcile locally, then push.

Example workflow:
1. `clasp pull`
2. edit `src/Code.js`
3. `clasp push`

## Notes

- Keep credentials and API keys out of committed files.
- For a public repo, do not commit `.clasp.json`.
- The script ID is not a secret, but keeping it local avoids binding the public repo to one live project.