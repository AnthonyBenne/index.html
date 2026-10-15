# Website scheduler (Google Apps Script)

Backup of the "Ben Anthony Website Scheduler" Apps Script project (office@benanthony.com Drive).
The live site calls its web-app deployment from index.html.

- `Code.gs`: GET returns open 30-min slots; POST books the call on Google Calendar and logs the lead to the "Ben Anthony CRM - Leads" sheet.
- `appsscript.json`: project manifest (Calendar v3 advanced service, runs as owner, anyone can call).

To deploy a change: paste `Code.gs` into the script editor, save, then Deploy > Manage deployments > edit the existing deployment > New version. Keep the same deployment so the URL in index.html doesn't change.
