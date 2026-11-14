# Apps Script

Collection of Google Apps Script projects and utilities.

## Projects

- **[Movie Tracker](./Movie%20Tracker/README.md)**: Automated Google Sheets movie watchlist & release tracker using IMDb, Rotten Tomatoes (Popcornmeter), OMDb, DVDsReleaseDates, Wikidata, and Wikipedia.
- **Asana Unicorn Image / Trigger**: Asana celebration image and trigger helpers.
- **Claude Spend Dashboard**: Usage and spend tracking dashboard.
- **Employee Tenure Leaderboard**: Employee tenure tracking script.
- **FedEx-UPS-Tracking-Status**: Shipment status tracking helper.
- **Google Calendar Color**: Calendar event color automation.
- **Google Group Inspector**: Group membership inspection utilities.
- **Import UKG Report to Sheet / Import Workday Report to Sheet**: HR report importers.
- **LostAssetEmail.gs / Send Email on Form Submit**: Email notification automations.
- **Menubar function / Weather Forecast / duplicateSheetWithProtection / generatePDF / sortOnEdit**: General Google Sheets & Workspace utilities.

## Secrets & Configuration

- Never commit `.clasp.json`, `.env`, API keys, tokens, or spreadsheet IDs.
- Use `.clasp.json.example` as a template for local `clasp` deployments and store secrets in `PropertiesService.getScriptProperties()`.
