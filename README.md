<div align="center">

# Spendwise

**Personal expense manager that lives entirely in your Google account.**

No server. No subscription. No third-party database. Your data stays in your own Google Sheets.

[![Made with Google Apps Script](https://img.shields.io/badge/Google_Apps_Script-4285F4?style=flat&logo=google&logoColor=white)](https://script.google.com)
[![Version](https://img.shields.io/badge/version-1.0.0-6FCF97?style=flat)](#)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

</div>

---

## What is Spendwise?

Spendwise is a full-featured personal finance tracker built as a Google Apps Script web app. You deploy it once to your Google account and it runs there permanently — no hosting costs, no external services, no data leaving your Drive.

The interface is a fast single-page app. The backend is pure Apps Script. All expenses, income, categories, and settings are stored in Google Sheets that you own and can read directly.

---

## Features

**Expense tracking**
- Add expenses with category, amount, date, payment method, and notes
- Edit and delete any entry inline
- Persistent category grid with emoji icons and custom monthly budgets

**Dashboard**
- Monthly spend total, average per day, top category at a glance
- Budget progress bars per category — color coded by percentage used
- 5 most recent transactions

**Analytics**
- 7 period presets: This Week, This Month, Last 30 Days, Quarter, Half Year, This Year, Last 12 Months
- Category breakdown with percentage of total
- Monthly comparison table — spend per category per month, with income and savings columns
- Spending trends chart

**Income tracking**
- Log income with categories (Salary, Bonus, Miscellaneous)
- Monthly summaries showing subtotals per category above the detail rows

**Weekly email report** *(optional)*
- Automatically sent on a day and time you choose
- Shows spend vs last week, top categories with visual bars, budget alerts, income summary
- Configured entirely from Settings — no code changes needed

**Settings**
- Currency symbol, default payment method, week start day
- Category editor — add, rename, reorder, set budgets
- Full data export to CSV across all historical data
- System panel — run SETUP, STATUS, and REPAIR from within the app

---

## Why Google Sheets as a database?

| Concern | Answer |
|---|---|
| **Data ownership** | Everything is in your Google Drive. Export, delete, or inspect it any time. |
| **Cost** | Free. Google Apps Script and Google Sheets have no usage cost at personal scale. |
| **Privacy** | No third party ever sees your financial data. |
| **Durability** | Google Sheets has versioning built in. You can roll back any accidental change. |
| **Transparency** | You can open the sheet and read the raw data directly — no black box. |

---

## How it works

```
Browser  ──── HTTPS ────▶  Google Apps Script (your account)
                                    │
                         ┌──────────┴──────────┐
                         │                     │
                   Config Sheet          Shard Sheets
                  (one, permanent)      (one per month)
                  ┌─────────────┐      ┌──────────────┐
                  │ Categories  │      │   Expenses   │
                  │ Settings    │      └──────────────┘
                  │ Income      │
                  │ Registry    │
                  └─────────────┘
```

Expenses are stored in **monthly shard sheets** — a separate Google Spreadsheet per month. This keeps performance fast as data grows and lets analytics skip irrelevant months entirely when computing date-range queries.

---

## Installation

### Prerequisites

- A Google account
- Access to [Google Apps Script](https://script.google.com)

### Steps

**1. Create a new Apps Script project**

Go to [script.google.com](https://script.google.com) → New Project.

**2. Add the files**

Delete the default `Code.gs` content. Create files matching the project structure and paste the contents of each file from this repository:

```
Code.gs
AdminOps.gs
index.html
shared-styles.html
shared-nav.html
page-add.html
page-dashboard.html
page-expenses.html
page-analytics.html
page-income.html
page-settings.html
```

To create an HTML file in Apps Script: click the **+** next to Files → HTML → enter the filename without the `.html` extension (Apps Script adds it automatically).

**3. Run SETUP()**

In the editor, open `AdminOps.gs`. Select `SETUP` from the function dropdown and click **Run**.

Apps Script will ask you to authorise the script — click **Review permissions**, choose your Google account, and click **Allow**.

SETUP() will:
- Create a **Config Sheet** in your Google Drive
- Create the first monthly **Shard Sheet** for the current month
- Seed all tabs with headers and default data
- Install a monthly auto-rotation trigger

Check the **Execution Log** — you will see the URLs of both created sheets.

**4. Deploy as a Web App**

- Click **Deploy** → **New Deployment**
- Type: **Web App**
- Execute as: **Me**
- Who has access: **Anyone** *(or "Anyone with Google account" for private access)*
- Click **Deploy** and copy the Web App URL

**5. Open the app**

Paste the URL into your browser. Spendwise is ready.

---

## Optional: Enable the Weekly Email Report

The weekly report requires Gmail send permission, which must be granted from the editor.

1. In the Apps Script editor, select any function (e.g. `STATUS`) and click **Run**
2. When the authorization dialog appears, click **Review permissions** → **Allow**
3. In the app, go to **Settings → Weekly Email Report**
4. Toggle **Enable**, enter your email address, choose day and time
5. Click **Save Report Settings**

The report fires automatically on your chosen schedule. Click **Send Test Email** to preview it immediately.

---

## First-run configuration

| Setting | Where | Recommendation |
|---|---|---|
| Currency symbol | Settings → General | Change from ₹ to your local symbol |
| Default payment | Settings → General | Set your most-used payment method |
| Categories | Settings → Categories & Budgets | Edit names, icons, and set monthly budget limits |

---

## Importing historical data

If you have existing expense data in a CSV, use `importFromCSV()` in `AdminOps.gs`:

1. Format your CSV with columns: `ID, Date, Category, Description, Amount, PaymentMethod, Notes, Timestamp`
2. Open `AdminOps.gs` and find `_getImportRows()` at the bottom of the file
3. Paste your CSV contents between the backticks
4. Run `importFromCSV()` from the editor — it skips rows with duplicate IDs
5. Run `fixShardRegistry()` to reorder the shard registry chronologically

---

## Maintenance

All routine maintenance is accessible from **Settings → System** in the app itself.

| Action | When | How |
|---|---|---|
| **STATUS** | Something seems wrong | Settings → System → Run STATUS |
| **REPAIR** | Expenses not loading, cache issues | Settings → System → Run REPAIR |
| **UPDATE** | After pulling a new version | Run `UPDATE()` in AdminOps.gs, then redeploy |
| **Rotate shard manually** | Auto-rotation missed | Settings → Data & Shards → Rotate Shard Now |

---

## Project structure

```
Code.gs              Runtime backend — all server functions
AdminOps.gs          Editor-only tools: setup, diagnostics, import
index.html           SPA shell — router, nav, floating add button
shared-styles.html   All CSS
shared-nav.html      Shared JS — toast, edit modal, payment methods, formatters
page-add.html        Add Expense page
page-dashboard.html  Dashboard
page-expenses.html   Expenses list with filters
page-analytics.html  Analytics and charts
page-income.html     Income tracker
page-settings.html   Settings, system panel, weekly report config
```

---

## Limitations

- **Single user** — designed for one Google account. Multi-user deployments are possible but not supported out of the box.
- **Apps Script quotas** — the free tier allows 6 minutes of script execution per day and 100 email recipients per day. For personal use this is never a constraint.
- **No real-time sync** — changes made in one browser tab don't auto-refresh another open tab.
- **No offline support** — requires an internet connection to the Google APIs.

---

## Tech stack

| Layer | Technology |
|---|---|
| Runtime | Google Apps Script (V8) |
| Frontend | Vanilla HTML / CSS / JS — single-page app |
| Charts | Chart.js (loaded lazily on Analytics page) |
| Backend | Apps Script server functions via `google.script.run` |
| Database | Google Sheets |
| Cache | `CacheService.getScriptCache()` |
| Email | `MailApp.sendEmail()` |
| Auth | Google account (implicit via deployment settings) |

---

## License

MIT — use it, modify it, build on it. A mention is appreciated but not required.