<!--
  Spendwise — personal expense tracker, budget tracker, income tracker
  Built with Google Apps Script and Google Sheets
  Free, self-hosted, no server, no database, no subscription
  Track expenses, manage budgets, view analytics, export CSV
  Alternative to Mint, YNAB, Toshl for Google Workspace users
-->

<div align="center">

# Spendwise

**Track your expenses and income — right inside your Google account.**

No server. No subscription. No third-party database.
Your data stays in Google Sheets that you own.

[![Google Apps Script](https://img.shields.io/badge/Google_Apps_Script-4285F4?style=flat&logo=google&logoColor=white)](https://script.google.com)
[![Version](https://img.shields.io/badge/version-1.3.0-6FCF97?style=flat)](#)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

</div>

---

Spendwise is a personal finance app that runs entirely on your Google account. You set it up once and use it like any web app — but everything is stored in your own Google Sheets and nothing ever leaves your Drive.

It's free, private, and yours to keep.

## What you get

**Expense & income tracking** — Log what you spend and earn. Categorise entries, set monthly budgets per category, and see how you're doing.

![Expense Form](https://raw.githubusercontent.com/nayanmehta03/spendwise-gas/main/screenshots/add_expense.png)

**Dashboard** — See your monthly total, average spend per day, budget usage per category, and recent transactions at a glance.

![Dashboard](https://raw.githubusercontent.com/nayanmehta03/spendwise-gas/main/screenshots/budgets.png)

**Analytics** — Compare spending across months, view category breakdowns, and spot trends with charts. Filter by week, month, quarter, half year, or full year.

![Analytics](https://raw.githubusercontent.com/nayanmehta03/spendwise-gas/main/screenshots/analytics.png)

**Standing instructions** — Add recurring expenses (rent, subscriptions, EMIs) and have them logged automatically on their due date. Every due date is tracked individually, so weekly and custom-interval items are logged every time they come round, and a run the trigger misses is caught up on the next one.

**Recurring reminders** — Get an email a day before anything recurring is due, with a **Log it** button that records the expense in one tap. Auto-debit items ride along as a heads-up, and anything still unlogged from the last 30 days is listed too. Configure it under Settings → Recurring Reminders.

**Weekly email report** — Get a summary of your week's spending delivered to your inbox. Shows top categories, budget alerts, income, and comparison to last week.

**Google Chat App Integration (New in v1.2.0)** — Log expenses directly from Google Chat using quick messages or slash commands like `/spend` and `/summary`. Receive a daily automated summary of your budget directly in your chat.

**Automated Gmail Receipt Ingestion (New in v1.2.0)** — Automatically scan your Gmail for purchase receipts (e.g. Swiggy, Zomato, Amazon, or custom keywords/senders) and import them. Features duplicate detection, keyword-to-category mapping, and robust execution logs.

**Desktop sidebar (New in v1.3.0)** — On screens 1024px and wider, a left rail lists every page with an icon and a short description, and content fills the width beside it. Collapse it to icons only and it stays that way next visit; narrower screens keep the hamburger menu.

**Full data export** — Download all your expenses as a CSV file any time from Settings.

---

## Upgrading from v1.1.0 (Keeping Data Intact)

Upgrading to v1.2.0 is **fully backwards compatible**. None of your existing expenses, categories, settings, or standing instructions will be touched.

Upgrading to v1.3.0 is **fully backwards compatible** too, and it is a front-end change only — the desktop sidebar. There is no sheet schema change, no new OAuth scope, nothing to re-authorize, and no need to re-run `SETUP()`. Update the HTML files, redeploy, and refresh. The steps below apply only if you are coming from v1.1.0.

### Step 1: Enable the manifest file
1. Open your existing Google Apps Script project at [script.google.com](https://script.google.com).
2. Click the gear icon (**Project Settings**) on the left sidebar.
3. Check the box **"Show 'appsscript.json' manifest file in editor"**.

### Step 2: Update the files
1. Go back to the Editor tab (**< >** icon).
2. Paste the updated contents for `Code.gs`, `AdminOps.gs`, `page-settings.html`, and `shared-styles.html` from this repository.
3. Create two new script files:
   - Click the **+** icon next to Files → **Script** → name it `ChatHandler` (do not add `.gs`) → paste `ChatHandler.js` content.
   - Click the **+** icon next to Files → **Script** → name it `EmailIngestion` (do not add `.gs`) → paste `EmailIngestion.js` content.
4. Open the `appsscript.json` file in the editor and replace its contents entirely with the project's `appsscript.json` (this adds the necessary OAuth scopes and enables the Chat service).

### Step 3: Run Setup / Repair
1. In the toolbar dropdown, select the function `SETUP` or `runRepair` and click **Run**.
2. This will securely verify your sheets and create three new configuration tabs in your Config sheet without modifying existing tabs:
   - `KeywordMap` (for auto-categorizing based on keywords)
   - `EmailSources` (for defining Gmail scan rules)
   - `Logs` (for chat and receipt import logs)
3. Redeploy your web app: Click **Deploy** → **Manage deployments** → click the edit pencil icon → select version **New Version** → click **Deploy**.
4. Refresh your web app. Go to Settings to view the new **Email Receipt Ingestion** and **Keyword Mappings** sections.

---

## Fresh Self-Hosting (Setup Guide)

You need a Google account. That's it.

### Step 1: Create a New Script Project
1. Go to [script.google.com](https://script.google.com) and click **New Project**.
2. Click the gear icon (**Project Settings**) on the left sidebar and check **"Show 'appsscript.json' manifest file in editor"**.
3. Go back to the editor.

### Step 2: Add Files
Create the following files in your Apps Script project and paste the contents from this repository:
- **Manifest**: `appsscript.json`
- **Script Files**: `Code.gs`, `AdminOps.gs`, `ChatHandler.gs`, `EmailIngestion.gs`
- **Shared HTML Files**: `index.html`, `shared-styles.html`, `shared-nav.html`
- **Page HTML Files**: `page-add.html`, `page-dashboard.html`, `page-expenses.html`, `page-analytics.html`, `page-income.html`, `page-standing.html`, `page-settings.html`

> *Note: In Apps Script, click **+** next to Files → select **HTML** or **Script** as appropriate. Type the name without extensions.*

### Step 3: Run SETUP
1. In the editor toolbar dropdown, select the function **SETUP** and click **Run**.
2. Authorize the script when prompted (click **Review permissions** → choose your account → **Advanced** → **Go to Spendwise (unsafe)** → **Allow**).
3. The script will create your database sheets automatically in Google Drive. Check the execution logs for success.

### Step 4: Deploy Web App
1. Click **Deploy** → **New Deployment**.
2. Select type: **Web App** (click gear icon next to "Select type" if Web App is not listed).
3. Set configuration:
   - Execute as: **Me**
   - Who has access: **Only myself** (Recommended for privacy).
4. Click **Deploy** and copy the **Web App URL**. Open it to access Spendwise!

> Opening the app at least once after deploying also activates the **Log it** buttons in recurring reminder emails — that first visit is how Spendwise learns its own deployment URL. If you ever redeploy to a new URL, delete the `WEBAPP_URL` script property and open the app again. `STATUS` warns you if the buttons aren't active yet.

---

## Configuring Email Receipt Ingestion

Gmail ingestion allows Spendwise to scan your inbox, parse receipt details, and automatically log expenses.

### 1. Configure Keyword Mappings
Go to Settings → **Keyword Mappings** to link keywords in transaction descriptions to categories. For example:
- `swiggy` → `Food & Drink`
- `zomato` → `Food & Drink`
- `amazon` → `Shopping`
- `uber` → `Transport`

### 2. Configure Email Sources
Go to Settings → **Email Receipt Ingestion** → click **+ Add Source** and configure:
- **Sender**: Email address of the vendor (e.g. `noreply@swiggy.in`).
- **Subject Filter**: Substring or text that matches the receipt email (e.g. `Order confirmation`, `Your Amazon.in order`).
- **Parser**: Select a specialized parser (`Swiggy`, `Zomato`, `Amazon`) or select `Generic Keyword Categorizer` for other vendors.
- **Enabled**: Check to activate the scanning rule.

### 3. Save & Enable
1. Enable the **Enable Email Ingestion** toggle switch.
2. Select your desired **Scan Interval** (e.g. every 15 min, 30 min, or every 60 min).
3. Click **Save Email Settings**. This automatically registers a background time-based trigger in your Apps Script account.
4. Click **Run Import Now** to test the scan immediately.
5. Click **View Logs** to verify that emails were processed and see if any transactions were imported or skipped as duplicates.

---

## Configuring Google Chat App

Link Spendwise to Google Chat to add expenses on-the-go or get daily budget progress summaries.

### 1. Link to a Google Cloud Project (GCP)
1. Open your Apps Script project editor.
2. Click Project Settings (gear icon) and copy your **Project Number** under GCP Project (if using a default project, you will need to link it to a standard GCP Project by clicking **Change project** and entering a GCP project ID).
3. Open the [Google Cloud Console](https://console.cloud.google.com/) for that project.

### 2. Enable Google Chat API
1. In Cloud Console, search for **Google Chat API** and click **Enable**.
2. Go to the **Configuration** tab of the Google Chat API.

### 3. Configure Chat Settings
Fill in the configuration fields:
- **App name**: Spendwise
- **Avatar URL**: (Use any image URL, e.g. a green wallet icon)
- **Description**: Spendwise Expense Tracker Bot
- **Interactive features**: Enable/Turn ON
- **Functionality**: Check **Receive 1:1 messages** (so you can DM the bot) and optionally **Join spaces**
- **Connection settings**: Select **Apps Script project** and paste your Apps Script **Deployment ID** (retrieve this from Apps Script editor: Deploy → Manage deployments → copy the Active Deployment ID).
- **Slash commands**: Add the following commands:
  - `/spend` (Description: `Log an expense. Format: /spend <amount> <desc> [category]`)
  - `/summary` (Description: `Get budget summary`)
  - `/help` (Description: `Show help guide`)
- Click **Save**.

### 4. Chat Commands
Open Google Chat, search for **Spendwise**, and start a conversation. Try these commands:
- `/spend 120 Uber ride to work` — Automatically logs `120` currency units, categorizes under `Transport` (using keyword mappings), and saves notes with a Chat reference.
- `/spend 500 grocery shopping Groceries` — Logs `500` under the `Groceries` category explicitly.
- `/summary` — Shows your total monthly spend, budget usage per category, and remaining balance.
- `/help` — Lists all commands.

### 5. Daily Chat Summary (Opt-in)
Go to your Settings sheet (or set `dailySummaryEnabled` to `true` and configure `dailySummaryTime` in the settings) and run `installDailySummaryTrigger()` from the script editor. This will push your budget summary directly to you in Google Chat every evening.

---

## Good to know

- **Safe Upgrades**: The SETUP and REPAIR scripts detect if your configuration is already present and will never overwrite, clear, or modify your existing transaction data.
- **Privacy First**: All Gmail scanning and Google Chat communication run entirely within your Google account. No external servers or database queries are performed.
- **Google Limits**: Apps Script triggers run on Google's free tier. Mail scanning checks read-only metadata and skips processing for messages already seen, ensuring you stay well within daily quota limits.

---

## Developing

If you'd rather edit locally than paste into the Apps Script editor, the project pushes with [clasp](https://github.com/google/clasp).

```bash
npm install -g @google/clasp
clasp login
clasp clone-script <your-scriptId>     # writes .clasp.json (gitignored)
```

Then set up your deployment target once — copy `deploy.config.example.json` to `deploy.config.json` and paste in your own deployment ID (the `AKfyc…` part of your web app URL, found under **Deploy → Manage deployments**). That file is gitignored, so your IDs never end up in a commit.

```powershell
.\deploy.ps1                                # push, version, redeploy
.\deploy.ps1 -DryRun                        # show what would be pushed
.\deploy.ps1 -Description "analytics fix"   # label the version
```

`deploy.ps1` always redeploys to the deployment ID you configured, so the live URL never changes — bookmarks, the Google Chat connection, and the **Log it** buttons in reminder emails keep working. Plain `clasp deploy` creates a *new* deployment with a *new* URL instead, so avoid it. From Git Bash, `./deploy.sh` forwards to the same script.

A few things worth knowing before you change files:

- **`.claspignore` is a whitelist.** It ignores everything, then re-includes each shipped file by name. A new `.js` or `.html` file won't deploy until you add a `!filename` line — `deploy.ps1` warns you when it spots one.
- **No build step and no modules.** Every `.js` file shares one global namespace server-side, so `import`/`require` don't exist and a duplicate function name across two files breaks the whole project.
- **Adding an OAuth scope** means editing `appsscript.json` *and* re-authorizing the app — it fails silently otherwise.

Contributor conventions live in `CLAUDE.md`.

---

## License

MIT — use it however you want.
