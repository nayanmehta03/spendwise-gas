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
[![Version](https://img.shields.io/badge/version-1.1.0-6FCF97?style=flat)](#)
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

**Standing instructions** — Add recurring expenses (rent, subscriptions, EMIs) and have them logged automatically on their due date.

**Weekly email report** — Get a summary of your week's spending delivered to your inbox. Shows top categories, budget alerts, income, and comparison to last week.

**Full data export** — Download all your expenses as a CSV file any time from Settings.

---

## Self-hosting (setup)

You need a Google account. That's it.

1. Go to [script.google.com](https://script.google.com) and create a **New Project**
2. Delete the default code. Create one file for each file in this repo and paste the contents:
   - `Code.gs` and `AdminOps.gs` (script files)
   - `index.html`, `shared-styles.html`, `shared-nav.html` (shared files)
   - `page-add.html`, `page-dashboard.html`, `page-expenses.html`, `page-analytics.html`, `page-income.html`, `page-standing.html`, `page-settings.html` (pages)

   > To add an HTML file: click **+** next to Files → **HTML** → type the name without `.html`

3. Open `AdminOps.gs`, select **SETUP** from the function dropdown, and click **Run**
4. Google will ask you to authorise — click **Review permissions** → choose your account → **Allow**
5. Check the **Execution Log** for confirmation. Two Google Sheets will be created in your Drive automatically.
6. Click **Deploy** → **New Deployment** → type: **Web App** → set:
   - Execute as: **Me**
   - Who has access: **Only myself**
7. Click **Deploy** and copy the URL. Open it. Done.

> **What does "Who has access" mean?**
>
> This controls who can *open* the web app URL — it does **not** share your spreadsheet data.
>
> | Option | Who can open the app | When to use |
> |---|---|---|
> | **Only myself** | Just your Google account | **Recommended.** Personal use. |
> | **Anyone with Google account** | Anyone signed into Google, but they see *your* data since it runs as you | Only if you want to share the app with family/friends |
> | **Anyone** | Opens without sign-in | Not recommended — anyone with the URL can see your expenses |
>
> In all cases the app runs as **you** and reads **your** sheets. "Only myself" is the safest choice.

### After setup

**Change your settings** — Go to Settings in the app to set your currency, default payment method, and edit categories/budgets.

**Enable weekly email reports** — Go to Settings → Weekly Email Report, toggle it on, and enter your email. You'll need to grant Gmail permission once from the script editor (Run any function → Allow when prompted).

**Import existing data** — If you have historical expenses in CSV format, open `AdminOps.gs`, paste the CSV contents into the `_getImportRows()` function, and run `importFromCSV()`.

**Something broken?** — Go to Settings → System and run **STATUS** (to check what's wrong) or **REPAIR** (to fix it). These are safe to run any time.

---

## Why Google Sheets?

- **It's free** — no hosting costs, no subscriptions
- **You own your data** — it's in your Google Drive, readable as a normal spreadsheet
- **It's private** — no third party sees your financial data
- **Built-in backups** — Google Sheets tracks version history automatically

---

## Good to know

- Designed for **one person** (single Google account)
- Needs an **internet connection** — no offline mode
- If you have the app open in two tabs, changes in one won't auto-show in the other — just refresh
- Google Apps Script has daily limits (6 min execution, 100 emails) but personal use never hits them

---

## License

MIT — use it however you want.
