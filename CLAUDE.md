# CLAUDE.md — working rules for Spendwise

Read this before touching anything. These rules exist because of three hard
constraints: **the repo is public**, **the runtime is Google Apps Script**, and
**one fixed deployment serves the live app**.

---

## 1. The repo is public

`github.com/nayanmehta03/spendwise-gas` is public. Anything committed is
world-readable **forever**, including in history. A leaked value is not fixed by
a follow-up commit.

**Never commit:**

- Script IDs, deployment IDs, spreadsheet IDs, Drive file IDs
- OAuth tokens, API keys, client secrets, `.clasprc.json`
- Real email addresses, real transaction data, real amounts, personal budgets
- Screenshots showing real numbers — scrub or use sample data first

**Before writing any file, ask: is this public or internal?**

| Public (committed) | Internal (gitignored) |
| --- | --- |
| `README.md` — user-facing setup & feature docs | `docs/**` — briefs, specs, research, plans |
| `*.js`, `*.html`, `appsscript.json` | `DESIGN_SPEC.md` |
| `deploy.ps1`, `deploy.sh`, `deploy.config.example.json` | `deploy.config.json`, `.clasp.json` |
| `.claspignore`, `.gitignore`, `CLAUDE.md` | `.claude/**`, `.vscode/**`, `.deploy-history.log` |

`.gitignore` is the single source of truth for that split. If you create a new
kind of file that holds account-bound or in-flux content, **add it to
`.gitignore` in the same change** — do not rely on remembering later.

Placeholders in committed files use the obvious-fake form:
`AKfyc__REPLACE_WITH_YOUR_DEPLOYMENT_ID__`, `you@example.com`.

---

## 2. The runtime is Google Apps Script

There is **no build step, no bundler, no npm, no `node_modules`**. What is in
the repo root is what runs.

- **No `import` / `export` / `require`.** All `.js` files share one flat global
  namespace on the server. A function declared in `AdminOps.js` is callable from
  `Code.js` with no ceremony — and a duplicate declaration across two files is a
  hard project-wide error.
- `.js` becomes `.gs` server-side. `.html` files are `HtmlService` templates.
- The frontend is a single-page shell: `index.html` pulls in every page and
  `shared-styles.html` via `<?!= include('name') ?>`. Client talks to server
  only through `google.script.run`.
- Runtime is **V8** (`appsscript.json`), so modern syntax is fine — but browser
  and Node globals are not. No `fetch`, no `localStorage` server-side; use
  `UrlFetchApp`, `PropertiesService`, `CacheService`.
- Adding an OAuth scope means editing `appsscript.json` **and** telling the user
  they must re-authorize — the app silently fails on a missing scope otherwise.

### File map

| File | Role |
| --- | --- |
| `Code.js` | Runtime backend — every function the web app calls |
| `AdminOps.js` | Editor-only: `SETUP()`, `STATUS()`, `RESET()`, trigger installers, diagnostics. Never called at runtime |
| `ChatHandler.js` | Google Chat app: `/spend`, `/summary`, daily summary |
| `EmailIngestion.js` | Gmail receipt import; parser-registry pattern |
| `index.html` | SPA shell, router, nav, FAB |
| `shared-styles.html` | All CSS |
| `shared-nav.html` | Overlay nav, toast, edit modal, shared client JS |
| `page-*.html` | One per screen |
| `shard-functions.js` | **Legacy. Excluded from push** — duplicates `AdminOps.js` functions. Do not revive |

### `.claspignore` is a whitelist

It ignores `**/**` and then re-includes each shipped file by name. **A new `.js`
or `.html` file will silently not deploy until you add a `!filename` line.** If
a change appears to have no effect live, check this first.

---

## 3. Deploying

One deployment serves the live app. Its ID is in `deploy.config.json`
(gitignored). Bookmarks, the Google Chat app connection, and the "Log it"
buttons in reminder emails all point at that URL.

```powershell
.\deploy.ps1                                  # push + version + redeploy
.\deploy.ps1 -DryRun                          # show the push set, change nothing
.\deploy.ps1 -Description "analytics fix"     # label the version
.\deploy.ps1 -SkipPush                        # redeploy what is already pushed
```

**Never run bare `clasp deploy` or `clasp create-deployment` without `-i`** — it
mints a *new* deployment with a *new* URL and silently orphans the live one.
`deploy.ps1` is the only sanctioned path; it verifies auth, blocks legacy files
from the push set, and warns before shipping uncommitted work.

**Do not deploy unless the user asks.** Deploying is outward-facing and
immediately affects the live app the user depends on daily.

If clasp reports `invalid_grant` / `invalid_rapt`, the token expired — the user
must run `clasp login` themselves. You cannot complete that OAuth flow.

---

## 4. Document everything

Every feature ships with its docs updated in the same change — not "later".

- **`README.md`** is the only user-facing doc. Update it when behaviour a user
  can see changes: a new feature, a new setup step, a new sheet tab, a new
  scope, a changed command. Match its existing voice — plain, second person, no
  marketing filler. Version-gated changes get an "Upgrading from vX" section.
- **`docs/`** is the internal workspace and is gitignored. Design specs,
  research, feature plans, and agent briefs go here. Content is *promoted* to
  `README.md` when it becomes something a user needs.
- **Code headers.** Every top-level `.js` file opens with a banner listing its
  role and entry points (see `Code.js`). Keep it accurate — it is the fastest
  map anyone gets. New exported-in-practice functions get a line there.
- Comments explain **why**, not what. Match the surrounding density; this
  codebase comments decisions and gotchas, not obvious statements.
- Bump the version in the file banner and the README badge together.

---

## 5. Working agreements

- **Multiple agents may be active on this repo at once** (UI work often runs in
  a separate session). Before editing a shared file — `shared-styles.html`,
  `index.html`, `shared-nav.html` — check `git status` and prefer narrow,
  targeted edits over rewrites.
- **Do not commit or push unless asked.** Report what changed and let the user
  decide.
- **Never `git add -A` / `git add .`** — stage named files. A blanket add is how
  an ignored-but-forced file reaches a public repo.
- Data safety is non-negotiable: `SETUP()` and `REPAIR` must stay idempotent and
  must never overwrite, clear, or reorder existing user rows. Any change to
  sheet structure needs a migration path that leaves old data intact.
- Prefer editing an existing function over adding a near-duplicate — the flat
  namespace makes drift expensive.
- Ask before: adding an OAuth scope, changing sheet schema, or anything that
  requires the user to re-run setup or re-authorize.
