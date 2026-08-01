# The Ledger — Setup Guide (Google Sheets edition)

This version runs on **Google Sheets + Apps Script** — the simplest, most
reliable option: your data lives in a Sheet you can open directly, and the
whole backend is one script file, no Firebase account, no billing plans,
no card, ever.

If you set this up before (an earlier version of this app), you can reuse
the exact same Google Sheet — just redeploy the new `AppsScript.gs`.
Nothing destructive happens to existing data; new columns/tabs are added
automatically as needed.

---

## 1. Set up the Google Sheet (5 minutes)

1. Go to **sheets.new** to create a blank Sheet.
2. **Extensions → Apps Script**. Delete the placeholder code, paste in the
   entire contents of `AppsScript.gs`.
3. Gear icon (**Project Settings**) → **Script Properties** → add:
   - `ADMIN_CODE` = any password you make up (e.g. `cedi2026`) — this is
     your personal Owner login code. Keep it somewhere safe.
   - `APP_URL` = leave blank for now (optional — fill in once hosted, so
     invite links work; see step 4).

## 2. Deploy the script (2 minutes)

1. **Deploy → New deployment → Web app**
2. Execute as: **Me** · Who has access: **Anyone**
3. Click **Deploy**, approve the permissions prompt (it's your own script)
4. Copy the URL — it ends in `/exec`. This is your Sheet's connection URL.

Whenever you update `AppsScript.gs` later: **Deploy → Manage deployments**
→ edit (pencil icon) → Version: **New version** → Deploy. This keeps the
same URL and just updates the code running behind it — a very common
Apps Script gotcha is editing the code and expecting it to take effect
without doing this step.

## 3. Host the app files (5 minutes)

Any static host works. Two easy options:

**GitHub Pages:**
1. Create a repo, upload `index.html`, `manifest.json`, `sw.js`, and the
   `icons/` folder into the repo root
2. Repo **Settings → Pages** → deploy from `main` branch, root
3. You'll get a URL like `https://yourname.github.io/ledger-app/`

**Netlify (drag-and-drop, fastest):**
1. Go to **app.netlify.com/drop**
2. Drag the folder (containing `index.html`, `manifest.json`, `sw.js`,
   `icons/`) onto the page
3. You get a live URL immediately

Either way — once it's live, open it on your phone and use "Add to Home
Screen" from the browser menu to install it like a real app.

## 4. First login

Open the hosted URL. On the login screen enter:
- **Sheet URL**: the `/exec` link from step 2
- **Name / Email**: whatever you like
- **Access code**: your `ADMIN_CODE`

You're now the Owner.

## 5. Invite people

**People** tab (Owner only) → **+ Invite someone** → email + role
(Viewer or Editor). They automatically get an email with their own access
code — and, if you filled in `APP_URL` in step 1, a direct link that logs
them straight in.

---

## What's in this version

- **Transactions** — Income, Expense, and **Transfer** (pick two
  accounts, one save auto-creates the linked debit/credit pair), tags,
  recurring (weekly/monthly/yearly, auto-generated), receipt photos
- **Accounts** — unlimited wallets/accounts, each with its own currency;
  tap one to see just its transactions
- **Categories** — default list plus your own
- **Budgets** — monthly limits per category, with live spent-so-far
- **Goals & Planning** — savings goals and one-off planning budgets
- **Loans** — borrowed/lent, repayment log, outstanding balance
- **Investments** — cost basis vs current value, gain/loss
- **Scale of Preference** — rank items by priority against a budget;
  **drag to reorder, or use the ↑↓ buttons** — both work, arrows are the
  reliable option on touchscreens where drag-and-drop can be finicky
- **Insights** — simple rule-based alerts (spending trends, budgets
  nearing/over, loan due dates) — arithmetic on your own numbers, not AI
- **Reports** — income statement / cash flow by date range, Print → Save
  as PDF, plus CSV export
- **Multi-currency** — Settings → fetch live exchange rates (free, no
  key) or enter your own; Dashboard converts everything to your base
  currency automatically
- **Dark mode** and an optional **PIN lock** (device-only convenience
  lock, with a "Forgot PIN? Turn it off" escape hatch — it can never
  permanently lock you out)
- **Hide balances** toggle (👁 icon on Dashboard) for a privacy mode
- **Excel/CSV backup** — export everything to one workbook any time, or
  import one back in

## Receipt photos

Stored via Google Drive — Sheets itself can't hold binary files, but
Drive can, in a "Ledger Receipts" folder auto-created the first time
someone uploads one. Anyone with Editor access or higher can attach a
photo; everyone with app access can view it.

## Sync, offline, and multiple people at once

The app saves to this device first (works offline, nothing is lost),
then pushes to the Sheet in the background. It also quietly re-checks
the Sheet every 45 seconds while open, and immediately when the device
comes back online, so everyone sees roughly-current data without doing
anything. Two people editing the exact same field at the exact same
moment: last save wins — normal behavior for this kind of sync, not a
bug.

## Troubleshooting

- **"Invalid email or access code"** — almost always means either (a)
  you edited the script but didn't redeploy a new version (see step 2),
  or (b) there's already a row for that email in the Sheet's Users tab
  from an earlier attempt — check there first.
- **Blank/broken page, only works after a reload** — hard-refresh once;
  the service worker is set to always fetch the freshest version, so
  this shouldn't recur after the first load post-deploy.
- **Wallet/transaction dropdown looks empty** — add at least one Account
  first (Accounts tab) before adding transactions, budgets, etc. — they
  all need an account to attach to.

## Costs

Entirely free. Google Sheets, Apps Script, and Google Drive have no
billing tier at all for this kind of personal/small-business usage —
there's no card to add anywhere in this setup.
