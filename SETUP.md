# The Ledger — Setup Guide

You have a folder now, not a single file — that's what makes it a real installable app:

```
index.html      the app
manifest.json   tells the phone/browser how to install it
sw.js           lets it work offline
icons/          app icon
AppsScript.gs   the Google Sheets bridge (deploy this separately, see below)
```

## Part 1 — Put the Google Sheet in place (5 minutes)

1. Go to **sheets.new** to create a fresh Google Sheet.
2. **Extensions → Apps Script**. Delete the placeholder code, paste in the entire contents of `AppsScript.gs`.
3. Gear icon (**Project Settings**) → **Script Properties** → add:
   - `ADMIN_CODE` = any password you make up, e.g. `cedi-2026-owner` — this is *your* permanent access code as Owner. Keep it safe.
   - `APP_URL` = leave blank for now; come back and fill this in once you've hosted the app in Part 2, so invite emails can include a one-click link.
4. **Deploy → New deployment → Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
   - Deploy, authorize the permissions (it's your own script on your own sheet), copy the URL — it ends in `/exec`.

## Part 2 — Host the app so it's installable

A `file://` HTML can't be "installed" — phones and browsers only offer that for something served over HTTPS. The good news: static hosting is free and takes a couple of minutes. Two easy options:

**Netlify (drag-and-drop, no account needed to try):**
1. Go to app.netlify.com/drop
2. Drag the whole folder (index.html, manifest.json, sw.js, icons/) onto the page
3. You get a live HTTPS URL immediately

**GitHub Pages (if you already use GitHub):**
1. Create a repo, upload the folder contents to it
2. Repo Settings → Pages → deploy from the main branch
3. Your app is live at `https://yourname.github.io/reponame/`

Either way, once it's live, open it on your phone — your browser will offer **"Add to Home Screen"** or **"Install app"**. That's it; it now behaves like a native app icon, and still works offline (it caches itself; the transactions themselves sync live when you have signal).

Go back to the Apps Script's `APP_URL` property and paste your hosted URL in, so future invite links work automatically.

## Part 3 — First login (you, the Owner)

Open the hosted app. On the login screen enter:
- **Sheet URL**: the `/exec` URL from Part 1
- **Name / Email**: whatever you like
- **Access code**: your `ADMIN_CODE`

You're now registered as Owner automatically.

## Part 4 — Invite other people

Go to **People** (only Owners see this tab) → **+ Invite someone** → enter their name, email, and role:
- **Viewer** — sees everything, can't add or change anything
- **Editor** — can add/edit/delete transactions, budgets, goals, investments, preferences
- (You can promote someone to **Owner** later from the same screen if needed)

They'll automatically get an email with their own access code — and, if you filled in `APP_URL`, a direct link that logs them straight in. Everyone sees the same shared transactions; roles just control who can change them, and that's enforced on the server side (Apps Script), not just hidden buttons.

## What's new in this version

- **Accounts** — unlimited wallets/accounts, each with its own type and currency (not just Personal/Business anymore)
- **Categories** — default income/expense categories, plus your own custom ones
- **Loans** — track money borrowed or lent, log repayments, see outstanding balance
- **Recurring transactions** — mark a transaction to repeat weekly/monthly/yearly; it auto-generates the next one when you open the app on or after its due date
- **Goals & Planning** — savings goals and one-off planning budgets (house construction, wedding, etc.) in one place
- **Multi-currency** — Settings -> base currency; fetch live rates (no API key needed) or enter your own; the Dashboard converts everything to your base currency automatically
- **Insights** — simple rule-based alerts (spending trend %, budget nearing/over, loan due dates, goal deadlines). This is arithmetic on your own numbers, not an AI model — flagged honestly so you know what it is
- **Reports** — income statement / cash flow by date range, with a Print -> Save as PDF button, plus CSV export on transactions, accounts, and loans
- **Dark mode** and an optional **PIN lock** (Settings) — the PIN is a convenience lock on this device only, not encryption
- **Activity log** — Owners see a running feed of who did what (People tab)

## Honest limits of this architecture

A few things from a full commercial finance-app spec genuinely do not fit a Google-Sheets-backed static app, so they are intentionally left out rather than faked:
- No native Android/iOS app-store build — the installed PWA is the realistic equivalent
- No fingerprint/Face ID — the PIN lock is the closest practical substitute
- No background push notifications when the app is fully closed — alerts show while the app is open
- Insights are rule-based, not a trained AI model
- No receipt photo storage — Sheets is not a file store

## Notes
- Everything is cached to the device first, so nothing is lost if you go offline — it syncs automatically once you're back online (and refreshes quietly every 45 seconds while the app is open).
- Settings → **Export/Import .xlsx** gives you a fully offline Excel backup that doesn't touch Google at all.
- If you ever need to revoke someone, remove them from the People tab — their code stops working immediately.
