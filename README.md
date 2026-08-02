# The Ledger — Setup

## ⚠️ If you already deployed the earlier version
This update changes the data model from a flat "Accounts" list to a
**Workspace → Wallet** hierarchy (Personal/Business/Family/… each with
their own wallets, budgets, goals, loans, and permissions). It is a
breaking schema change:

- Your existing `Users` tab and login still work fine — no need to
  re-bootstrap.
- Your old `Accounts` tab is left untouched but ignored (harmless).
- Old `Transactions`/`Budgets`/`Goals`/`Loans`/`Preferences`/`Investments`
  rows have no `workspaceId`, so they won't show up until you either
  re-enter them or manually add a `workspaceId` column value in the
  Sheet matching a new workspace's ID.
- Since you're still early in setup, the simplest path is usually a
  **fresh Google Sheet** (repeat step 1) rather than migrating old rows
  by hand. If you'd rather keep what's there, redeploy this `Code.gs`
  on the same Sheet, log in, create your first Workspace from the new
  **Workspaces** screen, then re-add wallets and re-enter (or hand-copy)
  any transactions you need to keep.

## 1. Backend (Google Sheets + Apps Script)
1. Create a new Google Sheet.
2. Extensions → Apps Script. Delete the default code, paste in `Code.gs`.
3. Project Settings (gear icon) → Script Properties → add a property
   `ADMIN_CODE` with a secret value only you know.
4. Deploy → New deployment → type **Web app**.
   - Execute as: **Me**
   - Who has access: **Anyone**
5. Copy the `.../exec` URL it gives you — this is your Server URL.
6. First time you run the script (e.g. via the deploy step), Google will
   ask you to authorize it — accept, so it can create sheet tabs, send
   mail, and use Drive for receipts.

## 2. Frontend
`index.html`, `sw.js`, `manifest.json`, `icon-192.png`, `icon-512.png`
must all sit in the **same folder** when hosted (GitHub Pages, Netlify,
Cloudflare Pages, or any static file host — all free, no build step).

## 3. First login
Open the hosted site → "First-time setup" tab → enter the Server URL,
your name, your email, and the `ADMIN_CODE`. This makes you the Owner
and creates all the sheet tabs automatically, including a starter
"Personal" workspace. Add more workspaces (Business, Family, Church,
etc.) any time from the **Workspaces** screen. From then on, invite
others from **People** — you'll choose which workspaces each person can
see, and they log in with the Server URL + their email + the access
code you send them.

## 4. Install as an app
On phone: open the site in Chrome/Safari → "Add to Home Screen".
On desktop: Chrome/Edge will show an install icon in the address bar.

## Notes
- Everything is local-first: changes save to the device immediately and
  sync to the Sheet in the background (~1.2s debounce), with a full
  pull every 45 seconds.
- Roles (Owner/Editor/Viewer) are enforced server-side in Apps Script,
  not just hidden in the UI.
- Insights are plain rule-based checks on your own data — no AI/ML.
- PIN lock is device-local only (hashed, stored in localStorage) and
  can always be turned off from the lock screen if forgotten.
