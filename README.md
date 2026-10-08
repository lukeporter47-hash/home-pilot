# PorterPilot

Multi-household household-management app, built entirely on **Google Apps
Script** — no Firebase, no separate backend to host or pay for.

## How the multi-tenancy works

The web app is deployed with:

- **Execute as:** User accessing the web app
- **Who has access:** Anyone with a Google account

Because the script runs *as the visitor*, their own Google sign-in is the
login — there's no separate auth step, and `CalendarApp` / `DriveApp` calls
inside it already operate on their own Google account. Each person's
household Spreadsheet ID is stored in their own `UserProperties` (private
storage, scoped to this script + that Google account).

Every data function in the project — Stock, Chores, Bills, Baby log,
Calendar, Meals, Reminders, and the Bar/Pets/Money/Contacts/Documents/
Packing modules — reads and writes through one helper, `ss_()`
(`Backend.gs`). That's the single point where multi-tenancy lives: instead
of a Sheet this script used to be bound to, `ss_()` now opens *this
visitor's* household Spreadsheet. Nothing else had to change.

- **`doGet(e)`** (`Backend.gs`) is the router: no household yet (or an
  `?hh=` invite link) → serves `Gate.html` (create/join); already set up →
  serves `Index.html`, the full app.
- **Creating a household** (`Household.gs`) makes a new Google Sheet with
  the same tab layout the original single-household app used (`People`,
  `Stock`, `StockLog`, `BabyLog`, `Chores`, `Bills`, `Meals`, `Reminders`),
  plus a hidden `Meta` tab marking it as a real PorterPilot household.
- **Joining a household** happens through an invite link:
  `<web app url>?hh=<spreadsheetId>`. The admin shares the Sheet with the
  new member first (Sheet's own File ▸ Share), then sends them the link
  from the "You're all set" screen — the script checks they actually have
  access before trusting it.

## Files

| File | Purpose |
|---|---|
| `Household.gs` | Account layer: create/join/leave a household, `ss_()`'s resolver key |
| `Backend.gs` | Core app logic (home, baby, stock, chores, bills, calendar, meals, reminders) + the `doGet` router |
| `Bar.gs`, `Pets.gs`, `Money.gs`, `Contacts.gs`, `Documents.gs`, `Packing.gs` | Extra modules, unchanged from the original — automatically multi-tenant via `ss_()` |
| `Notify.gs` | OneSignal push notifications (morning/evening digest). Each household's admin runs `installNotificationTriggers()` themselves (e.g. from the Apps Script editor) — the trigger then fires under their own authorization, so it resolves their own household correctly |
| `Gate.html` | Sign-in-free create/join household screen (PorterPilot branding) |
| `Index.html` | The full app, unchanged from the original single-household build |
| `appsscript.json` | Manifest — sets the web app execution/access mode |
| `assets/logo.png` | PorterPilot logo mark |

## Deploying

You'll need [`clasp`](https://github.com/google/clasp), Google's Apps Script
CLI:

```bash
npm install -g @google/clasp
clasp login
```

**First time, from scratch:**

```bash
clasp create --type webapp --title "PorterPilot"
clasp push
clasp deploy
```

**On an existing Apps Script project**, copy `.clasp.json.example` to
`.clasp.json`, fill in your `scriptId` (Apps Script editor ▸ Project
Settings), then:

```bash
clasp push
clasp deploy
```

`.clasp.json` is gitignored — it's personal to whoever is deploying, so
it's never committed.

After deploying, open the generated web app URL — that's the link you
share with other households. The first person to open it for a given
household becomes that household's admin.

## Status

- ✅ Multi-tenant account layer (create/join/leave a household)
- ✅ Full feature set carried over from the original single-household app:
  home screen, baby log, stock/shopping, chores, bills, calendar, meals,
  reminders, bar stock, pets, money/budget, contacts, documents, packing
- ⚠️ Push notifications (`Notify.gs`) are self-service per household — each
  admin runs `installNotificationTriggers()` once for their own household
- ⚠️ `Setup.gs` (the old single-Sheet onboarding script) is superseded by
  `Household.gs` and isn't part of this project anymore
- ⚠️ The IFTTT voice-logging webhook (`Voice.gs` in the original project)
  needs its own separate Apps Script deployment — it defines its own
  `doGet`/`doPost`, which would collide with this project's router if
  merged into the same one
