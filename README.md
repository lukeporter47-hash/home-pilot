# PorterPilot — Apps Script edition

Multi-household version of the household app, rebuilt on **Google Apps Script**
instead of Firebase. No external auth provider, no Firestore — just a Google
Sheet per household and a script that runs as whoever is visiting.

## How it works

The web app is deployed with:

- **Execute as:** User accessing the web app
- **Who has access:** Anyone with a Google account

Because the script runs *as the visitor*, their own Google sign-in is the
login — there's no separate auth step. Each person's household Spreadsheet ID
is remembered in their own `UserProperties` (private storage, scoped to this
script + that Google account), so it's there again next time they open the
app.

- **Creating a household** makes a new Google Sheet (`People`, `Stock`,
  `Chores`, `Bills`, `Meals`, `Reminders`, etc. — same tabs as the original
  single-household app) and remembers it as "your" household.
- **Joining a household** happens through an invite link:
  `<web app url>?hh=<spreadsheetId>`. The admin shares the Sheet with the new
  member first (File ▸ Share, in the Sheet itself), then sends them the link
  from the "You're all set" screen. The script checks they actually have
  access to that Sheet before trusting it.

This keeps everything inside Luke's/each family's own Google account — no
shared backend to pay for or run out of quota on, since Apps Script quotas
are per the *executing* user, not per the script owner.

## Files

| File | Purpose |
|---|---|
| `Code.gs` | Server-side logic: `doGet`, household create/join/leave |
| `Index.html` | The sign-in-less front end (uses `google.script.run`) |
| `appsscript.json` | Manifest — sets the web app execution/access mode |
| `assets/logo.png` | PorterPilot logo mark, for reference / app icons |

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
`.clasp.json`, fill in your `scriptId` (Apps Script editor ▸ Project Settings),
then:

```bash
clasp push
clasp deploy
```

`.clasp.json` is gitignored — it's personal to whoever is deploying, so it's
never committed.

After deploying, open the generated web app URL — that's the link you share
with your household (or publish later as the PWA/Play Store entry point).

## Status

This covers sign-in-free household creation and joining — the same ground
the Firebase proof of concept covered. The rest of the feature set (Baby log,
Stock/shopping, Calendar, Bills, Chores) from the original single-household
app still needs to be layered on top of this multi-tenant foundation.
