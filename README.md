# CCCC — Connor's Chrome Cookie Cutter

A PowerShell tool that inventories every cookie in your browsers and cuts out the trackers and dead ones, without touching your logins.

It reads cookie **metadata only**. Cookie values are never read or exported.

---

## What it does

| Mode | Command | What happens |
|---|---|---|
| **Analyze** (default) | `.\CCCC.ps1` | Copies each cookie database to a scratch folder (safe while browsers are open) and reports the whole jar: totals per browser/profile, persistent vs session, expired, known trackers, expiry distribution, top sites |
| **Delete** | `.\CCCC.ps1 -Delete -Trackers -Expired` | Backs up each store, then deletes cookies by rule |
| **Snapshot** | `.\CCCC.ps1 -Snapshot` | Copies every cookie store to a timestamped folder and deletes nothing |
| **Restore** | `.\CCCC.ps1 -Restore <snapshot>` | Puts a snapshot back. Takes a safety snapshot of the current state first, so a restore can itself be undone |

### Delete rules

| Flag | Deletes |
|---|---|
| `-Trackers` | Cookies from a built-in list of ad/analytics domains, **plus** tracker cookies matched by name (`_ga`, `_fbp`, `_uetvid`, `MUID`…) on any site, which catches first-party trackers a domain list misses |
| `-Expired` | Cookies already past their expiry that are still sitting on disk |
| `-Domain a.com,b.com` | Only those sites (wildcards ok: `*ads*`) |
| `-KeepDomain a.com,b.com` | Everything **except** those sites (an allow-list) |
| `-Wipe` | The entire cookie store |

`-Trackers -Expired` together means "tracker **or** expired". That's the nightly rule.

## Supported browsers

Chrome · Edge · Brave · Opera · Opera GX · Vivaldi · Chromium · Firefox. All profiles, or narrow with `-ProfileName`.

## Requirements

- Windows PowerShell 5.1 or later. No admin rights needed.
- A SQLite engine for analysis and selective deletes. Either one works:
  - `winget install SQLite.SQLite`
  - Python 3 (uses its built-in `sqlite3` module)

  Without one, only `-Wipe` works.

## Quick start

```powershell
# Inventory every browser and profile
.\CCCC.ps1

# Export every cookie row (no values) to CSV
.\CCCC.ps1 -ExportCsv .\cookies.csv

# Preview the nightly rule without changing anything
.\CCCC.ps1 -Delete -Trackers -Expired -WhatIf

# Run it for real, closing browsers that hold the files
.\CCCC.ps1 -Delete -Trackers -Expired -CloseBrowsers

# Keep two sites' logins, purge everything else in Chrome and Edge
.\CCCC.ps1 -Delete -Browser Chrome,Edge -KeepDomain github.com,google.com -CloseBrowsers
```

Full parameter help:

```powershell
Get-Help .\CCCC.ps1 -Full
```

## Run it nightly

Point a Windows Task Scheduler task at:

```
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\path\to\CCCC.ps1" -Delete -Trackers -Expired -CloseBrowsers -Force
```

`-CloseBrowsers` force-closes open browsers, so schedule it for a time you're not using them.

## What the first inventory found

From one machine, two browsers:

- **About 1 in 5** cookies were already expired but still on disk
- **About 1 in 8** came from known trackers
- **Ad exchanges** were the biggest tracker group by far
- The longest-lived cookies stopped at **about 13 months**, which is Chrome's 400-day cap

Next: logging nightly to measure how fast trackers come back after each clean.

## Safety

- **Backups first.** Every delete backs up the store unless you pass `-NoBackup`.
- **Preview anything** with `-WhatIf`.
- **Undo** with `-Restore`.
- **Cleanup.** Scratch copies of cookie databases are removed when the script exits.

> ⚠️ **Your CSV exports are your browsing history.** Keep them out of git. This repo's `.gitignore` blocks `*.csv` and `*.log`.

## Extending the tracker list

Add domains to `$TrackerDomains`, or cookie names to `$TrackerCookieNames`, near the top of the script. Names ending in `*` are treated as a prefix match.

---

Built by [WARD3V](https://github.com/WARD3V) · part of a public log of security and privacy builds.
