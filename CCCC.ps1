<#
.SYNOPSIS
    CCCC (Connor's Chrome Cookie Cutter) — inventory and purge browser cookies across every
    installed browser and profile.
.DESCRIPTION
    Two phases:

      ANALYZE (default)
        Copies each browser's cookie database to a scratch folder (safe
        while browsers are open) and reports the whole "cookie jar":
          - per browser/profile totals, sites, persistent vs session,
            expired count, known ad/tracking domains, DB size
          - expiry distribution (how long cookies will outlive you)
          - top sites across the jar and which browsers hold them
          - tracker domains present
        -ExportCsv writes every cookie row (no values) for spreadsheet work.

      DELETE (-Delete)
        Requires browsers closed (or -CloseBrowsers to kill them), backs up
        each DB, then removes cookies by rule:
          -Domain      only cookies for these sites (wildcards ok)
          -KeepDomain  everything EXCEPT these sites (allow-list)
          -Trackers    known ad/tracking domains, plus tracker cookie NAMES
                       (_ga, _fbp, _uetvid, MUID...) on any site
          -Expired     only cookies already past their expiry
          -Wipe        delete the whole cookie store file(s)
          -IncludeLegacy also clears the IE / WebView cookie cache
        -Trackers -Expired together means "tracker OR expired".

      SNAPSHOT (-Snapshot)
        Copies every selected cookie store (browsers may stay open) to
        <BackupDir>\<timestamp>\<Browser>_<Profile>\ plus a manifest.json,
        and deletes nothing. Analyze it later with -From, put it back with
        -Restore. Delete runs write the same layout automatically.

      RESTORE (-Restore <dir>)
        Copies a snapshot back over the live stores. Browsers must be
        closed (or -CloseBrowsers). The current live stores are snapshotted
        first unless -NoBackup, so a restore is itself reversible.

    Browsers: Chrome, Edge, Brave, Opera, Opera GX, Vivaldi, Chromium
    (all profiles), Firefox (all profiles). -ProfileName narrows to
    specific profiles.

    Selective work needs a SQLite engine. The script looks for
    sqlite3.exe (PATH or next to the script), then Python 3 (built-in
    sqlite3 module). Without either, only -Wipe works.

.PARAMETER Browser
    One or more of All, Chrome, Edge, Brave, Opera, OperaGX, Vivaldi,
    Chromium, Firefox. Default: All.
.PARAMETER ProfileName
    Only touch profiles matching these. Matches the profile folder name
    ("Default", "Profile 1") or the display name shown in Chrome ("Work").
    Wildcards ok. Omit to include every profile.
.PARAMETER Top
    How many sites to show in the top-sites table (Analyze). Default 25.
.PARAMETER ExportCsv
    Path for a CSV of every cookie (Analyze).
.PARAMETER PassThru
    Analyze: return the cookie objects. Delete: return the per-store
    result rows (used by the nightly wrapper to judge success).
.PARAMETER From
    Analyze a saved snapshot folder instead of the live browsers. Accepts a
    full path or just the timestamp folder name under BackupDir.
.PARAMETER Snapshot
    Copy every selected store to BackupDir\<timestamp>\ and stop.
.PARAMETER Restore
    Snapshot folder (full path or timestamp name under BackupDir) to copy
    back over the live stores.
.PARAMETER Domain
    Delete cookies whose host matches any of these. "example.com" matches
    example.com, .example.com and *.example.com. "*ads*" style wildcards ok.
.PARAMETER KeepDomain
    Delete everything except cookies matching these domains.
.PARAMETER Trackers
    Delete cookies on the built-in $TrackerDomains list. Combined with
    -Expired, a cookie is deleted if it is a tracker OR already expired.
.PARAMETER Expired
    Delete only cookies whose expiry is already in the past.
.PARAMETER Wipe
    Delete the entire cookie database file(s) for each targeted store.
.PARAMETER CloseBrowsers
    Kill running browser processes that hold the target stores.
.PARAMETER NoBackup
    Skip the pre-delete backup copy.
.PARAMETER BackupDir
    Where backups go. Default D:\SnapshotData\<timestamp>\<Browser>_<Profile>\.
.PARAMETER IncludeLegacy
    Also clear the legacy IE / WebView cookie cache (INetCookies).
.PARAMETER Force
    Suppress confirmation prompts (same as -Confirm:$false).

.EXAMPLE
    .\CCCC.ps1
    Inventory every browser and profile.
.EXAMPLE
    .\CCCC.ps1 -Browser Chrome
    Inventory Chrome only; the Profile column shows folder + display name.
.EXAMPLE
    .\CCCC.ps1 -Top 50 -ExportCsv D:\Lab\Logs\cookies.csv
.EXAMPLE
    .\CCCC.ps1 -Delete -Domain doubleclick.net,facebook.com -WhatIf
    Preview a targeted delete without touching anything.
.EXAMPLE
    .\CCCC.ps1 -Delete -Browser Chrome -ProfileName Default,'Profile 1' -Trackers -Expired -WhatIf
    Preview the nightly rule on two Chrome profiles.
.EXAMPLE
    .\CCCC.ps1 -Delete -Browser Chrome,Edge -KeepDomain github.com,google.com -CloseBrowsers
    Keep logins for two sites, purge everything else in Chrome and Edge.
.EXAMPLE
    .\CCCC.ps1 -Delete -Expired -Force
.EXAMPLE
    .\CCCC.ps1 -Delete -Wipe -CloseBrowsers -IncludeLegacy
    Scorched earth.
.EXAMPLE
    .\CCCC.ps1 -Snapshot
    Copy every store to D:\SnapshotData\<timestamp>\ without deleting.
.EXAMPLE
    .\CCCC.ps1 -From 20260101_120000
    Analyze that snapshot instead of the live browsers.
.EXAMPLE
    .\CCCC.ps1 -Restore 20260101_120000 -Browser Edge -CloseBrowsers
    Put Edge's cookies back the way they were in that snapshot.

.NOTES
    Windows PowerShell 5.1 compatible. Run as the user whose profile you
    are cleaning (no elevation needed).
    Cookie VALUES are never read or exported — only metadata.
    Session-restore ("continue where you left off") re-creates cookies
    when tabs reload; Firefox session restore also revives session
    cookies from sessionstore. Close tabs before a wipe if that matters.
    Oct 2026: added -ProfileName, -Trackers, and -PassThru for Delete
    (nightly scheduled-task support).
#>

[CmdletBinding(SupportsShouldProcess = $true,
               ConfirmImpact = 'High',
               DefaultParameterSetName = 'Analyze')]
param(
    [ValidateSet('All','Chrome','Edge','Brave','Opera','OperaGX',
                 'Vivaldi','Chromium','Firefox')]
    [string[]]$Browser = @('All'),

    # Not named $Profile: that is a PowerShell automatic variable.
    [string[]]$ProfileName,

    # ── Analyze ──
    [Parameter(ParameterSetName = 'Analyze')]
    [switch]$Analyze,

    [Parameter(ParameterSetName = 'Analyze')]
    [int]$Top = 25,

    [Parameter(ParameterSetName = 'Analyze')]
    [string]$ExportCsv,

    [Parameter(ParameterSetName = 'Analyze')]
    [Parameter(ParameterSetName = 'Delete')]
    [switch]$PassThru,

    [Parameter(ParameterSetName = 'Analyze')]
    [string]$From,

    # ── Snapshot / Restore ──
    [Parameter(ParameterSetName = 'Snapshot', Mandatory = $true)]
    [switch]$Snapshot,

    [Parameter(ParameterSetName = 'Restore', Mandatory = $true)]
    [string]$Restore,

    # ── Delete ──
    [Parameter(ParameterSetName = 'Delete', Mandatory = $true)]
    [switch]$Delete,

    [Parameter(ParameterSetName = 'Delete')]
    [string[]]$Domain,

    [Parameter(ParameterSetName = 'Delete')]
    [string[]]$KeepDomain,

    [Parameter(ParameterSetName = 'Delete')]
    [switch]$Trackers,

    [Parameter(ParameterSetName = 'Delete')]
    [switch]$Expired,

    [Parameter(ParameterSetName = 'Delete')]
    [switch]$Wipe,

    [Parameter(ParameterSetName = 'Delete')]
    [Parameter(ParameterSetName = 'Restore')]
    [switch]$CloseBrowsers,

    [Parameter(ParameterSetName = 'Delete')]
    [Parameter(ParameterSetName = 'Restore')]
    [switch]$NoBackup,

    [Parameter(ParameterSetName = 'Analyze')]
    [Parameter(ParameterSetName = 'Delete')]
    [Parameter(ParameterSetName = 'Snapshot')]
    [Parameter(ParameterSetName = 'Restore')]
    [string]$BackupDir = 'D:\SnapshotData',

    [Parameter(ParameterSetName = 'Delete')]
    [switch]$IncludeLegacy,

    [Parameter(ParameterSetName = 'Delete')]
    [Parameter(ParameterSetName = 'Restore')]
    [switch]$Force
)

Set-StrictMode -Version 2
if ($Force) { $ConfirmPreference = 'None' }

# ── CONFIGURATION ──

# Every browser we know how to find. Chromium-family stores live at
# <profile>\Network\Cookies (v96+) or <profile>\Cookies (older).
# Firefox stores live at <profile>\cookies.sqlite.
$Catalog = @(
    @{ Name = 'Chrome';   Engine = 'Chromium'; Process = 'chrome';  PathHint = '\Google\Chrome\'
       Root = (Join-Path $env:LOCALAPPDATA 'Google\Chrome\User Data') }
    @{ Name = 'Edge';     Engine = 'Chromium'; Process = 'msedge';  PathHint = '\Microsoft\Edge\'
       Root = (Join-Path $env:LOCALAPPDATA 'Microsoft\Edge\User Data') }
    @{ Name = 'Brave';    Engine = 'Chromium'; Process = 'brave';   PathHint = '\BraveSoftware\'
       Root = (Join-Path $env:LOCALAPPDATA 'BraveSoftware\Brave-Browser\User Data') }
    @{ Name = 'Vivaldi';  Engine = 'Chromium'; Process = 'vivaldi'; PathHint = '\Vivaldi\'
       Root = (Join-Path $env:LOCALAPPDATA 'Vivaldi\User Data') }
    @{ Name = 'Chromium'; Engine = 'Chromium'; Process = 'chrome';  PathHint = '\Chromium\'
       Root = (Join-Path $env:LOCALAPPDATA 'Chromium\User Data') }
    @{ Name = 'Opera';    Engine = 'Chromium'; Process = 'opera';   PathHint = '\Opera\'
       Root = (Join-Path $env:APPDATA 'Opera Software\Opera Stable'); FlatProfile = $true }
    @{ Name = 'OperaGX';  Engine = 'Chromium'; Process = 'opera';   PathHint = '\Opera GX\'
       Root = (Join-Path $env:APPDATA 'Opera Software\Opera GX Stable'); FlatProfile = $true }
    @{ Name = 'Firefox';  Engine = 'Gecko';    Process = 'firefox'; PathHint = '\Mozilla Firefox\'
       Root = (Join-Path $env:APPDATA 'Mozilla\Firefox\Profiles') }
)

# Registrable-domain heuristic: treat "x.co.uk" style as three labels.
$SecondLevelTlds = @('co','com','net','org','gov','edu','ac','ne','or','go','mil')

# Known ad / analytics / tracking registrable domains. Extend freely.
$TrackerDomains = @(
    'doubleclick.net','googlesyndication.com','googleadservices.com'
    'google-analytics.com','googletagmanager.com','googletagservices.com'
    'adnxs.com','adsrvr.org','criteo.com','criteo.net','taboola.com'
    'outbrain.com','scorecardresearch.com','quantserve.com','rubiconproject.com'
    'pubmatic.com','openx.net','casalemedia.com','bluekai.com','demdex.net'
    'everesttech.net','omtrdc.net','amazon-adsystem.com','hotjar.com'
    'mixpanel.com','segment.io','segment.com','fullstory.com','mouseflow.com'
    'crazyegg.com','facebook.net','fbcdn.net','doubleverify.com','moatads.com'
    'adform.net','yieldmo.com','sharethrough.com','indexww.com','liadm.com'
    'id5-sync.com','rlcdn.com','agkn.com','krxd.net','turn.com','3lift.com'
    'bidswitch.net','smartadserver.com','teads.tv','media.net','tapad.com'
    'zemanta.com','mediamath.com','adroll.com','bounceexchange.com'
    'onetrust.com','cookielaw.org','trustarc.com','permutive.com','chartbeat.com'
    'newrelic.com','nr-data.net','optimizely.com','clarity.ms'
    'ads-twitter.com','t.co'
    'a-mo.net','contextweb.com','stackadapt.com','aniview.com'   # added after first real inventory
    # Microsoft ad / telemetry hosts (subdomains only, so Bing, MSN and account sign-ins survive):
    # bat.bing.com = UET ad tag, c.bing.com / c.msn.com = MUID cross-site ID sync, msads.net = MS ads,
    # ads.linkedin.com = LinkedIn (Microsoft) ad tag. Xandr (adnxs.com) and Clarity (clarity.ms) are above.
    'bat.bing.com','c.bing.com','c.msn.com','msads.net','ads.linkedin.com'
)

# Tracker cookies identified by NAME. Many trackers drop their cookie on the site you
# are visiting (first-party), so a domain list never sees them. '*' = prefix match.
# MS: _uetsid/_uetvid (UET ads), _clck/_clsk/ANONCHK (Clarity), MUID (Microsoft ad ID).
# Others: Google Analytics/Ads, Meta, TikTok, Pinterest, Snap, Reddit.
$TrackerCookieNames = @(
    '_uetsid','_uetvid','_clck','_clsk','ANONCHK','MUID'
    '_ga','_ga_*','_gid','_gat*','_gcl_au','_gcl_aw','_gcl_dc'
    '_fbp','_fbc','_ttp','_tt_enable_cookie','_pin_unauth','_scid','_rdt_uuid'
)

function Test-TrackerHost([string]$HostName) {
    $h = $HostName.TrimStart('.').ToLowerInvariant()
    foreach ($d in $TrackerDomains) { if ($h -eq $d -or $h.EndsWith('.' + $d)) { return $true } }
    return $false
}

function Test-TrackerName([string]$CookieName) {
    foreach ($n in $TrackerCookieNames) { if ($CookieName -clike $n) { return $true } }
    return $false
}

$LegacyCookieDir = Join-Path $env:LOCALAPPDATA 'Microsoft\Windows\INetCookies'

$Script:Scratch  = Join-Path $env:TEMP ('CookieJar\' + (Get-Date -Format 'yyyyMMdd_HHmmss'))
$Script:PyHelper = $null
$Script:Sql      = $null

# ── HELPERS ──

function Write-Section([string]$Text) {
    Write-Host ''
    Write-Host ("── {0} ──" -f $Text) -ForegroundColor Cyan
}

function Convert-Unix([long]$Seconds) {
    if ($Seconds -le 0) { return $null }
    if ($Seconds -gt 253402300799) { $Seconds = 253402300799 }   # year 9999 cap
    try { return [DateTimeOffset]::FromUnixTimeSeconds($Seconds).LocalDateTime }
    catch { return $null }
}

function Get-RegistrableDomain([string]$HostName) {
    $h = $HostName.TrimStart('.').ToLowerInvariant()
    if ($h -match '^\d{1,3}(\.\d{1,3}){3}$' -or $h -notmatch '\.') { return $h }
    $parts = $h.Split('.')
    if ($parts.Count -ge 3 -and $parts[-1].Length -eq 2 -and
        ($SecondLevelTlds -contains $parts[-2])) {
        return ($parts[-3..-1] -join '.')
    }
    return ($parts[-2..-1] -join '.')
}

function Test-ProfileWanted([string]$Folder, [string]$Label) {
    # True when -ProfileName is unset, or a pattern matches the folder name,
    # the full label "Profile 1 (Work)", or the display name inside the parens.
    if (-not $ProfileName) { return $true }
    $display = $null
    if ($Label -match '\((?<n>[^()]+)\)$') { $display = $matches['n'] }
    foreach ($w in $ProfileName) {
        if ($Folder -like $w -or $Label -like $w) { return $true }
        if ($display -and $display -like $w) { return $true }
    }
    return $false
}

function Get-ChromiumProfileNames([string]$Root) {
    # Map profile folder ("Profile 1") -> display name from Local State.
    $map = @{}
    $ls = Join-Path $Root 'Local State'
    if (-not (Test-Path -LiteralPath $ls)) { return $map }
    try {
        $json  = Get-Content -LiteralPath $ls -Raw -Encoding UTF8 | ConvertFrom-Json
        $cache = $json.profile.info_cache
        foreach ($prop in $cache.PSObject.Properties) {
            $n = $prop.Value.name
            if ($n) { $map[$prop.Name] = [string]$n }
        }
    } catch { }
    return $map
}

function Get-CookieStores([string[]]$Want) {
    $stores = @()
    foreach ($b in $Catalog) {
        if (-not ($Want -contains 'All' -or $Want -contains $b.Name)) { continue }
        if (-not (Test-Path -LiteralPath $b.Root)) { continue }

        $flat = $b.ContainsKey('FlatProfile') -and $b.FlatProfile
        if ($flat) { $profileDirs = @(Get-Item -LiteralPath $b.Root) }
        else       { $profileDirs = @(Get-ChildItem -LiteralPath $b.Root -Directory -ErrorAction SilentlyContinue) }

        $names = @{}
        if ($b.Engine -eq 'Chromium' -and -not $flat) { $names = Get-ChromiumProfileNames $b.Root }

        foreach ($p in $profileDirs) {
            $dbs = @()
            if ($b.Engine -eq 'Chromium') {
                foreach ($rel in @('Network\Cookies','Cookies')) {
                    $c = Join-Path $p.FullName $rel
                    if (Test-Path -LiteralPath $c -PathType Leaf) { $dbs += $c }
                }
            } else {
                $c = Join-Path $p.FullName 'cookies.sqlite'
                if (Test-Path -LiteralPath $c -PathType Leaf) { $dbs += $c }
            }
            foreach ($db in $dbs) {
                $label = $p.Name
                if ($flat) { $label = 'Default' }
                elseif ($names.ContainsKey($p.Name)) { $label = "{0} ({1})" -f $p.Name, $names[$p.Name] }
                if (-not (Test-ProfileWanted $p.Name $label)) { continue }
                $stores += [pscustomobject]@{
                    Browser    = $b.Name
                    Engine     = $b.Engine
                    Profile    = $label
                    ProfileDir = $p.FullName
                    Db         = $db
                    Process    = $b.Process
                    PathHint   = $b.PathHint
                    SizeKB     = [math]::Round((Get-Item -LiteralPath $db).Length / 1KB, 1)
                }
            }
        }
    }
    return ,$stores
}

function Get-StoreSidecars([string]$Db) {
    # SQLite side files that must travel with the main DB.
    $out = @()
    foreach ($suffix in @('-journal','-wal','-shm')) {
        $s = $Db + $suffix
        if (Test-Path -LiteralPath $s -PathType Leaf) { $out += $s }
    }
    return ,$out
}

function Copy-FileShared([string]$From, [string]$To) {
    # Open with ReadWrite|Delete sharing so a running browser doesn't block us.
    $share = [IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete
    $in = New-Object IO.FileStream($From, [IO.FileMode]::Open, [IO.FileAccess]::Read, $share)
    try {
        $out = New-Object IO.FileStream($To, [IO.FileMode]::Create, [IO.FileAccess]::Write, [IO.FileShare]::None)
        try { $in.CopyTo($out) } finally { $out.Dispose() }
    } finally { $in.Dispose() }
}

function Copy-StoreForRead($Store) {
    # Returns path of a scratch copy (with sidecars) of the store DB.
    $safeProfile = ($Store.Profile -replace '[^\w\-]', '_')
    $dest = Join-Path $Script:Scratch ("{0}_{1}_{2}" -f $Store.Browser, $safeProfile, [IO.Path]::GetFileName($Store.Db))
    Copy-FileShared $Store.Db $dest
    foreach ($s in (Get-StoreSidecars $Store.Db)) {
        Copy-FileShared $s ($dest + $s.Substring($Store.Db.Length))
    }
    return $dest
}

function Get-BrowserProcesses($Store) {
    $procs = @(Get-Process -Name $Store.Process -ErrorAction SilentlyContinue)
    if ($Store.Process -eq 'chrome' -or $Store.Process -eq 'opera') {
        # chrome.exe is shared by Chrome and Chromium; opera.exe by Opera and GX.
        $procs = @($procs | Where-Object {
            $path = $null
            try { $path = $_.Path } catch { }
            (-not $path) -or ($path -like ('*' + $Store.PathHint + '*'))
        })
    }
    # Plain return (no unary comma): callers wrap in @() so an empty result counts 0.
    return $procs
}

# ── SNAPSHOTS ──
# Layout: <BackupDir>\<timestamp>\<Browser>_<Profile>\<Cookies | cookies.sqlite> (+ sidecars)
#         <BackupDir>\<timestamp>\manifest.json  (original paths, so -Restore knows where things go)

function ConvertTo-SafeName([string]$Text) { return ($Text -replace '[^\w\-]', '_') }

function Resolve-SnapshotDir([string]$Dir) {
    # Accept a full path or just a timestamp folder name under BackupDir.
    if (Test-Path -LiteralPath $Dir -PathType Container) { return (Get-Item -LiteralPath $Dir).FullName }
    $under = Join-Path $BackupDir $Dir
    if (Test-Path -LiteralPath $under -PathType Container) { return (Get-Item -LiteralPath $under).FullName }
    throw "Snapshot folder not found: $Dir  (also looked under $BackupDir)"
}

function Write-SnapshotManifest([string]$StampDir, $Entry) {
    $path = Join-Path $StampDir 'manifest.json'
    $entries = @()
    if (Test-Path -LiteralPath $path) {
        try { $entries = @((Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json)) } catch { $entries = @() }
    }
    $entries = @($entries | Where-Object { $_.Folder -ne $Entry.Folder }) + @($Entry)
    $json = ConvertTo-Json -InputObject @($entries) -Depth 4
    [IO.File]::WriteAllText($path, $json, (New-Object Text.UTF8Encoding($false)))
}

function Get-SnapshotStores([string]$Dir, [string[]]$Want) {
    # Rebuild store objects from a snapshot folder. Uses manifest.json when present,
    # otherwise infers Browser/Profile from the "<Browser>_<Profile>" folder name.
    $stores = @()
    $manifest = @{}
    $mpath = Join-Path $Dir 'manifest.json'
    if (Test-Path -LiteralPath $mpath) {
        try {
            foreach ($e in @((Get-Content -LiteralPath $mpath -Raw -Encoding UTF8 | ConvertFrom-Json))) { $manifest[[string]$e.Folder] = $e }
        } catch { }
    }
    foreach ($sub in (Get-ChildItem -LiteralPath $Dir -Directory)) {
        $db = $null; $engine = $null
        foreach ($cand in @(@('Cookies','Chromium'), @('cookies.sqlite','Gecko'))) {
            $p = Join-Path $sub.FullName $cand[0]
            if (Test-Path -LiteralPath $p -PathType Leaf) { $db = $p; $engine = $cand[1]; break }
        }
        if (-not $db) { continue }
        $browserName = $null; $profile = $null; $original = $null
        if ($manifest.ContainsKey($sub.Name)) {
            $m = $manifest[$sub.Name]
            $browserName = [string]$m.Browser; $profile = [string]$m.Profile; $original = [string]$m.OriginalDb
        } elseif ($sub.Name -match '^(?<b>[A-Za-z]+)_(?<p>.+)$') {
            $browserName = $matches['b']; $profile = $matches['p']
        } else { continue }
        $cat = $Catalog | Where-Object { $_.Name -eq $browserName } | Select-Object -First 1
        if (-not $cat) { continue }
        if (-not ($Want -contains 'All' -or $Want -contains $cat.Name)) { continue }
        $folderPart = $profile
        if ($profile -match '^(?<f>.+?) \(') { $folderPart = $matches['f'] }
        if (-not (Test-ProfileWanted $folderPart $profile)) { continue }
        $stores += [pscustomobject]@{
            Browser = $cat.Name; Engine = $engine; Profile = $profile; ProfileDir = $null
            Db = $db; Process = $cat.Process; PathHint = $cat.PathHint
            SizeKB = [math]::Round((Get-Item -LiteralPath $db).Length / 1KB, 1)
            OriginalDb = $original; Folder = $sub.Name; Snapshot = $true
        }
    }
    return ,$stores
}

# ── SQLITE ENGINE ──

function Find-SqlEngine {
    $exe = Get-Command 'sqlite3.exe' -ErrorAction SilentlyContinue
    if (-not $exe -and $PSScriptRoot) {
        $local = Join-Path $PSScriptRoot 'sqlite3.exe'
        if (Test-Path -LiteralPath $local) { $exe = Get-Command $local }
    }
    if ($exe) { return @{ Kind = 'sqlite3'; Path = $exe.Source; Args = @() } }

    foreach ($cand in @(@('py.exe', @('-3')), @('python.exe', @()))) {
        $cmd = Get-Command $cand[0] -ErrorAction SilentlyContinue
        if (-not $cmd) { continue }
        $probe = $null
        try {
            $probe = & $cmd.Source @($cand[1]) -c 'import sqlite3;print(sqlite3.sqlite_version)' 2>$null
        } catch { }
        if ($LASTEXITCODE -eq 0 -and $probe) {
            return @{ Kind = 'python'; Path = $cmd.Source; Args = $cand[1]; Version = ([string]$probe).Trim() }
        }
    }
    return $null
}

function Initialize-PyHelper {
    $py = @'
import sqlite3, sys, json
db, mode, sqlfile = sys.argv[1], sys.argv[2], sys.argv[3]
with open(sqlfile, "r", encoding="utf-8") as f:
    sql = f.read()
con = sqlite3.connect(db)
try:
    cur = con.execute(sql)
    if mode == "query":
        cols = [d[0] for d in cur.description]
        rows = [dict(zip(cols, r)) for r in cur.fetchall()]
        sys.stdout.write(json.dumps(rows))
    else:
        n = cur.rowcount
        con.commit()
        con.execute("VACUUM")
        sys.stdout.write(str(n))
finally:
    con.close()
'@
    $Script:PyHelper = Join-Path $Script:Scratch 'cookiesql.py'
    [IO.File]::WriteAllText($Script:PyHelper, $py, (New-Object Text.UTF8Encoding($false)))
}

function ConvertFrom-JsonBig([string]$Text) {
    try { return ,(ConvertFrom-Json $Text) }
    catch {
        Add-Type -AssemblyName System.Web.Extensions
        $ser = New-Object System.Web.Script.Serialization.JavaScriptSerializer
        $ser.MaxJsonLength = [int]::MaxValue
        return ,$ser.DeserializeObject($Text)
    }
}

function Invoke-CookieSql {
    param(
        [string]$Db,
        [string]$Sql,
        [ValidateSet('Query','Exec')][string]$Mode = 'Query'
    )
    $sqlFile = Join-Path $Script:Scratch ('q_' + [guid]::NewGuid().ToString('N') + '.sql')
    $body = $Sql
    if ($Script:Sql.Kind -eq 'sqlite3' -and $Mode -eq 'Exec') {
        $body = $Sql.TrimEnd(';') + '; SELECT changes(); VACUUM;'
    }
    [IO.File]::WriteAllText($sqlFile, $body, (New-Object Text.UTF8Encoding($false)))
    try {
        $out = $null
        if ($Script:Sql.Kind -eq 'python') {
            $out = & $Script:Sql.Path @($Script:Sql.Args) $Script:PyHelper $Db $Mode.ToLower() $sqlFile 2>&1
        } else {
            $read = '.read "' + ($sqlFile -replace '\\', '/') + '"'
            if ($Mode -eq 'Query') { $out = & $Script:Sql.Path -json $Db $read 2>&1 }
            else                   { $out = & $Script:Sql.Path $Db $read 2>&1 }
        }
        if ($LASTEXITCODE -ne 0) { throw ("SQLite call failed: " + (($out | Out-String).Trim())) }
        $text = ($out | Out-String).Trim()
        if ($Mode -eq 'Exec') {
            if ($text -match '(\d+)\s*$') { return [int]$matches[1] } else { return 0 }
        }
        if (-not $text -or $text -eq '[]') { return @() }
        return ,(ConvertFrom-JsonBig $text)
    } finally {
        Remove-Item -LiteralPath $sqlFile -WhatIf:$false -ErrorAction SilentlyContinue
    }
}

# ── COOKIE NORMALISATION ──

function Get-StoreCookies($Store, [string]$DbPath) {
    if ($Store.Engine -eq 'Chromium') {
        # expires_utc is microseconds since 1601-01-01; convert to unix seconds.
        $sql = @"
SELECT host_key AS Host, name AS Name, path AS Path,
       CASE WHEN has_expires = 1 THEN CAST(expires_utc / 1000000 AS INTEGER) - 11644473600 ELSE 0 END AS ExpiresUnix,
       is_persistent AS Persistent, is_secure AS Secure, is_httponly AS HttpOnly, samesite AS SameSite,
       length(encrypted_value) + length(value) AS Bytes,
       CAST(last_access_utc / 1000000 AS INTEGER) - 11644473600 AS LastAccessUnix,
       CAST(creation_utc / 1000000 AS INTEGER) - 11644473600 AS CreatedUnix
FROM cookies
"@
    } else {
        # Firefox keeps only persistent cookies on disk; session ones live in sessionstore.
        $sql = @"
SELECT host AS Host, name AS Name, path AS Path, expiry AS ExpiresUnix,
       1 AS Persistent, isSecure AS Secure, isHttpOnly AS HttpOnly, sameSite AS SameSite,
       length(value) AS Bytes,
       CAST(lastAccessed / 1000000 AS INTEGER) AS LastAccessUnix,
       CAST(creationTime / 1000000 AS INTEGER) AS CreatedUnix
FROM moz_cookies
"@
    }
    $rows = Invoke-CookieSql -Db $DbPath -Sql $sql -Mode Query
    $now  = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
    foreach ($r in $rows) {
        $hostName   = [string]$r.Host
        $site       = Get-RegistrableDomain $hostName
        $exp        = [long]$r.ExpiresUnix
        $persistent = ([int]$r.Persistent -eq 1) -and ($exp -gt 0)
        $expiresAt  = $null
        if ($persistent) { $expiresAt = Convert-Unix $exp }
        [pscustomobject]@{
            Browser    = $Store.Browser
            Profile    = $Store.Profile
            Host       = $hostName
            Site       = $site
            Name       = [string]$r.Name
            Path       = [string]$r.Path
            Persistent = $persistent
            Expired    = ($persistent -and ($exp -lt $now))
            Expires    = $expiresAt
            LastAccess = Convert-Unix ([long]$r.LastAccessUnix)
            Created    = Convert-Unix ([long]$r.CreatedUnix)
            Secure     = ([int]$r.Secure -eq 1)
            HttpOnly   = ([int]$r.HttpOnly -eq 1)
            SameSite   = [int]$r.SameSite
            Bytes      = [int]$r.Bytes
            Tracker    = ((Test-TrackerHost $hostName) -or (Test-TrackerName ([string]$r.Name)))
        }
    }
}

# ── ANALYZE ──

function Invoke-Analyze($Stores) {
    Write-Section 'COOKIE JAR INVENTORY'
    if ($From) { Write-Host ("  source: snapshot {0}" -f $From) -ForegroundColor Yellow }
    if ($Stores.Count -eq 0) { Write-Warning 'No cookie stores found for the selected browsers.'; return }

    $all      = New-Object System.Collections.Generic.List[object]
    $summary  = @()
    $locked   = @()

    foreach ($s in $Stores) {
        $running = $false
        if (-not $From) { $running = @(Get-BrowserProcesses $s).Count -gt 0 }
        $rows = @()
        $status = 'ok'
        try {
            $copy = Copy-StoreForRead $s
            $rows = @(Get-StoreCookies $s $copy)
        } catch {
            $status = 'LOCKED'
            $why = $_.Exception.Message
            if ($why -match 'being used by another process') { $why = 'file in use (browser holds an exclusive lock)' }
            $locked += [pscustomobject]@{ Browser = $s.Browser; Profile = $s.Profile; Error = $why }
        }
        foreach ($r in $rows) { $all.Add($r) }

        $sites = @($rows | Select-Object -ExpandProperty Site -Unique).Count
        $summary += [pscustomobject]@{
            Browser    = $s.Browser
            Profile    = $s.Profile
            Running    = $running
            Cookies    = $rows.Count
            Sites      = $sites
            Persistent = @($rows | Where-Object Persistent).Count
            Session    = @($rows | Where-Object { -not $_.Persistent }).Count
            Expired    = @($rows | Where-Object Expired).Count
            Trackers   = @($rows | Where-Object Tracker).Count
            SizeKB     = $s.SizeKB
            Status     = $status
        }
    }

    $summary | Format-Table Browser, Profile, Running, Cookies, Sites, Persistent, Session, Expired, Trackers, SizeKB, Status -AutoSize | Out-Host

    $total = $all.Count
    Write-Host ("Total: {0} cookies across {1} store(s), {2} distinct sites, {3} distinct hosts." -f `
        $total, $Stores.Count,
        @($all | Select-Object -ExpandProperty Site -Unique).Count,
        @($all | Select-Object -ExpandProperty Host -Unique).Count)

    if ($locked.Count -gt 0) {
        Write-Warning ('Could not read {0} store(s) — close that browser and re-run:' -f $locked.Count)
        $locked | Format-Table -AutoSize | Out-Host
    }
    if ($total -eq 0) { return }

    # Expiry distribution
    Write-Section 'EXPIRY DISTRIBUTION'
    $now = Get-Date
    $buckets = [ordered]@{
        'Session (gone on close)' = 0
        'Already expired'         = 0
        '< 1 day'                 = 0
        '1 - 30 days'             = 0
        '30 days - 1 year'        = 0
        '1 - 5 years'             = 0
        '> 5 years'               = 0
    }
    foreach ($c in $all) {
        if (-not $c.Persistent -or -not $c.Expires) { $buckets['Session (gone on close)']++; continue }
        $days = ($c.Expires - $now).TotalDays
        if     ($days -lt 0)    { $buckets['Already expired']++ }
        elseif ($days -lt 1)    { $buckets['< 1 day']++ }
        elseif ($days -lt 30)   { $buckets['1 - 30 days']++ }
        elseif ($days -lt 365)  { $buckets['30 days - 1 year']++ }
        elseif ($days -lt 1826) { $buckets['1 - 5 years']++ }
        else                    { $buckets['> 5 years']++ }
    }
    $buckets.GetEnumerator() | ForEach-Object {
        [pscustomobject]@{ Lifetime = $_.Key; Cookies = $_.Value; Pct = [math]::Round(100 * $_.Value / $total, 1) }
    } | Format-Table -AutoSize | Out-Host

    # Flags
    Write-Section 'FLAGS'
    [pscustomobject]@{
        Secure    = @($all | Where-Object Secure).Count
        HttpOnly  = @($all | Where-Object HttpOnly).Count
        Tracker   = @($all | Where-Object Tracker).Count
        Expired   = @($all | Where-Object Expired).Count
        TotalKB   = [math]::Round((($all | Measure-Object Bytes -Sum).Sum) / 1KB, 1)
    } | Format-Table -AutoSize | Out-Host

    # Top sites
    Write-Section ("TOP {0} SITES" -f $Top)
    $all | Group-Object Site | Sort-Object Count -Descending | Select-Object -First $Top | ForEach-Object {
        $g = $_
        [pscustomobject]@{
            Site     = $g.Name
            Cookies  = $g.Count
            Hosts    = @($g.Group | Select-Object -ExpandProperty Host -Unique).Count
            Browsers = (($g.Group | Select-Object -ExpandProperty Browser -Unique | Sort-Object) -join ',')
            Trackers = @($g.Group | Where-Object Tracker).Count
            Longest  = ($g.Group | Where-Object Expires | Sort-Object Expires -Descending | Select-Object -First 1).Expires
        }
    } | Format-Table -AutoSize | Out-Host

    # Trackers
    $trk = @($all | Where-Object Tracker)
    if ($trk.Count -gt 0) {
        Write-Section ("KNOWN TRACKERS, BY DOMAIN OR COOKIE NAME ({0} cookies)" -f $trk.Count)
        $trk | Group-Object Site | Sort-Object Count -Descending | ForEach-Object {
            [pscustomobject]@{
                Site     = $_.Name
                Cookies  = $_.Count
                Browsers = (($_.Group | Select-Object -ExpandProperty Browser -Unique | Sort-Object) -join ',')
            }
        } | Format-Table -AutoSize | Out-Host
        Write-Host 'Purge them with:' -ForegroundColor DarkGray
        Write-Host '  .\CCCC.ps1 -Delete -Trackers' -ForegroundColor DarkGray
    }

    # Legacy
    if (Test-Path -LiteralPath $LegacyCookieDir) {
        $legacyCount = @(Get-ChildItem -LiteralPath $LegacyCookieDir -File -Recurse -Force -ErrorAction SilentlyContinue).Count
        if ($legacyCount -gt 0) {
            Write-Section 'LEGACY IE / WEBVIEW CACHE'
            Write-Host ("{0} file(s) in {1}  (clear with -Delete -IncludeLegacy)" -f $legacyCount, $LegacyCookieDir)
        }
    }

    if ($ExportCsv) {
        $dir = Split-Path -Parent $ExportCsv
        if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        $all | Export-Csv -LiteralPath $ExportCsv -NoTypeInformation -Encoding UTF8
        Write-Host ''
        Write-Host ("Exported {0} rows -> {1}" -f $total, $ExportCsv) -ForegroundColor Green
    }
    if ($PassThru) { return $all }
}

# ── DELETE ──

function New-HostPredicate([string]$Column, [string[]]$Patterns) {
    $clauses = foreach ($p in $Patterns) {
        $d = $p.Trim().TrimStart('.').ToLowerInvariant() -replace "'", "''"
        if (-not $d) { continue }
        if ($d -match '[\*\?]') {
            $like = ($d -replace '\*', '%') -replace '\?', '_'
            "lower($Column) LIKE '$like' OR lower($Column) LIKE '.$like'"
        } else {
            "lower($Column) = '$d' OR lower($Column) = '.$d' OR lower($Column) LIKE '%.$d'"
        }
    }
    if (-not $clauses) { return $null }
    return '(' + ($clauses -join ' OR ') + ')'
}

function New-NamePredicate([string[]]$Patterns) {
    # Exact, case-sensitive cookie-name match; 'x*' = prefix. LIKE is escaped so the
    # '_' in names like _ga is literal (in SQL LIKE a bare '_' matches any character).
    $clauses = foreach ($n in $Patterns) {
        $v = $n -replace "'", "''"
        if ($v.EndsWith('*')) {
            $prefix = $v.TrimEnd('*') -replace '%', '\%' -replace '_', '\_'
            "name LIKE '$prefix%' ESCAPE '\'"
        } else {
            "name = '$v'"
        }
    }
    return '(' + ($clauses -join ' OR ') + ')'
}

function Get-DeleteWhere($Store) {
    $now = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
    if ($Store.Engine -eq 'Chromium') {
        $col = 'host_key'
        $nowChromium = ($now + 11644473600) * 1000000
        $expiredClause = "(has_expires = 1 AND expires_utc > 0 AND expires_utc < $nowChromium)"
    } else {
        $col = 'host'
        $expiredClause = "(expiry > 0 AND expiry < $now)"
    }
    $conds = @()
    if ($Domain)     { $p = New-HostPredicate $col $Domain;     if ($p) { $conds += $p } }
    if ($KeepDomain) { $p = New-HostPredicate $col $KeepDomain; if ($p) { $conds += ('NOT ' + $p) } }
    if ($Trackers) {
        # Tracker OR expired when both are set; trackers alone otherwise.
        $tp = '(' + (New-HostPredicate $col $TrackerDomains) + ' OR ' + (New-NamePredicate $TrackerCookieNames) + ')'
        if ($Expired) { $conds += "($tp OR $expiredClause)" } else { $conds += $tp }
    } elseif ($Expired) {
        $conds += $expiredClause
    }
    if ($conds.Count -eq 0) { return $null }
    return ($conds -join ' AND ')
}

function Backup-Store($Store, [string]$Stamp) {
    # Copies the DB + sidecars to BackupDir\<stamp>\<Browser>_<Profile>\ and records it in manifest.json.
    # Uses shared-read copies, so it works whether or not the browser is open.
    $folder   = ConvertTo-SafeName ("{0}_{1}" -f $Store.Browser, $Store.Profile)
    $stampDir = Join-Path $BackupDir $Stamp
    $dest     = Join-Path $stampDir $folder
    New-Item -ItemType Directory -Path $dest -Force -WhatIf:$false | Out-Null
    $files = @($Store.Db) + (Get-StoreSidecars $Store.Db)
    foreach ($f in $files) { Copy-FileShared $f (Join-Path $dest ([IO.Path]::GetFileName($f))) }
    Write-SnapshotManifest $stampDir ([pscustomobject]@{
        Browser    = $Store.Browser
        Profile    = $Store.Profile
        Engine     = $Store.Engine
        OriginalDb = $Store.Db
        Folder     = $folder
        DbFile     = [IO.Path]::GetFileName($Store.Db)
        Files      = @($files | ForEach-Object { [IO.Path]::GetFileName($_) })
        TakenAt    = (Get-Date).ToString('o')
    })
    return $dest
}

function Invoke-Delete($Stores) {
    Write-Section 'COOKIE JAR PURGE'
    if ($Stores.Count -eq 0) { Write-Warning 'No cookie stores found for the selected browsers/profiles.'; return }

    if (-not $Wipe -and -not ($Domain -or $KeepDomain -or $Expired -or $Trackers)) {
        Write-Host '  -Delete needs a rule. Pick one (add -WhatIf to preview, -CloseBrowsers if a browser is open):' -ForegroundColor Yellow
        Write-Host '    -Wipe                              everything, every store'                    -ForegroundColor Gray
        Write-Host '    -KeepDomain site1,site2            everything except those sites (keep logins)' -ForegroundColor Gray
        Write-Host '    -Domain site1,site2                only those sites (wildcards ok: *ads*)'      -ForegroundColor Gray
        Write-Host '    -Trackers                          the built-in ad/tracking domain list'       -ForegroundColor Gray
        Write-Host '    -Expired                           only cookies already past their expiry'     -ForegroundColor Gray
        Write-Host ''
        return
    }
    if (-not $Wipe -and -not $Script:Sql) {
        throw 'Selective delete needs a SQLite engine (sqlite3.exe on PATH, or Python 3). Use -Wipe, or: winget install SQLite.SQLite'
    }

    # 1. Browsers must be closed for the stores we touch.
    $byBrowser = $Stores | Group-Object Browser
    $skip = @{}
    foreach ($g in $byBrowser) {
        $procs = @(Get-BrowserProcesses $g.Group[0])
        if ($procs.Count -eq 0) { continue }
        if ($CloseBrowsers) {
            if ($PSCmdlet.ShouldProcess(("{0} ({1} process(es))" -f $g.Name, $procs.Count), 'Close browser')) {
                $procs | Stop-Process -Force -ErrorAction SilentlyContinue
                $deadline = (Get-Date).AddSeconds(15)
                while ((Get-Date) -lt $deadline -and @(Get-BrowserProcesses $g.Group[0]).Count -gt 0) { Start-Sleep -Milliseconds 300 }
                Start-Sleep -Milliseconds 500   # let file locks drop
                if (@(Get-BrowserProcesses $g.Group[0]).Count -gt 0) {
                    Write-Warning ("{0} still running — skipping its stores." -f $g.Name); $skip[$g.Name] = $true
                } else { Write-Host ("Closed {0}." -f $g.Name) -ForegroundColor Yellow }
            } else { $skip[$g.Name] = $true }
        } else {
            Write-Warning ("{0} is running ({1} process(es)). Close it or use -CloseBrowsers. Skipping." -f $g.Name, $procs.Count)
            $skip[$g.Name] = $true
        }
    }

    # 2. Per-store work.
    $stamp    = Get-Date -Format 'yyyyMMdd_HHmmss'
    $results  = @()
    $declined = 'declined'; if ($WhatIfPreference) { $declined = 'preview (-WhatIf)' }
    foreach ($s in $Stores) {
        $label = "{0} / {1}" -f $s.Browser, $s.Profile
        $r = [pscustomobject]@{ Browser = $s.Browser; Profile = $s.Profile; Before = $null; Deleted = 0; After = $null; Action = ''; Backup = '' }
        $results += $r
        if ($skip.ContainsKey($s.Browser)) { $r.Action = 'skipped (running)'; continue }

        try {
            $table = 'cookies'; if ($s.Engine -eq 'Gecko') { $table = 'moz_cookies' }

            if ($Wipe) {
                $files = @($s.Db) + (Get-StoreSidecars $s.Db)
                # Count rows first (when an engine is available) so the result table is informative.
                $count = $null
                if ($Script:Sql) {
                    try { $count = [int](Invoke-CookieSql -Db $s.Db -Sql "SELECT COUNT(*) AS n FROM $table" -Mode Query)[0].n } catch { }
                }
                $r.Before = $count
                $what = "Wipe cookie store ({0} file(s), {1} KB)" -f $files.Count, $s.SizeKB
                if ($count -ne $null) { $what = "Wipe cookie store ({0} cookies, {1} file(s), {2} KB)" -f $count, $files.Count, $s.SizeKB }
                if ($PSCmdlet.ShouldProcess($label, $what)) {
                    if (-not $NoBackup) { $r.Backup = Backup-Store $s $stamp }
                    foreach ($f in $files) { Remove-Item -LiteralPath $f -Force }
                    $r.Deleted = $count; $r.After = 0; $r.Action = 'wiped'
                } else { $r.After = $count; $r.Action = $declined }
                continue
            }
            $where = Get-DeleteWhere $s
            $total = [int](Invoke-CookieSql -Db $s.Db -Sql "SELECT COUNT(*) AS n FROM $table" -Mode Query)[0].n
            $hit   = [int](Invoke-CookieSql -Db $s.Db -Sql "SELECT COUNT(*) AS n FROM $table WHERE $where" -Mode Query)[0].n
            $r.Before = $total
            if ($hit -eq 0) { $r.After = $total; $r.Action = 'no matches'; continue }

            if ($PSCmdlet.ShouldProcess($label, ("Delete {0} of {1} cookies" -f $hit, $total))) {
                if (-not $NoBackup) { $r.Backup = Backup-Store $s $stamp }
                $n = Invoke-CookieSql -Db $s.Db -Sql "DELETE FROM $table WHERE $where" -Mode Exec
                $r.Deleted = $n
                $r.After   = [int](Invoke-CookieSql -Db $s.Db -Sql "SELECT COUNT(*) AS n FROM $table" -Mode Query)[0].n
                $r.Action  = 'deleted'
            } else { $r.After = $total; $r.Deleted = $hit; $r.Action = $declined }
        } catch {
            $r.Action = 'ERROR: ' + $_.Exception.Message
        }
    }

    # 3. Legacy IE / WebView cache.
    if ($IncludeLegacy) {
        if ($PSCmdlet.ShouldProcess('IE / WebView cookie cache', 'Clear via InetCpl.cpl')) {
            Start-Process -FilePath 'RunDll32.exe' -ArgumentList 'InetCpl.cpl,ClearMyTracksByProcess 2' -Wait -WindowStyle Hidden
            $results += [pscustomobject]@{ Browser = 'Legacy'; Profile = 'INetCookies'; Before = $null; Deleted = $null; After = $null; Action = 'cleared'; Backup = '' }
        }
    }

    Write-Section 'RESULT'
    $results | Format-Table Browser, Profile, Before, Deleted, After, Action -AutoSize | Out-Host
    # @() keeps an empty result a real (empty) array; under StrictMode, reading .Sum
    # from nothing throws in Windows PowerShell 5.1 (e.g. a -WhatIf preview).
    $doneRows = @($results | Where-Object { $_.Action -eq 'deleted' -or $_.Action -eq 'wiped' })
    $deletedTotal = 0
    if ($doneRows.Count -gt 0) { $deletedTotal = [int](($doneRows | Measure-Object Deleted -Sum).Sum) }
    if ($deletedTotal -gt 0) { Write-Host ("Deleted {0} cookies." -f $deletedTotal) -ForegroundColor Green }
    $backups = @($results | Where-Object { $_.Backup })
    if ($backups.Count -gt 0) { Write-Host ("Backups: {0}" -f (Join-Path $BackupDir $stamp)) -ForegroundColor DarkGray }
    if ($PassThru) { return $results }
}

# ── SNAPSHOT / RESTORE ──

function Invoke-Snapshot($Stores) {
    Write-Section 'COOKIE JAR SNAPSHOT'
    if ($Stores.Count -eq 0) { Write-Warning 'No cookie stores found for the selected browsers.'; return }
    $stamp   = Get-Date -Format 'yyyyMMdd_HHmmss'
    $results = @()
    foreach ($s in $Stores) {
        $r = [pscustomobject]@{ Browser = $s.Browser; Profile = $s.Profile; Files = 0; SizeKB = $s.SizeKB; Status = ''; Folder = '' }
        $results += $r
        $r.Files = (@($s.Db) + (Get-StoreSidecars $s.Db)).Count
        if ($WhatIfPreference) { $r.Status = 'preview (-WhatIf)'; continue }
        try { $r.Folder = Backup-Store $s $stamp; $r.Status = 'saved' }
        catch {
            $r.Status = 'LOCKED — close ' + $s.Browser + ' and retry'
            if ($_.Exception.Message -notmatch 'being used by another process') { $r.Status = 'ERROR: ' + $_.Exception.Message }
        }
    }
    $results | Format-Table Browser, Profile, Files, SizeKB, Status -AutoSize | Out-Host
    $saved = @($results | Where-Object { $_.Status -eq 'saved' }).Count
    if ($saved -gt 0) {
        Write-Host ("Snapshot: {0}   ({1} store(s))" -f (Join-Path $BackupDir $stamp), $saved) -ForegroundColor Green
        Write-Host ("  analyze:  .\CCCC.ps1 -From {0}" -f $stamp) -ForegroundColor DarkGray
        Write-Host ("  restore:  .\CCCC.ps1 -Restore {0} -CloseBrowsers" -f $stamp) -ForegroundColor DarkGray
    }
}

function Invoke-Restore([string]$Dir, $LiveStores) {
    Write-Section 'COOKIE JAR RESTORE'
    Write-Host ("  source: {0}" -f $Dir) -ForegroundColor Yellow
    $snap = Get-SnapshotStores $Dir $Browser
    if ($snap.Count -eq 0) { Write-Warning 'No stores found in that snapshot for the selected browsers.'; return }

    # 1. Resolve each snapshot store to a live target path (manifest first, folder-name match second).
    $jobs = @()
    foreach ($s in $snap) {
        $target = $null
        if ($s.OriginalDb) { $target = $s.OriginalDb }
        else {
            $match = $LiveStores | Where-Object {
                $_.Browser -eq $s.Browser -and (ConvertTo-SafeName ("{0}_{1}" -f $_.Browser, $_.Profile)) -eq $s.Folder
            } | Select-Object -First 1
            if ($match) { $target = $match.Db }
        }
        $jobs += [pscustomobject]@{ Snap = $s; Target = $target }
    }

    # 2. Browsers must be closed for the stores we overwrite.
    $skip = @{}
    foreach ($g in ($jobs | Group-Object { $_.Snap.Browser })) {
        $probe = $g.Group[0].Snap
        $procs = @(Get-BrowserProcesses $probe)
        if ($procs.Count -eq 0) { continue }
        if ($CloseBrowsers) {
            if ($PSCmdlet.ShouldProcess(("{0} ({1} process(es))" -f $g.Name, $procs.Count), 'Close browser')) {
                $procs | Stop-Process -Force -ErrorAction SilentlyContinue
                $deadline = (Get-Date).AddSeconds(15)
                while ((Get-Date) -lt $deadline -and @(Get-BrowserProcesses $probe).Count -gt 0) { Start-Sleep -Milliseconds 300 }
                Start-Sleep -Milliseconds 500
                if (@(Get-BrowserProcesses $probe).Count -gt 0) {
                    Write-Warning ("{0} still running — skipping its stores." -f $g.Name); $skip[$g.Name] = $true
                } else { Write-Host ("Closed {0}." -f $g.Name) -ForegroundColor Yellow }
            } else { $skip[$g.Name] = $true }
        } else {
            Write-Warning ("{0} is running ({1} process(es)). Close it or use -CloseBrowsers. Skipping." -f $g.Name, $procs.Count)
            $skip[$g.Name] = $true
        }
    }

    # 3. Copy back, safety-snapshotting the current live store first.
    $safetyStamp = (Get-Date -Format 'yyyyMMdd_HHmmss') + '_pre-restore'
    $results = @()
    foreach ($j in $jobs) {
        $s = $j.Snap
        $r = [pscustomobject]@{ Browser = $s.Browser; Profile = $s.Profile; Action = ''; Target = $j.Target }
        $results += $r
        if (-not $j.Target) { $r.Action = 'no matching live profile'; continue }
        if ($skip.ContainsKey($s.Browser)) { $r.Action = 'skipped (running)'; continue }
        try {
            $files = @($s.Db) + (Get-StoreSidecars $s.Db)
            $label = "{0} / {1}" -f $s.Browser, $s.Profile
            if ($PSCmdlet.ShouldProcess($label, ("Restore {0} file(s) over {1}" -f $files.Count, $j.Target))) {
                $targetDir = Split-Path -Parent $j.Target
                if (-not (Test-Path -LiteralPath $targetDir)) { New-Item -ItemType Directory -Path $targetDir -Force | Out-Null }
                if (-not $NoBackup -and (Test-Path -LiteralPath $j.Target)) {
                    $live = [pscustomobject]@{ Browser = $s.Browser; Profile = $s.Profile; Engine = $s.Engine; Db = $j.Target }
                    Backup-Store $live $safetyStamp | Out-Null
                }
                # Drop stale sidecars so the restored DB is not paired with a foreign journal.
                foreach ($old in (Get-StoreSidecars $j.Target)) { Remove-Item -LiteralPath $old -Force }
                foreach ($f in $files) { Copy-Item -LiteralPath $f -Destination (Join-Path $targetDir ([IO.Path]::GetFileName($f))) -Force }
                $r.Action = 'restored'
            } else {
                $r.Action = 'declined'; if ($WhatIfPreference) { $r.Action = 'preview (-WhatIf)' }
            }
        } catch { $r.Action = 'ERROR: ' + $_.Exception.Message }
    }

    Write-Section 'RESULT'
    $results | Format-Table Browser, Profile, Action, Target -AutoSize | Out-Host
    if (-not $NoBackup -and @($results | Where-Object { $_.Action -eq 'restored' }).Count -gt 0) {
        Write-Host ("Pre-restore safety snapshot: {0}" -f (Join-Path $BackupDir $safetyStamp)) -ForegroundColor DarkGray
    }
}

# ── MAIN ──

# Housekeeping calls pass -WhatIf:$false so a -WhatIf preview still runs end to end.
New-Item -ItemType Directory -Path $Script:Scratch -Force -WhatIf:$false | Out-Null
try {
    $Script:Sql = Find-SqlEngine
    if ($Script:Sql) {
        if ($Script:Sql.Kind -eq 'python') { Initialize-PyHelper }
        Write-Host ("SQLite engine: {0} ({1})" -f $Script:Sql.Kind, $Script:Sql.Path) -ForegroundColor DarkGray
    } else {
        Write-Warning 'No SQLite engine found (sqlite3.exe or Python 3). Analysis and selective delete are unavailable; -Wipe still works.'
        Write-Warning 'Install one with:  winget install SQLite.SQLite   (or)   winget install Python.Python.3.13'
    }

    $stores = Get-CookieStores $Browser
    switch ($PSCmdlet.ParameterSetName) {
        'Delete'   { Invoke-Delete $stores }
        'Snapshot' { Invoke-Snapshot $stores }
        'Restore'  { Invoke-Restore (Resolve-SnapshotDir $Restore) $stores }
        default {
            if ($From) {
                $From   = Resolve-SnapshotDir $From
                $stores = Get-SnapshotStores $From $Browser
            }
            if (-not $Script:Sql) {
                Write-Section 'COOKIE STORES (no engine — sizes only)'
                $stores | Format-Table Browser, Profile, SizeKB, Db -AutoSize | Out-Host
            } else {
                Invoke-Analyze $stores
            }
        }
    }
} finally {
    # Scratch holds copies of cookie DBs — always remove.
    Remove-Item -LiteralPath $Script:Scratch -Recurse -Force -WhatIf:$false -ErrorAction SilentlyContinue
}
