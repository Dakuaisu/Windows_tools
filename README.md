# Windows tools

A collection of standalone PowerShell utilities for Windows. Each script is
self-contained — there is nothing to install, no modules to import, and no
shared library between them. Copy one to another machine and it works.

Common conventions across every tool here:

- **Non-destructive by default.** Tools that could lose data either preview
  first, write an undo log, or only report and let you decide.
- **`-DryRun` (or preview mode) everywhere it matters.** Run it, read it, then
  run it for real.
- **Output goes to `%LOCALAPPDATA%\<ToolName>\`**, never next to the script,
  so the folder stays clean and reports survive moving the scripts around. If
  that folder cannot be written to — a redirected profile, a locked-down
  machine, an `-OutputDir` you do not have rights to — the tool falls back to
  `%TEMP%\<ToolName>\`, tells you it did, and carries on. The one exception is
  `Rename-Bulk`, which *refuses to run* rather than rename anything it could
  not write an undo log for.
- **Rolling history files are trimmed.** Anything a tool appends to on every
  run has a bound. Where the columns of a history file change between versions,
  the old file is archived beside it as `<name>_upto_<timestamp>.csv` rather
  than deleted or appended to until it errors.
- **Comment-based help.** `Get-Help .\Some-Tool.ps1 -Full` works on all of them.

---

## Layout

```
tools/
├─ Toolbox.cmd      double-click: one window for everything below
├─ Toolbox.ps1      the launcher itself
├─ diagnostics/     read-only "what is wrong with this machine" tools
├─ backup/          capture machine state so it can be rebuilt
├─ productivity/    day-to-day utilities
├─ shell/           Explorer / registry integrations
└─ DownloadsJanitor/  a shareable package, not a single script
```

| Tool | Location | One line |
| ---- | -------- | -------- |
| [Get-SystemSnapshot](#get-systemsnapshotps1) | `diagnostics/` | Whole-machine health report as a single HTML page |
| [Get-StartupImpact](#get-startupimpactps1) | `diagnostics/` | Everything that launches at boot or logon, in one table |
| [Test-InternetHealth](#test-internethealthps1) | `diagnostics/` | Latency, jitter, loss, DNS and TTFB, logged over time |
| [Get-BatteryReport](#get-batteryreportps1) | `diagnostics/` | Battery wear and capacity, tracked across months |
| [Export-InstalledSoftware](#export-installedsoftwareps1) | `backup/` | Inventory everything installed, generate a restore script |
| [Rename-Bulk](#rename-bulkps1) | `productivity/` | Batch rename with preview, collision checks and undo |
| [Set-FocusMode](#set-focusmodeps1) | `productivity/` | Close distractions, block sites, mute toasts — reversibly |
| [New-ProjectScaffold](#new-projectscaffoldps1) | `productivity/` | New project folder, git repo and starter files |
| [Claude Code context menu](#claude-code-context-menu) | `shell/` | "Open Claude Code here" right-click entry |
| [Tidy My Downloads](#downloads-janitor--downloadsjanitor) | `DownloadsJanitor/` | Sorts Downloads by type and month, reports old clutter, full undo |
| [Toolbox](#toolbox--one-window-for-all-of-them) | root | One window for all of the above |

---

## Toolbox — one window for all of them

**`Toolbox.cmd`** / **`Toolbox.ps1`**

Double-click `Toolbox.cmd`. Pick a tool on the left, fill in the form on the
right, press Run and watch the output appear at the bottom.

```powershell
.\Toolbox.ps1
```

Nothing in it is written per tool. **Every control is generated from the target
script's own `param()` block and comment-based help**, read through the
PowerShell parser — 65 controls across the nine tools, none of them hardcoded:

| In the script | In the window |
| ------------- | ------------- |
| `[switch]$Open` | a checkbox |
| `[ValidateSet('blank','node',…)]` | a dropdown |
| `[ValidateSet]` on a `[string[]]` | a multi-select list |
| `[int]` with `[ValidateRange(1,60)]` | a number box showing the range |
| `[string[]]$Targets` | a box taking one value per line |
| anything named `*Path`/`*Dir`/`*Root` | a text box with **Browse** |
| `.PARAMETER` help text | the grey hint under the label |

Add a parameter to a tool tomorrow and its form grows a control the next time
you open the window. **The dependency only points one way** — no script knows
this file exists, so the promise at the top of this README still holds: copy
any one of them to a machine that has never seen the toolbox and it runs.

Scripts with more than one parameter set get a **Mode** selector, and only the
parameters valid in that mode are shown — pick `Regex` on `Rename-Bulk` and
`-Template` disappears, so you cannot build a command PowerShell would reject.

A few details worth knowing:

- Anything you leave at its default is **left off the command line entirely**,
  so the script computes its own default rather than being handed a copy of it.
  The command actually run is echoed at the top of the output pane.
- Tools run as a separate PowerShell process with output streamed live, so the
  window stays responsive and **Stop** genuinely kills the run.
- **Run as administrator** relaunches that one tool elevated in its own console
  — output cannot be captured across the UAC boundary, and running the whole
  window elevated would put every report in the wrong profile.
- The `DownloadsJanitor` `.cmd` files open in their own window instead, because
  they ask questions and you need somewhere to answer them.

---

## Before you start

**Requirements:** Windows 10 or 11. Windows PowerShell 5.1 (built in) or
PowerShell 7 — every script works on both. No modules to install.

**If a script refuses to run**, PowerShell's execution policy is blocking it.
You do not need to change any system setting; just run it this way:

```powershell
powershell -ExecutionPolicy Bypass -File .\diagnostics\Get-SystemSnapshot.ps1
```

**If you were sent these files** over email, chat, or a USB stick, Windows
marks them as untrusted and will block them regardless of execution policy.
Unblock them once, from the `tools` folder:

```powershell
Get-ChildItem -Recurse -Filter *.ps1 | Unblock-File
```

**Administrator** is only needed for one thing: the site-blocking half of
`Set-FocusMode`. Everything else runs as a normal user, and `Set-FocusMode`
degrades gracefully — it skips site blocking with a warning and still does the
rest.

**Reading the help.** Every script documents itself:

```powershell
Get-Help .\productivity\Rename-Bulk.ps1 -Full
```

---

## diagnostics/

Read-only. These four tools inspect the machine and write reports. None of
them change a setting, install anything, or delete a file.

### Get-SystemSnapshot.ps1

**`diagnostics/Get-SystemSnapshot.ps1`**

One self-contained HTML page describing the machine right now: hardware, CPU,
memory modules, physical disks with SMART health, volumes with free-space bars,
graphics, network adapters, battery, Defender status, recent Windows updates,
devices in an error state, and the biggest memory consumers.

A **"Needs attention"** panel is built first and pinned to the top — low disk
space, unhealthy drives, pending reboots, stale updates, disabled real-time
protection. If that panel is empty, the machine passed every check.

The page has no external assets, so you can email it or attach it to a support
ticket. Keeping one per month and comparing them is the real use.

```powershell
.\diagnostics\Get-SystemSnapshot.ps1 -Open
```

| Option | Meaning |
| ------ | ------- |
| `-Open` | Open the report in your browser when done |
| `-TopProcesses <n>` | How many processes to list, by memory. Default 12, `0` to skip |
| `-LowDiskPercent <n>` | Free space below this raises an attention item. Default 10 |
| `-ProbeTimeoutSeconds <n>` | Give up on the two checks that can hang on a broken machine — Defender status and device enumeration — after this long, and carry on with the rest. Default 30 |
| `-Quiet` | No console output — for scheduled runs |
| `-OutputDir <path>` | Where the report lands |

**Output:** `%LOCALAPPDATA%\SystemSnapshot\snapshot_<timestamp>.html`, keeping
the 20 most recent.
**Exit code:** `0` nothing needs attention, `1` at least one attention item.

---

### Get-StartupImpact.ps1

**`diagnostics/Get-StartupImpact.ps1`**

Task Manager shows you Run keys and Startup folders. Half of what actually
slows a boot is a scheduled task or an auto-start service, so this folds all
of those into one table with publisher, signature status and target path —
along with two places Task Manager does not show at all: the policy `Run` keys,
and a Winlogon `Userinit`/`Shell` value that has been changed from its default.

Where an entry launches something *through* a wrapper — `cmd /c ...`,
`powershell -File ...`, `rundll32 <dll>` — the table reports what actually gets
run rather than the interpreter. That matters, because every one of those
wrappers is a Microsoft-signed binary in `System32` and would otherwise lend a
clean bill of health to whatever it starts.

> Active Setup is deliberately not swept. It does run once per user at logon,
> but on a real machine it is dozens of rows of Microsoft component
> registration, and burying three real findings under that is worse than
> leaving it out.

It also flags **suspicious** entries — unsigned binaries living in a
user-writable location, or entries whose target executable no longer exists.
That second case is common and worth cleaning: uninstalled software regularly
leaves a logon task behind pointing at nothing.

> The **Impact** column is this script's own heuristic (binary size, signature,
> install location, delayed-start flag). It is deliberately *not* Task
> Manager's number — Windows derives that from boot tracing that scripts
> cannot read. Treat it as a sorting hint, not a measurement.

```powershell
.\diagnostics\Get-StartupImpact.ps1
```

| Option | Meaning |
| ------ | ------- |
| `-SkipSignatureCheck` | Much faster; drops the Signed and Suspicious columns |
| `-OnlyEnabled` | Hide entries Windows has already disabled |
| `-All` | Include Microsoft services and `\Microsoft\` tasks (hundreds of rows) |
| `-IncludeServices:$false` | Drop auto-start services |
| `-IncludeTasks:$false` | Drop logon/boot scheduled tasks |
| `-OutputDir <path>` | Where the report lands |

**Output:** `%LOCALAPPDATA%\StartupImpact\startup-report_<timestamp>.txt` and
`.csv`, keeping the 12 most recent pairs.
**Exit code:** `1` if anything suspicious was found, else `0`.

---

### Test-InternetHealth.ps1

**`diagnostics/Test-InternetHealth.ps1`**

Four probe groups, in this order because that is the order the answers matter:

1. **Gateway** — separates "my wifi is bad" from "the internet is bad".
2. **DNS** — slow resolution feels exactly like slow internet, different fix.
3. **Latency** — ICMP to public anchors: average, min, max, jitter, loss.
4. **HTTP** — time to first byte on real HTTPS endpoints, which is what a
   browser actually waits on.

Every run appends a row to `history.csv`. That history is the point — it is
what you show an ISP that insists nothing is wrong. It keeps the most recent
`-HistoryLimit` rows (default 5000); pass `0` to keep everything.

There is a fifth check that does not fit the list above: a **captive portal**
probe. Every HTTPS target here simply *fails* behind a hotel or airport
sign-in page, which looks identical to a DNS fault, so the script also asks a
plain-HTTP endpoint whose exact response body it knows. Anything else
answering means something is sitting in the middle. Being unable to reach that
endpoint at all is treated as an ordinary outage, not a portal.

Throughput is only measured with `-SpeedTest`, because it downloads ~25 MB and
you do not want that on a metered connection or on every scheduled run.

```powershell
.\diagnostics\Test-InternetHealth.ps1
```

Watch a flaky connection for an hour, one pass a minute:

```powershell
.\diagnostics\Test-InternetHealth.ps1 -Count 60 -IntervalSeconds 60 -Quiet
```

| Option | Meaning |
| ------ | ------- |
| `-Count <n>` | Number of passes. Default 1 |
| `-IntervalSeconds <n>` | Wait between passes. Default 60 |
| `-PingCount <n>` | Echoes per target per pass. Default 10 |
| `-SpeedTest` | Also measure download throughput (~25 MB) |
| `-LatencyWarnMs <n>` | Average above this counts as degraded. Default 120 |
| `-LossWarnPercent <n>` | Loss at or above this counts as degraded. Default 2 |
| `-SkipCaptivePortalCheck` | Skip the sign-in-page probe |
| `-HistoryLimit <n>` | Rows kept in `history.csv`. Default 5000, `0` for unlimited |
| `-Quiet` | No per-probe output; still logs and sets the exit code |
| `-Targets`, `-HttpTargets`, `-DnsNames` | Override the default probe lists |

**Output:** `%LOCALAPPDATA%\InternetHealth\history.csv` and
`internet-health.log`.
**Exit code:** `0` healthy, `1` degraded, `2` no usable connection. A captive
portal counts as `2` — nothing reaches the internet until you sign in.

ICMP being blocked upstream while HTTPS works is recognised as a normal,
healthy configuration rather than reported as a fault.

---

### Get-BatteryReport.ps1

**`diagnostics/Get-BatteryReport.ps1`**

Wraps `powercfg /batteryreport` — both the HTML you can read and the XML you
can parse — and cross-checks it against the live WMI battery classes, because
the two disagree often enough to be worth seeing side by side.

The number that matters is **Health**: full-charge capacity as a percentage of
design capacity. Below ~80% a laptop starts wanting to stay plugged in; below
~60% the battery is effectively done.

Every run appends to `battery-history.csv`, so after a few months you see the
wear *curve* instead of a single meaningless snapshot. Running this on a
schedule is the whole idea. Desktops and VMs are detected and exit cleanly.

Machines with **two packs** — ThinkPads with a bridge battery, Surface Books,
some gaming laptops — are read as a set: capacities are summed, health is
computed from those totals, the cycle count is the highest of the packs, and
each pack also gets its own line so one failing cell is still visible.

```powershell
.\diagnostics\Get-BatteryReport.ps1 -Open
```

| Option | Meaning |
| ------ | ------- |
| `-Open` | Open the powercfg HTML report when done |
| `-Days <n>` | Days of history in the HTML report. Default 14, max 60 |
| `-WarnHealthPercent <n>` | Health below this counts as worn. Default 80 |
| `-HistoryLimit <n>` | Rows kept in `battery-history.csv`. Default 5000, `0` for unlimited |
| `-Quiet` | No console output |

**Output:** `%LOCALAPPDATA%\BatteryReport\` — timestamped HTML and XML reports
(12 most recent kept) plus `battery-history.csv`.
**Exit code:** `0` healthy, `1` worn, `2` no battery present, `3` battery data
unreadable.

---

## backup/

### Export-InstalledSoftware.ps1

**`backup/Export-InstalledSoftware.ps1`**

Read-only. Inventories everything installed and writes a dated folder
containing a generated **`restore.ps1`** you can run on a fresh Windows
install. Highest payoff per line of any tool here — you find out whether it
worked on the worst possible day, so run it now and again occasionally.

Sources collected:

- Registry uninstall keys (HKLM 64-bit, 32-bit, HKCU) — the authoritative
  Programs and Features list, including software winget has never heard of
- `winget export` → a real, importable JSON
- Microsoft Store / Appx packages
- VS Code, Insiders, Cursor and Windsurf extensions
- Global npm packages, pip packages, PowerShell Gallery modules
- Chocolatey and Scoop, if present

The generated `restore.ps1` is deliberately conservative. `winget import` and
the per-ecosystem installs are live commands; anything with no reliable
automated install path is emitted as a **commented checklist** at the bottom,
so you review it rather than trust it.

Deciding whether a Programs-and-Features entry is already covered by the winget
export is a guess — display names and winget package ids often share no words
at all. The matcher errs toward **listing** things: a checklist entry you did
not need costs a moment's reading, whereas an omitted one costs software you do
not notice is missing until you need it. The checklist header says how many
entries were assumed covered.

Every external tool it shells out to — winget, the editor CLIs, npm, pip,
choco, scoop, the Appx enumeration — runs under `-ToolTimeoutSeconds`. None of
them offer a timeout of their own and any of them can hang; when one does, that
collector is abandoned with a warning and the rest of the inventory still
completes.

```powershell
.\backup\Export-InstalledSoftware.ps1 -Open
```

On the new machine, from inside the inventory folder:

```powershell
.\restore.ps1 -WhatIf
```

| Option | Meaning |
| ------ | ------- |
| `-Open` | Open the output folder when done |
| `-Sources <list>` | Limit collectors, e.g. `-Sources Registry,Winget,VSCode` |
| `-IncludeVersions` | Pin exact versions in the winget export. Off by default — pinned versions often fail to restore because that build is no longer offered |
| `-ToolTimeoutSeconds <n>` | Give up on any one external tool after this long. Default 120 |
| `-OutputDir <path>` | Default `Documents\SoftwareInventory` |

**Output:** `Documents\SoftwareInventory\Inventory_<date>\` containing
`inventory.json`, `installed-programs.csv`, `winget-export.json`,
`store-apps.csv`, `pip-requirements.txt`, `*-extensions.txt`, `restore.ps1`
and a `README.txt` explaining each.

> winget ships as a Store app-execution alias that is not always on `PATH`.
> The script falls back to its known location, so this works even in shells
> where `winget` alone would not resolve.

---

## productivity/

### Rename-Bulk.ps1

**`productivity/Rename-Bulk.ps1`**

Batch rename by regex or by template. **Preview is the default** — the script
shows exactly what it would do and changes nothing until you add `-Apply`.
That inversion is deliberate: a bad bulk rename is tedious to unpick by hand,
and typing one extra switch costs nothing next to getting it wrong.

Safety properties worth knowing:

- **All-or-nothing validation.** Illegal characters, reserved device names
  (`CON`, `LPT1`, …), empty names, names colliding with each other or with a
  file already on disk — all caught *before* the first rename happens.
- **Swaps and rotations work** (`a→b`, `b→a`), and so do case-only renames,
  which a plain `Rename-Item` refuses on NTFS. Both go through a temporary
  name automatically.
- **Every applied run writes an undo log.** `-UndoLast` reverses the most
  recent one. The log folder is proved writable *before* the first rename, so a
  run can never turn out to be irreversible after the fact.
- **Folders are renamed after their contents**, so `-Recurse
  -IncludeDirectories` over a nested tree works: renaming a parent can never
  invalidate a path a later rename still needs. An undo replays the run
  backwards for the same reason.

Two ways to build the new name:

**Regex** — `.NET` regex against the file name, capture groups as `$1`, `$2`:

```powershell
.\productivity\Rename-Bulk.ps1 -Pattern '^IMG_' -Replacement 'photo-'
```

**Template** — tokens `{name}` `{ext}` `{n}` `{parent}` `{date}` `{time}`:

```powershell
.\productivity\Rename-Bulk.ps1 -Filter *.jpg -Template 'Holiday-{n}{ext}' -SortBy Date -Apply
```

Reverse the last run:

```powershell
.\productivity\Rename-Bulk.ps1 -UndoLast
```

| Option | Meaning |
| ------ | ------- |
| `-Apply` | Actually rename. Without it, preview only |
| `-Path <dir>` | Folder to work in. Default: current directory |
| `-Filter <wildcard>` | Which items to consider. Default `*` |
| `-Recurse` | Descend into subfolders |
| `-IncludeDirectories` | Rename folders too |
| `-CaseSensitive` | Make `-Pattern` case-sensitive (regex mode) |
| `-SortBy Name\|Date\|Size` | Ordering used to assign `{n}` |
| `-StartAt <n>`, `-Pad <n>` | First sequence number and zero-padding. Default 1 and 3 |
| `-UndoLast` / `-UndoFrom <csv>` | Reverse the last run, or a specific one |

**Output:** `%LOCALAPPDATA%\RenameBulk\undo_<timestamp>.csv`, 50 most recent
kept. A reversed log is renamed to `.reversed.csv` so `-UndoLast` cannot
re-run it.
**Exit code:** `1` if validation found problems (nothing was renamed), else `0`.
An undo that had to skip items — because they were moved, deleted, or their old
name has since been taken again — still exits `0` and reports the skips on the
console; nothing is ever overwritten to force one through.

---

### Set-FocusMode.ps1

**`productivity/Set-FocusMode.ps1`**

Turns distractions off and puts every one of them back afterwards. Everything
it changes is recorded in a state file **before** it is changed, and `-Off`
restores from that record rather than from assumptions — so you can close the
terminal, reboot, or come back three days later and `-Off` still knows exactly
what to undo.

What it touches:

- **Apps** — closes them politely with `CloseMainWindow`, the same signal as
  clicking the X, so nothing loses unsaved work. Apps that refuse are reported,
  not killed, unless you pass `-Force`.
- **Sites** — adds `0.0.0.0` entries to the hosts file between `BEGIN/END
  FOCUSMODE` markers, so removal is exact and the rest of your hosts file is
  never touched: it is read and written back in whatever encoding and line
  ending it already used, so non-ASCII content and hand-made edits survive
  byte for byte. Backed up first. **Needs an elevated shell**; without one this
  step is skipped with a warning and everything else still runs.
- **Notifications** — sets the per-user `ToastEnabled` flag to 0. The previous
  value is saved and restored. This is not Windows' own Focus Assist, which has
  no supported scripting interface, but it stops the thing that interrupts you.

First run writes a config file and tells you where it is. **Edit that, not the
script** — it holds the app list and the site list.

```powershell
.\productivity\Set-FocusMode.ps1 -On -DryRun
```

```powershell
.\productivity\Set-FocusMode.ps1 -On -Minutes 50
```

| Option | Meaning |
| ------ | ------- |
| `-On` / `-Off` / `-Status` | Enter, leave, or report without changing anything |
| `-Minutes <n>` | Auto-restore after this long. Ctrl+C restores immediately |
| `-Force` | Terminate apps that ignore a polite close request |
| `-DryRun` | Show what would happen, touch nothing |
| `-ConfigPath <file>` | Use a different config |

**Config:** `%LOCALAPPDATA%\FocusMode\config.json`
**State:** `%LOCALAPPDATA%\FocusMode\focus-state.json`, plus hosts backups in
`hosts-backups\` (20 kept).

> If `-Off` finds notifications suppressed but no focus session recorded, it
> deliberately **leaves that setting alone** and tells you. Turning them on
> would be a guess, and guessing wrong silently un-mutes someone who wanted
> them off.

---

### New-ProjectScaffold.ps1

**`productivity/New-ProjectScaffold.ps1`**

Creates a project folder with a README, a type-appropriate `.gitignore`, an
`.editorconfig`, a `CLAUDE.md`, an initialised git repo on `main` with an
initial commit, and the starter files for your stack — then opens it. The
point is to remove the ten minutes of boilerplate between "I have an idea" and
"I am writing the actual code", without generating a framework's worth of files
you then have to read and delete.

**An existing file is never overwritten.** If the target folder already has
contents the script stops unless you pass `-Force`, and even then it only adds
files that are not already there. That holds for `-Type dotnet` too: `dotnet
new console` is run *without* `--force`, so a folder that already contains a
project is reported and left alone rather than replaced.

Types: `blank`, `node`, `python`, `powershell`, `static`, `dotnet`.

```powershell
.\productivity\New-ProjectScaffold.ps1 -Name my-thing -DryRun
```

```powershell
.\productivity\New-ProjectScaffold.ps1 -Name scraper -Type python -CreateVenv -License MIT
```

| Option | Meaning |
| ------ | ------- |
| `-Name <name>` | Project name. Becomes the folder and seeds package names |
| `-Type <type>` | `blank` (default), `node`, `python`, `powershell`, `static`, `dotnet` |
| `-Root <path>` | Where to create it. Default `Documents\projects` |
| `-Description <text>` | Used in the README and package metadata |
| `-License MIT` | Add a LICENSE file. Default `None` |
| `-Author <name>` | Defaults to `git config user.name`, then your Windows user |
| `-CreateVenv` | Python only: also create `.venv` |
| `-NoGit` / `-NoCommit` | Skip the repo, or skip only the initial commit |
| `-OpenWith code\|claude\|explorer\|none` | What to launch. Defaults to `code` if on PATH |
| `-Force` | Allow scaffolding into a non-empty folder |
| `-DryRun` | List what would be created, touch nothing |

**Exit code:** `0` the project folder was scaffolded, `1` nothing was created
(the folder already had contents and `-Force` was not given, the name is not
usable as a Windows folder, or `-Root` could not be written to). Optional steps
that fail — dotnet, git, venv — are reported as warnings and still exit `0`,
because the folder itself was scaffolded.

Pairs well with the Explorer context-menu entry in `shell/` — scaffold, then
right-click the new folder and open Claude Code in it.

---

## shell/

### Claude Code context menu

**`shell/install-claude-code-context-menu.reg`**
**`shell/uninstall-claude-code-context-menu.reg`**

Adds **"Open Claude Code here"** to two right-click menus: the empty background
of an open folder window, and a folder icon itself. It launches the `claude`
CLI resolved from `PATH`, in the folder you clicked.

Installs under `HKEY_CURRENT_USER`, so **no administrator rights are needed**
and it only affects your account.

To install, double-click `install-claude-code-context-menu.reg` and accept the
prompt. Or from a terminal:

```powershell
reg import .\shell\install-claude-code-context-menu.reg
```

To remove it, run the uninstall file the same way. Requires the `claude` CLI to
be on your `PATH`.

---

## Downloads janitor — `DownloadsJanitor/`

**Tidy My Downloads** sorts loose files in the Downloads folder into
`<Category>\<yyyy-MM>` folders, and reports — never deletes — old clutter.

Unlike the other tools here it is a **package, not a single script**, because it
is built to be zipped up and handed to someone who has never opened a terminal.

```
DownloadsJanitor/
├─ READ ME FIRST.txt              one page, plain English
├─ Tidy My Downloads.cmd          the thing you double-click
├─ Put Everything Back.cmd        the undo
├─ Remove Downloads Janitor.cmd   removes the weekly schedule only
└─ tool/
   ├─ DownloadsJanitor.Setup.ps1        3-step wizard, all console UI
   ├─ DownloadsJanitor.Core.ps1         the engine
   ├─ DownloadsJanitor.Categories.ps1   648-extension map, 15 categories
   ├─ DownloadsJanitor.Undo.ps1         replays the move journal backwards
   └─ DownloadsJanitor.Uninstall.ps1    unregisters the scheduled task
```

`.cmd` launchers rather than `.ps1` because a double-clicked `.ps1` opens in
Notepad; the `.cmd` shells out with `-ExecutionPolicy Bypass`, which is the only
scope that also clears Mark-of-the-Web on a file that arrived by email.

### Safety model

The whole design exists to make one promise credible: **it will never lose a
file.**

- Nothing is ever deleted. There is no purge flag and there will not be one.
- **Preview is the default.** Moving requires `-Apply`; the wizard always shows
  a preview first and defaults to *no* at every prompt.
- `Move-Item` appears in exactly one function, which has no `-Force` parameter
  and so is structurally incapable of overwriting a file.
- Every move is written to a journal **before** it happens. If the journal
  cannot be written, nothing moves.
- **30-day grace period.** Anything downloaded in the last month is never
  touched, so the Downloads folder itself is always "the last 30 days".
- File *contents* are never read — no hashing, no sniffing. That is what stops
  OneDrive hydrating gigabytes of cloud-only files from a background task.
- Folders are never moved. An extracted app breaks the moment its parent
  path changes.
- Refuses to run against a drive root, profile root, Desktop, Documents,
  `%APPDATA%` and friends.

### Finding things again

Sorting by month is great for tidiness and useless when you cannot remember
the month. So every category folder gets an **`_All <Category>`** shortcut.
Double-click it and Explorer lists every file in that category in one flat
list, all months together, with thumbnails and sorting. Nothing runs — Windows
Search builds the list on open, so it is always current.

**It is a `.lnk`, not a `.search-ms`, and that matters.** The obvious
implementation is a saved-search file, and it silently does not work: on
Windows 11 `.search-ms` is associated with the ProgID `SearchFolder`, but

```
> ftype SearchFolder
File type 'SearchFolder' not found or no open command associated with it.
```

There is no open command registered, so double-clicking a `.search-ms` does
nothing at all — no error, no window — and even `explorer.exe file.search-ms`
refuses it. The `search-ms:` *protocol* works fine, so the shortcut hands the
URI to `explorer.exe` instead. Verified by enumerating Explorer windows over
COM and counting the items in each result.

Three details in the query are load-bearing, each found by testing:

| Detail | Why |
| ------ | --- |
| The folder path is `Uri.EscapeDataString`-encoded | Six category names contain `&`, the URI parameter separator. Unencoded, Explorer opens **nothing at all**. |
| Clauses separated by `%20`, never `+` | With `+` the query silently returns the wrong set (1 item instead of 77). |
| `-System.ItemType:Directory -System.FileExtension:.lnk` | The first drops the `yyyy-MM` folders from the results; the second stops the shortcut appearing in its own listing. |

These shortcuts are never moved by the tool and never listed as clutter, and
undo removes them along with the folders it created.

### Usage

```powershell
# preview only - changes nothing
.\DownloadsJanitor\tool\DownloadsJanitor.Core.ps1

# actually sort
.\DownloadsJanitor\tool\DownloadsJanitor.Core.ps1 -Apply

# put it all back
.\DownloadsJanitor\tool\DownloadsJanitor.Undo.ps1
```

Reports and the undo journal live in `%LOCALAPPDATA%\DownloadsJanitor\`. A
plain-English `Where are my files.html`, with a search box over every moved
file, is written into the Downloads folder itself on each real run.

Registered as the weekly scheduled task **Tidy My Downloads** (Sundays around
midday, 15-minute random delay).
