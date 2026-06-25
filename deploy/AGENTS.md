# Agent Instructions — ZebulonVSTO (deployed package)

You are looking at a **deployment bundle** of the ZebulonVSTO PowerPoint add-in,
not its source. Your job here is to help a user **install, run, diagnose, or
remove** it — not to build it. (Source: https://github.com/kimhono97/zebulon-vsto)

## What this is
A VSTO COM add-in for PowerPoint that does **UDP-based slide synchronization**:
one PowerPoint instance runs as **Sender** and broadcasts slide-navigation
commands; one or more **Receiver** instances execute them. Default UDP port
**8291**.

## Requirements (target machine)
- Windows, PowerPoint **2013+**
- **VSTO 2010 Runtime** ("Microsoft Visual Studio 2010 Tools for Office
  Runtime", microsoft.com/download id 48217). `Install.ps1` warns if missing.
- No admin rights needed — this is a **per-user (HKCU)** install.

## Package contents
```
ZebulonVSTO/        ZebulonVSTO.dll, .vsto, .dll.manifest, .dll.config, ZebulonVSTO.cer
Tools/              Send-SyncCommand.ps1, Start-SyncSession.ps1   (diagnostics; pure PowerShell)
Install.ps1         per-user install/update
Uninstall.ps1       per-user removal
README.txt          human-facing guide (Korean)
AGENTS.md           this file
```

## Install / update / uninstall
- **Install / update:** close PowerPoint, then
  `powershell -ExecutionPolicy Bypass -File .\Install.ps1`. It trusts the bundled
  self-signed cert (CurrentUser Root + TrustedPublisher — a Windows confirmation
  dialog may appear; the user clicks Yes), copies the add-in to
  `%LOCALAPPDATA%\ZebulonVSTO`, and registers it at
  `HKCU\Software\Microsoft\Office\PowerPoint\Addins\ZebulonVSTO`
  (`Manifest=<path>\ZebulonVSTO.vsto|vstolocal`, `LoadBehavior=3`). Updating is
  just re-running it over a freshly extracted newer package.
- **Uninstall:** close PowerPoint, then `... -File .\Uninstall.ps1` (removes the
  registry key, the install folder, and the trusted cert).

## Using the add-in
PowerPoint → **Add-Ins** ribbon tab → **"Zebulon"**. Pick **모드/Mode**
(송신/Sender or 수신/Receiver), set ports, **동기화 시작/Start**. The **콘솔/
Console** button opens a color-coded traffic log. Commands: `alert <text>`,
`select <n>`, `showslide <n>`, `hideslide`. The ribbon status line shows
`○ 중지됨` (stopped) / `● 수신·<port>` / `● 송신·→<ip>:<port>` / `⚠ <error>`.

## Diagnostics with Tools/
Pure-PowerShell UDP peers that mirror the add-in's roles (single `-Port`). They
work even though the shipped build has the debug console gated off, because a
Receiver always processes **incoming** REQUESTs.
```powershell
# Verify a deployed RECEIVER executes commands (start the add-in in Receiver mode first):
Tools\Send-SyncCommand.ps1 -Command "select 2" -Port 8291        # expect "Success", exit 0

# Watch what a deployed SENDER emits (set the add-in to Sender, Remote Port 8292):
Tools\Start-SyncSession.ps1 -Mode Receiver -Port 8292
```
**Same-machine caveat:** the add-in binds its Local Port (8291) in BOTH modes, so
a diagnostic peer on the same machine must use a **different** `-Port` than the
running add-in (or run on another machine).

## Troubleshooting
- **"Zebulon" tab missing:** fully restart PowerPoint. If still absent: confirm
  the VSTO 2010 Runtime is installed; check File → Options → Add-ins → Manage:
  COM Add-ins (it may have been disabled, or moved to the Disabled Items list).
- **"Unknown publisher" prompt:** the cert-trust step didn't complete — re-run
  `Install.ps1` and accept the certificate dialog.
- **Sync not working:** a firewall may be blocking the UDP port (default 8291);
  confirm the Sender's Remote IP/Port matches the Receiver's Local Port.
- **Console can't send custom commands:** expected — the shipped (Release) build
  disables debug command entry. Use a `Tools/` script as an external sender
  instead.

## Hard rules
- Do **not** try to build or `dotnet build` here — this is a binary bundle, not
  source. For code changes, work in the GitHub repo above.
- This add-in is .NET Framework / VSTO and Windows-only; there is no macOS/web
  variant.
