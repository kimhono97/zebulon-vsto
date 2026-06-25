# Agent Instructions — ZebulonVSTO

A **VSTO (Visual Studio Tools for Office) PowerPoint COM add-in** written in C# on **.NET Framework 4.7.2**. Its purpose is **UDP-based slide synchronization**: one PowerPoint instance runs as a *SENDER* and broadcasts slide-navigation commands; one or more *RECEIVER* instances execute them. A WPF console window provides live traffic logging and manual command entry.

## Quick Reference

| | |
|---|---|
| Language / runtime | C# / .NET Framework **4.7.2** (legacy non-SDK csproj) |
| Project type | VSTO PowerPoint add-in (`OutputType=Library`, `OfficeApplication=PowerPoint`, `LoadBehavior=3`) |
| Solution / project | `ZebulonVSTO.sln` → `ZebulonVSTO/ZebulonVSTO.csproj` (single project) |
| Build tool | **Visual Studio / MSBuild** — *not* `dotnet` (VSTO + COM interop) |
| Verified IDE | Visual Studio 2022+ — confirmed on **VS 2026 Community v18.1** |
| Dependencies | **framework-only** — no NuGet packages; JSON via built-in `System.Runtime.Serialization` |
| Host | PowerPoint **2013+**; default UDP port **8291** |
| Tests | xUnit unit tests for pure logic in `tests/ZebulonVSTO.Tests` (run via `dotnet test`); COM runtime still verified manually via F5 |
| License | MIT |

## Build & Debug

### Prerequisites
- Visual Studio 2022 or later (verified on **VS 2026 Community v18.1**) with:
  - **Office/SharePoint development** workload, **and**
  - the **VSTO (Visual Studio Tools for Office)** individual component.
  - ⚠️ Without the VSTO component the `Microsoft.VisualStudio.Tools.Office.targets` import fails and the build dies — this is the most common build error.
- Target machines (for running, not building) need the **VSTO 2010 Runtime** (bootstrapped via `Microsoft.VSTORuntime.4.0`).

### In the IDE (preferred)
1. Open `ZebulonVSTO.sln` — NuGet restores automatically.
2. Build **Debug** or **Release** (`Any CPU`).
3. Press **F5** → PowerPoint launches with the add-in registered. **This is the only runtime-test path** (a COM add-in cannot be tested headlessly).

### From the command line (CI / quick check)
```bash
# Build — VisualStudioVersion=10.0 is REQUIRED, see gotcha below.
# No NuGet restore is needed: the add-in has no package dependencies.
msbuild ZebulonVSTO.sln -p:Configuration=Release -p:VisualStudioVersion=10.0
```

> **CLI gotcha — `VisualStudioVersion=10.0`:** the VSTO/OfficeTools build targets only exist under the legacy `v10.0` MSBuild path. The csproj defaults `VisualStudioVersion` to `10.0` *only when empty*, but command-line MSBuild pre-sets it to the modern installed version (e.g. `18.x`), so the Office targets aren't found unless you force `-p:VisualStudioVersion=10.0`. The IDE resolves this automatically; the CLI does not.

### Expected non-fatal warning
- **MSB3327** ("signing certificate not found in the user certificate store") is expected and harmless — it only affects ClickOnce manifest signing. The DLL still compiles. See *Deployment*.

### Outputs
Debug → `ZebulonVSTO/bin/Debug/`, Release → `ZebulonVSTO/bin/Release/`. Each produces `ZebulonVSTO.dll` plus the VSTO/ClickOnce deployment set: `ZebulonVSTO.dll.manifest`, `ZebulonVSTO.vsto`, `ZebulonVSTO.dll.config`, `ZebulonVSTO.pdb`.

### Runtime-testing the sync feature
A single F5 instance can't exercise SENDER↔RECEIVER traffic by itself — you need a peer. Simplest single-machine test:
1. F5 → in PowerPoint, **Add-Ins** tab → **Zebulon** → **Sync** group: set **Mode = Receiver**, **Local Port = 8291**, click **Start Sync**. Open a deck with ≥2 slides.
2. From a separate terminal, send a UDP datagram to `127.0.0.1:8291` carrying a REQUEST `SyncMessage` (PascalCase fields, `Type:1`):
   ```powershell
   $json = '{"SenderIP":"127.0.0.1","SenderPort":8291,"ID":1,"Type":1,"Data":"select 2"}'
   $b = [Text.Encoding]::UTF8.GetBytes($json); $u = New-Object Net.Sockets.UdpClient
   $u.Send($b, $b.Length, "127.0.0.1", 8291) | Out-Null; $u.Close()
   ```
   Swap `Data` for `alert <text>`, `select <n>`, `showslide <n>`, or `hideslide`. The RECEIVER executes the command and replies with a RESPONSE (`Success`/`Failed`); the **Console** button shows the traffic log.
   Helper scripts under `tests/manual/` mirror the add-in's sync roles, parameterized by a single `-Port` (pure PowerShell, no add-in DLL — they drive / stand in for any peer as-is). A SENDER binds an ephemeral local port automatically (so it never collides with `-Port`); a RECEIVER binds `-Port`:
   ```powershell
   # SENDER — emit one command and print the RESPONSE (exit 0 = Success, 1 = Failed/timeout)
   ./tests/manual/Send-SyncCommand.ps1 -Command "select 2" -Port 8291

   # SENDER REPL — type commands until quit/bye/exit; 'help' lists them
   ./tests/manual/Start-SyncSession.ps1 -Mode Sender -Port 8291

   # RECEIVER — listen on -Port, log datagrams, reply (stands in for a 2nd PowerPoint).
   # -Reply Success|Failed|None, -RunSeconds N; -Ip sets the bind interface / send target.
   ./tests/manual/Start-SyncSession.ps1 -Mode Receiver -Port 8291
   ```
   Same-machine note: the add-in binds its Local Port (8291) in BOTH modes, so a PS peer on the same box must use a different `-Port` than the running add-in (or run on another machine).
3. Alternatively, run a second PowerPoint instance in **Sender** mode and drive it from its Sync Console.

### Unit tests
The pure, COM-free logic (`CommandParser`, `SyncMessage` serialization) has xUnit coverage in `tests/ZebulonVSTO.Tests`:
```bash
dotnet test tests/ZebulonVSTO.Tests
```
- The test project is **SDK-style (`net472`) and deliberately NOT in `ZebulonVSTO.sln`**, so it builds with the modern `dotnet` toolchain independently of the VSTO/MSBuild build. It **link-compiles** the source files under test (`Sync/SyncDefaults.cs`, `Sync/CommandParser.cs`, `Sync/Message.cs`) rather than referencing the add-in DLL — so `dotnet test` never touches the VSTO toolchain.
- Anything requiring PowerPoint/COM (`SyncManager` networking, `ThisAddIn` slide actions) is **not** unit-tested — verify those via F5 as above.
- `Serialized_ContainsFrozenWireFieldName` guards the on-the-wire JSON field names; if it fails, peer-to-peer sync is broken — fix the code, not the test.

## Project Structure

```
ZebulonVSTO.sln
ZebulonVSTO/
  ThisAddIn.cs              ← entry point: lifecycle + PowerPoint event handlers; implements ISyncLogger + ISlideController
  ThisAddIn.Designer.cs     ← VSTO-generated plumbing (Globals, ribbon collection) — DO NOT hand-edit
  ThisAddIn.Designer.xml    ← VSTO host blueprint (generates ThisAddIn.Designer.cs) — DO NOT hand-edit
  MainRibbon.cs             ← ribbon callbacks ([ComVisible] IRibbonExtensibility)
  MainRibbon.xml            ← declarative ribbon UI (embedded resource)
  Sync/
    SyncManager.cs          ← UDP transport singleton (send/receive, modes); collaborators injected via Attach()
    Message.cs              ← SyncMessage wire DTO + MessageType enum (DataContractJsonSerializer)
    CommandParser.cs        ← pure command-string parser → ParsedCommand/CommandKind (unit-tested)
    SyncContracts.cs        ← ISyncLogger + ISlideController (host-implemented seams; decouple SyncManager from Globals)
    SyncDefaults.cs         ← shared default IP/port constants (COM-free)
    SyncConsole.xaml(.cs)   ← WPF debug console (log + custom command input)
  Properties/               ← AssemblyInfo, (empty) Settings & Resources scaffolding
  app.config                ← empty configuration (no binding redirects; see Dependencies)
  ZebulonVSTO_TemporaryKey.pfx ← manifest signing key (intentionally tracked — see Deployment)
tests/
  ZebulonVSTO.Tests/        ← SDK-style net472 xUnit project (not in the .sln; run via `dotnet test`)
```

### How the sync feature works
- **`SyncManager`** (`Sync/SyncManager.cs`) is a lazily-created singleton. Default IP/port constants live in `Sync/SyncDefaults.cs` (`Port=8291`, `Localhost=127.0.0.1`, `Broadcast=255.255.255.255`) so the pure types stay COM-free. Modes: `SENDER` (emits slide-navigation requests), `RECEIVER` (executes them), `NONE`. Mode/ports/remote IP are locked while sync is running. It holds **no** compile-time dependency on `Globals.ThisAddIn`: the host wires an `ISyncLogger` and `ISlideController` via `Attach()` at startup.
- **Wire format** (`Sync/Message.cs`): `SyncMessage` serialized as UTF-8 JSON via the framework **`DataContractJsonSerializer`** (no third-party package). Fields `SenderIP`, `SenderPort`, `ID`, `Type`, `Data` are pinned with `[DataMember(Name=...)]`. `MessageType` enum is `CUSTOM=0, REQUEST=1, RESPONSE=2`. **The field names are a frozen on-the-wire contract between instances — do NOT rename them** (member *order* in the JSON is irrelevant; parsing is by name). The `Serialized_ContainsFrozenWireFieldName` test guards this.
- **Commands** (RECEIVER side): the raw command string is parsed by the pure `CommandParser.Parse` → `ParsedCommand` (`alert <text>`, `select <n>`, `showslide <n>`, `hideslide`; verbs case-insensitive via `ToLowerInvariant`). `SyncManager.ProcessRequest` switches on `ParsedCommand.Kind` and dispatches to `ISlideController`. RECEIVER always replies with a RESPONSE (`Success`/`Failed`).
- **Threading**: a single **background** thread runs the blocking `UdpClient.Receive` loop. Shutdown is **cooperative** (no `Thread.Abort`): `StopSync` flips a `volatile bool`, calls `UdpClient.Close()` to unblock `Receive` (caught as `ObjectDisposedException`), then `Thread.Join`s. Message IDs use `Interlocked.Increment`. **Any PowerPoint Interop call triggered by a received message MUST be marshalled to the UI thread via `ThisAddIn`'s `Dispatcher`** (see `SelectSlide`/`ShowSlide`/`HideSlide`/`Alert`) — never call Interop directly from the receive thread.
- **Ribbon** attaches to the built-in Add-Ins tab (`idMso="TabAddIns"`) relabeled **"Zebulon"** with groups `Info` and `Sync`. `MainRibbon.xml` (control IDs + `onAction`/`getEnabled`/… callback names) and `MainRibbon.cs` (callbacks) **must stay in sync**; the public callback method names are bound by the XML and must not be renamed. Adding a control means editing both, plus the `_enableMap` logic.
- **`DebugMode`** (`ThisAddIn.cs`, currently `true`) gates whether the console may send custom (non-SENDER) commands; it is passed to `SyncManager.Attach`. Defaulting it to `false` is a deferred deployment task.

## Dependencies

- **No NuGet packages.** The add-in is framework-only; there is no `packages.config` and no `packages/` restore step. JSON (de)serialization uses the built-in **`System.Runtime.Serialization`** (`DataContractJsonSerializer`) — added as a plain framework `<Reference>` in the csproj.
  - *History:* the sync feature previously depended on `System.Text.Json 7.0.2` plus its 8 transitive packages. That set was flagged **NU1903** and has been removed entirely in favour of the framework serializer. If you reintroduce a serializer, prefer the framework one; don't pull `System.Text.Json` back in without a deliberate review.
- **Office interop** (`Office`, `Microsoft.Office.Interop.PowerPoint`, both v15.0.0.0) use `EmbedInteropTypes=true`, so interop types are embedded and no PIAs ship. Keep `Private=False`; do not enable Copy Local.
- **`app.config`** is now an empty `<configuration/>` — the previous `System.Runtime.CompilerServices.Unsafe` binding redirect existed only for the removed JSON chain. It is still copied to output as `ZebulonVSTO.dll.config`; edit `app.config`, never the bin copy.

## Deployment

- Building **Release** emits the ClickOnce/VSTO deploy set into `bin/Release/`: `ZebulonVSTO.dll`, `ZebulonVSTO.dll.manifest`, `ZebulonVSTO.vsto`.
- Manifests are Authenticode-signed (`SignManifests=true`) with the repo-tracked self-signed **`ZebulonVSTO_TemporaryKey.pfx`** (thumbprint `290E123AB16DDD9CF1C892FD8390BD530B40328E`). For production, replace it with a real code-signing certificate and update `ManifestCertificateThumbprint`.
- Installs require the **VSTO 2010 Runtime** on the target machine.

## Conventions & Cautions

- **Never commit build artifacts**: `bin/`, `obj/`, `packages/`, `*.user`, `.vs/`, and VS upgrade reports (`UpgradeLog*.htm`) are gitignored. If you open the solution in a newer VS it may regenerate `UpgradeLog.htm` — it is ignored; do not commit it.
- **Keep `ZebulonVSTO_TemporaryKey.pfx` tracked.** `.gitignore` has a deliberate `!` exception for it; do not remove that exception or delete the file.
- **Do not hand-edit generated files**: `*.Designer.cs` (incl. `Properties/Resources.Designer.cs`, `Properties/Settings.Designer.cs`), `ThisAddIn.Designer.xml`, and `obj/*.g.cs` — they are VSTO/WPF-generated.
- **Naming follows standard C# conventions** (the old Hungarian `p`/`n`/`str`/`b` prefixes were removed): `_camelCase` for private fields, `PascalCase` for methods/properties/types/constants, `camelCase` for locals & parameters. The two exceptions are intentional and **must not** be "modernized": the `SyncMessage` wire field names (`SenderIP`/`SenderPort`/`ID`/`Type`/`Data`) and the `MainRibbon` public callback names bound by `MainRibbon.xml`.
- Comments default to **English** for new code; VSTO-generated files still contain Korean comments — leave generated files alone.
- Build the add-in with **MSBuild/Visual Studio only**, never `dotnet build` (VSTO + COM). The `tests/` project is the opposite — it is SDK-style and built/run with **`dotnet test`** only (it is not in the add-in solution).
- Commit messages: **English** (consistent with history).

## Documentation Language Policy

Choose documentation language by **audience**:

- **Human-facing documents → Korean.** Read by the team/end users. Examples: `readme.md`, design/architecture notes, end-user docs.
- **Agent-facing documents → English.** Primarily consumed by AI coding agents. Examples: this `AGENTS.md`, the thin `CLAUDE.md` / `GEMINI.md` pointers, **agent memory files, and skills**.
- Mixed-audience single doc: prefer English with a brief Korean summary at the top if needed.
- **Code comments**: follow the existing convention in the file; default to English for new code.
- **Commit messages**: English.

Rationale: English agent-facing content improves token efficiency and agent comprehension; Korean human-facing content serves the team.

## Multi-Agent Discoverability

`CLAUDE.md` and `GEMINI.md` are intentionally thin pointers to this file (per the [AGENTS.md convention](https://agents.md)). Keep them thin — put guidance here, not there.

## Known Follow-Ups

- **Deployment hardening (not yet done):** `DebugMode` still defaults to `true` (allows console-issued custom commands); drive it from configuration and default to `false` for release. Note: `DebugMode` is currently snapshotted into `SyncManager` at `Attach()` startup — if it becomes runtime-configurable, propagate changes to the manager (re-`Attach` or add an `AllowCustomCommands` setter) rather than relying on the startup value. Replace the self-signed `ZebulonVSTO_TemporaryKey.pfx` with a real code-signing certificate.
- **CI:** no pipeline yet. A GitHub Actions Windows runner does not ship the VSTO build targets by default, so a cloud full-build needs investigation; `dotnet test tests/ZebulonVSTO.Tests` runs anywhere and is the easy first CI step.
- The previous `System.Text.Json` NU1903 advisory is **resolved** (dependency removed — see *Dependencies*).
