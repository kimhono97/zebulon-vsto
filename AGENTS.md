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
| NuGet | classic **packages.config** (9 packages → `packages/`), *not* PackageReference |
| Host | PowerPoint **2013+**; default UDP port **8291** |
| Tests | none — runtime verification is manual via F5 into PowerPoint |
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
# 1. Restore (packages.config requires this flag)
msbuild ZebulonVSTO.sln -t:restore -p:RestorePackagesConfig=true

# 2. Build — VisualStudioVersion=10.0 is REQUIRED, see gotcha below
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
3. Alternatively, run a second PowerPoint instance in **Sender** mode and drive it from its Sync Console.

## Project Structure

```
ZebulonVSTO.sln
ZebulonVSTO/
  ThisAddIn.cs              ← entry point: lifecycle + PowerPoint event handlers + slide actions
  ThisAddIn.Designer.cs     ← VSTO-generated plumbing (Globals, ribbon collection) — DO NOT hand-edit
  ThisAddIn.Designer.xml    ← VSTO host blueprint (generates ThisAddIn.Designer.cs) — DO NOT hand-edit
  MainRibbon.cs             ← ribbon callbacks ([ComVisible] IRibbonExtensibility)
  MainRibbon.xml            ← declarative ribbon UI (embedded resource)
  Sync/
    SyncManager.cs          ← UDP transport singleton (send/receive, modes, commands)
    Message.cs              ← SyncMessage JSON DTO + MessageType enum
    SyncConsole.xaml(.cs)   ← WPF debug console (log + custom command input)
  Properties/               ← AssemblyInfo, (empty) Settings & Resources scaffolding
  app.config                ← one binding redirect (see Dependencies)
  packages.config           ← NuGet (classic)
  ZebulonVSTO_TemporaryKey.pfx ← manifest signing key (intentionally tracked — see Deployment)
```

### How the sync feature works
- **`SyncManager`** (`Sync/SyncManager.cs`) is a lazily-created singleton. Key constants: `DEFAULT_PORT=8291`, `DEFAULT_LOCAL=127.0.0.1`, `DEFAULT_TARGET=255.255.255.255`. Modes: `SENDER` (emits slide-navigation requests), `RECEIVER` (executes them), `NONE`. Mode/ports/remote IP are locked while sync is running.
- **Wire format** (`Sync/Message.cs`): `SyncMessage` serialized via `System.Text.Json` with **PascalCase** fields `SenderIP`, `SenderPort`, `ID`, `Type`, `Data`. `MessageType` enum is `CUSTOM=0, REQUEST=1, RESPONSE=2`. **Renaming/recasing any field breaks the on-the-wire contract between instances — keep them stable.**
- **Commands** (RECEIVER side, `ProcessRequest`): `alert <text>`, `select <n>`, `showslide <n>`, `hideslide`. RECEIVER always replies with a RESPONSE (`Success`/`Failed`).
- **Threading**: a single background thread runs the blocking `UdpClient.Receive` loop, stopped via `Thread.Abort()`. **Any PowerPoint Interop call triggered by a received message MUST be marshalled to the UI thread via `ThisAddIn`'s `Dispatcher`** (see `DoSelectSlide`/`DoSlideShow`/`DoSlideShowEnd`) — never call Interop directly from the receive thread.
- **Ribbon** attaches to the built-in Add-Ins tab (`idMso="TabAddIns"`) relabeled **"Zebulon"** with groups `Info` and `Sync`. `MainRibbon.xml` (control IDs) and `MainRibbon.cs` (callbacks) **must stay in sync**; adding a control means editing both, plus the `pEnableMap` logic.
- **`DEBUG_MODE`** (`ThisAddIn.cs`, currently `true`) gates whether the console may send custom (non-SENDER) commands.

## Dependencies

- **Classic `packages.config`** restore (not PackageReference); 9 packages, all `net472`, restored into repo-root `packages/`.
- **`System.Text.Json` 7.0.2** is the only first-class dependency (JSON for the sync feature). The other 8 (`Microsoft.Bcl.AsyncInterfaces`, `System.Memory`/`Buffers`/`Numerics.Vectors`, `Runtime.CompilerServices.Unsafe`, `Text.Encodings.Web`, `ValueTuple`, `Threading.Tasks.Extensions`) are its transitive .NET Framework support set — **keep them version-aligned; don't bump one in isolation.**
- **Office interop** (`Office`, `Microsoft.Office.Interop.PowerPoint`, both v15.0.0.0) use `EmbedInteropTypes=true`, so interop types are embedded and no PIAs ship. Keep `Private=False`; do not enable Copy Local.
- **`app.config`** has one binding redirect: `System.Runtime.CompilerServices.Unsafe → 6.0.0.0`. Changing dependency versions may require updating it. It is copied to output as `ZebulonVSTO.dll.config` — edit `app.config`, never the bin copy.
- 🔒 **Known security advisory:** `System.Text.Json 7.0.2` is flagged **NU1903 (high severity, GHSA-hh2w-p6rv-4g7w)**. A bump was intentionally deferred. Do **not** silently upgrade it; if upgrading, do it deliberately and re-verify the binding redirect and that the whole transitive set still restores and builds.

## Deployment

- Building **Release** emits the ClickOnce/VSTO deploy set into `bin/Release/`: `ZebulonVSTO.dll`, `ZebulonVSTO.dll.manifest`, `ZebulonVSTO.vsto`.
- Manifests are Authenticode-signed (`SignManifests=true`) with the repo-tracked self-signed **`ZebulonVSTO_TemporaryKey.pfx`** (thumbprint `290E123AB16DDD9CF1C892FD8390BD530B40328E`). For production, replace it with a real code-signing certificate and update `ManifestCertificateThumbprint`.
- Installs require the **VSTO 2010 Runtime** on the target machine.

## Conventions & Cautions

- **Never commit build artifacts**: `bin/`, `obj/`, `packages/`, `*.user`, `.vs/`, and VS upgrade reports (`UpgradeLog*.htm`) are gitignored. If you open the solution in a newer VS it may regenerate `UpgradeLog.htm` — it is ignored; do not commit it.
- **Keep `ZebulonVSTO_TemporaryKey.pfx` tracked.** `.gitignore` has a deliberate `!` exception for it; do not remove that exception or delete the file.
- **Do not hand-edit generated files**: `*.Designer.cs` (incl. `Properties/Resources.Designer.cs`, `Properties/Settings.Designer.cs`), `ThisAddIn.Designer.xml`, and `obj/*.g.cs` — they are VSTO/WPF-generated.
- Existing code uses a Hungarian-ish prefix convention (`p`=object, `n`=int, `str`=string, `b`=bool) and mixes Korean (VSTO-generated) and English comments. Match the surrounding file.
- Build with **MSBuild/Visual Studio only**, never `dotnet build` (VSTO + COM).
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

- `System.Text.Json 7.0.2` high-severity advisory (NU1903 / GHSA-hh2w-p6rv-4g7w) — deferred upgrade (see *Dependencies*).
