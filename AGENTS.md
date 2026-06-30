# Agent Instructions — ZebulonVSTO

A **VSTO (Visual Studio Tools for Office) PowerPoint COM add-in** written in C# on **.NET Framework 4.7.2**. It has two features:

1. **UDP-based slide synchronization**: one PowerPoint instance runs as a *SENDER* and broadcasts slide-navigation commands; one or more *RECEIVER* instances execute them. A WPF console window provides live traffic logging and manual command entry.
2. **Slide generation (Zebulon direct insert)**: a WPF wizard that inserts **Praise (찬양)** and **Word/scripture (말씀)** slides straight into the active deck — lyrics from the Zebulon **Provider**, scripture from the Zebulon **Web Bible API** — by filling box-marker (`$1`..`$N`) bind layouts. It is an in-app port of the external `zebulon-exporter`.

Both features keep all PowerPoint/COM interop isolated behind host-implemented interfaces (`ISlideController`/`ISlideBuilder`); the rest is COM-free and unit-tested.

## Quick Reference

| | |
|---|---|
| Language / runtime | C# / .NET Framework **4.7.2** (legacy non-SDK csproj) |
| Project type | VSTO PowerPoint add-in (`OutputType=Library`, `OfficeApplication=PowerPoint`, `LoadBehavior=3`) |
| Solution / project | `ZebulonVSTO.sln` → `ZebulonVSTO/ZebulonVSTO.csproj` (single project) |
| Build tool | **Visual Studio / MSBuild** — *not* `dotnet` (VSTO + COM interop) |
| Verified IDE | Visual Studio 2022+ — confirmed on **VS 2026 Community v18.1** |
| Dependencies | **framework-only** — no NuGet packages; JSON via built-in `System.Runtime.Serialization` |
| Host | PowerPoint **2013+**; default UDP sync port **8291**, discovery port **8290** |
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
1. F5 → in PowerPoint, **Add-Ins** tab → **Zebulon** → **Sync** group → click **동기화 시작**. In the wizard pick **수신 (Receiver)**, keep **수신 포트 = 8291**, click **시작**. Open a deck with ≥2 slides. (To exercise the SENDER auto-discovery list, run a `Responder` stand-in — see below — then in a second instance choose **송신 → 자동으로 찾기**.)
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

   # RESPONDER — answer discovery pings on 8290 so the wizard's auto list finds a peer.
   # -Role Receiver|Sender|Idle (advertised), -Port = advertised sync port. SO_REUSEADDR set.
   ./tests/manual/Start-SyncSession.ps1 -Mode Responder -Port 8291
   ```
   Same-machine note: the add-in binds its Local Port (8291) in BOTH modes, so a PS peer on the same box must use a different `-Port` than the running add-in (or run on another machine).
3. Alternatively, run a second PowerPoint instance in **Sender** mode and drive it from its Sync Console.

### Unit tests
The pure, COM-free logic (`CommandParser`, `SyncMessage` serialization, `DiscoveryProtocol` payloads, `InputValidation`) has xUnit coverage in `tests/ZebulonVSTO.Tests`:
```bash
dotnet test tests/ZebulonVSTO.Tests
```
- The test project is **SDK-style (`net472`) and deliberately NOT in `ZebulonVSTO.sln`**, so it builds with the modern `dotnet` toolchain independently of the VSTO/MSBuild build. It **link-compiles** the COM-free source files under test rather than referencing the add-in DLL — so `dotnet test` never touches the VSTO toolchain. Linked sets: the sync pure layer (`Sync/SyncDefaults.cs`, `CommandParser.cs`, `Message.cs`, `DiscoveryProtocol.cs`, `InputValidation.cs`) and the slide-gen pure layer (`Slides/SlideGenDefaults.cs`, `SlideGenModels.cs`, `LayoutMatching.cs`, `PraisePlanner.cs`, `TextTransforms.cs`, `LyricModels.cs`, `DisplayNames.cs`, `BibleCatalog.cs`, `RubyText.cs`, `WordModels.cs`, `SectionUtils.cs`, `VerseSelection.cs`, `WordAssembler.cs`, `WordPlanner.cs`).
- **Adding a new pure (COM-free) file under test → register it in BOTH csproj**: the add-in's `<Compile Include="Slides\Foo.cs" />` *and* the test project's `<Compile Include="..\..\ZebulonVSTO\Slides\Foo.cs" Link="UnderTest\Foo.cs" />`. Miss the second and the file silently isn't tested; miss the first and the add-in won't build. Keep pure logic free of `Microsoft.Office.Interop`/WPF so it stays link-compilable.
- Anything requiring PowerPoint/COM (`SyncManager` networking, `ThisAddIn` slide actions) is **not** unit-tested — verify those via F5 as above.
- `Serialized_ContainsFrozenWireFieldName` guards the on-the-wire JSON field names; if it fails, peer-to-peer sync is broken — fix the code, not the test.

## Project Structure

```
ZebulonVSTO.sln
ZebulonVSTO/
  ThisAddIn.cs              ← entry point: lifecycle + PowerPoint event handlers; implements ISyncLogger + ISlideController
  ThisAddIn.Slides.cs       ← partial of ThisAddIn: the ISlideBuilder (slide-gen Interop) half — UI-thread-marshalled
  ThisAddIn.Designer.cs     ← VSTO-generated plumbing (Globals, ribbon collection) — DO NOT hand-edit
  ThisAddIn.Designer.xml    ← VSTO host blueprint (generates ThisAddIn.Designer.cs) — DO NOT hand-edit
  MainRibbon.cs             ← ribbon callbacks ([ComVisible] IRibbonExtensibility)
  MainRibbon.xml            ← declarative ribbon UI (embedded resource)
  Sync/
    SyncManager.cs          ← UDP transport singleton (send/receive, modes); collaborators injected via Attach()
    Message.cs              ← SyncMessage wire DTO + MessageType enum (DataContractJsonSerializer)
    CommandParser.cs        ← pure command-string parser → ParsedCommand/CommandKind (unit-tested)
    DiscoveryProtocol.cs    ← pure ZSYNC1 ping/announce payload parser → DiscoveryPayload/DiscoveryRole (unit-tested)
    DiscoveryResponder.cs   ← always-on UDP responder on 8290 (answers DISCOVER with ANNOUNCE)
    DiscoveryScanner.cs     ← broadcasts DISCOVER, collects ANNOUNCE → DiscoveredPeer (used by the wizard)
    InputValidation.cs      ← pure port/IPv4 validators (unit-tested; shared by the wizard)
    SyncContracts.cs        ← ISyncLogger + ISlideController + IStatusObserver (host-implemented seams)
    SyncDefaults.cs         ← shared default IP/port constants incl. DiscoveryPort=8290 (COM-free)
    SetupWizard.xaml(.cs)   ← WPF setup wizard (mode → manual/auto-discovery → start)
    SyncConsole.xaml(.cs)   ← WPF debug console (log + custom command input)
  Slides/                   ← slide-generation feature (Zebulon direct insert) — pure layer is COM-free + unit-tested
    SlideGenContracts.cs    ← ISlideBuilder seam (host-implemented by ThisAddIn.Slides.cs)
    SlideGenDefaults.cs     ← WebBaseUrl/ProviderBaseUrl, box counts, "$N " marker convention (COM-free)
    SlideGenModels.cs       ← COM-free DTOs: LayoutDescriptor/BoxSignature/LayoutMatch/LayoutSelection/SlidePlanItem
    LayoutMatching.cs       ← pure: scan layouts for $1..$N bind candidates + the empty/separator layout
    SlideGenWindow.xaml(.cs)← WPF wizard (type → layout → data → position → generate); template download
    ProviderClient.cs       ← HTTP client for the Zebulon Provider (lyric list/fetch, template list/download)
    LyricModels.cs          ← Praise lyric DTOs; TextTransforms.cs ← Praise box text rules; PraisePlanner.cs ← Praise plan
    BibleClient.cs          ← HTTP client for {WebBaseUrl}/api/bible → BibleData (Word)
    BibleCatalog.cs / LanguageCatalog.cs ← 66-book / 18-version / language metadata (COM-free)
    WordModels.cs / WordAssembler.cs / WordPlanner.cs ← Word passage assembly + slide plan (ports PPTXFile_Word.addItem)
    WordSelectWindow.xaml(.cs)← per-passage dialog: language/version/ruby slots + book/chapter + tap-to-select verses
    VerseSelection.cs       ← pure tap-to-range state machine (ports the web SlideEditorWord verse picker)
    SectionUtils.cs         ← pure verse-range coalesce/format; RubyText.cs ← pure ruby (furigana) text handling
    DisplayNames.cs         ← pure path→display-name helpers; LyricPreviewWindow.xaml(.cs) ← lyric preview dialog
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
- **Ribbon** attaches to the built-in Add-Ins tab (`idMso="TabAddIns"`) relabeled **"Zebulon"**, groups `GroupInfo`/`GroupSync` (Korean labels). It is intentionally minimal: a large `BtnSync` (stopped → opens the **setup wizard**; running → `StopSync`), `BtnConsole`, and a 3-row status strip (`LblState`/`LblLocal`/`LblRemote`) aligned to the large button's height. `LblLocal`/`LblRemote` use `getVisible` to appear only while running (minimal when stopped, self/peer IP·port·role while running). The sender's remote line is set once at start; the receiver's "last sender" line is refreshed via the `IStatusObserver` seam (`SyncManager` → `ThisAddIn.OnPeerChanged` → dispatcher → `MainRibbon.RefreshPeerStatus` → `InvalidateControl`), and only on a peer **change** (no per-datagram flicker). All connection settings (mode, ports, remote IP) now live in the wizard, **not** the ribbon — the old `DdMode`/`Eb*` edit boxes and `OnTextChange`/`_enableMap` are gone; validation moved to `Sync/InputValidation.cs`. `MainRibbon.xml` (control IDs + callback names) and `MainRibbon.cs` (callbacks) **must stay in sync**; the public callback method names are bound by the XML and must not be renamed.
- **Setup wizard** (`Sync/SetupWizard.xaml`) is a modal WPF window (owned by the PowerPoint HWND) launched from `BtnSync`. Flow is **2-screen asymmetric**: mode select → (RECEIVER: pick a port) or (SENDER: a `자동/수동` toggle — auto-discovery results list with a broadcast-to-all option + explicit port, or manual IP/port entry). On finish it writes settings onto `SyncManager` and calls `StartSync`; on a start failure it stays open. A `SyncManager.RemoteLabel` (display-only, **not** on the wire) carries a discovered peer's host name for the ribbon.
- **Discovery** (peer auto-detect) is **additive** to the proven sync path. An always-on `DiscoveryResponder` binds **8290** (`SO_REUSEADDR`) from startup and answers each `DISCOVER` with an `ANNOUNCE` (its sync IP/port + role) unicast to the datagram source; `DiscoveryScanner` broadcasts `DISCOVER` ×3 and collects replies for the wizard. Payloads are the pure `ZSYNC1;key=value` format (`DiscoveryProtocol.cs`); `MessageType` gained `DISCOVER=3`/`ANNOUNCE=4` (appended — never reorder). **Auto-discovery only works new-build-to-new-build on the same subnet** (old peers run no responder and use a different port); **manual entry is the universal fallback**. Limited broadcast (`255.255.255.255`) can be blocked by AP isolation/firewalls — the wizard falls back to manual on 0 results.
- **`DebugMode`** (`ThisAddIn.cs`) gates whether the console may send custom (non-SENDER) commands; it is passed to `SyncManager.Attach`. It is **build-gated via `#if DEBUG`** — `true` in Debug, **`false` in Release** (the shipped build) — so production cannot inject console commands.

### How the slide-generation feature works
The slide-gen feature is **additive** to sync and follows the same COM-isolation discipline. It inserts slides **directly into the active deck** (no file export — that is what it ports from the external Exporter).

- **Seam**: `ISlideBuilder` (`Slides/SlideGenContracts.cs`) is the only contract that touches PowerPoint; it is implemented by the **`ThisAddIn.Slides.cs`** partial, the feature's *only* Interop code. Everything under `Slides/` except the two windows and `ThisAddIn.Slides.cs` is **COM-free and unit-tested**. Every Interop call is marshalled to the add-in UI thread via `_dispatcher` (same rule as `ISlideController` — never call Interop off the UI thread).
- **Entry / flow**: a ribbon button opens `SlideGenWindow` (modal, owned by the PowerPoint HWND, mirrors `SetupWizard`). Steps: pick **type** (Praise / Word) → pick a **bind layout** → pick **data** → pick **insert position** → generate. The type step can also **download a template** from the Provider, save it (user-chosen path), open it, and bind the wizard to that exact deck — `ISlideBuilder.OpenPresentation` returns the deck's `FullName`, and `ExecutePlan` targets **by `FullName`**, not "whatever is active" (a download-opened deck may not be active under the modal).
- **Layout model (box markers)**: the target deck must carry template layouts whose placeholders' text starts with **box markers `"$1 "`..`"$N "`** (`SlideGenDefaults.BoxMarker`), plus an **empty** (placeholder-free) layout used as a separator — the same convention the Exporter scans for. `LayoutMatching` (pure) reads COM-free `LayoutDescriptor`s and returns bind candidates — each box captured as a `BoxSignature` (placeholder **height+top geometry**) — and the empty layout. **Praise = 3 boxes `[KR, EN, CN]`; Word = 4 boxes `[book, KR, EN, CN]`** (`SlideGenDefaults.PraiseBoxCount`/`WordBoxCount`).
- **Plan → execute**: a pure planner builds an ordered `List<SlidePlanItem>` (each = a `LayoutKind` + per-box text + optional speaker note). `ThisAddIn.ExecutePlan` adds the slides and, for bind slides, writes box text by **resolving each box to its placeholder by geometry** (`FindPlaceholder`, tolerant float compare — Interop reports points as floats). Box→placeholder mapping and any geometry tweak live in the **Interop layer**, never the pure planners.
- **Praise**: lyrics come from the Zebulon **Provider** (`ProviderClient` → `Lyric`/`LyricModels`). `PraisePlanner` ports `PPTXFile_Praise.addItem` 1:1: per item → empty separator, a title slide, one bind slide per page (KR/EN/CN via `TextTransforms`), trailing separator. A **Praise-only** CN-box vertical-centering nudge runs in `ApplyBoxText` — see the caution under *Conventions*.
- **Word (scripture)**: text comes from the Zebulon **Web Bible API** (`BibleClient` → `GET {WebBaseUrl}/api/bible?v=&b=&c=` → `BibleData`). `WordSelectWindow` builds one passage at a time: up to 3 **language/version/ruby** slots (`LanguageCatalog`/`BibleCatalog` — 18 versions, 66 books), a book/chapter picker, and a **tap-to-select verse picker** — tap a start verse, then a later verse to extend the range, then commit; multiple disjoint ranges are kept and merged via `SectionUtils.Coalesce`. The picker's transition logic is the pure `VerseSelection`. `RubyText` ports the web ruby handling (base / furigana / both). `WordAssembler` slices chapter verses by range into a `WordItem`; `WordPlanner` ports `PPTXFile_Word.addItem` 1:1 (box 0 = headText, boxes 1/2/3 = `"{verse}. {line}"`, one bind slide per verse line, no title/notes).
- **Reference sources**: the feature is a **faithful 1:1 port** of two external repos — the **Exporter** `zebulon-exporter/py/pptx_exporter.py` (PPTX plan + placeholder geometry) and **Zebulon Web** `jym-workbox/src/zebulon/...` (Word selection UI/data, ruby, coalesce). Keep ports 1:1; when behavior must differ, **document the deliberate divergence in an XML comment** (existing examples: `WordAssembler` plain-texts every line because a placeholder can't render HTML; `SectionUtils.Coalesce` *drops* inverted ranges to match the web's `filter(r => r && r.end >= r.start)`).
- **Config**: `SlideGenDefaults.WebBaseUrl` (`https://jym-workbox.vercel.app`) and `ProviderBaseUrl` carry `TODO(confirm)` markers — verify before shipping. Runtime (COM insert) is **F5-only**; the pure layer is unit-tested (`tests/.../SlideGenTests.cs`, `WordTests.cs`).

## Dependencies

- **No NuGet packages.** The add-in is framework-only; there is no `packages.config` and no `packages/` restore step. JSON (de)serialization uses the built-in **`System.Runtime.Serialization`** (`DataContractJsonSerializer`) — added as a plain framework `<Reference>` in the csproj.
  - *History:* the sync feature previously depended on `System.Text.Json 7.0.2` plus its 8 transitive packages. That set was flagged **NU1903** and has been removed entirely in favour of the framework serializer. If you reintroduce a serializer, prefer the framework one; don't pull `System.Text.Json` back in without a deliberate review.
- **Office interop** (`Office`, `Microsoft.Office.Interop.PowerPoint`, both v15.0.0.0) use `EmbedInteropTypes=true`, so interop types are embedded and no PIAs ship. Keep `Private=False`; do not enable Copy Local.
- **`app.config`** is now an empty `<configuration/>` — the previous `System.Runtime.CompilerServices.Unsafe` binding redirect existed only for the removed JSON chain. It is still copied to output as `ZebulonVSTO.dll.config`; edit `app.config`, never the bin copy.

## Deployment

Release packaging is a **local** step — run `Build-Release.ps1` (the VSTO add-in can't build on hosted CI). It builds Release, exports the signing cert's public `.cer`, and bundles the output + `deploy/` templates + the `Tools/` diagnostic scripts (copied from `tests/manual/`) into `dist/ZebulonVSTO-<version>.zip` (gitignored).

- Building **Release** emits the VSTO deploy set into `bin/Release/`: `ZebulonVSTO.dll`, `ZebulonVSTO.dll.manifest`, `ZebulonVSTO.vsto`, `ZebulonVSTO.dll.config`. `DebugMode` is `false` in Release.
- **Install method** (per-user, no admin): `deploy/Install.ps1` trusts the bundled self-signed `.cer` (CurrentUser Root + TrustedPublisher), copies the add-in to `%LOCALAPPDATA%\ZebulonVSTO`, and registers it under `HKCU\Software\Microsoft\Office\PowerPoint\Addins\ZebulonVSTO` (`Manifest=…\ZebulonVSTO.vsto|vstolocal`, `LoadBehavior=3`). `deploy/Uninstall.ps1` reverses it. Update = re-run `Install.ps1` over a newer extracted package. Targets need the **VSTO 2010 Runtime**.
- Manifests are Authenticode-signed (`SignManifests=true`) with the repo-tracked self-signed **`ZebulonVSTO_TemporaryKey.pfx`** (`CN=ZebulonVSTO`, thumbprint `BDEB7C77F1C2347FF7B935F540D8BF513648F3DC`, valid through 2031-06-25). For production/external distribution, replace it with a real code-signing certificate and update `ManifestCertificateThumbprint`; then only the exported `.cer` changes and the install flow is unchanged.
- The deploy bundle ships its own operator-facing docs — `deploy/README.txt` (Korean) and `deploy/AGENTS.md` (English) — distinct from this repo-level file.
- **Release notes**: follow `docs/RELEASE_NOTES_GUIDE.md` (Korean, end-user facing) when writing them — published on GitHub Releases per tag `v<version>`, with the `dist/ZebulonVSTO-<version>.zip` attached.

## Conventions & Cautions

- **Never commit build artifacts**: `bin/`, `obj/`, `packages/`, `*.user`, `.vs/`, and VS upgrade reports (`UpgradeLog*.htm`) are gitignored. If you open the solution in a newer VS it may regenerate `UpgradeLog.htm` — it is ignored; do not commit it.
- **Keep `ZebulonVSTO_TemporaryKey.pfx` tracked.** `.gitignore` has a deliberate `!` exception for it; do not remove that exception or delete the file.
- **Do not hand-edit generated files**: `*.Designer.cs` (incl. `Properties/Resources.Designer.cs`, `Properties/Settings.Designer.cs`), `ThisAddIn.Designer.xml`, and `obj/*.g.cs` — they are VSTO/WPF-generated.
- **Naming follows standard C# conventions** (the old Hungarian `p`/`n`/`str`/`b` prefixes were removed): `_camelCase` for private fields, `PascalCase` for methods/properties/types/constants, `camelCase` for locals & parameters. The two exceptions are intentional and **must not** be "modernized": the `SyncMessage` wire field names (`SenderIP`/`SenderPort`/`ID`/`Type`/`Data`) and the `MainRibbon` public callback names bound by `MainRibbon.xml`.
- Comments default to **English** for new code; VSTO-generated files still contain Korean comments — leave generated files alone.
- Build the add-in with **MSBuild/Visual Studio only**, never `dotnet build` (VSTO + COM). The `tests/` project is the opposite — it is SDK-style and built/run with **`dotnet test`** only (it is not in the add-in solution).
- **Slide-gen box-geometry tweaks are template-kind-specific.** The shared Interop path (`ExecutePlan`/`ApplyBoxText`) serves *both* Praise and Word, so any Exporter-faithful geometry adjustment must be gated to the kind the Exporter actually applies it to. The CN-box vertical-centering nudge exists only in `PPTXFile_Praise.addItem`, so it is gated behind `LayoutSelection.CenterCnBox` (set true only for Praise). **Don't gate such tweaks on box *count*** (Praise=3 / Word=4 is a coincidence, not the intent) — a past bug had the nudge leak into Word via `sigs.Count >= 3` and shift Word's 2nd-language box. Check the source `addItem` (`zebulon-exporter/py/pptx_exporter.py`) to see which kind a behavior belongs to before sharing it.
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

- **Code-signing certificate:** the tracked self-signed `ZebulonVSTO_TemporaryKey.pfx` was renewed (`CN=ZebulonVSTO`, thumbprint `BDEB…F3DC`, valid **2026-06-25 → 2031-06-25**, empty pfx password per VS temp-key convention); `deploy/Install.ps1` pre-trusts its `.cer` per-user. Two remaining nice-to-haves: (1) it carries **no RFC3161 timestamp**, so manifests signed with it stop validating once the cert expires in 2031 — add a trusted timestamp at signing time to avoid that; (2) for distribution beyond a trusted internal LAN, replace it with a real/internal-CA Authenticode cert (update `ManifestCertificateThumbprint`, re-export the `.cer`). To re-issue the self-signed cert again: `New-SelfSignedCertificate -Type CodeSigningCert -Subject "CN=ZebulonVSTO" -CertStoreLocation Cert:\CurrentUser\My -NotAfter (Get-Date).AddYears(5) -KeyExportPolicy Exportable`, then update the thumbprint + re-export the pfx/.cer.
- **`DebugMode` runtime config:** it is snapshotted into `SyncManager` at `Attach()`. Today it's build-gated (`#if DEBUG`), which is fine. If it ever becomes runtime-configurable, propagate changes to the manager (re-`Attach` or add an `AllowCustomCommands` setter) rather than relying on the startup value.

## Done (recent)

- **Slide generation (Zebulon direct insert)**: a WPF wizard (`Slides/`) that inserts **Praise (찬양)** and **Word/scripture (말씀)** slides into the active deck via `$1..$N` box-marker bind layouts — an in-app 1:1 port of `zebulon-exporter` (PPTX plan/geometry) and `jym-workbox` (Word selection UI/data). COM isolated behind `ISlideBuilder` (`ThisAddIn.Slides.cs`); the pure layer is unit-tested. Word adds a Bible-API client, language/version/ruby slots, and a tap-to-select verse picker. See *How the slide-generation feature works*. Runtime is F5-only.
- **Auto-discovery + setup wizard + ribbon redesign**: peer auto-detect (always-on `DiscoveryResponder` on 8290 / `DiscoveryScanner` / pure `ZSYNC1` `DiscoveryProtocol`), a modal WPF `SetupWizard` (manual entry kept as the universal fallback), and a minimal ribbon with a live self/peer status strip refreshed via the `IStatusObserver` seam. Full design spec: `docs/auto-discovery-spec.md`. Auto path is new-build-to-new-build on the same subnet; manual works cross-version.
- **CI** (`.github/workflows/ci.yml`): runs `dotnet test tests/ZebulonVSTO.Tests` on `windows-latest` for pushes to main/update and PRs to main. The VSTO add-in build stays local (hosted runners lack the VSTO targets).
- **Release packaging**: `Build-Release.ps1` + `deploy/` (see *Deployment*).
- **`DebugMode`**: build-gated off in Release.
- The previous `System.Text.Json` NU1903 advisory is **resolved** (dependency removed — see *Dependencies*).
