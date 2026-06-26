# Auto-Discovery + Setup Wizard + Ribbon Redesign — Implementation Spec

> **요약 (Korean summary)**
> 기존 **수동 설정**은 그대로 두고, **자동 탐색(discovery)** 방식을 추가한다. 시작 버튼은 **2화면 마법사(Wizard)** 를 띄운다(역할 선택 → 설정). 마법사는 비대칭 구조다: **수신**은 포트만 정해 대기/광고, **송신**은 화면 내 토글로 자동(브로드캐스트 ping으로 수신기 탐색→목록→선택) 또는 수동(직접 입력)을 고른다. 탐색은 본질적으로 **송신기가 수신기를 찾는 기능**이다. Ribbon은 단순화하여 정지 중엔 최소 표시, 작동 중엔 자신·상대 상태(IP/Port/역할)를 3행으로 보여준다. 수신측 "상대"는 직근 송신자 IP를 변화 시에만 갱신해 표시한다. **수동은 항상 universal fallback**이며, 자동 탐색은 양쪽 모두 신버전일 때만 동작한다.

This document is the single source of truth for the feature. It consolidates three design rounds (discovery protocol → wizard → ribbon). It is a *spec*, not code; section 10 gives the file-by-file implementation plan.

---

## 1. Goals & Non-Goals

**Goals**
- Add an **automatic** setup path (mode → discover peers → pick → auto-fill connection settings → start) alongside the existing **manual** path.
- Move all connection settings off the ribbon into a **WPF setup wizard** launched by the Start button.
- Redesign the ribbon: minimal when stopped, rich self/peer status when running.

**Non-Goals (deferred)**
- "Inviting" an IDLE peer to start (decision B — deferred).
- Cross-session persistence of last-used settings (`Properties.Settings`) — v2.
- TX/RX counters / last-response indicator on the ribbon — v2.
- Multi-NIC directed-broadcast enumeration — v2 (v1 uses limited broadcast).
- Reverse-DNS hostname for the receiver's "last sender" line — v2 (v1 shows IP).

**Hard constraints**
- **Manual is the universal fallback.** It must work with mixed/old peers, broadcast-blocked networks, and firewalls.
- **Auto-discovery requires both ends on the new build** (old builds have no responder; see §2.2).
- Do not touch the frozen wire contract field names (`SenderIP/SenderPort/ID/Type/Data`) or the proven sync send/receive path. Discovery is **additive**.

---

## 2. Discovery Protocol

### 2.1 Ports & topology
| Port | Role | Lifetime |
|---|---|---|
| **8291** (`SyncDefaults.Port`) | Sync traffic (existing) | Bound only while sync runs; locked while running |
| **8290** (new, `SyncDefaults.DiscoveryPort`) | Discovery ping/announce | **Always-on responder** from add-in startup to shutdown, independent of sync |

Rationale: discovery must answer *before* a peer has started sync, so it cannot live on the sync socket (chicken-and-egg). Two ports keep the always-on responder isolated from the start/stop-locked sync socket. (A future consolidation onto one always-on socket is possible but is a larger lifecycle refactor — out of scope.)

### 2.2 `MessageType` extension (append only)
```
CUSTOM = 0, REQUEST = 1, RESPONSE = 2     // existing — DO NOT reorder
DISCOVER = 3                              // broadcast ping
ANNOUNCE = 4                              // unicast reply
```
Reuse the existing `SyncMessage` DTO and `DataContractJsonSerializer` — no new transport type.

**Backward compatibility:** old builds run no responder, so they never emit `ANNOUNCE` and never appear in scans. They also never receive `DISCOVER` on 8290 (different port from sync 8291), so the existing `ReceiveLoop` is untouched. Net: auto-discovery silently finds only new-build peers; manual remains the cross-version path.

### 2.3 Field semantics
Identity travels in the message body, not the transport, so it survives ephemeral source ports / NAT.

**DISCOVER** (scanner → broadcast `255.255.255.255:8290`)
| Field | Value | Note |
|---|---|---|
| `Type` | 3 | |
| `SenderIP` | scanner's `LocalIP` | used for self-filtering (§2.8) |
| `SenderPort` | scanner's discovery source port | informational |
| `ID` | new scan session id (`Interlocked.Increment`) | echoed back for correlation |
| `Data` | `ZSYNC1;q=1;want=RECEIVER` | `want` optional filter |

**ANNOUNCE** (responder → unicast back to the DISCOVER datagram's **actual UDP source endpoint**)
| Field | Value | Note |
|---|---|---|
| `Type` | 4 | |
| `SenderIP` | responder's **sync** `LocalIP` | scanner copies → `RemoteIP` |
| `SenderPort` | responder's **sync** `LocalPort` | scanner copies → `RemotePort` |
| `ID` | echoed scan session id | scanner drops non-matching (stale) replies |
| `Data` | `ZSYNC1;role=RECEIVER;host=PC-A;ver=1.2.0` | role/host/version |

`SenderIP:SenderPort` is the single source of truth for "where to reach me for sync"; `Data` carries only non-address metadata (no duplicate port).

### 2.4 Payload format & parser
Compact, versioned, delimited string (keeps `LogMessage` output readable; avoids JSON-in-JSON escaping):
```
ZSYNC1;key=value;key=value
```
- First token MUST equal magic+version `ZSYNC1`; otherwise the payload is rejected (namespaces our traffic, ignores foreign/old datagrams).
- Unknown keys are ignored (forward-compatible); missing keys fall back to defaults.
- `host` = `Environment.MachineName`; `ver` = assembly `ProductVersion`.

`DiscoveryPayload` (pure, COM-free, **unit-tested** like `CommandParser`):
```
DiscoveryPayload {
  bool   Valid       // magic matched
  string Role        // "RECEIVER" | "SENDER" | "IDLE"   (ANNOUNCE)
  string Host        // display name                      (ANNOUNCE)
  string Version     // assembly ProductVersion           (ANNOUNCE)
  string Want        // optional role filter              (DISCOVER)
  bool   IsQuery     // q=1                                (DISCOVER)
}
static DiscoveryPayload Parse(string data)
static string BuildDiscover(string want)
static string BuildAnnounce(string role, string host, string version)
```
Place in `Sync/DiscoveryProtocol.cs` (alongside `SyncDefaults`). Add `DiscoveryPort = 8290` to `SyncDefaults`.

### 2.5 Always-on responder (`Sync/DiscoveryResponder.cs`)
- Owns a `UdpClient` bound to `8290` with `ExclusiveAddressUse = false` + `ReuseAddress` (multi-instance per host for testing).
- Background receive thread modeled on `SyncManager.ReceiveLoop`; **cooperative shutdown** (`volatile bool` + `Close()` → `ObjectDisposedException`, no `Thread.Abort`).
- On a valid `DISCOVER`: build `ANNOUNCE` and unicast it to the **datagram's actual source endpoint** (the `ref IPEndPoint` from `Receive`), not to the message's `SenderIP`.
- Responder is pure networking (no PowerPoint interop) → **no Dispatcher** needed.
- Reads live state from `SyncManager.GetInstance()` to fill the announce:

| SyncManager state | Announced `role` | Announced `SenderPort` |
|---|---|---|
| running as RECEIVER | `RECEIVER` | sync `LocalPort` (truly listening) |
| running as SENDER | `SENDER` | sync `LocalPort` |
| stopped (NONE) | `IDLE` | sync `LocalPort` (default 8291) |

- Lifecycle: started in `ThisAddIn_Startup`, stopped in `ThisAddIn_Shutdown`.
- **Self-skip:** if `DISCOVER.SenderIP == own LocalIP`, do not reply (avoids the scanner listing itself; see §2.8 caveat).

### 2.6 Scanner (`Sync/DiscoveryScanner.cs`, used by the wizard)
- Opens a temporary `UdpClient` on an ephemeral port with `EnableBroadcast = true`.
- Sends `DISCOVER` **3×, 300 ms apart** (UDP is lossy); collection window ≈ **1.8 s**.
- Collects `ANNOUNCE` whose `ID` matches the current scan session; **dedupes by `SenderIP:SenderPort`**.
- Client-side **self-filter**: drop announces where `SenderIP == own LocalIP`.
- Exposes results as they arrive via a callback/event; the wizard marshals UI updates with `Dispatcher.Invoke` (pattern from `SyncConsole.AppendLogLine`).
- Scan is cancelable (cooperative stop) when the wizard closes or re-scans.

Discovered-peer record surfaced to the wizard:
```
DiscoveredPeer { string Host; string IP; int SyncPort; string Role; string Version; }
```

### 2.7 Sequence (typical: receivers join first, presenter then connects)
```
[RECEIVER PC-A]            [SENDER PC-P]                [RECEIVER PC-B]
 responder@8290 (always)    responder@8290 (always)      responder@8290 (always)
 Wizard:수신→Start           Wizard:송신→자동
  listening sync@8291          |
                              | (1) DISCOVER → 255.255.255.255:8290  (×3)
  <--- ANNOUNCE -------------- |  --- ANNOUNCE --->  (reply to UDP source)
  IP=PC-A:8291 role=RECEIVER   |   IP=PC-B:8291 role=RECEIVER
                              | (2) collect 1.8s, dedupe, drop self
                              | (3) list: PC-A(:8291), PC-B(:8291)
                              | (4) user picks one / "all (broadcast)"
                              | (5) set RemoteIP/RemotePort(/RemoteLabel)
                              | (6) StartSync(SENDER)
                              |---- existing sync path from here ---->
```

### 2.8 Edge cases & known limits
- **Multi-NIC:** `FindLocalIPAddress` picks the first IPv4 only; v1 broadcasts to `255.255.255.255`. Multi-NIC directed broadcast via `NetworkInterface.GetAllNetworkInterfaces()` is v2.
- **Broadcast blocked** (AP client isolation / firewalls): scan returns 0 → wizard offers re-scan and "switch to manual" (manual always works).
- **Same-host multi-instance:** requires `ReuseAddress`; unicast `ANNOUNCE` may reach only one socket, and self-filter by `LocalIP` hides same-host peers. This is a **test-only caveat**; real multi-machine use (distinct IPs) is unaffected. The PowerShell stand-in (§9) covers same-host testing.
- **Security/disclosure:** the always-on responder discloses host/role/version on the LAN. The `ZSYNC1` magic limits responses to our protocol. Acceptable for internal LAN use; documented.

---

## 3. Setup Wizard

### 3.1 Window
- WPF window (same hosting pattern as `SyncConsole`), **modal** via `ShowDialog`, `Owner` set to the PowerPoint main window HWND for correct z-order/modality.
- Modal is safe: discovery runs on a background thread with Dispatcher marshaling, so the UI thread is never blocked by network I/O.
- `Cancel` / window-close performs no changes and stops any in-flight scan.

### 3.2 State machine
```
              open
               │
               ▼
        ┌─────────────┐
        │  ModeSelect │
        └──┬───────┬──┘
      송신 │       │ 수신
           ▼       ▼
  ┌──────────────┐  ┌──────────────┐
  │  SenderSetup │  │ ReceiverSetup│
  │ (자동/수동토글)│  │  (포트 입력)  │
  └──┬────────┬──┘  └──────┬───────┘
 자동│        │수동        시작│
     ▼        ▼               ▼
 Scanning   ManualForm    Apply+StartSync(RECEIVER) → close
  │   │        │ 시작
  │   ▼        ▼ validate
  │ Results   Apply+StartSync(SENDER) → close
  │   │ select(1台/全員)+시작
  │   ▼
  │  Apply+StartSync(SENDER) → close
  ▼
 Empty → re-scan / switch-to-manual

 (any) Cancel → close, no change
```

### 3.3 Screen 1 — Mode select
Two large cards: **송신 (발표자)** / **수신 (청중)** with one-line descriptions. Card click selects and may auto-advance; `[다음 >]` and `[취소]`. Back-navigable.

### 3.4 Screen 2A — Receiver setup
| Control | Behavior |
|---|---|
| 로컬 IP (read-only) | `SyncMng.LocalIP` (auto-detected) |
| 수신 포트 | editable, default `SyncMng.LocalPort` (8291), validated 1–65535 |
| (optional) 주변 송신기 | passive count from responder; v1 may omit |
| `[< 뒤로] [취소] [시작]` | `시작` enabled only when port valid |

Apply: `Mode=RECEIVER; LocalPort=port; StartSync()`.

### 3.5 Screen 2B — Sender setup (auto/manual toggle in one screen)
Top segmented toggle: **(● 자동으로 찾기) (○ 직접 입력)**. Default = 자동.

**Auto sub-states**
- *Scanning*: list area shows spinner "검색 중… (약 1.8초)"; scan auto-starts on entry.
- *Results*: selectable rows + a top broadcast row:
  - `◉ 모든 수신기로 전송 (브로드캐스트 : [8291])` — **explicit port field** (decision: broadcast uses one explicit port; receivers on other ports are excluded, noted in-UI).
  - `○ {Host} {IP} : {SyncPort}  🟢수신` (selectable)
  - `⊘ {Host} {IP}  ⚪대기` (IDLE — shown greyed, not selectable)
  - `[↻ 다시 검색]`
- *Empty*: "수신기를 찾지 못했습니다" + `[↻ 다시 검색]` + `[직접 입력으로 전환]`.
- Advanced (collapsed): 로컬 포트(응답 수신), default 8291.

**Manual sub-state**
| Control | Validation |
|---|---|
| 원격 IP | IPv4 only (`IPAddress.TryParse` + `AddressFamily.InterNetwork`) |
| 원격 포트 | 1–65535 |
| 로컬 포트 (advanced) | 1–65535, default 8291 |

**Apply (sender):**
| Path | RemoteIP | RemotePort | RemoteLabel (display-only) |
|---|---|---|---|
| auto, unicast | peer.IP | peer.SyncPort | peer.Host |
| auto, broadcast | `255.255.255.255` | explicit port | (none → ribbon shows "전체(브로드캐스트)") |
| manual | typed IP | typed port | (none → ribbon shows IP) |
Then `Mode=SENDER; LocalPort=selfLocalPort; StartSync()`.

### 3.6 Validation (shared pure helper)
Extract the current ribbon validation (`TryParsePort`, IPv4 check from `MainRibbon.OnTextChange`) into a pure, testable helper (e.g. `Sync/InputValidation.cs`). Wizard forms disable `시작` and show inline red text on invalid input.

### 3.7 Start failure handling
`StartSync()` returns `false` (e.g. port in use). The wizard must **stay open** and show inline error "포트 사용 중일 수 있음" (mirrors current ribbon `⚠ 시작 실패`), not close.

---

## 4. Ribbon Redesign

### 4.1 Layout principle — Large = 3 × Normal height
The large Start/Stop button is 3 normal-rows tall. Put a **3-row status strip** beside it so heights align exactly.
```
       Col1 (Large)     Col2 (Tools)    Col3 (Status = 3 rows)
row1                     [콘솔]           LblState
row2   [BtnSync large]                    LblLocal
row3                                       LblRemote
```
Col2's lower two rows are intentionally empty (reserved for v2 TX/RX counters).

### 4.2 Two visual states
- **Stopped (minimal):** `LblState` = "○ 중지됨"; `LblLocal`/`LblRemote` hidden via `getVisible=false`. `BtnConsole` disabled.
- **Running (rich):** all three rows visible; `BtnConsole` enabled.

### 4.3 Status content per mode
| Row (control) | Stopped | Sender running | Receiver running |
|---|---|---|---|
| `LblState` | `○ 중지됨` | `● 송신 중` | `● 수신 중` |
| `LblLocal` | (hidden) | `로컬 {LocalIP} : {LocalPort}` | `로컬 {LocalIP} : {LocalPort}` |
| `LblRemote` | (hidden) | `원격 → {target}` | `송신자 {lastPeer}` |

- Sender `{target}`: broadcast → `전체(브로드캐스트) : {port}`; unicast with `RemoteLabel` → `{Host}({IP}) : {port}`; else `{IP} : {port}`.
- Receiver `{lastPeer}`: before first message → `대기 중…`; after → `{senderIP}`  *(IP only in v1; reverse-DNS hostname deferred — sync messages carry no host name)*.

### 4.4 Control & callback map (new ribbon XML)
| Control id | Type/size | Callbacks | Notes |
|---|---|---|---|
| `BtnAbout` | button large | `onAction=OnBtnAction` | unchanged |
| `BtnSync` | button large | `getImage`, `getLabel`, `getEnabled`, `onAction` | stopped→**open wizard**; running→**StopSync** |
| `BtnConsole` | button normal | `getEnabled`, `onAction` | enabled only while running |
| `LblState` | labelControl | `getLabel` | always visible |
| `LblLocal` | labelControl | `getLabel`, `getVisible` | visible only while running |
| `LblRemote` | labelControl | `getLabel`, `getVisible` | visible only while running |

**Removed** controls/callbacks: `DdMode`, `EbLocalIP`, `EbLocalPort`, `EbRemoteIP`, `EbRemotePort`, and their `OnTextChange` / `GetText` / `GetSelectedItemIndex` / `OnDdAction` handlers. Validation moves to the wizard (§3.6). Remember the rule: `MainRibbon.xml` ids and `MainRibbon.cs` public callback names must stay in lockstep.

### 4.5 `BtnSync` action change
```
OnBtnAction("BtnSync"):
  if SyncMng.IsRunning():  HideSyncConsole(); SyncMng.StopSync()
  else:                    open SetupWizard (ShowDialog)   // wizard calls StartSync on Finish
  then UpdateSyncSettingsUI()
```

---

## 5. Live Status Seam (the only non-trivial new wiring)

Ribbon labels are pull-based (`getLabel`), so updating them while running requires an explicit `Invalidate` trigger.

| Info | Update timing | Mechanism |
|---|---|---|
| state / local / **sender's target** | once, at start | existing `_ribbon.Invalidate()` after `StartSync` (target is static) |
| **receiver's "last sender"** | only when the sender identity changes | new `IStatusObserver` seam |

### 5.1 `IStatusObserver` (in `Sync/SyncContracts.cs`)
```
public interface IStatusObserver { void OnPeerChanged(); }
```
Wire it like the existing collaborators: extend `SyncManager.Attach(logger, controller, statusObserver, allowCustomCommands)` and update the `ThisAddIn_Startup` call.

### 5.2 `SyncManager` changes
- Add `_lastPeerIP`, `_lastPeerPort` + public `LastPeerIP`/`LastPeerPort`.
- In `ReceiveLoop`, on the RECEIVER `REQ` branch: if `received.SenderIP/SenderPort` differs from the stored pair, update it and call `_statusObserver?.OnPeerChanged()`. **Change-only → no per-packet invalidation / no flicker.**

### 5.3 Notification path & threading
```
ReceiveLoop (bg thread)
  └ _statusObserver.OnPeerChanged()
        → ThisAddIn.OnPeerChanged()                       // implements IStatusObserver
              → _dispatcher.BeginInvoke(() =>             // marshal to UI thread
                    _mainRibbon?.RefreshPeerStatus())     // ThisAddIn holds the MainRibbon ref
                       → _ribbon?.InvalidateControl("LblRemote")
```
- `ThisAddIn` must keep the `MainRibbon` instance it creates in `CreateRibbonExtensibilityObject` (currently discarded) in a field.
- `MainRibbon.RefreshPeerStatus()` invalidates `LblRemote` (and `LblState` if needed).
- Reuse the existing `_dispatcher` (already used for slide ops) — receive thread must never call ribbon APIs directly.

### 5.4 Display-only remote label
Add a display-only `RemoteLabel` (string) the wizard sets at start for the sender's unicast hostname. **Not on the wire**; purely for `LblRemote` rendering. Broadcast/manual leave it null (ribbon derives the text).

---

## 6. Data-flow summary (who sets what)
| Value | Set by | Read by |
|---|---|---|
| Mode, LocalPort, RemoteIP, RemotePort | Wizard apply | SyncManager (transport) + ribbon labels |
| RemoteLabel (display) | Wizard apply (sender unicast) | `LblRemote` |
| LastPeerIP/Port | `SyncManager.ReceiveLoop` (receiver) | `LblRemote` via seam |
| role/host/ver (announce) | DiscoveryResponder (reads SyncManager) | remote scanners |
| DiscoveredPeer list | DiscoveryScanner | Wizard auto results |

---

## 7. Testing
- **Unit (`tests/ZebulonVSTO.Tests`, `dotnet test`):** `DiscoveryProtocol.Parse`/`Build` round-trips, magic mismatch rejection, missing/unknown keys, `want` filtering. Keep `SyncMessage` frozen-field test green (new `Type` values must serialize/parse).
- **Manual PS stand-in:** extend `tests/manual/Start-SyncSession.ps1` with a **responder mode** (listen on 8290, reply `ANNOUNCE` to `DISCOVER`) so a single F5 instance can populate the wizard's auto list.
- **F5 scenarios:** (a) receiver Start → another instance/PS sender auto-discovers and connects; (b) broadcast-to-all with 2 receivers; (c) 0-result → manual fallback; (d) receiver "last sender" line updates once when traffic starts; (e) start failure (port in use) keeps wizard open.

---

## 8. Backward compatibility & migration
- Wire field names unchanged; `MessageType` only appended. Old peers ignore discovery (different port, no responder) and remain reachable via **manual**.
- Auto-discovery is new-build-to-new-build only; document this in `deploy/README.txt` (Korean) operator notes.

---

## 9. Open items / deferred
Invite-IDLE (B), cross-session persistence, ribbon TX/RX counters, multi-NIC directed broadcast, receiver reverse-DNS hostname. None block v1.

---

## 10. Implementation plan (file-by-file, suggested order)
1. **Protocol core (testable first):** `Sync/SyncDefaults.cs` (+`DiscoveryPort=8290`), `Sync/DiscoveryProtocol.cs` (`DiscoveryPayload` + parse/build), `Sync/InputValidation.cs` (extract validators). Tests in `tests/ZebulonVSTO.Tests`.
2. **Responder:** `Sync/DiscoveryResponder.cs`; start/stop in `ThisAddIn.cs`.
3. **Scanner:** `Sync/DiscoveryScanner.cs` (+ `DiscoveredPeer`).
4. **Wizard:** `Sync/SetupWizard.xaml(.cs)` (state machine, screens, validation, apply, start-failure handling).
5. **Seam + live status:** `IStatusObserver` in `Sync/SyncContracts.cs`; `SyncManager` last-peer tracking + `Attach` signature; `ThisAddIn` observer impl + held `MainRibbon` ref + `RemoteLabel`.
6. **Ribbon:** rewrite `MainRibbon.xml` (new controls, remove edit boxes/dropdown) and `MainRibbon.cs` (`getVisible`, new labels, `BtnSync`→wizard, drop removed callbacks, `RefreshPeerStatus`). Keep XML ids ↔ callback names in lockstep.
7. **Docs/tests:** update `AGENTS.md` (ribbon/discovery sections), `deploy/README.txt`; add PS responder mode.

> Build reminder: VSTO add-in builds with **MSBuild/VS only** (`-p:VisualStudioVersion=10.0` on CLI); the `tests/` project is `dotnet test` only.
