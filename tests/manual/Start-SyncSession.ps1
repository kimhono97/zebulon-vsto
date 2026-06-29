<#
.SYNOPSIS
    A PowerShell mirror of the ZebulonVSTO UDP sync feature. Runs as a SENDER
    (emits commands to a port) or a RECEIVER (listens on a port and replies),
    without needing a second PowerPoint. Pure PowerShell — no add-in DLL.

.DESCRIPTION
    Mirrors the add-in's roles, parameterized by a single -Port:

      -Mode Sender    -Port N   ->  emit commands to <Ip>:N  (REPL prompt)
      -Mode Receiver  -Port N   ->  listen on <Ip>:N, log datagrams, reply
      -Mode Responder           ->  answer peer-discovery pings (auto-detect)
      -Mode Probe               ->  broadcast discovery pings, list responders

    SENDER binds an ephemeral local port automatically (so it never collides with
    a receiver) and reads RESPONSEs there. Session commands: help, info, verbose,
    clear, quit/bye/exit.

    RECEIVER binds <Ip>:<Port> and (unless -Reply None) answers each REQUEST.
    Stop with q/Esc, Ctrl+C, or -RunSeconds.

    RESPONDER stands in for the add-in's always-on discovery responder: it listens
    on -DiscoveryPort (8290) and answers each DISCOVER with an ANNOUNCE advertising
    -Role (default Receiver) at sync port -Port. This lets a single F5 add-in
    instance populate the setup wizard's auto-discovery list. It sets SO_REUSEADDR,
    so it can share 8290 with a running add-in on the same host.

    PROBE is the mirror of Responder: it broadcasts a DISCOVER on -DiscoveryPort
    (8290) and prints every ANNOUNCE it receives (host / sync IP:port / role /
    version). Use it to prove that a running add-in's discovery responder (its
    "self-announce") actually answers — something the add-in cannot show by
    scanning, since it never lists itself.

    INTERACTIVE PROMPTS: run with no arguments and the script asks for Mode, then
    Ip and Port (press Enter to accept the shown default). Only the parameters you
    omit are prompted; pass any of -Mode/-Ip/-Port to skip its prompt. A redirected
    / non-interactive session never prompts — it uses the defaults — so automation
    does not block.

.PARAMETER Mode
    Sender (default), Receiver, Responder, or Probe.

.PARAMETER Port
    SENDER: the destination it emits to. RECEIVER: the port it listens on.
    RESPONDER: the sync port it advertises in its ANNOUNCE. Default 8291.

.PARAMETER Ip
    SENDER: destination IP (default 127.0.0.1; use 255.255.255.255 to broadcast).
    RECEIVER/RESPONDER: local interface to bind (default 0.0.0.0 = all interfaces).

.PARAMETER Reply
    RECEIVER only: answer REQUESTs with Success (default), Failed, or None.

.PARAMETER RunSeconds
    RECEIVER/RESPONDER: auto-stop after N seconds (0 = run until q/Esc/Ctrl+C).

.PARAMETER TimeoutMs
    SENDER only: how long to wait for each RESPONSE. Default 1000.

.PARAMETER Role
    RESPONDER only: the role to advertise — Receiver (default), Sender, or Idle.

.PARAMETER DiscoveryPort
    RESPONDER only: the discovery port to listen on. Default 8290.

.EXAMPLE
    .\Start-SyncSession.ps1 -Mode Sender -Port 8291
    zebulon> select 2

.EXAMPLE
    .\Start-SyncSession.ps1 -Mode Receiver -Port 8291

.EXAMPLE
    # Appear as a discoverable RECEIVER (sync port 8291) for the wizard's auto list
    .\Start-SyncSession.ps1 -Mode Responder -Port 8291

.EXAMPLE
    # Verify a running add-in's discovery responder answers (its self-announce)
    .\Start-SyncSession.ps1 -Mode Probe
#>
[CmdletBinding()]
param(
    [ValidateSet('Sender', 'Receiver', 'Responder', 'Probe')] [string] $Mode = 'Sender',
    [int]    $Port          = 8291,
    [string] $Ip            = '',
    [ValidateSet('Success', 'Failed', 'None')] [string] $Reply = 'Success',
    [int]    $RunSeconds    = 0,
    [int]    $TimeoutMs     = 1000,
    [ValidateSet('Receiver', 'Sender', 'Idle')] [string] $Role = 'Receiver',
    [int]    $DiscoveryPort = 8290
)

$ErrorActionPreference = 'Stop'

function Get-LocalIPv4 {
    foreach ($a in [System.Net.Dns]::GetHostAddresses([System.Net.Dns]::GetHostName())) {
        if ($a.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork) {
            return $a.IPAddressToString
        }
    }
    return '127.0.0.1'
}

function Resolve-SenderIp([string] $targetIp, [int] $targetPort) {
    # Choose the source IP that actually routes to the target (correct on
    # multi-homed hosts with Hyper-V/WSL/VPN NICs), not just the first DNS IPv4.
    # A connected UDP socket sends nothing but lets the OS pick the egress NIC.
    if ($targetIp -in @('127.0.0.1', 'localhost')) { return '127.0.0.1' }
    try {
        $probe = New-Object System.Net.Sockets.UdpClient
        $probe.Connect($targetIp, $targetPort)
        $ip = $probe.Client.LocalEndPoint.Address.ToString()
        $probe.Close()
        if ($ip -and $ip -ne '0.0.0.0') { return $ip }
    } catch { }
    return Get-LocalIPv4
}

function ConvertTo-JsonStringLiteral([string] $s) {
    if ($null -eq $s) { return '' }
    return $s.Replace('\', '\\').Replace('"', '\"')
}

function New-ListenerSocket([string] $ip, [int] $port) {
    try {
        if ([string]::IsNullOrEmpty($ip) -or $ip -eq '0.0.0.0') {
            return New-Object System.Net.Sockets.UdpClient($port)   # all interfaces
        }
        $endpoint = New-Object System.Net.IPEndPoint([System.Net.IPAddress]::Parse($ip), $port)
        return New-Object System.Net.Sockets.UdpClient($endpoint)
    } catch {
        Write-Host ("ERROR: cannot bind UDP {0}:{1} - {2}" -f $ip, $port, $_.Exception.GetBaseException().Message) -ForegroundColor Red
        Write-Host "That port is already in use. If your PowerPoint add-in has sync running it holds this port (it binds its Local Port in BOTH Sender and Receiver mode). Stop its sync / close PowerPoint, or pick a different -Port." -ForegroundColor Yellow
        return $null
    }
}

function Show-SenderHelp {
    Write-Host ''
    Write-Host 'Message commands (emitted to the receiver):' -ForegroundColor Yellow
    Write-Host '  select <n>      Select slide n in the editor window'
    Write-Host '  showslide <n>   Start/seek the slide show to slide n'
    Write-Host '  hideslide       End the slide show'
    Write-Host '  alert <text>    Pop an alert dialog with <text>'
    Write-Host '  (any other text is emitted verbatim as a command)'
    Write-Host ''
    Write-Host 'Session commands (handled locally, not emitted):' -ForegroundColor Yellow
    Write-Host '  help, ?         Show this help'
    Write-Host '  info            Show current target / local port / next id'
    Write-Host '  verbose         Toggle raw JSON echo on/off'
    Write-Host '  clear, cls      Clear the screen'
    Write-Host '  quit, bye, exit Leave the session'
    Write-Host ''
}

function Invoke-SenderMode {
    $targetIp = if ([string]::IsNullOrEmpty($Ip)) { '127.0.0.1' } else { $Ip }

    try {
        $udp = New-Object System.Net.Sockets.UdpClient(0)   # ephemeral local port
    } catch {
        Write-Host ("ERROR: cannot open a local UDP socket - {0}" -f $_.Exception.GetBaseException().Message) -ForegroundColor Red
        return
    }
    if ($targetIp -eq '255.255.255.255') { $udp.EnableBroadcast = $true }
    $udp.Client.ReceiveTimeout = $TimeoutMs
    $localPort = $udp.Client.LocalEndPoint.Port
    $myIp = Resolve-SenderIp $targetIp $Port
    $remote = New-Object System.Net.IPEndPoint([System.Net.IPAddress]::Any, 0)
    $nextId = 1
    $verbose = $false

    Write-Host ("ZebulonVSTO sync mirror [SENDER] -> emitting to {0}:{1}  (local :{2})" -f $targetIp, $Port, $localPort) -ForegroundColor Green
    Write-Host "Type 'help' for commands, 'quit' to leave."
    Show-SenderHelp

    try {
        while ($true) {
            Write-Host -NoNewline 'zebulon> ' -ForegroundColor Cyan
            $line = [Console]::ReadLine()
            if ($null -eq $line) { break }
            $line = $line.Trim()
            if ($line -eq '') { continue }

            $verb = ($line -split '\s+', 2)[0].ToLowerInvariant()

            if ($verb -in @('quit', 'bye', 'exit', 'q')) {
                break
            } elseif ($verb -in @('help', '?')) {
                Show-SenderHelp
            } elseif ($verb -eq 'info') {
                Write-Host ("  target={0}:{1}  localPort={2}  senderIp={3}  nextId={4}  verbose={5}  timeoutMs={6}" -f `
                    $targetIp, $Port, $localPort, $myIp, $nextId, $verbose, $TimeoutMs)
            } elseif ($verb -in @('verbose', 'v')) {
                $verbose = -not $verbose
                Write-Host ("  verbose = {0}" -f $verbose)
            } elseif ($verb -in @('clear', 'cls')) {
                Clear-Host
            } else {
                $id = $nextId
                $nextId++
                $data = ConvertTo-JsonStringLiteral $line
                $json = '{{"SenderIP":"{0}","SenderPort":{1},"ID":{2},"Type":1,"Data":"{3}"}}' -f $myIp, $localPort, $id, $data
                $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
                [void] $udp.Send($bytes, $bytes.Length, $targetIp, $Port)
                if ($verbose) { Write-Host ("  -> {0}" -f $json) -ForegroundColor DarkGray }
                try {
                    $respBytes = $udp.Receive([ref] $remote)
                    $respJson = [System.Text.Encoding]::UTF8.GetString($respBytes)
                    if ($verbose) { Write-Host ("  <- {0}" -f $respJson) -ForegroundColor DarkGray }
                    $resp = $respJson | ConvertFrom-Json
                    $tag = if ($resp.ID -eq $id) { '' } else { (" (id mismatch: got {0}, sent {1})" -f $resp.ID, $id) }
                    if ($resp.Type -eq 2 -and $resp.Data -eq 'Success') {
                        Write-Host ("  Success{0}" -f $tag) -ForegroundColor Green
                    } else {
                        Write-Host ("  {0}{1}" -f $resp.Data, $tag) -ForegroundColor Red
                    }
                } catch [System.Net.Sockets.SocketException] {
                    Write-Host ("  (no response within {0} ms)" -f $TimeoutMs) -ForegroundColor Yellow
                }
            }
        }
    } finally {
        $udp.Close()
        Write-Host 'bye.' -ForegroundColor Green
    }
}

function Invoke-ReceiverMode {
    $bindIp = if ([string]::IsNullOrEmpty($Ip)) { '0.0.0.0' } else { $Ip }

    $udp = New-ListenerSocket $bindIp $Port
    if ($null -eq $udp) { return }
    $udp.Client.ReceiveTimeout = 400
    $remote = New-Object System.Net.IPEndPoint([System.Net.IPAddress]::Any, 0)
    $myIp = if ($bindIp -and $bindIp -ne '0.0.0.0') { $bindIp } else { Get-LocalIPv4 }
    if ($bindIp -ne '0.0.0.0') {
        Write-Warning "Bound to a specific interface ($bindIp); limited broadcasts (255.255.255.255) will NOT be received. Use the default -Ip (all interfaces) to catch broadcast traffic."
    }

    $canPollKeys = $false
    try { $canPollKeys = -not [Console]::IsInputRedirected } catch { $canPollKeys = $false }

    $stopHint = if ($canPollKeys) { "Press 'q' or Esc to stop." } else { 'Press Ctrl+C to stop.' }
    if ($RunSeconds -gt 0) { $stopHint = "Auto-stops in $RunSeconds s. $stopHint" }
    Write-Host ("ZebulonVSTO sync mirror [RECEIVER] -> listening on {0}:{1}  (reply={2})" -f $bindIp, $Port, $Reply) -ForegroundColor Green
    Write-Host $stopHint

    $count = 0
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        while ($true) {
            if ($canPollKeys) {
                try {
                    if ([Console]::KeyAvailable) {
                        $k = [Console]::ReadKey($true)
                        if ($k.Key -eq 'Q' -or $k.Key -eq 'Escape') { break }
                    }
                } catch {
                    $canPollKeys = $false   # host has no real console keyboard; rely on Ctrl+C / RunSeconds
                }
            }
            if ($RunSeconds -gt 0 -and $sw.Elapsed.TotalSeconds -ge $RunSeconds) { break }

            try {
                $bytes = $udp.Receive([ref] $remote)
            } catch [System.Net.Sockets.SocketException] {
                continue   # ReceiveTimeout elapsed; loop to poll keys / clock
            }

            $raw = [System.Text.Encoding]::UTF8.GetString($bytes)
            $ts = (Get-Date).ToString('HH:mm:ss.fff')
            $msg = $null
            try { $msg = $raw | ConvertFrom-Json } catch { }

            if ($null -eq $msg) {
                Write-Host ("[{0}] <- {1}:{2}  (unparseable) {3}" -f $ts, $remote.Address, $remote.Port, $raw) -ForegroundColor DarkYellow
                continue
            }

            $kind = switch ([int]$msg.Type) { 0 { 'CUSTOM' } 1 { 'REQUEST' } 2 { 'RESPONSE' } default { "Type$($msg.Type)" } }
            Write-Host ("[{0}] <- {1}:{2}  {3} ID={4} Data='{5}'" -f $ts, $remote.Address, $remote.Port, $kind, $msg.ID, $msg.Data)

            if ([int]$msg.Type -ne 2 -and $Reply -ne 'None' -and $msg.SenderIP -and $msg.SenderPort) {
                $respJson = '{{"SenderIP":"{0}","SenderPort":{1},"ID":{2},"Type":2,"Data":"{3}"}}' -f $myIp, $Port, $msg.ID, $Reply
                $rb = [System.Text.Encoding]::UTF8.GetBytes($respJson)
                [void] $udp.Send($rb, $rb.Length, [string]$msg.SenderIP, [int]$msg.SenderPort)
                $color = if ($Reply -eq 'Success') { 'Green' } else { 'Red' }
                Write-Host ("           -> RESPONSE {0} to {1}:{2}" -f $Reply, $msg.SenderIP, $msg.SenderPort) -ForegroundColor $color
            }
            $count++
        }
    } finally {
        $udp.Close()
        Write-Host ("stopped. handled {0} message(s)." -f $count) -ForegroundColor Green
    }
}

function ConvertFrom-DiscoveryPayload([string] $data) {
    # Parse "ZSYNC1;key=value;..." — mirror of DiscoveryPayload.Parse.
    $result = @{ Valid = $false; Role = ''; Host = ''; Ver = ''; Want = ''; IsQuery = $false; InstanceId = '' }
    if ([string]::IsNullOrEmpty($data)) { return $result }
    $tokens = $data -split ';'
    if ($tokens[0] -cne 'ZSYNC1') { return $result }   # magic is case-sensitive
    $result.Valid = $true
    for ($i = 1; $i -lt $tokens.Length; $i++) {
        $t = $tokens[$i]
        $eq = $t.IndexOf('=')
        if ($eq -le 0) { continue }
        $k = $t.Substring(0, $eq)
        $v = $t.Substring($eq + 1)
        switch ($k) {
            'role' { $result.Role = $v }
            'host' { $result.Host = $v }
            'ver'  { $result.Ver = $v }
            'want' { $result.Want = $v }
            'q'    { $result.IsQuery = ($v -eq '1') }
            'id'   { $result.InstanceId = $v }
        }
    }
    return $result
}

function Invoke-ResponderMode {
    $bindIp = if ([string]::IsNullOrEmpty($Ip)) { '0.0.0.0' } else { $Ip }
    try {
        $udp = New-Object System.Net.Sockets.UdpClient
        $udp.ExclusiveAddressUse = $false
        $udp.Client.SetSocketOption(
            [System.Net.Sockets.SocketOptionLevel]::Socket,
            [System.Net.Sockets.SocketOptionName]::ReuseAddress, $true)
        $bindAddr = if ($bindIp -eq '0.0.0.0') { [System.Net.IPAddress]::Any } else { [System.Net.IPAddress]::Parse($bindIp) }
        $udp.Client.Bind((New-Object System.Net.IPEndPoint($bindAddr, $DiscoveryPort)))
    } catch {
        Write-Host ("ERROR: cannot bind discovery UDP {0}:{1} - {2}" -f $bindIp, $DiscoveryPort, $_.Exception.GetBaseException().Message) -ForegroundColor Red
        return
    }
    $udp.Client.ReceiveTimeout = 400
    $remote = New-Object System.Net.IPEndPoint([System.Net.IPAddress]::Any, 0)
    $myIp = if ($bindIp -and $bindIp -ne '0.0.0.0') { $bindIp } else { Get-LocalIPv4 }
    $roleWire = $Role.ToUpperInvariant()
    $myId = [Guid]::NewGuid().ToString('N')   # per-process id (self-recognition, not IP)

    $canPollKeys = $false
    try { $canPollKeys = -not [Console]::IsInputRedirected } catch { $canPollKeys = $false }
    $stopHint = if ($canPollKeys) { "Press 'q' or Esc to stop." } else { 'Press Ctrl+C to stop.' }
    if ($RunSeconds -gt 0) { $stopHint = "Auto-stops in $RunSeconds s. $stopHint" }
    Write-Host ("ZebulonVSTO discovery mirror [RESPONDER] -> listening on {0}:{1}, announcing role={2} at sync :{3}" -f $bindIp, $DiscoveryPort, $roleWire, $Port) -ForegroundColor Green
    Write-Host $stopHint

    $count = 0
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        while ($true) {
            if ($canPollKeys) {
                try {
                    if ([Console]::KeyAvailable) {
                        $k = [Console]::ReadKey($true)
                        if ($k.Key -eq 'Q' -or $k.Key -eq 'Escape') { break }
                    }
                } catch { $canPollKeys = $false }
            }
            if ($RunSeconds -gt 0 -and $sw.Elapsed.TotalSeconds -ge $RunSeconds) { break }

            try {
                $bytes = $udp.Receive([ref] $remote)
            } catch [System.Net.Sockets.SocketException] {
                continue   # ReceiveTimeout elapsed; loop to poll keys / clock
            }

            $raw = [System.Text.Encoding]::UTF8.GetString($bytes)
            $msg = $null
            try { $msg = $raw | ConvertFrom-Json } catch { }
            if ($null -eq $msg -or [int]$msg.Type -ne 3) { continue }   # only DISCOVER (Type 3)

            $payload = ConvertFrom-DiscoveryPayload ([string]$msg.Data)
            if (-not $payload.Valid -or -not $payload.IsQuery) { continue }
            # Skip only our own broadcast (by InstanceId, not IP — a same-host
            # add-in shares our IP but has a different id), and honor a role filter.
            if ($payload.InstanceId -eq $myId) { continue }
            if ($payload.Want -and $payload.Want.ToUpperInvariant() -ne $roleWire) { continue }

            $annData = "ZSYNC1;role={0};host={1};ver=ps;id={2}" -f $roleWire, $env:COMPUTERNAME, $myId
            $json = '{{"SenderIP":"{0}","SenderPort":{1},"ID":{2},"Type":4,"Data":"{3}"}}' -f $myIp, $Port, [int]$msg.ID, $annData
            $rb = [System.Text.Encoding]::UTF8.GetBytes($json)
            [void] $udp.Send($rb, $rb.Length, $remote.Address.ToString(), $remote.Port)
            $ts = (Get-Date).ToString('HH:mm:ss.fff')
            Write-Host ("[{0}] <- DISCOVER from {1}:{2}  -> ANNOUNCE {3} :{4}" -f $ts, $remote.Address, $remote.Port, $roleWire, $Port) -ForegroundColor Cyan
            $count++
        }
    } finally {
        $udp.Close()
        Write-Host ("stopped. answered {0} discovery ping(s)." -f $count) -ForegroundColor Green
    }
}

function Invoke-ProbeMode {
    try {
        $udp = New-Object System.Net.Sockets.UdpClient   # ephemeral local port
        $udp.EnableBroadcast = $true
        $udp.Client.ReceiveTimeout = 300
    } catch {
        Write-Host ("ERROR: cannot open a probe socket - {0}" -f $_.Exception.GetBaseException().Message) -ForegroundColor Red
        return
    }
    $myId = [Guid]::NewGuid().ToString('N')   # per-process id (self-recognition)
    $myIp = Get-LocalIPv4
    $broadcast = New-Object System.Net.IPEndPoint([System.Net.IPAddress]::Broadcast, $DiscoveryPort)
    $remote = New-Object System.Net.IPEndPoint([System.Net.IPAddress]::Any, 0)
    $seen = @{}

    $canPollKeys = $false
    try { $canPollKeys = -not [Console]::IsInputRedirected } catch { $canPollKeys = $false }
    $stopHint = if ($canPollKeys) { "Press 'q' or Esc to stop." } else { 'Press Ctrl+C to stop.' }
    if ($RunSeconds -gt 0) { $stopHint = "Auto-stops in $RunSeconds s. $stopHint" }
    Write-Host ("ZebulonVSTO discovery mirror [PROBE] -> broadcasting DISCOVER to 255.255.255.255:{0}" -f $DiscoveryPort) -ForegroundColor Green
    Write-Host $stopHint

    $count = 0
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $lastSend = [int64](-10000)   # force an immediate first ping
    try {
        while ($true) {
            if ($canPollKeys) {
                try {
                    if ([Console]::KeyAvailable) {
                        $k = [Console]::ReadKey($true)
                        if ($k.Key -eq 'Q' -or $k.Key -eq 'Escape') { break }
                    }
                } catch { $canPollKeys = $false }
            }
            if ($RunSeconds -gt 0 -and $sw.Elapsed.TotalSeconds -ge $RunSeconds) { break }

            # Re-broadcast a DISCOVER about once a second (UDP is lossy; peers may
            # start late). Replies are buffered between sends.
            if (($sw.ElapsedMilliseconds - $lastSend) -ge 1000) {
                $data = "ZSYNC1;q=1;id={0}" -f $myId
                $json = '{{"SenderIP":"{0}","SenderPort":{1},"ID":1,"Type":3,"Data":"{2}"}}' -f $myIp, $DiscoveryPort, $data
                $b = [System.Text.Encoding]::UTF8.GetBytes($json)
                try { [void] $udp.Send($b, $b.Length, $broadcast) } catch { }
                $lastSend = $sw.ElapsedMilliseconds
            }

            try {
                $bytes = $udp.Receive([ref] $remote)
            } catch [System.Net.Sockets.SocketException] {
                continue   # ReceiveTimeout elapsed; loop to re-ping / poll keys
            } catch [System.ObjectDisposedException] {
                break
            }

            $raw = [System.Text.Encoding]::UTF8.GetString($bytes)
            $msg = $null
            try { $msg = $raw | ConvertFrom-Json } catch { }
            if ($null -eq $msg -or [int]$msg.Type -ne 4) { continue }   # ANNOUNCE only

            $payload = ConvertFrom-DiscoveryPayload ([string]$msg.Data)
            if (-not $payload.Valid) { continue }
            if ($payload.InstanceId -eq $myId) { continue }   # our own (shouldn't happen)

            $key = "{0}:{1}" -f $msg.SenderIP, $msg.SenderPort
            if ($seen.ContainsKey($key)) { continue }
            $seen[$key] = $true
            $count++
            $ts = (Get-Date).ToString('HH:mm:ss.fff')
            $role = if ($payload.Role) { $payload.Role } else { '?' }
            Write-Host ("[{0}] ANNOUNCE  {1}  {2}:{3}  role={4}  ver={5}" -f `
                $ts, $payload.Host, $msg.SenderIP, $msg.SenderPort, $role, $payload.Ver) -ForegroundColor Green
        }
    } finally {
        $udp.Close()
        Write-Host ("stopped. discovered {0} peer(s)." -f $count) -ForegroundColor Green
    }
}

function Read-ValueWithDefault([string] $prompt, [string] $default) {
    $label = if ([string]::IsNullOrEmpty($default)) { $prompt } else { "$prompt [$default]" }
    $v = Read-Host $label
    if ([string]::IsNullOrWhiteSpace($v)) { return $default }
    return $v.Trim()
}

function Read-ModeInteractive {
    while ($true) {
        Write-Host ''
        Write-Host 'Select a mode:' -ForegroundColor Cyan
        Write-Host '  1) Sender     emit commands to a receiver (REPL prompt)'
        Write-Host '  2) Receiver   listen on a port and reply to commands'
        Write-Host '  3) Responder  answer discovery pings (auto-detect stand-in)'
        Write-Host '  4) Probe      broadcast discovery pings, list responders'
        $v = (Read-Host 'Enter 1/2/3/4 [1]').Trim().ToLowerInvariant()
        switch ($v) {
            ''          { return 'Sender' }
            '1'         { return 'Sender' }
            'sender'    { return 'Sender' }
            '2'         { return 'Receiver' }
            'receiver'  { return 'Receiver' }
            '3'         { return 'Responder' }
            'responder' { return 'Responder' }
            '4'         { return 'Probe' }
            'probe'     { return 'Probe' }
            default     { Write-Host 'Please enter 1, 2, 3, or 4.' -ForegroundColor Yellow }
        }
    }
}

function Read-PortInteractive([string] $prompt, [int] $default) {
    while ($true) {
        $v = (Read-Host ("{0} [{1}]" -f $prompt, $default)).Trim()
        if ([string]::IsNullOrEmpty($v)) { return $default }
        $n = 0
        if ([int]::TryParse($v, [ref] $n) -and $n -ge 1 -and $n -le 65535) { return $n }
        Write-Host 'Port must be an integer 1-65535.' -ForegroundColor Yellow
    }
}

# Interactively fill -Mode/-Ip/-Port when they were omitted AND the session is
# interactive. Explicitly passed args are kept as-is; a redirected / non-interactive
# session keeps the defaults so automation never blocks on a prompt.
$interactive = $false
try { $interactive = -not [Console]::IsInputRedirected } catch { $interactive = $false }

if ($interactive) {
    if (-not $PSBoundParameters.ContainsKey('Mode')) { $Mode = Read-ModeInteractive }

    # Probe targets -DiscoveryPort (8290) from an ephemeral socket — it uses
    # neither -Ip nor -Port, so don't prompt for them.
    if ($Mode -ne 'Probe') {
        if (-not $PSBoundParameters.ContainsKey('Ip')) {
            if ($Mode -eq 'Sender') {
                $Ip = Read-ValueWithDefault 'Destination IP (blank = 127.0.0.1; 255.255.255.255 = broadcast)' ''
            } else {
                $Ip = Read-ValueWithDefault 'Bind interface (blank = all interfaces / 0.0.0.0)' ''
            }
        }

        if (-not $PSBoundParameters.ContainsKey('Port')) {
            $portPrompt = switch ($Mode) {
                'Sender'    { 'Destination port' }
                'Receiver'  { 'Listen port' }
                'Responder' { 'Advertised sync port' }
            }
            $Port = Read-PortInteractive $portPrompt 8291
        }
    }
}

switch ($Mode) {
    'Receiver'  { Invoke-ReceiverMode }
    'Responder' { Invoke-ResponderMode }
    'Probe'     { Invoke-ProbeMode }
    default     { Invoke-SenderMode }
}
