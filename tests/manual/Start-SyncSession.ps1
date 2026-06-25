<#
.SYNOPSIS
    A PowerShell mirror of the ZebulonVSTO UDP sync feature. Runs as a SENDER
    (emits commands to a port) or a RECEIVER (listens on a port and replies),
    without needing a second PowerPoint. Pure PowerShell — no add-in DLL.

.DESCRIPTION
    Mirrors the add-in's two roles, parameterized by a single -Port:

      -Mode Sender   -Port N   ->  emit commands to <Ip>:N  (REPL prompt)
      -Mode Receiver -Port N   ->  listen on <Ip>:N, log datagrams, reply

    SENDER binds an ephemeral local port automatically (so it never collides with
    a receiver) and reads RESPONSEs there. Session commands: help, info, verbose,
    clear, quit/bye/exit.

    RECEIVER binds <Ip>:<Port> and (unless -Reply None) answers each REQUEST.
    Stop with q/Esc, Ctrl+C, or -RunSeconds.

.PARAMETER Mode
    Sender (default) or Receiver.

.PARAMETER Port
    The port. SENDER: the destination it emits to. RECEIVER: the port it listens
    on. Default 8291.

.PARAMETER Ip
    SENDER: destination IP (default 127.0.0.1; use 255.255.255.255 to broadcast).
    RECEIVER: local interface to bind (default 0.0.0.0 = all interfaces).

.PARAMETER Reply
    RECEIVER only: answer REQUESTs with Success (default), Failed, or None.

.PARAMETER RunSeconds
    RECEIVER only: auto-stop after N seconds (0 = run until q/Esc/Ctrl+C).

.PARAMETER TimeoutMs
    SENDER only: how long to wait for each RESPONSE. Default 1000.

.EXAMPLE
    .\Start-SyncSession.ps1 -Mode Sender -Port 8291
    zebulon> select 2

.EXAMPLE
    .\Start-SyncSession.ps1 -Mode Receiver -Port 8291
#>
[CmdletBinding()]
param(
    [ValidateSet('Sender', 'Receiver')] [string] $Mode = 'Sender',
    [int]    $Port       = 8291,
    [string] $Ip         = '',
    [ValidateSet('Success', 'Failed', 'None')] [string] $Reply = 'Success',
    [int]    $RunSeconds  = 0,
    [int]    $TimeoutMs   = 1000
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

if ($Mode -eq 'Receiver') { Invoke-ReceiverMode } else { Invoke-SenderMode }
