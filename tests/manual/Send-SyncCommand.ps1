<#
.SYNOPSIS
    Emit a single sync command to a ZebulonVSTO RECEIVER and print the RESPONSE.
    One-shot equivalent of Start-SyncSession.ps1 in Sender mode.

.DESCRIPTION
    Binds an ephemeral local UDP port, emits the command as a REQUEST to
    <Ip>:<Port>, and waits for the RESPONSE. Pure PowerShell — no add-in DLL.
    Exit code: 0 if the RESPONSE is "Success", 1 otherwise (Failed / timeout).

.PARAMETER Command
    The command string: "select <n>", "showslide <n>", "hideslide", "alert <text>".

.PARAMETER Port
    Destination port to emit to. Default 8291.

.PARAMETER Ip
    Destination IP. Default 127.0.0.1 (use 255.255.255.255 to broadcast).

.PARAMETER TimeoutMs
    How long to wait for the RESPONSE. Default 1000.

.PARAMETER Id
    Correlation ID; the receiver echoes it back. Default 1.

.EXAMPLE
    .\Send-SyncCommand.ps1 -Command "select 2"

.EXAMPLE
    .\Send-SyncCommand.ps1 -Command "alert hello" -Ip 192.168.0.20 -Port 8291
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)] [string] $Command,
    [int]    $Port      = 8291,
    [string] $Ip        = "127.0.0.1",
    [int]    $TimeoutMs  = 1000,
    [int]    $Id         = 1
)

$ErrorActionPreference = "Stop"

function Get-LocalIPv4 {
    foreach ($a in [System.Net.Dns]::GetHostAddresses([System.Net.Dns]::GetHostName())) {
        if ($a.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork) {
            return $a.IPAddressToString
        }
    }
    return "127.0.0.1"
}

function ConvertTo-JsonStringLiteral([string] $s) {
    if ($null -eq $s) { return "" }
    return $s.Replace('\', '\\').Replace('"', '\"')
}

function Resolve-SenderIp([string] $targetIp, [int] $targetPort) {
    # Choose the source IP that actually routes to the target (correct on
    # multi-homed hosts), not just the first DNS IPv4.
    if ($targetIp -in @("127.0.0.1", "localhost")) { return "127.0.0.1" }
    try {
        $probe = New-Object System.Net.Sockets.UdpClient
        $probe.Connect($targetIp, $targetPort)
        $ip = $probe.Client.LocalEndPoint.Address.ToString()
        $probe.Close()
        if ($ip -and $ip -ne "0.0.0.0") { return $ip }
    } catch { }
    return Get-LocalIPv4
}

$myIp = Resolve-SenderIp $Ip $Port

$udp = New-Object System.Net.Sockets.UdpClient(0)   # ephemeral local port
try {
    if ($Ip -eq "255.255.255.255") { $udp.EnableBroadcast = $true }
    $udp.Client.ReceiveTimeout = $TimeoutMs
    $localPort = $udp.Client.LocalEndPoint.Port

    # Type 1 = REQUEST (the kind a RECEIVER executes). We advertise our local
    # port as SenderPort so the RESPONSE routes back to this socket.
    $data = ConvertTo-JsonStringLiteral $Command
    $requestJson = '{{"SenderIP":"{0}","SenderPort":{1},"ID":{2},"Type":1,"Data":"{3}"}}' -f $myIp, $localPort, $Id, $data
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($requestJson)
    [void] $udp.Send($bytes, $bytes.Length, $Ip, $Port)
    Write-Host ("-> REQUEST to {0}:{1}  (local :{2})" -f $Ip, $Port, $localPort) -ForegroundColor Cyan
    Write-Host ("   {0}" -f $requestJson)

    $remote = New-Object System.Net.IPEndPoint([System.Net.IPAddress]::Any, 0)
    try {
        $respBytes = $udp.Receive([ref] $remote)
    } catch [System.Net.Sockets.SocketException] {
        Write-Host ""
        Write-Warning ("No RESPONSE within {0} ms." -f $TimeoutMs)
        Write-Host "Check that the receiver is running on ${Ip}:${Port} and a firewall isn't blocking UDP." -ForegroundColor Yellow
        exit 1
    }

    $respJson = [System.Text.Encoding]::UTF8.GetString($respBytes)
    Write-Host ("<- RESPONSE from {0}:{1}" -f $remote.Address, $remote.Port) -ForegroundColor Cyan
    Write-Host ("   {0}" -f $respJson)

    $resp = $respJson | ConvertFrom-Json
    if ($resp.ID -ne $Id) {
        Write-Warning ("Response ID {0} does not match request ID {1}." -f $resp.ID, $Id)
    }

    if ($resp.Type -eq 2 -and $resp.Data -eq "Success") {
        Write-Host ("RESULT: Success (ID {0})" -f $resp.ID) -ForegroundColor Green
        exit 0
    } else {
        Write-Host ("RESULT: {0} (Type={1}, ID={2})" -f $resp.Data, $resp.Type, $resp.ID) -ForegroundColor Red
        exit 1
    }
} finally {
    $udp.Close()
}
