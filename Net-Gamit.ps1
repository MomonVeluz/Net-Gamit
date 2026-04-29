<# 
Net-Gamit
Windows destination testing, network health, and WLAN diagnostics tool.

Run:
  powershell.exe -ExecutionPolicy Bypass -STA -File .\Net-Gamit.ps1
#>

if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    $hostPath = (Get-Process -Id $PID).Path
    Start-Process -FilePath $hostPath -ArgumentList @(
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-STA',
        '-File', "`"$PSCommandPath`""
    )
    exit
}

$ErrorActionPreference = 'Continue'

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Xaml

function New-TextSection {
    param(
        [Parameter(Mandatory)]
        [string]$Title,

        [AllowNull()]
        [object]$Content
    )

    $body = if ($null -eq $Content) {
        'No data returned.'
    }
    elseif ($Content -is [string]) {
        if ([string]::IsNullOrWhiteSpace($Content)) { 'No data returned.' } else { $Content.TrimEnd() }
    }
    else {
        ($Content | Out-String -Width 260).TrimEnd()
    }

    return @"

===== $Title =====
$body
"@
}

function Format-NetGamitObject {
    param(
        [AllowNull()]
        [object]$InputObject
    )

    if ($null -eq $InputObject) {
        return 'No data returned.'
    }

    $text = ($InputObject | Format-List * | Out-String -Width 260).TrimEnd()
    if ([string]::IsNullOrWhiteSpace($text)) { 'No data returned.' } else { $text }
}

function Invoke-NetGamitNativeCommand {
    param(
        [Parameter(Mandatory)]
        [string]$FilePath,

        [string[]]$Arguments = @()
    )

    try {
        $result = & $FilePath @Arguments 2>&1 | Out-String -Width 260
        if ([string]::IsNullOrWhiteSpace($result)) {
            return "Command completed with no output: $FilePath $($Arguments -join ' ')"
        }
        return $result.TrimEnd()
    }
    catch {
        return "Command failed: $FilePath $($Arguments -join ' ')`r`n$($_.Exception.Message)"
    }
}

function ConvertTo-NetGamitTargetInfo {
    param(
        [Parameter(Mandatory)]
        [string]$Target,

        [string]$PortText
    )

    $cleanTarget = $Target.Trim()
    if ([string]::IsNullOrWhiteSpace($cleanTarget)) {
        throw 'Please enter a target destination.'
    }

    $uri = $null
    $hasScheme = $cleanTarget -match '^[a-zA-Z][a-zA-Z0-9+\-.]*://'
    $looksLikeUrl = $hasScheme -or $cleanTarget -match '[/?#]' -or $cleanTarget -match '^[^@\s]+\.[A-Za-z]{2,}(:\d+)?(/.*)?$'

    if ($hasScheme) {
        if (-not [System.Uri]::TryCreate($cleanTarget, [System.UriKind]::Absolute, [ref]$uri)) {
            throw "The target URL is not valid: $cleanTarget"
        }
    }
    elseif ($cleanTarget -match '[/?#]' -or $cleanTarget -match ':\d+($|/)') {
        if (-not [System.Uri]::TryCreate("https://$cleanTarget", [System.UriKind]::Absolute, [ref]$uri)) {
            throw "The target value is not valid: $cleanTarget"
        }
    }

    $hostName = if ($uri) { $uri.Host } else { $cleanTarget }
    if ($hostName.StartsWith('[') -and $hostName.EndsWith(']')) {
        $hostName = $hostName.Trim('[', ']')
    }

    $ipAddress = $null
    $isIpAddress = [System.Net.IPAddress]::TryParse($hostName, [ref]$ipAddress)

    $portWasSpecified = -not [string]::IsNullOrWhiteSpace($PortText)
    $port = $null
    if ($portWasSpecified) {
        $parsedPort = 0
        if (-not [int]::TryParse($PortText.Trim(), [ref]$parsedPort) -or $parsedPort -lt 1 -or $parsedPort -gt 65535) {
            throw 'Port must be a number from 1 to 65535.'
        }
        $port = $parsedPort
    }
    elseif ($uri -and -not $uri.IsDefaultPort) {
        $port = $uri.Port
    }
    elseif ($uri -and $uri.Scheme -eq 'http') {
        $port = 80
    }
    else {
        $port = 443
    }

    $scheme = if ($uri) { $uri.Scheme } elseif ($port -eq 80) { 'http' } else { 'https' }
    $pathAndQuery = if ($uri) { $uri.PathAndQuery } else { '/' }
    if ([string]::IsNullOrWhiteSpace($pathAndQuery)) {
        $pathAndQuery = '/'
    }

    $httpUri = $null
    if ($looksLikeUrl) {
        $builder = [System.UriBuilder]::new($scheme, $hostName, $port, $pathAndQuery.TrimStart('/'))
        if (($scheme -eq 'https' -and $port -eq 443) -or ($scheme -eq 'http' -and $port -eq 80)) {
            $builder.Port = -1
        }
        $httpUri = $builder.Uri.AbsoluteUri
    }

    [pscustomobject]@{
        Original         = $cleanTarget
        Host             = $hostName
        IsIpAddress      = $isIpAddress
        IsDnsName        = -not $isIpAddress
        LooksLikeUrl     = $looksLikeUrl
        Scheme           = $scheme
        Port             = $port
        HttpUri          = $httpUri
        PortWasSpecified = $portWasSpecified
    }
}

function Resolve-NetGamitDnsName {
    param(
        [Parameter(Mandatory)]
        [string]$HostName
    )

    try {
        if (Get-Command Resolve-DnsName -ErrorAction SilentlyContinue) {
            $records = Resolve-DnsName -Name $HostName -ErrorAction Stop |
                Select-Object Name, Type, IPAddress, NameHost, QueryType, TTL, Section

            return [pscustomobject]@{
                Succeeded = $true
                Records   = $records
                Error     = $null
            }
        }

        $addresses = [System.Net.Dns]::GetHostAddresses($HostName) |
            ForEach-Object {
                [pscustomobject]@{
                    Name      = $HostName
                    Type      = $_.AddressFamily.ToString()
                    IPAddress = $_.IPAddressToString
                    NameHost  = $null
                    QueryType = $null
                    TTL       = $null
                    Section   = $null
                }
            }

        return [pscustomobject]@{
            Succeeded = ($addresses.Count -gt 0)
            Records   = $addresses
            Error     = $null
        }
    }
    catch {
        [pscustomobject]@{
            Succeeded = $false
            Records   = @()
            Error     = $_.Exception.Message
        }
    }
}

function Get-NetGamitFirstIPv4Address {
    param(
        [Parameter(Mandatory)]
        [object]$TargetInfo,

        [AllowNull()]
        [object]$DnsResult
    )

    $parsedIp = $null
    if ([System.Net.IPAddress]::TryParse($TargetInfo.Host, [ref]$parsedIp)) {
        if ($parsedIp.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork) {
            return $parsedIp.IPAddressToString
        }

        return $null
    }

    if ($null -eq $DnsResult -or -not $DnsResult.Succeeded) {
        return $null
    }

    foreach ($record in @($DnsResult.Records)) {
        if (-not $record.PSObject.Properties['IPAddress'] -or [string]::IsNullOrWhiteSpace([string]$record.IPAddress)) {
            continue
        }

        $recordIp = $null
        if ([System.Net.IPAddress]::TryParse([string]$record.IPAddress, [ref]$recordIp) -and $recordIp.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork) {
            return $recordIp.IPAddressToString
        }
    }

    return $null
}

function Test-NetGamitPing {
    param(
        [Parameter(Mandatory)]
        [string]$HostName
    )

    $nativePing = Invoke-NetGamitNativeCommand -FilePath 'ping.exe' -Arguments @('-4', '-n', '4', '-w', '1000', $HostName)

    try {
        $testConnection = @(Test-Connection -ComputerName $HostName -Count 4 -ErrorAction Stop)
        $latencies = @()
        $successfulReplies = @()
        foreach ($reply in $testConnection) {
            $replySucceeded = $false
            if ($reply.PSObject.Properties['StatusCode']) {
                $replySucceeded = ([int]$reply.StatusCode -eq 0)
            }
            elseif ($reply.PSObject.Properties['Status']) {
                $replySucceeded = ($reply.Status.ToString() -eq 'Success')
            }

            if ($reply.PSObject.Properties['Latency']) {
                if ($null -ne $reply.Latency) {
                    $latencies += [double]$reply.Latency
                    $replySucceeded = $true
                }
            }
            elseif ($reply.PSObject.Properties['ResponseTime']) {
                if ($null -ne $reply.ResponseTime) {
                    $latencies += [double]$reply.ResponseTime
                    $replySucceeded = $true
                }
            }

            if ($replySucceeded) {
                $successfulReplies += $reply
            }
        }

        $averageLatency = if ($latencies.Count -gt 0) {
            [math]::Round(($latencies | Measure-Object -Average).Average, 2)
        }
        else {
            $null
        }

        [pscustomobject]@{
            Succeeded      = ($successfulReplies.Count -gt 0)
            Replies        = $successfulReplies
            AverageLatency = $averageLatency
            NativeOutput   = $nativePing
            Error          = if ($successfulReplies.Count -gt 0) { $null } else { 'No successful ICMP echo replies were returned.' }
        }
    }
    catch {
        [pscustomobject]@{
            Succeeded      = $false
            Replies        = @()
            AverageLatency = $null
            NativeOutput   = $nativePing
            Error          = $_.Exception.Message
        }
    }
}

function Test-NetGamitTcpPort {
    param(
        [Parameter(Mandatory)]
        [string]$HostName,

        [Parameter(Mandatory)]
        [int]$Port,

        [int]$TimeoutMilliseconds = 5000
    )

    $client = [System.Net.Sockets.TcpClient]::new()
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    try {
        $asyncResult = $client.BeginConnect($HostName, $Port, $null, $null)
        $completed = $asyncResult.AsyncWaitHandle.WaitOne($TimeoutMilliseconds, $false)

        if (-not $completed) {
            $client.Close()
            return [pscustomobject]@{
                Succeeded        = $false
                ResponseTimeMs   = $stopwatch.ElapsedMilliseconds
                Error            = "TCP connection timed out after $TimeoutMilliseconds ms."
                FallbackUsed     = $true
                TestNetConnection = $null
            }
        }

        $client.EndConnect($asyncResult)
        [pscustomobject]@{
            Succeeded        = $true
            ResponseTimeMs   = $stopwatch.ElapsedMilliseconds
            Error            = $null
            FallbackUsed     = $true
            TestNetConnection = $null
        }
    }
    catch {
        [pscustomobject]@{
            Succeeded        = $false
            ResponseTimeMs   = $stopwatch.ElapsedMilliseconds
            Error            = $_.Exception.Message
            FallbackUsed     = $true
            TestNetConnection = $null
        }
    }
    finally {
        $stopwatch.Stop()
        if ($client) {
            $client.Close()
        }
    }
}

function Test-NetGamitPort {
    param(
        [Parameter(Mandatory)]
        [string]$HostName,

        [Parameter(Mandatory)]
        [ValidateSet('TCP', 'UDP')]
        [string]$Protocol,

        [Parameter(Mandatory)]
        [int]$Port
    )

    if ($Protocol -eq 'TCP') {
        $testNetConnection = $null
        if (Get-Command Test-NetConnection -ErrorAction SilentlyContinue) {
            try {
                $testNetConnection = Test-NetConnection -ComputerName $HostName -Port $Port -InformationLevel Detailed -WarningAction SilentlyContinue -ErrorAction Stop
                return [pscustomobject]@{
                    Protocol          = 'TCP'
                    Port              = $Port
                    Succeeded         = [bool]$testNetConnection.TcpTestSucceeded
                    ResponseTimeMs    = $null
                    Error             = $null
                    TestNetConnection = $testNetConnection
                    Note              = 'Test-NetConnection TCP result.'
                }
            }
            catch {
                $fallback = Test-NetGamitTcpPort -HostName $HostName -Port $Port
                $fallback | Add-Member -NotePropertyName Protocol -NotePropertyValue 'TCP' -Force
                $fallback | Add-Member -NotePropertyName Port -NotePropertyValue $Port -Force
                $fallback | Add-Member -NotePropertyName Note -NotePropertyValue "Test-NetConnection failed, TcpClient fallback used. $($_.Exception.Message)" -Force
                return $fallback
            }
        }

        $tcpFallback = Test-NetGamitTcpPort -HostName $HostName -Port $Port
        $tcpFallback | Add-Member -NotePropertyName Protocol -NotePropertyValue 'TCP' -Force
        $tcpFallback | Add-Member -NotePropertyName Port -NotePropertyValue $Port -Force
        $tcpFallback | Add-Member -NotePropertyName Note -NotePropertyValue 'TcpClient fallback used because Test-NetConnection is unavailable.' -Force
        return $tcpFallback
    }

    $udpClient = $null
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    try {
        $udpClient = [System.Net.Sockets.UdpClient]::new()
        $udpClient.Client.ReceiveTimeout = 2000
        $udpClient.Connect($HostName, $Port)

        $payload = [System.Text.Encoding]::ASCII.GetBytes("Net-Gamit UDP probe $(Get-Date -Format o)")
        [void]$udpClient.Send($payload, $payload.Length)

        $remoteEndpoint = [System.Net.IPEndPoint]::new([System.Net.IPAddress]::Any, 0)
        $received = $null
        $receivedText = $null
        $responseReceived = $false

        try {
            $received = $udpClient.Receive([ref]$remoteEndpoint)
            $receivedText = [System.Text.Encoding]::ASCII.GetString($received)
            $responseReceived = $true
        }
        catch [System.Net.Sockets.SocketException] {
            if ($_.Exception.SocketErrorCode -eq [System.Net.Sockets.SocketError]::TimedOut) {
                $responseReceived = $false
            }
            else {
                throw
            }
        }

        [pscustomobject]@{
            Protocol       = 'UDP'
            Port           = $Port
            Succeeded      = $responseReceived
            ResponseTimeMs = $stopwatch.ElapsedMilliseconds
            Error          = $null
            RemoteEndpoint = if ($responseReceived) { $remoteEndpoint.ToString() } else { $null }
            Response       = $receivedText
            Note           = if ($responseReceived) {
                'UDP response was received from the target service.'
            }
            else {
                'UDP payload was sent, but no reply was received. UDP services often do not answer generic probes, so this is inconclusive rather than a definite failure.'
            }
        }
    }
    catch {
        [pscustomobject]@{
            Protocol       = 'UDP'
            Port           = $Port
            Succeeded      = $false
            ResponseTimeMs = $stopwatch.ElapsedMilliseconds
            Error          = $_.Exception.Message
            RemoteEndpoint = $null
            Response       = $null
            Note           = 'UDP probe failed locally or the destination reported an error.'
        }
    }
    finally {
        $stopwatch.Stop()
        if ($udpClient) {
            $udpClient.Close()
        }
    }
}

function Get-NetGamitSslCertificate {
    param(
        [Parameter(Mandatory)]
        [string]$HostName,

        [int]$Port = 443,

        [string]$ConnectHost
    )

    $client = $null
    $sslStream = $null

    try {
        if ([string]::IsNullOrWhiteSpace($ConnectHost)) {
            $ConnectHost = $HostName
        }

        $client = [System.Net.Sockets.TcpClient]::new()
        $asyncResult = $client.BeginConnect($ConnectHost, $Port, $null, $null)
        $connected = $asyncResult.AsyncWaitHandle.WaitOne(6000, $false)
        if (-not $connected) {
            throw "Timed out connecting to $ConnectHost on TCP/$Port for SSL inspection."
        }
        $client.EndConnect($asyncResult)

        $validationCallback = {
            param($sender, $certificate, $chain, $sslPolicyErrors)
            return $true
        }

        $sslStream = [System.Net.Security.SslStream]::new($client.GetStream(), $false, $validationCallback)
        $sslStream.AuthenticateAsClient($HostName)
        $certificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($sslStream.RemoteCertificate)

        $now = Get-Date
        [pscustomobject]@{
            Succeeded       = $true
            Subject         = $certificate.Subject
            Issuer          = $certificate.Issuer
            NotBefore       = $certificate.NotBefore
            NotAfter        = $certificate.NotAfter
            DaysUntilExpiry = [math]::Round(($certificate.NotAfter - $now).TotalDays, 1)
            Thumbprint      = $certificate.Thumbprint
            SerialNumber    = $certificate.SerialNumber
            SignatureAlgo   = $certificate.SignatureAlgorithm.FriendlyName
            SslProtocol     = $sslStream.SslProtocol.ToString()
            ConnectedTo     = "$ConnectHost`:$Port"
            SniHost         = $HostName
            Errors          = $null
        }
    }
    catch {
        [pscustomobject]@{
            Succeeded       = $false
            Subject         = $null
            Issuer          = $null
            NotBefore       = $null
            NotAfter        = $null
            DaysUntilExpiry = $null
            Thumbprint      = $null
            SerialNumber    = $null
            SignatureAlgo   = $null
            SslProtocol     = $null
            ConnectedTo     = if ($ConnectHost) { "$ConnectHost`:$Port" } else { $null }
            SniHost         = $HostName
            Errors          = $_.Exception.Message
        }
    }
    finally {
        if ($sslStream) {
            $sslStream.Dispose()
        }
        if ($client) {
            $client.Close()
        }
    }
}

function Test-NetGamitHttpResponse {
    param(
        [Parameter(Mandatory)]
        [string]$Uri,

        [string]$ConnectHost
    )

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        if (-not [string]::IsNullOrWhiteSpace($ConnectHost)) {
            $parsedUri = [System.Uri]::new($Uri)
            $port = if ($parsedUri.IsDefaultPort) {
                if ($parsedUri.Scheme -eq 'https') { 443 } else { 80 }
            }
            else {
                $parsedUri.Port
            }

            $client = $null
            $stream = $null
            $reader = $null

            try {
                $client = [System.Net.Sockets.TcpClient]::new()
                $asyncResult = $client.BeginConnect($ConnectHost, $port, $null, $null)
                $connected = $asyncResult.AsyncWaitHandle.WaitOne(10000, $false)
                if (-not $connected) {
                    throw "Timed out connecting to $ConnectHost on TCP/$port."
                }
                $client.EndConnect($asyncResult)
                $client.ReceiveTimeout = 10000
                $client.SendTimeout = 10000

                $stream = $client.GetStream()
                if ($parsedUri.Scheme -eq 'https') {
                    $validationCallback = {
                        param($sender, $certificate, $chain, $sslPolicyErrors)
                        return $true
                    }
                    $sslStream = [System.Net.Security.SslStream]::new($stream, $false, $validationCallback)
                    $sslStream.AuthenticateAsClient($parsedUri.Host)
                    $stream = $sslStream
                }

                $pathAndQuery = if ([string]::IsNullOrWhiteSpace($parsedUri.PathAndQuery)) { '/' } else { $parsedUri.PathAndQuery }
                $hostHeader = if (($parsedUri.Scheme -eq 'https' -and $port -eq 443) -or ($parsedUri.Scheme -eq 'http' -and $port -eq 80)) {
                    $parsedUri.Host
                }
                else {
                    "$($parsedUri.Host):$port"
                }

                $requestText = "HEAD $pathAndQuery HTTP/1.1`r`nHost: $hostHeader`r`nUser-Agent: Net-Gamit/1.0`r`nAccept: */*`r`nConnection: close`r`n`r`n"
                $requestBytes = [System.Text.Encoding]::ASCII.GetBytes($requestText)
                $stream.Write($requestBytes, 0, $requestBytes.Length)
                $stream.Flush()

                $reader = [System.IO.StreamReader]::new($stream, [System.Text.Encoding]::ASCII)
                $statusLine = $reader.ReadLine()
                if ([string]::IsNullOrWhiteSpace($statusLine)) {
                    throw 'No HTTP response status line was returned.'
                }

                $headers = @()
                while ($true) {
                    $line = $reader.ReadLine()
                    if ($null -eq $line -or $line.Length -eq 0) {
                        break
                    }

                    $separator = $line.IndexOf(':')
                    if ($separator -gt 0) {
                        $headers += [pscustomobject]@{
                            Name  = $line.Substring(0, $separator).Trim()
                            Value = $line.Substring($separator + 1).Trim()
                        }
                    }
                }

                $stopwatch.Stop()
                $statusCode = $null
                $statusDescription = $statusLine
                if ($statusLine -match '^HTTP/\S+\s+(\d{3})\s*(.*)$') {
                    $statusCode = [int]$matches[1]
                    $statusDescription = $matches[2].Trim()
                }

                $contentLengthHeader = $headers | Where-Object { $_.Name -ieq 'Content-Length' } | Select-Object -First 1
                $contentLength = if ($contentLengthHeader -and $contentLengthHeader.Value -match '^\d+$') { [int64]$contentLengthHeader.Value } else { $null }

                return [pscustomobject]@{
                    Succeeded         = ($statusCode -ge 200 -and $statusCode -lt 400)
                    StatusCode        = $statusCode
                    StatusDescription = $statusDescription
                    ResponseTimeMs    = $stopwatch.ElapsedMilliseconds
                    ResponseUri       = $parsedUri.AbsoluteUri
                    ContentLength     = $contentLength
                    Headers           = $headers
                    ConnectedTo       = "$ConnectHost`:$port"
                    HostHeader        = $hostHeader
                    Error             = $null
                }
            }
            finally {
                if ($reader) {
                    $reader.Dispose()
                }
                elseif ($stream) {
                    $stream.Dispose()
                }

                if ($client) {
                    $client.Close()
                }
            }
        }

        $request = [System.Net.WebRequest]::Create($Uri)
        $request.Method = 'HEAD'
        $request.Timeout = 10000
        $request.AllowAutoRedirect = $false
        $request.UserAgent = 'Net-Gamit/1.0'

        try {
            $response = $request.GetResponse()
        }
        catch [System.Net.WebException] {
            if ($_.Exception.Response -and $_.Exception.Response.StatusCode -eq [System.Net.HttpStatusCode]::MethodNotAllowed) {
                $request = [System.Net.WebRequest]::Create($Uri)
                $request.Method = 'GET'
                $request.Timeout = 10000
                $request.AllowAutoRedirect = $false
                $request.UserAgent = 'Net-Gamit/1.0'
                $response = $request.GetResponse()
            }
            elseif ($_.Exception.Response) {
                $response = $_.Exception.Response
            }
            else {
                throw
            }
        }

        $stopwatch.Stop()
        $headers = @()
        foreach ($key in $response.Headers.AllKeys) {
            $headers += [pscustomobject]@{
                Name  = $key
                Value = $response.Headers[$key]
            }
        }

        $statusCode = [int]$response.StatusCode
        $statusDescription = $response.StatusDescription
        $contentLength = $response.ContentLength
        $responseUri = $response.ResponseUri.AbsoluteUri
        $response.Close()

        [pscustomobject]@{
            Succeeded         = ($statusCode -ge 200 -and $statusCode -lt 400)
            StatusCode        = $statusCode
            StatusDescription = $statusDescription
            ResponseTimeMs    = $stopwatch.ElapsedMilliseconds
            ResponseUri       = $responseUri
            ContentLength     = $contentLength
            Headers           = $headers
            ConnectedTo       = $null
            HostHeader        = $null
            Error             = $null
        }
    }
    catch {
        $stopwatch.Stop()
        [pscustomobject]@{
            Succeeded         = $false
            StatusCode        = $null
            StatusDescription = $null
            ResponseTimeMs    = $stopwatch.ElapsedMilliseconds
            ResponseUri       = $Uri
            ContentLength     = $null
            Headers           = @()
            ConnectedTo       = if ($ConnectHost) { $ConnectHost } else { $null }
            HostHeader        = $null
            Error             = $_.Exception.Message
        }
    }
}

function New-NetGamitNetworkAnalysis {
    param(
        [Parameter(Mandatory)]
        [object]$Data
    )

    $lines = [System.Collections.Generic.List[string]]::new()
    $target = $Data.TargetInfo
    $dns = $Data.Dns
    $ping = $Data.Ping
    $port = $Data.PortTest
    $http = $Data.Http
    $cert = $Data.Certificate
    $probeHost = $Data.ProbeHost
    $ipv4Failure = $Data.IPv4Failure

    $overallReachable = $false
    if ($Data.Protocol -eq 'TCP') {
        $overallReachable = [bool]($port.Succeeded -or ($http -and $http.Succeeded))
    }
    else {
        $overallReachable = [bool]($ping.Succeeded -or $port.Succeeded)
    }

    $ipv4Addresses = @()
    $ipv6Count = 0
    if ($target.IsDnsName -and $dns -and $dns.Succeeded) {
        foreach ($record in @($dns.Records)) {
            if (-not $record.PSObject.Properties['IPAddress'] -or [string]::IsNullOrWhiteSpace([string]$record.IPAddress)) {
                continue
            }

            $parsedAddress = $null
            if (-not [System.Net.IPAddress]::TryParse([string]$record.IPAddress, [ref]$parsedAddress)) {
                continue
            }

            if ($parsedAddress.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork) {
                $ipv4Addresses += $parsedAddress.IPAddressToString
            }
            elseif ($parsedAddress.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetworkV6) {
                $ipv6Count++
            }
        }

        $ipv4Addresses = @($ipv4Addresses | Select-Object -Unique)
    }

    $verdict = 'INCONCLUSIVE'
    $summary = 'The test did not prove end-to-end reachability. Review DNS, local routing, VPN/proxy state, firewall policy, and destination service availability.'

    if ($ipv4Failure) {
        $verdict = 'BLOCKED - NO IPV4 TARGET'
        $summary = 'Test Destination Node is operating in IPv4 mode, but no usable IPv4 address was available for this target.'
    }
    elseif ($target.IsDnsName -and -not $dns.Succeeded) {
        $verdict = 'BLOCKED - DNS FAILURE'
        $summary = 'Name resolution failed, so the target could not be tested by hostname.'
    }
    elseif ($overallReachable -and -not $ping.Succeeded) {
        $verdict = 'REACHABLE - SERVICE AVAILABLE'
        $summary = 'The tested service is reachable over IPv4. ICMP ping did not answer, which commonly means ICMP is filtered while the service port is allowed.'
    }
    elseif ($overallReachable) {
        $verdict = 'REACHABLE'
        $summary = 'The target is reachable over IPv4 for the tested path and service.'
    }
    elseif ($ping.Succeeded -and $Data.Protocol -eq 'TCP' -and -not $port.Succeeded) {
        $verdict = 'HOST REACHABLE - SERVICE UNAVAILABLE'
        $summary = 'The host responded to ICMP, but the tested TCP service did not accept the connection.'
    }
    elseif ($ping.Succeeded -and $Data.Protocol -eq 'UDP' -and -not $port.Succeeded) {
        $verdict = 'HOST REACHABLE - UDP NOT CONFIRMED'
        $summary = 'The host is reachable, but the generic UDP probe did not receive an application response. This can be normal for many UDP services.'
    }

    $dnsStatus = if (-not $target.IsDnsName) {
        'Skipped - target is already an IP address.'
    }
    elseif ($dns.Succeeded -and $ipv4Addresses.Count -gt 0) {
        "Resolved - $($ipv4Addresses.Count) IPv4 A record(s) found."
    }
    elseif ($dns.Succeeded) {
        'Resolved, but no IPv4 A record was found.'
    }
    else {
        "Failed - $($dns.Error)"
    }

    $icmpStatus = if ($ping.Succeeded) {
        if ($null -ne $ping.AverageLatency) {
            "Successful - average latency $($ping.AverageLatency) ms."
        }
        else {
            'Successful.'
        }
    }
    else {
        "No reply - $($ping.Error)"
    }

    $portStatus = if ($Data.Protocol -eq 'TCP') {
        if ($port.Succeeded) {
            "Open/reachable - TCP $($target.Port) accepted the test."
        }
        else {
            "Not reachable - TCP $($target.Port) did not accept the test. $($port.Error)"
        }
    }
    else {
        if ($port.Succeeded) {
            "Response received - UDP $($target.Port) answered the probe."
        }
        else {
            "No response - UDP $($target.Port) did not answer the generic probe."
        }
    }

    $httpStatus = $null
    if ($http) {
        if ($http.StatusCode) {
            $httpStatus = "HTTP/S returned $($http.StatusCode) $($http.StatusDescription) in $($http.ResponseTimeMs) ms."
        }
        else {
            $httpStatus = "HTTP/S did not complete in $($http.ResponseTimeMs) ms. $($http.Error)"
        }
    }

    $certStatus = $null
    if ($cert) {
        if ($cert.Succeeded) {
            if ($cert.DaysUntilExpiry -lt 0) {
                $certStatus = "Expired - certificate expired on $($cert.NotAfter)."
            }
            elseif ($cert.DaysUntilExpiry -le 30) {
                $certStatus = "Warning - certificate expires on $($cert.NotAfter) ($($cert.DaysUntilExpiry) days remaining)."
            }
            else {
                $certStatus = "Valid - certificate expires on $($cert.NotAfter) ($($cert.DaysUntilExpiry) days remaining)."
            }
        }
        else {
            $certStatus = "Inspection failed - $($cert.Errors)"
        }
    }

    $lines.Add('EXECUTIVE SUMMARY')
    $lines.Add('')
    $lines.Add("Overall result : $verdict")
    $lines.Add("Target         : $($target.Original)")
    $lines.Add("Resolved host  : $($target.Host)")
    $lines.Add("Tested service : $($Data.Protocol)/$($target.Port)")
    $lines.Add("IP mode        : IPv4 only")
    $lines.Add("IPv4 tested    : $(if ($probeHost) { $probeHost } else { 'None' })")
    $lines.Add('')
    $lines.Add('Summary')
    $lines.Add($summary)
    $lines.Add('')
    $lines.Add('Key Findings')
    $lines.Add("- DNS: $dnsStatus")
    if ($target.IsDnsName -and $ipv4Addresses.Count -gt 0) {
        $lines.Add("- IPv4 selection: First A record used for active probes: $probeHost.")
        if ($ipv6Count -gt 0) {
            $lines.Add("- IPv6 handling: $ipv6Count IPv6 record(s) were present but intentionally ignored.")
        }
    }
    elseif ($ipv4Failure) {
        $lines.Add("- IPv4 selection: $ipv4Failure")
    }
    $lines.Add("- ICMP ping: $icmpStatus")
    $lines.Add("- Service port: $portStatus")
    if ($httpStatus) {
        $lines.Add("- Web response: $httpStatus")
    }
    if ($certStatus) {
        $lines.Add("- SSL certificate: $certStatus")
    }
    $lines.Add('')
    $lines.Add('Interpretation')
    if ($ipv4Failure) {
        $lines.Add('No IPv4-based diagnostic can be completed until the target has an IPv4 address or an IPv4 A record.')
    }
    elseif ($target.IsDnsName -and -not $dns.Succeeded) {
        $lines.Add('Resolve the DNS issue first. Validate DNS server reachability, VPN/DNS suffix behavior, and the target hostname record.')
    }
    elseif ($overallReachable -and -not $ping.Succeeded) {
        $lines.Add('Treat the service as reachable. The failed ping should not be considered an outage by itself because TCP/HTTP connectivity succeeded.')
    }
    elseif ($overallReachable) {
        $lines.Add('No reachability issue was detected for the tested IPv4 service path.')
    }
    elseif ($ping.Succeeded -and $Data.Protocol -eq 'TCP' -and -not $port.Succeeded) {
        $lines.Add('Network path to the host exists, but the application service may be stopped, filtered, or listening on a different port.')
    }
    else {
        $lines.Add('The result is inconclusive. Check local default gateway, firewall/VPN/proxy policy, upstream filtering, and destination service health.')
    }
    $lines.Add('')
    $lines.Add('Recommended Action')
    if ($ipv4Failure) {
        $lines.Add('Use an IPv4 address directly, or update DNS so the hostname has an IPv4 A record.')
    }
    elseif ($target.IsDnsName -and -not $dns.Succeeded) {
        $lines.Add('Investigate DNS resolution before troubleshooting routing or application availability.')
    }
    elseif ($cert -and $cert.Succeeded -and $cert.DaysUntilExpiry -lt 0) {
        $lines.Add('Renew or replace the SSL certificate immediately.')
    }
    elseif ($cert -and $cert.Succeeded -and $cert.DaysUntilExpiry -le 30) {
        $lines.Add('Plan SSL certificate renewal soon to avoid service trust issues.')
    }
    elseif ($overallReachable -and -not $ping.Succeeded) {
        $lines.Add('No connectivity fix is required for the tested service. Treat ping failure as likely ICMP filtering unless other symptoms exist.')
    }
    elseif ($overallReachable) {
        $lines.Add('No immediate reachability action is required for this target and service.')
    }
    elseif ($ping.Succeeded -and $Data.Protocol -eq 'TCP' -and -not $port.Succeeded) {
        $lines.Add('Validate the destination application service, listening port, host firewall, and upstream firewall policy.')
    }
    else {
        $lines.Add('Continue troubleshooting from local routing and firewall/VPN policy, then validate the destination service.')
    }

    return ($lines -join [Environment]::NewLine)
}

function Invoke-NetGamitNetworkDiagnostics {
    param(
        [Parameter(Mandatory)]
        [string]$Target,

        [Parameter(Mandatory)]
        [ValidateSet('TCP', 'UDP')]
        [string]$Protocol,

        [string]$PortText
    )

    $targetInfo = ConvertTo-NetGamitTargetInfo -Target $Target -PortText $PortText
    $timestamp = Get-Date
    $raw = [System.Text.StringBuilder]::new()

    [void]$raw.AppendLine("Net-Gamit Test Destination Node")
    [void]$raw.AppendLine("Timestamp : $timestamp")
    [void]$raw.AppendLine("Target    : $($targetInfo.Original)")
    [void]$raw.AppendLine("Host      : $($targetInfo.Host)")
    [void]$raw.AppendLine("Protocol  : $Protocol")
    [void]$raw.AppendLine("Port      : $($targetInfo.Port)")
    [void]$raw.AppendLine("IP Mode   : IPv4 preferred/forced")

    $dns = $null
    if ($targetInfo.IsDnsName) {
        $dns = Resolve-NetGamitDnsName -HostName $targetInfo.Host
        [void]$raw.AppendLine((New-TextSection -Title 'DNS Lookup' -Content $(if ($dns.Succeeded) { $dns.Records | Format-Table -AutoSize | Out-String -Width 260 } else { $dns.Error })))
    }
    else {
        $dns = [pscustomobject]@{
            Succeeded = $true
            Records   = @()
            Error     = $null
        }
        [void]$raw.AppendLine((New-TextSection -Title 'DNS Lookup' -Content 'Skipped because target is an IP address.'))
    }

    $probeHost = Get-NetGamitFirstIPv4Address -TargetInfo $targetInfo -DnsResult $dns
    $ipv4Failure = $null
    if ([string]::IsNullOrWhiteSpace($probeHost)) {
        $ipv4Failure = if ($targetInfo.IsIpAddress) {
            "The target '$($targetInfo.Host)' is not an IPv4 address."
        }
        elseif ($dns.Succeeded) {
            "The hostname '$($targetInfo.Host)' did not return an IPv4 A record."
        }
        else {
            "DNS did not return a usable IPv4 address for '$($targetInfo.Host)'."
        }
    }

    if ($probeHost) {
        [void]$raw.AppendLine("Selected IPv4 for probes: $probeHost")
    }
    else {
        [void]$raw.AppendLine("Selected IPv4 for probes: none")
        [void]$raw.AppendLine("IPv4 selection issue: $ipv4Failure")
    }

    if ($ipv4Failure) {
        $ping = [pscustomobject]@{
            Succeeded      = $false
            Replies        = @()
            AverageLatency = $null
            NativeOutput   = 'Skipped because no IPv4 address is available for this target.'
            Error          = $ipv4Failure
        }
        $tracert = 'Skipped because no IPv4 address is available for this target.'
        $portTest = [pscustomobject]@{
            Protocol          = $Protocol
            Port              = $targetInfo.Port
            Succeeded         = $false
            ResponseTimeMs    = $null
            Error             = $ipv4Failure
            TestNetConnection = $null
            Note              = 'Skipped because Test Destination Node is configured for IPv4 mode.'
        }
        $http = $null
        $certificate = $null

        [void]$raw.AppendLine((New-TextSection -Title 'Ping' -Content $ping.NativeOutput))
        [void]$raw.AppendLine((New-TextSection -Title 'Tracert' -Content $tracert))
        [void]$raw.AppendLine((New-TextSection -Title "$Protocol Port Test" -Content (Format-NetGamitObject $portTest)))

        $data = [pscustomobject]@{
            Timestamp   = $timestamp
            TargetInfo  = $targetInfo
            Protocol    = $Protocol
            Dns         = $dns
            ProbeHost   = $probeHost
            IPv4Failure = $ipv4Failure
            Ping        = $ping
            Tracert     = $tracert
            PortTest    = $portTest
            Http        = $http
            Certificate = $certificate
        }

        $analysis = New-NetGamitNetworkAnalysis -Data $data
        $report = @"
NET-GAMIT TEST DESTINATION NODE REPORT
Generated: $timestamp

$analysis

DETAILS
$($raw.ToString().TrimEnd())
"@

        return [pscustomobject]@{
            Data     = $data
            RawText  = $raw.ToString().TrimEnd()
            Analysis = $analysis
            Report   = $report.TrimEnd()
        }
    }

    $ping = Test-NetGamitPing -HostName $probeHost
    [void]$raw.AppendLine((New-TextSection -Title 'Ping' -Content $ping.NativeOutput))

    $tracert = Invoke-NetGamitNativeCommand -FilePath 'tracert.exe' -Arguments @('-4', '-d', '-h', '15', '-w', '750', $probeHost)
    [void]$raw.AppendLine((New-TextSection -Title 'Tracert' -Content $tracert))

    $portTest = Test-NetGamitPort -HostName $probeHost -Protocol $Protocol -Port $targetInfo.Port
    [void]$raw.AppendLine((New-TextSection -Title "$Protocol Port Test" -Content (Format-NetGamitObject $portTest)))

    if ($Protocol -eq 'TCP' -and $portTest.TestNetConnection) {
        [void]$raw.AppendLine((New-TextSection -Title 'Test-NetConnection Detail' -Content (Format-NetGamitObject $portTest.TestNetConnection)))
    }

    $http = $null
    $certificate = $null
    if ($targetInfo.HttpUri -and $Protocol -eq 'TCP') {
        $http = Test-NetGamitHttpResponse -Uri $targetInfo.HttpUri -ConnectHost $probeHost
        [void]$raw.AppendLine((New-TextSection -Title 'HTTP/S Response' -Content (Format-NetGamitObject $http)))

        if ($targetInfo.Scheme -eq 'https' -or $targetInfo.Port -eq 443) {
            $certificate = Get-NetGamitSslCertificate -HostName $targetInfo.Host -Port $targetInfo.Port -ConnectHost $probeHost
            [void]$raw.AppendLine((New-TextSection -Title 'SSL Certificate' -Content (Format-NetGamitObject $certificate)))
        }
    }
    elseif ($targetInfo.HttpUri -and $Protocol -eq 'UDP') {
        [void]$raw.AppendLine((New-TextSection -Title 'HTTP/S Response' -Content 'Skipped because UDP was selected. HTTP/S uses TCP.'))
    }

    $data = [pscustomobject]@{
        Timestamp   = $timestamp
        TargetInfo  = $targetInfo
        Protocol    = $Protocol
        Dns         = $dns
        ProbeHost   = $probeHost
        IPv4Failure = $ipv4Failure
        Ping        = $ping
        Tracert     = $tracert
        PortTest    = $portTest
        Http        = $http
        Certificate = $certificate
    }

    $analysis = New-NetGamitNetworkAnalysis -Data $data
    $report = @"
NET-GAMIT TEST DESTINATION NODE REPORT
Generated: $timestamp

$analysis

DETAILS
$($raw.ToString().TrimEnd())
"@

    [pscustomobject]@{
        Data     = $data
        RawText  = $raw.ToString().TrimEnd()
        Analysis = $analysis
        Report   = $report.TrimEnd()
    }
}

function Convert-NetshWlanInterfaces {
    param(
        [Parameter(Mandatory)]
        [string]$Text
    )

    $interfaces = [System.Collections.Generic.List[object]]::new()
    $current = [ordered]@{}

    foreach ($line in ($Text -split "`r?`n")) {
        if ($line -match '^\s*Name\s*:\s*(.+)$' -and $current.Count -gt 0) {
            $interfaces.Add([pscustomobject]$current)
            $current = [ordered]@{}
        }

        if ($line -match '^\s*([^:]+?)\s*:\s*(.*)$') {
            $key = ($matches[1].Trim() -replace '\s+', '')
            $value = $matches[2].Trim()
            if ($current.Count -eq 0 -and $key -ne 'Name') {
                continue
            }

            if ($current.Contains($key)) {
                $key = "$key$($current.Count)"
            }
            $current[$key] = $value
        }
    }

    if ($current.Count -gt 0) {
        $interfaces.Add([pscustomobject]$current)
    }

    foreach ($interface in $interfaces) {
        $bssidValue = $null
        foreach ($bssidPropertyName in @('BSSID', 'APBSSID', 'APBssId')) {
            if ($interface.PSObject.Properties[$bssidPropertyName] -and -not [string]::IsNullOrWhiteSpace([string]$interface.$bssidPropertyName)) {
                $bssidValue = [string]$interface.$bssidPropertyName
                break
            }
        }

        $signalText = $interface.Signal
        $signalPercent = $null
        $rssiEstimate = $null
        $rssiDbm = $null
        $quality = $null
        if ($signalText -match '(\d+)%') {
            $signalPercent = [int]$matches[1]
            $rssiEstimate = [math]::Round(($signalPercent / 2) - 100, 0)
            $quality = if ($signalPercent -ge 80) {
                'Excellent'
            }
            elseif ($signalPercent -ge 60) {
                'Good'
            }
            elseif ($signalPercent -ge 40) {
                'Fair'
            }
            else {
                'Poor'
            }
        }

        if ($interface.PSObject.Properties['Rssi'] -and $interface.Rssi -match '-?\d+') {
            $rssiDbm = [int]$matches[0]
        }

        $receiveRate = $null
        foreach ($ratePropertyName in @('ReceiveRateMbps', 'Receiverate(Mbps)', 'ReceiveRate(Mbps)')) {
            if ($interface.PSObject.Properties[$ratePropertyName] -and $interface.$ratePropertyName -match '[\d.]+') {
                $receiveRate = [double]$matches[0]
                break
            }
        }

        $transmitRate = $null
        foreach ($ratePropertyName in @('TransmitRateMbps', 'Transmitrate(Mbps)', 'TransmitRate(Mbps)')) {
            if ($interface.PSObject.Properties[$ratePropertyName] -and $interface.$ratePropertyName -match '[\d.]+') {
                $transmitRate = [double]$matches[0]
                break
            }
        }

        $channelNumber = $null
        $band = $interface.Band
        if ($interface.Channel -match '^\d+$') {
            $channelNumber = [int]$interface.Channel
            if ([string]::IsNullOrWhiteSpace($band)) {
                $band = if ($channelNumber -ge 1 -and $channelNumber -le 14) {
                    '2.4 GHz (inferred from channel)'
                }
                elseif ($channelNumber -ge 32 -and $channelNumber -le 177) {
                    '5 GHz (inferred from channel)'
                }
                elseif ($channelNumber -gt 177) {
                    '6 GHz or high 5 GHz channel range (inferred)'
                }
                else {
                    'Unknown'
                }
            }
        }

        $interface | Add-Member -NotePropertyName BSSID -NotePropertyValue $bssidValue -Force
        $interface | Add-Member -NotePropertyName SignalPercent -NotePropertyValue $signalPercent -Force
        $interface | Add-Member -NotePropertyName RssiDbm -NotePropertyValue $rssiDbm -Force
        $interface | Add-Member -NotePropertyName ApproxRssiDbm -NotePropertyValue $rssiEstimate -Force
        $interface | Add-Member -NotePropertyName ReceiveRateMbps -NotePropertyValue $receiveRate -Force
        $interface | Add-Member -NotePropertyName TransmitRateMbps -NotePropertyValue $transmitRate -Force
        $interface | Add-Member -NotePropertyName SignalQuality -NotePropertyValue $quality -Force
        $interface | Add-Member -NotePropertyName InferredBand -NotePropertyValue $band -Force
    }

    return $interfaces
}

function Get-NetGamitAdapterHealth {
    $adapters = @()
    try {
        $netAdapters = Get-NetAdapter -ErrorAction Stop | Sort-Object Status, Name
        foreach ($adapter in $netAdapters) {
            $ipConfig = Get-NetIPConfiguration -InterfaceIndex $adapter.ifIndex -ErrorAction SilentlyContinue
            $adapterConfig = Get-CimInstance Win32_NetworkAdapterConfiguration -Filter "InterfaceIndex=$($adapter.ifIndex)" -ErrorAction SilentlyContinue
            $duplexSetting = $null
            try {
                $duplexProperty = Get-NetAdapterAdvancedProperty -Name $adapter.Name -ErrorAction SilentlyContinue |
                    Where-Object { $_.DisplayName -match 'Speed.*Duplex|Duplex' } |
                    Select-Object -First 1

                if ($duplexProperty) {
                    $duplexSetting = $duplexProperty.DisplayValue
                }
            }
            catch {
                $duplexSetting = $null
            }

            $ipv4 = @($ipConfig.IPv4Address | ForEach-Object { "$($_.IPAddress)/$($_.PrefixLength)" }) -join ', '
            $ipv6 = @($ipConfig.IPv6Address | ForEach-Object { "$($_.IPAddress)/$($_.PrefixLength)" }) -join ', '
            $gateway = @($ipConfig.IPv4DefaultGateway.NextHop + $ipConfig.IPv6DefaultGateway.NextHop | Where-Object { $_ }) -join ', '
            $dnsServers = @($ipConfig.DNSServer.ServerAddresses | Where-Object { $_ }) -join ', '

            $adapters += [pscustomobject]@{
                Name                 = $adapter.Name
                InterfaceDescription = $adapter.InterfaceDescription
                Status               = $adapter.Status
                LinkSpeed            = $adapter.LinkSpeed
                Duplex               = $duplexSetting
                MacAddress           = $adapter.MacAddress
                IfIndex              = $adapter.ifIndex
                MediaType            = $adapter.MediaType
                PhysicalMediaType    = $adapter.PhysicalMediaType
                IPv4                 = $ipv4
                IPv6                 = $ipv6
                Gateway              = $gateway
                DnsServers           = $dnsServers
                DhcpEnabled          = $adapterConfig.DHCPEnabled
                DhcpServer           = $adapterConfig.DHCPServer
                DnsDomain            = $adapterConfig.DNSDomain
                ServiceName          = $adapterConfig.ServiceName
            }
        }
    }
    catch {
        $adapters += [pscustomobject]@{
            Name                 = 'Adapter collection failed'
            InterfaceDescription = $_.Exception.Message
            Status               = $null
            LinkSpeed            = $null
            Duplex               = $null
            MacAddress           = $null
            IfIndex              = $null
            MediaType            = $null
            PhysicalMediaType    = $null
            IPv4                 = $null
            IPv6                 = $null
            Gateway              = $null
            DnsServers           = $null
            DhcpEnabled          = $null
            DhcpServer           = $null
            DnsDomain            = $null
            ServiceName          = $null
        }
    }

    return $adapters
}

function New-NetGamitSystemHealthAnalysis {
    param(
        [Parameter(Mandatory)]
        [object]$Data
    )

    $lines = [System.Collections.Generic.List[string]]::new()
    $activeAdapters = @($Data.Adapters | Where-Object { $_.Status -eq 'Up' })
    $connectedWlan = @($Data.WlanInterfaces | Where-Object { $_.State -match 'connected' -or $_.SSID })
    $defaultRoutes = @($Data.DefaultRoutes)

    $getAdapterType = {
        param($Adapter)

        if ($null -eq $Adapter) {
            return 'Unknown'
        }

        $adapterText = "$($Adapter.Name) $($Adapter.InterfaceDescription) $($Adapter.MediaType) $($Adapter.PhysicalMediaType)"
        if ($adapterText -match 'Wireless|Wi-Fi|WiFi|WLAN|802\.11') {
            return 'WLAN'
        }
        if ($adapterText -match 'VPN|Virtual|TAP|TUN|Hyper-V|VMware|VirtualBox|Loopback|Pseudo') {
            return 'Virtual/VPN'
        }
        if ($adapterText -match 'Bluetooth') {
            return 'Bluetooth'
        }

        return 'LAN'
    }

    $formatValue = {
        param($Value, $Fallback)
        if ([string]::IsNullOrWhiteSpace([string]$Value)) {
            return $Fallback
        }
        return [string]$Value
    }

    $primary = $null
    $primaryAdapter = $null
    $primaryType = 'Unknown'
    if ($defaultRoutes.Count -gt 0) {
        $primary = $defaultRoutes | Sort-Object TotalMetric | Select-Object -First 1
        $primaryAdapter = $activeAdapters | Where-Object { $_.Name -eq $primary.InterfaceAlias -or $_.IfIndex -eq $primary.ifIndex } | Select-Object -First 1
        $primaryType = & $getAdapterType $primaryAdapter
    }

    $lines.Add('NETWORK HEALTH SUMMARY')
    $lines.Add('')

    if ($primary) {
        $gatewayText = & $formatValue $primary.NextHop 'direct/none'
        $primaryLabel = if ($primaryType -eq 'WLAN') {
            'Wi-Fi/WLAN'
        }
        elseif ($primaryType -eq 'LAN') {
            'LAN/Ethernet'
        }
        else {
            $primaryType
        }

        $lines.Add("Primary internet path: $primaryLabel via '$($primary.InterfaceAlias)'")
        $lines.Add("Gateway/next hop: $gatewayText")
        $lines.Add("Route metric: $($primary.TotalMetric) - lower metric means Windows prefers this path.")

        if ($connectedWlan.Count -gt 0 -and $primaryType -ne 'WLAN') {
            $lines.Add("Note: Wi-Fi is connected, but it is not the preferred internet path based on the current default route.")
        }
    }
    else {
        $lines.Add('Primary internet path: Not found')
        $lines.Add('No IPv4 default route was detected, so internet/routed network access may not work.')
    }

    $lines.Add('')
    $lines.Add("Active interfaces: $($activeAdapters.Count)")
    if ($activeAdapters.Count -gt 0) {
        foreach ($adapter in ($activeAdapters | Sort-Object Name)) {
            $adapterType = & $getAdapterType $adapter
            $speed = & $formatValue $adapter.LinkSpeed 'speed not reported'
            $duplex = & $formatValue $adapter.Duplex 'duplex not reported'
            if ($adapterType -eq 'WLAN' -and $duplex -eq 'duplex not reported') {
                $duplex = 'not applicable/reported for Wi-Fi'
            }
            $ipv4 = & $formatValue $adapter.IPv4 'no IPv4 address'
            $gateway = & $formatValue $adapter.Gateway 'no default gateway on this adapter'

            $lines.Add("- $($adapter.Name): $adapterType, $($adapter.Status), $speed, duplex: $duplex, IPv4: $ipv4, gateway: $gateway")
        }
    }
    else {
        $lines.Add('- No active network adapters were detected.')
    }

    $lines.Add('')
    $adapterFindings = [System.Collections.Generic.List[string]]::new()
    $apipa = @($activeAdapters | Where-Object { $_.IPv4 -match '169\.254\.' })
    if ($apipa.Count -gt 0) {
        $adapterFindings.Add("APIPA address detected on $(@($apipa.Name) -join ', '), which usually means DHCP did not provide an address.")
    }

    if ($primary -and $primaryAdapter) {
        if ([string]::IsNullOrWhiteSpace($primaryAdapter.IPv4)) {
            $adapterFindings.Add("Primary interface '$($primaryAdapter.Name)' has no IPv4 address listed.")
        }
        if ([string]::IsNullOrWhiteSpace($primaryAdapter.DnsServers)) {
            $adapterFindings.Add("Primary interface '$($primaryAdapter.Name)' has no DNS servers listed.")
        }
    }

    $activeWithoutIpv4 = @($activeAdapters | Where-Object { [string]::IsNullOrWhiteSpace($_.IPv4) })
    if ($activeWithoutIpv4.Count -gt 0) {
        $adapterFindings.Add("Active adapter(s) without IPv4: $(@($activeWithoutIpv4.Name) -join ', '). This is fine only if they are IPv6-only, virtual, or special-purpose adapters.")
    }

    if ($adapterFindings.Count -eq 0) {
        $lines.Add('Adapter health: No obvious adapter/IP configuration errors were found in the active interfaces.')
    }
    else {
        $lines.Add('Adapter health: Needs attention')
        foreach ($finding in $adapterFindings) {
            $lines.Add("- $finding")
        }
    }

    $lines.Add('')
    if ($connectedWlan.Count -gt 0) {
        $wlan = $connectedWlan | Select-Object -First 1
        $ssid = & $formatValue $wlan.SSID 'unknown SSID'
        $bssid = & $formatValue $wlan.BSSID 'unknown BSSID'
        $radio = & $formatValue $wlan.RadioType 'radio type not reported'
        $band = & $formatValue $wlan.InferredBand 'band not reported'
        $signal = & $formatValue $wlan.Signal 'signal not reported'
        $rssi = if ($null -ne $wlan.RssiDbm) {
            "RSSI $($wlan.RssiDbm) dBm"
        }
        elseif ($null -ne $wlan.ApproxRssiDbm) {
            "about $($wlan.ApproxRssiDbm) dBm estimated"
        }
        else {
            'RSSI not reported'
        }
        $signalQuality = & $formatValue $wlan.SignalQuality 'unknown'

        $lines.Add('Wi-Fi status: Connected')
        $lines.Add("SSID: $ssid")
        $lines.Add("AP/BSSID: $bssid")
        $lines.Add("Radio/band/channel: $radio, $band, channel $($wlan.Channel)")
        $lines.Add("Signal: $signal ($signalQuality, $rssi)")

        if ($wlan.SignalPercent -ne $null -and $wlan.SignalPercent -ge 80) {
            $lines.Add('Wi-Fi interpretation: Signal is excellent. Wi-Fi signal strength is not likely to be the cause of slowness or drops.')
        }
        elseif ($wlan.SignalPercent -ne $null -and $wlan.SignalPercent -ge 60) {
            $lines.Add('Wi-Fi interpretation: Signal is good. It should be stable for normal work.')
        }
        elseif ($wlan.SignalPercent -ne $null -and $wlan.SignalPercent -ge 40) {
            $lines.Add('Wi-Fi interpretation: Signal is fair. Latency or throughput may degrade, especially under load.')
        }
        elseif ($wlan.SignalPercent -ne $null) {
            $lines.Add('Wi-Fi interpretation: Signal is poor. Expect roaming, retransmissions, reduced throughput, or intermittent drops.')
        }
        else {
            $lines.Add('Wi-Fi interpretation: Signal quality could not be determined from netsh output.')
        }
    }
    elseif ($activeAdapters | Where-Object { $_.InterfaceDescription -match 'Wireless|Wi-Fi|WiFi|802\.11' -and $_.Status -eq 'Up' }) {
        $lines.Add('Wi-Fi status: Wireless adapter is up, but netsh did not report an active WLAN connection.')
    }
    else {
        $lines.Add('Wi-Fi status: Not connected or no WLAN adapter was reported.')
    }

    $lines.Add('')
    if ($defaultRoutes.Count -gt 1) {
        $lines.Add("Routing note: $($defaultRoutes.Count) default routes exist. Windows should prefer '$($primary.InterfaceAlias)' because it has the lowest route metric.")
    }
    else {
        $lines.Add('Routing note: A single default route is present.')
    }

    $lines.Add('')
    $lines.Add('Bottom line:')
    if ($adapterFindings.Count -eq 0 -and $primary) {
        $lines.Add("The machine appears network-ready. Primary internet traffic should leave through '$($primary.InterfaceAlias)' ($primaryType).")
    }
    elseif (-not $primary) {
        $lines.Add('The machine may not have internet access because no default route was found.')
    }
    else {
        $lines.Add("The machine has network connectivity, but review the adapter finding(s) above before assuming the path is healthy.")
    }

    return ($lines -join [Environment]::NewLine)
}

function Invoke-NetGamitSystemHealth {
    $timestamp = Get-Date
    $raw = [System.Text.StringBuilder]::new()
    [void]$raw.AppendLine("Net-Gamit Network Health Check")
    [void]$raw.AppendLine("Timestamp : $timestamp")
    [void]$raw.AppendLine("Computer  : $env:COMPUTERNAME")
    [void]$raw.AppendLine("User      : $env:USERDOMAIN\$env:USERNAME")

    $adapters = @(Get-NetGamitAdapterHealth)
    [void]$raw.AppendLine((New-TextSection -Title 'Network Adapters and IP Configuration' -Content ($adapters | Format-Table -AutoSize | Out-String -Width 260)))

    $profiles = @()
    try {
        $profiles = @(Get-NetConnectionProfile -ErrorAction Stop | Select-Object Name, InterfaceAlias, InterfaceIndex, NetworkCategory, IPv4Connectivity, IPv6Connectivity)
    }
    catch {
        $profiles = @([pscustomobject]@{ Error = $_.Exception.Message })
    }
    [void]$raw.AppendLine((New-TextSection -Title 'Network Connection Profiles' -Content ($profiles | Format-Table -AutoSize | Out-String -Width 260)))

    $dnsClientServers = @()
    try {
        $dnsClientServers = @(Get-DnsClientServerAddress -ErrorAction Stop | Select-Object InterfaceAlias, InterfaceIndex, AddressFamily, ServerAddresses)
    }
    catch {
        $dnsClientServers = @([pscustomobject]@{ Error = $_.Exception.Message })
    }
    [void]$raw.AppendLine((New-TextSection -Title 'DNS Client Server Addresses' -Content ($dnsClientServers | Format-Table -AutoSize | Out-String -Width 260)))

    $wlanRaw = Invoke-NetGamitNativeCommand -FilePath 'netsh.exe' -Arguments @('wlan', 'show', 'interfaces')
    $wlanInterfaces = @(Convert-NetshWlanInterfaces -Text $wlanRaw)
    [void]$raw.AppendLine((New-TextSection -Title 'WLAN Details from netsh' -Content $wlanRaw))
    if ($wlanInterfaces.Count -gt 0) {
        [void]$raw.AppendLine((New-TextSection -Title 'Parsed WLAN Summary' -Content ($wlanInterfaces | Format-Table Name, State, SSID, BSSID, RadioType, Channel, InferredBand, Signal, SignalQuality, RssiDbm, ApproxRssiDbm -AutoSize | Out-String -Width 260)))
    }

    $defaultRoutes = @()
    $routes = @()
    try {
        $routes = @(Get-NetRoute -AddressFamily IPv4 -ErrorAction Stop |
            Sort-Object DestinationPrefix, RouteMetric, InterfaceMetric |
            Select-Object DestinationPrefix, NextHop, InterfaceAlias, ifIndex, RouteMetric, InterfaceMetric, @{
                Name = 'TotalMetric'
                Expression = { $_.RouteMetric + $_.InterfaceMetric }
            })

        $defaultRoutes = @($routes | Where-Object { $_.DestinationPrefix -eq '0.0.0.0/0' } | Sort-Object TotalMetric)
    }
    catch {
        $routes = @([pscustomobject]@{ Error = $_.Exception.Message })
    }
    [void]$raw.AppendLine((New-TextSection -Title 'Routing Table - PowerShell View' -Content ($routes | Format-Table -AutoSize | Out-String -Width 260)))

    $routePrint = Invoke-NetGamitNativeCommand -FilePath 'route.exe' -Arguments @('print')
    [void]$raw.AppendLine((New-TextSection -Title 'route print' -Content $routePrint))

    $tcpConnections = @()
    try {
        $tcpConnections = @(Get-NetTCPConnection -ErrorAction Stop |
            Select-Object LocalAddress, LocalPort, RemoteAddress, RemotePort, State, OwningProcess, AppliedSetting)
    }
    catch {
        $tcpConnections = @([pscustomobject]@{ Error = $_.Exception.Message })
    }
    [void]$raw.AppendLine((New-TextSection -Title 'TCP Connection State Summary' -Content ($tcpConnections | Group-Object State | Sort-Object Name | Select-Object Name, Count | Format-Table -AutoSize | Out-String -Width 260)))
    [void]$raw.AppendLine((New-TextSection -Title 'TCP Listening Ports' -Content ($tcpConnections | Where-Object { $_.State -eq 'Listen' } | Sort-Object LocalPort | Format-Table -AutoSize | Out-String -Width 260)))
    [void]$raw.AppendLine((New-TextSection -Title 'TCP Established Connections' -Content ($tcpConnections | Where-Object { $_.State -eq 'Established' } | Sort-Object RemoteAddress, RemotePort | Format-Table -AutoSize | Out-String -Width 260)))

    $udpEndpoints = @()
    try {
        $udpEndpoints = @(Get-NetUDPEndpoint -ErrorAction Stop | Select-Object LocalAddress, LocalPort, OwningProcess)
    }
    catch {
        $udpEndpoints = @([pscustomobject]@{ Error = $_.Exception.Message })
    }
    [void]$raw.AppendLine((New-TextSection -Title 'UDP Listening Endpoints' -Content ($udpEndpoints | Sort-Object LocalPort | Format-Table -AutoSize | Out-String -Width 260)))

    $netstat = Invoke-NetGamitNativeCommand -FilePath 'netstat.exe' -Arguments @('-ano')
    [void]$raw.AppendLine((New-TextSection -Title 'netstat -ano' -Content $netstat))

    $neighbors = @()
    try {
        $neighbors = @(Get-NetNeighbor -ErrorAction Stop | Select-Object InterfaceAlias, IPAddress, LinkLayerAddress, State, AddressFamily)
    }
    catch {
        $neighbors = @([pscustomobject]@{ Error = $_.Exception.Message })
    }
    [void]$raw.AppendLine((New-TextSection -Title 'ARP / Neighbor Cache' -Content ($neighbors | Format-Table -AutoSize | Out-String -Width 260)))

    $firewallProfiles = @()
    try {
        $firewallProfiles = @(Get-NetFirewallProfile -ErrorAction Stop | Select-Object Name, Enabled, DefaultInboundAction, DefaultOutboundAction, AllowInboundRules, NotifyOnListen)
    }
    catch {
        $firewallProfiles = @([pscustomobject]@{ Error = $_.Exception.Message })
    }
    [void]$raw.AppendLine((New-TextSection -Title 'Windows Firewall Profiles' -Content ($firewallProfiles | Format-Table -AutoSize | Out-String -Width 260)))

    $proxy = Invoke-NetGamitNativeCommand -FilePath 'netsh.exe' -Arguments @('winhttp', 'show', 'proxy')
    [void]$raw.AppendLine((New-TextSection -Title 'WinHTTP Proxy' -Content $proxy))

    $ipconfigAll = Invoke-NetGamitNativeCommand -FilePath 'ipconfig.exe' -Arguments @('/all')
    [void]$raw.AppendLine((New-TextSection -Title 'ipconfig /all' -Content $ipconfigAll))

    $data = [pscustomobject]@{
        Timestamp         = $timestamp
        Adapters          = $adapters
        Profiles          = $profiles
        DnsClientServers  = $dnsClientServers
        WlanRaw           = $wlanRaw
        WlanInterfaces    = $wlanInterfaces
        Routes            = $routes
        DefaultRoutes     = $defaultRoutes
        TcpConnections    = $tcpConnections
        UdpEndpoints      = $udpEndpoints
        FirewallProfiles  = $firewallProfiles
        Neighbors         = $neighbors
        Proxy             = $proxy
        IpconfigAll       = $ipconfigAll
    }

    $analysis = New-NetGamitSystemHealthAnalysis -Data $data
    $report = @"
NET-GAMIT NETWORK HEALTH REPORT
Generated: $timestamp

$analysis

DETAILS
$($raw.ToString().TrimEnd())
"@

    [pscustomobject]@{
        Data     = $data
        RawText  = $raw.ToString().TrimEnd()
        Analysis = $analysis
        Report   = $report.TrimEnd()
    }
}

function Get-NetGamitWlanEventReason {
    param(
        [AllowNull()]
        [string]$Message
    )

    if ([string]::IsNullOrWhiteSpace($Message)) {
        return $null
    }

    $patterns = @(
        'Reason\s*:\s*(?<reason>[^\r\n]+)',
        'Reason Code\s*:\s*(?<reason>[^\r\n]+)',
        'Failure Reason\s*:\s*(?<reason>[^\r\n]+)',
        'Reason for disconnect(?:ion)?\s*:\s*(?<reason>[^\r\n]+)',
        'The reason is\s*(?<reason>[^\r\n]+)'
    )

    foreach ($pattern in $patterns) {
        if ($Message -match $pattern) {
            return $matches['reason'].Trim()
        }
    }

    return $null
}

function Get-NetGamitWlanEventCategory {
    param(
        [AllowNull()]
        [string]$ProviderName,

        [int]$Id,

        [AllowNull()]
        [string]$Message
    )

    $text = "$ProviderName $Id $Message"
    if ($ProviderName -match 'WLAN-AutoConfig') {
        if ($Id -in 8001, 11000, 11001) {
            return 'Association'
        }
        if ($Id -in 8003, 11004, 11005) {
            return 'Disassociation'
        }
        if ($Id -in 8002, 11006, 11010, 12011, 12012, 12013) {
            return 'Failure'
        }
    }

    if ($text -match '(?i)successfully connected|connected to a wireless network|connection.*complete|association.*success') {
        return 'Association'
    }
    if ($text -match '(?i)disconnect|disassociated|deauth|roam') {
        return 'Disassociation'
    }
    if ($text -match '(?i)driver|hardware|miniport|device.*not|device.*failed|reset|ndis|netwtw|netw[a-z0-9]*') {
        return 'Hardware/Driver'
    }
    if ($text -match '(?i)fail|failed|failure|unable|timeout|authentication|802\.1x|reason code') {
        return 'Failure'
    }

    return 'Informational'
}

function Get-NetGamitWlanEvents {
    param(
        [datetime]$Since = (Get-Date).AddHours(-6)
    )

    $events = [System.Collections.Generic.List[object]]::new()
    $wlanLogs = @(
        'Microsoft-Windows-WLAN-AutoConfig/Operational',
        'Microsoft-Windows-WiFiNetworkManager/Operational',
        'Microsoft-Windows-NetworkProfile/Operational'
    )

    foreach ($logName in $wlanLogs) {
        try {
            $log = Get-WinEvent -ListLog $logName -ErrorAction SilentlyContinue
            if (-not $log) {
                continue
            }

            $logEvents = @(Get-WinEvent -FilterHashtable @{ LogName = $logName; StartTime = $Since } -MaxEvents 300 -ErrorAction Stop)
            foreach ($event in $logEvents) {
                $message = ($event.Message -replace "`r?`n", ' ').Trim()
                $events.Add([pscustomobject]@{
                    TimeCreated      = $event.TimeCreated
                    LogName          = $event.LogName
                    ProviderName     = $event.ProviderName
                    Id               = $event.Id
                    LevelDisplayName = $event.LevelDisplayName
                    Category         = Get-NetGamitWlanEventCategory -ProviderName $event.ProviderName -Id $event.Id -Message $message
                    Reason           = Get-NetGamitWlanEventReason -Message $event.Message
                    Message          = $message
                })
            }
        }
        catch {
            $events.Add([pscustomobject]@{
                TimeCreated      = Get-Date
                LogName          = $logName
                ProviderName     = 'Net-Gamit'
                Id               = 0
                LevelDisplayName = 'Error'
                Category         = 'Collection Error'
                Reason           = $null
                Message          = $_.Exception.Message
            })
        }
    }

    try {
        $systemEvents = @(Get-WinEvent -FilterHashtable @{ LogName = 'System'; StartTime = $Since } -MaxEvents 500 -ErrorAction Stop |
            Where-Object { $_.ProviderName -match 'WLAN|Wi-?Fi|Wireless|Netwtw|NETw|NDIS|NativeWifi' -or $_.Message -match 'WLAN|Wi-?Fi|Wireless|Netwtw|802\.11|miniport|driver' })

        foreach ($event in $systemEvents) {
            $message = ($event.Message -replace "`r?`n", ' ').Trim()
            $events.Add([pscustomobject]@{
                TimeCreated      = $event.TimeCreated
                LogName          = $event.LogName
                ProviderName     = $event.ProviderName
                Id               = $event.Id
                LevelDisplayName = $event.LevelDisplayName
                Category         = Get-NetGamitWlanEventCategory -ProviderName $event.ProviderName -Id $event.Id -Message $message
                Reason           = Get-NetGamitWlanEventReason -Message $event.Message
                Message          = $message
            })
        }
    }
    catch {
        $events.Add([pscustomobject]@{
            TimeCreated      = Get-Date
            LogName          = 'System'
            ProviderName     = 'Net-Gamit'
            Id               = 0
            LevelDisplayName = 'Error'
            Category         = 'Collection Error'
            Reason           = $null
            Message          = $_.Exception.Message
        })
    }

    return @($events | Sort-Object TimeCreated -Descending)
}

function Format-NetGamitDuration {
    param(
        [AllowNull()]
        [timespan]$Duration
    )

    if ($null -eq $Duration) {
        return 'unknown'
    }

    if ($Duration.TotalHours -ge 1) {
        return ('{0}h {1}m' -f [math]::Floor($Duration.TotalHours), $Duration.Minutes)
    }

    return ('{0}m {1}s' -f [math]::Floor($Duration.TotalMinutes), $Duration.Seconds)
}

function New-NetGamitWlanAnalysis {
    param(
        [Parameter(Mandatory)]
        [object]$Data
    )

    $lines = [System.Collections.Generic.List[string]]::new()
    $now = Get-Date
    $wlan = @($Data.WlanInterfaces | Where-Object { $_.State -match 'connected' -or $_.SSID } | Select-Object -First 1)
    if ($wlan.Count -eq 0) {
        $wlan = @($Data.WlanInterfaces | Select-Object -First 1)
    }
    $wlan = $wlan | Select-Object -First 1

    $events = @($Data.Events)
    $associationEvents = @($events | Where-Object { $_.Category -eq 'Association' } | Sort-Object TimeCreated -Descending)
    $disassociationEvents = @($events | Where-Object { $_.Category -eq 'Disassociation' } | Sort-Object TimeCreated -Descending)
    $failureEvents = @($events | Where-Object { $_.Category -eq 'Failure' } | Sort-Object TimeCreated -Descending)
    $driverEvents = @($events | Where-Object { $_.Category -eq 'Hardware/Driver' } | Sort-Object TimeCreated -Descending)
    $collectionErrors = @($events | Where-Object { $_.Category -eq 'Collection Error' })

    $isConnected = ($wlan -and $wlan.State -match 'connected')
    $ssid = if ($wlan -and $wlan.SSID) { $wlan.SSID } else { 'Not connected / not reported' }
    $bssid = if ($wlan -and $wlan.BSSID) { $wlan.BSSID } else { 'Not reported' }
    $radio = if ($wlan -and $wlan.RadioType) { $wlan.RadioType } else { 'Not reported' }
    $band = if ($wlan -and $wlan.InferredBand) { $wlan.InferredBand } else { 'Not reported' }
    $channel = if ($wlan -and $wlan.Channel) { $wlan.Channel } else { 'Not reported' }
    $signal = if ($wlan -and $wlan.Signal) { $wlan.Signal } else { 'Not reported' }
    $rssi = if ($wlan -and $null -ne $wlan.RssiDbm) { "$($wlan.RssiDbm) dBm" } elseif ($wlan -and $null -ne $wlan.ApproxRssiDbm) { "about $($wlan.ApproxRssiDbm) dBm estimated" } else { 'Not reported' }
    $quality = if ($wlan -and $wlan.SignalQuality) { $wlan.SignalQuality } else { 'Unknown' }
    $receiveRate = if ($wlan -and $null -ne $wlan.ReceiveRateMbps) { "$($wlan.ReceiveRateMbps) Mbps" } else { 'Not reported' }
    $transmitRate = if ($wlan -and $null -ne $wlan.TransmitRateMbps) { "$($wlan.TransmitRateMbps) Mbps" } else { 'Not reported' }

    $lastAssociation = $associationEvents | Select-Object -First 1
    $associationText = 'Not determined from the last 6 hours of WLAN events.'
    if ($isConnected -and $lastAssociation) {
        $associationText = "$($lastAssociation.TimeCreated) ($(Format-NetGamitDuration -Duration ($now - $lastAssociation.TimeCreated)) ago)"
    }
    elseif ($isConnected) {
        $associationText = 'Currently connected; no association event was found in the last 6 hours, so the connection may be older than the event window or the log entry was unavailable.'
    }

    $verdict = 'INCONCLUSIVE'
    $summary = 'WLAN adapter data was collected, but connection quality could not be fully determined.'

    if (-not $wlan) {
        $verdict = 'NO WLAN INTERFACE DATA'
        $summary = 'No WLAN interface details were returned by netsh.'
    }
    elseif (-not $isConnected) {
        $verdict = 'NOT CONNECTED'
        $summary = 'The WLAN adapter is present, but it is not currently associated to an SSID.'
    }
    elseif ($driverEvents.Count -gt 0) {
        $verdict = 'CONNECTED - DRIVER/HARDWARE EVENTS OBSERVED'
        $summary = 'The WLAN adapter is currently connected, but recent driver or hardware-related WLAN events were found.'
    }
    elseif ($failureEvents.Count -gt 0 -or $disassociationEvents.Count -gt 0) {
        $verdict = 'CONNECTED - RECENT WLAN EVENTS OBSERVED'
        $summary = 'The WLAN adapter is currently connected, but recent WLAN failures or disconnections were recorded.'
    }
    elseif ($wlan.SignalPercent -ne $null -and $wlan.SignalPercent -lt 40) {
        $verdict = 'CONNECTED - POOR SIGNAL'
        $summary = 'The WLAN adapter is connected, but signal quality is poor and may cause instability.'
    }
    elseif ($isConnected) {
        $verdict = 'CONNECTED - HEALTHY'
        $summary = 'The WLAN adapter is associated and no recent driver, hardware, or disconnection indicators were found in the collected event window.'
    }

    $lines.Add('WLAN DIAGNOSTICS SUMMARY')
    $lines.Add('')
    $lines.Add("Overall result : $verdict")
    $lines.Add("Adapter        : $(if ($wlan) { $wlan.Name } else { 'Not reported' })")
    $lines.Add("SSID           : $ssid")
    $lines.Add("AP/BSSID       : $bssid")
    $lines.Add("Radio/Band     : $radio / $band")
    $lines.Add("Channel        : $channel")
    $lines.Add("Signal         : $signal ($quality, RSSI $rssi)")
    $lines.Add("Rx/Tx rate     : $receiveRate / $transmitRate")
    $lines.Add("Associated at  : $associationText")
    $lines.Add('')
    $lines.Add('Summary')
    $lines.Add($summary)
    $lines.Add('')
    $lines.Add('Key Findings')
    $lines.Add("- WLAN state: $(if ($isConnected) { 'Connected' } else { 'Not connected or not reported' }).")
    $lines.Add("- Events reviewed: $($events.Count) WLAN-related event(s) from the last 6 hours.")
    $lines.Add("- Association events: $($associationEvents.Count).")
    $lines.Add("- Disassociation events: $($disassociationEvents.Count).")
    $lines.Add("- Failure events: $($failureEvents.Count).")
    $lines.Add("- Driver/hardware indicators: $($driverEvents.Count).")
    if ($collectionErrors.Count -gt 0) {
        $lines.Add("- Collection warnings: $($collectionErrors.Count) event source(s) could not be read.")
    }

    $lines.Add('')
    $lines.Add('Recent Association / Disassociation History')
    if ($associationEvents.Count -gt 0) {
        foreach ($event in ($associationEvents | Select-Object -First 3)) {
            $lines.Add("- Associated: $($event.TimeCreated) | Event $($event.Id) | $($event.ProviderName)")
        }
    }
    else {
        $lines.Add('- No association event found in the last 6 hours.')
    }

    if ($disassociationEvents.Count -gt 0) {
        foreach ($event in ($disassociationEvents | Select-Object -First 5)) {
            $reason = if ($event.Reason) { $event.Reason } else { 'Reason not specified in event message.' }
            $lines.Add("- Disassociated: $($event.TimeCreated) | Event $($event.Id) | Reason: $reason")
        }
    }
    else {
        $lines.Add('- No disassociation event found in the last 6 hours.')
    }

    $lines.Add('')
    $lines.Add('Interpretation')
    if ($driverEvents.Count -gt 0) {
        $lines.Add('Recent event logs include driver, miniport, NDIS, or hardware-related WLAN indicators. Review the driver version, adapter device status, and vendor driver stability.')
    }
    elseif ($failureEvents.Count -gt 0) {
        $lines.Add('Recent WLAN failures were found, but they do not clearly point to a hardware or driver fault from the collected messages. Review authentication, signal, roaming, and access point availability.')
    }
    elseif ($disassociationEvents.Count -gt 0) {
        $lines.Add('Recent WLAN disconnections occurred. Review the listed reason codes/messages, signal quality, roaming behavior, and AP stability.')
    }
    elseif ($isConnected -and $wlan.SignalPercent -ne $null -and $wlan.SignalPercent -ge 60) {
        $lines.Add('The current WLAN connection appears stable from the collected signal and event data.')
    }
    elseif ($isConnected) {
        $lines.Add('The current WLAN connection is active, but signal quality or event evidence is not strong enough to declare it fully healthy.')
    }
    else {
        $lines.Add('The adapter is not currently associated. Validate WLAN radio state, profile configuration, authentication, and driver/device status.')
    }

    $lines.Add('')
    $lines.Add('Recommended Action')
    if ($driverEvents.Count -gt 0) {
        $lines.Add('Update or reinstall the WLAN driver, verify adapter device status, and review System log driver events for repeat failures.')
    }
    elseif ($failureEvents.Count -gt 0 -or $disassociationEvents.Count -gt 0) {
        $lines.Add('Correlate the event timestamps with user impact, then validate signal quality, AP health, authentication, roaming, and power-save behavior.')
    }
    elseif ($isConnected) {
        $lines.Add('No immediate WLAN repair action is required based on the collected data.')
    }
    else {
        $lines.Add('Reconnect to the intended SSID and rerun WLAN Diagnostics after association.')
    }

    return ($lines -join [Environment]::NewLine)
}

function Invoke-NetGamitWlanDiagnostics {
    $timestamp = Get-Date
    $since = $timestamp.AddHours(-6)
    $raw = [System.Text.StringBuilder]::new()
    [void]$raw.AppendLine('Net-Gamit WLAN Diagnostics')
    [void]$raw.AppendLine("Timestamp    : $timestamp")
    [void]$raw.AppendLine("Event window : $since to $timestamp")
    [void]$raw.AppendLine("Computer     : $env:COMPUTERNAME")
    [void]$raw.AppendLine("User         : $env:USERDOMAIN\$env:USERNAME")

    $netshResults = [ordered]@{}
    $netshCommands = [ordered]@{
        'netsh wlan show interfaces'           = @('wlan', 'show', 'interfaces')
        'netsh wlan show drivers'              = @('wlan', 'show', 'drivers')
        'netsh wlan show settings'             = @('wlan', 'show', 'settings')
        'netsh wlan show profiles'             = @('wlan', 'show', 'profiles')
        'netsh wlan show networks mode=bssid'  = @('wlan', 'show', 'networks', 'mode=bssid')
        'netsh wlan show filters'              = @('wlan', 'show', 'filters')
        'netsh wlan show hostednetwork'        = @('wlan', 'show', 'hostednetwork')
        'netsh wlan show wirelesscapabilities' = @('wlan', 'show', 'wirelesscapabilities')
        'netsh wlan show wlanreport'           = @('wlan', 'show', 'wlanreport')
    }

    foreach ($name in $netshCommands.Keys) {
        $output = Invoke-NetGamitNativeCommand -FilePath 'netsh.exe' -Arguments $netshCommands[$name]
        $netshResults[$name] = $output
        [void]$raw.AppendLine((New-TextSection -Title $name -Content $output))
    }

    $wlanInterfaces = @(Convert-NetshWlanInterfaces -Text $netshResults['netsh wlan show interfaces'])
    if ($wlanInterfaces.Count -gt 0) {
        [void]$raw.AppendLine((New-TextSection -Title 'Parsed WLAN Interface Summary' -Content ($wlanInterfaces | Format-Table Name, State, SSID, BSSID, RadioType, Channel, InferredBand, Signal, SignalQuality, RssiDbm, ReceiveRateMbps, TransmitRateMbps -AutoSize | Out-String -Width 260)))
    }

    $wlanAdapters = @()
    try {
        $wlanAdapters = @(Get-NetAdapter -ErrorAction Stop | Where-Object { $_.Name -match 'Wi-?Fi|WLAN|Wireless' -or $_.InterfaceDescription -match 'Wi-?Fi|WLAN|Wireless|802\.11' } | Select-Object Name, InterfaceDescription, Status, LinkSpeed, MacAddress, ifIndex, MediaType, PhysicalMediaType, DriverInformation)
    }
    catch {
        $wlanAdapters = @([pscustomobject]@{ Error = $_.Exception.Message })
    }
    [void]$raw.AppendLine((New-TextSection -Title 'PowerShell WLAN Adapters' -Content ($wlanAdapters | Format-Table -AutoSize | Out-String -Width 260)))

    $wlanIpConfig = @()
    try {
        $aliases = @($wlanAdapters | Where-Object { $_.Name } | Select-Object -ExpandProperty Name)
        foreach ($alias in $aliases) {
            $wlanIpConfig += Get-NetIPConfiguration -InterfaceAlias $alias -ErrorAction SilentlyContinue
        }
    }
    catch {
        $wlanIpConfig = @([pscustomobject]@{ Error = $_.Exception.Message })
    }
    [void]$raw.AppendLine((New-TextSection -Title 'WLAN IP Configuration' -Content (Format-NetGamitObject $wlanIpConfig)))

    $wlanAdvanced = @()
    try {
        foreach ($adapter in @($wlanAdapters | Where-Object { $_.Name })) {
            $wlanAdvanced += Get-NetAdapterAdvancedProperty -Name $adapter.Name -ErrorAction SilentlyContinue |
                Select-Object Name, DisplayName, DisplayValue, RegistryKeyword, RegistryValue
        }
    }
    catch {
        $wlanAdvanced = @([pscustomobject]@{ Error = $_.Exception.Message })
    }
    [void]$raw.AppendLine((New-TextSection -Title 'WLAN Advanced Adapter Properties' -Content ($wlanAdvanced | Format-Table -AutoSize | Out-String -Width 260)))

    $wlanStats = @()
    try {
        foreach ($adapter in @($wlanAdapters | Where-Object { $_.Name })) {
            $wlanStats += Get-NetAdapterStatistics -Name $adapter.Name -ErrorAction SilentlyContinue |
                Select-Object Name, ReceivedBytes, SentBytes, ReceivedUnicastPackets, SentUnicastPackets, ReceivedDiscardedPackets, OutboundDiscardedPackets, ReceivedPacketErrors, OutboundPacketErrors
        }
    }
    catch {
        $wlanStats = @([pscustomobject]@{ Error = $_.Exception.Message })
    }
    [void]$raw.AppendLine((New-TextSection -Title 'WLAN Adapter Statistics' -Content ($wlanStats | Format-Table -AutoSize | Out-String -Width 260)))

    $wlanPnp = @()
    try {
        $wlanPnp = @(Get-PnpDevice -Class Net -ErrorAction Stop | Where-Object { $_.FriendlyName -match 'Wi-?Fi|WLAN|Wireless|802\.11' } | Select-Object Status, Class, FriendlyName, InstanceId)
    }
    catch {
        $wlanPnp = @([pscustomobject]@{ Error = $_.Exception.Message })
    }
    [void]$raw.AppendLine((New-TextSection -Title 'WLAN Plug and Play Device Status' -Content ($wlanPnp | Format-Table -AutoSize | Out-String -Width 260)))

    $events = @(Get-NetGamitWlanEvents -Since $since)
    [void]$raw.AppendLine((New-TextSection -Title 'WLAN Related Events - Last 6 Hours' -Content ($events | Select-Object TimeCreated, LogName, ProviderName, Id, LevelDisplayName, Category, Reason, Message | Format-Table -Wrap -AutoSize | Out-String -Width 260)))

    $data = [pscustomobject]@{
        Timestamp      = $timestamp
        Since          = $since
        NetshResults   = $netshResults
        WlanInterfaces = $wlanInterfaces
        WlanAdapters   = $wlanAdapters
        WlanIpConfig   = $wlanIpConfig
        WlanAdvanced   = $wlanAdvanced
        WlanStats      = $wlanStats
        WlanPnp        = $wlanPnp
        Events         = $events
    }

    $analysis = New-NetGamitWlanAnalysis -Data $data
    $report = @"
NET-GAMIT WLAN DIAGNOSTICS REPORT
Generated: $timestamp

$analysis

DETAILS
$($raw.ToString().TrimEnd())
"@

    [pscustomobject]@{
        Data     = $data
        RawText  = $raw.ToString().TrimEnd()
        Analysis = $analysis
        Report   = $report.TrimEnd()
    }
}

function New-NetGamitCombinedReport {
    param(
        [AllowNull()]
        [object]$NetworkResult,

        [AllowNull()]
        [object]$HealthResult,

        [AllowNull()]
        [object]$WlanResult
    )

    if ($null -eq $NetworkResult -and $null -eq $HealthResult -and $null -eq $WlanResult) {
        return ''
    }

    $timestamp = Get-Date
    $sections = [System.Collections.Generic.List[string]]::new()
    $sections.Add("NET-GAMIT COMPLETE REPORT")
    $sections.Add("Generated: $timestamp")
    $sections.Add('')

    if ($NetworkResult) {
        $sections.Add($NetworkResult.Report)
    }

    if (($NetworkResult -and $HealthResult) -or ($NetworkResult -and $WlanResult)) {
        $sections.Add('')
    }

    if ($HealthResult) {
        $sections.Add($HealthResult.Report)
    }

    if (($NetworkResult -or $HealthResult) -and $WlanResult) {
        $sections.Add('')
    }

    if ($WlanResult) {
        $sections.Add($WlanResult.Report)
    }

    return ($sections -join [Environment]::NewLine).TrimEnd()
}

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Net-Gamit - Windows Connectivity Tool"
        Height="820" Width="1180" MinHeight="720" MinWidth="980"
        WindowStartupLocation="CenterScreen"
        Background="#F4F4F4"
        FontFamily="Segoe UI">
    <Window.Resources>
        <SolidColorBrush x:Key="ThermoRedBrush" Color="#E31B23"/>
        <SolidColorBrush x:Key="ThermoRedDarkBrush" Color="#B5121B"/>
        <SolidColorBrush x:Key="ThermoBlackBrush" Color="#1D1D1B"/>
        <SolidColorBrush x:Key="ThermoCharcoalBrush" Color="#343434"/>
        <SolidColorBrush x:Key="ThermoGrayBrush" Color="#6E6E6E"/>
        <SolidColorBrush x:Key="AppBackgroundBrush" Color="#F4F4F4"/>
        <SolidColorBrush x:Key="PanelBackgroundBrush" Color="#FFFFFF"/>
        <SolidColorBrush x:Key="PanelBorderBrush" Color="#D5D5D5"/>
        <SolidColorBrush x:Key="MutedTextBrush" Color="#4A4A4A"/>
        <SolidColorBrush x:Key="DisabledSurfaceBrush" Color="#D7D7D7"/>
        <SolidColorBrush x:Key="DisabledBorderBrush" Color="#C9CDD2"/>
        <SolidColorBrush x:Key="DisabledTextBrush" Color="#7A8088"/>
        <Style TargetType="TextBlock">
            <Setter Property="Foreground" Value="{StaticResource ThermoBlackBrush}"/>
        </Style>
        <Style TargetType="Button">
            <Setter Property="Padding" Value="14,8"/>
            <Setter Property="Margin" Value="6,0,0,0"/>
            <Setter Property="MinWidth" Value="92"/>
            <Setter Property="Background" Value="{StaticResource ThermoRedBrush}"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderBrush" Value="{StaticResource ThermoRedBrush}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="{StaticResource ThermoRedDarkBrush}"/>
                    <Setter Property="BorderBrush" Value="{StaticResource ThermoRedDarkBrush}"/>
                </Trigger>
                <Trigger Property="IsEnabled" Value="False">
                    <Setter Property="Background" Value="{StaticResource DisabledSurfaceBrush}"/>
                    <Setter Property="BorderBrush" Value="{StaticResource DisabledBorderBrush}"/>
                    <Setter Property="Foreground" Value="{StaticResource DisabledTextBrush}"/>
                    <Setter Property="Cursor" Value="Arrow"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Padding" Value="8"/>
            <Setter Property="BorderBrush" Value="#B9B9B9"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Background" Value="{StaticResource PanelBackgroundBrush}"/>
            <Setter Property="Foreground" Value="{StaticResource ThermoBlackBrush}"/>
        </Style>
        <Style TargetType="ComboBox">
            <Setter Property="Padding" Value="7"/>
            <Setter Property="BorderBrush" Value="#B9B9B9"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Background" Value="{StaticResource PanelBackgroundBrush}"/>
            <Setter Property="Foreground" Value="{StaticResource ThermoBlackBrush}"/>
        </Style>
        <Style TargetType="TabControl">
            <Setter Property="Background" Value="{StaticResource AppBackgroundBrush}"/>
            <Setter Property="BorderBrush" Value="{StaticResource PanelBorderBrush}"/>
        </Style>
        <Style TargetType="TabItem">
            <Setter Property="Padding" Value="12,6"/>
            <Setter Property="Background" Value="#ECECEC"/>
            <Setter Property="Foreground" Value="{StaticResource ThermoBlackBrush}"/>
            <Setter Property="BorderBrush" Value="{StaticResource PanelBorderBrush}"/>
            <Style.Triggers>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="{StaticResource PanelBackgroundBrush}"/>
                    <Setter Property="Foreground" Value="{StaticResource ThermoRedBrush}"/>
                    <Setter Property="BorderBrush" Value="{StaticResource ThermoRedBrush}"/>
                    <Setter Property="FontWeight" Value="SemiBold"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="ProgressBar">
            <Setter Property="Foreground" Value="{StaticResource ThermoRedBrush}"/>
            <Setter Property="Background" Value="#E5E5E5"/>
            <Setter Property="BorderBrush" Value="#CFCFCF"/>
        </Style>
        <Style x:Key="MonoTextBox" TargetType="TextBox">
            <Setter Property="FontFamily" Value="Consolas"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="AcceptsReturn" Value="True"/>
            <Setter Property="AcceptsTab" Value="True"/>
            <Setter Property="VerticalScrollBarVisibility" Value="Auto"/>
            <Setter Property="HorizontalScrollBarVisibility" Value="Auto"/>
            <Setter Property="TextWrapping" Value="NoWrap"/>
            <Setter Property="IsReadOnly" Value="True"/>
            <Setter Property="Padding" Value="10"/>
            <Setter Property="BorderBrush" Value="#C7C7C7"/>
            <Setter Property="Background" Value="{StaticResource PanelBackgroundBrush}"/>
            <Setter Property="Foreground" Value="#111111"/>
        </Style>
    </Window.Resources>
    <Grid Margin="18">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <Border Grid.Row="0" Background="{StaticResource ThermoBlackBrush}" BorderBrush="{StaticResource ThermoRedBrush}" BorderThickness="0,0,0,5" CornerRadius="8" Padding="18" Margin="0,0,0,14">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel>
                    <TextBlock Text="Net-Gamit" FontSize="28" FontWeight="Bold" Foreground="White"/>
                    <TextBlock Text="Basic Windows Connectivity Tool By MomonVeluz" Foreground="#EFEFEF" FontSize="14" Margin="0,3,0,0"/>
                </StackPanel>
                <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                    <Button Name="ExportReportButton" Content="Export Report" Background="#D7D7D7" BorderBrush="#C9CDD2" Foreground="#7A8088" IsEnabled="False"/>
                    <Button Name="ClearButton" Content="Clear" Background="#D7D7D7" BorderBrush="#C9CDD2" Foreground="#7A8088" IsEnabled="False"/>
                </StackPanel>
            </Grid>
        </Border>

        <TabControl Grid.Row="1" Name="MainTabs" Background="{StaticResource AppBackgroundBrush}">
            <TabItem Header="Test Destination Node">
                <Grid Margin="0,14,0,0">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>

                    <Border Grid.Row="0" Background="{StaticResource PanelBackgroundBrush}" BorderBrush="{StaticResource PanelBorderBrush}" BorderThickness="1" CornerRadius="8" Padding="14" Margin="0,0,0,12">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="120"/>
                                <ColumnDefinition Width="120"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <StackPanel Grid.Column="0" Margin="0,0,12,0">
                                <TextBlock Text="Target destination" FontWeight="SemiBold" Margin="0,0,0,5"/>
                                <TextBox Name="TargetBox" Height="34" ToolTip="Enter an IP address, hostname, or URL such as 8.8.8.8, example.com, or https://example.com"/>
                            </StackPanel>
                            <StackPanel Grid.Column="1" Margin="0,0,12,0">
                                <TextBlock Text="Protocol" FontWeight="SemiBold" Margin="0,0,0,5"/>
                                <ComboBox Name="ProtocolBox" Height="34">
                                    <ComboBoxItem Content="TCP" IsSelected="True"/>
                                    <ComboBoxItem Content="UDP"/>
                                </ComboBox>
                            </StackPanel>
                            <StackPanel Grid.Column="2" Margin="0,0,12,0">
                                <TextBlock Text="Port" FontWeight="SemiBold" Margin="0,0,0,5"/>
                                <TextBox Name="PortBox" Height="34" Text="443"/>
                            </StackPanel>
                            <StackPanel Grid.Column="3" Orientation="Horizontal" VerticalAlignment="Bottom">
                                <Button Name="RunDiagnosticButton" Content="Run Test" MinWidth="138"/>
                            </StackPanel>
                        </Grid>
                    </Border>

                    <Grid Grid.Row="1">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="2*"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>
                        <TextBlock Grid.Row="0" Grid.Column="0" Text="Test data" FontWeight="SemiBold" Margin="0,0,0,6"/>
                        <TextBlock Grid.Row="0" Grid.Column="1" Text="Analysis and Conclusion" FontWeight="SemiBold" Margin="12,0,0,6"/>
                        <TextBox Grid.Row="1" Grid.Column="0" Name="NetworkOutputBox" Style="{StaticResource MonoTextBox}"/>
                        <TextBox Grid.Row="1" Grid.Column="1" Name="NetworkAnalysisBox" Style="{StaticResource MonoTextBox}" Margin="12,0,0,0" TextWrapping="Wrap"/>
                    </Grid>
                </Grid>
            </TabItem>

            <TabItem Header="Network Health Check">
                <Grid Margin="0,14,0,0">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    <Border Grid.Row="0" Background="{StaticResource PanelBackgroundBrush}" BorderBrush="{StaticResource PanelBorderBrush}" BorderThickness="1" CornerRadius="8" Padding="14" Margin="0,0,0,12">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <TextBlock Text="Summarize the machine's active network path, adapters, Wi-Fi signal, routes, sockets, firewall, proxy, and native Windows command output." VerticalAlignment="Center" Foreground="{StaticResource MutedTextBrush}"/>
                            <Button Grid.Column="1" Name="RunHealthButton" Content="Run Health Check" MinWidth="142"/>
                        </Grid>
                    </Border>
                    <Grid Grid.Row="1">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="2*"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>
                        <TextBlock Grid.Row="0" Grid.Column="0" Text="Network health data" FontWeight="SemiBold" Margin="0,0,0,6"/>
                        <TextBlock Grid.Row="0" Grid.Column="1" Text="Analysis and Conclusion" FontWeight="SemiBold" Margin="12,0,0,6"/>
                        <TextBox Grid.Row="1" Grid.Column="0" Name="HealthOutputBox" Style="{StaticResource MonoTextBox}"/>
                        <TextBox Grid.Row="1" Grid.Column="1" Name="HealthAnalysisBox" Style="{StaticResource MonoTextBox}" Margin="12,0,0,0" TextWrapping="Wrap"/>
                    </Grid>
                </Grid>
            </TabItem>

            <TabItem Header="WLAN Diagnostics">
                <Grid Margin="0,14,0,0">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    <Border Grid.Row="0" Background="{StaticResource PanelBackgroundBrush}" BorderBrush="{StaticResource PanelBorderBrush}" BorderThickness="1" CornerRadius="8" Padding="14" Margin="0,0,0,12">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <TextBlock Text="Collect WLAN adapter details, netsh wlan output, PowerShell adapter data, and WLAN-related event logs from the last 6 hours." VerticalAlignment="Center" Foreground="{StaticResource MutedTextBrush}"/>
                            <Button Grid.Column="1" Name="RunWlanButton" Content="Run WLAN Diagnostics" MinWidth="168"/>
                        </Grid>
                    </Border>
                    <Grid Grid.Row="1">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="2*"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>
                        <TextBlock Grid.Row="0" Grid.Column="0" Text="WLAN diagnostic data" FontWeight="SemiBold" Margin="0,0,0,6"/>
                        <TextBlock Grid.Row="0" Grid.Column="1" Text="Analysis and Conclusion" FontWeight="SemiBold" Margin="12,0,0,6"/>
                        <TextBox Grid.Row="1" Grid.Column="0" Name="WlanOutputBox" Style="{StaticResource MonoTextBox}"/>
                        <TextBox Grid.Row="1" Grid.Column="1" Name="WlanAnalysisBox" Style="{StaticResource MonoTextBox}" Margin="12,0,0,0" TextWrapping="Wrap"/>
                    </Grid>
                </Grid>
            </TabItem>

            <TabItem Header="Reports">
                <Grid Margin="0,14,0,0">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    <Border Grid.Row="0" Background="{StaticResource PanelBackgroundBrush}" BorderBrush="{StaticResource PanelBorderBrush}" BorderThickness="1" CornerRadius="8" Padding="14" Margin="0,0,0,12">
                        <TextBlock Text="Reports are generated automatically after Test Destination Node, Network Health Check, or WLAN Diagnostics completes." Foreground="{StaticResource MutedTextBrush}"/>
                    </Border>
                    <TextBox Grid.Row="1" Name="ReportBox" Style="{StaticResource MonoTextBox}" TextWrapping="NoWrap"/>
                </Grid>
            </TabItem>
        </TabControl>

        <Border Grid.Row="2" Background="#EDEDED" BorderBrush="{StaticResource PanelBorderBrush}" BorderThickness="1" CornerRadius="6" Padding="10" Margin="0,12,0,0">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="220"/>
                </Grid.ColumnDefinitions>
                <TextBlock Name="StatusText" Text="Ready." VerticalAlignment="Center" Foreground="{StaticResource MutedTextBrush}"/>
                <ProgressBar Grid.Column="1" Name="BusyProgress" Height="16" IsIndeterminate="False"/>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

$reader = [System.Xml.XmlNodeReader]::new($xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

$TargetBox = $window.FindName('TargetBox')
$ProtocolBox = $window.FindName('ProtocolBox')
$PortBox = $window.FindName('PortBox')
$RunDiagnosticButton = $window.FindName('RunDiagnosticButton')
$NetworkOutputBox = $window.FindName('NetworkOutputBox')
$NetworkAnalysisBox = $window.FindName('NetworkAnalysisBox')
$RunHealthButton = $window.FindName('RunHealthButton')
$HealthOutputBox = $window.FindName('HealthOutputBox')
$HealthAnalysisBox = $window.FindName('HealthAnalysisBox')
$RunWlanButton = $window.FindName('RunWlanButton')
$WlanOutputBox = $window.FindName('WlanOutputBox')
$WlanAnalysisBox = $window.FindName('WlanAnalysisBox')
$ReportBox = $window.FindName('ReportBox')
$ExportReportButton = $window.FindName('ExportReportButton')
$ClearButton = $window.FindName('ClearButton')
$StatusText = $window.FindName('StatusText')
$BusyProgress = $window.FindName('BusyProgress')
$MainTabs = $window.FindName('MainTabs')

$script:LastNetworkResult = $null
$script:LastHealthResult = $null
$script:LastWlanResult = $null
$script:LastNetworkReportText = $null
$script:LastHealthReportText = $null
$script:LastWlanReportText = $null
$script:ActiveWorkers = [System.Collections.Generic.List[object]]::new()

function Get-NetGamitWorkerScript {
    param(
        [Parameter(Mandatory)]
        [scriptblock]$Work
    )

    $functionNames = @(
        'New-TextSection',
        'Format-NetGamitObject',
        'Invoke-NetGamitNativeCommand',
        'ConvertTo-NetGamitTargetInfo',
        'Resolve-NetGamitDnsName',
        'Get-NetGamitFirstIPv4Address',
        'Test-NetGamitPing',
        'Test-NetGamitTcpPort',
        'Test-NetGamitPort',
        'Get-NetGamitSslCertificate',
        'Test-NetGamitHttpResponse',
        'New-NetGamitNetworkAnalysis',
        'Invoke-NetGamitNetworkDiagnostics',
        'Convert-NetshWlanInterfaces',
        'Get-NetGamitAdapterHealth',
        'New-NetGamitSystemHealthAnalysis',
        'Invoke-NetGamitSystemHealth',
        'Get-NetGamitWlanEventReason',
        'Get-NetGamitWlanEventCategory',
        'Get-NetGamitWlanEvents',
        'Format-NetGamitDuration',
        'New-NetGamitWlanAnalysis',
        'Invoke-NetGamitWlanDiagnostics'
    )

    $builder = [System.Text.StringBuilder]::new()
    [void]$builder.AppendLine('$ErrorActionPreference = ''Continue''')
    [void]$builder.AppendLine('$ProgressPreference = ''SilentlyContinue''')

    foreach ($functionName in $functionNames) {
        $command = Get-Command -Name $functionName -CommandType Function -ErrorAction Stop
        [void]$builder.AppendLine("function $functionName {")
        [void]$builder.AppendLine($command.ScriptBlock.ToString())
        [void]$builder.AppendLine('}')
    }

    [void]$builder.AppendLine('& {')
    [void]$builder.AppendLine($Work.ToString())
    [void]$builder.AppendLine('}')

    return $builder.ToString()
}

function Set-NetGamitBusy {
    param(
        [bool]$IsBusy,
        [string]$Message
    )

    $RunDiagnosticButton.IsEnabled = -not $IsBusy
    $RunHealthButton.IsEnabled = -not $IsBusy
    $RunWlanButton.IsEnabled = -not $IsBusy
    Set-NetGamitReportActionState -IsBusy $IsBusy
    $BusyProgress.IsIndeterminate = $IsBusy
    $StatusText.Text = $Message
}

function Test-NetGamitReportHasContent {
    if ($null -eq $ReportBox) {
        return $false
    }

    return -not [string]::IsNullOrWhiteSpace($ReportBox.Text)
}

function Set-NetGamitReportActionState {
    param(
        [bool]$IsBusy = $false
    )

    $hasReport = Test-NetGamitReportHasContent
    $canUseReportActions = (-not $IsBusy -and $hasReport)
    $ExportReportButton.IsEnabled = $canUseReportActions
    $ClearButton.IsEnabled = (-not $IsBusy -and $hasReport)

    if ($canUseReportActions) {
        $ExportReportButton.Background = '#E31B23'
        $ExportReportButton.BorderBrush = '#E31B23'
        $ExportReportButton.Foreground = 'White'
        $ClearButton.Background = '#343434'
        $ClearButton.BorderBrush = '#343434'
        $ClearButton.Foreground = 'White'
    }
    else {
        $ExportReportButton.Background = '#D7D7D7'
        $ExportReportButton.BorderBrush = '#C9CDD2'
        $ExportReportButton.Foreground = '#7A8088'
        $ClearButton.Background = '#D7D7D7'
        $ClearButton.BorderBrush = '#C9CDD2'
        $ClearButton.Foreground = '#7A8088'
    }
}

function Get-NetGamitUsefulText {
    param(
        [AllowNull()]
        [string]$Text
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return $null
    }

    $trimmed = $Text.Trim()
    $placeholderPatterns = @(
        '^Working\.\.\.$',
        '^Running diagnostics against .*\.\.\.$',
        '^Collecting network health data\.\.\.$',
        '^Collecting system health data\.\.\.$'
    )

    foreach ($pattern in $placeholderPatterns) {
        if ($trimmed -match $pattern) {
            return $null
        }
    }

    return $trimmed
}

function Get-NetGamitResultPropertyText {
    param(
        [AllowNull()]
        [object]$Result,

        [Parameter(Mandatory)]
        [string]$PropertyName
    )

    if ($null -eq $Result) {
        return $null
    }

    if (($Result -is [System.Array]) -or (($Result -is [System.Collections.IEnumerable]) -and -not ($Result -is [string]) -and -not $Result.PSObject.Properties[$PropertyName])) {
        foreach ($item in $Result) {
            $value = Get-NetGamitResultPropertyText -Result $item -PropertyName $PropertyName
            if (-not [string]::IsNullOrWhiteSpace($value)) {
                return $value
            }
        }
        return $null
    }

    if ($Result.PSObject.Properties[$PropertyName]) {
        return Get-NetGamitUsefulText -Text ([string]$Result.$PropertyName)
    }

    return $null
}

function Select-NetGamitResultObject {
    param(
        [AllowNull()]
        [object]$Result
    )

    if ($null -eq $Result) {
        return $null
    }

    if (($Result -is [System.Array]) -or (($Result -is [System.Collections.IEnumerable]) -and -not ($Result -is [string]) -and -not ($Result.PSObject.Properties['Report'] -or $Result.PSObject.Properties['RawText'] -or $Result.PSObject.Properties['Analysis']))) {
        $firstItem = $null
        foreach ($item in $Result) {
            if ($null -eq $firstItem) {
                $firstItem = $item
            }

            if ($item -and ($item.PSObject.Properties['Report'] -or $item.PSObject.Properties['RawText'] -or $item.PSObject.Properties['Analysis'])) {
                return $item
            }
        }

        return $firstItem
    }

    return $Result
}

function New-NetGamitPaneReport {
    param(
        [Parameter(Mandatory)]
        [string]$Title,

        [AllowNull()]
        [string]$Analysis,

        [AllowNull()]
        [string]$Details
    )

    $analysisText = Get-NetGamitUsefulText -Text $Analysis
    $detailsText = Get-NetGamitUsefulText -Text $Details

    if ([string]::IsNullOrWhiteSpace($analysisText) -and [string]::IsNullOrWhiteSpace($detailsText)) {
        return $null
    }

    $sections = [System.Collections.Generic.List[string]]::new()
    $sections.Add($Title)
    $sections.Add("Generated: $(Get-Date)")
    $sections.Add('')

    if (-not [string]::IsNullOrWhiteSpace($analysisText)) {
        $sections.Add('SUMMARY AND INTERPRETATION')
        $sections.Add($analysisText)
        $sections.Add('')
    }

    if (-not [string]::IsNullOrWhiteSpace($detailsText)) {
        $sections.Add('DETAILS')
        $sections.Add($detailsText)
    }

    return ($sections -join [Environment]::NewLine).TrimEnd()
}

function Update-NetGamitReportView {
    $networkReport = Get-NetGamitUsefulText -Text $script:LastNetworkReportText
    if ([string]::IsNullOrWhiteSpace($networkReport)) {
        $networkReport = Get-NetGamitResultPropertyText -Result $script:LastNetworkResult -PropertyName 'Report'
    }
    if ([string]::IsNullOrWhiteSpace($networkReport)) {
        $networkReport = New-NetGamitPaneReport `
            -Title 'NET-GAMIT TEST DESTINATION NODE REPORT' `
            -Analysis $NetworkAnalysisBox.Text `
            -Details $NetworkOutputBox.Text
    }

    $healthReport = Get-NetGamitUsefulText -Text $script:LastHealthReportText
    if ([string]::IsNullOrWhiteSpace($healthReport)) {
        $healthReport = Get-NetGamitResultPropertyText -Result $script:LastHealthResult -PropertyName 'Report'
    }
    if ([string]::IsNullOrWhiteSpace($healthReport)) {
        $healthReport = New-NetGamitPaneReport `
            -Title 'NET-GAMIT NETWORK HEALTH REPORT' `
            -Analysis $HealthAnalysisBox.Text `
            -Details $HealthOutputBox.Text
    }

    $wlanReport = Get-NetGamitUsefulText -Text $script:LastWlanReportText
    if ([string]::IsNullOrWhiteSpace($wlanReport)) {
        $wlanReport = Get-NetGamitResultPropertyText -Result $script:LastWlanResult -PropertyName 'Report'
    }
    if ([string]::IsNullOrWhiteSpace($wlanReport)) {
        $wlanReport = New-NetGamitPaneReport `
            -Title 'NET-GAMIT WLAN DIAGNOSTICS REPORT' `
            -Analysis $WlanAnalysisBox.Text `
            -Details $WlanOutputBox.Text
    }

    if ([string]::IsNullOrWhiteSpace($networkReport) -and [string]::IsNullOrWhiteSpace($healthReport) -and [string]::IsNullOrWhiteSpace($wlanReport)) {
        $ReportBox.Text = ''
        Set-NetGamitReportActionState
        return
    }

    $sections = [System.Collections.Generic.List[string]]::new()
    $sections.Add('NET-GAMIT COMPLETE REPORT')
    $sections.Add("Generated: $(Get-Date)")
    $sections.Add('')

    if (-not [string]::IsNullOrWhiteSpace($networkReport)) {
        $sections.Add($networkReport)
    }

    if (-not [string]::IsNullOrWhiteSpace($healthReport)) {
        if ($sections.Count -gt 3) {
            $sections.Add('')
        }
        $sections.Add($healthReport)
    }

    if (-not [string]::IsNullOrWhiteSpace($wlanReport)) {
        if ($sections.Count -gt 3) {
            $sections.Add('')
        }
        $sections.Add($wlanReport)
    }

    $ReportBox.Text = ($sections -join [Environment]::NewLine).TrimEnd()
    Set-NetGamitReportActionState
}

function Start-NetGamitBackgroundTask {
    param(
        [Parameter(Mandatory)]
        [scriptblock]$Work,

        [Parameter(Mandatory)]
        [scriptblock]$OnComplete,

        [Parameter(Mandatory)]
        [string]$StatusMessage,

        [hashtable]$Variables = @{}
    )

    Set-NetGamitBusy -IsBusy $true -Message $StatusMessage

    if ($null -eq $script:ActiveWorkers) {
        $script:ActiveWorkers = [System.Collections.Generic.List[object]]::new()
    }

    $completeClosure = $OnComplete.GetNewClosure()
    $runspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $runspace.ApartmentState = 'MTA'
    $runspace.ThreadOptions = 'ReuseThread'
    $runspace.Open()

    foreach ($key in $Variables.Keys) {
        $runspace.SessionStateProxy.SetVariable($key, $Variables[$key])
    }

    $powerShell = [System.Management.Automation.PowerShell]::Create()
    $powerShell.Runspace = $runspace
    [void]$powerShell.AddScript((Get-NetGamitWorkerScript -Work $Work))

    $asyncResult = $powerShell.BeginInvoke()

    $task = [pscustomobject]@{
        PowerShell  = $powerShell
        Runspace    = $runspace
        AsyncResult = $asyncResult
        Timer       = $null
    }

    $timer = [System.Windows.Threading.DispatcherTimer]::new()
    $timer.Interval = [TimeSpan]::FromMilliseconds(250)
    $timer.Add_Tick({
        if (-not $asyncResult.IsCompleted) {
            return
        }

        $timer.Stop()
        try {
            $resultCollection = $powerShell.EndInvoke($asyncResult)
            if ($resultCollection.Count -eq 0 -and $powerShell.Streams.Error.Count -gt 0) {
                $errorText = ($powerShell.Streams.Error | ForEach-Object { $_.ToString() }) -join [Environment]::NewLine
                throw $errorText
            }

            $result = if ($resultCollection.Count -eq 1) {
                $resultCollection[0]
            }
            else {
                $resultCollection
            }

            Set-NetGamitBusy -IsBusy $false -Message 'Ready.'
            & $completeClosure $result
        }
        catch {
            Set-NetGamitBusy -IsBusy $false -Message 'Ready.'
            [System.Windows.MessageBox]::Show($_.Exception.Message, 'Net-Gamit Error', 'OK', 'Error') | Out-Null
        }
        finally {
            try {
                $powerShell.Dispose()
            }
            catch {
            }

            try {
                $runspace.Close()
                $runspace.Dispose()
            }
            catch {
            }

            if ($script:ActiveWorkers) {
                [void]$script:ActiveWorkers.Remove($task)
            }
        }
    }.GetNewClosure())

    $task.Timer = $timer
    $script:ActiveWorkers.Add($task)
    $timer.Start()
}

$RunDiagnosticButton.Add_Click({
    $target = $TargetBox.Text
    $portText = $PortBox.Text
    $selectedProtocolItem = [System.Windows.Controls.ComboBoxItem]$ProtocolBox.SelectedItem
    $protocol = [string]$selectedProtocolItem.Content

    $NetworkOutputBox.Text = "Testing destination node $target..."
    $NetworkAnalysisBox.Text = 'Working...'

    Start-NetGamitBackgroundTask `
        -StatusMessage "Testing destination node $target..." `
        -Variables @{
            target = $target
            protocol = $protocol
            portText = $portText
        } `
        -Work {
            Invoke-NetGamitNetworkDiagnostics -Target $target -Protocol $protocol -PortText $portText
        } `
        -OnComplete {
            param($result)
            $result = Select-NetGamitResultObject -Result $result
            $script:LastNetworkResult = $result
            $script:LastNetworkReportText = Get-NetGamitResultPropertyText -Result $result -PropertyName 'Report'
            $NetworkOutputBox.Text = Get-NetGamitResultPropertyText -Result $result -PropertyName 'RawText'
            $NetworkAnalysisBox.Text = Get-NetGamitResultPropertyText -Result $result -PropertyName 'Analysis'
            Update-NetGamitReportView
            $StatusText.Text = "Test Destination Node completed at $(Get-Date -Format 'HH:mm:ss')."
        }
})

$RunHealthButton.Add_Click({
    $HealthOutputBox.Text = 'Collecting network health data...'
    $HealthAnalysisBox.Text = 'Working...'

    Start-NetGamitBackgroundTask `
        -StatusMessage 'Running network health check...' `
        -Work {
            Invoke-NetGamitSystemHealth
        } `
        -OnComplete {
            param($result)
            $result = Select-NetGamitResultObject -Result $result
            $script:LastHealthResult = $result
            $script:LastHealthReportText = Get-NetGamitResultPropertyText -Result $result -PropertyName 'Report'
            $HealthOutputBox.Text = Get-NetGamitResultPropertyText -Result $result -PropertyName 'RawText'
            $HealthAnalysisBox.Text = Get-NetGamitResultPropertyText -Result $result -PropertyName 'Analysis'
            Update-NetGamitReportView
            $StatusText.Text = "Network health check completed at $(Get-Date -Format 'HH:mm:ss')."
        }
})

$RunWlanButton.Add_Click({
    $WlanOutputBox.Text = 'Collecting WLAN diagnostic data...'
    $WlanAnalysisBox.Text = 'Working...'

    Start-NetGamitBackgroundTask `
        -StatusMessage 'Running WLAN diagnostics...' `
        -Work {
            Invoke-NetGamitWlanDiagnostics
        } `
        -OnComplete {
            param($result)
            $result = Select-NetGamitResultObject -Result $result
            $script:LastWlanResult = $result
            $script:LastWlanReportText = Get-NetGamitResultPropertyText -Result $result -PropertyName 'Report'
            $WlanOutputBox.Text = Get-NetGamitResultPropertyText -Result $result -PropertyName 'RawText'
            $WlanAnalysisBox.Text = Get-NetGamitResultPropertyText -Result $result -PropertyName 'Analysis'
            Update-NetGamitReportView
            $StatusText.Text = "WLAN diagnostics completed at $(Get-Date -Format 'HH:mm:ss')."
        }
})

$ExportReportButton.Add_Click({
    Update-NetGamitReportView

    if ([string]::IsNullOrWhiteSpace($ReportBox.Text)) {
        [System.Windows.MessageBox]::Show('Run Test Destination Node, Network Health Check, or WLAN Diagnostics before exporting a report.', 'Net-Gamit', 'OK', 'Information') | Out-Null
        return
    }

    $dialog = [Microsoft.Win32.SaveFileDialog]::new()
    $dialog.Title = 'Export Net-Gamit Report'
    $dialog.Filter = 'Text report (*.txt)|*.txt|Markdown report (*.md)|*.md|All files (*.*)|*.*'
    $dialog.FileName = "Net-Gamit-Report-$(Get-Date -Format 'yyyyMMdd-HHmmss').txt"

    if ($dialog.ShowDialog() -eq $true) {
        [System.IO.File]::WriteAllText($dialog.FileName, $ReportBox.Text, [System.Text.Encoding]::UTF8)
        $StatusText.Text = "Report exported to $($dialog.FileName)"
    }
})

$ClearButton.Add_Click({
    $NetworkOutputBox.Clear()
    $NetworkAnalysisBox.Clear()
    $HealthOutputBox.Clear()
    $HealthAnalysisBox.Clear()
    $WlanOutputBox.Clear()
    $WlanAnalysisBox.Clear()
    $ReportBox.Clear()
    $script:LastNetworkResult = $null
    $script:LastHealthResult = $null
    $script:LastWlanResult = $null
    $script:LastNetworkReportText = $null
    $script:LastHealthReportText = $null
    $script:LastWlanReportText = $null
    Set-NetGamitReportActionState
    $StatusText.Text = 'Cleared current session data.'
})

$window.Add_Closed({
    foreach ($task in @($script:ActiveWorkers)) {
        try {
            if ($task.Timer) {
                $task.Timer.Stop()
            }
            if ($task.PowerShell -and $task.AsyncResult -and -not $task.AsyncResult.IsCompleted) {
                $task.PowerShell.Stop()
            }
            if ($task.PowerShell) {
                $task.PowerShell.Dispose()
            }
            if ($task.Runspace) {
                $task.Runspace.Close()
                $task.Runspace.Dispose()
            }
        }
        catch {
            continue
        }
    }
})

[void]$window.ShowDialog()
