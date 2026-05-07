# Net-Gamit

<img width="1162" height="806" alt="image" src="https://github.com/user-attachments/assets/780f34a4-99a9-4509-96b1-fc9f55346b44" />



Net-Gamit is a native Windows PowerShell + WPF GUI for destination-node testing, network health checks, WLAN diagnostics, and process/socket review.

## Run

For users who only need to double-click the application, distribute:

```text
dist\Net-Gamit.exe
```

The executable embeds `Net-Gamit.ps1` and starts it with:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File <embedded Net-Gamit script>
```

Double-click `Launch-Net-Gamit.cmd`, or run:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File .\Net-Gamit.ps1
```

## Build The EXE

```powershell
New-Item -ItemType Directory -Force .\dist | Out-Null
& "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe" /nologo /target:winexe /platform:anycpu /optimize+ /reference:System.Windows.Forms.dll /out:.\dist\Net-Gamit.exe /resource:.\Net-Gamit.ps1,NetGamitScript.ps1 .\tools\NetGamitLauncher.cs
```

## What It Collects

- Test Destination Node checks against an IP address, hostname, or URL.
- TCP or UDP target selection with a configurable port.
- Ping, tracert, Test-NetConnection/TcpClient fallback, DNS lookup, HTTP/S response timing, and SSL certificate details.
- Active adapters, IP/subnet/gateway/MAC/DNS/DHCP details.
- WLAN SSID, BSSID, signal, estimated RSSI, radio type, channel, and inferred band from `netsh wlan`.
- Routing table, TCP/UDP sockets, `netstat -ano`, neighbor cache, firewall profiles, WinHTTP proxy, and `ipconfig /all`.
- A cleaner summary-style Network Health Check interpretation that highlights the primary internet path, active adapters, adapter health, speed/duplex, and Wi-Fi signal quality.
- WLAN Diagnostics with netsh WLAN data, wireless adapter PowerShell data, recent WLAN event log review, association/disassociation history, and driver/hardware failure indicators.
- Processes tab with TCP/UDP sockets mapped to owning process names, PIDs, paths, open TCP listening ports, UDP endpoints, and TCP connection age when Windows exposes it.
- A combined interpretation/report that can be exported from the GUI.
