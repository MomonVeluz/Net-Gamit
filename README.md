# Net-Gamit

<img width="1373" height="874" alt="image" src="https://github.com/user-attachments/assets/99b8f51c-cebd-4164-af5b-d01176474a0b" />


Net-Gamit is a native Windows PowerShell + WPF GUI for destination-node testing, network health checks, and WLAN diagnostics.

## Run

Double-click `Launch-Net-Gamit.cmd`, or run:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File .\Net-Gamit.ps1
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
- A combined interpretation/report that can be exported from the GUI.
