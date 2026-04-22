# Windows Network Speed Monitor

Small Windows desktop monitor for live network traffic and active TCP connections.

## Features

- Live incoming, outgoing, and combined network speeds
- Current, minimum, and maximum rates since launch
- Total bytes received and sent
- Active network adapters table
- Established TCP connections with process name, remote IP, remote port, and best-effort hostname lookup
- Sortable tables that keep their sort order during refresh
- Toggle between `Mbps` and `MBps`

## Files

- `Start_Network_Monitor.bat`: launcher
- `scripts/NetworkMonitor.ps1`: main app

## Run

```powershell
.\Start_Network_Monitor.bat
```

If PowerShell execution policy blocks the script, the launcher already uses `-ExecutionPolicy Bypass` for this app only.

## Notes

- Hostnames come from reverse DNS when available.
- Exact website URLs are not available from normal Windows socket data, so the app shows remote IPs and hostnames rather than full pages.
- The app uses built-in Windows PowerShell networking cmdlets such as `Get-NetAdapterStatistics` and `Get-NetTCPConnection`.
