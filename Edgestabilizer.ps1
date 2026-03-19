# If Edge is running right now (best option)
(Get-Process msedge -ErrorAction SilentlyContinue | Select-Object -First 1).Path

# Try common locations
Get-Item "C:\Program Files\Microsoft\Edge\Application\msedge.exe" -ErrorAction SilentlyContinue
Get-Item "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" -ErrorAction SilentlyContinue
Get-Item "$env:LOCALAPPDATA\Microsoft\Edge\Application\msedge.exe" -ErrorAction SilentlyContinue

# If Edge is on PATH
where.exe msedge

# Registry (if present)
reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe" /ve