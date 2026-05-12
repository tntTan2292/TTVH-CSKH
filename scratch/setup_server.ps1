# VNPost Dashboard Server Setup Script
# 1. Open Firewall Ports for Local Network Access
New-NetFirewallRule -DisplayName "VNPost_DOC_8088" -Direction Inbound -LocalPort 8088 -Protocol TCP -Action Allow -ErrorAction SilentlyContinue
New-NetFirewallRule -DisplayName "VNPost_DOC_8010" -Direction Inbound -LocalPort 8010 -Protocol TCP -Action Allow -ErrorAction SilentlyContinue

# 2. Create Startup Shortcut
$startupPath = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup\VNPost_DOC.lnk"
$vbsPath = "d:\Antigravity - Project - TTVH\CSKH\silent_start.vbs"

$ws = New-Object -ComObject WScript.Shell
$s = $ws.CreateShortcut($startupPath)
$s.TargetPath = "wscript.exe"
$s.Arguments = "`"$vbsPath`""
$s.Save()

Write-Host "✅ Server Configuration Complete!"
Write-Host "✅ Firewall Ports 8088 & 8010 are OPEN."
Write-Host "✅ Auto-startup shortcut created in Startup folder."
