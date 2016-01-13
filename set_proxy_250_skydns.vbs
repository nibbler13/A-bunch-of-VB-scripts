Dim objShell, RegLocate

On Error Resume Next

Set objShell = WScript.CreateObject("WScript.Shell")
On Error Resume Next

RegLocate = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyServer"
objShell.RegWrite RegLocate,"172.16.166.250:3128","REG_SZ"

RegLocate = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable"
objShell.RegWrite RegLocate,"1","REG_DWORD"

RegLocate = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyOverride"
objShell.RegWrite RegLocate,"172.*.*.*;<local>","REG_SZ"

WScript.Quit