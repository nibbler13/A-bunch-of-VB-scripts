Dim objShell, RegLocate

On Error Resume Next

Set objShell = WScript.CreateObject("WScript.Shell")
On Error Resume Next

RegLocate = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyServer"
objShell.RegWrite RegLocate,"127.0.0.1:8080","REG_SZ"

RegLocate = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable"
objShell.RegWrite RegLocate,"0","REG_DWORD"

WScript.Quit