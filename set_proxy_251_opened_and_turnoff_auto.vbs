On Error Resume Next

Set oReg   = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
sKeyPath   = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\Connections"
sValueName = "DefaultConnectionSettings"

' Get registry value where each byte is a different setting.
oReg.GetBinaryValue &H80000001, sKeyPath, sValueName, bValue

' Check byte to see if detect is currently on.
If (bValue(8) And 8) = 8 Then

  ' Turn off detect and write back settings value.
  bValue(8) = bValue(8) And Not 8
  oReg.SetBinaryValue &H80000001, sKeyPath, sValueName, bValue

End If

Set oReg = Nothing

Dim objShell, RegLocate

On Error Resume Next

Set objShell = WScript.CreateObject("WScript.Shell")
On Error Resume Next

RegLocate = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyServer"
objShell.RegWrite RegLocate,"172.16.166.251:3128","REG_SZ"

RegLocate = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable"
objShell.RegWrite RegLocate,"1","REG_DWORD"

RegLocate = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyOverride"
objShell.RegWrite RegLocate,"172.*.*.*;<local>","REG_SZ"

WScript.Quit