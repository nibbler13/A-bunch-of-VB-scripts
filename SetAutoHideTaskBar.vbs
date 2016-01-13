Option Explicit
Const HKCU = &H80000001
Dim objReg
Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}root\default:StdRegProv")
Dim objWMI
Set objWMI = GetObject("winmgmts:{impersonationLevel=impersonate}root\cimv2")
' Adjust the first bit of the taskbar settings
Dim arrVal()
objReg.GetBinaryValue HKCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\StuckRects2", "Settings", arrVal
arrVal(8) = (arrVal(8) AND &h07) OR &h01
objReg.SetBinaryValue HKCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\StuckRects2", "Settings", arrVal
' Restart Explorer for the settings to take effect.
Dim objProcess, colProcesses
Set colProcesses = objWMI.ExecQuery("Select * from Win32_Process Where Name='explorer.exe'")
For Each objProcess In colProcesses
objProcess.Terminate()
Next