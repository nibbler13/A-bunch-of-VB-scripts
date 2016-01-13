on error resume next
Set wshShell = Wscript.CreateObject("Wscript.Shell")
strUserName = wshShell.ExpandEnvironmentStrings("%USERNAME%")

If strUserName = "nn-admin" Then
	strSystemDrive = wshShell.ExpandEnvironmentStrings("%SYSTEMDRIVE%")
	strComputer = "."
	Set objWMIService = GetObject ("winmgmts:\\" & strComputer & "\root\cimv2")
	Set colPrinters = objWMIService.ExecQuery ("Select * From Win32_Printer")
	For Each objPrinter in colPrinters
		If objPrinter.Attributes And 4 Then 
			Set objFSO = CreateObject("Scripting.FileSystemObject")
			if not objFSO.FolderExists(strSystemDrive & "\Temp") then
				objFSO.CreateFolder (strSystemDrive & "\Temp")
			end if
			outFile = strSystemDrive & "\Temp\DefaultPrinterName.txt"
			Set objFile = objFSO.CreateTextFile(outFile, true)
			objFile.Write objPrinter.Name
			objFile.Close
		End If
	Next
End If