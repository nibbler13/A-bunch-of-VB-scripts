on error resume next
Set wshShell = Wscript.CreateObject("Wscript.Shell")
strSystemDrive = wshShell.ExpandEnvironmentStrings("%SYSTEMDRIVE%")
Set objFSO = CreateObject("Scripting.FileSystemObject")

If objFSO.FileExists(strSystemDrive & "\Temp\DefaultPrinterName.txt") Then
	set file = objFSO.OpenTextFile(strSystemDrive & "\Temp\DefaultPrinterName.txt")
	defaultPrinter = file.ReadAll
	strComputer = "."
	printerFinded = 0
	Set objWMIService = GetObject ("winmgmts:\\" & strComputer & "\root\cimv2")
	Set colPrinters = objWMIService.ExecQuery ("Select * From Win32_Printer")
	For Each objPrinter in colPrinters
		if objPrinter.Name = defaultPrinter then
			Set WshNetwork = WScript.CreateObject("WScript.Network") 
			WshNetwork.SetDefaultPrinter defaultPrinter
			printerFinded = 1
		end if
	Next
	
	if printerFinded = 0 and InStr(defaultPrinter, "\\") <> 0 then
		Set WshNetwork = WScript.CreateObject("WScript.Network")
		WshNetwork.AddWindowsPrinterConnection defaultPrinter
		WshNetwork.SetDefaultPrinter defaultPrinter
	end if
End If