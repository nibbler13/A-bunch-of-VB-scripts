Const ForReading = 1
Const ForWriting = 2

'install_papercut_service.vbs "\\nnkk-fs\PrintLogs\"
strRootNetworkDirectory = "\\nnkk-fs\PrintLogs\"
'\\nnkk-fs\PrintLogs\

'Создаем FSO и получаем имя компьютера
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objNet = CreateObject("WScript.Network")
strCompName = objNet.ComputerName

'Если не существует, то создаем папку на диске C для программы и ее конфига
If Not objFSO.FolderExists("C:\PrintCutLogger") Then
	objFSO.CreateFolder ("C:\PrintCutLogger")
End If

'Если не существует, то создаем папку для отчетов на файл-сервере
If Not objFSO.FolderExists(strRootNetworkDirectory & strCompName) Then
	objFSO.CreateFolder (strRootNetworkDirectory & strCompName)
End If

'Если не существует, то копируем чистый конфиг из дистрибутива и прописываем в него путь хранения
If objFSO.FileExists("C:\PrintCutLogger\papercut-logger.conf") Then
	dim filesys
	Set filesys = CreateObject("Scripting.FileSystemObject")
	filesys.DeleteFile("C:\PrintCutLogger\papercut-logger.conf")
End If

objFSO.CopyFile (strRootNetworkDirectory & "_distrib\papercut-logger.conf"), "C:\PrintCutLogger\"
Set objFIle = objFSO.OpenTextFile("C:\PrintCutLogger\papercut-logger.conf", ForReading)
strText = objFile.ReadAll
objFile.Close
strNewText = Replace(strText, "CsvLogFilePath=", ("CsvLogFilePath=" & strRootNetworkDirectory & strCompName))
Set objFile = objFSO.OpenTextFile("C:\PrintCutLogger\papercut-logger.conf", ForWriting)
objFile.WriteLine strNewText
objFile.Close

'Если не существует, то копируем программу мониторинга печати
If Not objFSO.FileExists("C:\PrintCutLogger\pcpl.exe") Then
	objFSO.CopyFile (strRootNetworkDirectory & "_distrib\pcpl.exe"), "C:\PrintCutLogger\"
End If

'Проверка наличия нужного процесса в системе
Set wmi = GetObject("winmgmts://./root/cimv2")
isServiceInstalled = FALSE
svcQry = "SELECT * from Win32_Service"
Set objOutParams = wmi.ExecQuery(svcQry)
For Each objSvc in objOutParams
	Select Case objSvc.Name
		Case "PCPrintLogger"
			isServiceInstalled = TRUE
	End Select
Next

'Если процесс отсутствует, то устанавливаем его и запускаем
If Not isServiceInstalled Then
	Set objShell = WScript.CreateObject("WScript.Shell")
	objShell.Run """C:\PrintCutLogger\pcpl.exe"" /install", 0, True
	objShell.Run """C:\PrintCutLogger\pcpl.exe"" /config-spooler", 0, True
	Set objShell = Nothing

	Set svc = wmi.Get("Win32_Service.Name='PCPrintLogger'")
	svc.StartService
End If