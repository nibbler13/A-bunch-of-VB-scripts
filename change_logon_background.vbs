'----------------------------------------------------------------------------------------------------------------------------
'Script Name : SetLogonWallPaper.vbs
'Author   : Matthew Beattie
'Created   : 09/06/10
'Description : This script sets the Logon WallPaper based on the operating system type.
'----------------------------------------------------------------------------------------------------------------------------
'Initialization Section. Define and Create Global Variables.
'----------------------------------------------------------------------------------------------------------------------------
Option Explicit
Dim objFSO, wshShell, systemPath, fileSpec
On Error Resume Next
  Set objFSO  = CreateObject("Scripting.FileSystemObject")
  Set wshShell = CreateObject("WScript.Shell")
  systemPath  = wshShell.ExpandEnvironmentStrings("%WinDir%")
  If Err.Number <> 0 Then
   WScript.Quit
  End If
On Error Goto 0
'----------------------------------------------------------------------------------------------------------------------------
'Main Processing Section  
'----------------------------------------------------------------------------------------------------------------------------
On Error Resume Next
  ProcessScript
  If Err.Number <> 0 Then
   WScript.Quit
  End If
On Error Goto 0
'----------------------------------------------------------------------------------------------------------------------------
'Name    : ProcessScript -> Primary Function that controls all other script processing.  
'Parameters : None     ->  
'Return   : None     ->  
'----------------------------------------------------------------------------------------------------------------------------
Function ProcessScript
  Dim regKey, version, systemType
  regKey  = "HKLM\Software\Microsoft\Windows NT\CurrentVersion\ProductName"
  version = ReadRegistry(regKey)
  Select Case version
   Case "Windows 7 Professional"
     fileSpec  = systemPath & "\\system32\oobe\info\Backgrounds\backgroundDefault.jpg"
     systemType = 0
	 
	Dim newFolder, fileSys, objShell
	newFolder = systemPath & "\system32\oobe\info"
	set objShell = createobject("Scripting.FileSystemObject")
	set filesys = CreateObject("Scripting.FileSystemObject") 
	if not objShell.FolderExists(newFolder) then
		filesys.CreateFolder(newFolder) 
	end if
	newFolder = systemPath & "\system32\oobe\info\Backgrounds"
	if not objShell.FolderExists(newFolder) then
		filesys.CreateFolder(newFolder) 
	end if 
	If filesys.FileExists("\\172.16.166.3\mail\backgroundDefault.jpg") Then
		fileSys.CopyFile "\\172.16.166.3\mail\backgroundDefault.jpg", newFolder & "\", True
	End If 
	 
	 
   Case "Microsoft Windows XP"
     systemType = 1
     fileSpec  = systemPath & "\Web\Wallpaper\Bliss.bmp"
   Case Else
   msgbox "exit"
     Exit Function
  End Select
  '-------------------------------------------------------------------------------------------------------------------------
  'Set the Logon Wallpaper based on the operating system type.
  '-------------------------------------------------------------------------------------------------------------------------
  If Not SetLogonWallPaper(systemType) Then
   Exit Function
  End If
End Function
'----------------------------------------------------------------------------------------------------------------------------
'Name    : ReadRegistry -> Read the value of a registry key or value.
'Parameters : key     -> Name of the key (ending in "\") or value to read.
'Return   : ReadRegistry -> Value of key or value read from the local registry (blank is not found).
'----------------------------------------------------------------------------------------------------------------------------
Function ReadRegistry(ByVal key)
  Dim result
  If StrComp(Left (key, 4), "HKU\", vbTextCompare) = 0 Then
   Key = "HKEY_USERS" & Mid(key, 4)
  End If
  On Error Resume Next
   ReadRegistry = WshShell.RegRead (key)
   If Err.Number <> 0 Then
     ReadRegistry = ""
   End If
  On Error Goto 0
End Function
'----------------------------------------------------------------------------------------------------------------------------
'Name    : SetLogonWallPaper -> Sets the Logon Wallpaper registry settings based on the operating system type.
'Parameters : systemType    -> Integer identifying the operating system type.
'Return   : SetLogonWallPaper -> Returns True if the LogonWallPaper registry settings were updated otherwise False.
'----------------------------------------------------------------------------------------------------------------------------
Function SetLogonWallPaper(systemType)
  Dim elements, element, regKey, regValue, regType
  SetLogonWallPaper = False
  '-------------------------------------------------------------------------------------------------------------------------
  'Define the registry settings based on the systemType
  '-------------------------------------------------------------------------------------------------------------------------
  Select Case systemType
   Case 0
     elements = Array("HKLM\Software\Microsoft\Windows\CurrentVersion\Authentication\LogonUI\Background\OEMBackground,1,REG_DWORD")
   Case 1
     elements = Array("HKEY_USERS\.Default\Control Panel\Desktop\TileWallpaper,1,REG_SZ", _
             "HKEY_USERS\.Default\Control Panel\Desktop\WallpaperStyle,2,REG_SZ", _
             "HKEY_USERS\.Default\Control Panel\Desktop\Wallpaper," & fileSpec & ",REG_SZ")
   Case Else
     Exit Function
  End Select
  '-------------------------------------------------------------------------------------------------------------------------
  'Configure the registry settings.
  '-------------------------------------------------------------------------------------------------------------------------
  For Each element In elements
   On Error Resume Next
     regKey  = Split(element, ",")(0)
     regValue = Split(element, ",")(1)
     regType = Split(element, ",")(2)

     wshShell.RegWrite regKey, regValue, regType
     If Err.Number <> 0 Then
      Exit Function
     End If
   On Error Goto 0
  Next
  SetLogonWallPaper = True
End Function
'----------------------------------------------------------------------------------------------------------------------------