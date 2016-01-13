Set objFSO = CreateObject("Scripting.FileSystemObject")
objStartFolder = "\\nnkk-fs\doks"

Set objOutputFile = objFSO.CreateTextFile ("C:\report.csv") 
Set objFolder = objFSO.GetFolder(objStartFolder)
'Wscript.Echo objFolder.Path
objOutputFile.WriteLine(objFolder.Path) 
Set colFiles = objFolder.Files
For Each objFile in colFiles
    'Wscript.Echo objFile.Name
	objOutputFile.WriteLine("first file&" & objFile.Path & "&" & objFile.Name & "&" & objFile.Size) 
Next

ShowSubfolders objFSO.GetFolder(objStartFolder)

objOutputFile.Close
Wscript.Echo

Sub ShowSubFolders(Folder)
    For Each Subfolder in Folder.SubFolders
        'Wscript.Echo Subfolder.Path
		objOutputFile.WriteLine("subfolder&" & Subfolder.Path & "&&" & Subfolder.Size) 
        Set objFolder = objFSO.GetFolder(Subfolder.Path)
        Set colFiles = objFolder.Files
        For Each objFile in colFiles
            'Wscript.Echo objFile.Name
			objOutputFile.WriteLine("file&" & objFile.Path & "&" & objFile.Name & "&" & objFile.Size) 
        Next
        'Wscript.Echo
        ShowSubFolders Subfolder
    Next
End Sub