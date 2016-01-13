set fso = createobject("Scripting.FileSystemObject") 	
startFolder = "\\172.16.166.5\toDesktopInfoscreen\"
endFolder = "C:\Temp\"
set objFolder = fso.GetFolder(startFolder)
set colFiles = objFolder.Files
For Each objFile in colFiles
	If Not objFile.Name = "Thumbs.db" Then
		fso.copyfile startFolder & objFile.Name, endFolder
	End If
Next