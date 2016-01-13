set fso = createobject("Scripting.FileSystemObject")   
set rootfolder = fso.getfolder("\\172.16.166.3\mail")	
set oShell = WScript.CreateObject ("WScript.Shell")
desktop = oShell.SpecialFolders("desktop")
APPDATA = oShell.SpecialFolders("APPDATA")
     dst_folder = "c:\temp\"
fso.copyfile rootfolder & "\*.*" , dst_folder, 1
for each rf in rootfolder.subfolders
  on error resume next
  call fso.copyfolder(rf, dst_folder & "\" & rf.name)
next

Set oWsh = CreateObject("Wscript.Shell")
oWsh.Run("c:\temp\outauto.exe")