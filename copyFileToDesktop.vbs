set fso = createobject("Scripting.FileSystemObject")   
set rootfolder = fso.getfolder("\\172.16.166.1\toDesktopRegistryLab")	
set oShell = WScript.CreateObject ("WScript.Shell")
desktop = oShell.SpecialFolders("desktop")
APPDATA = oShell.SpecialFolders("APPDATA")
     dst_folder = desktop
fso.copyfile rootfolder & "\*.*" , dst_folder
for each rf in rootfolder.subfolders
  on error resume next
  call fso.copyfolder(rf, dst_folder & "\" & rf.name)
next