Option Explicit

'Schdule Task to append Database
Dim fso, oD

Set fso = CreateObject("Scripting.FileSystemObject")
Set oD = fso.GetDrive(fso.GetDriveName(WScript.ScriptFullName))

'Check whether space is enough for appending
If oD.FreeSpace/1024/1024/1024 < 2 then
    WScript.Quit
End If

Dim CurrDir_path
CurrDir_path = fso.GetParentFolderName(Wscript.ScriptFullName)

Dim oAccess
Set oAccess = CreateObject("access.application")

With oAccess
    .Visible = True
	
	On Error Resume Next
    .OpenCurrentDatabase CurrDir_path & "\SampleDb.mdb"
    
	If Err.Number = 0 Then
		.DoCmd.RunMacro "ScheduleTask"
		
		.CloseCurrentDatabase
		
	End If 'Err.Number = 0
	
    .Quit
	
End With 'oAccess

Set oAccess = Nothing


WScript.Quit (0)