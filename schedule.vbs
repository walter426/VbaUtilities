Option Explicit

Dim fso, oD

Set fso = CreateObject("Scripting.FileSystemObject")
Set oD = fso.GetDrive(fso.GetDriveName(WScript.ScriptFullName))

If oD.FreeSpace/1024/1024/1024 < 3 then
	'WScript.Echo oD.FreeSpace/1024/1024/1024
	WScript.Quit
End If


Dim CurrDir_path
CurrDir_path = fso.GetParentFolderName(Wscript.ScriptFullName)

Dim oAccess
Set oAccess = CreateObject("access.application")

With oAccess
	.Visible = True

	Dim DbSet
	DbSet = Array("\sample.mdb")
	
	Dim DbName
	Dim MapTblCount
	Dim FailedReason
	
	For each DbName in DbSet
		Dim Db_path
		Db_path = CurrDir_path & DbName
		
		On Error Resume Next
		.OpenCurrentDatabase CurrDir_path & DbName

		If Err.Number = 0 Then
			Dim RS
			Set RS = .CurrentDb.OpenRecordset("MapTblSet")
			
			With RS
				.MoveFirst
				.MoveLast
				
				MapTblCount = .RecordCount
				
				.Close
			End With
	
			FailedReason = .Run("ST_DLDailyRawData")

			.CloseCurrentDatabase
			
			
			If Err.Number = 0 And FailedReason = "" And MapTblCount >= 0 Then
				Dim MapTblIdx
				
				'Check Loss
				For MapTblIdx = 0 to MapTblCount - 1
					.OpenCurrentDatabase Db_path
					
					If Err.Number <> 0 Then
						Exit For
					End If

					FailedReason = .Run("ST_DataLossCheck_RawData", MapTblIdx)
					.CloseCurrentDatabase
					
					If Err.Number <> 0 or FailedReason <> "" Then
						Exit For
					End If
				Next
				
				
				'Append Database
				If Err.Number = 0 and FailedReason = "" Then
					For MapTblIdx = 0 to MapTblCount - 1
						.OpenCurrentDatabase Db_path
						
						If Err.Number <> 0 Then
							Exit For
						End If

						FailedReason = .Run("ST_ProcessRdAndAppendDb", MapTblIdx)
						.CloseCurrentDatabase

						If Err.Number <> 0 or FailedReason <> "" Then
							Exit For
						End If
					Next
				End If
				
			End If 'Err.Number = 0 And MapTblCount < 0
			
		End If 'Err.Number = 0
		
	Next
	
	.Quit
	
End With 'oAccess

Set oAccess = Nothing


WScript.Quit (0)


Public Function GetSTE(oAccess)
    Set RS = oAccess.CurrentDb.OpenRecordset("ScheduleTaskErr")
			
	With RS
		.MoveFirst
		GetSTE = .Fields("Err").Value
		.Close
	End With
    
End Function