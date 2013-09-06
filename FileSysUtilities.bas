'/***************************************************************************
'         VBA Utilities
'                             -------------------
'    begin                : 2013-07-23
'    copyright            : (C) 2013 by Walter Tsui
'    email                : waltertech426@gmail.com
' ***************************************************************************/

'/***************************************************************************
' *                                                                         *
' *   This program is free software; you can redistribute it and/or modify  *
' *   it under the terms of the GNU General Public License as published by  *
' *   the Free Software Foundation; either version 2 of the License, or     *
' *   (at your option) any later version.                                   *
' *                                                                         *
' ***************************************************************************/

Attribute VB_Name = "FileSysUtilities"
Option Compare Database

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Const ZipTool_local_path = "\7za\7za"

'Check whether a file exists
Function FileExists(ByVal strFile As String, Optional bFindFolders As Boolean) As Boolean
    'Purpose:   Return True if the file exists, even if it is hidden.
    'Arguments: strFile: File name to look for. Current directory searched if no path included.
    '           bFindFolders. If strFile is a folder, FileExists() returns False unless this argument is True.
    'Note:      Does not look inside subdirectories for the file.
    'Author:    Allen Browne. http://allenbrowne.com June, 2006.
    Dim lngAttributes As Long

    'Include read-only files, hidden files, system files.
    lngAttributes = (vbReadOnly Or vbHidden Or vbSystem)

    If bFindFolders Then
        lngAttributes = (lngAttributes Or vbDirectory) 'Include folders as well.
    Else
        'Strip any trailing slash, so Dir does not look inside the folder.
        Do While Right$(strFile, 1) = "\"
            strFile = Left$(strFile, Len(strFile) - 1)
        Loop
    End If

    'If Dir() returns something, the file exists.
    On Error Resume Next
    FileExists = (Len(Dir(strFile, lngAttributes)) > 0)
End Function

'Copy File without error msg
Public Sub CopyFileBypassErr(Src As String, des As String)
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    'object.copyfile,source,destination,file overright(True is default)
    objFSO.CopyFile Src, des, True
    
    Set objFSO = Nothing
End Sub

'Unzip multiple files in directory
Public Function ExtractZipInDir(SrcDir As String, DesDir As String, Optional Criteria As String = "", Optional DeleteZipFile As Boolean = False) As String
    On Error GoTo Err_ExtractZip
    
    Dim FailedReason As String
    
    Dim Result As String
    
    Criteria = SrcDir & Criteria
    Result = Dir(Criteria)
    
    
    Do While Len(Result) > 0
        Call ExtractZip(SrcDir & Result, DesDir, DeleteZipFile)
        Result = Dir
    Loop

Exit_ExtractZip:
    ExtractZipInDir = FailedReason
    Exit Function

Err_ExtractZip:
    FailedReason = Err.Description
    Resume Exit_ExtractZip

End Function

'Unzip a file
Public Function ExtractZip(Src As String, DesDir As String, Optional DeleteZipFile As Boolean = False) As String
    On Error GoTo Err_ExtractZip
    
    Dim FailedReason As String
    
    Dim ZipTool_path As String
    ZipTool_path = [CurrentProject].[Path] & ZipTool_local_path
    
    Dim ShellCmd As String
    Dim Success As Boolean

    
    ShellCmd = ZipTool_path & " x " & Src & " -o" & DesDir & " -ry"
    'MsgBox ShellCmd
    Success = ShellAndWait(ShellCmd, vbHide)

    If Success = True And DeleteZipFile = True Then
        Kill Src
    End If

Exit_ExtractZip:
    ExtractZip = FailedReason
    Exit Function

Err_ExtractZip:
    FailedReason = Err.Description
    Resume Exit_ExtractZip

End Function

'Ftp upload file
Public Function FTPUpload(sSite, sUsername, sPassword, sLocalFile, sRemotePath, Optional Delay As Integer = 1000) As String
    'This script is provided under the Creative Commons license located
    'at http://creativecommons.org/licenses/by-nc/2.5/ . It may not
    'be used for commercial purposes with out the expressed written consent
    'of NateRice.com
    
    Const OpenAsDefault = -2
    Const FailIfNotExist = 0
    Const ForReading = 1
    Const ForWriting = 2
    
    Dim oFTPScriptFSO As Object
    Dim oFTPScriptShell As Object
    
    Set oFTPScriptFSO = CreateObject("Scripting.FileSystemObject")
    Set oFTPScriptShell = CreateObject("WScript.Shell")
    
    sRemotePath = Trim(sRemotePath)
    sLocalFile = Trim(sLocalFile)
    
    '----------Path Checks---------
    'Here we willcheck the path, if it contains
    'spaces then we need to add quotes to ensure
    'it parses correctly.
    If InStr(sRemotePath, " ") > 0 Then
        If Left(sRemotePath, 1) <> """" And Right(sRemotePath, 1) <> """" Then
            sRemotePath = """" & sRemotePath & """"
        End If
    End If
    
    If InStr(sLocalFile, " ") > 0 Then
        If Left(sLocalFile, 1) <> """" And Right(sLocalFile, 1) <> """" Then
            sLocalFile = """" & sLocalFile & """"
        End If
    End If
    
    'Check to ensure that a remote path was
    'passed. If it's blank then pass a "\"
    If Len(sRemotePath) = 0 Then
        'Please note that no premptive checking of the
        'remote path is done. If it does not exist for some
        'reason. Unexpected results may occur.
        sRemotePath = "\"
    End If
    
    'Check the local path and file to ensure
    'that either the a file that exists was
    'passed or a wildcard was passed.
    If InStr(sLocalFile, "*") Then
        If InStr(sLocalFile, " ") Then
            FTPUpload = "Error: Wildcard uploads do not work if the path contains a " & _
                        "space." & vbCrLf
            FTPUpload = FTPUpload & "This is a limitation of the Microsoft FTP client."
            Exit Function
        End If
    ElseIf Len(sLocalFile) = 0 Or Not oFTPScriptFSO.FileExists(sLocalFile) Then
        'nothing to upload
        FTPUpload = "Error: File Not Found."
        Exit Function
    End If
    '--------END Path Checks---------
    
    'build input file for ftp command
    Dim sFTPScript As String
    
    sFTPScript = sFTPScript & "USER " & sUsername & vbCrLf
    sFTPScript = sFTPScript & sPassword & vbCrLf
    sFTPScript = sFTPScript & "cd " & sRemotePath & vbCrLf
    sFTPScript = sFTPScript & "binary" & vbCrLf
    sFTPScript = sFTPScript & "prompt n" & vbCrLf
    sFTPScript = sFTPScript & "put " & sLocalFile & vbCrLf
    sFTPScript = sFTPScript & "quit" & vbCrLf & "quit" & vbCrLf & "quit" & vbCrLf
    
    
    Dim sFTPTemp As String
    Dim sFTPTempFile As String
    Dim sFTPResults As String
    
    sFTPTemp = oFTPScriptShell.ExpandEnvironmentStrings("%TEMP%")
    sFTPTempFile = sFTPTemp & "\" & oFTPScriptFSO.GetTempName
    sFTPResults = sFTPTemp & "\" & oFTPScriptFSO.GetTempName
    
    'Write the input file for the ftp command
    'to a temporary file.
    Dim fFTPScript As Object
    
    Set fFTPScript = oFTPScriptFSO.CreateTextFile(sFTPTempFile, True)
    fFTPScript.WriteLine (sFTPScript)
    fFTPScript.Close
    Set fFTPScript = Nothing
    
    oFTPScriptShell.Run "%comspec% /c FTP -n -s:" & sFTPTempFile & " " & sSite & _
                        " > " & sFTPResults, 0, True
    
    Sleep Delay
    
    'Check results of transfer.
    Dim fFTPResults As Object
    
    Set fFTPResults = oFTPScriptFSO.OpenTextFile(sFTPResults, ForReading, _
    FailIfNotExist, OpenAsDefault)
    
    Dim sResults As String
    sResults = fFTPResults.ReadAll
    
    fFTPResults.Close
    
    If InStr(sResults, "226 Transfer complete.") > 0 Then
        FTPUpload = ""
    ElseIf InStr(sResults, "File not found") > 0 Then
        FTPUpload = "Error: File Not Found"
    ElseIf InStr(sResults, "cannot log in.") > 0 Then
        FTPUpload = "Error: Login Failed."
    Else
        FTPUpload = "Error: Unknown."
    End If
    
    oFTPScriptFSO.DeleteFile (sFTPTempFile)
    oFTPScriptFSO.DeleteFile (sFTPResults)
    
    Set oFTPScriptFSO = Nothing
    
    oFTPScriptShell.CurrentDirectory = sOriginalWorkingDirectory
    Set oFTPScriptShell = Nothing
    
End Function

'Ftp download file
Function FTPDownload(sSite, sUsername, sPassword, sLocalPath, sRemotePath, sRemoteFile, Optional Delay As Integer = 1000) As String
    Const OpenAsDefault = -2
    Const FailIfNotExist = 0
    Const ForReading = 1
    Const ForWriting = 2
    
    Dim oFTPScriptFSO As Object
    Dim oFTPScriptShell As Object
    
    Set oFTPScriptFSO = CreateObject("Scripting.FileSystemObject")
    Set oFTPScriptShell = CreateObject("WScript.Shell")
    
    
    sRemotePath = Trim(sRemotePath)
    sLocalPath = Trim(sLocalPath)
    
    '----------Path Checks---------
    If InStr(sRemotePath, " ") > 0 Then
        If Left(sRemotePath, 1) <> """" And Right(sRemotePath, 1) <> """" Then
            sRemotePath = """" & sRemotePath & """"
        End If
    End If
    
    If Len(sRemotePath) = 0 Then
        sRemotePath = "\"
    End If
    
    'If the local path was blank. Pass the current working direcory.
    If Len(sLocalPath) = 0 Then
        sLocalPath = oFTPScriptShell.CurrentDirectory
    End If
    
    If Not oFTPScriptFSO.FolderExists(sLocalPath) Then
        'destination not found
        FTPDownload = "Error: Local Folder Not Found."
      Exit Function
    End If
    
    Dim sOriginalWorkingDirectory As String
    sOriginalWorkingDirectory = oFTPScriptShell.CurrentDirectory
    oFTPScriptShell.CurrentDirectory = sLocalPath
    '--------END Path Checks---------
    
    'build input file for ftp command
    Dim sFTPScript As String
    sFTPScript = ""
    
    sFTPScript = sFTPScript & "USER " & sUsername & vbCrLf
    sFTPScript = sFTPScript & sPassword & vbCrLf
    sFTPScript = sFTPScript & "cd " & sRemotePath & vbCrLf
    sFTPScript = sFTPScript & "binary" & vbCrLf
    sFTPScript = sFTPScript & "prompt n" & vbCrLf
    sFTPScript = sFTPScript & "mget " & sRemoteFile & vbCrLf
    sFTPScript = sFTPScript & "quit" & vbCrLf & "quit" & vbCrLf & "quit" & vbCrLf
    
    
    Dim sFTPTemp As String
    Dim sFTPTempFile As String
    Dim sFTPResults As String
    
    sFTPTemp = oFTPScriptShell.ExpandEnvironmentStrings("%TEMP%")
    sFTPTempFile = sFTPTemp & "\" & oFTPScriptFSO.GetTempName
    sFTPResults = sFTPTemp & "\" & oFTPScriptFSO.GetTempName
    
    'Write the input file for the ftp command to a temporary file.
    Dim fFTPScript As Object
    
    Set fFTPScript = oFTPScriptFSO.CreateTextFile(sFTPTempFile, True)
    fFTPScript.WriteLine (sFTPScript)
    fFTPScript.Close
    Set fFTPScript = Nothing
    
    oFTPScriptShell.Run "%comspec% /c FTP -n -s:" & sFTPTempFile & " " & sSite & _
    " > " & sFTPResults, 0, True
    
    Sleep Delay

    
    'Check results of transfer.
    Dim fFTPResults As Object
    Set fFTPResults = oFTPScriptFSO.OpenTextFile(sFTPResults, ForReading, _
                      FailIfNotExist, OpenAsDefault)
                      
    Dim sResults As String
    sResults = fFTPResults.ReadAll
    fFTPResults.Close
    
    If InStr(sResults, "226 Transfer complete.") > 0 Then
        FTPDownload = ""
    ElseIf InStr(sResults, "File not found") > 0 Then
        FTPDownload = "Error: File Not Found"
    ElseIf InStr(sResults, "cannot log in.") > 0 Then
        FTPDownload = "Error: Login Failed."
    Else
        FTPDownload = "Error: Unknown."
    End If
    
    oFTPScriptFSO.DeleteFile (sFTPTempFile)
    oFTPScriptFSO.DeleteFile (sFTPResults)
    
    Set oFTPScriptFSO = Nothing
    
    oFTPScriptShell.CurrentDirectory = sOriginalWorkingDirectory
    Set oFTPScriptShell = Nothing
    
End Function

'Count Row Number of a text file
Public Function CountRowsInText(file_name As String) As Long
    On Error GoTo Err_CountRowsInText
    
    Dim fso As Object
    Dim File As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set File = fso.OpenTextFile(file_name, 1)


    Dim RowCnt As Long
    Dim str_line As String
    
    RowCnt = 0

    Do Until File.AtEndOfStream = True
        RowCnt = RowCnt + 1
        str_line = File.ReadLine
        
    Loop


    File.Close


Exit_CountRowsInText:
    CountRowsInText = RowCnt
    Exit Function

Err_CountRowsInText:
    RowCnt = -1
    Call ShowMsgBox(Err.Description)
    Resume Exit_CountRowsInText

End Function

'Split a Text File into multiple text files of specified row count(default: 65535)
Public Function SplitTextFile(src As String, Optional des_fmt As String, Optional RowCntPerFile As Long = 65535, Optional file_idx_start As Integer = 0, Optional NumOfHdrRows As Long = 0, Optional DeleteSrc As Boolean = False) As String
    On Error GoTo Err_SplitTextFile
    
    Dim FailedReason As String


    If Len(Dir(src)) = 0 Then
        FailedReason = src
        GoTo Exit_SplitTextFile
    End If

    If RowCntPerFile < NumOfHdrRows + 1 Then
        FailedReason = "RowCntPerFile < NumOfHdrRows + 1"
        GoTo Exit_SplitTextFile
    End If


    'if no need to split, return
    Dim RowCnt_src As Long
    RowCnt_src = CountRowsInText(src)
    
    If RowCnt_src <= RowCntPerFile Then
        GoTo Exit_SplitTextFile
    End If
    
    
    'Check whether there exists files which name is same to the splitted files
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim des_dir As String
    Dim des_name As String
    Dim des_ext As String
    Dim des_path As String
    
    des_dir = fso.GetParentFolderName(src)
    des_name = fso.GetFileName(src)
    des_ext = fso.GetExtensionName(src)

    If des_fmt = "" Then
        des_fmt = Left(des_name, Len(des_name) - Len("." & des_ext)) & "_*"
    End If

    
    Dim NumOfSplit As Integer
    
    If RowCnt_src <= RowCntPerFile Then
        NumOfSplit = 0
    Else
        NumOfSplit = Int((RowCnt_src - RowCntPerFile) / (RowCntPerFile + 1 - NumOfHdrRows)) + 1
    End If
    
    
    Dim file_idx_end As Integer
    file_idx_end = file_idx_start + NumOfSplit 'Int(RowCnt_src / (RowCntPerFile + 1 - NumOfHdrRows))
    
    
    Dim file_idx As Integer
    
    For file_idx = file_idx_start To file_idx_end
        des_path = des_dir & "\" & Replace(des_fmt, "*", str(file_idx)) & "." & des_ext
        
        If Len(Dir(des_path)) > 0 Then
            Exit For
        End If
        
    Next file_idx
    
    
    If Len(Dir(des_path)) > 0 Then
        FailedReason = des_path
        GoTo Exit_SplitTextFile
    End If
    
    
    
    'Obtain header rows for later files and create the first splitted file
    Dim File_src As Object
    Dim FileNum_des As Integer
    Dim str_line As String
    Dim HdrRows As String

    Set File_src = fso.OpenTextFile(src, 1)
    des_path = des_dir & "\" & Replace(des_fmt, "*", str(file_idx_start)) & "." & des_ext
    FileNum_des = FreeFile
    Open des_path For Output As #FileNum_des
    
    RowCnt = 0
    
    Do Until RowCnt >= NumOfHdrRows Or File_src.AtEndOfStream = True
        RowCnt = RowCnt + 1
        str_line = File_src.ReadLine
        Print #FileNum_des, str_line
        HdrRows = HdrRows & str_line
        
    Loop

    
    Do Until RowCnt >= RowCntPerFile Or File_src.AtEndOfStream = True
        RowCnt = RowCnt + 1
        Print #FileNum_des, File_src.ReadLine
        
    Loop
    
    Close #FileNum_des
    

    'Start to split
    For file_idx = file_idx_start + 1 To file_idx_end
        If File_src.AtEndOfStream = True Then
            Exit For
        End If
            
        des_path = des_dir & "\" & Replace(des_fmt, "*", str(file_idx)) & "." & des_ext
        FileNum_des = FreeFile
        Open des_path For Output As #FileNum_des

        RowCnt = NumOfHdrRows
        Print #FileNum_des, HdrRows

        Do Until RowCnt >= RowCntPerFile Or File_src.AtEndOfStream = True
            RowCnt = RowCnt + 1
            Print #FileNum_des, File_src.ReadLine
            
        Loop

        Close #FileNum_des
        
    Next file_idx

    
    File_src.Close


    If DeleteSrc = True Then
        Kill src
    End If
    
    
Exit_SplitTextFile:
    SplitTextFile = FailedReason
    Exit Function

Err_SplitTextFile:
    FailedReason = Err.Description
    Resume Exit_SplitTextFile
    
End Function

'Delete rows in a text file
Public Function DeleteRowInText(file_name As String, StartRow As Long, EndRow As Long) As String
    On Error GoTo Err_DeleteRowInText
    
    Dim FailedReason As String
    
    If EndRow < StartRow Then
        EndRow = StartRow
    End If
    
    Dim temp_file_name As String
    temp_file_name = file_name & "_temp"
    
    On Error Resume Next
    Kill temp_file_name
    On Error GoTo Err_DeleteRowInText
    
    Dim temp_file_PortNum As Integer
    temp_file_PortNum = FreeFile
    
    Open temp_file_name For Output As #temp_file_PortNum
    
    Dim fso As Object
    Dim File As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set File = fso.OpenTextFile(file_name, 1)

    Dim row As Long
    Dim str_line As String
    
    row = 0

    Do Until File.AtEndOfStream = True 'EOF(2)
        row = row + 1
        
        str_line = File.ReadLine
    
        If row >= StartRow And row <= EndRow Then
            GoTo Loop_DeleteRowInText_1
        End If
        
        Print #temp_file_PortNum, str_line
        
Loop_DeleteRowInText_1:
    Loop

    File.Close
    
    Close #temp_file_PortNum
    
    Kill file_name
    
    Name temp_file_name As file_name
    

Exit_DeleteRowInText:
    DeleteRowInText = FailedReason
    Exit Function
    
Err_DeleteRowInText:
    FailedReason = Err.Description
    GoTo Exit_DeleteRowInText
    
End Function


'Replace multiple strings in multiple files in a folder
Function ReplaceStrInFolder(folder_name As String, Arr_f As Variant, Arr_r As Variant, Optional StartRow As Long = 0) As String
    On Error GoTo Err_ReplaceStrInFolder
    
    Dim FailedReason As String
    
    Dim file_name As String
        
    file_name = Dir(folder_name & "\")
        
    Do Until file_name = ""
        file_name = folder_name & "\" & file_name
        Call ReplaceStrInFile(file_name, Arr_f, Arr_r, StartRow)
        file_name = Dir()
    Loop

Exit_ReplaceStrInFolder:
    ReplaceStrInFolder = FailedReason
    Exit Function
    
Err_ReplaceStrInFolder:
    FailedReason = Err.Description
    GoTo Exit_ReplaceStrInFolder
    
End Function

'Replace multiple strings in a file
Function ReplaceStrInFile(file_name As String, Arr_f As Variant, Arr_r As Variant, Optional StartRow As Long = 0) As String
    On Error GoTo Err_ReplaceStrInFile
    
    Dim FailedReason As String
    
    Dim temp_file_name As String
    temp_file_name = file_name & "_temp"
    
    On Error Resume Next
    Kill temp_file_name
    On Error GoTo Err_ReplaceStrInFile
    
    Close
    Open temp_file_name For Output As #1
    
    Dim fso As Object
    Dim File As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set File = fso.OpenTextFile(file_name, 1)

    Dim row As Long
    Dim str_line As String
    Dim i As Integer
    Dim str_f As String
    Dim str_r As String
    
    row = 0

    Do Until File.AtEndOfStream = True 'EOF(2)
        row = row + 1
        
        str_line = File.ReadLine
    
        If row < StartRow Then
            GoTo Loop_ReplaceStrInFile_1
        End If
    
        For i = 0 To UBound(Arr_f)
            str_f = Arr_f(i)
            str_r = Arr_r(i)
            
            str_line = Replace(str_line, str_f, str_r)

        Next i
        
Loop_ReplaceStrInFile_1:

        Print #1, str_line
        
    Loop

    File.Close
    
    Close
    
    Kill file_name
    
    Name temp_file_name As file_name
    

Exit_ReplaceStrInFile:
    ReplaceStrInFile = FailedReason
    Exit Function
    
Err_ReplaceStrInFile:
    FailedReason = Err.Description
    GoTo Exit_ReplaceStrInFile
    
End Function