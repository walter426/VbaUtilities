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

Attribute VB_Name = "ShellUtilities"
Option Compare Database

Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function GlobalFindAtomA Lib "kernel32" (ByVal lpstr As String) As Integer
Declare Function GlobalAddAtomA Lib "kernel32" (ByVal lpstr As String) As Integer
Declare Function GlobalFindAtomW Lib "kernel32" (ByVal lpstr As String) As Integer
Declare Function GlobalAddAtomW Lib "kernel32" (ByVal lpstr As String) As Integer
Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long

Const INFINITE = &HFFFF
Const SYNCHRONIZE = &H100000

'Start a Shell command and wait for it to finish, hiding while it is running.
Public Function ShellAndWait(ByVal cmd As String, _
    ByVal window_style As VbAppWinStyle) As Boolean
    Dim process_id As Long
    Dim process_handle As Long

    ' Start the program.
    On Error GoTo ShellError
    
    ShellAndWait = False
        
    process_id = Shell(cmd, window_style)
    On Error GoTo 0

    ' Wait for the program to finish.
    ' Get the process handle.
    process_handle = OpenProcess(SYNCHRONIZE, 0, process_id)
    
    If process_handle <> 0 Then
        WaitForSingleObject process_handle, INFINITE
        CloseHandle process_handle
        
        ShellAndWait = True
    End If

    Exit Function

ShellError:
    'MsgBox "Error starting task " & _
    '    txtProgram.text & vbCrLf & _
    '    Err.Description, vbOKOnly Or vbExclamation, _
    '    "Error"
    ShellAndWait = False
End Function

'Send multiples shell commands with timeout
Public Function Shell_SendKeysWithTimeout(oshell As Object, CmdTxt As String, Timeout As Integer) As String
    On Error GoTo Err_Shell_SendKeysWithTimeout
    
    Dim FailedReason As String
    
    
    Dim CmdSet As Variant
    CmdSet = SplitStrIntoArray(CmdTxt, Chr(10))
    
    Dim cmd_idx As Integer
    
    For cmd_idx = 0 To UBound(CmdSet)
        If CmdSet(cmd_idx) = "" Then
            GoTo Next_Shell_SendKeysWithTimeout
        End If
        
        
        With oshell
            .SendKeys (CmdSet(cmd_idx) & vbCrLf)
            Sleep Timeout
        End With 'oShell
        
Next_Shell_SendKeysWithTimeout:
    Next cmd_idx


Exit_Shell_SendKeysWithTimeout:
    Shell_SendKeysWithTimeout = FailedReason
    Exit Function

Err_Shell_SendKeysWithTimeout:
    FailedReason = Err.Description
    Resume Exit_Shell_SendKeysWithTimeout
    
End Function


