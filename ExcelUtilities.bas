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

Attribute VB_Name = "Excelutilities"
Option Compare Database

'Check whether specified worksheet exists or not in specified workbook
Public Function WorkSheetExist(oWb As Workbook, SheetName As String) As Boolean
    WorkSheetExist = False
    
    Dim ws As Worksheet

    For Each ws In oWb.Worksheets
        If SheetName = ws.Name Then
            WorkSheetExist = True
            Exit For
        End If
    Next ws
    
End Function

'Convert Column Number To Column Letter
Public Function ColumnLetter(oWs As Worksheet, Col As Long) As String
     '-----------------------------------------------------------------
    Dim sColumn As String
    On Error Resume Next
    sColumn = Split(oWs.Columns(Col).Address(, False), ":")(1)
    On Error GoTo 0
    ColumnLetter = sColumn
End Function

Public Function LinkToWorksheetInWorkbook(Wb_path As String, SheetNameList As Variant, Optional SheetNameLocalList As Variant, Optional HasFieldNames As Boolean = True) As String
    On Error GoTo Err_LinkToWorksheetInWorkbook

    Dim FailedReason As String

    If Len(Dir(Wb_path)) = 0 Then
        FailedReason = Wb_path
        GoTo Exit_LinkToWorksheetInWorkbook
    End If
    

    If VarType(SheetNameLocalList) <> vbArray + vbVariant Then
        SheetNameLocalList = SheetNameList
    End If


    Dim FullNameList() As Variant
    Dim SheetNameAndRangeList() As Variant
    
    ReDim FullNameList(0 To UBound(SheetNameList))
    ReDim SheetNameAndRangeList(0 To UBound(SheetNameList))

    Dim oExcel As Excel.Application
    Set oExcel = CreateObject("Excel.Application")
    
    With oExcel
        Dim oWb As Workbook
        Set oWb = .Workbooks.Open(Filename:=Wb_path)

        With oWb
            Dim SheetNameIdx As Integer
            Dim SheetName As String
            Dim FullName As String
            Dim ShtColCnt As Long
            Dim col_idx As Long
            Dim SheetNameAndRange As String
            
            For SheetNameIdx = 0 To UBound(SheetNameList)
                SheetName = SheetNameList(SheetNameIdx)
                DelTable (SheetNameLocalList(SheetNameIdx))
                  
                On Error Resume Next
                .Worksheets(SheetName).Activate
                On Error GoTo Next_SheetNameIdx_1
                
                With .ActiveSheet.UsedRange
                    ShtColCnt = .Columns.count
                    
                    If HasFieldNames = True Then
                        For col_idx = 1 To ShtColCnt
                            If IsEmpty(.Cells(1, col_idx)) = True Then
                                ShtColCnt = col_idx - 1
                                Exit For
                            End If
                        Next col_idx
                    End If
                    
                    SheetNameAndRange = SheetName & "!A1:" & ColumnLetter(oWb.ActiveSheet, ShtColCnt) & .Rows.count
                    
                End With '.ActiveSheet.UsedRange
                
                FullNameList(SheetNameIdx) = .FullName
                SheetNameAndRangeList(SheetNameIdx) = SheetNameAndRange
                
Next_SheetNameIdx_1:
            Next SheetNameIdx

            .Close False
            
        End With 'oWb
        
        .Quit
        
    End With 'oExcel
    

    For SheetNameIdx = 0 To UBound(SheetNameList)
        SheetName = SheetNameLocalList(SheetNameIdx)
        FullName = FullNameList(SheetNameIdx)
        SheetNameAndRange = SheetNameAndRangeList(SheetNameIdx)

        On Error Resume Next
        DoCmd.TransferSpreadsheet acLink, , SheetName, FullName, True, SheetNameAndRange
        On Error GoTo Next_SheetNameIdx_2
        
Next_SheetNameIdx_2:
    Next SheetNameIdx
    
    On Error GoTo Err_LinkToWorksheetInWorkbook
    
    
Exit_LinkToWorksheetInWorkbook:
    LinkToWorksheetInWorkbook = FailedReason
    Exit Function

Err_LinkToWorksheetInWorkbook:
    FailedReason = Err.Description
    Resume Exit_LinkToWorksheetInWorkbook
End Function

'Export a table to one or more worksheets in case row count over 65535
Public Function ExportTblToSht(oExcel As Object, Wb_path, Tbl_name As String, Sht_name As String) As String
    On Error GoTo Err_ExportTblToSht
    
    Dim FailedReason As String
    FailedReason = ""

    If Len(Dir(Wb_path)) = 0 Or TableExist(Tbl_name) = False Then
        FailedReason = Tbl_name & " does not exist"
        GoTo Exit_ExportTblToSht
    End If

    
    Dim MaxRowPerSht As Long
    Dim RecordCount As Long
        
    MaxRowPerSht = 65534
    RecordCount = Table_RecordCount(Tbl_name)
    
    
    If RecordCount <= 0 Then
        GoTo Exit_ExportTblToSht
    
    ElseIf RecordCount <= MaxRowPerSht Then
        DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, Tbl_name, Wb_path, True, Sht_name
    
    Else
        Dim Tbl_COPY_name As String
        Tbl_COPY_name = Tbl_name & "_COPY"
        
        DelTable (Tbl_COPY_name)
        
        Dim SQL_cmd As String

        SQL_cmd = "SELECT * " & vbCrLf & _
                    "INTO [" & Tbl_COPY_name & "] " & vbCrLf & _
                    "FROM [" & Tbl_name & "]" & vbCrLf & _
                    ";"
        
        RunSQL_CmdWithoutWarning (SQL_cmd)
        
        SQL_cmd = "ALTER TABLE [" & Tbl_COPY_name & "] " & vbCrLf & _
                    "ADD record_idx COUNTER " & vbCrLf & _
                    ";"
        
        RunSQL_CmdWithoutWarning (SQL_cmd)
        
        
        Dim ShtCount As Integer
        Dim sht_idx As Integer
        Dim Sht_part_name As String
        Dim Tbl_part_name As String
        
        ShtCount = Int(RecordCount / MaxRowPerSht)
        
        For sht_idx = 0 To ShtCount
            Sht_part_name = Sht_name
            
            If sht_idx > 0 Then
                Sht_part_name = Sht_part_name & "_" & sht_idx
            End If
                
            Tbl_part_name = Tbl_name & "_" & sht_idx
            
            DelTable (Tbl_part_name)
            
            SQL_cmd = "SELECT * " & vbCrLf & _
                        "INTO [" & Tbl_part_name & "] " & vbCrLf & _
                        "FROM [" & Tbl_COPY_name & "]" & vbCrLf & _
                        "WHERE [record_idx] >= " & sht_idx * MaxRowPerSht + 1 & vbCrLf & _
                        "AND [record_idx] <= " & (sht_idx + 1) * MaxRowPerSht & vbCrLf & _
                        ";"
        
            'MsgBox SQL_cmd
            RunSQL_CmdWithoutWarning (SQL_cmd)
        
        
            SQL_cmd = "ALTER TABLE [" & Tbl_part_name & "] " & vbCrLf & _
                        "DROP COLUMN [record_idx] " & vbCrLf & _
                        ";"
    
            RunSQL_CmdWithoutWarning (SQL_cmd)
            
            
            DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, Tbl_part_name, Wb_path, True, Sht_part_name
            
            DelTable (Tbl_part_name)
            
        Next sht_idx
    
        DelTable (Tbl_COPY_name)
        
    End If

    
Exit_ExportTblToSht:
    ExportTblToSht = FailedReason
    Exit Function

Err_ExportTblToSht:
    FailedReason = Err.Description
    Resume Exit_ExportTblToSht
    
End Function

'Replace String in a range of a worksheet that enclose any excel error in a function
Public Function ReplaceStrInWsRng(oWsRng As Range, What As Variant, Replacement As Variant, Optional LookAt As Variant, Optional SearchOrder As Variant, Optional MatchCase As Variant, Optional MatchByte As Variant, Optional SearchFormat As Variant, Optional ReplaceFormat As Variant) As String
    On Error GoTo Err_ReplaceStrInWsRng
    
    Dim FailedReason As String
    
    With oWsRng
        .Application.DisplayAlerts = False

        .Replace What, Replacement, LookAt, SearchOrder, MatchCase, MatchByte, SearchFormat, ReplaceFormat

        .Application.DisplayAlerts = True
        
    End With '.oWsRng


Exit_ReplaceStrInWsRng:
    ReplaceStrInWsRng = FailedReason
    Exit Function

Err_ReplaceStrInWsRng:
    FailedReason = Err.Description
    Resume Exit_ReplaceStrInWsRng
    
End Function
