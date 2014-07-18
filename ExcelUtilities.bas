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

'Link multiple worksheets in workbooks
Public Function LinkToWorksheetInWorkbook(Wb_path As String, ByVal SheetNameList As Variant, Optional ByVal SheetNameLocalList As Variant, Optional ByVal ShtSeriesList As Variant, Optional HasFieldNames As Boolean = True) As String
    On Error GoTo Err_LinkToWorksheetInWorkbook

    Dim FailedReason As String

    If Len(Dir(Wb_path)) = 0 Then
        FailedReason = Wb_path
        GoTo Exit_LinkToWorksheetInWorkbook
    End If
    

    If VarType(SheetNameLocalList) <> vbArray + vbVariant Then
        SheetNameLocalList = SheetNameList
    End If
    

    'Prepare worksheets to be linked.
    If UBound(SheetNameList) <> UBound(SheetNameLocalList) Then
        FailedReason = "No. of elements in SheetNameList and SheetNameLocalList are not equal"
        GoTo Exit_LinkToWorksheetInWorkbook
    End If
    

    'Link worksheets
    Dim FullNameList() As Variant
    Dim SheetNameAndRangeList() As Variant

    Dim oExcel As Excel.Application
    Set oExcel = CreateObject("Excel.Application")
    
    With oExcel
        Dim oWb As Workbook
        Set oWb = .Workbooks.Open(Filename:=Wb_path, ReadOnly:=True)

        With oWb
            'Prepare to link worksheets in series
            If VarType(ShtSeriesList) = vbArray + vbVariant Then
            
                Dim ShtSeries As Variant
                
                Dim ShtSeries_name As String
                Dim ShtSeries_local_name As String
                Dim ShtSeries_start_idx As Integer
                Dim ShtSeries_end_idx As Integer
                
                Dim WsInS_idx As Integer
                Dim WsInS_cnt As Integer
                
                For Each ShtSeries In ShtSeriesList
                    ShtSeries_name = ShtSeries(0)
                    ShtSeries_local_name = ShtSeries(1)
                    ShtSeries_start_idx = ShtSeries(2)
                    ShtSeries_end_idx = ShtSeries(3)
                    
                    If ShtSeries_local_name = "" Then
                        ShtSeries_local_name = ShtSeries_name
                    End If
                    
                    
                    If ShtSeries_end_idx < ShtSeries_start_idx Then
                        ShtSeries_end_idx = .Worksheets.count - 1
                    End If
                    
                    
                    WsInS_cnt = 0
                
                    For WsInS_idx = ShtSeries_start_idx To ShtSeries_end_idx
                        If WorkSheetExist(oWb, Replace(ShtSeries_name, "*", WsInS_idx)) = True Then
                            WsInS_cnt = WsInS_cnt + 1
                        Else
                            Exit For
                        End If
                        
                    Next WsInS_idx
                
    
                    If WsInS_cnt > 0 Then
                        For WsInS_idx = 0 To WsInS_cnt
                            FailedReason = AppendArray(SheetNameList, Array(Replace(ShtSeries_name, "*", WsInS_idx)))
                            FailedReason = AppendArray(SheetNameLocalList, Array(Replace(ShtSeries_local_name, "*", WsInS_idx)))
                        Next WsInS_idx
                        
                    End If
                    
                Next ShtSeries
                
            End If
            
            'Link worksheets
            ReDim FullNameList(0 To UBound(SheetNameList))
            ReDim SheetNameAndRangeList(0 To UBound(SheetNameList))
    
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
        If SheetNameLocalList(SheetNameIdx) <> "" Then
            SheetName = SheetNameLocalList(SheetNameIdx)
        Else
            SheetName = SheetNameList(SheetNameIdx)
        End If
        
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
Public Function ExportTblToSht(Wb_path, Tbl_name As String, sht_name As String) As String
    On Error GoTo Err_ExportTblToSht
    
    Dim FailedReason As String

    If TableExist(Tbl_name) = False Then
        FailedReason = Tbl_name & " does not exist"
        GoTo Exit_ExportTblToSht
    End If

    
    If Len(Dir(Wb_path)) = 0 Then
        Dim oExcel As Excel.Application
        Set oExcel = CreateObject("Excel.Application")
    
        With oExcel
            Dim oWb As Workbook
            Set oWb = .Workbooks.Add
            
            
            With oWb
                .SaveAs Wb_path
                .Close
            End With 'oWb_DailyRpt
            
            .Quit
            
        End With 'oExcel
        
        Set oExcel = Nothing
        
    End If
    
    
    Dim MaxRowPerSht As Long
    Dim RecordCount As Long
        
    MaxRowPerSht = 65534
    RecordCount = Table_RecordCount(Tbl_name)
    
    
    If RecordCount <= 0 Then
        GoTo Exit_ExportTblToSht
    
    ElseIf RecordCount <= MaxRowPerSht Then
        DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, Tbl_name, Wb_path, True, sht_name
    
    Else
        'handle error msg, "File sharing lock count exceeded. Increase MaxLocksPerFile registry entry"
        DAO.DBEngine.SetOption dbMaxLocksPerFile, 40000
        
        
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
            Sht_part_name = sht_name
            
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

'Append Access SQL Object(Table, Query) to Excel worksheet, and activate it 
Public Function AppendSqlObjToAndActivateWs(oWs As Worksheet, SqlObj_name As String, Optional AddBorder As Boolean = False) As String
    On Error GoTo Err_AppendSqlObjToAndActivateWs
    
    Dim FailedReason As String

    If TableExist(SqlObj_name) = False And QueryExist(SqlObj_name) Then
        FailedReason = SqlObj_name & " does not exist!"
        GoTo Exit_AppendSqlObjToAndActivateWs
    End If
    
    With oWs
        'Have to activate the worksheet for copying query with no error!
        .Activate
        
        'Store the new start row
        Dim RowEnd_old As Long
        Dim RowStart_new As Long
        
        RowEnd_old = .UsedRange.Rows.count
        RowStart_new = RowEnd_old + 1
        
        'Append SqlObj_name to the sheet
        
        'Create Recordset object
        Dim rs As DAO.Recordset
        Set rs = CurrentDb.OpenRecordset(SqlObj_name, dbOpenSnapshot)

        .Range("A" & CStr(.UsedRange.Rows.count + 1)).CopyFromRecordset rs, 65534


        'Copy format from previous rows to new rows
        .Range(.Cells(RowEnd_old, 1), .Cells(RowEnd_old, .UsedRange.Columns.count)).Copy
        .Range(.Cells(RowStart_new, 1), .Cells(.UsedRange.Rows.count, .UsedRange.Columns.count)).PasteSpecial Paste:=xlPasteFormats
        

        If AddBorder = True Then
            'Add border at the last row
            With .UsedRange.Rows(.UsedRange.Rows.count).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With '.UsedRange.Rows(.UsedRange.Rows.count).Borders(xlEdgeBottom)
        End If
        
        
    End With 'oWs
    
    
Exit_AppendSqlObjToAndActivateWs:
    AppendSqlObjToAndActivateWs = FailedReason
    Exit Function

Err_AppendSqlObjToAndActivateWs:
    FailedReason = Err.Description
    Resume Exit_AppendSqlObjToAndActivateWs
    
End Function