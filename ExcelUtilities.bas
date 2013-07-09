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
