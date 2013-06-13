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
Public Sub LinkToWorksheetInWorkbook(Wb_path As String, SheetNameList As Variant)
    On Error GoTo Err_LinkToWorksheetInWorkbook

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
            Dim SheetNameAndRange As String
            
            For SheetNameIdx = 0 To UBound(SheetNameList)
                SheetName = SheetNameList(SheetNameIdx)
                DelTable (SheetName)
                  
                On Error Resume Next
                .Worksheets(SheetName).Activate
                On Error GoTo Next_SheetNameIdx_1
                
                With .ActiveSheet.UsedRange
                     SheetNameAndRange = SheetName & "!A1:" & ColumnLetter(oWb.ActiveSheet, .Columns.count) & .Rows.count
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
        SheetName = SheetNameList(SheetNameIdx)
        FullName = FullNameList(SheetNameIdx)
        SheetNameAndRange = SheetNameAndRangeList(SheetNameIdx)

        On Error Resume Next
        DoCmd.TransferSpreadsheet acLink, , SheetName, FullName, True, SheetNameAndRange
        On Error GoTo Next_SheetNameIdx_2
        
Next_SheetNameIdx_2:
    Next SheetNameIdx
    
    On Error GoTo Err_LinkToWorksheetInWorkbook
    
Exit_LinkToWorksheetInWorkbook:
    Exit Sub

Err_LinkToWorksheetInWorkbook:
    MsgBox Err.Description
    Resume Exit_LinkToWorksheetInWorkbook
End Sub
