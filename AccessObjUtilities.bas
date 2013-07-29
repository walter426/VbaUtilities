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

Attribute VB_Name = "AccessObjUtilities"
Option Compare Database

'Delete Table
Public Sub DelTable(TableName As String)
'Delete table function

    If TableExist(TableName) = True Then
        DoCmd.DeleteObject acTable, TableName
    End If
End Sub

'Delete Table by sub string
Public Sub DelTables_BySubStr(sub_str As String)
    Dim tdf As TableDef
    
    With CurrentDb
        For Each tdf In .TableDefs
            If InStr(tdf.Name, sub_str) > 0 Then
                DoCmd.SetWarnings False
                DoCmd.DeleteObject acTable, tdf.Name
                DoCmd.SetWarnings True
            End If
        Next tdf
    End With 'CurrentDb
End Sub

'Check whether table exists or not
Public Function TableExist(TableName As String) As Boolean
    TableExist = False
    
    Dim tdf As TableDef
    
    With CurrentDb
        For Each tdf In .TableDefs
            If TableName = tdf.Name Then
                    TableExist = True
                    Exit For
            End If
        Next tdf
    End With 'CurrentDb
End Function

'Delete Query
Public Sub DelQuery(QueryName As String)
    Dim QryExist As Boolean
    
    QryExist = QueryExist(QueryName)
    
    If QryExist = True Then
        DoCmd.SetWarnings False
        DoCmd.DeleteObject acQuery, QueryName
        DoCmd.SetWarnings True
    End If
End Sub

'Check whether query exist or not
Public Function QueryExist(QueryName As String) As Boolean
    QueryExist = False
    
    Dim QDf As QueryDef
    
    With CurrentDb
        For Each QDf In .QueryDefs
            If QueryName = QDf.Name Then
                    QueryExist = True
                    Exit For
            End If
        Next QDf
    End With 'CurrentDb
End Function

'Obtain record counts of a table
Public Function Table_RecordCount(Tbl_name As String) As Long
    Table_RecordCount = -1
    
    If TableExist(Tbl_name) = False Then
        GoTo Exit_Table_RecordCount
    End If
        
    
    Table_RecordCount = SQL_Obj_RecordCount(Tbl_name)
    
Exit_Table_RecordCount:
End Function

'Obtain record counts of a query
Public Function Query_RecordCount(Qry_name As String) As Long
    Query_RecordCount = -1
    
    If QueryExist(Qry_name) = False Then
        GoTo Exit_Query_RecordCount
    End If
        
    
    Query_RecordCount = SQL_Obj_RecordCount(Qry_name)
    
Exit_Query_RecordCount:
End Function

'Obtain record counts of a SQL object
Public Function SQL_Obj_RecordCount(SQL_Obj_name As String) As Long
    SQL_Obj_RecordCount = -1
    
    Dim RS As DAO.Recordset
    Set RS = CurrentDb.OpenRecordset(SQL_Obj_name)
    
    With RS
    
        If .EOF = True Then
            SQL_Obj_RecordCount = 0
        
        Else
            .MoveFirst
            .MoveLast
            
            SQL_Obj_RecordCount = .RecordCount
        End If
        
        .Close
    End With
    
End Function

'Check whether a table is valid or not
Public Function TableValid(TableName As String) As Boolean
    TableValid = False
    
    If TableExist(TableName) = False Then
        Exit Function
    End If
    
    If Table_RecordCount(TableName) <= 0 Then
        Exit Function
    End If
    
    TableValid = True
    
End Function

'Check whether a query is valid or not
Public Function QueryValid(Qryname As String) As Boolean
    QueryValid = False
    
    If QueryExist(Qryname) = False Then
        Exit Function
    End If
    
    If Query_RecordCount(Qryname) <= 0 Then
        Exit Function
    End If
    
    QueryValid = True
    
End Function

'Remove all link tables
Public Sub RemoveLink()
    Dim dbs As Database, tdf As TableDef
    ' Return Database variable that points to current database.
    Set dbs = CurrentDb
    For Each tdf In dbs.TableDefs
        If (tdf.Attributes = dbAttachedTable) Then
            DoCmd.DeleteObject acTable, tdf.Name
        End If
    Next tdf

End Sub

'Obtain a string with all columns names of a table
Public Function ObtainTblFldNameStr(Tbl_name As String)
    If TableExist(Tbl_name) = False Then
            GoTo Exit_ObtainTblFldNameStr
    End If
    
     
    With CurrentDb
        Dim td As TableDef
        Dim fld As Field
        
        Set td = .TableDefs(Tbl_name)
    
        If td.Fields.count <= 0 Then
            GoTo Exit_ObtainTblFldNameStr
        End If
    
        For Each fld In td.Fields
            ObtainTblFldNameStr = ObtainTblFldNameStr + ", " + "[" + fld.Name + "]"
        Next
    
        ObtainTblFldNameStr = Right(ObtainTblFldNameStr, Len(ObtainTblFldNameStr) - 1)
    
    End With 'CurrentDb
    
Exit_ObtainTblFldNameStr:
    Exit Function
End Function

'Find a column in a table
Public Function FindColInTbl(Tbl_name As String, Col_name As String) As Integer
    FindColInTbl = -1
    
    If TableExist(Tbl_name) = False Then
            GoTo Exit_FindColInTbl
    End If
    
     
    With CurrentDb
        Dim td As TableDef
        Dim fld As Field
        
        Set td = .TableDefs(Tbl_name)
    
        If td.Fields.count <= 0 Then
            GoTo Exit_FindColInTbl
        End If
    
        Dim Col_Idx As Integer
        Col_Idx = 0
        
        For Each fld In td.Fields
            If Col_name = fld.Name Then
                FindColInTbl = Col_Idx
                Exit For
            End If
            
            Col_Idx = Col_Idx + 1
        Next
    
    End With 'CurrentDb
    
Exit_FindColInTbl:
    Exit Function
End Function

'Export Access Table to Text file
Public Sub ExportTableToTxt(Tbl_name As String, OutputPathFile As String, Delim As String, HasFldName As Boolean)
    If OutputPathFile = "" Or Tbl_name = "" Then
        Exit Sub
    End If
    
    If TableExist(Tbl_name) = False Then
        Exit Sub
    End If
    
    If Delim = "" Then
        Delim = " "
    End If

    'If Dir(OutputPathFile, vbDirectory) = "" Then
    '    Exit Sub
    'End If

    Open OutputPathFile For Output As #1
    
    With CurrentDb
        Dim line As String
        line = ""
    
        If HasFldName = True Then
            Dim TD_Tbl As TableDef
            Set TD_Tbl = .TableDefs(Tbl_name)
            
            Dim fld As Field
            
            For Each fld In TD_Tbl.Fields
                line = line & fld.Name & Delim
            Next
            
            line = Left(line, Len(line) - Len(Delim))
            Print #1, line
        End If
        
        
        Dim RS_Tbl As Recordset
        Set RS_Tbl = .OpenRecordset(Tbl_name)
        
        With RS_Tbl
            .MoveFirst
        
            Dim FldIdx As Integer
            Dim fld_str As String
            
            Do Until .EOF
                FldIdx = 0
                line = ""
                
                For FldIdx = 0 To .Fields.count - 1
                    If .Fields(FldIdx).Type = dbDate Then
                        fld_str = Format(str(.Fields(FldIdx).Value), "MM/DD/YY")
                    Else
                        fld_str = .Fields(FldIdx).Value
                    End If
                    
                    If Len(fld_str) = 0 Then
                        fld_str = "0"
                    End If
                    
                    line = line & fld_str & Delim
                Next
                
                line = Left(line, Len(line) - Len(Delim))
                Print #1, line
                
                .MoveNext
            Loop
        End With 'RS_Tbl
    End With 'CurrentDb
    
    Close
    
End Sub

'Convert Access Table into HTML Format
Public Function ConvertTblToHtml(Tbl_name As String, Html As String) As String
    On Error GoTo Err_ConvertTblToHtml
    
    Dim FailedReason As String
    
    If TableValid(Tbl_name) = False Then
        FailedReason = Tbl_name & "is not valid"
        GoTo Exit_ConvertTblToHtml
    End If
    

    Html = Html & "<table border = ""1"", style = ""font-size:9pt;"">" & vbCrLf

    
    Dim RS_Tbl As DAO.Recordset
    Set RS_Tbl = CurrentDb.OpenRecordset(Tbl_name)
    
    'Create table
    With RS_Tbl
        Dim fld_idx As Integer
    
        'Create header
        Html = Html & "<tr>" & vbCrLf
        
        For fld_idx = 0 To .Fields.count - 1
            Html = Html & "<th bgcolor = #c0c0c0>" & .Fields(fld_idx).Name & "</th>" & vbCrLf
        Next fld_idx 'For fld_idx = 0 To .Fields.count - 1
        
        Html = Html & "</tr>"
        
        
        'Create rows
        .MoveFirst
        
        Do Until .EOF
            Html = Html & "<tr>" & vbCrLf
            
            For fld_idx = 0 To .Fields.count - 1
                Html = Html & "<td>" & .Fields(fld_idx).Value & "</td>" & vbCrLf
            Next fld_idx 'For fld_idx = 0 To .Fields.count - 1
            
            Html = Html & "</tr>" & vbCrLf
            
            .MoveNext
        Loop

        .Close
        
    End With 'RS_TblD
    
    
    Html = Html & "</table>"
    
    
Exit_ConvertTblToHtml:
    ConvertTblToHtml = FailedReason
    Exit Function

Err_ConvertTblToHtml:
    MsgBox Err.Description
    Resume Exit_ConvertTblToHtml
End Function
