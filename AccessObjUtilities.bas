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
Public Function RemoveLink() As String
    On Error GoTo Err_RemoveLink
    
    Dim FailedReason As String
    
    Dim dbs As Database, tdf As TableDef
    ' Return Database variable that points to current database.
    Set dbs = CurrentDb
    For Each tdf In dbs.TableDefs
        If (tdf.Attributes = dbAttachedTable) Then
            DoCmd.DeleteObject acTable, tdf.Name
        End If
    Next tdf

Exit_RemoveLink:
    RemoveLink = FailedReason
    Exit Function

Err_RemoveLink:
    FailedReason = Err.Description
    Resume Exit_RemoveLink
    
End Function

'Get the current path of a link table
Public Function GetLinkTblPath(Tbl_name As String) As String
    On Error GoTo Exit_GetLinkTblPath
    
    Dim LinkTblPath As String
    
    LinkTblPath = CurrentDb.TableDefs(Tbl_name).Connect
    LinkTblPath = Right(LinkTblPath, Len(LinkTblPath) - (InStr(1, LinkTblPath, "DATABASE=") + 8)) & "\" & CurrentDb.TableDefs(Tbl_name).SourceTableName
    
    GetLinkTblPath = LinkTblPath
    
Exit_GetLinkTblPath:
    Exit Function

Err_GetLinkTblPath:
    ShowMsgBox (Err.Description)
    GetLinkTblPath = ""
    Resume Exit_GetLinkTblPath
    
End Function

'Get Link Table connection Info
Public Function GetLinkTblConnInfo(Tbl_name As String, param As String) As String
    On Error GoTo Exit_GetLinkTblConnInfo
    
    Dim LinkTblConnInfo As Variant
    LinkTblConnInfo = SplitStrIntoArray(CurrentDb.TableDefs(Tbl_name).Connect, ";")

    Dim param_idx As Integer
    Dim LinkTblConnParam As String
    
    param = param & "="

    For param_idx = 0 To UBound(LinkTblConnInfo)
        LinkTblConnParam = LinkTblConnInfo(param_idx)

        If Left(LinkTblConnParam, Len(param)) = param Then
            GetLinkTblConnInfo = Right(LinkTblConnParam, Len(LinkTblConnParam) - Len(param))
            Exit For
        End If
    Next param_idx
    
    
Exit_GetLinkTblConnInfo:
    Exit Function

Err_GetLinkTblConnInfo:
    ShowMsgBox (Err.Description)
    GetLinkTblConnInfo = ""
    Resume Exit_GetLinkTblConnInfo
    
End Function

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

'Export Table to Text file
Public Function ExportTableToTxt(Tbl_name As String, des As String, Optional Delim As String = " ", Optional HasFldName As Boolean = True, Optional NullStr As String = "", Optional DateFmt As String = "MM/DD/YY", Optional TimeFmt As String = "h:mm") As String
    On Error GoTo Err_ExportTableToTxt
    
    Dim FailedReason As String
    
    If des = "" Or Tbl_name = "" Then
        FailedReason = "Input is invalid"
        GoTo Exit_ExportTableToTxt
    End If
    
    If TableExist(Tbl_name) = False Then
        FailedReason = Tbl_name & " does not exist"
        GoTo Exit_ExportTableToTxt
    End If
    
    If Delim = "" Then
        Delim = " "
    End If

    
    Dim des_PortNum As Integer
    des_PortNum = FreeFile

    Open des For Output As #des_PortNum

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
            Print #des_PortNum, line
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
                    If IsNull(.Fields(FldIdx)) = True Then
                        fld_str = NullStr
                    
                    ElseIf .Fields(FldIdx).Type = dbDate Then
                        If .Fields(FldIdx).Value > 1 Then
                            fld_str = Format(str(.Fields(FldIdx).Value), DateFmt)
                        Else
                            fld_str = Format(str(.Fields(FldIdx).Value), TimeFmt)
                        End If
                        
                    Else
                        fld_str = .Fields(FldIdx).Value
                        
                    End If
                    
                    
                    line = line & fld_str & Delim
                    
                Next
                
                line = Left(line, Len(line) - Len(Delim))
                Print #des_PortNum, line
                
                .MoveNext
                
            Loop
            
        End With 'RS_Tbl
        
        .Close
        
    End With 'CurrentDb
    
    Close #des_PortNum


Exit_ExportTableToTxt:
    ExportTableToTxt = FailedReason
    Exit Function

Err_ExportTableToTxt:
    FailedReason = Err.Description
    Resume Exit_ExportTableToTxt
        
End Function
