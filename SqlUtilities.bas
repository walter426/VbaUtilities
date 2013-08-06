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

Attribute VB_Name = "SqlUtilities"
Option Compare Database

'Run SQL command without warning msg
Public Sub RunSQL_CmdWithoutWarning(SQL_cmd As String)
    DoCmd.SetWarnings False
    Application.SetOption "Confirm Action Queries", False
    
    DoCmd.RunSQL SQL_cmd

    Application.SetOption "Confirm Action Queries", True
    DoCmd.SetWarnings True
End Sub

'Re-Select table columns
Public Sub ModifyTbl_ReSelect(Tbl_name, str_select)
    Dim Tbl_T_name As String
    Tbl_T_name = Tbl_name & "_temp"
    
    DelTable (Tbl_T_name)
    
    SQL_cmd = "SELECT " & str_select & " " & vbCrLf & _
                "INTO [" & Tbl_T_name & "]" & vbCrLf & _
                "FROM [" & Tbl_name & "]" & vbCrLf & _
                ";"
    
    RunSQL_CmdWithoutWarning (SQL_cmd)
    
    
    DelTable (Tbl_name)
    DoCmd.Rename Tbl_name, acTable, Tbl_T_name
    
End Sub

'Update multiple columns of a table under the same condition
Public Function UpdateTblColBatchly(Tbl_src_name As String, Str_Col_Update As String, SQL_Format_Set As String, SQL_Format_Where As String) As String
    On Error GoTo Err_UpdateTblColBatchly
    
    Dim FailedReason As String

    If TableExist(Tbl_src_name) = False Then
        FailedReason = Tbl_src_name & "does not exist!"
        GoTo Exit_UpdateTblColBatchly
    End If
    
    
    Str_Col_Update = Trim(Str_Col_Update)
    
    
    If Str_Col_Update = "*" Then
        With CurrentDb
            Dim RS_Tbl_src As Recordset
            Set RS_Tbl_src = .OpenRecordset(Tbl_src_name)
            
            Dim fld_idx As Integer
            
            With RS_Tbl_src
                For fld_idx = 0 To .Fields.count - 1
                    Call UpdateTblCol(Tbl_src_name, .Fields(fld_idx).Name, SQL_Format_Set, SQL_Format_Where)
                Next fld_idx
                
                .Close
            End With 'RS_Tbl_src
            
            .Close
        End With 'CurrentDb
        
    Else
        Dim Col_Update As Variant
        Col_Update = SplitStrIntoArray(Str_Col_Update, ",")
    
        
        Dim Col_Idx As Integer
        Dim ColName As String
        
        For Col_Idx = 0 To UBound(Col_Update)
            ColName = Col_Update(Col_Idx)
            Call UpdateTblCol(Tbl_src_name, ColName, SQL_Format_Set, SQL_Format_Where)
        Next Col_Idx
        
    End If


Exit_UpdateTblColBatchly:
    UpdateTblColBatchly = FailedReason
    Exit Function

Err_UpdateTblColBatchly:
    Call ShowMsgBox(Err.Description)
    Resume Exit_UpdateTblColBatchly
    
End Function

'Update a column of a table under a specified condition
Public Function UpdateTblCol(Tbl_src_name As String, ColName As String, SQL_Format_Set, SQL_Format_Where As String) As String
    On Error GoTo Err_UpdateTblCol
    
    Dim FailedReason As String

    If TableExist(Tbl_src_name) = False Then
        FailedReason = Tbl_src_name & "does not exist!"
        GoTo Exit_UpdateTblCol
    End If
    
    
    Dim SQL_Seg_Set As String
    Dim SQL_Seg_Where As String
    Dim SQL_cmd As String
    

    SQL_Seg_Set = "SET [" & ColName & "] = " & Replace(SQL_Format_Set, "*", "[" & ColName & "]") & " "
    
    If SQL_Format_Where <> "" Then
        SQL_Seg_Where = "WHERE " & Replace(SQL_Format_Where, "*", "[" & ColName & "]")
    End If
    
    SQL_cmd = "UPDATE " & Tbl_src_name & " " & vbCrLf & _
                SQL_Seg_Set & " " & vbCrLf & _
                SQL_Seg_Where & " " & vbCrLf & _
                ";"
    
    RunSQL_CmdWithoutWarning (SQL_cmd)

Exit_UpdateTblCol:
    UpdateTblCol = FailedReason
    Exit Function

Err_UpdateTblCol:
    Call ShowMsgBox(Err.Description)
    Resume Exit_UpdateTblCol
    
End Function

'Create Table with dedicated Column and Expressions from a source table
Public Function CreateTbl_ColAndExpr(Tbl_src_name As String, Str_Col_Id As String, Str_Col_Order As String, SQL_Seg_ColAndExpr As String, SQL_Seg_Where As String, Tbl_output_name As String) As String
    On Error GoTo Err_CreateTbl_ColAndExpr
    
    Dim FailedReason As String

    If TableExist(Tbl_src_name) = False Then
        FailedReason = Tbl_src_name & "does not exist!"
        GoTo Exit_CreateTbl_ColAndExpr
    End If
    
    Dim Col_Id As Variant
    Dim Col_Order As Variant
    
    Col_Id = SplitStrIntoArray(Str_Col_Id, ",")
    Col_Order = SplitStrIntoArray(Str_Col_Order, ",")
    
    
    DelTable (Tbl_output_name)
    
    
    Dim SQL_Seg_Select As String
    Dim SQL_Seg_OrderBy As String
    
    SQL_Seg_Select = "SELECT "
    SQL_Seg_OrderBy = ""
    
    Dim Col_Idx As Integer
    
    For Col_Idx = 0 To UBound(Col_Id)
        SQL_Seg_Select = SQL_Seg_Select & "[" & Col_Id(Col_Idx) & "], "
    Next Col_Idx
    
    
    SQL_Seg_Select = SQL_Seg_Select & SQL_Seg_ColAndExpr
    
    
    If UBound(Col_Order) >= 0 Then
        SQL_Seg_OrderBy = "ORDER BY "
        
        For Col_Idx = 0 To UBound(Col_Order)
            SQL_Seg_OrderBy = SQL_Seg_OrderBy & "[" & Col_Order(Col_Idx) & "], "
        Next Col_Idx
        
        SQL_Seg_OrderBy = Left(SQL_Seg_OrderBy, Len(SQL_Seg_OrderBy) - 2)
        
    End If
    
    If SQL_Seg_Where <> "" Then
        SQL_Seg_Where = "WHERE " & SQL_Seg_Where
    End If
    
    
    Dim SQL_cmd As String
    
    SQL_cmd = SQL_Seg_Select & " " & vbCrLf & _
                "INTO [" & Tbl_output_name & "] " & vbCrLf & _
                "FROM [" & Tbl_src_name & "] " & vbCrLf & _
                SQL_Seg_Where & " " & vbCrLf & _
                SQL_Seg_OrderBy & " " & vbCrLf & _
                ";"
    
    'MsgBox SQL_cmd
    RunSQL_CmdWithoutWarning (SQL_cmd)

Exit_CreateTbl_ColAndExpr:
    CreateTbl_ColAndExpr = FailedReason
    Exit Function

Err_CreateTbl_ColAndExpr:
    Call ShowMsgBox(Err.Description)
    Resume Exit_CreateTbl_ColAndExpr
    
End Function

'Create Table of group function, there is a default Group function for all columns, columns can be specified to different group fucntion
Public Function CreateTbl_Group(Tbl_input_name As String, Str_Col_Group As String, Str_Col_Order As String, Str_GroupFunc_all As String, Str_Col_Sum As String, Str_Col_Avg As String, Str_Col_Max As String, SQL_Seg_Where As String, Tbl_output_name As String) As String
    On Error GoTo Err_CreateTbl_Group
    
    Dim FailedReason As String
    
    If TableValid(Tbl_input_name) = False Then
        FailedReason = Tbl_input_name & " is not valid!"
        GoTo Exit_CreateTbl_Group
    End If

    If Len(Str_Col_Group) = 0 Then
        GoTo Exit_CreateTbl_Group
    End If
    
    
    Str_GroupFunc_all = Trim(Str_GroupFunc_all)
    
    
    Dim Col_Idx As Integer
    
    Dim Col_Group As Variant
    Dim Col_Order As Variant
    Dim Col_Sum As Variant
    Dim Col_Avg As Variant
    Dim Col_Max As Variant
    
    Col_Group = SplitStrIntoArray(Str_Col_Group, ",")
    Col_Order = SplitStrIntoArray(Str_Col_Order, ",")
    Col_Sum = SplitStrIntoArray(Str_Col_Sum, ",")
    Col_Avg = SplitStrIntoArray(Str_Col_Avg, ",")
    Col_Max = SplitStrIntoArray(Str_Col_Max, ",")
    

    DelTable (Tbl_output_name)


    With CurrentDb
        Dim RS_Tbl_input As Recordset
        Set RS_Tbl_input = .OpenRecordset(Tbl_input_name)
        
        With RS_Tbl_input
            Dim SQL_Seg_Select As String
            Dim SQL_Seg_GroupBy As String
            Dim SQL_Seg_OrderBy As String
            
            SQL_Seg_Select = "SELECT "
            SQL_Seg_GroupBy = "GROUP BY "
            SQL_Seg_OrderBy = ""
            
            
            Dim fld_idx As Integer
            Dim fld_name As String
            
            Dim IsColForGroupBy As Boolean
            
            Dim NumOfCol_Group_found As Integer
            NumOfCol_Group_found = 0
            
            Dim Col_GroupBy As Variant
            
            Dim GroupFunc As String
            
            For fld_idx = 0 To .Fields.count - 1
                fld_name = .Fields(fld_idx).Name
                IsColForGroupBy = False
                
                If NumOfCol_Group_found <= UBound(Col_Group) Then
                    If FindStrInArray(Col_Group, fld_name) > -1 Then
                        SQL_Seg_GroupBy = SQL_Seg_GroupBy & "[" & fld_name & "], "
                        IsColForGroupBy = True
                        NumOfCol_Group_found = NumOfCol_Group_found + 1
                    End If
                End If
                
                    
                If IsColForGroupBy = True Then
                    SQL_Seg_Select = SQL_Seg_Select & "[" & fld_name & "], "
                    
                ElseIf .Fields(fld_idx).Type <> dbText And .Fields(fld_idx).Type <> dbDate Then
                    If FindStrInArray(Col_Sum, fld_name) > -1 Then
                        GroupFunc = "SUM"
                        
                    ElseIf FindStrInArray(Col_Avg, fld_name) > -1 Then
                        GroupFunc = "AVG"
                        
                    ElseIf FindStrInArray(Col_Max, fld_name) > -1 Then
                        GroupFunc = "MAX"
                        
                    Else
                        GroupFunc = Str_GroupFunc_all
                        
                    End If
                    
                    If GroupFunc <> "" Then
                        SQL_Seg_Select = SQL_Seg_Select & GroupFunc & "([" & Tbl_input_name & "].[" & fld_name & "]) AS [" & fld_name & "], "
                    End If
                    
                End If
                
Next_CreateTbl_Group_1:
            Next fld_idx
            
            
            SQL_Seg_Select = Left(SQL_Seg_Select, Len(SQL_Seg_Select) - 2)
            SQL_Seg_GroupBy = Left(SQL_Seg_GroupBy, Len(SQL_Seg_GroupBy) - 2)
            
            .Close
            
        End With 'RS_Tbl_input
        
        
        If UBound(Col_Order) >= 0 Then
            SQL_Seg_OrderBy = "ORDER BY "
            
            For Col_Idx = 0 To UBound(Col_Order)
                SQL_Seg_OrderBy = SQL_Seg_OrderBy & "[" & Col_Order(Col_Idx) & "], "
            Next Col_Idx
            
            SQL_Seg_OrderBy = Left(SQL_Seg_OrderBy, Len(SQL_Seg_OrderBy) - 2)
            
        End If
        
        
        If SQL_Seg_Where <> "" Then
            SQL_Seg_Where = "WHERE " & SQL_Seg_Where
        End If
        
        Dim SQL_cmd As String
        
        SQL_cmd = SQL_Seg_Select & " " & vbCrLf & _
                    "INTO [" & Tbl_output_name & "] " & vbCrLf & _
                    "FROM [" & Tbl_input_name & "] " & vbCrLf & _
                    SQL_Seg_Where & " " & vbCrLf & _
                    SQL_Seg_GroupBy & " " & vbCrLf & _
                    SQL_Seg_OrderBy & " " & vbCrLf & _
                    ";"
        
        'MsgBox SQL_cmd
        RunSQL_CmdWithoutWarning (SQL_cmd)
        
        .Close
        
    End With 'CurrentDb

Exit_CreateTbl_Group:
    CreateTbl_Group = FailedReason
    Exit Function

Err_CreateTbl_Group:
    Call ShowMsgBox(Err.Description)
    Resume Exit_CreateTbl_Group
    
End Function

'Create a set of grouped table, the grouping config is set in a specified table
Public Function CreateTbls_TblToSum(Tbl_MT_name As String) As String
    On Error GoTo Err_CreateTbls_TblToSum
    
    Dim FailedReason As String
    
    If TableExist(Tbl_MT_name) = False Then
        FailedReason = Tbl_MT_name & " does not exist!"
        GoTo Exit_CreateTbls_TblToSum
    End If


    With CurrentDb
        Dim RS_Tbl_MT As Recordset
        Set RS_Tbl_MT = .OpenRecordset(Tbl_MT_name)
        
        With RS_Tbl_MT
            Dim Tbl_src_name As String
            Dim Tbl_sum_name As String
            Dim Str_Col_Order As String
            Dim Str_GroupFunc_all As String
            Dim Str_Col_Sum As String
            Dim Str_Col_Avg As String
            Dim Str_Col_Max As String
            
            .MoveFirst
        
            Do Until .EOF
            
                If .Fields("Enable").Value = False Then
                    GoTo Loop_CreateTbls_TblToSum_1
                End If
            
                Tbl_src_name = .Fields("Tbl_src").Value
                
                If TableExist(Tbl_src_name) = False Then
                    GoTo Loop_CreateTbls_TblToSum_1
                End If


                Tbl_sum_name = .Fields("Tbl_sum").Value
                
                If Len(Tbl_sum_name) = 0 Then
                    GoTo Loop_CreateTbls_TblToSum_1
                End If
                
                
                If IsNull(.Fields("Col_Group").Value) = True Then
                    GoTo Loop_CreateTbls_TblToSum_1
                End If
                
                If IsNull(.Fields("Col_Order").Value) = True Then
                    Str_Col_Order = ""
                Else
                    Str_Col_Order = .Fields("Col_Order").Value
                End If
                
                If IsNull(.Fields("GroupFunc_all").Value) = True Then
                    Str_GroupFunc_all = ""
                Else
                    Str_GroupFunc_all = .Fields("GroupFunc_all").Value
                End If
                
                If IsNull(.Fields("Col_Sum").Value) = True Then
                    Str_Col_Sum = ""
                Else
                    Str_Col_Sum = .Fields("Col_Sum").Value
                End If
                
                If IsNull(.Fields("Col_Avg").Value) = True Then
                    Str_Col_Avg = ""
                Else
                    Str_Col_Avg = .Fields("Col_Avg").Value
                End If
                
                If IsNull(.Fields("Col_Max").Value) = True Then
                    Str_Col_Max = ""
                Else
                    Str_Col_Max = .Fields("Col_Max").Value
                End If
                
                
                Call CreateTbl_Group(Tbl_src_name, .Fields("Col_Group").Value, Str_Col_Order, Str_GroupFunc_all, Str_Col_Sum, Str_Col_Avg, Str_Col_Max, "", Tbl_sum_name)

Loop_CreateTbls_TblToSum_1:
                .MoveNext
            Loop
            
            .Close
            
        End With 'RS_Tbl_MT
        
        .Close
        
    End With 'CurrentDb


Exit_CreateTbls_TblToSum:
    CreateTbls_TblToSum = FailedReason
    Exit Function

Err_CreateTbls_TblToSum:
    Call ShowMsgBox(Err.Description)
    Resume Exit_CreateTbls_TblToSum
    
End Function


'Create table which are joined from two tables having the same columns for joining
Public Function CreateTbl_JoinTwoTbl(Tbl_src_1_name As String, Tbl_src_2_name As String, JoinCond As String, Str_Col_Join As String, Str_Col_Order As String, Tbl_output_name As String) As String
    On Error GoTo Err_CreateTbl_JoinTwoTbl
    
    Dim FailedReason As String

    If TableExist(Tbl_src_1_name) = False Then
        FailedReason = Tbl_src_1_name & "does not exist!"
        GoTo Exit_CreateTbl_JoinTwoTbl
    End If
    
    If TableExist(Tbl_src_2_name) = False Then
        FailedReason = Tbl_src_2_name & "does not exist!"
        GoTo Exit_CreateTbl_JoinTwoTbl
    End If
    
    
    Dim Col_Join As Variant
    Col_Join = SplitStrIntoArray(Str_Col_Join, ",")
    
    If UBound(Col_Join) < 0 Then
        GoTo Exit_CreateTbl_JoinTwoTbl
    End If
    
    Dim Col_Order As Variant
    Col_Order = SplitStrIntoArray(Str_Col_Order, ",")
    
    
    DelTable (Tbl_output_name)
    
    
    Dim SQL_Seg_Select As String
    Dim SQL_Seg_OrderBy As String
    
    SQL_Seg_Select = "SELECT "
    SQL_Seg_OrderBy = ""
    
    With CurrentDb
        Dim RS_Tbl_src As Recordset
        Set RS_Tbl_src = .OpenRecordset(Tbl_src_1_name)
        
        Dim fld_idx As Integer
        Dim fld_name As String
        
        With RS_Tbl_src
            For fld_idx = 0 To .Fields.count - 1
                fld_name = .Fields(fld_idx).Name
                SQL_Seg_Select = SQL_Seg_Select & "[" & Tbl_src_1_name & "].[" & fld_name & "], "
            Next fld_idx
            
            .Close
        End With 'RS_Tbl_src
        
        
        Set RS_Tbl_src = .OpenRecordset(Tbl_src_2_name)
        
        With RS_Tbl_src
            Dim NumOfCol_Join_found As Integer
            NumOfCol_Join_found = 0
                
                
            For fld_idx = 0 To .Fields.count - 1
                fld_name = .Fields(fld_idx).Name
                
                If NumOfCol_Join_found <= UBound(Col_Join) And FindStrInArray(Col_Join, fld_name) > -1 Then
                    NumOfCol_Join_found = NumOfCol_Join_found + 1
                Else
                    SQL_Seg_Select = SQL_Seg_Select & "[" & Tbl_src_2_name & "].[" & fld_name & "], "
                End If
            Next fld_idx
            
            .Close
        End With 'RS_Tbl_src
    End With 'CurrentDb
    
    SQL_Seg_Select = Left(SQL_Seg_Select, Len(SQL_Seg_Select) - 2)
    
    
    Dim SQL_Seg_JoinOn As String
    SQL_Seg_JoinOn = "("
    
    Dim Col_Idx As Integer
    
    For Col_Idx = 0 To UBound(Col_Join)
        SQL_Seg_JoinOn = SQL_Seg_JoinOn & "[" & Tbl_src_1_name & "].[" & Col_Join(Col_Idx) & "] = [" & Tbl_src_2_name & "].[" & Col_Join(Col_Idx) & "] AND "
    Next Col_Idx

    SQL_Seg_JoinOn = Left(SQL_Seg_JoinOn, Len(SQL_Seg_JoinOn) - 4) & ")"

    
    If UBound(Col_Order) >= 0 Then
        SQL_Seg_OrderBy = "ORDER BY "
        
        For Col_Idx = 0 To UBound(Col_Order)
            SQL_Seg_OrderBy = SQL_Seg_OrderBy & "[" & Tbl_src_1_name & "].[" & Col_Order(Col_Idx) & "], "
        Next Col_Idx
        
        SQL_Seg_OrderBy = Left(SQL_Seg_OrderBy, Len(SQL_Seg_OrderBy) - 2)
        
    End If
    
    
    Dim SQL_cmd As String
    
    SQL_cmd = SQL_Seg_Select & " " & vbCrLf & _
                "INTO [" & Tbl_output_name & "] " & vbCrLf & _
                "FROM [" & Tbl_src_1_name & "] " & JoinCond & " JOIN [" & Tbl_src_2_name & "] " & vbCrLf & _
                "ON " & SQL_Seg_JoinOn & vbCrLf & _
                SQL_Seg_OrderBy & " " & vbCrLf & _
                ";"
    
    'MsgBox SQL_cmd
    RunSQL_CmdWithoutWarning (SQL_cmd)

Exit_CreateTbl_JoinTwoTbl:
    CreateTbl_JoinTwoTbl = FailedReason
    Exit Function

Err_CreateTbl_JoinTwoTbl:
    Call ShowMsgBox(Err.Description)
    Resume Exit_CreateTbl_JoinTwoTbl
    
End Function

'Create table which is cancatenated from two tables of the same structure
Public Function CreateTbl_ConcatTwoTbl(Tbl_src_1_name As String, Type_1 As String, Tbl_src_2_name As String, Type_2 As String, Tbl_output_name As String) As String
    On Error GoTo Err_CreateTbl_ConcatTwoTbl
    
    Dim FailedReason As String

    If TableExist(Tbl_src_1_name) = False Then
        FailedReason = Tbl_src_1_name & "does not exist!"
        GoTo Exit_CreateTbl_ConcatTwoTbl
    End If


    If TableExist(Tbl_src_2_name) = False Then
        FailedReason = Tbl_src_2_name & "does not exist!"
        GoTo Exit_CreateTbl_ConcatTwoTbl
    End If

    DelTable (Tbl_output_name)

    
    Dim SQL_Seq_Type_1 As String
    Dim SQL_Seq_Type_2 As String
    
        
    If Type_1 = "" Or Type_2 = "" Then
        SQL_Seq_Type_1 = ""
        SQL_Seq_Type_2 = ""
    
    Else
        SQL_Seq_Type_1 = Chr(34) & Type_1 & Chr(34) & " AS [Type], "
        SQL_Seq_Type_2 = Chr(34) & Type_2 & Chr(34) & " AS [Type], "
        
    End If
        

    Dim SQL_cmd As String
    
    SQL_cmd = "SELECT " & Chr(34) & "null" & Chr(34) & " AS [Type], " & Tbl_src_1_name & ".* " & vbCrLf & _
                "INTO " & Tbl_output_name & " " & vbCrLf & _
                "FROM " & Tbl_src_1_name & " " & vbCrLf & _
                "WHERE 1 = 0 " & vbCrLf & _
                ";"
                
    RunSQL_CmdWithoutWarning (SQL_cmd)

    
    SQL_cmd = "INSERT INTO " & Tbl_output_name & " " & vbCrLf & _
                "SELECT " & SQL_Seq_Type_1 & "[" & Tbl_src_1_name & "].* " & vbCrLf & _
                "FROM [" & Tbl_src_1_name & "] " & vbCrLf & _
                ";"
    
    RunSQL_CmdWithoutWarning (SQL_cmd)
    
    
    SQL_cmd = "INSERT INTO " & Tbl_output_name & " " & vbCrLf & _
                 "SELECT " & SQL_Seq_Type_2 & "[" & Tbl_src_2_name & "].* " & vbCrLf & _
                "FROM [" & Tbl_src_2_name & "] " & vbCrLf & _
                ";"
    
    RunSQL_CmdWithoutWarning (SQL_cmd)

    
Exit_CreateTbl_ConcatTwoTbl:
    CreateTbl_ConcatTwoTbl = FailedReason
    Exit Function

Err_CreateTbl_ConcatTwoTbl:
    Call ShowMsgBox(Err.Description)
    Resume Exit_CreateTbl_ConcatTwoTbl
    
End Function

