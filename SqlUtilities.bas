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

'To create an expression that consists of a set of vector columns aggregated in a specified pattern
Public Function CreateSqlSeg_VectorColAgg(col_pattern As String, str_agg As String, Idx_start As Integer, Idx_end As Integer, Optional wildcard As String = "#") As String
    On Error GoTo Err_CreateSqlSeg_VectorColAgg
    
    Dim SQL_Seg As String
    Dim col_idx As Integer
    
    For col_idx = Idx_start To Idx_end
        SQL_Seg = SQL_Seg & Replace(col_pattern, wildcard, col_idx) & " " & str_agg & " "
    Next col_idx
    
    SQL_Seg = Left(SQL_Seg, Len(SQL_Seg) - 2)
    
    
Exit_CreateSqlSeg_VectorColAgg:
    CreateSqlSeg_VectorColAgg = SQL_Seg
    Exit Function

Err_CreateSqlSeg_VectorColAgg:
    ShowMsgBox (Err.Description)
    Resume Exit_CreateSqlSeg_VectorColAgg
    
End Function

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
    FailedReason = Err.Description
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
    FailedReason = Err.Description
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
    
    
    If SQL_Seg_ColAndExpr <> "" Then
        SQL_Seg_Select = SQL_Seg_Select & SQL_Seg_ColAndExpr
    Else
        SQL_Seg_Select = Left(SQL_Seg_Select, Len(SQL_Seg_Select) - 2)
    End If
    
    
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
    FailedReason = Err.Description
    Resume Exit_CreateTbl_ColAndExpr
    
End Function

'Create Table of group function, there is a default Group function for all columns, columns can be specified to different group fucntion
Public Function CreateTbl_Group(Tbl_input_name As String, Tbl_output_name As String, Str_Col_Group As String, Optional Str_GroupFunc_all As String = "", Optional GF_all_dbTypes As Variant = "", Optional Str_Col_UnSelected As String = "", Optional ByVal GroupFunc_Col_Pairs As Variant = "", Optional SQL_Seg_Where As String = "", Optional Str_Col_Order As String = "") As String
    On Error GoTo Err_CreateTbl_Group
    
    Dim FailedReason As String
    
    If TableValid(Tbl_input_name) = False Then
        FailedReason = Tbl_input_name & " is not valid!"
        GoTo Exit_CreateTbl_Group
    End If

    If Len(Str_Col_Group) = 0 Then
        FailedReason = "No Any Group Columns"
        GoTo Exit_CreateTbl_Group
    End If

    
    If Str_GroupFunc_all <> "" Then
        If UBound(GF_all_dbTypes) < 0 Then
            FailedReason = "No db Type is assigned for the general group function"
            GoTo Exit_CreateTbl_Group
        End If
    End If


    If VarType(GroupFunc_Col_Pairs) <> vbArray + vbVariant Then
        If Str_GroupFunc_all = "" Then
            FailedReason = "No Any Group Functions for all or specified columns"
            GoTo Exit_CreateTbl_Group
        Else
            GroupFunc_Col_Pairs = Array()
        End If
    End If
         
    Dim GF_C_P_idx As Integer

    For GF_C_P_idx = 0 To UBound(GroupFunc_Col_Pairs)
        GroupFunc_Col_Pairs(GF_C_P_idx)(1) = SplitStrIntoArray(GroupFunc_Col_Pairs(GF_C_P_idx)(1) & "", ",")
    Next GF_C_P_idx
    
    
    Str_GroupFunc_all = Trim(Str_GroupFunc_all)
    
    
    Dim col_idx As Integer
    
    Dim Col_Group As Variant
    Dim Col_UnSelected As Variant
    Dim Col_Order As Variant

    
    Col_Group = SplitStrIntoArray(Str_Col_Group, ",")
    Col_UnSelected = SplitStrIntoArray(Str_Col_UnSelected, ",")
    Col_Order = SplitStrIntoArray(Str_Col_Order, ",")
    

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
            
            Dim GroupFunc_Col_Pair As Variant
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
                    
                ElseIf FindStrInArray(Col_UnSelected, fld_name) < 0 Then
                    GroupFunc = ""
                    
                    For Each GroupFunc_Col_Pair In GroupFunc_Col_Pairs
                        If FindStrInArray(GroupFunc_Col_Pair(1), fld_name) > -1 Then
                            GroupFunc = GroupFunc_Col_Pair(0)
                        End If

                    Next GroupFunc_Col_Pair
                    
 
                    If GroupFunc = "" And Str_GroupFunc_all <> "" Then
                        For Each GF_all_dbType In GF_all_dbTypes
                            If .Fields(fld_idx).Type = GF_all_dbType Then
                                GroupFunc = Str_GroupFunc_all
                            End If
                        
                        Next GF_all_dbType
                        
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
            
            For col_idx = 0 To UBound(Col_Order)
                SQL_Seg_OrderBy = SQL_Seg_OrderBy & "[" & Col_Order(col_idx) & "], "
            Next col_idx
            
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
    FailedReason = Err.Description
    Resume Exit_CreateTbl_Group
    
End Function


'Create a set of grouped table, the grouping config is set in a specified table
Public Function CreateTbls_Group(Tbl_MT_name As String) As String
    On Error GoTo Err_CreateTbls_Group
    
    Dim FailedReason As String
    
    If TableExist(Tbl_MT_name) = False Then
        FailedReason = Tbl_MT_name & " does not exist!"
        GoTo Exit_CreateTbls_Group
    End If


    With CurrentDb
        Dim RS_Tbl_MT As Recordset
        Set RS_Tbl_MT = .OpenRecordset(Tbl_MT_name)
        
        With RS_Tbl_MT
            Dim FailedReason_1 As String
            
            Dim Tbl_src_name As String
            Dim Tbl_Group_name As String
            
            Dim Str_Col_Group As String
            Dim Str_Col_UnSelected As String
            Dim Str_GroupFunc_all As String
            Dim GF_all_dbTypes As Variant
            
            Dim GroupFunc_Col_Pairs As Variant
            Dim SQL_Seg_Where As String
            Dim Str_Col_Order As String
            
            .MoveFirst
        
            Do Until .EOF
            
                If .Fields("Enable").Value = False Then
                    GoTo Loop_CreateTbls_Group_1
                End If
            
                Tbl_src_name = .Fields("Tbl_src").Value
                
                If TableExist(Tbl_src_name) = False Then
                    GoTo Loop_CreateTbls_Group_1
                End If

                Tbl_Group_name = .Fields("Tbl_Group").Value
                
                If Len(Tbl_Group_name) = 0 Then
                    GoTo Loop_CreateTbls_Group_1
                End If
                
                If IsNull(.Fields("Col_Group").Value) = True Then
                    GoTo Loop_CreateTbls_Group_1
                End If
                
                If IsNull(.Fields("GroupFunc_all").Value) = True Then
                    Str_GroupFunc_all = ""
                Else
                    Str_GroupFunc_all = .Fields("GroupFunc_all").Value
                End If
                
                GF_all_dbTypes = Array(dbInteger, dbLong, dbSingle, dbDouble, dbDecimal)
                
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
                
                GroupFunc_Col_Pairs = Array(Array("SUM", Str_Col_Sum), _
                                            Array("AVG", Str_Col_Avg), _
                                            Array("MAX", Str_Col_Max))
                
                If IsNull(.Fields("Col_Order").Value) = True Then
                    Str_Col_Order = ""
                Else
                    Str_Col_Order = .Fields("Col_Order").Value
                End If
                
                If IsNull(.Fields("Cond").Value) = True Then
                    SQL_Seg_Where = ""
                Else
                    SQL_Seg_Where = .Fields("Cond").Value
                End If

                
                FailedReason_1 = CreateTbl_Group(Tbl_src_name, Tbl_Group_name, .Fields("Col_Group").Value, Str_GroupFunc_all:=Str_GroupFunc_all, GF_all_dbTypes:=GF_all_dbTypes, GroupFunc_Col_Pairs:=GroupFunc_Col_Pairs, Str_Col_Order:=Str_Col_Order)
                
                If FailedReason_1 <> "" Then
                    FailedReason = FailedReason & Tbl_Group_name & ": " & FailedReason_1 & vbCrLf
                End If

Loop_CreateTbls_Group_1:
                .MoveNext
            Loop
            
            .Close
            
        End With 'RS_Tbl_MT
        
        .Close
        
    End With 'CurrentDb


Exit_CreateTbls_Group:
    CreateTbls_Group = FailedReason
    Exit Function

Err_CreateTbls_Group:
    FailedReason = Err.Description
    Resume Exit_CreateTbls_Group
    
End Function


'Create table which is joined from two tables
Public Function CreateTbl_JoinTwoTbl(Tbl_src_1_name As String, Tbl_src_2_name As String, JoinCond As String, ColSet_Join_1 As Variant, ColSet_Join_2 As Variant, Tbl_des_name As String, Optional ColSet_src_1 As Variant = Null, Optional ColSet_src_2 As Variant = Null, Optional ColSet_Order As Variant = Null) As String
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
    
    If IsNull(ColSet_Join_1) = True Then
        GoTo Exit_CreateTbl_JoinTwoTbl
    End If
    
    If IsNull(ColSet_Join_2) = True Then
        GoTo Exit_CreateTbl_JoinTwoTbl
    End If
    

    DelTable (Tbl_des_name)
    
    
    Dim Col_Idx As Integer
    
    
    With CurrentDb
        If IsNull(ColSet_src_1) = True Then
            Dim RS_Tbl_src As Recordset
            Set RS_Tbl_src = .OpenRecordset(Tbl_src_1_name)
            
            Dim fld_idx As Integer
            Dim fld_name As String
            
            ColSet_src_1 = Array()
            
            With RS_Tbl_src
                For fld_idx = 0 To .Fields.count - 1
                    fld_name = .Fields(fld_idx).Name
                    Call AppendArray(ColSet_src_1, Array("[" & fld_name & "]"))
                Next fld_idx
                
                .Close
                
            End With 'RS_Tbl_src
        End If
        

        Set RS_Tbl_src = .OpenRecordset(Tbl_src_2_name)
        
        With RS_Tbl_src
            If IsNull(ColSet_src_2) = True Then
                Dim NumOfColSet_Join_found As Integer
                NumOfColSet_Join_found = 0
                
                ColSet_src_2 = Array()
                    
                For fld_idx = 0 To .Fields.count - 1
                    fld_name = .Fields(fld_idx).Name

                    If NumOfColSet_Join_found <= UBound(ColSet_Join_2) And FindStrInArray(ColSet_Join_2, fld_name) > -1 Then
                        NumOfColSet_Join_found = NumOfColSet_Join_found + 1
                    Else
                        Call AppendArray(ColSet_src_2, Array("[" & fld_name & "]"))
                    End If
                Next fld_idx
            End If

            .Close
            
        End With 'RS_Tbl_src
    End With 'CurrentDb
    

    Dim SQL_Seg_Select As String
    SQL_Seg_Select = "SELECT " & "[" & Tbl_src_1_name & "]." & Join(ColSet_src_1, ", [" & Tbl_src_1_name & "].") & ", " & "[" & Tbl_src_2_name & "]." & Join(ColSet_src_2, ", [" & Tbl_src_2_name & "].")

    Dim SQL_Seg_JoinOn As String
    SQL_Seg_JoinOn = "("

    For Col_Idx = LBound(ColSet_Join_1) To UBound(ColSet_Join_1)
        SQL_Seg_JoinOn = SQL_Seg_JoinOn & "[" & Tbl_src_1_name & "].[" & ColSet_Join_1(Col_Idx) & "] = [" & Tbl_src_2_name & "].[" & ColSet_Join_2(Col_Idx) & "] AND "
    Next Col_Idx

    SQL_Seg_JoinOn = Left(SQL_Seg_JoinOn, Len(SQL_Seg_JoinOn) - 4) & ")"

    
    Dim SQL_Seg_OrderBy As String
    SQL_Seg_OrderBy = ""
    
    If IsNull(ColSet_Order) = False Then
        SQL_Seg_OrderBy = "ORDER BY "
        
        For Col_Idx = LBound(ColSet_Order) To UBound(ColSet_Order)
            SQL_Seg_OrderBy = SQL_Seg_OrderBy & "[" & Tbl_src_1_name & "].[" & ColSet_Order(Col_Idx) & "], "
        Next Col_Idx
        
        SQL_Seg_OrderBy = Left(SQL_Seg_OrderBy, Len(SQL_Seg_OrderBy) - 2)
        
    End If
    
    
    Dim SQL_cmd As String
    
    SQL_cmd = SQL_Seg_Select & " " & vbCrLf & _
                "INTO [" & Tbl_des_name & "] " & vbCrLf & _
                "FROM [" & Tbl_src_1_name & "] " & JoinCond & " JOIN [" & Tbl_src_2_name & "] " & vbCrLf & _
                "ON " & SQL_Seg_JoinOn & vbCrLf & _
                SQL_Seg_OrderBy & " " & vbCrLf & _
                ";"

    RunSQL_CmdWithoutWarning (SQL_cmd)


Exit_CreateTbl_JoinTwoTbl:
    CreateTbl_JoinTwoTbl = FailedReason
    Exit Function

Err_CreateTbl_JoinTwoTbl:
    FailedReason = Err.Description
    Resume Exit_CreateTbl_JoinTwoTbl
    
End Function

'Create table which is cancatenated from multiple tables of the same structure
Public Function CreateTbl_ConcatTbls(Tbl_src_Set As Variant, Tbl_des_name As String, Optional Type_Set As Variant = "") As String
    On Error GoTo Err_CreateTbl_ConcatTbls
    
    Dim FailedReason As String

    If UBound(Tbl_src_Set) < 0 Then
        FailedReason = "No table in the table set"
        GoTo Exit_CreateTbl_ConcatTbls
    End If

    
    Dim Tbl_src_name As Variant
    
    For Each Tbl_src_name In Tbl_src_Set
        If TableExist(Tbl_src_name & "") = False Then
            FailedReason = Tbl_src_name & " does not exist!"
            GoTo Exit_CreateTbl_ConcatTbls
        End If
    
    Next
    

    'Initialize Tbl_des
    DelTable (Tbl_des_name)
    
    Dim SQL_cmd As String
    
    Tbl_src_name = Tbl_src_Set(0)
    
    SQL_cmd = "SELECT " & Chr(34) & "null" & Chr(34) & " AS [Type], " & Tbl_src_name & ".* " & vbCrLf & _
                "INTO " & Tbl_des_name & " " & vbCrLf & _
                "FROM " & Tbl_src_name & " " & vbCrLf & _
                "WHERE 1 = 0 " & vbCrLf & _
                ";"
                
    RunSQL_CmdWithoutWarning (SQL_cmd)


    'Start Append
    Dim tbl_idx As Integer
    Dim SQL_Seq_Type As String
    
    For tbl_idx = 0 To UBound(Tbl_src_Set)
        Tbl_src_name = Tbl_src_Set(tbl_idx)
        
        If VarType(Type_Set) > vbArray And Type_Set(tbl_idx) = "" Then
            SQL_Seq_Type = ""
        Else
            SQL_Seq_Type = Chr(34) & Type_Set(tbl_idx) & Chr(34) & " AS [Type], "
        End If
        
        SQL_cmd = "INSERT INTO " & Tbl_des_name & " " & vbCrLf & _
                    "SELECT " & SQL_Seq_Type & "[" & Tbl_src_name & "].* " & vbCrLf & _
                    "FROM [" & Tbl_src_name & "] " & vbCrLf & _
                    ";"
        
        RunSQL_CmdWithoutWarning (SQL_cmd)
        
    Next
    
    
    If UBound(Type_Set) < 0 Then
        SQL_cmd = "ALTER TABLE [" & Tbl_des_name & "] " & vbCrLf & _
                    "DROP COLUMN [Type]" & vbCrLf & _
                    ";"
 
        RunSQL_CmdWithoutWarning (SQL_cmd)
        
    End If
    
    
Exit_CreateTbl_ConcatTbls:
    CreateTbl_ConcatTbls = FailedReason
    Exit Function

Err_CreateTbl_ConcatTbls:
    FailedReason = Err.Description
    Resume Exit_CreateTbl_ConcatTbls
    
End Function

'Execute SQLite Command Set
Public Function ExecuteSQLiteCmdSet(SQLiteDb_path As String, CmdSet As String) As String
    On Error GoTo Err_ExecuteSQLiteCmdSet
    
    Dim FailedReason As String

    If FileExists(SQLiteDb_path) = False Then
        FailedReason = SQLiteDb_path
        GoTo Exit_ExecuteSQLiteCmdSet
    End If
    
    
    'Create a SQLite Command file, and then parse it into the Python SQLite Command Parser for execution
    Dim SQLiteCmdFile_path As String
    Dim iFileNum_SQLiteCmd As Integer
    
    SQLiteCmdFile_path = [CurrentProject].[Path] & "\" & "SQLiteCmd.txt"
    iFileNum_SQLiteCmd = FreeFile()
    
    If FileExists(SQLiteCmdFile_path) = True Then
        Kill SQLiteCmdFile_path
    End If
    
    Open SQLiteCmdFile_path For Output As iFileNum_SQLiteCmd
    Print #iFileNum_SQLiteCmd, CmdSet
    Close #iFileNum_SQLiteCmd
    
    
    Dim SQLiteCmdLog_path As String
    SQLiteCmdLog_path = [CurrentProject].[Path] & "\SQLiteCmd.log"
    
    If FileExists(SQLiteCmdLog_path) = True Then
        Kill SQLiteCmdLog_path
    End If
    
    
    ShellCmd = "python " & [CurrentProject].[Path] & "\SQLiteCmdParser.py " & SQLiteDb_path & " " & SQLiteCmdFile_path & " " & SQLiteCmdLog_path
    Call ShellAndWait(ShellCmd, vbHide)

    If FileExists(SQLiteCmdLog_path) = False Then
        FailedReason = "SQLiteCmdParser"
        GoTo Exit_ExecuteSQLiteCmdSet
    End If
    
    
    Dim iFileNum_SQLiteCmdLog As Integer
    Dim SQLiteCmdLog_line As String
    
    iFileNum_SQLiteCmdLog = FreeFile()
    
    Open SQLiteCmdLog_path For Input As iFileNum_SQLiteCmdLog
    
    If Not EOF(iFileNum_SQLiteCmdLog) Then
        Line Input #iFileNum_SQLiteCmdLog, SQLiteCmdLog_line
    End If
    
    If SQLiteCmdLog_line <> "done" Then
        FailedReason = SQLiteCmdLog_path
        GoTo Exit_ExecuteSQLiteCmdSet
    End If
    
    Close iFileNum_SQLiteCmdLog
    
    Kill SQLiteCmdFile_path
    Kill SQLiteCmdLog_path


Exit_ExecuteSQLiteCmdSet:
    ExecuteSQLiteCmdSet = FailedReason
    Exit Function

Err_ExecuteSQLiteCmdSet:
    Call ShowMsgBox(Err.Description)
    Resume Exit_ExecuteSQLiteCmdSet
    
End Function


'Append Table into a SQLite database
Public Function AppendTblToSQLite(Tbl_src_name As String, Tbl_des_name As String) As String
    On Error GoTo Err_AppendTblToSQLite
    
    Dim FailedReason As String
    
    If TableExist(Tbl_src_name) = False Then
        FailedReason = Tbl_src_name
        GoTo Exit_AppendTblToSQLite
    End If
    
    If TableExist(Tbl_des_name) = False Then
        FailedReason = Tbl_des_name
        GoTo Exit_AppendTblToSQLite
    End If
    
    
    'Create Db
    Dim TempDb_path As String
    TempDb_path = [CurrentProject].[Path] & "\TempDb.mdb"
    
    If FileExists(TempDb_path) = True Then
        Kill TempDb_path
    End If
    
    Call CreateDatabase(TempDb_path, dbLangGeneral)
    
    
    'Copy Table into the TempDb
    Dim SQL_cmd As String
    
    SQL_cmd = "SELECT * " & vbCrLf & _
                "INTO [" & Tbl_des_name & "]" & vbCrLf & _
                "IN '" & TempDb_path & "'" & vbCrLf & _
                "FROM [" & Tbl_src_name & "] " & vbCrLf & _
                ";"
    
    RunSQL_CmdWithoutWarning (SQL_cmd)


    'Convert TempDb into SQLite
    Dim SQLiteDb_path As String
    SQLiteDb_path = [CurrentProject].[Path] & "\TempDb.sqlite"
    
    If FileExists(SQLiteDb_path) = True Then
        Kill SQLiteDb_path
    End If
    
    Dim ShellCmd As String
    ShellCmd = "java -jar " & [CurrentProject].[Path] & "\mdb-sqlite.jar " & TempDb_path & " " & SQLiteDb_path
    Call ShellAndWait(ShellCmd, vbHide)
    
    SQL_cmd = "ATTACH """ & SQLiteDb_path & """ AS TempDb;" & vbCrLf & _
                "INSERT INTO [" & Tbl_des_name & "] SELECT * FROM TempDb.[" & Tbl_des_name & "];"
    
    FailedReason = ExecuteSQLiteCmdSet(GetLinkTblConnInfo(Tbl_des_name, "DATABASE"), SQL_cmd)
    
    If FailedReason <> "" Then
        GoTo Exit_AppendTblToSQLite
    End If
    
    
    Kill SQLiteDb_path
    Kill TempDb_path


Exit_AppendTblToSQLite:
    AppendTblToSQLite = FailedReason
    Exit Function

Err_AppendTblToSQLite:
    Call ShowMsgBox(Err.Description)
    Resume Exit_AppendTblToSQLite
    
End Function