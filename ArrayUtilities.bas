Attribute VB_Name = "ArrayUtilities"
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

Option Compare Database

Public Function GetItemInArray(Array_src As Variant, idx As Long) As Variant
    GetItemInArray = Array_src(idx)
End Function

Public Function GetItemInStrArray(StrArray_src As String, separator As String, idx As Long) As Variant
    GetItemInStrArray = GetItemInArray(SplitStrIntoArray(StrArray_src, separator), idx)
End Function

'Find item in an array
Public Function FindItemInArray(Array_src As Variant, item As String) As Long
    FindItemInArray = -1


    Dim i As Long
    
    For i = LBound(Array_src) To UBound(Array_src)
        If Array_src(i) = item Then
            FindItemInArray = i
            Exit For
        End If
    Next i
    
End Function

'Append items to an Array
Public Function AppendArray(Array_src As Variant, Array_append As Variant) As String
    Dim FailedReason As String

    Dim i As Long
    
    For i = LBound(Array_append) To UBound(Array_append)
        ReDim Preserve Array_src(LBound(Array_src) To UBound(Array_src) + 1)
        Array_src(UBound(Array_src)) = Array_append(i)
    Next i


Exit_AppendArray:
    AppendArray = FailedReason
    Exit Function

Err_AppendArray:
    FailedReason = Err.Description
    Resume Exit_AppendArray
    
End Function

'Delete item in an array by index
Public Sub DeleteArrayItem(arr As Variant, index As Long)
    Dim i As Long
    
    For i = index To UBound(arr) - 1
        arr(i) = arr(i + 1)
    Next
    
    ' VB will convert this to 0 or to an empty string.
    arr(UBound(arr)) = Empty
    ReDim Preserve arr(LBound(arr) To UBound(arr) - 1)
    
End Sub
