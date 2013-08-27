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

'Find item in an array
Public Function FindItemInArray(Array_src As Variant, item As String) As Long
    FindItemInArray = -1


    Dim i As Long
    
    For i = 0 To UBound(Array_src)
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
    
    For i = 0 To UBound(Array_append)
        ReDim Preserve Array_src(0 To UBound(Array_src) + 1)
        Array_src(UBound(Array_src)) = Array_append(i)
    Next i


Exit_AppendArray:
    AppendArray = FailedReason
    Exit Function

Err_AppendArray:
    FailedReason = Err.Description
    Resume Exit_AppendArray
    
End Function
