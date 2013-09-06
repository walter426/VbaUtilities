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

Attribute VB_Name = "StrUtilities"
Option Compare Database

'Split a string into array by separator
Public Function SplitStrIntoArray(str As String, separator As String) As Variant
    Dim Arr As Variant
    
    If Len(str) > 0 Then
        Arr = Split(str, separator)
        
        Dim i As Integer
        
        For i = 0 To UBound(Arr)
            Arr(i) = Trim(Arr(i))
        Next i
    Else
        Arr = Array()
    End If
    
    SplitStrIntoArray = Arr
    
End Function

'Find string in an array
Public Function FindStrInArray(Array_str As Variant, str As String) As Integer
    FindStrInArray = -1


    Dim i As Integer
    
    For i = 0 To UBound(Array_str)
        If str = Array_str(i) Then
            FindStrInArray = i
            Exit For
        End If
    Next i
    
End Function

'Replace substring by regular expression
'Ref: Microsoft VBScript Regular Expressions 5.5
Public Function Replace_RE(str As String, Pattern_f As String, substr_r As String) As String
    On Error GoTo Exit_Replace_RE
    
    Replace_RE = str
    
    Dim RE As RegExp
    Set RE = CreateObject("vbscript.regexp")
    
    With RE
        .MultiLine = True
        .Global = True
        .IgnoreCase = False
        .Pattern = Pattern_f
        
        Replace_RE = .Replace(str, substr_r)
        
    End With
    
Exit_Replace_RE:
    Exit Function

Err_Replace_RE:
    ShowMsgBox (Err.Description)
    Resume Exit_Replace_RE
End Function

