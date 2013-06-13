Attribute VB_Name = "General Utilities"
Option Compare Database
Option Explicit

Public NotShowMsgBox As Boolean

Public Function EnableMsgBox()
    NotShowMsgBox = False
End Function

Public Function DisableMsgBox()
    NotShowMsgBox = True
End Function

Public Function ShowMsgBox(str As String) As Boolean
    
    If NotShowMsgBox = False Then
        MsgBox str
    End If
    
    ShowMsgBox = NotShowMsgBox
    
End Function

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
