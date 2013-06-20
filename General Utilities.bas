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