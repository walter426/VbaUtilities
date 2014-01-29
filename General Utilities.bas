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

Attribute VB_Name = "General Utilities"
Option Compare Database
Option Explicit

Public NotShowMsgBox As Boolean

'Enable user-defined MsgBox
Public Function EnableMsgBox()
    NotShowMsgBox = False
End Function

'Disable user-defined MsgBox
Public Function DisableMsgBox()
    NotShowMsgBox = True
End Function

'Display string in a msgbox depending on the user-defined flag
Public Function ShowMsgBox(str As String) As Boolean
    
    If NotShowMsgBox = False Then
        MsgBox str
    End If
    
    ShowMsgBox = NotShowMsgBox
    
End Function