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

Attribute VB_Name = "MathUtilities"
Option Compare Database

'Ceiling
Public Function Ceiling(X)
    Ceiling = Int(X) - (X - Int(X) > 0)
End Function

'Log on base 10
Public Function Log10(X)
    Log10 = Log(X) / Log(10#)
End Function
