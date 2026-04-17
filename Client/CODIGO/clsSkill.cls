VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSkill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************************
'Argentum Online
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'**************************************************************************

'
' Module author: D'Artagnan (14/04/2015)
'

Option Explicit

Private bNaturalPoints As Byte
Private bAssignedPoints As Byte

Public Sub Initialize(Optional ByVal bNatural As Byte = 0, Optional ByVal bAssigned As Byte = 0)
'******************************************
'Author: D'Artagnan
'Date: 14/04/2015
'
'******************************************
On Error GoTo ErrHandler
  
    bNaturalPoints = bNatural
    bAssignedPoints = bAssigned
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Initialize de clsSkill.cls")
End Sub

Public Property Get Assigned() As Byte
'******************************************
'Author: D'Artagnan
'Date: 14/04/2015
'
'******************************************
    Assigned = bAssignedPoints
End Property

Public Property Let Assigned(ByVal value As Byte)
'******************************************
'Author: D'Artagnan
'Date: 14/04/2015
'
'******************************************
    bAssignedPoints = value
End Property

Public Property Get Both() As Byte
'******************************************
'Author: D'Artagnan
'Date: 14/04/2015
'
'******************************************
    Both = bNaturalPoints + bAssignedPoints
End Property

Public Property Get Natural() As Byte
'******************************************
'Author: D'Artagnan
'Date: 14/04/2015
'
'******************************************
    Natural = bNaturalPoints
End Property

Public Property Let Natural(ByVal value As Byte)
'******************************************
'Author: D'Artagnan
'Date: 14/04/2015
'
'******************************************
    bNaturalPoints = value
End Property

Public Sub Reset()
'******************************************
'Author: D'Artagnan
'Date: 14/04/2015
'
'******************************************
On Error GoTo ErrHandler
  
    bNaturalPoints = 0
    bAssignedPoints = 0
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Reset de clsSkill.cls")
End Sub
