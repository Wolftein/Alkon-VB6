VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "diccionario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************
' diccionario.cls
'
' Designed and implemented by Mariono Barrou (El Oso)
'**************************************************************

'**************************************************************************
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

'clase diccionario
'basico, plain sin queso ni papa fritas


'mi idea cuando hice esto, lo encontre en el rigido :p. Hecha por el oso

'****
' ZaMa: Como se usaba solo para contar votos en guilds, lo adapte para que tenga mejor performanse.
' En realidad se podian usar 2 rutinas nada mas, la que agregaba el voto (clave) y la que cuenta y devuelve
' a los ganadores :S.
'****
Option Explicit

Private Const MAX_ELEM = 100

Private Type tDiccElem
    key As String
    Counter As Integer
End Type

Private p_elementos(1 To MAX_ELEM) As tDiccElem 'visual basic es una mierda para usar memoria dinamica, asi que uso esto
Private NumKeys As Integer

Public Property Get CantElem() As Integer
    CantElem = NumKeys
End Property

Public Sub AddKey(ByVal key As String)
'***************************************************
'Author: ZaMa
'Last Modification: 14/11/2010
'Increses the key counter if key already exists, else creates it.
'***************************************************
On Error GoTo ErrHandler
  
    
    Dim KeyIndex As Long
    Dim Found As Boolean
    
    If LenB(key) = 0 Then Exit Sub
    
    key = UCase$(key)
    
    If NumKeys < MAX_ELEM Then
        
        For KeyIndex = 1 To NumKeys
            If key = p_elementos(KeyIndex).key Then
                ' Found it, increase counter
                p_elementos(KeyIndex).Counter = p_elementos(KeyIndex).Counter + 1
                Found = True
                Exit For
            End If
        Next KeyIndex
        
        If Not Found Then
            NumKeys = NumKeys + 1
            p_elementos(NumKeys).Counter = 1
            p_elementos(NumKeys).key = UCase$(key)
        End If
    
    End If
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub AddKey de clsdicc.cls")
End Sub

Public Function At(ByVal key As String) As Integer
On Error GoTo ErrHandler
  
Dim I As Integer

    key = UCase$(key)

    At = Null
    For I = 1 To NumKeys
        If key = p_elementos(I).key Then
            At = p_elementos(I).Counter
            Exit Function
        End If
    Next I

  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function At de clsdicc.cls")
End Function

Public Function AtIndex(ByVal I As Integer) As String
On Error GoTo ErrHandler
  
    AtIndex = p_elementos(I).key
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function AtIndex de clsdicc.cls")
End Function

Public Function MaxKeyValue(ByRef MaxValueCounter As Integer) As String
'***************************************************
'Author: ZaMa
'Last Modification: 14/11/2010
'Returns the keys which have the biggest values.
'***************************************************
On Error GoTo ErrHandler
  
    
    Dim KeyIndex As Long
    Dim MaxValue As Integer
    Dim key As String
    
    MaxValue = -1
    
    For KeyIndex = 1 To NumKeys
        
        If MaxValue < p_elementos(KeyIndex).Counter Then
            ' New max
            MaxValueCounter = 1
            key = p_elementos(KeyIndex).key
            MaxValue = p_elementos(KeyIndex).Counter
        
        ElseIf MaxValue = p_elementos(KeyIndex).Counter Then
            ' Key value already exists
            MaxValueCounter = MaxValueCounter + 1
            key = key & "," & p_elementos(KeyIndex).key
        End If
        
    Next KeyIndex

    MaxKeyValue = key

  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function MaxKeyValue de clsdicc.cls")
End Function

Public Sub DumpAll()
On Error GoTo ErrHandler
  
Dim I As Integer

For I = 1 To MAX_ELEM
    p_elementos(I).key = vbNullString
    p_elementos(I).Counter = Null
Next I
NumKeys = 0

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub DumpAll de clsdicc.cls")
End Sub
