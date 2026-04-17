VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCola"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Argentum Online 0.14.0
'Copyright (C) 2002 Márquez Pablo Ignacio
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
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit
'                    Metodos publicos
'
' Public sub Push(byval I as variant) mete el elemento I
' al final de la cola.
'
' Public Function Pop As Variant: quita de la cola el primer elem
' y lo devuelve
'
' Public Function VerElemento(ByVal Index As Integer) As Variant
' muestra el elemento numero Index de la cola sin quitarlo
'
' Public Function PopByVal() As Variant: muestra el primer
' elemento de la cola sin quitarlo
'
' Public Property Get Longitud() As Integer: devuelve la
' cantidad de elementos que tiene la cola.

Private Const FRENTE = 1

Private MAXELEM As Integer

Private Cola As Collection

Public Property Let MaxLenght(ByVal Max As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 14/11/2010
'Sets the Max queue lenght.
'***************************************************
    MAXELEM = Max
End Property

Public Sub Reset()
On Error GoTo ErrHandler
  
On Error Resume Next

Dim I As Integer
For I = 1 To Me.Longitud
    Cola.Remove FRENTE
Next I

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Reset de ModCola.cls")
End Sub

Public Property Get Longitud() As Integer
Longitud = Cola.Count
End Property

Private Function IndexValido(ByVal I As Integer) As Boolean
On Error GoTo ErrHandler
  
IndexValido = I >= 1 And I <= Me.Longitud
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function IndexValido de ModCola.cls")
End Function

Private Sub Class_Initialize()
Set Cola = New Collection
  
End Sub

Public Function VerElemento(ByVal Index As Integer, Optional ByVal ConvertUpper As Boolean = True) As String
On Error GoTo ErrHandler
  
On Error Resume Next
    If IndexValido(Index) Then
        'Pablo
        If ConvertUpper Then
            VerElemento = UCase$(Cola.Item(Index))
        Else
            VerElemento = Cola.Item(Index)
        End If
        
        '/Pablo
        'VerElemento = Cola(Index)
    Else
        VerElemento = 0
    End If
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function VerElemento de ModCola.cls")
End Function


Public Sub Push(ByVal Detalle As String, Optional ByVal ConvertUpper As Boolean = True)
'***************************************************
'Author: Unknown
'Last Modification: 14/11/2010
'14/11/2010: ZaMa - Ahora se valida si supera un maximo previamente establecido.
'***************************************************
On Error GoTo ErrHandler
  
On Error Resume Next
    'Mete elemento en la cola
    
    Dim aux As String
    aux = Time$ & " "
    
    If ConvertUpper Then
        aux = aux & UCase$(Detalle)
    Else
        aux = aux & Detalle
    End If
    
    ' Validate that not exceed Max lenght
    If MAXELEM <> 0 Then
        If Cola.Count = MAXELEM Then
            Call Cola.Remove(FRENTE) ' Quito el primer elemento
        End If
    End If
    
    Call Cola.Add(aux)
        
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Push de ModCola.cls")
End Sub

Public Function PopByVal() As String
On Error GoTo ErrHandler
  
On Error Resume Next

'Quita elemento de la cola
If Cola.Count > 0 Then
    PopByVal = Cola.Item(1)
Else
    PopByVal = 0
End If

  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function PopByVal de ModCola.cls")
End Function

Public Function Existe(ByVal Nombre As String) As Boolean
On Error GoTo ErrHandler
  
On Error Resume Next
  
 Dim V As String
 Dim I As Integer
 Dim NombreEnMayusculas As String
 NombreEnMayusculas = UCase$(Nombre)
 
 For I = 1 To Me.Longitud
        'Pablo
        V = mid$(Me.VerElemento(I), 10, Len(Me.VerElemento(I)))
        '/Pablo
        'V = Me.VerElemento(I)
        If V = NombreEnMayusculas Then
            Existe = True
            Exit Function
        End If
 Next
 Existe = False

  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function Existe de ModCola.cls")
End Function

Public Sub Quitar(ByVal Nombre As String)
On Error GoTo ErrHandler
  
On Error Resume Next
Dim V As String
Dim I As Integer
Dim NombreEnMayusculas As String
 
NombreEnMayusculas = UCase$(Nombre)
 
For I = 1 To Me.Longitud
    'Pablo
    V = mid$(Me.VerElemento(I), 10, Len(Me.VerElemento(I)))
    '/Pablo
    'V = Me.VerElemento(I)
    If V = NombreEnMayusculas Then
        Call Cola.Remove(I)
        Exit Sub
    End If
Next I
 
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Quitar de ModCola.cls")
End Sub

Private Sub Class_Terminate()
'Destruimos el objeto Cola
On Error GoTo ErrHandler
  
Set Cola = Nothing
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Class_Terminate de ModCola.cls")
End Sub
