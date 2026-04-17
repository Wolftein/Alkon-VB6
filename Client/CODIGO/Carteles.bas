Attribute VB_Name = "Carteles"
'Argentum Online 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
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

Private Const XPosCartel As Integer = 100
Private Const YPosCartel As Integer = 100
Private Const MAXLONG As Integer = 36

'Carteles
Public Cartel As Boolean
Public Leyenda As String
Public LeyendaFormateada() As String
Public textura As Integer


Sub InitCartel(Ley As String, Grh As Integer)
On Error GoTo ErrHandler
  
If Not Cartel Then
    Leyenda = Ley
    textura = Grh
    Cartel = True
    ReDim LeyendaFormateada(0 To (Len(Ley) \ (MAXLONG \ 2)))
                
    Dim I As Integer, k As Integer, anti As Integer
    anti = 1
    k = 0
    I = 0
    Call DarFormato(Leyenda, I, k, anti)
    I = 0
    Do While LeyendaFormateada(I) <> "" And I < UBound(LeyendaFormateada)
        
       I = I + 1
    Loop
    ReDim Preserve LeyendaFormateada(0 To I - 1)
Else
    Exit Sub
End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub InitCartel de Carteles.bas")
End Sub


Private Function DarFormato(s As String, I As Integer, k As Integer, anti As Integer)
On Error GoTo ErrHandler
  
If anti + I <= Len(s) + 1 Then
    If ((I >= MAXLONG) And mid$(s, anti + I, 1) = " ") Or (anti + I = Len(s)) Then
        LeyendaFormateada(k) = mid(s, anti, I + 1)
        k = k + 1
        anti = anti + I + 1
        I = 0
    Else
        I = I + 1
    End If
    Call DarFormato(s, I, k, anti)
End If
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function DarFormato de Carteles.bas")
End Function

Private Function XOffSetFromTexture(Id As Integer) As Integer
On Error GoTo ErrHandler
  
    If Id = 4987 Then
        XOffSetFromTexture = 10
    End If
    If Id = 514 Then
        XOffSetFromTexture = 20
    End If
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function XOffSetFromTexture de Carteles.bas")
End Function

Private Function YOffSetFromTexture(Id As Integer) As Integer
On Error GoTo ErrHandler
  
    If Id = 4987 Then
        YOffSetFromTexture = 20
    End If
    If Id = 514 Then
        YOffSetFromTexture = 55
    End If
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function YOffSetFromTexture de Carteles.bas")
End Function

Sub DibujarCartel()
On Error GoTo ErrHandler
  
If Not Cartel Then Exit Sub
Dim X As Integer, Y As Integer
X = XPosCartel + XOffSetFromTexture(textura)
Y = YPosCartel + YOffSetFromTexture(textura)
Call Mod_TileEngine.DrawGrhIndex(textura, XPosCartel, YPosCartel, 0#, 0)
Dim J As Integer, OffsetX As Integer, OffsetY As Integer

OffsetX = X + FuentesJuego.FuenteBase.Tamanio
OffsetY = Y + FuentesJuego.FuenteBase.Tamanio

For J = 0 To UBound(LeyendaFormateada)
    DrawText OffsetX, OffsetY, 0#, LeyendaFormateada(J), &HFFFFFFFF, eRendererAlignmentLeftTop, FuentesJuego.FuenteBase
    OffsetY = OffsetY + FuentesJuego.FuenteBase.Tamanio
Next
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub DibujarCartel de Carteles.bas")
End Sub

