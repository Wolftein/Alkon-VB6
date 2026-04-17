Attribute VB_Name = "Matematicas"
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

Public Function Min(ByVal First As Long, ByVal Second As Long) As Long

    If First <= Second Then
        Min = First
    Else
        Min = Second
    End If

End Function

Public Function Max(ByVal First As Long, ByVal Second As Long) As Long

    If First >= Second Then
        Max = First
    Else
        Max = Second
    End If

End Function

Public Function Porcentaje(ByVal Total As Long, ByVal Porc As Long) As Long
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    Porcentaje = (Total * Porc) / 100
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function Porcentaje de Matematicas.bas")
End Function

Function Distancia(ByRef wp1 As WorldPos, ByRef wp2 As WorldPos) As Long
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    'Encuentra la distancia entre dos WorldPos
    Distancia = Abs(wp1.X - wp2.X) + Abs(wp1.Y - wp2.Y) + (Abs(wp1.Map - wp2.Map) * 100)
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function Distancia de Matematicas.bas")
End Function

Function Distance(ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer) As Double
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

'Encuentra la distancia entre dos puntos

Distance = Sqr(((Y1 - Y2) ^ 2 + (X1 - X2) ^ 2))

  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function Distance de Matematicas.bas")
End Function

Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/06/2006
'Generates a random number in the range given - recoded to use longs and work properly with ranges
'**************************************************************
On Error GoTo ErrHandler
  
    RandomNumber = Fix(rnd * (UpperBound - LowerBound + 1)) + LowerBound
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function RandomNumber de Matematicas.bas")
End Function

Function RandomDecimalNumber(ByVal LowerBound As Long, ByVal UpperBound As Variant) As Double
'**************************************************************
'Author: Nightw
'Last Modify Date: 23/06/2015
'Generates a random decimal number in the range given
'**************************************************************
On Error GoTo ErrHandler
  

    RandomDecimalNumber = rnd * (UpperBound - LowerBound) + LowerBound

  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function RandomDecimalNumber de Matematicas.bas")
End Function

