Attribute VB_Name = "Queue"
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

Public Type tVertice
    X As Integer
    Y As Integer
End Type

Private Const MAXELEM As Integer = 1000

Private m_array() As tVertice
Private m_lastelem As Integer
Private m_firstelem As Integer
Private m_size As Integer

Public Function IsEmpty() As Boolean
On Error GoTo ErrHandler
  
IsEmpty = m_size = 0
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function IsEmpty de Queue.bas")
End Function

Public Function IsFull() As Boolean
On Error GoTo ErrHandler
  
IsFull = m_lastelem = MAXELEM
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function IsFull de Queue.bas")
End Function

Public Function Push(ByRef Vertice As tVertice) As Boolean
On Error GoTo ErrHandler
  

If Not IsFull Then
    
    If IsEmpty Then m_firstelem = 1
    
    m_lastelem = m_lastelem + 1
    m_size = m_size + 1
    m_array(m_lastelem) = Vertice
    
    Push = True
Else
    Push = False
End If

  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function Push de Queue.bas")
End Function

Public Function Pop() As tVertice
On Error GoTo ErrHandler
  

If Not IsEmpty Then
    
    Pop = m_array(m_firstelem)
    m_firstelem = m_firstelem + 1
    m_size = m_size - 1
    
    If m_firstelem > m_lastelem And m_size = 0 Then
            m_lastelem = 0
            m_firstelem = 0
            m_size = 0
    End If
   
End If

  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function Pop de Queue.bas")
End Function

Public Sub InitQueue()
On Error GoTo ErrHandler
  
ReDim m_array(MAXELEM) As tVertice
m_lastelem = 0
m_firstelem = 0
m_size = 0
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub InitQueue de Queue.bas")
End Sub

