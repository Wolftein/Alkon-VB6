VERSION 5.00
Begin VB.Form frmCantidad 
   BackColor       =   &H80000000&
   BorderStyle     =   0  'None
   ClientHeight    =   1440
   ClientLeft      =   1635
   ClientTop       =   4410
   ClientWidth     =   3240
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   96
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   216
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCantidad 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   525
      MaxLength       =   5
      TabIndex        =   0
      Top             =   450
      Width           =   2220
   End
   Begin VB.Image imgTirarTodo 
      Height          =   375
      Left            =   1680
      Tag             =   "1"
      Top             =   915
      Width           =   1320
   End
   Begin VB.Image imgTirar 
      Height          =   375
      Left            =   225
      Tag             =   "1"
      Top             =   915
      Width           =   1320
   End
End
Attribute VB_Name = "frmCantidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private clsFormulario As clsFormMovementManager

Private cBotonTirar As clsGraphicalButton
Private cBotonTirarTodo As clsGraphicalButton
Public LastButtonPressed As clsGraphicalButton

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
On Error GoTo ErrHandler
  
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
     Me.Picture = LoadPicture(DirInterfaces & SELECTED_UI & "VentanaCantidad.jpg")
    
    Call LoadButtons
    
    Call modCustomCursors.SetFormCursorDefault(Me)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Form_Load de frmCantidad.frm")
End Sub

Private Sub LoadButtons()
On Error GoTo ErrHandler
  

    Dim GrhPath As String
    
    GrhPath = DirInterfaces & SELECTED_UI
    
    Set cBotonTirar = New clsGraphicalButton
    Set cBotonTirarTodo = New clsGraphicalButton
    
    Set LastButtonPressed = New clsGraphicalButton

    Call cBotonTirar.Initialize(imgTirar, GrhPath & "BotonTirar.jpg", _
                                        GrhPath & "BotonTirarRollover.jpg", _
                                        GrhPath & "BotonTirarClick.jpg", Me)
                                    
    Call cBotonTirarTodo.Initialize(imgTirarTodo, GrhPath & "BotonTirarTodo.jpg", _
                                        GrhPath & "BotonTirarTodoRollover.jpg", _
                                        GrhPath & "BotonTirarTodoClick.jpg", Me)

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadButtons de frmCantidad.frm")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub imgTirar_Click()
On Error GoTo ErrHandler
  
    If LenB(txtCantidad.text) > 0 Then
        If Not IsNumeric(txtCantidad.text) Then Exit Sub  'Should never happen
        Call WriteDrop(Inventario.SelectedItem, frmCantidad.txtCantidad.text)
        frmCantidad.txtCantidad.text = ""
    End If
    
    CerrarVentana
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgTirar_Click de frmCantidad.frm")
End Sub

Private Sub imgTirarTodo_Click()
On Error GoTo ErrHandler
  
    If Inventario.SelectedItem = 0 Then Exit Sub
    
    If Inventario.SelectedItem <> FLAGORO Then
        Call WriteDrop(Inventario.SelectedItem, Inventario.Amount(Inventario.SelectedItem))
        CerrarVentana
    Else
        If UserGLD > 10000 Then
            Call WriteDrop(Inventario.SelectedItem, 10000)
            CerrarVentana
        Else
            Call WriteDrop(Inventario.SelectedItem, UserGLD)
            CerrarVentana
        End If
    End If

    frmCantidad.txtCantidad.text = ""
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgTirarTodo_Click de frmCantidad.frm")
End Sub

Private Sub txtCantidad_Change()
On Error GoTo ErrHandler
    If Val(txtCantidad.text) < 0 Then
        txtCantidad.text = "1"
    End If
    
    If Val(txtCantidad.text) > MAX_INVENTORY_OBJS Then
        txtCantidad.text = "10000"
    End If
    
    Exit Sub
    
ErrHandler:
    'If we got here the user may have pasted (Shift + Insert) a REALLY large number, causing an overflow, so we set amount back to 1
    txtCantidad.text = "1"
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
On Error GoTo ErrHandler
  
    If (KeyAscii <> 8) Then
        If (KeyAscii < 48 Or KeyAscii > 57) Then
            KeyAscii = 0
        End If
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub txtCantidad_KeyPress de frmCantidad.frm")
End Sub

Private Sub txtCantidad_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
  
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then CerrarVentana
End Sub

Private Sub CerrarVentana()
On Error GoTo ErrHandler
  
    Unload Me
    If frmMain.Visible Then frmMain.SetFocus
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CerrarVentana de frmCantidad.frm")
End Sub
