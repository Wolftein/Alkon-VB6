VERSION 5.00
Begin VB.Form frmPunishmentAdm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aplicando pena..."
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6465
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   6465
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnCancel 
      Caption         =   "&Cancelar"
      Height          =   255
      Left            =   3600
      TabIndex        =   8
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton btnAccept 
      Caption         =   "&Aceptar"
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox txtNotes 
      Height          =   1455
      Left            =   1560
      MaxLength       =   255
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   2640
      Width           =   4215
   End
   Begin VB.TextBox txtReason 
      Height          =   1335
      Left            =   1560
      MaxLength       =   255
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1200
      Width           =   4215
   End
   Begin VB.ComboBox cboPunishmentList 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   660
      Width           =   4455
   End
   Begin VB.Label lblUserToPunish 
      Caption         =   "NOMBRE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   9
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label lblPunishmentTypeText 
      AutoSize        =   -1  'True
      Caption         =   "Aplicando TIPOPENA a "
      Height          =   195
      Left            =   2160
      TabIndex        =   6
      Top             =   240
      Width           =   1740
   End
   Begin VB.Label lblNotes 
      Caption         =   "Notas administrativas"
      Height          =   735
      Left            =   360
      TabIndex        =   5
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label lblReason 
      Caption         =   "Razon:"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblPunishmentType 
      Caption         =   "Tipo de Pena:"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   1095
   End
End
Attribute VB_Name = "frmPunishmentAdm"
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
Option Base 0


Public userToPunish As String
Public punishmentType As Byte


Private Sub btnAccept_Click()
    Call Engine_Audio.PlayInterface("click.wav")
On Error GoTo ErrHandler
  
    
    If cboPunishmentList.ListIndex = -1 Then
        Call MsgBox("Seleccione una pena.")
        Exit Sub
    End If
    
    Select Case punishmentType
        Case ePunishmentSubType.Ban
            Call WriteBanChar(userToPunish, txtReason.text, txtNotes.text, punishmentList(cboPunishmentList.ListIndex).Id)
        Case ePunishmentSubType.Jail
            Call WriteJail(userToPunish, txtReason.text, txtNotes.text, punishmentList(cboPunishmentList.ListIndex).Id)
        Case ePunishmentSubType.Warning
            Call WriteWarnUser(userToPunish, txtReason, txtNotes.text, punishmentList(cboPunishmentList.ListIndex).Id)
    End Select
    
    
    cboPunishmentList.Clear
    
    Unload Me
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub btnAccept_Click de frmPunishmentAdm.frm")
End Sub

Private Sub btnCancel_Click()
    Call Engine_Audio.PlayInterface("click.wav")
On Error GoTo ErrHandler
  
    cboPunishmentList.Clear
    Unload Me
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub btnCancel_Click de frmPunishmentAdm.frm")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then CloseWindow
End Sub

Private Sub Form_Load()
    Dim I As Integer
On Error GoTo ErrHandler
  
        
    For I = 0 To UBound(punishmentList) - 1
        cboPunishmentList.AddItem punishmentList(I).Name
    Next I
    
    Select Case punishmentType
        Case 1
            lblPunishmentTypeText = "Aplicando CARCEL a "
        Case 2
            lblPunishmentTypeText = "Aplicando BAN a "
    End Select
    
    lblUserToPunish = userToPunish
    
    Call modCustomCursors.SetFormCursorDefault(Me)

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Form_Load de frmPunishmentAdm.frm")
End Sub
Private Sub CloseWindow()
On Error GoTo ErrHandler
  
    Unload Me
    If frmMain.Visible Then frmMain.SetFocus
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CloseWindow de frmPunishmentAdm.frm")
End Sub

