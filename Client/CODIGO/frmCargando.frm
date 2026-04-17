VERSION 5.00
Begin VB.Form frmCargando 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000000&
   BorderStyle     =   0  'None
   ClientHeight    =   7650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10020
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmCargando.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   510
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   668
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image imgExtras 
      Appearance      =   0  'Flat
      Height          =   900
      Left            =   8160
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   900
   End
   Begin VB.Image imgSonido 
      Appearance      =   0  'Flat
      Height          =   900
      Left            =   6960
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   900
   End
   Begin VB.Image imgClases 
      Appearance      =   0  'Flat
      Height          =   900
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   900
   End
   Begin VB.Image imgDats 
      Appearance      =   0  'Flat
      Height          =   900
      Left            =   4560
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   900
   End
   Begin VB.Image imgMotorGrafico 
      Appearance      =   0  'Flat
      Height          =   900
      Left            =   960
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   900
   End
   Begin VB.Image imgAnimaciones 
      Appearance      =   0  'Flat
      Height          =   900
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   900
   End
   Begin VB.Image imgConstantes 
      Appearance      =   0  'Flat
      Height          =   900
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   900
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4080
      TabIndex        =   0
      Top             =   5880
      Width           =   3135
   End
End
Attribute VB_Name = "frmCargando"
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
  
Private Sub Form_Load()
On Error GoTo ErrHandler
    Me.Picture = LoadPicture(DirInterfaces & SELECTED_UI & "VentanaCargando.jpg")
    
    Call modCustomCursors.SetFormCursorDefault(Me)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Form_Load de frmCargando.frm")
End Sub

