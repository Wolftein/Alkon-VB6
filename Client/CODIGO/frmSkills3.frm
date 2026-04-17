VERSION 5.00
Begin VB.Form frmSkills3 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9000
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   450
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image imgMas18 
      Height          =   300
      Left            =   8100
      Top             =   3840
      Width           =   300
   End
   Begin VB.Image imgMenos18 
      Height          =   300
      Left            =   7260
      Top             =   3840
      Width           =   300
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   18
      Left            =   7635
      TabIndex        =   19
      Top             =   3840
      Width           =   405
   End
   Begin VB.Image imgSastreria 
      Height          =   330
      Left            =   4920
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Image imgCombateSinArmas 
      Height          =   330
      Left            =   4950
      Top             =   3405
      Width           =   2175
   End
   Begin VB.Image imgCombateDistancia 
      Height          =   330
      Left            =   4950
      Top             =   3030
      Width           =   2295
   End
   Begin VB.Image imgDomar 
      Height          =   330
      Left            =   4950
      Top             =   2655
      Width           =   1875
   End
   Begin VB.Image imgHerreria 
      Height          =   330
      Left            =   4950
      Top             =   2280
      Width           =   1065
   End
   Begin VB.Image imgCarpinteria 
      Height          =   330
      Left            =   4950
      Top             =   1905
      Width           =   1395
   End
   Begin VB.Image imgMineria 
      Height          =   330
      Left            =   4950
      Top             =   1530
      Width           =   1005
   End
   Begin VB.Image imgPesca 
      Height          =   330
      Left            =   4950
      Top             =   1170
      Width           =   780
   End
   Begin VB.Image imgEscudos 
      Height          =   330
      Left            =   4950
      Top             =   810
      Width           =   2295
   End
   Begin VB.Image imgTalar 
      Height          =   330
      Left            =   615
      Top             =   3780
      Width           =   780
   End
   Begin VB.Image imgSupervivencia 
      Height          =   330
      Left            =   615
      Top             =   3405
      Width           =   1740
   End
   Begin VB.Image imgOcultarse 
      Height          =   330
      Left            =   615
      Top             =   3030
      Width           =   1260
   End
   Begin VB.Image imgApunialar 
      Height          =   330
      Left            =   615
      Top             =   2655
      Width           =   1260
   End
   Begin VB.Image imgMeditar 
      Height          =   330
      Left            =   615
      Top             =   2280
      Width           =   1065
   End
   Begin VB.Image imgCombateArmas 
      Height          =   330
      Left            =   615
      Top             =   1905
      Width           =   2325
   End
   Begin VB.Image imgEvasion 
      Height          =   330
      Left            =   615
      Top             =   1530
      Width           =   2325
   End
   Begin VB.Image imgRobar 
      Height          =   330
      Left            =   615
      Top             =   1170
      Width           =   900
   End
   Begin VB.Image imgMagia 
      Height          =   330
      Left            =   615
      Top             =   810
      Width           =   870
   End
   Begin VB.Label lblHelp 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1110
      Left            =   555
      TabIndex        =   18
      Top             =   4770
      Width           =   7860
   End
   Begin VB.Image imgCancelar 
      Height          =   195
      Left            =   8655
      Top             =   165
      Width           =   195
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   3495
      TabIndex        =   17
      Top             =   840
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   3495
      TabIndex        =   16
      Top             =   1215
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   3495
      TabIndex        =   15
      Top             =   1575
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   3495
      TabIndex        =   14
      Top             =   1950
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   3495
      TabIndex        =   13
      Top             =   2325
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   3495
      TabIndex        =   12
      Top             =   2700
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   3495
      TabIndex        =   11
      Top             =   3075
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   3495
      TabIndex        =   10
      Top             =   3450
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   3495
      TabIndex        =   9
      Top             =   3825
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   7635
      TabIndex        =   8
      Top             =   840
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   11
      Left            =   7635
      TabIndex        =   7
      Top             =   1215
      Width           =   405
   End
   Begin VB.Image imgMas1 
      Height          =   300
      Left            =   3960
      Top             =   780
      Width           =   300
   End
   Begin VB.Image imgMas2 
      Height          =   300
      Left            =   3960
      Top             =   1155
      Width           =   300
   End
   Begin VB.Image imgMenos2 
      Height          =   300
      Left            =   3120
      Top             =   1155
      Width           =   300
   End
   Begin VB.Image imgMas3 
      Height          =   300
      Left            =   3960
      Top             =   1515
      Width           =   300
   End
   Begin VB.Image imgMenos3 
      Height          =   300
      Left            =   3120
      Top             =   1515
      Width           =   300
   End
   Begin VB.Image imgMas4 
      Height          =   300
      Left            =   3960
      Top             =   1890
      Width           =   300
   End
   Begin VB.Image imgMenos4 
      Height          =   300
      Left            =   3120
      Top             =   1890
      Width           =   300
   End
   Begin VB.Image imgMas5 
      Height          =   300
      Left            =   3960
      Top             =   2265
      Width           =   300
   End
   Begin VB.Image imgMenos5 
      Height          =   300
      Left            =   3120
      Top             =   2265
      Width           =   300
   End
   Begin VB.Image imgMas6 
      Height          =   300
      Left            =   3960
      Top             =   2640
      Width           =   300
   End
   Begin VB.Image imgMenos6 
      Height          =   300
      Left            =   3120
      Top             =   2640
      Width           =   300
   End
   Begin VB.Image imgMas7 
      Height          =   300
      Left            =   3960
      Top             =   3015
      Width           =   300
   End
   Begin VB.Image imgMenos7 
      Height          =   300
      Left            =   3120
      Top             =   3015
      Width           =   300
   End
   Begin VB.Image imgMas8 
      Height          =   300
      Left            =   3960
      Top             =   3390
      Width           =   300
   End
   Begin VB.Image imgMenos8 
      Height          =   300
      Left            =   3120
      Top             =   3390
      Width           =   300
   End
   Begin VB.Image imgMas9 
      Height          =   300
      Left            =   3960
      Top             =   3765
      Width           =   300
   End
   Begin VB.Image imgMenos9 
      Height          =   300
      Left            =   3120
      Top             =   3765
      Width           =   300
   End
   Begin VB.Image imgMas10 
      Height          =   300
      Left            =   8100
      Top             =   780
      Width           =   300
   End
   Begin VB.Image imgMenos10 
      Height          =   300
      Left            =   7260
      Top             =   780
      Width           =   300
   End
   Begin VB.Image imgMas11 
      Height          =   300
      Left            =   8100
      Top             =   1155
      Width           =   300
   End
   Begin VB.Image imgMenos11 
      Height          =   300
      Left            =   7260
      Top             =   1155
      Width           =   300
   End
   Begin VB.Image imgMas12 
      Height          =   300
      Left            =   8100
      Top             =   1515
      Width           =   300
   End
   Begin VB.Image imgMenos12 
      Height          =   300
      Left            =   7260
      Top             =   1515
      Width           =   300
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   12
      Left            =   7635
      TabIndex        =   6
      Top             =   1575
      Width           =   405
   End
   Begin VB.Image imgMas13 
      Height          =   300
      Left            =   8100
      Top             =   1890
      Width           =   300
   End
   Begin VB.Image imgMenos13 
      Height          =   300
      Left            =   7260
      Top             =   1890
      Width           =   300
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   13
      Left            =   7635
      TabIndex        =   5
      Top             =   1950
      Width           =   405
   End
   Begin VB.Image imgMas14 
      Height          =   300
      Left            =   8100
      Top             =   2265
      Width           =   300
   End
   Begin VB.Image imgMenos14 
      Height          =   300
      Left            =   7260
      Top             =   2265
      Width           =   300
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   14
      Left            =   7635
      TabIndex        =   4
      Top             =   2325
      Width           =   405
   End
   Begin VB.Image imgMas15 
      Height          =   300
      Left            =   8100
      Top             =   2640
      Width           =   300
   End
   Begin VB.Image imgMenos15 
      Height          =   300
      Left            =   7260
      Top             =   2640
      Width           =   300
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   15
      Left            =   7635
      TabIndex        =   3
      Top             =   2700
      Width           =   405
   End
   Begin VB.Image imgMas16 
      Height          =   300
      Left            =   8100
      Top             =   3015
      Width           =   300
   End
   Begin VB.Image imgMenos16 
      Height          =   300
      Left            =   7260
      Top             =   3015
      Width           =   300
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   16
      Left            =   7635
      TabIndex        =   2
      Top             =   3075
      Width           =   405
   End
   Begin VB.Image imgMenos1 
      Height          =   300
      Left            =   3120
      Top             =   780
      Width           =   300
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   17
      Left            =   7635
      TabIndex        =   1
      Top             =   3450
      Width           =   405
   End
   Begin VB.Image imgMas17 
      Height          =   300
      Left            =   8100
      Top             =   3390
      Width           =   300
   End
   Begin VB.Image imgMenos17 
      Height          =   300
      Left            =   7260
      Top             =   3390
      Width           =   300
   End
   Begin VB.Image imgAceptar 
      Height          =   375
      Left            =   3120
      Top             =   6105
      Width           =   2775
   End
   Begin VB.Label puntos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1000"
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
      Height          =   300
      Left            =   4650
      TabIndex        =   0
      Top             =   195
      Width           =   615
   End
End
Attribute VB_Name = "frmSkills3"
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

Private cBotonMas(1 To NUMSKILLS) As clsGraphicalButton
Private cBotonMenos(1 To NUMSKILLS) As clsGraphicalButton
Private cSkillNames(1 To NUMSKILLS) As clsGraphicalButton
Private cBtonAceptar As clsGraphicalButton
Private cBotonCancelar As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton

Private bPuedeMagia As Boolean
Private bPuedeMeditar As Boolean
Private bPuedeEscudo As Boolean
Private bPuedeCombateDistancia As Boolean

Private vsHelp(1 To NUMSKILLS) As String

Private Sub Form_Load()
    
On Error GoTo ErrHandler
  
    MirandoAsignarSkills = True
    
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Dim lSkill As Long
    For lSkill = 1 To NUMSKILLS
        text1(lSkill).Caption = UserSkills(lSkill).Both
    Next lSkill
    
    Alocados = SkillPoints
    puntos.Caption = SkillPoints

    'Flags para saber que skills se modificaron
    ReDim Flags(1 To NUMSKILLS)
    
    Call ValidarSkills
    
    Me.Picture = LoadPicture(DirInterfaces & SELECTED_UI & "VentanaSkills.jpg")
    Call LoadButtons
    
    Call LoadHelp
    
    Call modCustomCursors.SetFormCursorDefault(Me)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Form_Load de frmSkills3.frm")
End Sub

Private Sub LoadButtons()
On Error GoTo ErrHandler
  
    Dim GrhPath As String
    Dim I As Long
    
    GrhPath = DirInterfaces & SELECTED_UI

    For I = 1 To NUMSKILLS
        Set cBotonMas(I) = New clsGraphicalButton
        Set cBotonMenos(I) = New clsGraphicalButton
        Set cSkillNames(I) = New clsGraphicalButton
    Next I
    
    Set cBtonAceptar = New clsGraphicalButton
    Set cBotonCancelar = New clsGraphicalButton
    
    Set LastButtonPressed = New clsGraphicalButton
    
    Call cBtonAceptar.Initialize(imgAceptar, GrhPath & "BotonSkillsAceptar.jpg", _
                                    GrhPath & "BotonSkillsAceptarRollover.jpg", _
                                    GrhPath & "BotonSkillsAceptarClick.jpg", Me)

    Call cBotonCancelar.Initialize(imgCancelar, GrhPath & "BotonSkillsSalir.jpg", _
                                    GrhPath & "BotonSkillsSalirRollover.jpg", _
                                    GrhPath & "BotonSkillsSalirClick.jpg", Me)

    Call cBotonMas(1).Initialize(imgMas1, GrhPath & "BotonSkillsMas.jpg", _
                                    GrhPath & "BotonSkillsMasRollover.jpg", _
                                    GrhPath & "BotonSkillsMasClick.jpg", Me, _
                                    GrhPath & "BotonSkillsMasDisabled.jpg", Not bPuedeMagia)

    Call cBotonMas(2).Initialize(imgMas2, GrhPath & "BotonSkillsMas.jpg", _
                                    GrhPath & "BotonSkillsMasRollover.jpg", _
                                    GrhPath & "BotonSkillsMasClick.jpg", Me)

    Call cBotonMas(3).Initialize(imgMas3, GrhPath & "BotonSkillsMas.jpg", _
                                    GrhPath & "BotonSkillsMasRollover.jpg", _
                                    GrhPath & "BotonSkillsMasClick.jpg", Me)

    Call cBotonMas(4).Initialize(imgMas4, GrhPath & "BotonSkillsMas.jpg", _
                                    GrhPath & "BotonSkillsMasRollover.jpg", _
                                    GrhPath & "BotonSkillsMasClick.jpg", Me)
    
    Call cBotonMas(5).Initialize(imgMas5, GrhPath & "BotonSkillsMas.jpg", _
                                    GrhPath & "BotonSkillsMasRollover.jpg", _
                                    GrhPath & "BotonSkillsMasClick.jpg", Me, _
                                    GrhPath & "BotonSkillsMasDisabled.jpg", Not bPuedeMeditar)

    Call cBotonMas(6).Initialize(imgMas6, GrhPath & "BotonSkillsMas.jpg", _
                                    GrhPath & "BotonSkillsMasRollover.jpg", _
                                    GrhPath & "BotonSkillsMasClick.jpg", Me)

    Call cBotonMas(7).Initialize(imgMas7, GrhPath & "BotonSkillsMas.jpg", _
                                    GrhPath & "BotonSkillsMasRollover.jpg", _
                                    GrhPath & "BotonSkillsMasClick.jpg", Me)

    Call cBotonMas(8).Initialize(imgMas8, GrhPath & "BotonSkillsMas.jpg", _
                                    GrhPath & "BotonSkillsMasRollover.jpg", _
                                    GrhPath & "BotonSkillsMasClick.jpg", Me)
    
    Call cBotonMas(9).Initialize(imgMas9, GrhPath & "BotonSkillsMas.jpg", _
                                    GrhPath & "BotonSkillsMasRollover.jpg", _
                                    GrhPath & "BotonSkillsMasClick.jpg", Me)
                                    
    Call cBotonMas(10).Initialize(imgMas10, GrhPath & "BotonSkillsMas.jpg", _
                                    GrhPath & "BotonSkillsMasRollover.jpg", _
                                    GrhPath & "BotonSkillsMasClick.jpg", Me, _
                                    GrhPath & "BotonSkillsMasDisabled.jpg", Not bPuedeEscudo)

    Call cBotonMas(11).Initialize(imgMas11, GrhPath & "BotonSkillsMas.jpg", _
                                    GrhPath & "BotonSkillsMasRollover.jpg", _
                                    GrhPath & "BotonSkillsMasClick.jpg", Me)
    
    Call cBotonMas(12).Initialize(imgMas12, GrhPath & "BotonSkillsMas.jpg", _
                                    GrhPath & "BotonSkillsMasRollover.jpg", _
                                    GrhPath & "BotonSkillsMasClick.jpg", Me)

    Call cBotonMas(13).Initialize(imgMas13, GrhPath & "BotonSkillsMas.jpg", _
                                    GrhPath & "BotonSkillsMasRollover.jpg", _
                                    GrhPath & "BotonSkillsMasClick.jpg", Me)

    Call cBotonMas(14).Initialize(imgMas14, GrhPath & "BotonSkillsMas.jpg", _
                                    GrhPath & "BotonSkillsMasRollover.jpg", _
                                    GrhPath & "BotonSkillsMasClick.jpg", Me)

    Call cBotonMas(15).Initialize(imgMas15, GrhPath & "BotonSkillsMas.jpg", _
                                    GrhPath & "BotonSkillsMasRollover.jpg", _
                                    GrhPath & "BotonSkillsMasClick.jpg", Me)
    
    Call cBotonMas(16).Initialize(imgMas16, GrhPath & "BotonSkillsMas.jpg", _
                                    GrhPath & "BotonSkillsMasRollover.jpg", _
                                    GrhPath & "BotonSkillsMasClick.jpg", Me, _
                                    GrhPath & "BotonSkillsMasDisabled.jpg", Not bPuedeCombateDistancia)
    
    Call cBotonMas(17).Initialize(imgMas17, GrhPath & "BotonSkillsMas.jpg", _
                                  GrhPath & "BotonSkillsMasRollover.jpg", _
                                  GrhPath & "BotonSkillsMasClick.jpg", Me)
    
    Call cBotonMenos(1).Initialize(imgMenos1, GrhPath & "BotonSkillsMenos.jpg", _
                                    GrhPath & "BotonSkillsMenosRollover.jpg", _
                                    GrhPath & "BotonSkillsMenosClick.jpg", Me, _
                                    GrhPath & "BotonSkillsMenosDisabled.jpg", Not bPuedeMagia)

    Call cBotonMenos(2).Initialize(imgMenos2, GrhPath & "BotonSkillsMenos.jpg", _
                                    GrhPath & "BotonSkillsMenosRollover.jpg", _
                                    GrhPath & "BotonSkillsMenosClick.jpg", Me)

    Call cBotonMenos(3).Initialize(imgMenos3, GrhPath & "BotonSkillsMenos.jpg", _
                                    GrhPath & "BotonSkillsMenosRollover.jpg", _
                                    GrhPath & "BotonSkillsMenosClick.jpg", Me)

    Call cBotonMenos(4).Initialize(imgMenos4, GrhPath & "BotonSkillsMenos.jpg", _
                                    GrhPath & "BotonSkillsMenosRollover.jpg", _
                                    GrhPath & "BotonSkillsMenosClick.jpg", Me)
    
    Call cBotonMenos(5).Initialize(imgMenos5, GrhPath & "BotonSkillsMenos.jpg", _
                                    GrhPath & "BotonSkillsMenosRollover.jpg", _
                                    GrhPath & "BotonSkillsMenosClick.jpg", Me, _
                                    GrhPath & "BotonSkillsMenosDisabled.jpg", Not bPuedeMeditar)

    Call cBotonMenos(6).Initialize(imgMenos6, GrhPath & "BotonSkillsMenos.jpg", _
                                    GrhPath & "BotonSkillsMenosRollover.jpg", _
                                    GrhPath & "BotonSkillsMenosClick.jpg", Me)

    Call cBotonMenos(7).Initialize(imgMenos7, GrhPath & "BotonSkillsMenos.jpg", _
                                    GrhPath & "BotonSkillsMenosRollover.jpg", _
                                    GrhPath & "BotonSkillsMenosClick.jpg", Me)

    Call cBotonMenos(8).Initialize(imgMenos8, GrhPath & "BotonSkillsMenos.jpg", _
                                    GrhPath & "BotonSkillsMenosRollover.jpg", _
                                    GrhPath & "BotonSkillsMenosClick.jpg", Me)
    
    Call cBotonMenos(9).Initialize(imgMenos9, GrhPath & "BotonSkillsMenos.jpg", _
                                    GrhPath & "BotonSkillsMenosRollover.jpg", _
                                    GrhPath & "BotonSkillsMenosClick.jpg", Me)

    Call cBotonMenos(10).Initialize(imgMenos10, GrhPath & "BotonSkillsMenos.jpg", _
                                    GrhPath & "BotonSkillsMenosRollover.jpg", _
                                    GrhPath & "BotonSkillsMenosClick.jpg", Me, _
                                    GrhPath & "BotonSkillsMenosDisabled.jpg", Not bPuedeEscudo)

    Call cBotonMenos(11).Initialize(imgMenos11, GrhPath & "BotonSkillsMenos.jpg", _
                                    GrhPath & "BotonSkillsMenosRollover.jpg", _
                                    GrhPath & "BotonSkillsMenosClick.jpg", Me)
    
    Call cBotonMenos(12).Initialize(imgMenos12, GrhPath & "BotonSkillsMenos.jpg", _
                                    GrhPath & "BotonSkillsMenosRollover.jpg", _
                                    GrhPath & "BotonSkillsMenosClick.jpg", Me)

    Call cBotonMenos(13).Initialize(imgMenos13, GrhPath & "BotonSkillsMenos.jpg", _
                                    GrhPath & "BotonSkillsMenosRollover.jpg", _
                                    GrhPath & "BotonSkillsMenosClick.jpg", Me)

    Call cBotonMenos(14).Initialize(imgMenos14, GrhPath & "BotonSkillsMenos.jpg", _
                                    GrhPath & "BotonSkillsMenosRollover.jpg", _
                                    GrhPath & "BotonSkillsMenosClick.jpg", Me)

    Call cBotonMenos(15).Initialize(imgMenos15, GrhPath & "BotonSkillsMenos.jpg", _
                                    GrhPath & "BotonSkillsMenosRollover.jpg", _
                                    GrhPath & "BotonSkillsMenosClick.jpg", Me)
    
    Call cBotonMenos(16).Initialize(imgMenos16, GrhPath & "BotonSkillsMenos.jpg", _
                                    GrhPath & "BotonSkillsMenosRollover.jpg", _
                                    GrhPath & "BotonSkillsMenosClick.jpg", Me, _
                                    GrhPath & "BotonSkillsMenosDisabled.jpg", Not bPuedeCombateDistancia)

    Call cBotonMenos(17).Initialize(imgMenos17, GrhPath & "BotonSkillsMenos.jpg", _
                                    GrhPath & "BotonSkillsMenosRollover.jpg", _
                                    GrhPath & "BotonSkillsMenosClick.jpg", Me)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadButtons de frmSkills3.frm")
End Sub

Private Sub SumarSkillPoint(ByVal SkillIndex As Integer)
On Error GoTo ErrHandler
  
    If Alocados > 0 Then

        If Val(text1(SkillIndex).Caption) < MAXSKILLPOINTS Then
            text1(SkillIndex).Caption = Val(text1(SkillIndex).Caption) + 1
            Flags(SkillIndex) = Flags(SkillIndex) + 1
            Alocados = Alocados - 1
        End If
            
    End If
    
    puntos.Caption = Alocados
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SumarSkillPoint de frmSkills3.frm")
End Sub

Private Sub RestarSkillPoint(ByVal SkillIndex As Integer)
On Error GoTo ErrHandler
  
    If Alocados < SkillPoints Then
        
        If Val(text1(SkillIndex).Caption) > 0 And Flags(SkillIndex) > 0 Then
            text1(SkillIndex).Caption = Val(text1(SkillIndex).Caption) - 1
            Flags(SkillIndex) = Flags(SkillIndex) - 1
            Alocados = Alocados + 1
        End If
    End If
    
    puntos.Caption = Alocados
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub RestarSkillPoint de frmSkills3.frm")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrHandler
  
    LastButtonPressed.ToggleToNormal
    lblHelp.Caption = ""
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Form_MouseMove de frmSkills3.frm")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MirandoAsignarSkills = False
End Sub

Private Sub imgAceptar_Click()
On Error GoTo ErrHandler
  
    Dim skillChanges(NUMSKILLS) As Byte
    Dim I As Long
    
    For I = 1 To NUMSKILLS
        If CByte(text1(I).Caption) > UserSkills(I).Both Then
            If MsgBox("La asignación de puntos es irreversible, ¿desea continuar?", vbYesNo) = vbYes Then
                Exit For
            Else
                Exit Sub
            End If
        End If
    Next I

    For I = 1 To NUMSKILLS
        skillChanges(I) = CByte(text1(I).Caption) - UserSkills(I).Both
        
        'Actualizamos nuestros datos locales
        UserSkills(I).Assigned = UserSkills(I).Assigned + skillChanges(I)
    Next I
    
    Call WriteModifySkills(skillChanges())
    
    If Alocados = 0 Then Call frmMain.LightSkillStar(False)
    
    SkillPoints = Alocados
    
    CerrarVentana
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgAceptar_Click de frmSkills3.frm")
End Sub

Private Sub imgApunialar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.Apuñalar)
End Sub

Private Sub imgCancelar_Click()
    CerrarVentana
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
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CerrarVentana de frmSkills3.frm")
End Sub

Private Sub imgCarpinteria_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.Carpinteria)
End Sub

Private Sub imgCombateArmas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.Armas)
End Sub

Private Sub imgCombateDistancia_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.Proyectiles)
End Sub

Private Sub imgCombateSinArmas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.Wrestling)
End Sub

Private Sub imgDomar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.Domar)
End Sub

Private Sub imgEscudos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.Defensa)
End Sub

Private Sub imgEvasion_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.Tacticas)
End Sub

Private Sub imgHerreria_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.Herreria)
End Sub

Private Sub imgMagia_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.Magia)
End Sub

Private Sub imgMas1_Click()
    Call SumarSkillPoint(1)
End Sub

Private Sub imgMas10_Click()
    Call SumarSkillPoint(10)
End Sub

Private Sub imgMas11_Click()
    Call SumarSkillPoint(11)
End Sub

Private Sub imgMas12_Click()
    Call SumarSkillPoint(12)
End Sub

Private Sub imgMas13_Click()
    Call SumarSkillPoint(13)
End Sub

Private Sub imgMas14_Click()
    Call SumarSkillPoint(14)
End Sub

Private Sub imgMas15_Click()
    Call SumarSkillPoint(15)
End Sub

Private Sub imgMas16_Click()
    Call SumarSkillPoint(16)
End Sub

Private Sub imgMas17_Click()
    Call SumarSkillPoint(17)
End Sub

Private Sub imgMas18_Click()
    Call SumarSkillPoint(18)
End Sub

Private Sub imgMas2_Click()
    Call SumarSkillPoint(2)
End Sub

Private Sub imgMas3_Click()
    Call SumarSkillPoint(3)
End Sub

Private Sub imgMas4_Click()
    Call SumarSkillPoint(4)
End Sub

Private Sub imgMas5_Click()
    Call SumarSkillPoint(5)
End Sub

Private Sub imgMas6_Click()
    Call SumarSkillPoint(6)
End Sub

Private Sub imgMas7_Click()
    Call SumarSkillPoint(7)
End Sub

Private Sub imgMas8_Click()
    Call SumarSkillPoint(8)
End Sub

Private Sub imgMas9_Click()
    Call SumarSkillPoint(9)
End Sub

Private Sub imgMeditar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.Meditar)
End Sub

Private Sub imgMenos1_Click()
    Call RestarSkillPoint(1)
End Sub

Private Sub imgMenos10_Click()
    Call RestarSkillPoint(10)
End Sub

Private Sub imgMenos11_Click()
    Call RestarSkillPoint(11)
End Sub

Private Sub imgMenos12_Click()
    Call RestarSkillPoint(12)
End Sub

Private Sub imgMenos13_Click()
    Call RestarSkillPoint(13)
End Sub

Private Sub imgMenos14_Click()
    Call RestarSkillPoint(14)
End Sub

Private Sub imgMenos15_Click()
    Call RestarSkillPoint(15)
End Sub

Private Sub imgMenos16_Click()
    Call RestarSkillPoint(16)
End Sub

Private Sub imgMenos17_Click()
    Call RestarSkillPoint(17)
End Sub

Private Sub imgMenos18_Click()
    Call RestarSkillPoint(18)
End Sub

Private Sub imgMenos2_Click()
    Call RestarSkillPoint(2)
End Sub

Private Sub imgMenos3_Click()
    Call RestarSkillPoint(3)
End Sub

Private Sub imgMenos4_Click()
    Call RestarSkillPoint(4)
End Sub

Private Sub imgMenos5_Click()
    Call RestarSkillPoint(5)
End Sub

Private Sub imgMenos6_Click()
    Call RestarSkillPoint(6)
End Sub

Private Sub imgMenos7_Click()
    Call RestarSkillPoint(7)
End Sub

Private Sub imgMenos8_Click()
    Call RestarSkillPoint(8)
End Sub

Private Sub imgMenos9_Click()
    Call RestarSkillPoint(9)
End Sub

Private Sub LoadHelp()
On Error GoTo ErrHandler
  
    
    vsHelp(eSkill.Magia) = "Magia:" & vbCrLf & _
                            "- Representa la habilidad de un personaje de las áreas mágica." & vbCrLf & _
                            "- Indica la variedad de hechizos que es capaz de dominar el personaje."
    If Not bPuedeMagia Then
        vsHelp(eSkill.Magia) = vsHelp(eSkill.Magia) & vbCrLf & _
                                "* Habilidad inhabilitada para tu clase."
    End If
    
    vsHelp(eSkill.Robar) = "Robar:" & vbCrLf & _
                            "- Habilidades de hurto. Nunca por medio de la violencia." & vbCrLf & _
                            "- Indica la probabilidad de éxito del personaje al intentar apoderarse de oro de otro, en caso de ser Ladrón, tambien podrá apoderarse de items."
    
    vsHelp(eSkill.Tacticas) = "Evasión en Combate:" & vbCrLf & _
                                "- Representa la habilidad general para moverse en combate entre golpes enemigos sin morir o tropezar en el intento." & vbCrLf & _
                                "- Indica la posibilidad de evadir un golpe físico del personaje."
    
    vsHelp(eSkill.Armas) = "Combate con Armas:" & vbCrLf & _
                            "- Representa la habilidad del personaje para manejar armas de combate cuerpo a cuerpo." & vbCrLf & _
                            "- Indica la probabilidad de impactar al oponente con armas cuerpo a cuerpo."
    
    vsHelp(eSkill.Meditar) = "Meditar:" & vbCrLf & _
                                "- Representa la capacidad del personaje de concentrarse para abstrarse dentro de su mente, y así revitalizar su fuerza espiritual." & vbCrLf & _
                                "- Indica la velocidad a la que el personaje recupera maná (Clases mágicas)."
    
    If Not bPuedeMeditar Then
        vsHelp(eSkill.Meditar) = vsHelp(eSkill.Meditar) & vbCrLf & _
                                "* Habilidad inhabilitada para tu clase."
    End If

    vsHelp(eSkill.Apuñalar) = "Apuñalar:" & vbCrLf & _
                                "- Representa la destreza para inflingir daño grave con armas cortas." & vbCrLf & _
                                "- Indica la posibilidad de apuñalar al enemigo en un ataque. El Asesino es la única clase que no necesitará 10 skills para comenzar a entrenar esta habilidad."

    vsHelp(eSkill.Ocultarse) = "Ocultarse:" & vbCrLf & _
                                "- La habilidad propia de un personaje para mimetizarse con el medio y evitar se perciba su presencia." & vbCrLf & _
                                "- Indica la facilidad con la que uno puede desaparecer de la vista de los demás y por cuanto tiempo."
    
    vsHelp(eSkill.Supervivencia) = "Superivencia:" & vbCrLf & _
                                    "- Es el conjunto de habilidades necesarias para sobrevivir fuera de una ciudad en base a lo que la naturaleza ofrece." & vbCrLf & _
                                    "- Permite conocer la salud de las criaturas guiándose exclusivamente por su aspecto, así como encender fogatas junto a las que descansar."
    
    vsHelp(eSkill.Talar) = "Talar:" & vbCrLf & _
                            "- Es la habilidad en el uso del hacha para evitar desperdiciar leña y maximizar la efectividad de cada golpe dado." & vbCrLf & _
                            "- Indica la probabilidad de obtener leña por golpe."
    
    vsHelp(eSkill.Defensa) = "Defensa con Escudos:" & vbCrLf & _
                                "- Es la habilidad de interponer correctamente el escudo ante cada embate enemigo para evitar ser impactado sin perder el equilibrio y poder responder rápidamente con la otra mano." & vbCrLf & _
                                "- Indica las probabilidades de bloquear un impacto con el escudo."
    
    If Not bPuedeEscudo Then
        vsHelp(eSkill.Defensa) = vsHelp(eSkill.Defensa) & vbCrLf & _
                                "* Habilidad inhabilitada para tu clase."
    End If


    vsHelp(eSkill.Pesca) = "Pesca:" & vbCrLf & _
                            "- Es el conjunto de conocimientos básicos para poder armar un señuelo, poner la carnada en el anzuelo y saber dónde buscar peces." & vbCrLf & _
                            "- Indica la probabilidad de tener éxito en cada intento de pescar."
    
    vsHelp(eSkill.Mineria) = "Minería:" & vbCrLf & _
                                "- Es el conjunto de conocimientos sobre los distintos minerales, el dónde se obtienen, cómo deben ser extraídos y trabajados." & vbCrLf & _
                                "- Indica la probabilidad de tener éxito en cada intento de minar y la capacidad, o no de convertir estos minerales en lingotes."
    
    vsHelp(eSkill.Carpinteria) = "Carpintería:" & vbCrLf & _
                                    "- Es el conjunto de conocimientos para saber serruchar, lijar, encolar y clavar madera con un buen nivel de terminación." & vbCrLf & _
                                    "- Indica la habilidad en el manejo de estas herramientas, el que tan bueno se es en el oficio de carpintero."
    
    vsHelp(eSkill.Herreria) = "Herrería:" & vbCrLf & _
                                "- Es el conjunto de conocimientos para saber procesar cada tipo de mineral para fundirlo, forjarlo y crear aleaciones." & vbCrLf & _
                                "- Indica la habilidad en el manejo de estas técnicas, el que tan bueno se es en el oficio de herrero."
    
    vsHelp(eSkill.Domar) = "Domar Animales:" & vbCrLf & _
                                "- Es la habilidad en el trato con animales para que estos te sigan y ayuden en combate." & vbCrLf & _
                                "- Indica la posibilidad de lograr domar a una criatura y qué clases de criaturas se puede domar."
    
    vsHelp(eSkill.Proyectiles) = "Combate a distancia:" & vbCrLf & _
                                "- Es el manejo de las armas de largo alcance." & vbCrLf & _
                                "- Indica la probabilidad de éxito para impactar a un enemigo con este tipo de armas."
    
    If Not bPuedeCombateDistancia Then
        vsHelp(eSkill.Proyectiles) = vsHelp(eSkill.Proyectiles) & vbCrLf & _
                                "* Habilidad inhabilitada para tu clase."
    End If

    vsHelp(eSkill.Wrestling) = "Combate sin armas:" & vbCrLf & _
                                "- Es la habilidad del personaje para entrar en combate sin arma alguna salvo sus propios brazos." & vbCrLf & _
                                "- Indica la probabilidad de éxito para impactar a un enemigo estando desarmado. El Bandido y Ladrón tienen habilidades extras asociadas a esta habilidad."
    
    vsHelp(eSkill.Sastreria) = "Sastrería:" & vbCrLf & _
                            "- Es el conjunto de conocimientos para saber convertir cualquier tipo de piel en ropajes." & vbCrLf & _
                            "- Indica la habilidad en el manejo de estas técnicas, el que tan bueno se es en el oficio de sastre."

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadHelp de frmSkills3.frm")
End Sub

Private Sub imgMineria_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.Mineria)
End Sub

Private Sub imgOcultarse_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.Ocultarse)
End Sub

Private Sub imgPesca_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.Pesca)
End Sub

Private Sub imgRobar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.Robar)
End Sub

Private Sub imgSastreria_Click()
    Call ShowHelp(eSkill.Sastreria)
End Sub

Private Sub imgSupervivencia_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.Supervivencia)
End Sub

Private Sub imgTalar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.Talar)
End Sub

Private Sub lblHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub ShowHelp(ByVal eeSkill As eSkill)
On Error GoTo ErrHandler
  
    lblHelp.Caption = vsHelp(eeSkill)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ShowHelp de frmSkills3.frm")
End Sub

Private Sub ValidarSkills()
On Error GoTo ErrHandler
  

    bPuedeMagia = True
    bPuedeMeditar = True
    bPuedeEscudo = True
    bPuedeCombateDistancia = True

    Select Case PlayerData.Class
        Case eClass.Warrior, eClass.Worker, eClass.Thief, eClass.Hunter
            bPuedeMagia = False
            bPuedeMeditar = False
        
        Case eClass.Mage, eClass.Druid
            bPuedeEscudo = False
            bPuedeCombateDistancia = False
            
    End Select
    
    ' Magia
    imgMas1.Enabled = bPuedeMagia
    imgMenos1.Enabled = bPuedeMagia

    ' Meditar
    imgMas5.Enabled = bPuedeMeditar
    imgMenos5.Enabled = bPuedeMeditar

    ' Escudos
    imgMas10.Enabled = bPuedeEscudo
    imgMenos10.Enabled = bPuedeEscudo

    ' Proyectiles
    imgMas16.Enabled = bPuedeCombateDistancia
    imgMenos16.Enabled = bPuedeCombateDistancia
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ValidarSkills de frmSkills3.frm")
End Sub

