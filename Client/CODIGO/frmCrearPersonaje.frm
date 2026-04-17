VERSION 5.00
Begin VB.Form frmCrearPersonaje 
   BackColor       =   &H8000000A&
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11985
   ClipControls    =   0   'False
   Icon            =   "frmCrearPersonaje.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   799
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tAnimacion 
      Interval        =   1000
      Left            =   600
      Top             =   360
   End
   Begin VB.ComboBox lstProfesion 
      BackColor       =   &H00000000&
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
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":030A
      Left            =   4680
      List            =   "frmCrearPersonaje.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   4080
      Width           =   2625
   End
   Begin VB.ComboBox lstGenero 
      BackColor       =   &H00000000&
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
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":030E
      Left            =   4680
      List            =   "frmCrearPersonaje.frx":0318
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   4635
      Width           =   2625
   End
   Begin VB.ComboBox lstRaza 
      BackColor       =   &H00000000&
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
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":032B
      Left            =   4680
      List            =   "frmCrearPersonaje.frx":032D
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   3525
      Width           =   2625
   End
   Begin VB.TextBox txtNombre 
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
      Height          =   300
      Left            =   3900
      MaxLength       =   15
      TabIndex        =   0
      Top             =   2880
      Width           =   4095
   End
   Begin VB.PictureBox picPJ 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   1440
      ScaleHeight     =   65
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   4
      Top             =   4185
      Width           =   615
   End
   Begin VB.PictureBox picHead 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   2
      Left            =   2115
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   12
      Top             =   3315
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picHead 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   1
      Left            =   1560
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   13
      Top             =   3315
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picHead 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   0
      Left            =   1005
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   11
      Top             =   3315
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgWebsiteLink 
      Height          =   615
      Left            =   3960
      Top             =   8160
      Width           =   4455
   End
   Begin VB.Image imgDiscord 
      Height          =   525
      Left            =   11175
      Top             =   8235
      Width           =   600
   End
   Begin VB.Image imgArcoStar 
      Height          =   195
      Index           =   5
      Left            =   11040
      Top             =   4575
      Width           =   225
   End
   Begin VB.Image imgArcoStar 
      Height          =   195
      Index           =   4
      Left            =   10815
      Top             =   4575
      Width           =   225
   End
   Begin VB.Image imgArcoStar 
      Height          =   195
      Index           =   3
      Left            =   10590
      Top             =   4575
      Width           =   225
   End
   Begin VB.Image imgArcoStar 
      Height          =   195
      Index           =   2
      Left            =   10365
      Top             =   4575
      Width           =   225
   End
   Begin VB.Image imgArcoStar 
      Height          =   195
      Index           =   1
      Left            =   10140
      Top             =   4575
      Width           =   225
   End
   Begin VB.Image imgArmasStar 
      Height          =   195
      Index           =   5
      Left            =   11040
      Top             =   4245
      Width           =   225
   End
   Begin VB.Image imgArmasStar 
      Height          =   195
      Index           =   4
      Left            =   10815
      Top             =   4245
      Width           =   225
   End
   Begin VB.Image imgArmasStar 
      Height          =   195
      Index           =   3
      Left            =   10590
      Top             =   4245
      Width           =   225
   End
   Begin VB.Image imgArmasStar 
      Height          =   195
      Index           =   2
      Left            =   10365
      Top             =   4245
      Width           =   225
   End
   Begin VB.Image imgEscudosStar 
      Height          =   195
      Index           =   5
      Left            =   11040
      Top             =   3915
      Width           =   225
   End
   Begin VB.Image imgEscudosStar 
      Height          =   195
      Index           =   4
      Left            =   10815
      Top             =   3915
      Width           =   225
   End
   Begin VB.Image imgEscudosStar 
      Height          =   195
      Index           =   3
      Left            =   10590
      Top             =   3915
      Width           =   225
   End
   Begin VB.Image imgEscudosStar 
      Height          =   195
      Index           =   2
      Left            =   10365
      Top             =   3915
      Width           =   225
   End
   Begin VB.Image imgVidaStar 
      Height          =   195
      Index           =   5
      Left            =   11040
      Top             =   3585
      Width           =   225
   End
   Begin VB.Image imgVidaStar 
      Height          =   195
      Index           =   4
      Left            =   10815
      Top             =   3585
      Width           =   225
   End
   Begin VB.Image imgVidaStar 
      Height          =   195
      Index           =   3
      Left            =   10590
      Top             =   3585
      Width           =   225
   End
   Begin VB.Image imgVidaStar 
      Height          =   195
      Index           =   2
      Left            =   10365
      Top             =   3585
      Width           =   225
   End
   Begin VB.Image imgMagiaStar 
      Height          =   195
      Index           =   5
      Left            =   11040
      Top             =   3255
      Width           =   225
   End
   Begin VB.Image imgMagiaStar 
      Height          =   195
      Index           =   4
      Left            =   10815
      Top             =   3255
      Width           =   225
   End
   Begin VB.Image imgMagiaStar 
      Height          =   195
      Index           =   3
      Left            =   10590
      Top             =   3255
      Width           =   225
   End
   Begin VB.Image imgMagiaStar 
      Height          =   195
      Index           =   2
      Left            =   10365
      Top             =   3255
      Width           =   225
   End
   Begin VB.Image imgArmasStar 
      Height          =   195
      Index           =   1
      Left            =   10140
      Top             =   4245
      Width           =   225
   End
   Begin VB.Image imgEscudosStar 
      Height          =   195
      Index           =   1
      Left            =   10140
      Top             =   3915
      Width           =   225
   End
   Begin VB.Image imgVidaStar 
      Height          =   195
      Index           =   1
      Left            =   10140
      Top             =   3585
      Width           =   225
   End
   Begin VB.Image imgMagiaStar 
      Height          =   195
      Index           =   1
      Left            =   10140
      Top             =   3255
      Width           =   225
   End
   Begin VB.Image imgEvasionStar 
      Height          =   195
      Index           =   5
      Left            =   11070
      Top             =   2925
      Width           =   225
   End
   Begin VB.Image imgEvasionStar 
      Height          =   195
      Index           =   4
      Left            =   10845
      Top             =   2925
      Width           =   225
   End
   Begin VB.Image imgEvasionStar 
      Height          =   195
      Index           =   3
      Left            =   10620
      Top             =   2925
      Width           =   225
   End
   Begin VB.Image imgEvasionStar 
      Height          =   195
      Index           =   2
      Left            =   10395
      Top             =   2925
      Width           =   225
   End
   Begin VB.Image imgEvasionStar 
      Height          =   195
      Index           =   1
      Left            =   10170
      Top             =   2925
      Width           =   225
   End
   Begin VB.Label lblEspecialidad 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Especialidad"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   9180
      TabIndex        =   14
      Top             =   5325
      Width           =   2160
   End
   Begin VB.Label lblAtributoFinal 
      BackStyle       =   0  'Transparent
      Caption         =   "18"
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
      Height          =   195
      Index           =   5
      Left            =   7680
      TabIndex        =   10
      Top             =   5355
      Width           =   225
   End
   Begin VB.Label lblAtributoFinal 
      BackStyle       =   0  'Transparent
      Caption         =   "18"
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
      Height          =   195
      Index           =   4
      Left            =   6795
      TabIndex        =   9
      Top             =   5355
      Width           =   225
   End
   Begin VB.Label lblAtributoFinal 
      BackStyle       =   0  'Transparent
      Caption         =   "18"
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
      Height          =   195
      Index           =   3
      Left            =   5880
      TabIndex        =   8
      Top             =   5355
      Width           =   225
   End
   Begin VB.Label lblAtributoFinal 
      BackStyle       =   0  'Transparent
      Caption         =   "18"
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
      Height          =   195
      Index           =   2
      Left            =   4980
      TabIndex        =   7
      Top             =   5355
      Width           =   225
   End
   Begin VB.Label lblAtributoFinal 
      BackStyle       =   0  'Transparent
      Caption         =   "18"
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
      Height          =   195
      Index           =   1
      Left            =   4080
      TabIndex        =   6
      Top             =   5355
      Width           =   225
   End
   Begin VB.Label lblHelp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1320
      Left            =   3075
      TabIndex        =   5
      Top             =   6300
      Width           =   5880
   End
   Begin VB.Image imgVolver 
      Height          =   525
      Left            =   840
      Top             =   6780
      Width           =   1230
   End
   Begin VB.Image imgCrear 
      Height          =   525
      Left            =   9840
      Top             =   6780
      Width           =   1230
   End
   Begin VB.Image imgGenero 
      Height          =   180
      Left            =   5580
      Top             =   4440
      Width           =   630
   End
   Begin VB.Image imgClase 
      Height          =   180
      Left            =   5640
      Top             =   3870
      Width           =   480
   End
   Begin VB.Image imgRaza 
      Height          =   180
      Left            =   5640
      Top             =   3315
      Width           =   480
   End
   Begin VB.Image imgEspecialidad 
      Height          =   240
      Left            =   9480
      Top             =   5010
      Width           =   1425
   End
   Begin VB.Image imgArcos 
      Height          =   255
      Left            =   9420
      Top             =   4545
      Width           =   495
   End
   Begin VB.Image imgArmas 
      Height          =   255
      Left            =   9390
      Top             =   4215
      Width           =   525
   End
   Begin VB.Image imgEscudos 
      Height          =   255
      Left            =   9270
      Top             =   3885
      Width           =   660
   End
   Begin VB.Image imgVida 
      Height          =   255
      Left            =   9540
      Top             =   3540
      Width           =   390
   End
   Begin VB.Image imgMagia 
      Height          =   255
      Left            =   9420
      Top             =   3240
      Width           =   510
   End
   Begin VB.Image imgEvasion 
      Height          =   255
      Left            =   9300
      Top             =   2910
      Width           =   660
   End
   Begin VB.Image imgConstitucion 
      Height          =   255
      Left            =   7290
      Top             =   5100
      Width           =   990
   End
   Begin VB.Image imgCarisma 
      Height          =   240
      Left            =   6615
      Top             =   5100
      Width           =   630
   End
   Begin VB.Image imgInteligencia 
      Height          =   240
      Left            =   5520
      Top             =   5100
      Width           =   960
   End
   Begin VB.Image imgAgilidad 
      Height          =   240
      Left            =   4740
      Top             =   5100
      Width           =   705
   End
   Begin VB.Image imgFuerza 
      Height          =   240
      Left            =   3930
      Top             =   5100
      Width           =   525
   End
   Begin VB.Image imgNombre 
      Height          =   270
      Left            =   4530
      Top             =   2520
      Width           =   2715
   End
   Begin VB.Image picCharacterRight 
      Height          =   270
      Left            =   2295
      Top             =   4485
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image picCharacterLeft 
      Height          =   270
      Left            =   915
      Top             =   4485
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image picHeadRight 
      Height          =   270
      Left            =   2625
      Top             =   3375
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image picHeadLeft 
      Height          =   270
      Left            =   600
      Top             =   3375
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image Image1 
      Height          =   3120
      Left            =   8880
      Stretch         =   -1  'True
      Top             =   9120
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.Image imgHogar 
      Height          =   480
      Left            =   5640
      Picture         =   "frmCrearPersonaje.frx":032F
      Top             =   9120
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmCrearPersonaje"
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

Private cBotonPasswd As clsGraphicalButton
Private cBotonMail As clsGraphicalButton
Private cBotonNombre As clsGraphicalButton
Private cBotonConfirmPasswd As clsGraphicalButton
Private cBotonAtributos As clsGraphicalButton
Private cBotonD As clsGraphicalButton
Private cBotonM As clsGraphicalButton
Private cBotonF As clsGraphicalButton
Private cBotonFuerza As clsGraphicalButton
Private cBotonAgilidad As clsGraphicalButton
Private cBotonInteligencia As clsGraphicalButton
Private cBotonCarisma As clsGraphicalButton
Private cBotonConstitucion As clsGraphicalButton
Private cBotonEvasion As clsGraphicalButton
Private cBotonMagia As clsGraphicalButton
Private cBotonVida As clsGraphicalButton
Private cBotonEscudos As clsGraphicalButton
Private cBotonArmas As clsGraphicalButton
Private cBotonArcos As clsGraphicalButton
Private cBotonEspecialidad As clsGraphicalButton
Private cBotonPuebloOrigen As clsGraphicalButton
Private cBotonRaza As clsGraphicalButton
Private cBotonClase As clsGraphicalButton
Private cBotonGenero As clsGraphicalButton
Private cBotonAlineacion As clsGraphicalButton
Private cBotonVolver As clsGraphicalButton
Private cBotonCrear As clsGraphicalButton
Private cButtonHeadLeft As clsGraphicalButton
Private cButtonHeadRight As clsGraphicalButton

Private cButtonBodyLeft As clsGraphicalButton
Private cButtonBodyRight As clsGraphicalButton

Private cButtonDiscord As clsGraphicalButton


Public LastButtonPressed As clsGraphicalButton

Private picFullStar As Picture
Private picHalfStar As Picture
Private picGlowStar As Picture

Private Enum eHelp
    iePasswd
    ieMail
    ieNombre
    ieConfirmPasswd
    ieAtributos
    ieFuerza
    ieAgilidad
    ieInteligencia
    ieCarisma
    ieConstitucion
    ieEvasion
    ieMagia
    ieVida
    ieEscudos
    ieArmas
    ieArcos
    ieEspecialidad
    iePuebloOrigen
    ieRaza
    ieClase
    ieGenero
    ieAlineacion
End Enum

Private vHelp(21) As String
Private vEspecialidades() As String

Private Type tModRaza
    Fuerza As Single
    Agilidad As Single
    Inteligencia As Single
    Carisma As Single
    Constitucion As Single
End Type

Private Type tModClase
    Evasion As Double
    AtaqueArmas As Double
    AtaqueProyectiles As Double
    DañoArmas As Double
    DañoProyectiles As Double
    Escudo As Double
    Magia As Double
    Vida As Double
    Hit As Double
End Type

Private ModRaza() As tModRaza
Private ModClase() As tModClase

Private NroRazas As Integer
Private NroClases As Integer

Private Cargando As Boolean

Private currentGrh As Long
Private Dir As E_Heading

Private mDevice(0 To 3) As Long


Private Sub Form_Load()
On Error GoTo ErrHandler
  
    Cargando = True
    Call LoadCharInfo
    Call CargarEspecialidades
    
    Call IniciarGraficos
    Call CargarCombos
    
    Call LoadHelp

    Dir = SOUTH
    
    Cargando = False
    
    PlayerData.Gender = 0
    PlayerData.Race = 0
    UserEmail = ""
    UserHead = 0
    
    txtNombre.MaxLength = MAX_NICKNAME_SIZE
    
    Dim I As Integer
        
    For I = LBound(mDevice) To UBound(mDevice) - 1
        mDevice(I) = Aurora_Graphics.CreatePassFromDisplay(picHead(I).hwnd, picHead(I).ScaleWidth, picHead(I).ScaleHeight)
    Next I
    
     mDevice(UBound(mDevice)) = Aurora_Graphics.CreatePassFromDisplay(picPj.hwnd, picPj.ScaleWidth, picPj.ScaleHeight)

#If EnableSecurity Then
    Call ProtectForm(Me)
#End If

    ' Select the 3 dropdowns default elements
    lstRaza.ListIndex = 0
    lstProfesion.ListIndex = 0
    lstGenero.ListIndex = 0

    
    Call modCustomCursors.SetFormCursorDefault(Me)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Form_Load de frmCrearPersonaje.frm")
End Sub

Private Sub CargarEspecialidades()
On Error GoTo ErrHandler
  

    ReDim vEspecialidades(1 To NroClases)
    
    vEspecialidades(eClass.Thief) = "Robar y Ocultarse"
    vEspecialidades(eClass.Assasin) = "Apuñalar"
    vEspecialidades(eClass.Bandit) = "Combate Sin Armas"
    vEspecialidades(eClass.Druid) = "Domar"
    vEspecialidades(eClass.Worker) = "Extracción y Construcción"
    vEspecialidades(eClass.Bard) = "Hechicería y Evasión"
    vEspecialidades(eClass.Cleric) = "Hechicería y Armas"
    vEspecialidades(eClass.Mage) = "Hechicería"
    vEspecialidades(eClass.Warrior) = "Durabilidad y Armas"
    vEspecialidades(eClass.Paladin) = "Combate Con Armas"
    vEspecialidades(eClass.Hunter) = "Combate Con Arcos"
       
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CargarEspecialidades de frmCrearPersonaje.frm")
End Sub

Private Sub IniciarGraficos()
On Error GoTo ErrHandler
  

    Dim GrhPath As String
    GrhPath = DirInterfaces & SELECTED_UI
    
    Me.Picture = LoadPicture(GrhPath & "VentanaCrearPersonaje.jpg")
    
    Set cBotonPasswd = New clsGraphicalButton
    Set cBotonMail = New clsGraphicalButton
    Set cBotonNombre = New clsGraphicalButton
    Set cBotonConfirmPasswd = New clsGraphicalButton
    Set cBotonAtributos = New clsGraphicalButton
    Set cBotonD = New clsGraphicalButton
    Set cBotonM = New clsGraphicalButton
    Set cBotonF = New clsGraphicalButton
    Set cBotonFuerza = New clsGraphicalButton
    Set cBotonAgilidad = New clsGraphicalButton
    Set cBotonInteligencia = New clsGraphicalButton
    Set cBotonCarisma = New clsGraphicalButton
    Set cBotonConstitucion = New clsGraphicalButton
    Set cBotonEvasion = New clsGraphicalButton
    Set cBotonMagia = New clsGraphicalButton
    Set cBotonVida = New clsGraphicalButton
    Set cBotonEscudos = New clsGraphicalButton
    Set cBotonArmas = New clsGraphicalButton
    Set cBotonArcos = New clsGraphicalButton
    Set cBotonEspecialidad = New clsGraphicalButton
    Set cBotonPuebloOrigen = New clsGraphicalButton
    Set cBotonRaza = New clsGraphicalButton
    Set cBotonClase = New clsGraphicalButton
    Set cBotonGenero = New clsGraphicalButton
    Set cBotonAlineacion = New clsGraphicalButton
    Set cBotonVolver = New clsGraphicalButton
    Set cBotonCrear = New clsGraphicalButton
    
    Set cButtonHeadLeft = New clsGraphicalButton
    Set cButtonHeadRight = New clsGraphicalButton
    Set cButtonBodyLeft = New clsGraphicalButton
    Set cButtonBodyRight = New clsGraphicalButton
    
    Set cButtonDiscord = New clsGraphicalButton
    
    Set LastButtonPressed = New clsGraphicalButton
                              
    Call cBotonVolver.Initialize(ImgVolver, GrhPath & "BotonVolver.jpg", GrhPath & "BotonVolver.jpg", _
                                    GrhPath & "BotonVolver.jpg", Me)
                                    
    Call cBotonCrear.Initialize(ImgCrear, GrhPath & "BotonCrear.jpg", GrhPath & "BotonCrear.jpg", _
                                   GrhPath & "BotonCrear.jpg", Me)
                                   
    Call cButtonDiscord.Initialize(imgDiscord, GrhPath & "BotonDiscord.jpg", GrhPath & "BotonDiscord.jpg", _
                                   GrhPath & "BotonDiscord.jpg", Me)
                                   
                                   
    Call cButtonHeadLeft.Initialize(picHeadLeft, GrhPath & "BotonFlechaIzquierda_2.jpg", GrhPath & "BotonFlechaIzquierda_2.jpg", _
                                   GrhPath & "BotonFlechaIzquierda_2.jpg", Me)
                                   
    Call cButtonHeadRight.Initialize(picHeadRight, GrhPath & "BotonFlechaDerecha_2.jpg", GrhPath & "BotonFlechaDerecha_2.jpg", _
                                   GrhPath & "BotonFlechaDerecha_2.jpg", Me)
                                   
    Call cButtonBodyLeft.Initialize(picCharacterLeft, GrhPath & "BotonFlechaIzquierda_2.jpg", GrhPath & "BotonFlechaIzquierda_2.jpg", _
                                   GrhPath & "BotonFlechaIzquierda_2.jpg", Me)
                                   
    Call cButtonBodyRight.Initialize(picCharacterRight, GrhPath & "BotonFlechaDerecha_2.jpg", GrhPath & "BotonFlechaDerecha_2.jpg", _
                                   GrhPath & "BotonFlechaDerecha_2.jpg", Me)

    Set picFullStar = LoadPicture(GrhPath & "EstrellaSimple.jpg")
    Set picHalfStar = LoadPicture(GrhPath & "EstrellaMitad.jpg")
    Set picGlowStar = LoadPicture(GrhPath & "EstrellaBrillante.jpg")
    
    
    

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub IniciarGraficos de frmCrearPersonaje.frm")
End Sub

Private Sub CargarCombos()
On Error GoTo ErrHandler
  
    Dim I As Integer
    
    lstProfesion.Clear
    
    For I = LBound(ListEnabledClasses) To EnabledClassesQty
        lstProfesion.AddItem ListaClases(ListEnabledClasses(I))
    Next I
    
    lstRaza.Clear
    
    For I = LBound(ListaRazas()) To NroRazas
        lstRaza.AddItem ListaRazas(I)
    Next I
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CargarCombos de frmCrearPersonaje.frm")
End Sub

Function CheckData() As Boolean
On Error GoTo ErrHandler

    If AccountConnecting.UserRace = 0 Then
        Call frmMessageBox.ShowMessage("Seleccione la raza del personaje.")
        Exit Function
    End If
        
    If AccountConnecting.UserClass = 0 Then
        Call frmMessageBox.ShowMessage("Seleccione la clase del personaje.")
        Exit Function
    End If
    
    If AccountConnecting.UserGender = 0 Then
        Call frmMessageBox.ShowMessage("Seleccione el género del personaje.")
        Exit Function
    End If
    
    If Len(UserName) > MAX_NICKNAME_SIZE Then
        Call frmMessageBox.ShowMessage("El nombre debe tener " & MAX_NICKNAME_SIZE & " letras o menos.")
        Exit Function
    End If
    
    CheckData = True

  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function CheckData de frmCrearPersonaje.frm")
End Function



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ClearLabel
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim I As Long
    
    For I = LBound(mDevice) To UBound(mDevice)
        Call Aurora_Graphics.DeletePass(mDevice(I))
    Next I
    
End Sub


Private Sub ImgCrear_Click()

On Error GoTo ErrHandler
  
    Dim I As Integer
    Dim CharAscii As Byte
    
    UserName = complexNameToSimple(txtNombre.text, True)
            
    If Right$(UserName, 1) = " " Then
        UserName = RTrim$(UserName)
    End If
    
    If Len(UserName) > MAX_NICKNAME_SIZE Then
        Call frmMessageBox.ShowMessage("El nombre de tu personaje no puede superar los " & MAX_NICKNAME_SIZE & " caracteres")
        Exit Sub
    End If
    
    If Not CheckData Then Exit Sub
    
    AccountConnecting.UserRace = lstRaza.ListIndex + 1
    AccountConnecting.UserGender = lstGenero.ListIndex + 1
    AccountConnecting.UserClass = ListEnabledClasses(lstProfesion.ListIndex + 1)
    
    bShowTutorial = True

    Call modAccount.Prepare_And_Connect(E_MODO.AccountCreateChar)

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgCrear_Click de frmCrearPersonaje.frm")
End Sub

Private Sub imgEspecialidad_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieEspecialidad)
End Sub

Private Sub imgNombre_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieNombre)
End Sub

Private Sub imgAtributos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieAtributos)
End Sub


Private Sub imgFuerza_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieFuerza)
End Sub

Private Sub imgAgilidad_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieAgilidad)
End Sub

Private Sub imgInteligencia_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieInteligencia)
End Sub

Private Sub imgCarisma_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieCarisma)
End Sub

Private Sub imgConstitucion_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieConstitucion)
End Sub

Private Sub imgArcos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieArcos)
End Sub

Private Sub imgArmas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieArmas)
End Sub

Private Sub imgEscudos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieEscudos)
End Sub

Private Sub imgEvasion_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieEvasion)
End Sub

Private Sub imgMagia_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieMagia)
End Sub

Private Sub imgVida_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieVida)
End Sub

Private Sub imgRaza_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieRaza)
End Sub

Private Sub imgClase_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieClase)
End Sub

Private Sub imgGenero_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieGenero)
End Sub

Private Sub imgalineacion_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieAlineacion)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then CerrarVentana
End Sub

Private Sub ImgVolver_Click()
    Call CerrarVentana
End Sub

Private Sub CerrarVentana()
On Error GoTo ErrHandler
  
    Call Engine_Audio.PlayMusic(MP3_Inicio & ".mp3")
    bShowTutorial = False

    Unload Me
    
    frmAccount.Visible = True
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CerrarVentana de frmCrearPersonaje.frm")
End Sub


Private Sub lstGenero_Click()
    AccountConnecting.UserGender = lstGenero.ListIndex + 1
On Error GoTo ErrHandler
  
    Call DarCuerpoYCabeza
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub lstGenero_Click de frmCrearPersonaje.frm")
End Sub

Private Sub lstProfesion_Click()
On Error GoTo ErrHandler
  
    AccountConnecting.UserClass = lstProfesion.ListIndex + 1
    
    Call UpdateStats
    Call UpdateEspecialidad(AccountConnecting.UserClass)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub lstProfesion_Click de frmCrearPersonaje.frm")
End Sub

Private Sub UpdateEspecialidad(ByVal eClase As eClass)
On Error GoTo ErrHandler
  
    lblEspecialidad.Caption = vEspecialidades(ListEnabledClasses(eClase))
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub UpdateEspecialidad de frmCrearPersonaje.frm")
End Sub

Private Sub lstRaza_Click()
    AccountConnecting.UserRace = lstRaza.ListIndex + 1
On Error GoTo ErrHandler
  
    Call DarCuerpoYCabeza
    
    Call UpdateStats
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub lstRaza_Click de frmCrearPersonaje.frm")
End Sub

Private Sub picCharacterLeft_Click()
On Error GoTo ErrHandler
    Dir = CheckDir(Dir + 1)
    Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub picCharacterLeft_Click de frmCrearPersonaje.frm")
End Sub

Private Sub picCharacterRight_Click()
On Error GoTo ErrHandler
    Dir = CheckDir(Dir - 1)
    Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub picCharacterRight_Click de frmCrearPersonaje.frm")
End Sub

Private Sub picHead_Click(Index As Integer)
On Error GoTo ErrHandler

    If Index = 0 Then
        UserHead = CheckCabeza(UserHead - 1)
    ElseIf Index = 2 Then
        UserHead = CheckCabeza(UserHead + 1)
    End If

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub picHead_Click de frmCrearPersonaje.frm")
End Sub


Private Sub DrawHead(ByVal Head As Integer, ByVal PicIndex As Integer)
On Error GoTo ErrHandler
  
    Dim Grh As Integer
    Grh = HeadData(Head).Head(Dir).GrhIndex
    
    Call UIBegin(mDevice(PicIndex), frmCrearPersonaje.picHead(PicIndex).ScaleWidth, frmCrearPersonaje.picHead(PicIndex).ScaleHeight, &H0)

    Call Mod_TileEngine.DrawGrhIndex(Grh, picHead(PicIndex).ScaleWidth / 2 - GrhData(Grh).pixelWidth / 2, 1, 0, 0)
      
    Call UIEnd
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub DrawHead de frmCrearPersonaje.frm")
End Sub

Private Sub picHeadLeft_Click()
On Error GoTo ErrHandler

    UserHead = CheckCabeza(UserHead - 1)
    
    Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub picHeadLeft_Click de frmCrearPersonaje.frm")
End Sub

Private Sub picHeadRight_Click()
On Error GoTo ErrHandler

    UserHead = CheckCabeza(UserHead + 1)
    
  Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ErrHandler de frmCrearPersonaje.frm")
End Sub

Private Sub tAnimacion_Timer()

    Dim Grh As Long
    Dim X As Long
    Dim Y As Long
    Static Frame As Byte
    
    If currentGrh = 0 Or UserHead = 0 Then Exit Sub
    Dim I   As Long

    UserHead = CheckCabeza(UserHead)
    
    Frame = Frame + 1
    If Frame >= GrhData(currentGrh).NumFrames Then Frame = 1

    Call DrawHead(CheckCabeza(UserHead + 1), 2)
    Call DrawHead(UserHead, 1)
    Call DrawHead(CheckCabeza(UserHead - 1), 0)
            
    Call UIBegin(mDevice(3), picPJ.ScaleWidth, picPJ.ScaleHeight, &H0)

    Grh = GrhData(currentGrh).Frames(Frame)
                
    X = picPJ.Width / 2 - GrhData(Grh).pixelWidth / 2
    Y = (picPJ.Height - GrhData(Grh).pixelHeight)
        
    Call Mod_TileEngine.DrawGrhIndex(Grh, X, Y, -1#, 0)
                    
    Grh = HeadData(UserHead).Head(Dir).GrhIndex
            
    X = picPJ.Width / 2 - GrhData(Grh).pixelWidth / 2
    Y = Y + BodyData(UserBody).HeadOffset.Y - 5
        
    Call Mod_TileEngine.DrawGrhIndex(Grh, X, Y, 0#, 0)
        
    Call UIEnd

End Sub

Private Sub txtNombre_Change()
    txtNombre.text = LTrim$(txtNombre.text)
    
    If Len(txtNombre.Text) > MAX_NICKNAME_SIZE Then
        txtNombre.Text = mid$(txtNombre.Text, 0, MAX_NICKNAME_SIZE)
    End If

End Sub

Private Sub DarCuerpoYCabeza()
On Error GoTo ErrHandler
  

    Dim bVisible As Boolean
    Dim PicIndex As Integer
    Dim LineIndex As Integer
    
    Select Case AccountConnecting.UserGender
        Case eGenero.Hombre
            Select Case AccountConnecting.UserRace
                Case eRaza.Humano
                    UserHead = HUMANO_H_PRIMER_CABEZA
                    UserBody = HUMANO_H_CUERPO_DESNUDO
                    
                Case eRaza.Elfo
                    UserHead = ELFO_H_PRIMER_CABEZA
                    UserBody = ELFO_H_CUERPO_DESNUDO
                    
                Case eRaza.ElfoOscuro
                    UserHead = DROW_H_PRIMER_CABEZA
                    UserBody = DROW_H_CUERPO_DESNUDO
                    
                Case eRaza.Enano
                    UserHead = ENANO_H_PRIMER_CABEZA
                    UserBody = ENANO_H_CUERPO_DESNUDO
                    
                Case eRaza.Gnomo
                    UserHead = GNOMO_H_PRIMER_CABEZA
                    UserBody = GNOMO_H_CUERPO_DESNUDO
                    
                Case Else
                    UserHead = 0
                    UserBody = 0
            End Select
            
        Case eGenero.Mujer
            Select Case AccountConnecting.UserRace
                Case eRaza.Humano
                    UserHead = HUMANO_M_PRIMER_CABEZA
                    UserBody = HUMANO_M_CUERPO_DESNUDO
                    
                Case eRaza.Elfo
                    UserHead = ELFO_M_PRIMER_CABEZA
                    UserBody = ELFO_M_CUERPO_DESNUDO
                    
                Case eRaza.ElfoOscuro
                    UserHead = DROW_M_PRIMER_CABEZA
                    UserBody = DROW_M_CUERPO_DESNUDO
                    
                Case eRaza.Enano
                    UserHead = ENANO_M_PRIMER_CABEZA
                    UserBody = ENANO_M_CUERPO_DESNUDO
                    
                Case eRaza.Gnomo
                    UserHead = GNOMO_M_PRIMER_CABEZA
                    UserBody = GNOMO_M_CUERPO_DESNUDO
                    
                Case Else
                    UserHead = 0
                    UserBody = 0
            End Select
        Case Else
            UserHead = 0
            UserBody = 0
    End Select
    
    bVisible = UserHead <> 0 And UserBody <> 0
    
    picHeadLeft.Visible = bVisible
    picHeadRight.Visible = bVisible
    picCharacterLeft.Visible = bVisible
    picCharacterRight.Visible = bVisible
    
    For PicIndex = 0 To 2
        picHead(PicIndex).Visible = bVisible
    Next PicIndex
    
    currentGrh = BodyData(UserBody).Walk(Dir).GrhIndex
    If currentGrh > 0 Then _
        tAnimacion.Interval = Round(GrhData(currentGrh).Speed / GrhData(currentGrh).NumFrames)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub DarCuerpoYCabeza de frmCrearPersonaje.frm")
End Sub

Private Function CheckCabeza(ByVal Head As Integer) As Integer
On Error GoTo ErrHandler
  

Select Case AccountConnecting.UserGender
    Case eGenero.Hombre
        Select Case AccountConnecting.UserRace
            Case eRaza.Humano
                If Head > HUMANO_H_ULTIMA_CABEZA Then
                    CheckCabeza = HUMANO_H_PRIMER_CABEZA + (Head - HUMANO_H_ULTIMA_CABEZA) - 1
                ElseIf Head < HUMANO_H_PRIMER_CABEZA Then
                    CheckCabeza = HUMANO_H_ULTIMA_CABEZA - (HUMANO_H_PRIMER_CABEZA - Head) + 1
                Else
                    CheckCabeza = Head
                End If
                
            Case eRaza.Elfo
                If Head > ELFO_H_ULTIMA_CABEZA Then
                    CheckCabeza = ELFO_H_PRIMER_CABEZA + (Head - ELFO_H_ULTIMA_CABEZA) - 1
                ElseIf Head < ELFO_H_PRIMER_CABEZA Then
                    CheckCabeza = ELFO_H_ULTIMA_CABEZA - (ELFO_H_PRIMER_CABEZA - Head) + 1
                Else
                    CheckCabeza = Head
                End If
                
            Case eRaza.ElfoOscuro
                If Head > DROW_H_ULTIMA_CABEZA Then
                    CheckCabeza = DROW_H_PRIMER_CABEZA + (Head - DROW_H_ULTIMA_CABEZA) - 1
                ElseIf Head < DROW_H_PRIMER_CABEZA Then
                    CheckCabeza = DROW_H_ULTIMA_CABEZA - (DROW_H_PRIMER_CABEZA - Head) + 1
                Else
                    CheckCabeza = Head
                End If
                
            Case eRaza.Enano
                If Head > ENANO_H_ULTIMA_CABEZA Then
                    CheckCabeza = ENANO_H_PRIMER_CABEZA + (Head - ENANO_H_ULTIMA_CABEZA) - 1
                ElseIf Head < ENANO_H_PRIMER_CABEZA Then
                    CheckCabeza = ENANO_H_ULTIMA_CABEZA - (ENANO_H_PRIMER_CABEZA - Head) + 1
                Else
                    CheckCabeza = Head
                End If
                
            Case eRaza.Gnomo
                If Head > GNOMO_H_ULTIMA_CABEZA Then
                    CheckCabeza = GNOMO_H_PRIMER_CABEZA + (Head - GNOMO_H_ULTIMA_CABEZA) - 1
                ElseIf Head < GNOMO_H_PRIMER_CABEZA Then
                    CheckCabeza = GNOMO_H_ULTIMA_CABEZA - (GNOMO_H_PRIMER_CABEZA - Head) + 1
                Else
                    CheckCabeza = Head
                End If
                
            Case Else
                AccountConnecting.UserRace = lstRaza.ListIndex + 1
                CheckCabeza = CheckCabeza(Head)
        End Select
        
    Case eGenero.Mujer
        Select Case AccountConnecting.UserRace
            Case eRaza.Humano
                If Head > HUMANO_M_ULTIMA_CABEZA Then
                    CheckCabeza = HUMANO_M_PRIMER_CABEZA + (Head - HUMANO_M_ULTIMA_CABEZA) - 1
                ElseIf Head < HUMANO_M_PRIMER_CABEZA Then
                    CheckCabeza = HUMANO_M_ULTIMA_CABEZA - (HUMANO_M_PRIMER_CABEZA - Head) + 1
                Else
                    CheckCabeza = Head
                End If
                
            Case eRaza.Elfo
                If Head > ELFO_M_ULTIMA_CABEZA Then
                    CheckCabeza = ELFO_M_PRIMER_CABEZA + (Head - ELFO_M_ULTIMA_CABEZA) - 1
                ElseIf Head < ELFO_M_PRIMER_CABEZA Then
                    CheckCabeza = ELFO_M_ULTIMA_CABEZA - (ELFO_M_PRIMER_CABEZA - Head) + 1
                Else
                    CheckCabeza = Head
                End If
                
            Case eRaza.ElfoOscuro
                If Head > DROW_M_ULTIMA_CABEZA Then
                    CheckCabeza = DROW_M_PRIMER_CABEZA + (Head - DROW_M_ULTIMA_CABEZA) - 1
                ElseIf Head < DROW_M_PRIMER_CABEZA Then
                    CheckCabeza = DROW_M_ULTIMA_CABEZA - (DROW_M_PRIMER_CABEZA - Head) + 1
                Else
                    CheckCabeza = Head
                End If
                
            Case eRaza.Enano
                If Head > ENANO_M_ULTIMA_CABEZA Then
                    CheckCabeza = ENANO_M_PRIMER_CABEZA + (Head - ENANO_M_ULTIMA_CABEZA) - 1
                ElseIf Head < ENANO_M_PRIMER_CABEZA Then
                    CheckCabeza = ENANO_M_ULTIMA_CABEZA - (ENANO_M_PRIMER_CABEZA - Head) + 1
                Else
                    CheckCabeza = Head
                End If
                
            Case eRaza.Gnomo
                If Head > GNOMO_M_ULTIMA_CABEZA Then
                    CheckCabeza = GNOMO_M_PRIMER_CABEZA + (Head - GNOMO_M_ULTIMA_CABEZA) - 1
                ElseIf Head < GNOMO_M_PRIMER_CABEZA Then
                    CheckCabeza = GNOMO_M_ULTIMA_CABEZA - (GNOMO_M_PRIMER_CABEZA - Head) + 1
                Else
                    CheckCabeza = Head
                End If
                
            Case Else
                AccountConnecting.UserRace = lstRaza.ListIndex + 1
                CheckCabeza = CheckCabeza(Head)
        End Select
    Case Else
        AccountConnecting.UserGender = lstGenero.ListIndex + 1
        CheckCabeza = CheckCabeza(Head)
End Select
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function CheckCabeza de frmCrearPersonaje.frm")
End Function

Private Function CheckDir(ByRef Dir As E_Heading) As E_Heading

On Error GoTo ErrHandler
  
    If Dir > E_Heading.WEST Then Dir = E_Heading.NORTH
    If Dir < E_Heading.NORTH Then Dir = E_Heading.WEST
    
    CheckDir = Dir
    
    currentGrh = BodyData(UserBody).Walk(Dir).GrhIndex
    If currentGrh > 0 Then _
        tAnimacion.Interval = Round(GrhData(currentGrh).Speed / GrhData(currentGrh).NumFrames)

  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function CheckDir de frmCrearPersonaje.frm")
End Function

Private Sub LoadHelp()
On Error GoTo ErrHandler
  
    vHelp(eHelp.iePasswd) = "La contraseña que utilizarás para conectar tu personaje al juego."
    vHelp(eHelp.ieMail) = "Es sumamente importante que ingreses una dirección de correo electrónico válida, ya que en el caso de perder la contraseña de tu personaje, se te enviará cuando lo requieras, a esa dirección."
    vHelp(eHelp.ieNombre) = "Sé cuidadoso al seleccionar el nombre de tu personaje. Argentum es un juego de rol, un mundo mágico y fantástico, y si seleccionás un nombre obsceno o con connotación política, los administradores borrarán tu personaje y no habrá ninguna posibilidad de recuperarlo."
    vHelp(eHelp.ieConfirmPasswd) = "La contraseña que utilizarás para conectar tu personaje al juego."
    vHelp(eHelp.ieAtributos) = "Son las cualidades que definen tu personaje."
    vHelp(eHelp.ieFuerza) = "De ella dependerá qué tan potentes serán tus golpes, tanto con armas de cuerpo a cuerpo, a distancia o sin armas."
    vHelp(eHelp.ieAgilidad) = "Este atributo intervendrá en qué tan bueno seas, tanto evadiendo como acertando golpes, respecto de otros personajes como de las criaturas a las q te enfrentes."
    vHelp(eHelp.ieInteligencia) = "Influirá de manera directa en cuánto maná ganarás por nivel."
    vHelp(eHelp.ieCarisma) = "Será necesario tanto para la relación con otros personajes (entrenamiento en parties) como con las criaturas (domar animales)."
    vHelp(eHelp.ieConstitucion) = "Afectará a la cantidad de vida que podrás ganar por nivel."
    vHelp(eHelp.ieEvasion) = "Evalúa la habilidad esquivando ataques físicos."
    vHelp(eHelp.ieMagia) = "Puntúa la cantidad de maná que se tendrá."
    vHelp(eHelp.ieVida) = "Valora la cantidad de salud que se podrá llegar a tener."
    vHelp(eHelp.ieEscudos) = "Estima la habilidad para rechazar golpes con escudos."
    vHelp(eHelp.ieArmas) = "Evalúa la habilidad en el combate cuerpo a cuerpo con armas."
    vHelp(eHelp.ieArcos) = "Evalúa la habilidad en el combate a distancia con arcos. "
    vHelp(eHelp.ieEspecialidad) = "Indica la fortaleza de un personaje en torno a sus habilidades."
    vHelp(eHelp.ieRaza) = "De la raza que elijas dependerá cómo se modifiquen los atributos. Podés cambiar de raza para poder visualizar cómo se modifican los distintos atributos."
    vHelp(eHelp.ieClase) = "La clase influirá en las características principales que tenga tu personaje, asi como en las magias e items que podrá utilizar. Las estrellas que ves abajo te mostrarán en qué habilidades se destaca la misma."
    vHelp(eHelp.ieGenero) = "Indica si el personaje será masculino o femenino. Esto influye en los items que podrá equipar."
    vHelp(eHelp.ieAlineacion) = "Indica si el personaje seguirá la senda del mal o del bien. (Actualmente deshabilitado)"
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadHelp de frmCrearPersonaje.frm")
End Sub

Private Sub ClearLabel()
On Error GoTo ErrHandler
  
    LastButtonPressed.ToggleToNormal
    lblHelp = ""
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ClearLabel de frmCrearPersonaje.frm")
End Sub

Private Sub txtNombre_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieNombre)
  
End Sub

Public Sub UpdateStats()
On Error GoTo ErrHandler
  
    Call UpdateRazaMod
    Call UpdateStars
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub UpdateStats de frmCrearPersonaje.frm")
End Sub

Private Sub UpdateRazaMod()
On Error GoTo ErrHandler
  
    Dim SelRaza As Integer
    Dim I As Integer
       
   SelRaza = lstRaza.ListIndex + 1
   
    With ModRaza(SelRaza)
        lblAtributoFinal(eAtributos.Fuerza).Caption = Val(18 + .Fuerza)
        lblAtributoFinal(eAtributos.Agilidad).Caption = Val(18 + .Agilidad)
        lblAtributoFinal(eAtributos.Inteligencia).Caption = Val(18 + .Inteligencia)
        lblAtributoFinal(eAtributos.Carisma).Caption = Val(18 + .Carisma)
        lblAtributoFinal(eAtributos.Constitucion).Caption = Val(18 + .Constitucion)
    End With
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub UpdateRazaMod de frmCrearPersonaje.frm")
End Sub

Private Sub UpdateStars()
On Error GoTo ErrHandler
  
    Dim NumStars As Double
    
    If AccountConnecting.UserClass = 0 Then Exit Sub
    
    ' Estrellas de evasion
    NumStars = (2.454 + 0.073 * Val(lblAtributoFinal(eAtributos.Agilidad).Caption)) * ModClase(AccountConnecting.UserClass).Evasion
    Call SetStars(imgEvasionStar, NumStars * 2)
    
    ' Estrellas de magia
    NumStars = ModClase(AccountConnecting.UserClass).Magia * Val(lblAtributoFinal(eAtributos.Inteligencia).Caption) * 0.085
    Call SetStars(imgMagiaStar, NumStars * 2)
    
    ' Estrellas de vida
    NumStars = 0.24 + (Val(lblAtributoFinal(eAtributos.Constitucion).Caption) * 0.5 - ModClase(AccountConnecting.UserClass).Vida) * 0.475
    Call SetStars(imgVidaStar, NumStars * 2)
    
    ' Estrellas de escudo
    NumStars = 4 * ModClase(AccountConnecting.UserClass).Escudo
    Call SetStars(imgEscudosStar, NumStars * 2)
    
    ' Estrellas de armas
    NumStars = (0.509 + 0.01185 * Val(lblAtributoFinal(eAtributos.Fuerza).Caption)) * ModClase(AccountConnecting.UserClass).Hit * _
                ModClase(AccountConnecting.UserClass).DañoArmas + 0.119 * ModClase(AccountConnecting.UserClass).AtaqueArmas * _
                Val(lblAtributoFinal(eAtributos.Agilidad).Caption)
    Call SetStars(imgArmasStar, NumStars * 2)
    
    ' Estrellas de arcos
    NumStars = (0.4915 + 0.01265 * Val(lblAtributoFinal(eAtributos.Fuerza).Caption)) * _
                ModClase(AccountConnecting.UserClass).DañoProyectiles * ModClase(AccountConnecting.UserClass).Hit + 0.119 * ModClase(AccountConnecting.UserClass).AtaqueProyectiles * _
                Val(lblAtributoFinal(eAtributos.Agilidad).Caption)
    Call SetStars(imgArcoStar, NumStars * 2)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub UpdateStars de frmCrearPersonaje.frm")
End Sub

Private Sub SetStars(ByRef ImgContainer As Object, ByVal NumStars As Integer)
On Error GoTo ErrHandler
  
    Dim FullStars As Integer
    Dim HasHalfStar As Boolean
    Dim Index As Integer
    Dim Counter As Integer

    If NumStars > 0 Then
        
        If NumStars > 10 Then NumStars = 10
        
        FullStars = Int(NumStars / 2)
        
        ' Tienen brillo extra si estan todas
        If FullStars = 5 Then
            For Index = 1 To FullStars
                ImgContainer(Index).Picture = picGlowStar
            Next Index
        Else
            ' Numero impar? Entonces hay que poner "media estrella"
            If (NumStars Mod 2) > 0 Then HasHalfStar = True
            
            ' Muestro las estrellas enteras
            If FullStars > 0 Then
                For Index = 1 To FullStars
                    ImgContainer(Index).Picture = picFullStar
                Next Index
                
                Counter = FullStars
            End If
            
            ' Muestro la mitad de la estrella (si tiene)
            If HasHalfStar Then
                Counter = Counter + 1
                
                ImgContainer(Counter).Picture = picHalfStar
            End If
            
            ' Si estan completos los espacios, no borro nada
            If Counter <> 5 Then
                ' Limpio las que queden vacias
                For Index = Counter + 1 To 5
                    Set ImgContainer(Index).Picture = Nothing
                Next Index
            End If
            
        End If
    Else
        ' Limpio todo
        For Index = 1 To 5
            Set ImgContainer(Index).Picture = Nothing
        Next Index
    End If

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SetStars de frmCrearPersonaje.frm")
End Sub

Private Sub LoadCharInfo()
On Error GoTo ErrHandler
  
    Dim SearchVar As String
    Dim I As Integer
    
    NroRazas = UBound(ListaRazas())
    NroClases = UBound(ListaClases())

    ReDim ModRaza(1 To NroRazas)
    ReDim ModClase(1 To NroClases)
    
    'Modificadores de Clase
    For I = 1 To NroClases
        With ModClase(I)
            SearchVar = ListaClases(I)
            
            .Evasion = Val(GetVar(IniPath & "CharInfo.dat", "MODEVASION", SearchVar))
            .AtaqueArmas = Val(GetVar(IniPath & "CharInfo.dat", "MODATAQUEARMAS", SearchVar))
            .AtaqueProyectiles = Val(GetVar(IniPath & "CharInfo.dat", "MODATAQUEPROYECTILES", SearchVar))
            .DañoArmas = Val(GetVar(IniPath & "CharInfo.dat", "MODDAÑOARMAS", SearchVar))
            .DañoProyectiles = Val(GetVar(IniPath & "CharInfo.dat", "MODDAÑOPROYECTILES", SearchVar))
            .Escudo = Val(GetVar(IniPath & "CharInfo.dat", "MODESCUDO", SearchVar))
            .Hit = Val(GetVar(IniPath & "CharInfo.dat", "HIT", SearchVar))
            .Magia = Val(GetVar(IniPath & "CharInfo.dat", "MODMAGIA", SearchVar))
            .Vida = Val(GetVar(IniPath & "CharInfo.dat", "MODVIDA", SearchVar))
        End With
    Next I
    
    'Modificadores de Raza
    For I = 1 To NroRazas
        With ModRaza(I)
            SearchVar = Replace(ListaRazas(I), " ", "")
        
            .Fuerza = Val(GetVar(IniPath & "CharInfo.dat", "MODRAZA", SearchVar + "Fuerza"))
            .Agilidad = Val(GetVar(IniPath & "CharInfo.dat", "MODRAZA", SearchVar + "Agilidad"))
            .Inteligencia = Val(GetVar(IniPath & "CharInfo.dat", "MODRAZA", SearchVar + "Inteligencia"))
            .Carisma = Val(GetVar(IniPath & "CharInfo.dat", "MODRAZA", SearchVar + "Carisma"))
            .Constitucion = Val(GetVar(IniPath & "CharInfo.dat", "MODRAZA", SearchVar + "Constitucion"))
        End With
    Next I

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadCharInfo de frmCrearPersonaje.frm")
End Sub

