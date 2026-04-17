VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11970
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   798
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.PictureBox picStoreButton 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4080
      ScaleHeight     =   495
      ScaleWidth      =   975
      TabIndex        =   34
      Top             =   8400
      Width           =   975
   End
   Begin ARGENTUM.AOPictureBox picInv 
      Height          =   2475
      Left            =   8940
      TabIndex        =   37
      Top             =   2580
      Width           =   2430
      _ExtentX        =   4286
      _ExtentY        =   4366
   End
   Begin ARGENTUM.AOPictureBox picMain 
      Height          =   6240
      Left            =   150
      TabIndex        =   36
      Top             =   2070
      Width           =   8160
      _ExtentX        =   14393
      _ExtentY        =   11007
   End
   Begin VB.PictureBox picSM 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   4
      Left            =   11415
      MousePointer    =   99  'Custom
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   32
      Top             =   8445
      Width           =   420
   End
   Begin VB.Timer CheckSavedPackets 
      Interval        =   1000
      Left            =   840
      Top             =   240
   End
   Begin VB.Timer tmrBlink 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   2280
      Top             =   240
   End
   Begin VB.PictureBox picSM 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   3
      Left            =   10995
      MousePointer    =   99  'Custom
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   24
      Top             =   8445
      Width           =   420
   End
   Begin VB.PictureBox picSM 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   2
      Left            =   10590
      MousePointer    =   99  'Custom
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   23
      Top             =   8445
      Width           =   420
   End
   Begin VB.PictureBox picSM 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   1
      Left            =   10215
      MousePointer    =   99  'Custom
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   22
      Top             =   8445
      Width           =   420
   End
   Begin VB.PictureBox picSM 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   0
      Left            =   9840
      MousePointer    =   99  'Custom
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   21
      Top             =   8445
      Width           =   420
   End
   Begin VB.Timer macrotrabajo 
      Enabled         =   0   'False
      Left            =   1800
      Top             =   240
   End
   Begin VB.Timer TrainingMacro 
      Enabled         =   0   'False
      Interval        =   3121
      Left            =   360
      Top             =   240
   End
   Begin VB.Timer Second 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1320
      Top             =   240
   End
   Begin RichTextLib.RichTextBox RecTxt 
      Height          =   1476
      Index           =   0
      Left            =   135
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   128
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   2593
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox RecTxt 
      Height          =   1485
      Index           =   1
      Left            =   135
      TabIndex        =   28
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   128
      Visible         =   0   'False
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   2619
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":0387
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox RecTxt 
      Height          =   1485
      Index           =   2
      Left            =   135
      TabIndex        =   29
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   128
      Visible         =   0   'False
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   2619
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":0404
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox RecTxt 
      Height          =   1485
      Index           =   3
      Left            =   135
      TabIndex        =   30
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   128
      Visible         =   0   'False
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   2619
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":0481
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox SendCMSTXT 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   120
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1622
      Visible         =   0   'False
      Width           =   8205
   End
   Begin VB.TextBox SendTxt 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   135
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1622
      Visible         =   0   'False
      Width           =   8205
   End
   Begin ARGENTUM.UCColdDownList hlst 
      Height          =   2760
      Left            =   8880
      TabIndex        =   35
      Top             =   2520
      Visible         =   0   'False
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   4868
      BarBackColor    =   0
      BackColor       =   0
      ForeColor       =   16777215
      BarOffsetY      =   -2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblMasteryPoints 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "255"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   210
      Left            =   10095
      TabIndex        =   33
      Top             =   1080
      Width           =   270
   End
   Begin VB.Label lblPorcLvl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "33.33%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   225
      Left            =   9120
      TabIndex        =   31
      Top             =   1440
      Width           =   2025
   End
   Begin VB.Shape shpExp 
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   195
      Left            =   9165
      Top             =   1455
      Width           =   2040
   End
   Begin VB.Image imgCerrar 
      Height          =   195
      Left            =   11655
      Top             =   150
      Width           =   195
   End
   Begin VB.Image imgMoveMagicUp 
      Height          =   225
      Left            =   11520
      Top             =   3240
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgMoveMagicDown 
      Height          =   225
      Left            =   11520
      Top             =   3480
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgPestania 
      Height          =   315
      Index           =   1
      Left            =   1410
      MousePointer    =   99  'Custom
      Top             =   1636
      Width           =   1260
   End
   Begin VB.Image imgPestania 
      Height          =   315
      Index           =   2
      Left            =   2520
      MousePointer    =   99  'Custom
      Top             =   1636
      Width           =   1245
   End
   Begin VB.Image imgPestania 
      Height          =   315
      Index           =   0
      Left            =   360
      MousePointer    =   99  'Custom
      Top             =   1636
      Width           =   1245
   End
   Begin VB.Image imgMapa 
      Height          =   195
      Left            =   10560
      Top             =   6720
      Width           =   1035
   End
   Begin VB.Image imgClanes 
      Height          =   195
      Left            =   10575
      Top             =   8040
      Width           =   1035
   End
   Begin VB.Image imgEstadisticas 
      Height          =   195
      Left            =   10575
      Top             =   7695
      Width           =   1035
   End
   Begin VB.Image imgOpciones 
      Height          =   195
      Left            =   10575
      Top             =   7365
      Width           =   1035
   End
   Begin VB.Image imgGrupo 
      Height          =   195
      Left            =   10575
      Top             =   7035
      Width           =   1035
   End
   Begin VB.Image imgAsignarSkill 
      Height          =   450
      Left            =   10695
      MousePointer    =   99  'Custom
      Top             =   900
      Width           =   450
   End
   Begin VB.Label lblDropGold 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   10560
      MousePointer    =   99  'Custom
      TabIndex        =   27
      Top             =   6240
      Width           =   255
   End
   Begin VB.Label lblMinimizar 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   11280
      MousePointer    =   99  'Custom
      TabIndex        =   26
      Top             =   240
      Width           =   255
   End
   Begin VB.Label lblCerrar 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   11160
      MousePointer    =   99  'Custom
      TabIndex        =   25
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   10320
      MousePointer    =   99  'Custom
      TabIndex        =   17
      Top             =   1920
      Width           =   1125
   End
   Begin VB.Image cmdInfo 
      Height          =   615
      Left            =   10680
      MousePointer    =   99  'Custom
      Top             =   5400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblMapName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   8700
      TabIndex        =   20
      Top             =   8445
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BetaTester"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   195
      Left            =   8760
      TabIndex        =   19
      Top             =   360
      Width           =   2625
   End
   Begin VB.Label lblLvl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "255"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   285
      Left            =   10020
      TabIndex        =   18
      Top             =   870
      Width           =   405
   End
   Begin VB.Image CmdLanzar 
      Height          =   615
      Left            =   8760
      MousePointer    =   99  'Custom
      Top             =   5400
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   8880
      MousePointer    =   99  'Custom
      TabIndex        =   16
      Top             =   1920
      Width           =   1395
   End
   Begin VB.Label GldLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   10920
      TabIndex        =   15
      Top             =   6270
      Width           =   90
   End
   Begin VB.Label lblStrg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   9720
      TabIndex        =   9
      Top             =   6225
      Width           =   210
   End
   Begin VB.Label lblDext 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   9120
      TabIndex        =   8
      Top             =   6225
      Width           =   210
   End
   Begin VB.Label Coord 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "000 X:00 Y: 00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   8595
      TabIndex        =   7
      Top             =   8550
      Width           =   1335
   End
   Begin VB.Label lblWeapon 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "99/99"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   7065
      TabIndex        =   6
      Top             =   8550
      Width           =   615
   End
   Begin VB.Label lblShielder 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "99/99"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5130
      TabIndex        =   5
      Top             =   8553
      Width           =   855
   End
   Begin VB.Label lblHelm 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "99/99"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2940
      TabIndex        =   4
      Top             =   8553
      Width           =   855
   End
   Begin VB.Label lblArmor 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "99/99"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1170
      TabIndex        =   3
      Top             =   8553
      Width           =   855
   End
   Begin VB.Image imgScroll 
      Height          =   240
      Index           =   1000
      Left            =   11400
      MousePointer    =   99  'Custom
      Top             =   3225
      Width           =   225
   End
   Begin VB.Image InvEqu 
      Height          =   4230
      Left            =   8715
      Top             =   1875
      Width           =   2970
   End
   Begin VB.Label lblMana 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "9999/9999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   8700
      TabIndex        =   11
      Top             =   7095
      Width           =   1095
   End
   Begin VB.Label lblEnergia 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "999/999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   8700
      TabIndex        =   10
      Top             =   6750
      Width           =   1095
   End
   Begin VB.Label lblVida 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "999/999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   8700
      TabIndex        =   12
      Top             =   7440
      Width           =   1095
   End
   Begin VB.Label lblHambre 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "999/999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   8700
      TabIndex        =   13
      Top             =   7785
      Width           =   1095
   End
   Begin VB.Label lblSed 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "999/999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   8700
      TabIndex        =   14
      Top             =   8130
      Width           =   1095
   End
   Begin VB.Shape shpEnergia 
      FillStyle       =   0  'Solid
      Height          =   120
      Left            =   8580
      Top             =   6795
      Width           =   1320
   End
   Begin VB.Shape shpMana 
      FillStyle       =   0  'Solid
      Height          =   120
      Left            =   8580
      Top             =   7140
      Width           =   1320
   End
   Begin VB.Shape shpVida 
      FillStyle       =   0  'Solid
      Height          =   120
      Left            =   8580
      Top             =   7485
      Width           =   1320
   End
   Begin VB.Shape shpHambre 
      FillStyle       =   0  'Solid
      Height          =   120
      Left            =   8580
      Top             =   7830
      Width           =   1320
   End
   Begin VB.Shape shpSed 
      BorderColor     =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   120
      Left            =   8580
      Top             =   8175
      Width           =   1320
   End
   Begin VB.Image imgPestania 
      Height          =   315
      Index           =   3
      Left            =   3645
      MousePointer    =   99  'Custom
      Top             =   1636
      Width           =   1260
   End
   Begin VB.Shape MainViewShp 
      BorderColor     =   &H00000000&
      Height          =   6240
      Left            =   150
      Top             =   2070
      Visible         =   0   'False
      Width           =   8160
   End
   Begin VB.Menu mnuObj 
      Caption         =   "Objeto"
      Visible         =   0   'False
      Begin VB.Menu mnuTirar 
         Caption         =   "Tirar"
      End
      Begin VB.Menu mnuUsar 
         Caption         =   "Usar"
      End
      Begin VB.Menu mnuEquipar 
         Caption         =   "Equipar"
      End
   End
   Begin VB.Menu mnuNpc 
      Caption         =   "NPC"
      Visible         =   0   'False
      Begin VB.Menu mnuNpcDesc 
         Caption         =   "Descripcion"
      End
      Begin VB.Menu mnuNpcComerciar 
         Caption         =   "Comerciar"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMain"
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

Public tX As Byte
Public tY As Byte
Public MouseX As Long
Public MouseY As Long
Public MouseBoton As Long
Public MouseShift As Long
Public ShiftKey As Boolean
Private clicX As Long
Private clicY As Long

Private Type tPestania
    ' Pics holders
    picNormal As Picture
    picSelected As Picture
    picRollover As Picture
    picNewMessages As Picture
    
    ' Flags
    Rollover As Boolean
    Selected As Boolean
    NewMessages As Boolean
End Type

Private MainTabs() As tPestania
Private eeCurrentSelectedTab As eConsoleType
Private eeCurrentRolloverTab As eConsoleType

Private LeftClicX As Long
Private LeftClicY As Long

Public IsPlaying As Byte

Private clsFormulario As clsFormMovementManager

Private cBotonDiamArriba As clsGraphicalButton
Private cBotonDiamAbajo As clsGraphicalButton
Private cBotonMapa As clsGraphicalButton
Private cBotonGrupo As clsGraphicalButton
Private cBotonOpciones As clsGraphicalButton
Private cBotonEstadisticas As clsGraphicalButton
Private cBotonClanes As clsGraphicalButton
Private cBotonAsignarSkill As clsGraphicalButton
Private cBotonMoverHechiArriba As clsGraphicalButton
Private cBotonMoverHechiAbajo As clsGraphicalButton
Private cBotonCerrarJuego As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton

Public picSkillStar As Picture

Private picOpenStore As Picture
Private picCloseStore As Picture

Private bLastBrightBlink As Boolean

Private WithEvents dragInventory As clsGraphicalInventory
Attribute dragInventory.VB_VarHelpID = -1

Public ExpPercLabelHoverValueSet As Boolean
Dim RealLabelX As Integer, RealLabelY As Integer


'Usado para controlar que no se dispare el binding de la tecla CTRL cuando se usa CTRL+Tecla.
Dim CtrlMaskOn As Boolean

Private Devices(0 To 4) As Long
Private Images(0 To 4) As Integer

Private Sub dragInventory_dragDone(ByVal originalSlot As Integer, ByVal newSlot As Integer, _
                                   ByVal moveType As eMoveType)
On Error GoTo ErrHandler
  
    Call Protocol.WriteMoveItem(originalSlot, newSlot, eMoveType.Inventory)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub dragInventory_dragDone de frmMain.frm")
End Sub

Public Sub Inicializar()
On Error GoTo ErrHandler
  
    
    If NoRes Then
        ' Handles Form movement (drag and drop).
        Set clsFormulario = New clsFormMovementManager
        clsFormulario.Initialize Me, 120
    End If
    
    ' Set the default cursor as soon as we start the form
    Call modCustomCursors.SetFormCursorDefault(Me)

    Me.Picture = LoadPicture(DirInterfaces & SELECTED_UI & "VentanaPrincipal.JPG")
    InvEqu.Picture = LoadPicture(DirInterfaces & SELECTED_UI & "CentroInventario.jpg")
    
    Set picOpenStore = LoadPicture(DirInterfaces & SELECTED_UI & "BotonTiendaVerde.jpg")
    Set picCloseStore = LoadPicture(DirInterfaces & SELECTED_UI & "BotonTiendaRojo.jpg")

    Call LoadButtons
    Call LoadTabs
       
    Set dragInventory = Inventario
    
    Me.Left = 0
    Me.Top = 0
    
    ' Detect links in console
    Dim cs As Integer
    
    For cs = 0 To 3
        EnableURLDetect RecTxt(cs).hwnd, Me.hwnd, cs
    Next cs

    Call Inventario.Initialize(frmMain.picInv, MAX_INVENTORY_SLOTS, , , , , , , , , True)

    CtrlMaskOn = False

    Dialogos.Font = FuentesJuego.FuenteBase
    
    Me.ScaleWidth = 800
    Me.ScaleHeight = 600
    
    For cs = 0 To 4
        Devices(cs) = Aurora_Graphics.CreatePassFromDisplay(picSM(cs).hwnd, picSM(cs).ScaleWidth, picSM(cs).ScaleHeight)
    Next cs
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Inicializar de frmMain.frm")
End Sub

Private Sub LoadButtons()
On Error GoTo ErrHandler
  
    Dim GrhPath As String
    Dim I As Integer
    
    GrhPath = DirInterfaces & SELECTED_UI

    Set cBotonDiamArriba = New clsGraphicalButton
    Set cBotonDiamAbajo = New clsGraphicalButton
    Set cBotonGrupo = New clsGraphicalButton
    Set cBotonOpciones = New clsGraphicalButton
    Set cBotonEstadisticas = New clsGraphicalButton
    Set cBotonClanes = New clsGraphicalButton
    Set cBotonAsignarSkill = New clsGraphicalButton
    Set cBotonMapa = New clsGraphicalButton
    Set cBotonMoverHechiArriba = New clsGraphicalButton
    Set cBotonMoverHechiAbajo = New clsGraphicalButton
    Set cBotonCerrarJuego = New clsGraphicalButton
    
    Set LastButtonPressed = New clsGraphicalButton
    
    
    'Call cBotonDiamArriba.Initialize(imgInvScrollUp, "", _
    '                                GrhPath & "BotonDiamArribaF.bmp", _
    '                                GrhPath & "BotonDiamArribaF.bmp", Me)

    'Call cBotonDiamAbajo.Initialize(imgInvScrollDown, "", _
    '                                GrhPath & "BotonDiamAbajoF.bmp", _
    '                                GrhPath & "BotonDiamAbajoF.bmp", Me)
    
    Call cBotonMapa.Initialize(imgMapa, GrhPath & "BotonMapa.jpg", _
                                    GrhPath & "BotonMapaRollover.jpg", _
                                    GrhPath & "BotonMapaClick.jpg", Me)
                                    
    Call cBotonGrupo.Initialize(imgGrupo, GrhPath & "BotonGrupo.jpg", _
                                    GrhPath & "BotonGrupoRollover.jpg", _
                                    GrhPath & "BotonGrupoClick.jpg", Me, , , , , _
                                    GrhPath & "BotonGrupoNuevoMensaje.jpg")

    Call cBotonOpciones.Initialize(imgOpciones, GrhPath & "BotonOpciones.jpg", _
                                    GrhPath & "BotonOpcionesRollover.jpg", _
                                    GrhPath & "BotonOpcionesClick.jpg", Me)

    Call cBotonEstadisticas.Initialize(imgEstadisticas, GrhPath & "BotonEstadisticas.jpg", _
                                    GrhPath & "BotonEstadisticasRollover.jpg", _
                                    GrhPath & "BotonEstadisticasClick.jpg", Me)

    Call cBotonClanes.Initialize(imgClanes, GrhPath & "BotonClanes.jpg", _
                                    GrhPath & "BotonClanesRollover.jpg", _
                                    GrhPath & "BotonClanesClick.jpg", Me)
                                    
    Call cBotonAsignarSkill.Initialize(imgAsignarSkill, GrhPath & "BotonSkill.jpg", _
                                    GrhPath & "BotonSkillRollover.jpg", _
                                    GrhPath & "BotonSkillClick.jpg", Me)
                                    
    Call cBotonMoverHechiArriba.Initialize(imgMoveMagicUp, GrhPath & "BotonMoverHechiArriba.jpg", _
                                    GrhPath & "BotonMoverHechiArribaRollover.jpg", _
                                    GrhPath & "BotonMoverHechiArribaClick.jpg", Me)

    Call cBotonMoverHechiAbajo.Initialize(imgMoveMagicDown, GrhPath & "BotonMoverHechiAbajo.jpg", _
                                    GrhPath & "BotonMoverHechiAbajoRollover.jpg", _
                                    GrhPath & "BotonMoverHechiAbajoClick.jpg", Me)
                                    
                                    
    Call cBotonCerrarJuego.Initialize(imgCerrar, GrhPath & "BotonCruzSalir.jpg", _
                                    GrhPath & "BotonCruzSalirRollover.jpg", _
                                    GrhPath & "BotonCruzSalirClick.jpg", Me)
                                    
                                    


    'Set picSkillStar = LoadPicture(GrhPath & "BotonAsignarSkills.bmp")

    imgAsignarSkill.Visible = SkillPoints > 0
    
    imgAsignarSkill.MouseIcon = picMouseIcon
    lblDropGold.MouseIcon = picMouseIcon
    lblCerrar.MouseIcon = picMouseIcon
    lblMinimizar.MouseIcon = picMouseIcon
    CmdLanzar.MouseIcon = picMouseIcon
    cmdInfo.MouseIcon = picMouseIcon
    Label7.MouseIcon = picMouseIcon
    Label4.MouseIcon = picMouseIcon
    
    
    For I = 0 To 4
        picSM(I).MouseIcon = picMouseIcon
    Next I
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadButtons de frmMain.frm")
End Sub

Private Sub LoadTabs()
On Error GoTo ErrHandler
  
    
    Dim GrhPath As String
    GrhPath = DirInterfaces & SELECTED_UI
    
    ReDim MainTabs(eConsoleType.Last - 1)
    
    With MainTabs(eConsoleType.General)
        Set .picNormal = LoadPicture(GrhPath & "PestaniaGeneralApagado.jpg")
        Set .picSelected = LoadPicture(GrhPath & "PestaniaGeneralSeleccionado.jpg")
        Set .picNewMessages = LoadPicture(GrhPath & "PestaniaGeneralMensajesNuevos.jpg")
        Set .picRollover = LoadPicture(GrhPath & "PestaniaGeneralRollover.jpg")
        
        Set imgPestania(eConsoleType.General) = .picSelected
        imgPestania(eConsoleType.General).MouseIcon = picMouseIcon
        
        .Selected = True
    End With
    
    With MainTabs(eConsoleType.Acciones)
        Set .picNormal = LoadPicture(DirInterfaces & SELECTED_UI & "PestaniaAccionesApagado.jpg")
        Set .picSelected = LoadPicture(DirInterfaces & SELECTED_UI & "PestaniaAccionesSeleccionado.jpg")
        Set .picNewMessages = LoadPicture(DirInterfaces & SELECTED_UI & "PestaniaAccionesMensajesNuevos.jpg")
        Set .picRollover = LoadPicture(DirInterfaces & SELECTED_UI & "PestaniaAccionesRollover.jpg")
        
        Set imgPestania(eConsoleType.Acciones) = .picNormal
        imgPestania(eConsoleType.Acciones).MouseIcon = picMouseIcon
    End With
    
    With MainTabs(eConsoleType.Agrupaciones)
        Set .picNormal = LoadPicture(DirInterfaces & SELECTED_UI & "PestaniaAgrupacionesApagado.jpg")
        Set .picSelected = LoadPicture(DirInterfaces & SELECTED_UI & "PestaniaAgrupacionesSeleccionado.jpg")
        Set .picNewMessages = LoadPicture(DirInterfaces & SELECTED_UI & "PestaniaAgrupacionesMensajesNuevos.jpg")
        Set .picRollover = LoadPicture(DirInterfaces & SELECTED_UI & "PestaniaAgrupacionesRollover.jpg")
        
        Set imgPestania(eConsoleType.Agrupaciones) = .picNormal
        imgPestania(eConsoleType.Agrupaciones).MouseIcon = picMouseIcon
    End With
    
    With MainTabs(eConsoleType.Custom)
        Set .picNormal = LoadPicture(DirInterfaces & SELECTED_UI & "PestaniaPersonalizadaApagado.jpg")
        Set .picSelected = LoadPicture(DirInterfaces & SELECTED_UI & "PestaniaPersonalizadaSeleccionado.jpg")
        Set .picNewMessages = LoadPicture(DirInterfaces & SELECTED_UI & "PestaniaPersonalizadaMensajesNuevos.jpg")
        Set .picRollover = LoadPicture(DirInterfaces & SELECTED_UI & "PestaniaPersonalizadaRollover.jpg")
        
        Set imgPestania(eConsoleType.Custom) = .picNormal
        imgPestania(eConsoleType.Custom).MouseIcon = picMouseIcon
    End With

    eeCurrentSelectedTab = eConsoleType.General
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadTabs de frmMain.frm")
End Sub

Public Sub LightSkillStar(ByVal bTurnOn As Boolean)
On Error GoTo ErrHandler
  
        imgAsignarSkill.Visible = bTurnOn
   
  Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LightSkillStar de frmMain.frm")
End Sub

Public Sub MoveSpell(ByVal Direction As Byte)
On Error GoTo ErrHandler
  
    If hlst.Visible = True Then
        Dim Success As Boolean
        Dim ItemToMove As Integer
        ItemToMove = hlst.ListIndex + 1
        Select Case Direction
            Case 1 'subir
                Success = hlst.TryMoveItemUp()
            Case 0 'bajar
                Success = hlst.TryMoveItemDown()
        End Select
        If Success Then
            Call WriteMoveSpell(Direction = 1, ItemToMove)
        End If
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub MoveSpell de frmMain.frm")
End Sub

Public Sub ActivarMacroHechizos()
On Error GoTo ErrHandler
  
    If Not hlst.Visible Then
        Call AddtoRichTextBox(frmMain.RecTxt(0), "Debes tener seleccionado el hechizo para activar el auto-lanzar", 0, 200, 200, False, True, True, eMessageType.combate)
        Exit Sub
    End If
    
    TrainingMacro.Interval = PlayerData.Intervals.SpellCastMacro
    TrainingMacro.Enabled = True
    Call AddtoRichTextBox(frmMain.RecTxt(0), "Auto lanzar hechizos activado", 0, 200, 200, False, True, True, eMessageType.combate)
    Call ControlSM(eSMType.mSpells, True)

      
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ActivarMacroHechizos de frmMain.frm")
End Sub

Public Sub DesactivarMacroHechizos()
On Error GoTo ErrHandler
  
    TrainingMacro.Enabled = False
    Call AddtoRichTextBox(frmMain.RecTxt(0), "Auto lanzar hechizos desactivado", 0, 150, 150, False, True, True, eMessageType.combate)
    Call ControlSM(eSMType.mSpells, False)

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub DesactivarMacroHechizos de frmMain.frm")
End Sub

Public Sub ControlSM(ByVal Index As Byte, ByVal Mostrar As Boolean)
On Error GoTo ErrHandler
    
    Select Case Index
        Case eSMType.sResucitation
                
            If Mostrar Then
                Images(Index) = 4978
                
                Call AddtoRichTextBox(frmMain.RecTxt(0), MENSAJE_SEGURO_RESU_ON, 0, 255, 0, True, False, True, eMessageType.Info)
                picSM(Index).ToolTipText = "Seguro de resucitación activado."
            Else
                Images(Index) = 4982
                Call AddtoRichTextBox(frmMain.RecTxt(0), MENSAJE_SEGURO_RESU_OFF, 255, 0, 0, True, False, True, eMessageType.Info)
                picSM(Index).ToolTipText = "Seguro de resucitación desactivado."
            End If
            
        Case eSMType.sSafemode
        
            If Mostrar Then
                Images(Index) = 4979
                Call AddtoRichTextBox(frmMain.RecTxt(0), MENSAJE_SEGURO_ACTIVADO, 0, 255, 0, True, False, True, eMessageType.Info)
                picSM(Index).ToolTipText = "Seguro activado."
            Else
                Images(Index) = 4983
                Call AddtoRichTextBox(frmMain.RecTxt(0), MENSAJE_SEGURO_DESACTIVADO, 255, 0, 0, True, False, True, eMessageType.Info)
                picSM(Index).ToolTipText = "Seguro desactivado."
                
                If charlist(UserCharIndex).Criminal = 1 Then
                    Call AddtoRichTextBox(frmMain.RecTxt(0), MENSAJE_SEGURO_ADVIERTE, _
                                          65, 190, 156, False, False, True, eMessageType.Info)
                End If
            End If
            
        Case eSMType.mSpells
            If Mostrar Then
                Images(Index) = 4980
                picSM(Index).ToolTipText = "Macro de hechizos activado."
            Else
                Images(Index) = 4984
                picSM(Index).ToolTipText = "Macro de hechizos desactivado."
            End If
            
        Case eSMType.mWork
            If Mostrar Then
                Images(Index) = 4981
                picSM(Index).ToolTipText = "Macro de trabajo activado."
            Else
                Images(Index) = 4985
                picSM(Index).ToolTipText = "Macro de trabajo desactivado."
            End If
        Case eSMType.mPets
            Images(Index) = 26147
            
            picSM(Index).ToolTipText = "Abrir ventana de mascotas"
    End Select

    SMStatus(Index) = Mostrar
    
    Call Invalidate(picSM(Index).hwnd)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ControlSM de frmMain.frm")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = 16 And (Shift = 1)) Then
On Error GoTo ErrHandler
  
        ShiftKey = True
    End If

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Form_KeyDown de frmMain.frm")
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    
'***************************************************
'Autor: Unknown
'Last Modification: 18/11/2010
'18/11/2009: ZaMa - Ahora se pueden poner comandos en los mensajes personalizados (execpto guildchat y privados)
'18/11/2010: Amraphen - Agregué el handle correspondiente para las nuevas configuraciones de teclas (CTRL+0..9).
'***************************************************
On Error GoTo ErrHandler
  
#If EnableSecurity Then
    If LOGGING Then Call CheatingDeath.StoreKey(KeyCode, False)
#End If

    'Cambio de pestañas (D'Artagnan)
    If (Not SendTxt.Visible) And (Not SendCMSTXT.Visible) Then
    
        'Verificamos si se está presionando la tecla CTRL.
        If Shift = 2 Then
            If KeyCode >= vbKey0 And KeyCode <= vbKey9 Then
                If KeyCode = vbKey0 Then
                    'Si es CTRL+0 muestro la ventana de configuración de teclas.
                    Call frmCustomKeys.Show(, Me)
                    
                ElseIf KeyCode >= vbKey1 And KeyCode <= vbKey9 Then
                    'Si es CTRL+1..9 cambio la configuración.
                    If KeyCode - vbKey0 = CustomKeys.CurrentConfig Then Exit Sub
                    
                    CustomKeys.CurrentConfig = KeyCode - vbKey0
                    
                    Dim sMsg As String
                    
                    sMsg = "¡Se ha cargado la configuración "
                    If CustomKeys.CurrentConfig = 0 Then
                        sMsg = sMsg & "default"
                    Else
                        sMsg = sMsg & "perzonalizada número " & CStr(CustomKeys.CurrentConfig)
                    End If
                    sMsg = sMsg & "!"

                    Call ShowConsoleMsg(sMsg, 255, 255, 255, True)
                End If
                
                CtrlMaskOn = True
                Exit Sub
            End If
        End If
        
        If KeyCode = 16 Then
            ShiftKey = False
        End If
        
        
        'Cambio de pestañas (D'Artagnan)
        If KeyCode = vbKeyTab Then
            If GetAsyncKeyState(vbKeyControl) Then
                If eeCurrentSelectedTab < eConsoleType.Last - 1 Then
                    Call imgPestania_Click(eeCurrentSelectedTab + 1)
                Else
                    Call imgPestania_Click(0)
                End If
            End If
        End If
        
        If KeyCode = vbKeyControl Then
            'Chequeo que no se haya usado un CTRL + tecla antes de disparar las bindings.
            If CtrlMaskOn Then
                CtrlMaskOn = False
                Exit Sub
            End If
        End If
        
        'Checks if the key is valid
        If LenB(CustomKeys.ReadableName(KeyCode)) > 0 Then
            Select Case KeyCode
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleMusic)
                    GameConfig.Sounds.bMusicEnabled = Not GameConfig.Sounds.bMusicEnabled
                    Engine_Audio.MusicEnabled = GameConfig.Sounds.bMusicEnabled
                    
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleSound)
                    GameConfig.Sounds.bSoundEffectsEnabled = Not GameConfig.Sounds.bSoundEffectsEnabled
                    Engine_Audio.EffectEnabled = GameConfig.Sounds.bSoundEffectsEnabled
                    IsPlaying = PlayLoop.plNone

                Case CustomKeys.BindedKey(eKeyType.mKeyGetObject)
                    Call AgarrarItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyEquipObject)
                    Call EquiparItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleNames)
                    Nombres = (Nombres + 1) Mod 3
                    GameConfig.Extras.NameStyle = Nombres
                    
                Case CustomKeys.BindedKey(eKeyType.mKeyTamAnimal)
                    If UserEstado = 1 Then
                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                        End With
                    Else
                        Call WriteWork(eSkill.Domar)
                    End If
                    
                Case CustomKeys.BindedKey(eKeyType.mKeySteal)
                    If UserEstado = 1 Then
                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                        End With
                    Else
                        Call WriteWork(eSkill.Robar)
                    End If
                    
                Case CustomKeys.BindedKey(eKeyType.mKeyHide)
                    If UserEstado = 1 Then
                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                        End With
                    Else
                        Call WriteWork(eSkill.Ocultarse)
                    End If
                                    
                Case CustomKeys.BindedKey(eKeyType.mKeyDropObject)
                    Call TirarItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyUseObject)
                    If macrotrabajo.Enabled Then Call DesactivarMacroTrabajo
                        
                    If MainTimer.Check(TimersIndex.UseItemWithU) Then
                        Call UsarItem
                    End If
                
                Case CustomKeys.BindedKey(eKeyType.mKeyRequestRefresh)
                    If MainTimer.Check(TimersIndex.SendRPU) Then
                        Call WriteRequestPositionUpdate
                        Beep
                    End If
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleSafeMode)
                    Call WriteSafeToggle

                Case CustomKeys.BindedKey(eKeyType.mKeyToggleResuscitationSafe)
                    Call WriteResuscitationToggle
            End Select
        Else
            
            'Evito que se muestren los mensajes personalizados cuando se cambie una configuración de teclas.
            If Shift = 2 Then Exit Sub
            
            Select Case KeyCode
                'Custom messages!
                Case vbKey0 To vbKey9
                    Dim CustomMessage As String
                    
                    CustomMessage = CustomMessages.Message((KeyCode - 39) Mod 10)
                    If LenB(CustomMessage) <> 0 Then
                        ' No se pueden mandar mensajes personalizados de clan o privado!
                        If UCase$(Left$(CustomMessage, 5)) <> "/CMSG" And _
                            Left$(CustomMessage, 1) <> "\" Then
                            
                            Call ParseUserCommand(CustomMessage)
                        End If
                    End If
            End Select
        End If
    End If
    
    Select Case KeyCode
        Case CustomKeys.BindedKey(eKeyType.mKeyTalkWithGuild)
            If SendTxt.Visible Then Exit Sub
            
            If (Not Comerciando) And (Not MirandoAsignarSkills) And _
              (Not frmMSG.Visible) And (Not MirandoForo) And _
              (Not MirandoEstadisticas) And (Not frmCantidad.Visible) Then
                SendCMSTXT.Visible = True
                SendCMSTXT.SetFocus
            End If
        
        Case CustomKeys.BindedKey(eKeyType.mKeyTakeScreenShot)
            Call ScreenCapture
            
        Case CustomKeys.BindedKey(eKeyType.mKeyShowOptions)
            Call frmOpciones.Show(vbModeless, frmMain)
        
        Case CustomKeys.BindedKey(eKeyType.mKeyMeditate)
            If UserMinMAN = UserMaxMAN Then Exit Sub
            
            If UserEstado = 1 Then
                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                End With
                Exit Sub
            End If
            
            If MainTimer.Check(TimersIndex.Meditate) Then
                Call RequestMeditate
            Else
                AddtoRichTextBox frmMain.RecTxt(0), "No tan rápido..!", 255, 255, 255, False, False, True, eMessageType.Info
            End If
        
        Case CustomKeys.BindedKey(eKeyType.mKeyCastSpellMacro)
            If UserEstado = 1 Then
                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                End With
                Exit Sub
            End If
            
            If TrainingMacro.Enabled Then
                DesactivarMacroHechizos
            Else
                ActivarMacroHechizos
            End If
        
        Case CustomKeys.BindedKey(eKeyType.mKeyWorkMacro)
            If UserEstado = 1 Then
                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                End With
                Exit Sub
            End If
            
            If macrotrabajo.Enabled Then
                Call DesactivarMacroTrabajo
            Else
                Call ActivarMacroTrabajo
            End If
        
        Case CustomKeys.BindedKey(eKeyType.mKeyExitGame)
            If UserParalizado Then 'Inmo
                Call modMessages.ShowConsoleMessage(modMessages.eMessageId.Cant_Quit_Paralized)
                Exit Sub
            End If
        
            If frmMain.macrotrabajo.Enabled Then Call DesactivarMacroTrabajo
            Call WriteQuit
            
        'Case CustomKeys.BindedKey(eKeyType.mKeyPetPanel)
        '    Unload frmMascotas
        '    Call frmMascotas.Show(, Me)
            
        Case CustomKeys.BindedKey(eKeyType.mKeyAttack)
            If Shift <> 0 Then Exit Sub
            
            If Not esGM(UserCharIndex) Then
                
                If Not MainTimer.Check(TimersIndex.Arrows, False) Then Exit Sub 'Check if arrows interval has finished.
                If Not MainTimer.Check(TimersIndex.CastSpell, False) Then 'Check if spells interval has finished.
                    If Not MainTimer.Check(TimersIndex.CastAttack, False) Then Exit Sub 'Corto intervalo Golpe-Hechizo
                Else
                    If Not MainTimer.Check(TimersIndex.Attack, False) Then Exit Sub
                End If
                
            End If
            
            If UserDescansar Or UserMeditar Then Exit Sub
            
            If TrainingMacro.Enabled Then Call DesactivarMacroHechizos
            If macrotrabajo.Enabled Then Call DesactivarMacroTrabajo

            If frmCustomKeys.Visible Then Exit Sub 'Chequeo si está visible la ventana de configuración de teclas.
            
            Call frmMain.hlst.StartSpellAfterMelee
            
            Call RecalculateMousePointerForSpell
            
            Call WriteAttack
            
        Case CustomKeys.BindedKey(eKeyType.mKeyTalk)
            If SendCMSTXT.Visible Then Exit Sub
            
            If (Not Comerciando) And (Not MirandoAsignarSkills) And _
              (Not frmMSG.Visible) And (Not MirandoForo) And _
              (Not MirandoEstadisticas) And (Not frmCantidad.Visible) Then
                SendTxt.Visible = True
                SendTxt.SetFocus
            End If
            
    End Select
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Form_KeyUp de frmMain.frm")
End Sub

Private Sub Form_Load()
    Me.Width = 800 * Screen.TwipsPerPixelX
On Error GoTo ErrHandler
  
    Me.Height = 602 * Screen.TwipsPerPixelY

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Form_Load de frmMain.frm")
End Sub

Private Sub imgCerrar_Click()

On Error GoTo ErrHandler

    prgRun = False
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgCerrar_Click de frmMain.frm")
End Sub



Private Sub imgMoveMagicDown_Click()
    Call MoveSpell(0)
  
End Sub

Private Sub imgMoveMagicUp_Click()
    Call MoveSpell(1)

End Sub

Private Sub lblLvl_Click()
On Error GoTo ErrHandler
  
    If UserLvl = 42 Then
        Call AddtoRichTextBox(frmMain.RecTxt(0), "Hás logrado llegar a nivel máximo. Ahora, por cada barra de experiencia que completes ganarás un punto de maestría, canjeable por beneficios dentro del juego.", 0, 200, 200, False, True, True, eMessageType.Trabajo)
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub lblLvl_Click de frmMain.frm")
End Sub

Private Sub lblPorcLvl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RealLabelX = X / 15
    RealLabelY = Y / 15
    

    With lblPorcLvl
        If (RealLabelX <= 1) Or (RealLabelY <= 1) Or (RealLabelX >= .Width - 2) Or (RealLabelY >= .Height - 2) Then
            ExpPercLabelHoverValueSet = False
            Call ShowLevelCompletionPerc
        Else
            If Not ExpPercLabelHoverValueSet Then
                ExpPercLabelHoverValueSet = True
                Call ShowLevelExpRequired
            End If
        End If
    End With
End Sub

Private Sub picMain_DragEnd(ByVal Source As ARGENTUM.AOPictureBox, ByVal Shift As Boolean, ByVal X As Single, ByVal Y As Single)

On Error GoTo ErrHandler
  
    If (Source.hwnd = picMain.hwnd Or Comerciando Or ViewingFormCantMove) Then Exit Sub
  
    If Shift Then
        Call TirarItem
    Else
        Dim tX As Byte
        Dim tY As Byte
        Call ConvertCPtoTP(X, Y, tX, tY)
    
        Protocol.WriteDropXY Inventario.slotDragged, 1, tX, tY
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub picMain_DragFinish de frmMain.frm")
End Sub

Private Sub picMain_GotFocus()
If SendCMSTXT.Visible Then
On Error GoTo ErrHandler
  
SendCMSTXT.SetFocus
End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub picMain_GotFocus de frmMain.frm")
End Sub

Private Sub picMain_KeyPress(KeyAscii As Integer)

On Error GoTo ErrHandler
  
If SendCMSTXT.Visible Then
SendCMSTXT.SetFocus
    
    SendCMSTXT_KeyPress KeyAscii
End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub picMain_KeyPress de frmMain.frm")
End Sub

Private Sub picMain_KeyUp(KeyCode As Integer, Shift As Integer)
'If SendCMSTXT.Visible Then
'SendCMSTXT.SetFocus
On Error GoTo ErrHandler
  
    
 '   SendCMSTXT_KeyUp KeyCode, Shift
'End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub picMain_KeyUp de frmMain.frm")
End Sub

Private Sub picMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If SendCMSTXT.Visible Then
On Error GoTo ErrHandler
  
SendCMSTXT.SetFocus
End If
If SendTxt.Visible Then
    SendTxt.SetFocus
End If

MouseBoton = Button
    MouseShift = Shift
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub picMain_MouseDown de frmMain.frm")
End Sub

Private Sub picMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If SendCMSTXT.Visible Then
On Error GoTo ErrHandler
  
SendCMSTXT.SetFocus
End If
If SendTxt.Visible Then
    SendTxt.SetFocus
End If
    clicX = X
    clicY = Y
  
  Exit Sub

ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub picMain_MouseUp de frmMain.frm")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If prgRun = True Then
On Error GoTo ErrHandler
  
        prgRun = False
        Cancel = 1
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Form_QueryUnload de frmMain.frm")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DisableURLDetect
On Error GoTo ErrHandler
  
#If Testeo = 0 Then
    'modMusicHook.DelMusicHook
#End If
      
    Dim cs As Integer
    For cs = 0 To 4
       Call Aurora_Graphics.DeletePass(Devices(cs))
    Next cs
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Form_Unload de frmMain.frm")
End Sub

Private Sub imgAsignarSkill_Click()
    OrigenSkills = eOrigenSkills.ieAsignacion
On Error GoTo ErrHandler
  
    Call WriteRequestSkills
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgAsignarSkill_Click de frmMain.frm")
End Sub

Private Sub imgGrupo_Click()
    If Not Mod_General.PartyInvitationEmpty Then
        Call Mod_General.PartyPendingInvitation
        Call Mod_General.PartyTempInviClear
        Exit Sub
    End If
    Call frmSelectGrupo.Show(False, Me)
End Sub

Private Sub imgClanes_Click()
On Error GoTo ErrHandler

    If Not Guilds.InvitationEmpty Then
        Call Guilds.GuildPendingInvitation
        Exit Sub
    End If
    If PlayerData.Guild.IdGuild = 0 Then
        Call ShowConsoleMsg("No pertences a ningún clan.", 100, 100, 100, False, False)
        Exit Sub
    End If
    
    Call frmGuildMain.LoadForm(frmGuildInformation, "")
    Call frmGuildMain.ShowPartial
  Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgClanes_Click de frmMain.frm")
End Sub

Private Sub imgEstadisticas_Click()
    OrigenSkills = eOrigenSkills.ieEstadisticas
On Error GoTo ErrHandler
  
    Call WriteRequestStadictis
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgEstadisticas_Click de frmMain.frm")
End Sub

Private Sub imgMapa_Click()
On Error GoTo ErrHandler

    Call ShellExecute(0, "Open", MAP_URL, "", App.path, SW_SHOWNORMAL)
    Exit Sub
    
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgMapa_Click de frmMain.frm")
End Sub

Private Sub imgOpciones_Click()
On Error GoTo ErrHandler

    Call frmOpciones.Show(vbModeless, frmMain)
    Exit Sub
    
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgOpciones_Click de frmMain.frm")
End Sub

Private Sub InvEqu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal

End Sub

Private Sub lblMinimizar_Click()
    Me.WindowState = 1
  
End Sub

Private Sub macrotrabajo_Timer()
On Error GoTo ErrHandler
  
    If Inventario.SelectedItem = 0 Then
        Call DesactivarMacroTrabajo
        Exit Sub
    End If
    
    'Macros are disabled if not using Argentum!
    'If Not Application.IsAppActive() Then  'Implemento lo propuesto por GD, se puede usar macro aun que se esté en otra ventana
    '    Call DesactivarMacroTrabajo
    '    Exit Sub
    'End If
    
    If UsingSkill = eSkill.Pesca Or UsingSkill = eSkill.Talar Or UsingSkill = eSkill.Mineria Or _
                UsingSkill = FundirMetal Or (UsingSkill = eSkill.Herreria And Not MirandoHerreria) Then
        Call WriteWorkLeftClick(tX, tY, UsingSkill)
        UsingSkill = 0
    End If
    
    'If Inventario.OBJType(Inventario.SelectedItem) = eObjType.otWeapon Then
     If Not MirandoCarpinteria Then Call UsarItem
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub macrotrabajo_Timer de frmMain.frm")
End Sub

Public Sub ActivarMacroTrabajo()
On Error GoTo ErrHandler
  
    macrotrabajo.Enabled = PlayerData.Intervals.WorkMacro
    macrotrabajo.Enabled = True
    Call AddtoRichTextBox(frmMain.RecTxt(0), "Macro Trabajo ACTIVADO", 0, 200, 200, False, True, True, eMessageType.Trabajo)
    Call ControlSM(eSMType.mWork, True)

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ActivarMacroTrabajo de frmMain.frm")
End Sub

Public Sub DesactivarMacroTrabajo()
On Error GoTo ErrHandler
  
    macrotrabajo.Enabled = False
    UsingSkill = 0
    MousePointer = vbDefault
    Call AddtoRichTextBox(frmMain.RecTxt(0), "Macro Trabajo DESACTIVADO", 0, 200, 200, False, True, True, eMessageType.Trabajo)
    Call ControlSM(eSMType.mWork, False)

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub DesactivarMacroTrabajo de frmMain.frm")
End Sub

Private Sub mnuEquipar_Click()
    Call EquiparItem
  
End Sub

Private Sub mnuNPCComerciar_Click()
On Error GoTo ErrHandler
  
    Call WriteLeftClick(tX, tY)
    Call WriteCommerceStart
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub mnuNPCComerciar_Click de frmMain.frm")
End Sub

Private Sub mnuNpcDesc_Click()
    Call WriteLeftClick(tX, tY)
  
End Sub

Private Sub mnuTirar_Click()
    Call TirarItem

End Sub

Private Sub mnuUsar_Click()
    Call UsarItem

End Sub

Private Sub Coord_Click()
    Call AddtoRichTextBox(frmMain.RecTxt(0), "Estas coordenadas son tu ubicación en el mapa. Utiliza la letra L para corregirla si esta no se corresponde con la del servidor por efecto del Lag.", 255, 255, 255, False, False, True, eMessageType.Info)

End Sub

Private Sub picSM_DblClick(Index As Integer)
On Error GoTo ErrHandler
  
Select Case Index

    Case eSMType.sResucitation
        Call WriteResuscitationToggle
        
    Case eSMType.sSafemode
        Call WriteSafeToggle
        
    Case eSMType.mSpells
        If UserEstado = 1 Then
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
            End With
            Exit Sub
        End If
        
        If TrainingMacro.Enabled Then
            Call DesactivarMacroHechizos
        Else
            Call ActivarMacroHechizos
        End If
        
    Case eSMType.mWork
        If UserEstado = 1 Then
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
            End With
            Exit Sub
        End If
        
        If macrotrabajo.Enabled Then
            Call DesactivarMacroTrabajo
        Else
            Call ActivarMacroTrabajo
        End If
    Case eSMType.mPets
        Unload frmMascotas
        Call frmMascotas.Show(vbModeless, frmMain)
        Exit Sub
        
End Select

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub picSM_DblClick de frmMain.frm")
End Sub


Private Sub picSM_Paint(Index As Integer)
    Call UIBegin(Devices(Index), picSM(Index).ScaleWidth, picSM(Index).ScaleHeight, &H0)

    Call DrawGrhIndex(Images(Index), 0, 0, 0, 0)
                
    Call UIEnd
End Sub

Private Sub picStoreButton_Click()
    If Not PlayerData.CurrentMap.CraftingStoreAllowed Then Exit Sub
    
    Call WriteWorkerStore_WorkerStoreGetRecipes
End Sub

Private Sub RecTxt_Change(Index As Integer)

'el .SetFocus causaba errores al salir y volver a entrar
On Error GoTo ErrHandler
  

If Not Application.IsAppActive() Then Exit Sub
    
If SendTxt.Visible Then
    SendTxt.SetFocus
ElseIf Me.SendCMSTXT.Visible Then
    SendCMSTXT.SetFocus
ElseIf (Not Comerciando) And (Not MirandoAsignarSkills) And _
    (Not frmMSG.Visible) And (Not MirandoForo) And _
    (Not MirandoEstadisticas) And (Not frmCantidad.Visible) And (Not MirandoParty) Then
             
    If picInv.Visible Then
        picInv.SetFocus
    ElseIf hlst.Visible Then
        hlst.SetFocus
    End If
    
End If
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub RecTxt_Change de frmMain.frm")
End Sub

Private Sub RecTxt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If picInv.Visible Then
On Error GoTo ErrHandler
  
    picInv.SetFocus
Else
    hlst.SetFocus
End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub RecTxt_KeyDown de frmMain.frm")
End Sub

Private Sub RecTxt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    StartCheckingLinks
On Error GoTo ErrHandler
  
    RestoreTabs
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub RecTxt_MouseMove de frmMain.frm")
End Sub

Private Sub SendTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    
    ' Control + Shift
    If Shift = 3 Then
        On Error GoTo ErrHandler
        
        ' Only allow numeric keys
        If KeyCode >= vbKey0 And KeyCode <= vbKey9 Then
            
            ' Get Msg Number
            Dim NroMsg As Integer
            NroMsg = KeyCode - vbKey0 - 1
            
            ' Pressed "0", so Msg Number is 9
            If NroMsg = -1 Then NroMsg = 9
            
            'Como es KeyDown, si mantenes _
             apretado el mensaje llena la consola
            If CustomMessages.Message(NroMsg) = SendTxt.text Then
                Exit Sub
            End If
            
            CustomMessages.Message(NroMsg) = SendTxt.text
            
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡¡""" & SendTxt.Text & """ fue guardado como mensaje personalizado " & NroMsg + 1 & "!!", .red, .green, .blue, .bold, .italic)
            End With
            
        End If
        
    End If
    
    Exit Sub
    
ErrHandler:
    'Did detected an invalid message??
    If Err.Number = CustomMessages.InvalidMessageErrCode Then
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("El Mensaje es inválido. Modifiquelo por favor.", .red, .green, .blue, .bold, .italic)
        End With
    End If
    
End Sub

Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)
    'Send text
On Error GoTo ErrHandler
  
    If KeyCode = vbKeyReturn Or (KeyCode = CustomKeys.BindedKey(eKeyType.mKeyTalk)) Then
        If LenB(stxtbuffer) <> 0 Then Call ParseUserCommand(stxtbuffer)
        
        stxtbuffer = ""
        SendTxt.text = ""
        KeyCode = 0
        SendTxt.Visible = False
        
        If picInv.Visible Then
            picInv.SetFocus
        Else
            hlst.SetFocus
        End If
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SendTxt_KeyUp de frmMain.frm")
End Sub

Private Sub Second_Timer()
    If Not DialogosClanes Is Nothing Then DialogosClanes.PassTimer
  
    HeartBeatTime = HeartBeatTime + 1
    
    If HeartBeatTime >= 5 Then
        Call WritePing
        HeartBeatTime = 0
    End If
        
        Call modQuests.RefreshObjectives
End Sub

'[END]'

''''''''''''''''''''''''''''''''''''''
'     ITEM CONTROL                   '
''''''''''''''''''''''''''''''''''''''

Private Sub TirarItem()
On Error GoTo ErrHandler
    
    If Comerciando Or ViewingFormCantMove Then Exit Sub
    
    If UserEstado = 1 Then
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
        End With
    Else
        If (Inventario.SelectedItem > 0 And Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Or (Inventario.SelectedItem = FLAGORO) Then
            If Inventario.Amount(Inventario.SelectedItem) = 1 Then
                Call WriteDrop(Inventario.SelectedItem, 1)
            Else
                If Inventario.Amount(Inventario.SelectedItem) > 1 Then
                    If Not Comerciando Then frmCantidad.Show , frmMain
                End If
            End If
        End If
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub TirarItem de frmMain.frm")
End Sub

Private Sub AgarrarItem()
On Error GoTo ErrHandler
  
    If UserEstado = 1 Then
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
        End With
    Else
        Call WritePickUp
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub AgarrarItem de frmMain.frm")
End Sub

Private Sub UsarItem()
On Error GoTo ErrHandler
  
    If pausa Then Exit Sub
    
    If Comerciando Or ViewingFormCantMove Then Exit Sub
    
    If TrainingMacro.Enabled Then DesactivarMacroHechizos
    
    If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then _
        Call WriteUseItem(Inventario.SelectedItem)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub UsarItem de frmMain.frm")
End Sub

Private Sub EquiparItem()
On Error GoTo ErrHandler
  
    If UserEstado = 1 Then
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
        End With
    Else
        If Comerciando Or ViewingFormCantMove Then Exit Sub
        
        If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then _
        Call WriteEquipItem(Inventario.SelectedItem)
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub EquiparItem de frmMain.frm")
End Sub


Private Sub tmrBlink_Timer()
On Error GoTo ErrHandler
  
    If bLastBrightBlink Then
        frmMain.lblStrg.ForeColor = getStrenghtColor(15)
        frmMain.lblDext.ForeColor = getDexterityColor(15)
    Else
        frmMain.lblStrg.ForeColor = getStrenghtColor(UserFuerza)
        frmMain.lblDext.ForeColor = getDexterityColor(UserAgilidad)
    End If
    
    bLastBrightBlink = Not bLastBrightBlink
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub tmrBlink_Timer de frmMain.frm")
End Sub

''''''''''''''''''''''''''''''''''''''
'     HECHIZOS CONTROL               '
''''''''''''''''''''''''''''''''''''''

Private Sub TrainingMacro_Timer()
    If Not hlst.Visible Then
On Error GoTo ErrHandler
  
        DesactivarMacroHechizos
        Exit Sub
    End If
    
    'Macros are disabled if focus is not on Argentum!
    If Not Application.IsAppActive() Then
        DesactivarMacroHechizos
        Exit Sub
    End If
    
    If Comerciando Then Exit Sub
    
    If hlst.List(hlst.ListIndex) <> "(None)" And MainTimer.Check(TimersIndex.CastSpell, False) Then
        Call WriteCastSpell(hlst.ListIndex + 1)
        Call WriteWork(eSkill.Magia)
    End If
    
    Call ConvertCPtoTP(MouseX, MouseY, tX, tY)
    
    If UsingSkill = Magia And Not MainTimer.Check(TimersIndex.CastSpell) Then Exit Sub
    
    If UsingSkill = Proyectiles And Not MainTimer.Check(TimersIndex.Attack) Then Exit Sub
    
    Call WriteWorkLeftClick(tX, tY, UsingSkill)
    UsingSkill = 0
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub TrainingMacro_Timer de frmMain.frm")
End Sub

Private Sub cmdLanzar_Click()
    If hlst.List(hlst.ListIndex) <> "(None)" And MainTimer.Check(TimersIndex.Work, False) Then
On Error GoTo ErrHandler
  
        If UserEstado = 1 Then
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
            End With
        Else
            Call WriteCastSpell(hlst.ListIndex + 1)
            UsaMacro = True
        End If
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub cmdLanzar_Click de frmMain.frm")
End Sub

Private Sub CmdLanzar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UsaMacro = False
On Error GoTo ErrHandler
  
    CnTd = 0
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CmdLanzar_MouseMove de frmMain.frm")
End Sub

Private Sub cmdINFO_Click()
    If hlst.ListIndex <> -1 Then
On Error GoTo ErrHandler
  
        Call WriteSpellInfo(hlst.ListIndex + 1)
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub cmdINFO_Click de frmMain.frm")
End Sub

Private Sub picMain_Click()
If SendCMSTXT.Visible Then
On Error GoTo ErrHandler
  
SendCMSTXT.SetFocus
End If

If SendTxt.Visible Then
    SendTxt.SetFocus
End If
    If Cartel Then Cartel = False
    
#If EnableSecurity Then
    If LOGGING Then Call CheatingDeath.StoreKey(MouseBoton, True)
#End If

    If Not Comerciando Then

        Call ConvertCPtoTP(MouseX, MouseY, tX, tY)
        
        If MouseShift = 0 Then
            If MouseBoton <> vbRightButton Then
                '[ybarra]
                If UsaMacro Then
                    CnTd = CnTd + 1
                    If CnTd = 3 Then
                        Call WriteUseSpellMacro
                        CnTd = 0
                    End If
                    UsaMacro = False
                End If
                '[/ybarra]
                If UsingSkill = 0 Then
                    If MainTimer.Check(TimersIndex.Click) Then
                        Call WriteLeftClick(tX, tY)
                    End If
                Else
                
                    If TrainingMacro.Enabled Then Call DesactivarMacroHechizos
                    If macrotrabajo.Enabled Then Call DesactivarMacroTrabajo
   
                    'Splitted because VB isn't lazy!
                    If (Not esGM(UserCharIndex)) Then
                        If UsingSkill = Proyectiles Then
                            If Not MainTimer.Check(TimersIndex.Arrows) Then
                                Call modCustomCursors.SetFormCursorDefault(Me)
                                UsingSkill = 0
                                With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                    Call AddtoRichTextBox(frmMain.RecTxt(0), "No puedes lanzar proyectiles tan rápido.", .red, .green, .blue, .bold, .italic, , eMessageType.combate)
                                End With
                                Exit Sub
                            End If
                       ElseIf UsingSkill = Magia Then
    
                            If Not frmMain.CanUseSpell(CastedSpellIndex - 1) Then 'Check if spells interval has finished.
                                Call modCustomCursors.SetFormCursorDefault(Me)
                                UsingSkill = 0
                                With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                    Call AddtoRichTextBox(frmMain.RecTxt(0), "No puedes lanzar hechizos tan rapido.", .red, .green, .blue, .bold, .italic, , eMessageType.combate)
                                End With
                                Exit Sub
                            End If
                            
                        ElseIf (UsingSkill = Pesca Or UsingSkill = Robar Or UsingSkill = Talar Or UsingSkill = Mineria Or UsingSkill = FundirMetal) Then
                            If Not MainTimer.Check(TimersIndex.Work) Then
                                Call modCustomCursors.SetFormCursorDefault(Me)
                                UsingSkill = 0
                                Exit Sub
                            End If
                        End If
                    End If
                
                    Call modCustomCursors.SetFormCursorDefault(Me)
                    Call WriteWorkLeftClick(tX, tY, UsingSkill)
                    
                    UsingSkill = 0
                End If
            Else

                ' Descastea
                If UsingSkill = Magia Or UsingSkill = Proyectiles Then
                    Call modCustomCursors.SetFormCursorDefault(Me)
                    UsingSkill = 0
                ElseIf GameConfig.Extras.bRightClickEnabled Then
                    ' Store the place right clicked
                    LeftClicX = clicX
                    LeftClicY = clicY
                    
                    If MainTimer.Check(TimersIndex.Click) Then
                        Call WriteRightClick(tX, tY)
                    End If
                End If
   
            End If
        ElseIf (MouseShift And 1) = 1 Then
            If Not CustomKeys.KeyAssigned(KeyCodeConstants.vbKeyShift) Then
                If MouseBoton = vbLeftButton Then
                    Call WriteWarpChar("YO", PlayerData.CurrentMap.Number, tX, tY)
                End If
            End If
        End If
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub picMain_Click de frmMain.frm")
End Sub

Private Sub picMain_DblClick()
'**************************************************************
'Author: Unknown
'Last Modify Date: 12/27/2007
'12/28/2007: ByVal - Chequea que la ventana de comercio y boveda no este abierta al hacer doble clic a un comerciante, sobrecarga la lista de items.
'**************************************************************
On Error GoTo ErrHandler
  
    Call ConvertCPtoTP(MouseX, MouseY, tX, tY)
    If Not MirandoForo And Not Comerciando Then 'frmComerciar.Visible And Not frmBancoObj.Visible Then
        If MapData(tX, tY).ObjGrh.GrhIndex = 600 Then
            Call WriteWorkLeftClick(tX, tY, eSkill.Herreria)
        Else
            Call WriteDoubleClick(tX, tY)
        End If
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub picMain_DblClick de frmMain.frm")
End Sub

Private Sub picMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error GoTo ErrHandler
  
    MouseX = X
    MouseY = Y
    
    'Trim to fit screen
    If MouseX < 0 Then
        MouseX = 0
    ElseIf MouseX > picMain.Width Then
        MouseX = picMain.Width
    End If
    
    'Trim to fit screen
    If MouseY < 0 Then
        MouseY = 0
    ElseIf MouseY > picMain.Height Then
        MouseY = picMain.Height
    End If
    
    LastButtonPressed.ToggleToNormal
    
    RestoreTabs
    
    ' Disable links checking (not over consola)
    StopCheckingLinks

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub picMain_MouseMove de frmMain.frm")
End Sub

Private Sub RestoreTabs()
    ' Restore Tab
On Error GoTo ErrHandler
  
    If eeCurrentRolloverTab <> eConsoleType.Last Then
        With MainTabs(eeCurrentRolloverTab)
            If .NewMessages Then
                Set imgPestania(eeCurrentRolloverTab) = .picNewMessages
            Else
                Set imgPestania(eeCurrentRolloverTab) = .picNormal
            End If
        End With
        
        eeCurrentRolloverTab = eConsoleType.Last
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub RestoreTabs de frmMain.frm")
End Sub

Private Sub hlst_KeyDown(KeyCode As Integer, Shift As Integer)
       KeyCode = 0
  
End Sub

Private Sub hlst_KeyPress(KeyAscii As Integer)
       KeyAscii = 0

End Sub

Private Sub hlst_KeyUp(KeyCode As Integer, Shift As Integer)
        KeyCode = 0

End Sub

Private Sub lblDropGold_Click()

On Error GoTo ErrHandler
  
    Inventario.SelectGold
    If UserGLD > 0 Then
        If Not Comerciando Then frmCantidad.Show , frmMain
    End If
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub lblDropGold_Click de frmMain.frm")
End Sub

Private Sub Label4_Click()
    Call Engine_Audio.PlayInterface(SND_CLICK)

On Error GoTo ErrHandler
  
    Call ShowInventory
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Label4_Click de frmMain.frm")
End Sub

Public Sub ShowInventory()

    InvEqu.Picture = LoadPicture(DirInterfaces & SELECTED_UI & "CentroInventario.jpg")

    ' Activo controles de inventario
    picInv.Visible = True

    ' Desactivo controles de hechizo
    hlst.Visible = False
    cmdInfo.Visible = False
    CmdLanzar.Visible = False
    
    imgMoveMagicUp.Visible = False
    imgMoveMagicDown.Visible = False
    
End Sub

Private Sub Label7_Click()
    
    ' D'Artagnan: Only magic classes can open spells panel
On Error GoTo ErrHandler
  
    If PlayerData.Class = eClass.Warrior Or PlayerData.Class = eClass.Thief Or PlayerData.Class = eClass.Worker Or PlayerData.Class = eClass.Hunter Then Exit Sub
    
    Call Engine_Audio.PlayInterface(SND_CLICK)

    InvEqu.Picture = LoadPicture(DirInterfaces & SELECTED_UI & "Centrohechizos.jpg") 'LoadPicture(App.path & "\Graficos\Centrohechizos.jpg")
    
    ' Activo controles de hechizos
    hlst.Visible = True
    cmdInfo.Visible = True
    CmdLanzar.Visible = True
    
    imgMoveMagicUp.Visible = True
    imgMoveMagicDown.Visible = True
    
    ' Desactivo controles de inventario
    picInv.Visible = False

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Label7_Click de frmMain.frm")
End Sub

Private Sub picInv_DblClick()

On Error GoTo ErrHandler
  
    If MirandoCarpinteria Or MirandoHerreria Then Exit Sub
    
    If Not MainTimer.Check(TimersIndex.UseItemWithDblClick) Then Exit Sub
    
    If macrotrabajo.Enabled Then Call DesactivarMacroTrabajo
    
    Call UsarItem
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub picInv_DblClick de frmMain.frm")
End Sub
  
Private Sub picInv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrHandler
    Call Engine_Audio.PlayInterface(SND_CLICK)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub picInv_MouseUp de frmMain.frm")
End Sub

Private Sub SendTxt_Change()
'**************************************************************
'Author: Unknown
'Last Modify Date: 3/06/2006
'3/06/2006: Maraxus - impedí se inserten caractéres no imprimibles
'**************************************************************
On Error GoTo ErrHandler
  
    If Len(SendTxt.text) > 160 Then
        stxtbuffer = "Soy un cheater, avisenle a un gm"
    Else
        'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
        Dim I As Long
        Dim Temp As String
        Dim CharAscii As Integer
        
        For I = 1 To Len(SendTxt.text)
            CharAscii = Asc(mid$(SendTxt.text, I, 1))
            If CharAscii >= vbKeySpace And CharAscii <= 250 Then
                Temp = Temp & Chr$(CharAscii)
            End If
        Next I

        If Temp <> SendTxt.text Then
            'We only set it if it's different, otherwise the event will be raised
            'constantly and the client will crush
            SendTxt.text = Temp
        End If
        
        stxtbuffer = SendTxt.text
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SendTxt_Change de frmMain.frm")
End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
  
End Sub

Private Sub SendCMSTXT_KeyUp(KeyCode As Integer, Shift As Integer)
    'Send text
On Error GoTo ErrHandler
  
    If (KeyCode = vbKeyReturn) Or (KeyCode = CustomKeys.BindedKey(eKeyType.mKeyTalkWithGuild)) Then
        'Say
        If stxtbuffercmsg <> "" Then
            Call ParseUserCommand("/CMSG " & stxtbuffercmsg)
        End If

        stxtbuffercmsg = ""
        SendCMSTXT.text = ""
        KeyCode = 0
        Me.SendCMSTXT.Visible = False
        
        If picInv.Visible Then
            picInv.SetFocus
        Else
            hlst.SetFocus
        End If
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SendCMSTXT_KeyUp de frmMain.frm")
End Sub

Private Sub SendCMSTXT_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
  
End Sub

Private Sub SendCMSTXT_Change()
On Error GoTo ErrHandler
    If Len(SendCMSTXT.text) > 160 Then
        stxtbuffercmsg = "Soy un cheater, avisenle a un GM"
    Else
        'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
        Dim I As Long
        Dim Temp As String
        Dim CharAscii As Integer
        
        For I = 1 To Len(SendCMSTXT.text)
            CharAscii = Asc(mid$(SendCMSTXT.text, I, 1))
            If CharAscii >= vbKeySpace And CharAscii <= 250 Then
                Temp = Temp & Chr$(CharAscii)
            End If
        Next I

        If Temp <> SendCMSTXT.text Then
            'We only set it if it's different, otherwise the event will be raised
            'constantly and the client will crush
            SendCMSTXT.text = Temp
        End If
        
        stxtbuffercmsg = SendCMSTXT.text
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SendCMSTXT_Change de frmMain.frm")
End Sub


Public Sub NewMessageMainTab(ByVal eePestania As eConsoleType)
On Error GoTo ErrHandler
  
    
    If MainTabs(eePestania).NewMessages Or _
        MainTabs(eePestania).Selected Or _
        MainTabs(eePestania).Rollover Then Exit Sub
        
    Set imgPestania(eePestania).Picture = MainTabs(eePestania).picNewMessages
    MainTabs(eePestania).NewMessages = True
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub NewMessageMainTab de frmMain.frm")
End Sub

Private Sub imgPestania_Click(Index As Integer)
    
On Error GoTo ErrHandler
  
    If Index = eeCurrentSelectedTab Then Exit Sub

    ' Remove previous
    With MainTabs(eeCurrentSelectedTab)
        Set imgPestania(eeCurrentSelectedTab).Picture = .picNormal
        RecTxt(eeCurrentSelectedTab).Visible = False
        .Selected = False
    End With
    
    ' Update clicked one
    eeCurrentSelectedTab = Index
    With MainTabs(eeCurrentSelectedTab)
        Set imgPestania(eeCurrentSelectedTab).Picture = .picSelected
        RecTxt(eeCurrentSelectedTab).Visible = True
        .Selected = True
        .NewMessages = False
        .Rollover = False
    End With
    
    imgPestania(eeCurrentSelectedTab).ZOrder 0
    
    If Index = eeCurrentRolloverTab Then _
        eeCurrentRolloverTab = eConsoleType.Last
        
    Call Engine_Audio.PlayInterface(SND_CLICK)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgPestania_Click de frmMain.frm")
End Sub

Private Sub imgPestania_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ' No updates
On Error GoTo ErrHandler
  
    If Index = eeCurrentRolloverTab Then Exit Sub
    
    ' Last=None
    If eeCurrentRolloverTab <> eConsoleType.Last Then
        With MainTabs(eeCurrentRolloverTab)
            If .NewMessages Then
                Set imgPestania(eeCurrentRolloverTab) = .picNewMessages
            Else
                Set imgPestania(eeCurrentRolloverTab) = .picNormal
            End If
            
            .Rollover = False
        End With
    End If
    
    If Index <> eeCurrentSelectedTab Then
        eeCurrentRolloverTab = Index
        With MainTabs(Index)
            Set imgPestania(Index) = .picRollover
            .Rollover = True
        End With
    Else
        eeCurrentRolloverTab = eConsoleType.Last
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgPestania_MouseMove de frmMain.frm")
End Sub

Public Sub Disconnect()
On Error GoTo ErrHandler

    Call Protocol.Shutdown

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Disconnect de frmMain.frm")
End Sub

Public Sub UpdateStaBar()
On Error GoTo ErrHandler
  
    Dim bWidth As Byte
    
    If UserMaxSTA > 0 Then
        bWidth = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 88)
    End If
    
    shpEnergia.Width = 88 - bWidth
    shpEnergia.Left = 573 + (88 - shpEnergia.Width)
    shpEnergia.Visible = (bWidth <> 88)
      
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub UpdateStaBar de frmMain.frm")
End Sub

Public Sub UpdateManBar()
On Error GoTo ErrHandler
  
    Dim bWidth As Byte
    
    If UserMaxMAN > 0 Then
        bWidth = (((UserMinMAN / 100) / (UserMaxMAN / 100)) * 88)
    End If
        
    shpMana.Width = 88 - bWidth
    shpMana.Left = 573 + (88 - shpMana.Width)
    shpMana.Visible = (bWidth <> 88)
      
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub UpdateManBar de frmMain.frm")
End Sub

Public Sub UpdateHPBar()
On Error GoTo ErrHandler
  
    Dim bWidth As Byte
    
    If UserMaxHP > 0 Then
        bWidth = (((UserMinHP / 100) / (UserMaxHP / 100)) * 88)
    End If
    
    shpVida.Width = 88 - bWidth
    shpVida.Left = 573 + (88 - shpVida.Width)
    shpVida.Visible = (bWidth <> 88)
      
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub UpdateHPBar de frmMain.frm")
End Sub


Public Sub ShowStoreButton(ByVal ShowButton As Boolean)
    
    picStoreButton.AutoSize = True
    
    picStoreButton.Picture = IIf(PlayerData.OnDemandCraftingStoreOpen, picCloseStore, picOpenStore)
    
    picStoreButton.Visible = ShowButton
    picStoreButton.Left = frmMain.picMain.Left + (frmMain.picMain.Width / 2) - (picStoreButton.Width / 2)
    picStoreButton.Top = frmMain.picMain.Top + frmMain.picMain.Height - picStoreButton.Height + 20
    picStoreButton.MouseIcon = picMouseIcon
    picStoreButton.MousePointer = vbCustom
    
End Sub

Public Sub hlst_CooldownFinish(Index As Integer)

    Call RecalculateMousePointerForSpell
    
End Sub

Public Sub hlst_SpellAfterMeleeCooldownFinish()

    Call RecalculateMousePointerForSpell
    
End Sub

Public Sub RecalculateMousePointerForSpell()
    If Not UsingSkill = eSkill.Magia Then Exit Sub
    
    frmMain.MousePointer = MousePointerConstants.vbCustom
    frmMain.MouseIcon = GetMousePointerForAction(eMousePointerAction.Spell, IIf(CanUseSpell(CastedSpellIndex - 1), eMousePointerModifier.Normal, eMousePointerModifier.Disabled))
End Sub

Public Function CanUseSpell(ByVal SpellIndex As Integer) As Boolean
    Dim CurrentSpellIsReady As Boolean
    Dim DefaultSpellIsReady As Boolean
    Dim SpellAfterMeleeIsReady As Boolean
    
    CurrentSpellIsReady = True
    
    If Not UsingSkill = eSkill.Magia Then Exit Function
    
    SpellAfterMeleeIsReady = hlst.SpellAfterMeleeIsReady()
    CurrentSpellIsReady = IIf(SpellIndex >= 0, hlst.ItemIsReady(SpellIndex), False)
    DefaultSpellIsReady = IIf(SpellIndex <> charlist(UserCharIndex).LastSpellCast, hlst.DefaultIsReady(), True)

    CanUseSpell = (SpellAfterMeleeIsReady = True And DefaultSpellIsReady = True And CurrentSpellIsReady = True)
End Function


