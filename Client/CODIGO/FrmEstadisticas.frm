VERSION 5.00
Begin VB.Form frmEstadisticas 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Estadisticas"
   ClientHeight    =   6675
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   6975
   ClipControls    =   0   'False
   Icon            =   "FrmEstadisticas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   445
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   18
      Left            =   5415
      Top             =   5400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Skills 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Height          =   195
      Index           =   18
      Left            =   4965
      TabIndex        =   34
      Top             =   5385
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgCerrar 
      Height          =   195
      Left            =   6630
      Tag             =   "1"
      Top             =   150
      Width           =   195
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   17
      Left            =   5415
      Top             =   5145
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   16
      Left            =   5415
      Top             =   4890
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   15
      Left            =   5415
      Top             =   4635
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   14
      Left            =   5415
      Top             =   4380
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   13
      Left            =   5415
      Top             =   4125
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   12
      Left            =   5415
      Top             =   3870
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   11
      Left            =   5415
      Top             =   3615
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   10
      Left            =   5415
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   9
      Left            =   5415
      Top             =   3105
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   8
      Left            =   5415
      Top             =   2850
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   7
      Left            =   5415
      Top             =   2595
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   6
      Left            =   5415
      Top             =   2340
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   5
      Left            =   5415
      Top             =   2085
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   4
      Left            =   5415
      Top             =   1830
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   3
      Left            =   5415
      Top             =   1575
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   2
      Left            =   5415
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   1
      Left            =   5415
      Top             =   1065
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Left            =   2130
      TabIndex        =   33
      Top             =   6060
      Width           =   555
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Clase"
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
      Left            =   1800
      TabIndex        =   32
      Top             =   5835
      Width           =   960
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Left            =   2130
      TabIndex        =   31
      Top             =   5610
      Width           =   825
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Left            =   2130
      TabIndex        =   30
      Top             =   5385
      Width           =   825
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Left            =   2130
      TabIndex        =   29
      Top             =   5160
      Width           =   825
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Index           =   0
      Left            =   2130
      TabIndex        =   28
      Top             =   4935
      Width           =   825
   End
   Begin VB.Label Skills 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Height          =   195
      Index           =   17
      Left            =   4965
      TabIndex        =   27
      Top             =   5130
      Width           =   315
   End
   Begin VB.Label Skills 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Height          =   195
      Index           =   16
      Left            =   4965
      TabIndex        =   26
      Top             =   4875
      Width           =   315
   End
   Begin VB.Label Skills 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Height          =   195
      Index           =   15
      Left            =   4965
      TabIndex        =   25
      Top             =   4620
      Width           =   315
   End
   Begin VB.Label Skills 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Height          =   195
      Index           =   14
      Left            =   4965
      TabIndex        =   24
      Top             =   4365
      Width           =   315
   End
   Begin VB.Label Skills 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Height          =   195
      Index           =   13
      Left            =   4965
      TabIndex        =   23
      Top             =   4110
      Width           =   315
   End
   Begin VB.Label Skills 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Height          =   195
      Index           =   12
      Left            =   4965
      TabIndex        =   22
      Top             =   3855
      Width           =   315
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ciudadano"
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
      Height          =   195
      Index           =   7
      Left            =   1560
      TabIndex        =   21
      Top             =   4080
      Width           =   1245
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Height          =   195
      Index           =   6
      Left            =   1785
      TabIndex        =   20
      Top             =   3840
      Width           =   315
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Height          =   195
      Index           =   5
      Left            =   1785
      TabIndex        =   19
      Top             =   3615
      Width           =   315
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Height          =   195
      Index           =   4
      Left            =   1785
      TabIndex        =   18
      Top             =   3390
      Width           =   315
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Height          =   195
      Index           =   2
      Left            =   1785
      TabIndex        =   17
      Top             =   3180
      Width           =   315
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Height          =   195
      Index           =   1
      Left            =   1785
      TabIndex        =   16
      Top             =   2970
      Width           =   315
   End
   Begin VB.Label Skills 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Height          =   195
      Index           =   11
      Left            =   4965
      TabIndex        =   15
      Top             =   3600
      Width           =   315
   End
   Begin VB.Label Skills 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Height          =   195
      Index           =   10
      Left            =   4965
      TabIndex        =   14
      Top             =   3345
      Width           =   315
   End
   Begin VB.Label Skills 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Height          =   195
      Index           =   9
      Left            =   4965
      TabIndex        =   13
      Top             =   3090
      Width           =   315
   End
   Begin VB.Label Skills 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Height          =   195
      Index           =   8
      Left            =   4965
      TabIndex        =   12
      Top             =   2835
      Width           =   315
   End
   Begin VB.Label Skills 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Height          =   195
      Index           =   7
      Left            =   4965
      TabIndex        =   11
      Top             =   2580
      Width           =   315
   End
   Begin VB.Label Skills 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Height          =   195
      Index           =   6
      Left            =   4965
      TabIndex        =   10
      Top             =   2325
      Width           =   315
   End
   Begin VB.Label Skills 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Height          =   195
      Index           =   5
      Left            =   4965
      TabIndex        =   9
      Top             =   2070
      Width           =   315
   End
   Begin VB.Label Skills 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Height          =   195
      Index           =   4
      Left            =   4965
      TabIndex        =   8
      Top             =   1815
      Width           =   315
   End
   Begin VB.Label Skills 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Height          =   195
      Index           =   3
      Left            =   4965
      TabIndex        =   7
      Top             =   1560
      Width           =   315
   End
   Begin VB.Label Skills 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Height          =   195
      Index           =   2
      Left            =   4965
      TabIndex        =   6
      Top             =   1305
      Width           =   315
   End
   Begin VB.Label Skills 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Height          =   195
      Index           =   1
      Left            =   4965
      TabIndex        =   5
      Top             =   1050
      Width           =   315
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
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
      Height          =   195
      Index           =   5
      Left            =   1785
      TabIndex        =   4
      Top             =   2040
      Width           =   210
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
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
      Height          =   195
      Index           =   4
      Left            =   1785
      TabIndex        =   3
      Top             =   1800
      Width           =   210
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
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
      Height          =   195
      Index           =   3
      Left            =   1785
      TabIndex        =   2
      Top             =   1560
      Width           =   210
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
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
      Height          =   195
      Index           =   2
      Left            =   1785
      TabIndex        =   1
      Top             =   1320
      Width           =   210
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
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
      Height          =   195
      Index           =   1
      Left            =   1785
      TabIndex        =   0
      Top             =   1080
      Width           =   210
   End
End
Attribute VB_Name = "frmEstadisticas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.6
'
'Copyright (C) 2002 MÃ¡rquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 MatÃ­as Fernando PequeÃ±o
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
'Calle 3 nÃºmero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'CÃ³digo Postal 1900
'Pablo Ignacio MÃ¡rquez

Option Explicit

Private clsFormulario As clsFormMovementManager

Private cBotonCerrar As clsGraphicalButton
Public LastButtonPressed As clsGraphicalButton

Private Const ANCHO_BARRA As Byte = 73 'pixeles
Private Const BAR_LEFT_POS As Integer = 361 'pixeles

Private Sub Iniciar_Labels()
'Iniciamos los labels con los valores de los atributos y los skills
On Error GoTo ErrHandler
  
Dim I As Integer
Dim Ancho As Integer

For I = 1 To NUMATRIBUTOS
    Atri(I).Caption = UserAtributos(I)
Next

For I = 1 To NUMSKILLS

    With UserSkills(I)
        Skills(I).Caption = .Both
        Skills(I).ToolTipText = .Natural & " naturales, " & .Assigned & " asignados."

        If .Both < MAXSKILLPOINTS Then
            Ancho = IIf(PorcentajeSkills(I) = 0, ANCHO_BARRA, (100 - PorcentajeSkills(I)) / 100 * ANCHO_BARRA)
        Else
            Ancho = 0
        End If

    End With

    shpSkillsBar(I).Width = Ancho
    shpSkillsBar(I).Left = BAR_LEFT_POS + ANCHO_BARRA - Ancho
Next

With charlist(UserCharIndex)
    Label4(7).ForeColor = GetNickColorByAlignment(.Alignment)
    Select Case .Alignment
        Case eCharacterAlignment.FactionLegion
            Label4(7).Caption = "Legion Oscura"
        Case eCharacterAlignment.FactionRoyal
            Label4(7).Caption = "Armada Real"
        Case eCharacterAlignment.Neutral
            Label4(7).Caption = "Neutral"
        Case Else
            Label4(7).Caption = "Newbie"
    End Select
End With


With UserEstadisticas
    Label6(0).Caption = .CriminalesMatados
    Label6(1).Caption = .CiudadanosMatados
    Label6(2).Caption = .UsuariosMatados
    Label6(3).Caption = .NpcsMatados
    Label6(4).Caption = .Clase
    Label6(5).Caption = .PenaCarcel
End With

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Iniciar_Labels de FrmEstadisticas.frm")
End Sub

Private Sub Form_Load()
    
On Error GoTo ErrHandler
  
    MirandoEstadisticas = True
    
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Iniciar_Labels
    
    Me.Picture = LoadPicture(DirInterfaces & SELECTED_UI & "VentanaEstadisticas.jpg")
    
    Call LoadButtons
    
    Call modCustomCursors.SetFormCursorDefault(Me)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Form_Load de FrmEstadisticas.frm")
End Sub

Private Sub LoadButtons()
On Error GoTo ErrHandler
  
    
    Dim GrhPath As String
    
    GrhPath = DirInterfaces & SELECTED_UI
    
    Set cBotonCerrar = New clsGraphicalButton
    Set LastButtonPressed = New clsGraphicalButton
    
    Call cBotonCerrar.Initialize(imgCerrar, GrhPath & "BotonEstadisticasSalir.jpg", _
                                    GrhPath & "BotonEstadisticasSalirRollover.jpg", _
                                    GrhPath & "BotonEstadisticasSalirClick.jpg", Me)

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadButtons de FrmEstadisticas.frm")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    CerrarVentana
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MirandoEstadisticas = False
End Sub

Private Sub imgCerrar_Click()
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
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CerrarVentana de FrmEstadisticas.frm")
End Sub
