VERSION 5.00
Begin VB.Form frmDuelo4v4 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4830
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7185
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Nick7 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   285
      Left            =   1440
      MaxLength       =   35
      TabIndex        =   7
      Top             =   5600
      Width           =   2775
   End
   Begin VB.TextBox Nick6 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   285
      Left            =   1440
      MaxLength       =   35
      TabIndex        =   6
      Top             =   5040
      Width           =   2775
   End
   Begin VB.TextBox Nick3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   285
      Left            =   1440
      MaxLength       =   35
      TabIndex        =   3
      Top             =   3270
      Width           =   2775
   End
   Begin VB.TextBox Nick2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   285
      Left            =   1440
      MaxLength       =   35
      TabIndex        =   2
      Top             =   2750
      Width           =   2775
   End
   Begin VB.TextBox Nick1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   285
      Left            =   1440
      MaxLength       =   35
      TabIndex        =   1
      Top             =   2200
      Width           =   2775
   End
   Begin VB.TextBox Oro 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   1440
      MaxLength       =   7
      TabIndex        =   0
      Text            =   "0"
      Top             =   900
      Width           =   2775
   End
   Begin VB.TextBox Nick4 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   285
      Left            =   1440
      MaxLength       =   35
      TabIndex        =   4
      Top             =   3960
      Width           =   2775
   End
   Begin VB.TextBox Nick5 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   285
      Left            =   1440
      MaxLength       =   35
      TabIndex        =   5
      Top             =   4520
      Width           =   2775
   End
   Begin VB.Image Resu 
      Height          =   420
      Left            =   3720
      Top             =   1500
      Width           =   420
   End
   Begin VB.Image Drop 
      Height          =   420
      Left            =   1440
      Top             =   1500
      Width           =   420
   End
   Begin VB.Image imgRetar 
      Height          =   885
      Left            =   550
      Top             =   6050
      Width           =   3795
   End
   Begin VB.Image imgCerrar 
      Height          =   195
      Left            =   4500
      Top             =   150
      Width           =   195
   End
End
Attribute VB_Name = "frmDuelo4v4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private clsFormulario As clsFormMovementManager
Private cBotonCerrar As clsGraphicalButton
Private cBotonRetar As clsGraphicalButton
Private Dropi As Boolean
Private Resui As Boolean
Private Tic As Picture
Private Tac As Picture
Public LastButtonPressed As clsGraphicalButton

Private Sub Drop_Click()
Dropi = Not Dropi
On Error GoTo ErrHandler
  
If Dropi Then
    Drop.Picture = Tac
Else
    Drop.Picture = Tic
End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Drop_Click de frmDuelo4v4.frm")
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then CloseWindow
End Sub

Private Sub Oro_Change()
On Error GoTo ErrHandler
    If Val(Oro.text) < 0 Then
        Oro.text = "1"
    End If
    
    If Val(Oro.text) > 9000000 Then
        Oro.text = "9000000"
    End If
    
    Exit Sub
    
ErrHandler:
    'If we got here the user may have pasted (Shift + Insert) a REALLY large number, causing an overflow, so we set amount back to 1
    Oro.text = "1"
End Sub

Private Sub Resu_Click()
Resui = Not Resui
On Error GoTo ErrHandler
  
If Resui Then
    Resu.Picture = Tac
Else
    Resu.Picture = Tic
End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Resu_Click de frmDuelo4v4.frm")
End Sub

Private Sub Form_Load()
    Set clsFormulario = New clsFormMovementManager
On Error GoTo ErrHandler
  
    clsFormulario.Initialize Me
    Call ActivarBotones
    Me.Picture = LoadPicture(DirInterfaces & SELECTED_UI & "VentanaDuelo4v4.jpg")
    Set Tic = LoadPicture(DirInterfaces & SELECTED_UI & "BotonDueloAmigosTic.jpg")
    Set Tac = LoadPicture(DirInterfaces & SELECTED_UI & "BotonDueloAmigosTac.jpg")
    Drop.Picture = Tic
    Dropi = False
    Resu.Picture = Tac
    Resui = True
    
    Call modCustomCursors.SetFormCursorDefault(Me)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Form_Load de frmDuelo4v4.frm")
End Sub

Sub ActivarBotones()
On Error GoTo ErrHandler
  
    Dim GrhPath As String
    
    GrhPath = DirInterfaces & SELECTED_UI

    Set cBotonCerrar = New clsGraphicalButton
    Set cBotonRetar = New clsGraphicalButton
    
    Set LastButtonPressed = New clsGraphicalButton

    Call cBotonCerrar.Initialize(imgCerrar, GrhPath & "BotonCruzSalir.jpg", _
                                    GrhPath & "BotonCruzSalirRollover.jpg", _
                                    GrhPath & "BotonCruzSalirClick.jpg", Me)
                                    
    Call cBotonRetar.Initialize(imgRetar, GrhPath & "BotonEnviarDuelo.jpg", _
                                    GrhPath & "BotonEnviarDueloRollover.jpg", _
                                    GrhPath & "BotonEnviarDueloClick.jpg", Me)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ActivarBotones de frmDuelo4v4.frm")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LastButtonPressed.ToggleToNormal
End Sub

Private Sub imgCerrar_Click()
Me.Visible = False
End Sub

Private Sub imgRetar_Click()
On Error GoTo ErrHandler
  
    Nick1.text = RTrim$(LTrim$(Nick1.text))
    Nick2.text = RTrim$(LTrim$(Nick2.text))
    Nick3.text = RTrim$(LTrim$(Nick3.text))
    Nick4.text = RTrim$(LTrim$(Nick4.text))
    Nick5.text = RTrim$(LTrim$(Nick5.text))
    Nick6.text = RTrim$(LTrim$(Nick6.text))
    Nick7.text = RTrim$(LTrim$(Nick7.text))
    If Not Len(Nick1.text) >= 1 Or Not Len(Nick2.text) >= 1 Or _
        Not Len(Nick3.text) >= 1 Or Not Len(Nick4.text) >= 1 Or _
        Not Len(Nick5.text) >= 1 Or Not Len(Nick6.text) >= 1 Or _
        Not Len(Nick7.text) >= 1 Then Exit Sub
    Call WriteRetar(4, Val(Oro.text), Dropi, Nick1.text, Nick2.text, Nick3.text, Resui, Nick4.text, Nick5.text, Nick6.text, Nick7.text)
    Me.Visible = False
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgRetar_Click de frmDuelo4v4.frm")
End Sub
Private Sub CloseWindow()
On Error GoTo ErrHandler
  
    Unload Me
    If frmDuelos.Visible Then
        frmDuelos.SetFocus
    Else
        If frmMain.Visible Then
            frmMain.SetFocus
        End If
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CloseWindow de frmDuelo4v4.frm")
End Sub



