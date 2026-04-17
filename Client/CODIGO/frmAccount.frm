VERSION 5.00
Begin VB.Form frmAccount 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Argentum Online"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmAccount.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.PictureBox picPj 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      DrawStyle       =   5  'Transparent
      HasDC           =   0   'False
      Height          =   1260
      Index           =   7
      Left            =   9285
      ScaleHeight     =   84
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   84
      TabIndex        =   7
      Top             =   6360
      Width           =   1260
   End
   Begin VB.PictureBox picPj 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      DrawStyle       =   5  'Transparent
      HasDC           =   0   'False
      Height          =   1260
      Index           =   6
      Left            =   9285
      ScaleHeight     =   84
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   84
      TabIndex        =   6
      Top             =   4815
      Width           =   1260
   End
   Begin VB.PictureBox picPj 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      DrawStyle       =   5  'Transparent
      HasDC           =   0   'False
      Height          =   1260
      Index           =   5
      Left            =   9285
      ScaleHeight     =   84
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   84
      TabIndex        =   5
      Top             =   3270
      Width           =   1260
   End
   Begin VB.PictureBox picPj 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      DrawStyle       =   5  'Transparent
      HasDC           =   0   'False
      Height          =   1260
      Index           =   4
      Left            =   9285
      ScaleHeight     =   84
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   84
      TabIndex        =   4
      Top             =   1725
      Width           =   1260
   End
   Begin VB.PictureBox picPj 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      DrawStyle       =   5  'Transparent
      HasDC           =   0   'False
      Height          =   1260
      Index           =   3
      Left            =   1440
      ScaleHeight     =   84
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   84
      TabIndex        =   3
      Top             =   6360
      Width           =   1260
   End
   Begin VB.PictureBox picPj 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      DrawStyle       =   5  'Transparent
      HasDC           =   0   'False
      Height          =   1260
      Index           =   2
      Left            =   1440
      ScaleHeight     =   84
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   84
      TabIndex        =   2
      Top             =   4800
      Width           =   1260
   End
   Begin VB.PictureBox picPj 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      DrawStyle       =   5  'Transparent
      HasDC           =   0   'False
      Height          =   1260
      Index           =   1
      Left            =   1440
      ScaleHeight     =   84
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   84
      TabIndex        =   1
      Top             =   3270
      Width           =   1260
   End
   Begin VB.PictureBox picPj 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      DrawStyle       =   5  'Transparent
      ForeColor       =   &H00000000&
      HasDC           =   0   'False
      Height          =   1260
      Index           =   0
      Left            =   1440
      ScaleHeight     =   84
      ScaleLeft       =   35
      ScaleMode       =   0  'User
      ScaleWidth      =   84
      TabIndex        =   0
      Top             =   1725
      Width           =   1260
   End
   Begin VB.Image imgWebsiteLink 
      Height          =   615
      Left            =   3960
      Top             =   8160
      Width           =   4455
   End
   Begin VB.Label lblGuildName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "<Guild>"
      ForeColor       =   &H00C0C0C0&
      Height          =   300
      Index           =   7
      Left            =   7140
      TabIndex        =   23
      Top             =   7080
      Width           =   1800
   End
   Begin VB.Label lblGuildName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "<Guild>"
      ForeColor       =   &H00C0C0C0&
      Height          =   300
      Index           =   6
      Left            =   7140
      TabIndex        =   22
      Top             =   5520
      Width           =   1800
   End
   Begin VB.Label lblGuildName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "<Guild>"
      ForeColor       =   &H00C0C0C0&
      Height          =   300
      Index           =   5
      Left            =   7140
      TabIndex        =   21
      Top             =   3960
      Width           =   1800
   End
   Begin VB.Label lblGuildName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "<Guild>"
      ForeColor       =   &H00C0C0C0&
      Height          =   300
      Index           =   4
      Left            =   7140
      TabIndex        =   20
      Top             =   2400
      Width           =   1800
   End
   Begin VB.Label lblGuildName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "<Guild>"
      ForeColor       =   &H00C0C0C0&
      Height          =   300
      Index           =   3
      Left            =   3030
      TabIndex        =   19
      Top             =   7080
      Width           =   1800
   End
   Begin VB.Label lblGuildName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "<Guild>"
      ForeColor       =   &H00C0C0C0&
      Height          =   300
      Index           =   2
      Left            =   3030
      TabIndex        =   18
      Top             =   5520
      Width           =   1800
   End
   Begin VB.Label lblGuildName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "<Guild>"
      ForeColor       =   &H00C0C0C0&
      Height          =   300
      Index           =   1
      Left            =   3030
      TabIndex        =   17
      Top             =   3960
      Width           =   1800
   End
   Begin VB.Label lblGuildName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "<Guild>"
      ForeColor       =   &H00C0C0C0&
      Height          =   300
      Index           =   0
      Left            =   3030
      TabIndex        =   16
      Top             =   2400
      Width           =   1800
   End
   Begin VB.Image imgDisconnect 
      Height          =   570
      Left            =   5475
      Top             =   7200
      Width           =   1260
   End
   Begin VB.Image imgCambiarContrasena 
      Height          =   555
      Left            =   240
      Top             =   8205
      Width           =   630
   End
   Begin VB.Image imgDiscord 
      Height          =   555
      Left            =   11160
      Top             =   8205
      Width           =   630
   End
   Begin VB.Image imgCrearPersonaje 
      Height          =   570
      Left            =   5475
      Top             =   4950
      Width           =   1260
   End
   Begin VB.Image imgBorrarPersonaje 
      Height          =   570
      Left            =   5475
      Top             =   5655
      Width           =   1260
   End
   Begin VB.Image imgEntrarPersonaje 
      Height          =   570
      Left            =   5475
      Top             =   3960
      Width           =   1260
   End
   Begin VB.Label lblInfoChar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nick"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   300
      Index           =   7
      Left            =   7140
      TabIndex        =   15
      Top             =   6840
      Width           =   1800
   End
   Begin VB.Label lblInfoChar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nick"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   300
      Index           =   6
      Left            =   7140
      TabIndex        =   14
      Top             =   5280
      Width           =   1800
   End
   Begin VB.Label lblInfoChar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nick"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   300
      Index           =   5
      Left            =   7140
      TabIndex        =   13
      Top             =   3720
      Width           =   1800
   End
   Begin VB.Label lblInfoChar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nick"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   300
      Index           =   4
      Left            =   7140
      TabIndex        =   12
      Top             =   2160
      Width           =   1800
   End
   Begin VB.Label lblInfoChar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nick"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   300
      Index           =   3
      Left            =   3030
      TabIndex        =   11
      Top             =   6840
      Width           =   1800
   End
   Begin VB.Label lblInfoChar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nick"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   300
      Index           =   2
      Left            =   3030
      TabIndex        =   10
      Top             =   5280
      Width           =   1800
   End
   Begin VB.Label lblInfoChar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nick"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   180
      Index           =   1
      Left            =   3030
      TabIndex        =   9
      Top             =   3720
      Width           =   1800
   End
   Begin VB.Label lblInfoChar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nick"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   180
      Index           =   0
      Left            =   3030
      TabIndex        =   8
      Top             =   2160
      Width           =   1800
   End
End
Attribute VB_Name = "frmAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cBotonEntrarPersonaje    As clsGraphicalButton
Private cBotonCambiarContrasenia As clsGraphicalButton
Private cBotonCrearPersonaje     As clsGraphicalButton
Private cBotonBorrarPersonaje    As clsGraphicalButton
Private cButtonDisconnect        As clsGraphicalButton
Private cButtonDiscord           As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrHandler
  
    If KeyCode = vbKeyEscape Then
        Call CloseFormAndBack
    End If
  
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Form_KeyUp de frmAccount.frm")
End Sub

Private Sub CloseFormAndBack()
    Me.Visible = False
    Call AccountReset
    frmConnect.Visible = True
    Call frmConnect.txtNombre.SetFocus
End Sub


Private Sub Form_Load()
On Error GoTo ErrHandler

    Call modCustomCursors.SetFormCursorDefault(Me)

    Me.Picture = LoadPicture(DirInterfaces & SELECTED_UI & "VentanaCuenta.jpg")

    Set cBotonEntrarPersonaje = New clsGraphicalButton
    Set cBotonCambiarContrasenia = New clsGraphicalButton
    Set cBotonCrearPersonaje = New clsGraphicalButton
    Set cBotonBorrarPersonaje = New clsGraphicalButton
    Set cButtonDisconnect = New clsGraphicalButton
    Set cButtonDiscord = New clsGraphicalButton
    
    Set LastButtonPressed = New clsGraphicalButton
    
                                    
    Call cBotonEntrarPersonaje.Initialize(imgEntrarPersonaje, DirInterfaces & SELECTED_UI & "BotonIngresar.jpg", _
                                    DirInterfaces & SELECTED_UI & "BotonIngresar.jpg", _
                                    DirInterfaces & SELECTED_UI & "BotonIngresar.jpg", Me)
                                    
    Call cBotonCambiarContrasenia.Initialize(imgCambiarContrasena, DirInterfaces & SELECTED_UI & "BotonLlaveAmarilla.jpg", _
                                    DirInterfaces & SELECTED_UI & "BotonLlaveAmarilla.jpg", _
                                    DirInterfaces & SELECTED_UI & "BotonLlaveAmarilla.jpg", Me)
                                    
    Call cBotonCrearPersonaje.Initialize(imgCrearPersonaje, DirInterfaces & SELECTED_UI & "BotonCrear.jpg", _
                                    DirInterfaces & SELECTED_UI & "BotonCrear.jpg", _
                                    DirInterfaces & SELECTED_UI & "BotonCrear.jpg", Me)
    
    Call cBotonBorrarPersonaje.Initialize(imgBorrarPersonaje, DirInterfaces & SELECTED_UI & "BotonBorrar.jpg", _
                                    DirInterfaces & SELECTED_UI & "BotonBorrar.jpg", _
                                    DirInterfaces & SELECTED_UI & "BotonBorrar.jpg", Me)
                                    
    Call cButtonDisconnect.Initialize(imgDisconnect, DirInterfaces & SELECTED_UI & "BotonDesconectar.jpg", _
                                    DirInterfaces & SELECTED_UI & "BotonDesconectar.jpg", _
                                    DirInterfaces & SELECTED_UI & "BotonDesconectar.jpg", Me)
                                    
    Call cButtonDiscord.Initialize(imgDiscord, DirInterfaces & SELECTED_UI & "BotonDiscord.jpg", _
                                        DirInterfaces & SELECTED_UI & "BotonDiscord.jpg", _
                                        DirInterfaces & SELECTED_UI & "BotonDiscord.jpg", Me)
            
    
    Dim I As Integer

    For I = LBound(modAccount.mDevice) To UBound(modAccount.mDevice)
        modAccount.mDevice(I) = Aurora_Graphics.CreatePassFromDisplay(picPj(I).hwnd, picPj(I).ScaleWidth, picPj(I).ScaleHeight)
        
        lblGuildName(I).Caption = vbNullString
        lblInfoChar(I).Caption = vbNullString
    Next I
    
    Call ResetAllInfo(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Form_Load de frmAccount.frm")
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim I As Integer

    For I = LBound(modAccount.mDevice) To UBound(modAccount.mDevice)
        Call Aurora_Graphics.DeletePass(modAccount.mDevice(I))
    Next I

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
  
End Sub

Public Sub imgBorrarPersonaje_Click()
On Error GoTo ErrHandler
  
    If Acc_Data.Acc_Char_Selected > 0 And MainTimer.Check(TimersIndex.Action) Then
        Call frmDeleteCharValidation.Show(, Me)
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgBorrarPersonaje_Click de frmAccount.frm")
End Sub

Private Sub imgCambiarContrasena_Click()
    Call frmAccountChangePassword.Show(, Me)
End Sub

Private Sub imgCrearPersonaje_Click()
On Error GoTo ErrHandler
  
    If Not MainTimer.Check(TimersIndex.Action) Then Exit Sub
    
    Dim I As Long
    Dim bAvailableSlot As Boolean
    
    ' Look for a free slot. However, this is checked by the server.
    For I = 1 To UBound(Acc_Data.Acc_Char)
        If LenB(Acc_Data.Acc_Char(I).Char_Name) = 0 Then
            bAvailableSlot = True
            Exit For
        End If
    Next I
    
    If bAvailableSlot Then
        'Acc_Data.Acc_Waiting_Response = True
        EstadoLogin = E_MODO.Dados
        Call HandleLogin
    Else
        Call MsgBox("Has alcanzado el número máximo de personajes.", vbInformation)
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgCrearPersonaje_Click de frmAccount.frm")
End Sub

Private Sub imgDisconnect_Click()
    Call CloseFormAndBack
End Sub

Private Sub imgDiscord_Click()
    Call Mod_General.OpenDiscordLink
End Sub

Private Sub imgEntrarPersonaje_Click()
On Error GoTo ErrHandler
        Call Entrar
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgEntrarPersonaje_Click de frmAccount.frm")
End Sub


Private Sub picPj_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrHandler

    Dim Last As Integer
    
    If Acc_Data.Acc_Char_Selected <> Index + 1 And LenB(Acc_Data.Acc_Char(Index + 1).Char_Name) Then
        Last = Acc_Data.Acc_Char_Selected
        Acc_Data.Acc_Char_Selected = Index + 1
        
        If (Last <> 0) Then
            Call Invalidate(picPj(Last - 1).hwnd)
        End If

        Call Invalidate(picPj(Index).hwnd)

    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub picPj_MouseUp de frmAccount.frm")
End Sub

Private Sub Entrar()
' G Toyz: 21/04

On Error GoTo ErrHandler

    If Acc_Data.Acc_Char_Selected > 0 And MainTimer.Check(TimersIndex.Action) Then
        ' Get user name
        UserName = Acc_Data.Acc_Char(Acc_Data.Acc_Char_Selected).Char_Name
        uName = complexNameToSimple(UserName, False)
         
        If UserName <> "" Then
            Call modAccount.Prepare_And_Connect(E_MODO.AccountLoginChar)
        End If
    End If

Exit Sub

ErrHandler:
        Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Entrar de frmAccount.frm")
End Sub

Private Sub picPJ_dblclick(Index As Integer)

' G Toyz - 21/04

On Error GoTo ErrHandler

    Call Entrar
    
Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub picPJ_dblclick de frmAccount.frm")
End Sub



Private Sub picPj_Paint(Index As Integer)
    
    Dim color As Long
    
    With Acc_Data.Acc_Char(Index + 1)
            
        Call UIBegin(mDevice(Index), frmAccount.picPJ(Index).ScaleWidth, frmAccount.picPJ(Index).ScaleHeight, &H0)

        If .Char_Name <> vbNullString Then
                          
            If (Acc_Data.Acc_Char_Selected = Index + 1) Then
                color = &HFFFFFFFF
            Else
                color = &HFF909090
            End If
                       
            Call Draw_Char_Slot(Index + 1, color)
        End If
                       
       Call UIEnd
    End With
End Sub
