VERSION 5.00
Begin VB.Form frmAccountCreate 
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtSecurityCode 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      IMEMode         =   3  'DISABLE
      Left            =   4125
      TabIndex        =   4
      ToolTipText     =   "Código de seguridad en caso de olvidar su contraseña."
      Top             =   6270
      Width           =   3795
   End
   Begin VB.TextBox txtEmail 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      IMEMode         =   3  'DISABLE
      Left            =   4125
      TabIndex        =   3
      Top             =   5430
      Width           =   3795
   End
   Begin VB.TextBox txtRePassword 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      IMEMode         =   3  'DISABLE
      Left            =   4125
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   4590
      Width           =   3795
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      IMEMode         =   3  'DISABLE
      Left            =   4125
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3750
      Width           =   3795
   End
   Begin VB.TextBox txtName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Left            =   4125
      TabIndex        =   0
      Top             =   2910
      Width           =   3795
   End
   Begin VB.Image imgCreateAccount 
      Height          =   1560
      Left            =   9210
      Top             =   6435
      Width           =   2250
   End
   Begin VB.Image imgClose 
      Height          =   420
      Left            =   4995
      Tag             =   "1"
      Top             =   8250
      Width           =   1650
   End
End
Attribute VB_Name = "frmAccountCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cCreateAccountButton As clsGraphicalButton
Private cCloseButton As clsGraphicalButton
Public LastButtonPressed As clsGraphicalButton

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then CloseWindow
End Sub

Private Sub ImgClose_Click()
    Call frmMain.Disconnect
On Error GoTo ErrHandler
  
    Unload Me
    frmConnect.SetFocus
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgClose_Click de frmAccountCreate.frm")
End Sub

Private Sub Form_Load()
    Me.Picture = LoadPicture(DirInterfaces & SELECTED_UI & "VentanaCrearCuenta.jpg")
    Call LoadButtons
    
    Call modCustomCursors.SetFormCursorDefault(Me)
End Sub

Private Sub LoadButtons()
On Error GoTo ErrHandler
  
    Dim GrhPath As String
    
    GrhPath = DirInterfaces & SELECTED_UI
    
    Set cCreateAccountButton = New clsGraphicalButton
    Set cCloseButton = New clsGraphicalButton
    Set LastButtonPressed = New clsGraphicalButton
    
    Call cCreateAccountButton.Initialize(Me.imgCreateAccount, GrhPath & "BotonCrearCuenta.gif", _
                                         GrhPath & "BotonCrearCuentaRollover.gif", _
                                         GrhPath & "BotonCrearCuentaClick.gif", Me)
                                         
    Call cCloseButton.Initialize(Me.imgClose, GrhPath & "RegresarRollover.gif", _
                                 GrhPath & "RegresarRollover.gif", _
                                 GrhPath & "RegresarRollover.gif", Me)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadButtons de frmAccountCreate.frm")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub imgCreateAccount_Click()
On Error GoTo ErrHandler
  
    If ValidInput(txtName.text, txtPassword.text, txtRePassword.text, _
                  txtEmail.text, txtSecurityCode.text) Then
        If txtPassword.text = txtRePassword.text Then
            If CheckMailString(txtEmail.text) Then
                If Not txtSecurityCode.text = txtName.text Then
                    Call modAccount.Set_Acc_Data_To_Create
                    Call modAccount.Prepare_And_Connect(E_MODO.AccountCreate)
                Else
                    Call MsgBox("El token no puede ser igual al nombre de la cuenta")
                End If
            Else
                Call MsgBox("Ingrese una dirección de correo electrónico válida.")
            End If
        Else
            Call MsgBox("Las contraseñas no coinciden.")
        End If
    Else
        Call MsgBox("Complete todos los campos.")
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgCreateAccount_Click de frmAccountCreate.frm")
End Sub

Private Sub CloseWindow()
On Error GoTo ErrHandler
  
    Unload Me
    frmConnect.SetFocus
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CloseWindow de frmAccountCreate.frm")
End Sub

