VERSION 5.00
Begin VB.Form frmBancoAccChangePass 
   BorderStyle     =   0  'None
   Caption         =   "Boveda"
   ClientHeight    =   3225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3585
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   3585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtPass2 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1400
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1750
      Width           =   1815
   End
   Begin VB.TextBox txtPass1 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1400
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox txtToken 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1400
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   660
      Width           =   1815
   End
   Begin VB.Image imgCerrar 
      Height          =   195
      Left            =   3280
      Top             =   150
      Width           =   195
   End
   Begin VB.Image imgCambiar 
      Height          =   345
      Left            =   210
      Top             =   2520
      Width           =   3195
   End
End
Attribute VB_Name = "frmBancoAccChangePass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager
Private cBotonCerrar As clsGraphicalButton
Private cBotonCambiar As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then CloseWindow
End Sub

Private Sub Form_Load()
    Set clsFormulario = New clsFormMovementManager
On Error GoTo ErrHandler
  
    clsFormulario.Initialize Me
    
    Me.Picture = LoadPicture(DirInterfaces & SELECTED_UI & "VentanaBovedaCuentaCambiarContrasenia.jpg")
    Call LoadButtons
    
    Call modCustomCursors.SetFormCursorDefault(Me)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Form_Load de frmBancoAccChangePass.frm")
End Sub

Sub LoadButtons()
On Error GoTo ErrHandler
  
    Dim GrhPath As String
    
    GrhPath = DirInterfaces & SELECTED_UI

    
    Set cBotonCerrar = New clsGraphicalButton
    Set cBotonCambiar = New clsGraphicalButton
    
    Set LastButtonPressed = New clsGraphicalButton

    Call cBotonCambiar.Initialize(imgCambiar, GrhPath & "BotonBovedaCuentaCambiar.jpg", _
                                    GrhPath & "BotonBovedaCuentaCambiarRollover.jpg", _
                                    GrhPath & "BotonBovedaCuentaCambiarClick.jpg", Me)

    Call cBotonCerrar.Initialize(imgCerrar, GrhPath & "BotonCruzSalir.jpg", _
                                    GrhPath & "BotonCruzSalirRollover.jpg", _
                                    GrhPath & "BotonCruzSalirClick.jpg", Me)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadButtons de frmBancoAccChangePass.frm")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LastButtonPressed.ToggleToNormal
  
End Sub

Private Sub imgCambiar_Click()
On Error GoTo ErrHandler
  
    If txtToken.text = vbNullString Then
        Call MsgBox("Debe ingresar el Token.")
        Exit Sub
    End If
    
    If Len(txtPass1.text) > 16 Then
        Call MsgBox("La contraseña no puede tener más de 16 caracteres de largo.")
        Exit Sub
    End If
    
    If txtPass1.text <> txtPass2.text Then
        Call MsgBox("La confirmación es diferente a la contraseña.")
        Exit Sub
    End If
    
    If Not AsciiValidos(txtPass1.text) Then
        Call MsgBox("No puedes utilizar caracteres inválidos.")
        Exit Sub
    End If
    
    Call WriteAccBankChangePass(txtToken.text, txtPass1.text)
    Unload Me
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgCambiar_Click de frmBancoAccChangePass.frm")
End Sub

Private Sub imgCerrar_Click()
    Unload Me
On Error GoTo ErrHandler
  
    frmMain.SetFocus
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgCerrar_Click de frmBancoAccChangePass.frm")
End Sub
Private Sub CloseWindow()
On Error GoTo ErrHandler
  
    Unload Me
    If frmBancoAcc.Visible Then
        frmBancoAcc.SetFocus
    Else
        If frmBancoAccPass.Visible Then
            frmBancoAccPass.SetFocus
        Else
            If frmMain.Visible Then
                frmMain.SetFocus
            End If
        End If
    End If
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CloseWindow de frmBancoAccChangePass.frm")
End Sub
