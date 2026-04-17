VERSION 5.00
Begin VB.Form frmAccountChangePassword 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Cambiar Contraseña"
   ClientHeight    =   3555
   ClientLeft      =   0
   ClientTop       =   -15
   ClientWidth     =   4755
   Icon            =   "frmAccountChangePassword.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Palette         =   "frmAccountChangePassword.frx":030A
   ScaleHeight     =   237
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   317
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtNewPassword 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   265
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   380
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2460
      Width           =   4000
   End
   Begin VB.TextBox txtNewPassword 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   265
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   380
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1750
      Width           =   4000
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   265
      IMEMode         =   3  'DISABLE
      Left            =   380
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1025
      Width           =   4000
   End
   Begin VB.Image imgCerrar 
      Height          =   210
      Left            =   4380
      Top             =   195
      Width           =   210
   End
   Begin VB.Image imgAceptar 
      Height          =   375
      Left            =   990
      Top             =   2790
      Width           =   2775
   End
End
Attribute VB_Name = "frmAccountChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cBotonSalir As clsGraphicalButton
Private cBotonAceptar As clsGraphicalButton
Public LastButtonPressed As clsGraphicalButton

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then CerrarVentana
End Sub

Private Sub Form_Load()
    Me.Picture = LoadPicture(DirInterfaces & SELECTED_UI & "VentanaCambiarContrasenia.jpg")
    Call LoadButtons
    
    Call modCustomCursors.SetFormCursorDefault(Me)
End Sub

Private Sub LoadButtons()
On Error GoTo ErrHandler
  
    
    Dim GrhPath As String
    
    GrhPath = DirInterfaces & SELECTED_UI
    
    Set cBotonSalir = New clsGraphicalButton
    Set cBotonAceptar = New clsGraphicalButton
    Set LastButtonPressed = New clsGraphicalButton
                         
    Call cBotonSalir.Initialize(imgCerrar, GrhPath & "BotonCerrarForm.jpg", _
                                    GrhPath & "BotonCerrarFormRollover.jpg", _
                                    GrhPath & "BotonCerrarFormClick.jpg", Me)
                                    
    Call cBotonAceptar.Initialize(imgAceptar, GrhPath & "BotonAceptarCambioContrasenia.jpg", _
                                    GrhPath & "BotonAceptarCambioContraseniaRollover.jpg", _
                                    GrhPath & "BotonAceptarCambioContraseniaClick.jpg", Me)
                          
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadButtons de frmAccountChangePassword.frm")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Public Sub imgAceptar_Click()
On Error GoTo ErrHandler
    If ValidInput(txtPassword.text, txtNewPassword(0).text, txtNewPassword(1).text) Then
  
        If txtNewPassword(0).text = txtNewPassword(1).text Then
            ' Inverval per connection.
            If Not MainTimer.Check(TimersIndex.Action) Then Exit Sub
            
            With Acc_Data
                ' Prepare data.
                .Acc_Password = txtPassword.text
                .Acc_New_Password = txtNewPassword(0).text

                .Acc_Password = MD5.GetMD5String(.Acc_Password)
                Call MD5.MD5Reset
                
                .Acc_New_Password = MD5.GetMD5String(.Acc_New_Password)
                Call MD5.MD5Reset

            End With
            
            ' Connect.
            Call modAccount.Prepare_And_Connect(E_MODO.AccountChangePassword)
        Else
            Call MsgBox("Las contraseñas no coinciden.")
        End If
    Else
        Call MsgBox("Complete todos los campos.")
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgAceptar_Click de frmAccountChangePassword.frm")
End Sub

Private Sub imgCerrar_Click()
    Call CerrarVentana
End Sub

Private Sub CerrarVentana()
On Error GoTo ErrHandler
  
    Unload Me
    
    If frmAccount.Visible Then frmAccount.SetFocus
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CerrarVentana de frmAccountChangePassword.frm")
End Sub
