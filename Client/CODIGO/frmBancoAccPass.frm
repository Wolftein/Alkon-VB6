VERSION 5.00
Begin VB.Form frmBancoAccPass 
   BorderStyle     =   0  'None
   Caption         =   "Boveda"
   ClientHeight    =   2100
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4155
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   4155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtPass 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   460
      Width           =   2415
   End
   Begin VB.Image imgCerrar 
      Height          =   195
      Left            =   3840
      Top             =   150
      Width           =   195
   End
   Begin VB.Image imgCambiar 
      Height          =   345
      Left            =   500
      Top             =   1440
      Width           =   3195
   End
   Begin VB.Image imgIngresar 
      Height          =   345
      Left            =   500
      Top             =   1080
      Width           =   3195
   End
End
Attribute VB_Name = "frmBancoAccPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager
Private cBotonCerrar As clsGraphicalButton
Private cBotonCambiar As clsGraphicalButton
Private cBotonIngresar As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then CloseWindow
End Sub

Private Sub Form_Load()
    Set clsFormulario = New clsFormMovementManager
On Error GoTo ErrHandler
  
    clsFormulario.Initialize Me
    
    Me.Picture = LoadPicture(DirInterfaces & SELECTED_UI & "VentanaBovedaCuentaPass.jpg")
    Call LoadButtons
    
    Call modCustomCursors.SetFormCursorDefault(Me)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Form_Load de frmBancoAccPass.frm")
End Sub

Sub LoadButtons()
On Error GoTo ErrHandler
  
    Dim GrhPath As String
    
    GrhPath = DirInterfaces & SELECTED_UI

    
    Set cBotonCerrar = New clsGraphicalButton
    Set cBotonCambiar = New clsGraphicalButton
    Set cBotonIngresar = New clsGraphicalButton
    
    Set LastButtonPressed = New clsGraphicalButton

    Call cBotonCambiar.Initialize(imgCambiar, GrhPath & "BotonBovedaCuentaCambiar.jpg", _
                                    GrhPath & "BotonBovedaCuentaCambiarRollover.jpg", _
                                    GrhPath & "BotonBovedaCuentaCambiarClick.jpg", Me)

    Call cBotonIngresar.Initialize(imgIngresar, GrhPath & "BotonBovedaCuentaIngresar.jpg", _
                                    GrhPath & "BotonBovedaCuentaIngresarRollover.jpg", _
                                    GrhPath & "BotonBovedaCuentaIngresarClick.jpg", Me)
                                    
    Call cBotonCerrar.Initialize(imgCerrar, GrhPath & "BotonCruzSalir.jpg", _
                                    GrhPath & "BotonCruzSalirRollover.jpg", _
                                    GrhPath & "BotonCruzSalirClick.jpg", Me)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadButtons de frmBancoAccPass.frm")
End Sub
Private Sub imgCambiar_Click()
    frmBancoAccChangePass.Show , frmMain
  
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LastButtonPressed.ToggleToNormal

End Sub

Private Sub imgCerrar_Click()
On Error GoTo ErrHandler
  
    Unload Me
    frmMain.SetFocus
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgCerrar_Click de frmBancoAccPass.frm")
End Sub

Private Sub imgIngresar_Click()
    Call WriteAccBankStart(txtPass.text)
End Sub
Private Sub CloseWindow()
On Error GoTo ErrHandler
  
    Unload Me
    If frmMain.Visible Then frmMain.SetFocus
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CloseWindow de frmBancoAccPass.frm")
End Sub
