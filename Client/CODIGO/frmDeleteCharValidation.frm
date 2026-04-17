VERSION 5.00
Begin VB.Form frmDeleteCharValidation 
   BorderStyle     =   0  'None
   ClientHeight    =   1860
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6840
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   124
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   456
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtToken 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   690
      Width           =   2970
   End
   Begin VB.Image imgCancelar 
      Height          =   450
      Left            =   4080
      Top             =   1200
      Width           =   2250
   End
   Begin VB.Image imgConfirmar 
      Height          =   450
      Left            =   360
      Top             =   1200
      Width           =   2250
   End
End
Attribute VB_Name = "frmDeleteCharValidation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager

Private cBotonCancelar As clsGraphicalButton
Private cBotonConfirmar As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton

Private Sub txtToken_Validator()
End Sub

Private Sub Form_Load()
' Handles Form movement (drag and drop).
On Error GoTo ErrHandler
  
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Me.Picture = LoadPicture(DirInterfaces & SELECTED_UI & "VentanaBorrarChar.jpg")
    Call LoadButtons
    
    Call modCustomCursors.SetFormCursorDefault(Me)
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Form_Load de frmDeleteCharValidation.frm")
End Sub

Private Sub LoadButtons()
On Error GoTo ErrHandler
  
    Dim GrhPath As String
    
    GrhPath = DirInterfaces & SELECTED_UI

    Set cBotonCancelar = New clsGraphicalButton
    Set cBotonConfirmar = New clsGraphicalButton
    
    Set LastButtonPressed = New clsGraphicalButton
    
    
    Call cBotonCancelar.Initialize(imgCancelar, GrhPath & "BotonDeleteCharCancelar.jpg", _
                                    GrhPath & "BotonDeleteCharCancelarRollover.jpg", _
                                    GrhPath & "BotonDeleteCharCancelarClick.jpg", Me)

    Call cBotonConfirmar.Initialize(imgConfirmar, GrhPath & "BotonDeleteCharConfirmar.jpg", _
                                    GrhPath & "BotonDeleteCharConfirmarRollover.jpg", _
                                    GrhPath & "BotonDeleteCharConfirmarClick.jpg", Me)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadButtons de frmDeleteCharValidation.frm")
End Sub

Private Sub imgCancelar_Click()
    CerrarVentana
End Sub

Private Sub imgConfirmar_Click()
On Error GoTo ErrHandler
    
    If Not MainTimer.Check(TimersIndex.Action) Then Exit Sub
    
    If txtToken = "" Then
        MsgBox "Debes indicar un token para borrar el personaje"
        Exit Sub
    End If
    If MsgBox("Esta seguro que desea eliminar a " & Acc_Data.Acc_Char(Acc_Data.Acc_Char_Selected).Char_Name & "?", vbYesNo, "Eliminar Personaje") = vbYes Then
        Call Set_Acc_Data_To_Delete
        Call modAccount.Prepare_And_Connect(E_MODO.AccountDeleteChar)
    End If
    Call CerrarVentana
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgConfirmar_Click de frmDeleteCharValidation.frm")
End Sub

Private Sub CerrarVentana()
On Error GoTo ErrHandler
  
    Unload Me
    If frmMain.Visible Then frmMain.SetFocus
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CerrarVentana de frmDeleteCharValidation.frm")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
  
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then CerrarVentana
End Sub

