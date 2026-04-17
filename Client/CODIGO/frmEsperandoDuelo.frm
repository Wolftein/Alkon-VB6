VERSION 5.00
Begin VB.Form frmEsperandoDuelo 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   ScaleHeight     =   1800
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image imgCancelar 
      Height          =   390
      Left            =   1440
      Top             =   960
      Width           =   1875
   End
End
Attribute VB_Name = "frmEsperandoDuelo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager
Private BotonCancelar As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub Form_Load()
    
On Error GoTo ErrHandler
  
Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Call LoadButtons
    
    Call modCustomCursors.SetFormCursorDefault(Me)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Form_Load de frmEsperandoDuelo.frm")
End Sub

Private Sub LoadButtons()
On Error GoTo ErrHandler
  
    Dim GrhPath As String
    
    GrhPath = DirInterfaces & SELECTED_UI

    Set BotonCancelar = New clsGraphicalButton
    
    Set LastButtonPressed = New clsGraphicalButton

    Call BotonCancelar.Initialize(imgCancelar, GrhPath & "BotonDueloCancelarEspera.jpg", _
                                    GrhPath & "BotonDueloCancelarEsperaRollover.jpg", _
                                    GrhPath & "BotonDueloCancelarEsperaClick.jpg", Me)
    

    Me.Picture = LoadPicture(GrhPath & "VentanaEsperandoDuelo.jpg")
    Me.Icon = frmMain.Icon
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadButtons de frmEsperandoDuelo.frm")
End Sub

Private Sub imgCancelar_Click()
    Call WriteCancelarEspera
End Sub
