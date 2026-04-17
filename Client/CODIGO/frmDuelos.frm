VERSION 5.00
Begin VB.Form frmDuelos 
   BorderStyle     =   0  'None
   ClientHeight    =   2580
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4815
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image img4vs4 
      Height          =   225
      Left            =   3480
      Top             =   1900
      Width           =   735
   End
   Begin VB.Image imgPublico1vs1 
      Height          =   225
      Left            =   2040
      Top             =   900
      Width           =   735
   End
   Begin VB.Image img3vs3 
      Height          =   225
      Left            =   2520
      Top             =   1900
      Width           =   735
   End
   Begin VB.Image img2vs2 
      Height          =   225
      Left            =   1560
      Top             =   1900
      Width           =   735
   End
   Begin VB.Image img1vs1 
      Height          =   225
      Left            =   600
      Top             =   1900
      Width           =   735
   End
   Begin VB.Image imgCerrar 
      Height          =   195
      Left            =   4485
      Top             =   150
      Width           =   195
   End
End
Attribute VB_Name = "frmDuelos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager
Private cBotonCerrar As clsGraphicalButton
Private cBotonPublico1v1 As clsGraphicalButton
Private cBotonPrivado1v1 As clsGraphicalButton
Private cBotonPrivado2v2 As clsGraphicalButton
Private cBotonPrivado3v3 As clsGraphicalButton
Private cBotonPrivado4v4 As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then CloseWindow
End Sub

Private Sub Form_Load()
Set clsFormulario = New clsFormMovementManager
On Error GoTo ErrHandler
  
    clsFormulario.Initialize Me
    
    Me.Picture = LoadPicture(DirInterfaces & SELECTED_UI & "VentanaDuelos.jpg")
    Call ActivarBotones
    
    Call modCustomCursors.SetFormCursorDefault(Me)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Form_Load de frmDuelos.frm")
End Sub

Sub ActivarBotones()
On Error GoTo ErrHandler
  
    Dim GrhPath As String
    
    GrhPath = DirInterfaces & SELECTED_UI

    
    Set cBotonCerrar = New clsGraphicalButton
    Set cBotonPublico1v1 = New clsGraphicalButton
    Set cBotonPrivado1v1 = New clsGraphicalButton
    Set cBotonPrivado2v2 = New clsGraphicalButton
    Set cBotonPrivado3v3 = New clsGraphicalButton
    Set cBotonPrivado4v4 = New clsGraphicalButton
    
    Set LastButtonPressed = New clsGraphicalButton

    Call cBotonPublico1v1.Initialize(imgPublico1vs1, GrhPath & "BotonDuelo1v1.jpg", _
                                    GrhPath & "BotonDuelo1v1Rollover.jpg", _
                                    GrhPath & "BotonDuelo1v1Click.jpg", Me)
                               
    Call cBotonPrivado1v1.Initialize(img1vs1, GrhPath & "BotonDuelo1v1.jpg", _
                                    GrhPath & "BotonDuelo1v1Rollover.jpg", _
                                    GrhPath & "BotonDuelo1v1Click.jpg", Me)
                                    
    Call cBotonPrivado2v2.Initialize(img2vs2, GrhPath & "BotonDuelo2v2.jpg", _
                                    GrhPath & "BotonDuelo2v2Rollover.jpg", _
                                    GrhPath & "BotonDuelo2v2Click.jpg", Me)
                                    
    Call cBotonPrivado3v3.Initialize(img3vs3, GrhPath & "BotonDuelo3v3.jpg", _
                                    GrhPath & "BotonDuelo3v3Rollover.jpg", _
                                    GrhPath & "BotonDuelo3v3Click.jpg", Me)
                                    
    Call cBotonPrivado4v4.Initialize(img4vs4, GrhPath & "BotonDuelo4v4.jpg", _
                                    GrhPath & "BotonDuelo4v4Rollover.jpg", _
                                    GrhPath & "BotonDuelo4v4Click.jpg", Me)
                                    
    Call cBotonCerrar.Initialize(imgCerrar, GrhPath & "BotonCruzSalir.jpg", _
                                    GrhPath & "BotonCruzSalirRollover.jpg", _
                                    GrhPath & "BotonCruzSalirClick.jpg", Me)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ActivarBotones de frmDuelos.frm")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub img1vs1_Click()
On Error GoTo ErrHandler
  
    frmDuelo1v1.Show , frmMain
    Unload Me
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub img1vs1_Click de frmDuelos.frm")
End Sub

Private Sub img2vs2_Click()
On Error GoTo ErrHandler
  
    frmDuelo2v2.Show , frmMain
    Unload Me
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub img2vs2_Click de frmDuelos.frm")
End Sub

Private Sub img3vs3_Click()
On Error GoTo ErrHandler
  
    frmDuelo3v3.Show , frmMain
    Unload Me
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub img3vs3_Click de frmDuelos.frm")
End Sub

Private Sub img4vs4_Click()
On Error GoTo ErrHandler
  
    frmDuelo4v4.Show , frmMain
    Unload Me
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub img4vs4_Click de frmDuelos.frm")
End Sub

Private Sub imgCerrar_Click()
    Unload Me
  
End Sub

Private Sub imgPublico1vs1_Click()
    Call WriteDueloPublico
On Error GoTo ErrHandler
  
    Unload Me
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgPublico1vs1_Click de frmDuelos.frm")
End Sub

Private Sub CloseWindow()
On Error GoTo ErrHandler
  
    Unload Me
    If frmMain.Visible Then frmMain.SetFocus
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CloseWindow de frmDuelos.frm")
End Sub
