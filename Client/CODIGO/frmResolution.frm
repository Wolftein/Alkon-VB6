VERSION 5.00
Begin VB.Form frmResolution 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   1845
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6840
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmResolution.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   123
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   456
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image imgCheck 
      Height          =   210
      Left            =   315
      Top             =   1260
      Width           =   225
   End
   Begin VB.Image imgNo 
      Height          =   540
      Left            =   2040
      Top             =   1080
      Width           =   1200
   End
   Begin VB.Image imgYes 
      Height          =   540
      Left            =   3600
      Top             =   1080
      Width           =   1200
   End
End
Attribute VB_Name = "frmResolution"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager

Private cBotonSi As clsGraphicalButton
Private cBotonNo As clsGraphicalButton
Private cBotonCheck As clsGraphicalButton
Public LastButtonPressed As clsGraphicalButton


Private checked As Boolean
Private imgTemp As Picture

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then CloseWindow
End Sub

Private Sub Form_Load()
    Me.Picture = LoadPicture(DirInterfaces & SELECTED_UI & "VentanaResolucion.jpg")
    checked = False
On Error GoTo ErrHandler

    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
  
    Call LoadButtons
    
    Call modCustomCursors.SetFormCursorDefault(Me)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Form_Load de frmResolution.frm")
End Sub

Private Sub LoadButtons()
On Error GoTo ErrHandler
  
    Dim GrhPath As String

    GrhPath = DirInterfaces & SELECTED_UI
    
    Set cBotonSi = New clsGraphicalButton
    Set cBotonNo = New clsGraphicalButton
    Set cBotonCheck = New clsGraphicalButton
    Set LastButtonPressed = New clsGraphicalButton
    
    Call cBotonSi.Initialize(imgYes, GrhPath & "BotonSi.jpg", _
                                        GrhPath & "BotonSi.jpg", _
                                        GrhPath & "BotonSi.jpg", Me)
                                        
    Call cBotonNo.Initialize(imgNo, GrhPath & "BotonNo.jpg", _
                                        GrhPath & "BotonNo.jpg", _
                                        GrhPath & "BotonNo.jpg", Me)
                                        
    Call cBotonSi.SoundEnabled(False)
    Call cBotonNo.SoundEnabled(False)
    Call cBotonCheck.SoundEnabled(False)
    Call LastButtonPressed.SoundEnabled(False)
                                                               
    Set imgTemp = LoadPicture(GrhPath & "CheckEnabled.jpg")
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadButtons de frmResolution.frm")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub imgCheck_Click()
    checked = Not checked
On Error GoTo ErrHandler
  
    
    If checked Then
        imgCheck.Picture = imgTemp
    Else
        Set imgCheck.Picture = Nothing
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgCheck_Click de frmResolution.frm")
End Sub

Private Sub SetOption(ByVal bFullScreen As Boolean)

    GameConfig.Extras.bAskForResolutionChange = Not checked
    GameConfig.Graphics.bUseFullScreen = bFullscreen
            
    Me.Visible = False
    Unload Me

End Sub

Private Sub imgNo_Click()
    Call SetOption(False)
End Sub

Private Sub imgYes_Click()
    Call SetOption(True)
End Sub

Private Sub CloseWindow()
On Error GoTo ErrHandler
  
    Unload Me
    If frmMain.Visible Then frmMain.SetFocus
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CloseWindow de frmResolution.frm")
End Sub
