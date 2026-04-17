VERSION 5.00
Begin VB.Form frmSelectGrupo 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2115
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2070
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2115
   ScaleWidth      =   2070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image imgInvite 
      Appearance      =   0  'Flat
      Height          =   630
      Left            =   360
      Top             =   480
      Width           =   1425
   End
   Begin VB.Image imgCerrar 
      Height          =   210
      Left            =   1650
      Top             =   210
      Width           =   210
   End
   Begin VB.Image imgParty 
      Appearance      =   0  'Flat
      Height          =   630
      Left            =   240
      Top             =   1200
      Width           =   1545
   End
End
Attribute VB_Name = "frmSelectGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager
Private cBotonInvite As clsGraphicalButton
Private cBotonParty As clsGraphicalButton
Private cBotonSalir As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton

Private Sub ImgInvite_Click()
    Call WritePartyInviteMember("")
    
    Unload Me
End Sub

Private Sub Form_Load()
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
        
    Call LoadButtons
    
    Call modCustomCursors.SetFormCursorDefault(Me)
    
End Sub

Private Sub imgParty_Click()
    
    Call WriteRequestPartyForm
    
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Call CerrarVentana
End Sub

Private Sub CerrarVentana()
On Error GoTo ErrHandler
  
    Unload Me
    If frmMain.Visible Then frmMain.SetFocus
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CerrarVentana de frmParty.frm")
End Sub

Private Sub imgCerrar_Click()
    Call CerrarVentana
End Sub

Private Sub LoadButtons()
On Error GoTo ErrHandler
  
    Dim GrhPath As String
    
    GrhPath = DirInterfaces & SELECTED_UI
    
    Set cBotonInvite = New clsGraphicalButton
    Set cBotonParty = New clsGraphicalButton
    Set cBotonSalir = New clsGraphicalButton
    
    Set LastButtonPressed = New clsGraphicalButton
    
    Me.Picture = LoadPicture(GrhPath & "VentanaSelGrupos.jpg")
                       
    Call cBotonParty.Initialize(imgInvite, GrhPath & "BotonGuildMemberInvite.jpg", _
                                    GrhPath & "BotonGuildMemberInvite.jpg", _
                                    GrhPath & "BotonGuildMemberInvite.jpg", Me)
                                    
    Call cBotonParty.Initialize(imgParty, GrhPath & "BotonSelGrupoPartyNormal.jpg", _
                                    GrhPath & "BotonSelGrupoPartyRollover.jpg", _
                                    GrhPath & "BotonSelGrupoPartyNormal.jpg", Me)
                                    

    Call cBotonSalir.Initialize(imgCerrar, GrhPath & "BotonOpcionesSalir.jpg", _
                                    GrhPath & "BotonOpcionesSalirRollover.jpg", _
                                    GrhPath & "BotonOpcionesSalirClick.jpg", Me)
                                                            
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadButtons de frmOpciones.frm")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub
