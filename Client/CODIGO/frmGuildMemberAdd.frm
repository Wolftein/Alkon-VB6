VERSION 5.00
Begin VB.Form frmGuildMemberAdd 
   Appearance      =   0  'Flat
   BackColor       =   &H00292929&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8145
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   346
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   543
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtNamePlayer 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   2880
      TabIndex        =   0
      Top             =   2280
      Width           =   2340
   End
   Begin VB.Image imgInviteButton 
      Height          =   600
      Left            =   3480
      Top             =   2625
      Width           =   1215
   End
End
Attribute VB_Name = "frmGuildMemberAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public cInviteButton As clsGraphicalButton
Public LastButtonPressed As clsGraphicalButton

Public GrhPath As String

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then CloseWindow
End Sub
Private Sub CloseWindow()
On Error GoTo ErrHandler
  
    Unload Me
    If frmGuildMain.Visible Then
        Unload frmGuildMain
    End If
    If frmMain.Visible Then frmMain.SetFocus
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CloseWindow de frmGuildMemberAdd.frm")
End Sub

Private Sub Form_Load()
    Call modCustomCursors.SetFormCursorDefault(Me)
    
    GrhPath = DirInterfaces & SELECTED_UI
    
    Me.Picture = LoadPicture(GrhPath & "VentanaGuildInvite.jpg")
    
    Call InitButtons
End Sub

Private Sub InitButtons()

    GrhPath = DirInterfaces & SELECTED_UI

    Set cInviteButton = New clsGraphicalButton
    Set LastButtonPressed = New clsGraphicalButton
    
    Call cInviteButton.Initialize(imgInviteButton, GrhPath & "BotonInvitar.jpg", _
                                    GrhPath & "BotonInvitar.jpg", _
                                    GrhPath & "BotonInvitar.jpg", Me)

End Sub

Private Sub Form_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub imgInviteButton_Click()
    If TxtNamePlayer.Text <> "" Then
        Call WriteGuildMember(PlayerData.Guild.IdGuild, UserIndex, eMemberAction.SendInvitation, TxtNamePlayer.Text)
    End If
    Exit Sub
End Sub
