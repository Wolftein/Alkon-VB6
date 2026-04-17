VERSION 5.00
Begin VB.Form frmGuildCreateStep3 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   350
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblGuildName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del Clan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   3120
      Width           =   3975
   End
   Begin VB.Image imgOpenGuild 
      Height          =   570
      Left            =   2400
      Top             =   4440
      Width           =   1290
   End
End
Attribute VB_Name = "frmGuildCreateStep3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager
Private cButtonOpenGuild As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton

Private Sub Form_Load()

    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me, , False
    
    Call LoadControls

    lblGuildName.Caption = GuildCreation.Name
    
    Call modCustomCursors.SetFormCursorDefault(Me)
End Sub

Private Sub ImgClose_Click()
    Unload frmGuildCreateStep1
    Unload frmGuildCreateStep2
    Unload frmGuildCreateStep3
    Call frmGuildMain.ShowFull
End Sub

Private Sub LoadControls()
    
    Set cButtonOpenGuild = New clsGraphicalButton
    
    Set LastButtonPressed = New clsGraphicalButton
    
    Dim GrhPath As String
    GrhPath = DirInterfaces & SELECTED_UI
    
    Me.Picture = LoadPicture(GrhPath & "VentanaGuildCreationStep3.jpg")
    
    Call cButtonOpenGuild.Initialize(imgOpenGuild, GrhPath & "BotonVerClan.jpg", _
                                    GrhPath & "BotonVerClan.jpg", _
                                    GrhPath & "BotonVerClan.jpg", Me)
                                        
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then CloseWindow
End Sub

Private Sub CloseWindow()
On Error GoTo ErrHandler
  
    Unload Me
    If frmMain.Visible Then frmMain.SetFocus
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CloseWindow de frmGuildCreateStep3.frm")
End Sub

Private Sub imgOpenGuild_Click()
    On Error GoTo ErrHandler

    If Not Guilds.InvitationEmpty Then
        Call Guilds.GuildPendingInvitation
        Exit Sub
    End If
    
    If PlayerData.Guild.IdGuild = 0 Then
        Call ShowConsoleMsg("No pertences a ningún clan.", 100, 100, 100, False, False)
        Exit Sub
    End If
    
    Call frmGuildMain.LoadForm(frmGuildInformation, "")
    Call frmGuildMain.ShowPartial
    
    Call CloseWindow
    Exit Sub
    
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgOpenGuild_Click de frmGuildCreateStep3.frm")
End Sub
