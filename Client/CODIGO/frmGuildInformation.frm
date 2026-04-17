VERSION 5.00
Begin VB.Form frmGuildInformation 
   BackColor       =   &H00292929&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5115
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8145
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   341
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   543
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblMemberCount 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5520
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label lblContribution 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2925
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label lblGuildLeader 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Guild Leader"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   1080
      Width           =   4695
   End
End
Attribute VB_Name = "frmGuildInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_GotFocus()
    Call FillControlInfo
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then CloseWindow
End Sub

Private Sub Form_Load()

    Call LoadControls
    
    Call FillControlInfo
    
    Call modCustomCursors.SetFormCursorDefault(Me)
End Sub


Private Sub LoadControls()

    Dim GrhPath As String
    GrhPath = DirInterfaces & SELECTED_UI
    
    Me.Picture = LoadPicture(GrhPath & "VentanaGuildInfo.jpg")
    
End Sub

Private Sub FillControlInfo()

    lblGuildLeader.Caption = GetGuildLeaderName()
    lblMemberCount.Caption = PlayerData.Guild.MemberCount & "/" & PlayerData.Guild.MaxMemberQty
    lblContribution.Caption = PlayerData.Guild.ContributionAvailable & "/" & PlayerData.Guild.MaxContributionAvailable
    
End Sub

Private Function GetGuildLeaderName() As String
    
    Dim I As Integer
    
    For I = 1 To PlayerData.Guild.MemberCount
        If PlayerData.Guild.Members(I).RoleId = 1 Then
            GetGuildLeaderName = PlayerData.Guild.Members(I).UserName
            Exit Function
        End If
    
    Next I
    
End Function

Private Sub CloseWindow()
On Error GoTo ErrHandler
  
    Unload Me
    If frmGuildMain.Visible Then
        Unload frmGuildMain
    End If
    If frmMain.Visible Then frmMain.SetFocus
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CloseWindow de frmGuildInformation.frm")
End Sub
