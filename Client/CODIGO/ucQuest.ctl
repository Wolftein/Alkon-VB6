VERSION 5.00
Begin VB.UserControl ucQuest 
   BackColor       =   &H00292929&
   ClientHeight    =   525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3915
   ScaleHeight     =   525
   ScaleWidth      =   3915
   Begin VB.Label LblNumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "1"
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
      Height          =   495
      Left            =   50
      TabIndex        =   1
      Top             =   160
      Width           =   405
   End
   Begin VB.Label LblName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
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
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   160
      Width           =   3360
   End
End
Attribute VB_Name = "ucQuest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private QuestId As Integer

Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

Public Sub SetQuest(ByVal Id As Integer, ByVal Name As String)
    QuestId = Id
    LblName.Caption = Name
    LblNumber.Caption = QuestId
End Sub

Private Sub LblName_Click()
    Call Edit
End Sub


Private Sub LblNumber_Click()
    Call Edit
End Sub

Private Sub Edit()
    Dim Quest As tQuest
    Dim I As Integer
    
    'For I = 1 To GameMetadata.GuildQuestsQty
    '    If GameMetadata.GuildQuests(I).Id = QuestId Then
    '        Quest = GameMetadata.GuildQuests(I)
    '        Exit For
    '    End If
    'Next I
    
    Call frmGuildQuests.SetHiddenShapeVisibility(False)
    
    With Quest
        frmGuildQuests.LblActualQuestDescription.Caption = GameMetadata.GuildQuests(QuestId).Desc
        frmGuildQuests.LblActualQuestName.Caption = GameMetadata.GuildQuests(QuestId).Title
        frmGuildQuests.SelectedQuestId = QuestId
    End With
    
End Sub

Private Sub UserControl_Initialize()

    Dim GrhPath As String
    GrhPath = DirInterfaces & SELECTED_UI
    
    UserControl.Picture = LoadPicture(GrhPath & "VentanaGuildQuestItem.jpg")

    LblNumber.MousePointer = vbCustom
    LblNumber.MouseIcon = picMouseIcon
    
    LblName.MousePointer = vbCustom
    LblName.MouseIcon = picMouseIcon
    
    UserControl.MousePointer = vbCustom
    UserControl.MouseIcon = picMouseIcon
End Sub

