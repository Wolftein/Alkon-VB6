VERSION 5.00
Begin VB.Form frmGuildQuests 
   BackColor       =   &H00292929&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5160
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8055
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar VScroll 
      Height          =   4005
      Left            =   4200
      Max             =   100
      TabIndex        =   2
      Top             =   600
      Width           =   255
   End
   Begin VB.PictureBox PboQuests 
      BackColor       =   &H00292929&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4005
      Left            =   0
      ScaleHeight     =   4005
      ScaleWidth      =   4485
      TabIndex        =   1
      Top             =   600
      Width           =   4485
   End
   Begin VB.PictureBox PboContainer 
      BackColor       =   &H00292929&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4000
      Left            =   0
      ScaleHeight     =   4005
      ScaleWidth      =   4485
      TabIndex        =   0
      Top             =   600
      Width           =   4485
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Misiones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   240
      Width           =   3615
   End
   Begin VB.Shape ShapeHideElements 
      BorderStyle     =   0  'Transparent
      Height          =   4935
      Left            =   4680
      Top             =   120
      Width           =   3255
   End
   Begin VB.Image ImgIniciar 
      Height          =   555
      Left            =   5745
      Top             =   4320
      Width           =   1290
   End
   Begin VB.Label LblActualQuestDescription 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   1935
      Left            =   4890
      TabIndex        =   4
      Top             =   800
      Width           =   2895
   End
   Begin VB.Label LblActualQuestName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre Misión"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   4800
      TabIndex        =   3
      Top             =   200
      Width           =   3135
   End
End
Attribute VB_Name = "frmGuildQuests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ControlNamePattern As String = "quest_"
Public SelectedQuestId As Integer

Private MemberListCtl As Control

Private cButtonStartQuest As clsGraphicalButton
Public LastButtonPressed As clsGraphicalButton


Public LastControlPosition As Integer

Public Sub ShowData()
    
    LastControlPosition = 200
    
    Call CleanQuestElementControls
   
    If GameMetadata.GuildQuestsQty <= 0 Then
        Exit Sub
    End If
    
    Call ShowAvailableQuests(LastControlPosition)

    Call SetHiddenShapeVisibility(True)
    
    If PboContainer.Height > LastControlPosition Then
        VScroll.Visible = False
    End If
    
    VScroll.LargeChange = PboContainer.Height / LastControlPosition * 100
    VScroll.value = 0
    VScroll.Top = 400
    
    PboQuests.Height = LastControlPosition
    PboQuests.Left = 0
    PboQuests.Top = 0
End Sub

Private Sub Form_Load()
    Call LoadControls

    Call ShowData
    
    Call modCustomCursors.SetFormCursorDefault(Me)
End Sub

Public Sub CleanQuestElementControls()
    Dim I As Integer
    Dim ControlName As String
    
    For I = Controls.Count - 1 To 0 Step -1
        ControlName = Me.Controls(I).Name
        If InStr(1, ControlName, ControlNamePattern) = 1 Then
            Call Me.Controls.Remove(ControlName)
        End If
    Next I
End Sub

Public Sub LoadControls()

    Dim GrhPath As String
    GrhPath = DirInterfaces & SELECTED_UI
    
    Me.Picture = LoadPicture(GrhPath & "VentanaGuildQuests.jpg")
    
    Set cButtonStartQuest = New clsGraphicalButton
    Set LastButtonPressed = New clsGraphicalButton
    
    
    Call cButtonStartQuest.Initialize(ImgIniciar, GrhPath & "BotonIniciar.jpg", _
                                    GrhPath & "BotonIniciar.jpg", _
                                    GrhPath & "BotonIniciar.jpg", Me)
                                    'GrhPath & "BotonIniciar.jpg", _
                                    'GrhPath & "BotonIniciar.jpg", Me)
                                   
    Call cButtonStartQuest.EnableButton(PlayerData.Guild.IdRolOwn = ID_ROLE_LEADER)
    
End Sub


Public Sub SetHiddenShapeVisibility(ByVal Visible As Boolean)
    
    ShapeHideElements.BackColor = &H292929
    ShapeHideElements.Visible = Visible
    ShapeHideElements.BackStyle = IIf(Visible, 1, 0)
    
End Sub

Private Function ShowAvailableQuests(ByRef LastControlPosition As Integer) As Integer
    Dim I As Integer
   
    Dim QuestNotCompletedOrRepeatabale As Boolean
    Dim CorrelativesCompleted As Boolean
    Dim IsAlignmentCorrect As Boolean
    
    For I = 1 To GameMetadata.GuildQuestsQty

        With GameMetadata.GuildQuests(I)
        
            QuestNotCompletedOrRepeatabale = Not IsQuestCompleted(.Id) Or (IsQuestCompleted(.Id) And .RepetitionQuantity > 0)
            CorrelativesCompleted = HasCorrelativesCompleted(.Id)
      
            If .Alignment > 0 Then
                IsAlignmentCorrect = (PlayerData.Guild.Alignment = .Alignment)
            Else
                IsAlignmentCorrect = True
            End If
            
            If QuestNotCompletedOrRepeatabale And CorrelativesCompleted And IsAlignmentCorrect Then
                LastControlPosition = LastControlPosition + ShowQuest(GameMetadata.GuildQuests(I), LastControlPosition)
            End If
        End With

    Next I
    
End Function

Private Function HasCorrelativesCompleted(ByVal QuestId As Integer) As Boolean
    Dim I As Integer
    Dim J As Integer
    Dim RequiredQuestsCompletedQty As Integer
    Dim CorrelativesCompleted() As Boolean
    
    
    With GameMetadata.GuildQuests(QuestId)
    
        If .CorrelativesQuantity <= 0 Then
            HasCorrelativesCompleted = True
            Exit Function
        End If
        
        ReDim CorrelativesCompleted(1 To GameMetadata.GuildQuests(QuestId).CorrelativesQuantity)
            
        For I = 1 To GameMetadata.GuildQuests(QuestId).CorrelativesQuantity
            For J = 1 To PlayerData.Guild.Quest.CompletedQuantiy
                If GameMetadata.GuildQuests(QuestId).Correlatives(I).IdQuest = PlayerData.Guild.Quest.Completed(J) Then
                    CorrelativesCompleted(I) = True
                End If
            Next J
        Next I
        
        For I = 1 To .CorrelativesQuantity
            If CorrelativesCompleted(I) = False Then Exit Function
        Next I
        
        HasCorrelativesCompleted = True
    
    End With
    
End Function

Private Function IsQuestCompleted(ByVal QuestId As Integer) As Boolean
    Dim I As Integer
    
    IsQuestCompleted = False
    
    For I = 1 To PlayerData.Guild.Quest.CompletedQuantiy
        If PlayerData.Guild.Quest.Completed(I) = QuestId Then
            IsQuestCompleted = True
            Exit Function
        End If
    Next I
    
End Function


Private Function ShowQuest(ByRef Quest As tQuest, ByVal Position As Long) As Long
    Dim MemberListCtl As Control


    Set MemberListCtl = Controls.Add("ARGENTUM.ucQuest", ControlNamePattern & Quest.Id)

    With MemberListCtl

       ' Fix this. If the UserControl full name is bigger than 39 chars, this won't work
        ' because of a runtime error 1741: https://windows10dll.nirsoft.net/msvbvm60_dll.html
        .Top = IIf(LastControlPosition = 1, 0, LastControlPosition + .Height - 15)
        .Left = 0
        .Visible = True

        SetParent .hwnd, PboQuests.hwnd
       .Move 200, Position, .Width, .Height

        Call MemberListCtl.SetQuest(Quest.Id, Quest.Title)
    End With

    ShowQuest = MemberListCtl.Height

    DoEvents
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set PboQuests = Nothing
    Set MemberListCtl = Nothing
End Sub

Private Sub ImgIniciar_Click()
    If SelectedQuestId = 0 Then Exit Sub
    If Not cButtonStartQuest.IsEnabled Then Exit Sub
    
    Call WriteGuildQuest(SelectedQuestId)
End Sub

Private Sub VScroll_Change()
    Call DoScroll
End Sub

Private Sub VScroll_Scroll()
    Call DoScroll
End Sub

Private Sub DoScroll()
    Dim Top As Double
    Top = (PboQuests.Height - PboContainer.Height) * VScroll.value / 100
    PboQuests.Top = -Top
End Sub
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
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CloseWindow de frmGuildQuests.frm")
End Sub


