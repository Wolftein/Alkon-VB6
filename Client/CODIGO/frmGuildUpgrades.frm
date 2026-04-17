VERSION 5.00
Begin VB.Form frmGuildUpgrades 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Mejoras de Clan"
   ClientHeight    =   5115
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8145
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5115
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ARGENTUM.AOPictureBox PicGQstReq 
      Height          =   480
      Left            =   5880
      TabIndex        =   6
      Top             =   4320
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   847
   End
   Begin ARGENTUM.AOPictureBox PicGUpgReq 
      Height          =   480
      Left            =   3240
      TabIndex        =   5
      Top             =   4320
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   847
   End
   Begin ARGENTUM.AOPictureBox picGUpgrades 
      Height          =   1995
      Left            =   700
      TabIndex        =   4
      Top             =   780
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   3519
   End
   Begin VB.Image ImgInfo 
      Height          =   525
      Left            =   1080
      Top             =   3000
      Width           =   1230
   End
   Begin VB.Label LblUpgradeDesc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción de la misión"
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
      Height          =   1040
      Left            =   3180
      TabIndex        =   3
      Top             =   1200
      Width           =   4610
   End
   Begin VB.Label LblGoldCost 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4365
      TabIndex        =   2
      Top             =   3900
      Width           =   1215
   End
   Begin VB.Label LblContributionCost 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   6720
      TabIndex        =   1
      Top             =   3900
      Width           =   1215
   End
   Begin VB.Label LblUpgradeName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Upgrade Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3120
      TabIndex        =   0
      Top             =   600
      Width           =   4695
   End
End
Attribute VB_Name = "frmGuildUpgrades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cButtonAquireUpgrade As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton

Public WithEvents GUpgrades As clsGraphicalInventory
Attribute GUpgrades.VB_VarHelpID = -1
Public WithEvents GUpgReq As clsGraphicalInventory
Attribute GUpgReq.VB_VarHelpID = -1
Public WithEvents GQstReq As clsGraphicalInventory
Attribute GQstReq.VB_VarHelpID = -1

Private Const MAX_UPGRADES As Integer = 20
Private Const GRAPH_QUEST As Integer = 542
Private Const GRAPH_UNKNOWN_UPG As Integer = 542


Private Sub Initialize()
    
    If GUpgrades Is Nothing Then
        
        Set GUpgrades = New clsGraphicalInventory
        
        Call GUpgrades.Initialize(frmGuildUpgrades.picGUpgrades, MAX_UPGRADES, , , , 10, , , , , True, ItemSeparatorSizeInPixels:=1)
        
        Call LoadUpgrades
        
    End If

    Call InitializateGUpgReq
    Call InitializateGQstReq

    Exit Sub
    
End Sub


Private Sub Form_Load()

    Call LoadControls
    
    Call Initialize
    
    Call modCustomCursors.SetFormCursorDefault(Me)
End Sub

Public Sub LoadControls()

    Set cButtonAquireUpgrade = New clsGraphicalButton
    Set LastButtonPressed = New clsGraphicalButton
    
    Dim GrhPath As String
    GrhPath = DirInterfaces & SELECTED_UI
    
    Me.Picture = LoadPicture(GrhPath & "VentanaGuildUpgrades.jpg")
    
    Call cButtonAquireUpgrade.Initialize(ImgInfo, GrhPath & "BotonComprar.jpg", _
                                    GrhPath & "BotonComprar.jpg", _
                                    GrhPath & "BotonComprar.jpg", Me)
                                    'GrhPath & "BotonComprar.jpg", _
                                    'GrhPath & "BotonComprar.jpg", Me)

End Sub

Private Sub ImgInfo_Click()
    Dim I As Integer
   
    If GUpgrades.SelectedItem = 0 Then Exit Sub
    
    For I = 1 To GetQtyGuildUpgrades()
        If GUpgrades.Valor(GUpgrades.SelectedItem) = PlayerData.Guild.Upgrades(I).IdUpgrade Then
            MsgBox "Ya posee esta mejora."
            Exit Sub
        End If
    Next I
        
    Call WriteGuildUpgrade(GUpgrades.Valor(GUpgrades.SelectedItem))
End Sub

Private Sub picGUpgrades_Click()

    Dim UpgradeIndex As Integer
    
    If GUpgrades.SelectedItem = 0 Then Exit Sub
        
    UpgradeIndex = GUpgrades.Valor(GUpgrades.SelectedItem)
    
    Call Engine_Audio.PlayInterface(SND_CLICK)
    
    If UpgradeIndex <> 0 Then
    
        LblUpgradeName.Caption = GuildUpgrades(UpgradeIndex).Name
        LblUpgradeDesc.Caption = GuildUpgrades(UpgradeIndex).Description
        LblContributionCost = GuildUpgrades(UpgradeIndex).ContributionCost
        LblGoldCost = GuildUpgrades(UpgradeIndex).GoldCost

        Call Load_ReqUpgrade(UpgradeIndex)
        
        Call Load_ReqQuest(UpgradeIndex)
    End If

End Sub

Public Sub LoadUpgrades()
    Dim MaxUpgradesList As Integer, Obtained As Boolean
    Dim I As Integer, J As Integer, k As Integer
    Dim ItemPos As Integer, MaxUpgradeGroup As Integer
    
    MaxUpgradesList = GetQtyGuildUpgradesList()
    MaxUpgradeGroup = GetQtyGuildUpgradesGroup()

    ItemPos = 1

    If Not GUpgrades Is Nothing Then
       
        For I = 1 To MaxUpgradeGroup ' groups
            
            For J = 1 To GuildUpgradesGroup(I).UpgradeQty 'upgrade
                Obtained = False
                For k = 1 To GetQtyGuildUpgrades() 'upgrade Obtained
                    If PlayerData.Guild.Upgrades(k).IdUpgrade = GuildUpgradesGroup(I).Upgrades(J) Then
                        Obtained = True
                    End If
                Next k
                
                Call GUpgrades.SetItem(ItemPos, _
                    0, _
                    0, _
                    Obtained, _
                    IIf(GuildUpgrades(GuildUpgradesGroup(I).Upgrades(J)).IconGraph <> 0, GuildUpgrades(GuildUpgradesGroup(I).Upgrades(J)).IconGraph, GRAPH_UNKNOWN_UPG), _
                    0, _
                    0, _
                    0, _
                    0, _
                    0, _
                    GuildUpgradesGroup(I).Upgrades(J), _
                    GuildUpgrades(GuildUpgradesGroup(I).Upgrades(J)).Name, _
                    0, _
                    True)
                
                ItemPos = ItemPos + 1
                
            Next J
       
        Next I
    End If
    
End Sub
Private Sub Load_ReqUpgrade(ByVal UpgradeIndex As Integer)
    Dim I As Integer, J As Integer
    Dim QtyUpgradeReq As Integer
    Dim UpgradesQty As Integer
    Dim Obtained As Boolean

    Call GUpgReq.Release
    Set GUpgReq = Nothing
    
    Call InitializateGUpgReq

    If ((Not GuildUpgrades(UpgradeIndex).UpgradeRequired) = -1) Then
        QtyUpgradeReq = 0
    Else
        QtyUpgradeReq = UBound(GuildUpgrades(UpgradeIndex).UpgradeRequired)
    End If
    
    UpgradesQty = GetQtyGuildUpgrades()

    If QtyUpgradeReq = 0 Then
        Exit Sub
    End If
    
    If Not GUpgReq Is Nothing Then
        For I = 1 To QtyUpgradeReq
            Obtained = False
            For J = 1 To UpgradesQty
                If PlayerData.Guild.Upgrades(J).IdUpgrade = GuildUpgrades(UpgradeIndex).UpgradeRequired(I) Then
                    Obtained = True
                End If
            Next J
            Call GUpgReq.SetItem(I, _
                        0, _
                        0, _
                        False, _
                        IIf(GuildUpgrades(GuildUpgrades(UpgradeIndex).UpgradeRequired(I)).IconGraph <> 0, GuildUpgrades(GuildUpgrades(UpgradeIndex).UpgradeRequired(I)).IconGraph, GRAPH_UNKNOWN_UPG), _
                        0, _
                        0, _
                        0, _
                        0, _
                        0, _
                        0, _
                        GuildUpgrades(GuildUpgrades(UpgradeIndex).UpgradeRequired(I)).Name, _
                        0, _
                        Obtained)
        Next I
    End If
    
    Exit Sub
End Sub

Private Sub Load_ReqQuest(ByVal UpgradeIndex As Integer)
    Dim I As Integer, J As Integer, k As Integer
    Dim QtyQstReq As Integer
    'Dim QuestQty As Integer
    
    Call GQstReq.Release
    Set GQstReq = Nothing
    
    Call InitializateGQstReq
    
    If ((Not GuildUpgrades(UpgradeIndex).QuestRequired) = -1) Then
        QtyQstReq = 0
    Else
        QtyQstReq = UBound(GuildUpgrades(UpgradeIndex).QuestRequired)
    End If
    
    If QtyQstReq = 0 Then
        Exit Sub
    End If
       
    If Not GQstReq Is Nothing Then
        For I = 1 To QtyQstReq
            Call GQstReq.SetItem(I, _
                        0, _
                        0, _
                        False, _
                        GRAPH_QUEST, _
                        0, _
                        0, _
                        0, _
                        0, _
                        0, _
                        0, _
                        GuildUpgrades(UpgradeIndex).QuestRequired(I).Title, _
                        0, _
                        GuildUpgrades(UpgradeIndex).QuestRequired(I).Obtained)
        Next I
    End If
    
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call GUpgrades.Release
    Call GUpgReq.Release
    Call GQstReq.Release

    Set GUpgrades = Nothing
    Set GUpgReq = Nothing
    Set GQstReq = Nothing
    
End Sub

Private Sub InitializateGUpgReq()

    If GUpgReq Is Nothing Then
        
        Set GUpgReq = New clsGraphicalInventory
        
        Call GUpgReq.Initialize(frmGuildUpgrades.PicGUpgReq, MAX_UPGRADES, , , , 10, , , , , True)
    End If

End Sub

Private Sub InitializateGQstReq()

    If GQstReq Is Nothing Then
        
        Set GQstReq = New clsGraphicalInventory
        
        Call GQstReq.Initialize(frmGuildUpgrades.PicGQstReq, MAX_UPGRADES, , , , 10, , , , , True)
    End If

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
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CloseWindow de frmGuildUpgrades.frm")
End Sub
