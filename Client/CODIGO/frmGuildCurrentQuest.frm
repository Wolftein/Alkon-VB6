VERSION 5.00
Begin VB.Form frmGuildQuestActive 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4980
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   ScaleHeight     =   332
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   536
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ARGENTUM.AOPictureBox PicRewardsInventory 
      Height          =   495
      Left            =   1185
      TabIndex        =   8
      Top             =   3735
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   873
   End
   Begin ARGENTUM.AOPictureBox PicInventory 
      Height          =   1920
      Left            =   1185
      TabIndex        =   7
      Top             =   1425
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   3387
   End
   Begin VB.TextBox TxtDepositQuantity 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   220
      Left            =   2040
      TabIndex        =   6
      Text            =   "1"
      Top             =   3480
      Width           =   735
   End
   Begin VB.Timer tmrRemainingTime 
      Interval        =   1000
      Left            =   120
      Top             =   840
   End
   Begin ARGENTUM.ucQuestObjectives QuestObjectives 
      Height          =   3340
      Left            =   5040
      TabIndex        =   5
      Top             =   1275
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   4551
   End
   Begin VB.Image ImgShowInventory 
      Height          =   525
      Left            =   4350
      ToolTipText     =   "Abrir Inventario"
      Top             =   1920
      Width           =   600
   End
   Begin VB.Image ImgDepositItem 
      Height          =   300
      Left            =   2880
      ToolTipText     =   "Depositar Item Seleccionado"
      Top             =   3450
      Width           =   300
   End
   Begin VB.Label LblQuestDescription 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "This is the quest description"
      ForeColor       =   &H00FFFFFF&
      Height          =   1890
      Left            =   630
      TabIndex        =   4
      Top             =   1320
      Width           =   3615
   End
   Begin VB.Label lblCurrentStageNumber 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "1 / 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6200
      TabIndex        =   3
      Top             =   360
      Width           =   1095
   End
   Begin VB.Image ImgCancelQuest 
      Height          =   525
      Left            =   3360
      Top             =   4320
      Width           =   1230
   End
   Begin VB.Label LblRewards 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "3000 Experiencia"
      ForeColor       =   &H00FFFFFF&
      Height          =   1245
      Left            =   1200
      TabIndex        =   2
      ToolTipText     =   "Cancelar la misión"
      Top             =   3720
      Width           =   3015
   End
   Begin VB.Label LblRemainingTime 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "00:20:35"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label LblQuestName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre Misión"
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
      Height          =   435
      Left            =   2160
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "frmGuildQuestActive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Guild.Forms")
Option Explicit

Public cButtonCancelMission As clsGraphicalButton
Public cButtonDepositItem As clsGraphicalButton

Public cButtonShowInventory As clsGraphicalButton
Public UserInventory As clsGraphicalInventory
Public RewardsInventory As clsGraphicalInventory
Dim InventoryOpen As Boolean

Public LastButtonPressed As clsGraphicalButton

Private Sub ShowOrHideInventory(ByVal ShowInventory As Boolean)
    Dim GrhPath As String
    GrhPath = DirInterfaces & SELECTED_UI
    
    If Not ShowInventory Then
        frmGuildQuestActive.Picture = LoadPicture(GrhPath & "VentanaGuildActiveQuest.jpg")
        
        Call cButtonShowInventory.Initialize(ImgShowInventory, GrhPath & "BotonIconoGold.jpg", _
                                   GrhPath & "BotonIconoGold.jpg", _
                                   GrhPath & "BotonIconoGold.jpg", Me)
                                   
        ImgShowInventory.ToolTipText = "Abrir Inventario"
        
        
    Else
        frmGuildQuestActive.Picture = LoadPicture(GrhPath & "VentanaGuildActiveQuestDeposit.jpg")

        Call cButtonShowInventory.Initialize(ImgShowInventory, GrhPath & "BotonIconoGoldRed.jpg", _
                                   GrhPath & "BotonIconoGoldRed.jpg", _
                                   GrhPath & "BotonIconoGoldRed.jpg", Me)
                                   
        ImgShowInventory.ToolTipText = "Cerrar Inventario"

    End If
    
    ' Show or hide the inventory controls, and do the oposite with the normal controls
    TxtDepositQuantity.text = 1
    
    PicInventory.Visible = ShowInventory
    TxtDepositQuantity.Visible = ShowInventory
    ImgDepositItem.Visible = ShowInventory
    'ImgDepositIAllItems.Visible = ShowInventory
    InventoryOpen = ShowInventory
        
    LblRewards.Visible = Not ShowInventory
    LblQuestDescription.Visible = Not ShowInventory
    ImgCancelQuest.Visible = Not ShowInventory
    PicRewardsInventory.Visible = Not ShowInventory
    

End Sub

Private Sub Form_Activate()
    Call modQuests.RefreshObjectives
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set UserInventory = Nothing
End Sub

Private Sub Form_Load()

    Call LoadControls

    Call ShowData
    
    Call LoadObjectives
    
    Call ShowOrHideInventory(False)
            
    tmrRemainingTime.Interval = 1000
    tmrRemainingTime.Enabled = True
    
    Call modCustomCursors.SetFormCursorDefault(Me)
End Sub

Private Sub LoadObjectives()
    Call QuestObjectives.Initialize
End Sub

Public Sub ShowData()

    If PlayerData.Guild.Quest.Id <= 0 Then Exit Sub

    LblQuestName.Caption = GameMetadata.GuildQuests(PlayerData.Guild.Quest.Id).Title
    
    LblQuestDescription.Caption = GameMetadata.GuildQuests(PlayerData.Guild.Quest.Id).Desc
    
    lblCurrentStageNumber.Caption = GetQuestStageNumberText()
    'LblObjectives.Caption = GetQuestObjectivesText()
    LblRewards.Caption = GetQuestRewardsText()
    
    ' Calculate the remaining time for the first time
    Call CalculateRemainingTime
    
    ' Loads the current user inventory into the small inventory we have in the item's deposit section
    ' Note: Only the items that CAN be deposited into the active quest requirements will be shown
    Call LoadUserInventory
    
    Call ShowOrHideInventory(False)
    
    Call DrawRewardsInventory
    
    ' Enable the timer
    tmrRemainingTime.Enabled = True
    
End Sub



Private Sub ImgCancelQuest_Click()
    If Not cButtonCancelMission.IsEnabled Then
        Exit Sub
    End If
    Call WriteGuildQuest(0)
End Sub

Public Function GetQuestObjectivesText(Optional ByVal ShowEndNpcText As Boolean = True) As String
    Dim I As Integer
    
    With GameMetadata.GuildQuests(PlayerData.Guild.Quest.Id).Stages(PlayerData.Guild.Quest.CurrentStage)
        If .NpcsKillsQuantity > 0 Then
            For I = 1 To .NpcsKillsQuantity
                GetQuestObjectivesText = GetQuestObjectivesText & PlayerData.Guild.Quest.CurrentStageProgress.NpcKilled(I) & " / " & .NpcKill(I).Quantity & " - Matar " & .NpcKill(I).Quantity & " " & GameMetadata.Npcs(.NpcKill(I).NpcIndex).Name & vbCrLf
            Next I
        End If
        
        If .ObjsCollectQuantity > 0 Then
            For I = 0 To PlayerData.Guild.Quest.CurrentStageProgress.ObjsCollected.ItemsCount - 1
                GetQuestObjectivesText = GetQuestObjectivesText & PlayerData.Guild.Quest.CurrentStageProgress.ObjsCollected.Items(I).Quantity & " / " & PlayerData.Guild.Quest.CurrentStageProgress.ObjsCollected.Items(I).RequiredQuantity _
                                       & " - Obtener " & PlayerData.Guild.Quest.CurrentStageProgress.ObjsCollected.Items(I).RequiredQuantity & " " & PlayerData.Guild.Quest.CurrentStageProgress.ObjsCollected.Items(I).ObjIndex & vbCrLf
            Next I
        End If
                
        If .Frags.Neutral.Qty > 0 Then
            GetQuestObjectivesText = GetQuestObjectivesText & PlayerData.Guild.Quest.CurrentStageProgress.FragsNeutralQty & " / " & .Frags.Neutral.Qty & " - Matar " & .Frags.Neutral.Qty & " neutrales" & vbCrLf
        End If
        
        If .Frags.Army.Qty > 0 Then
            GetQuestObjectivesText = GetQuestObjectivesText & PlayerData.Guild.Quest.CurrentStageProgress.FragsArmyQty & " / " & .Frags.Army.Qty & " - Matar " & .Frags.Army.Qty & " miembros de la Armada Real" & vbCrLf
        End If
                
        If .Frags.Legion.Qty > 0 Then
            GetQuestObjectivesText = GetQuestObjectivesText & PlayerData.Guild.Quest.CurrentStageProgress.FragsLegionQty & " / " & .Frags.Legion.Qty & " - Matar " & .Frags.Legion.Qty & " miembros de la Legion del Mal" & vbCrLf
        End If
        
        If .EndNpc.NpcIndex > 0 And ShowEndNpcText Then
            GetQuestObjectivesText = GetQuestObjectivesText & "Hablar con " & GameMetadata.Npcs(.EndNpc.NpcIndex).Name & vbCrLf
        End If
    
    
    End With
End Function

Public Function CalculateRemainingTime()
  
    Dim TimeRemaining As String
    
    ' If form is not visible/active, then we stop the timer as we don't need to calculate anything
    If frmGuildQuestActive.Visible = False Then
        LblRemainingTime.Caption = "00:00:00"
        tmrRemainingTime.Enabled = False
        Exit Function
    End If
    
    TimeRemaining = modQuests.GetQuestRemainingTime()
    
    ' Just in case something happened with the quest stop notification, if the remaining time is < than
    ' 0 seconds then we stop the timer and exit.
    If TimeRemaining = "00:00:00" Then
        LblRemainingTime.Caption = "00:00:00"
        tmrRemainingTime.Enabled = False
        Exit Function
    End If
    
    ' Show the remaining time.
    LblRemainingTime.Caption = TimeRemaining
    
End Function

Public Function GetQuestStageNumberText() As String

    GetQuestStageNumberText = PlayerData.Guild.Quest.CurrentStage & " / " & GameMetadata.GuildQuests(PlayerData.Guild.Quest.Id).StageQuantity
    
End Function

Public Function GetQuestRewardsText() As String

    With GameMetadata.GuildQuests(PlayerData.Guild.Quest.Id).Stages(PlayerData.Guild.Quest.CurrentStage)
        If .Rewards.Exp > 0 Then
            GetQuestRewardsText = GetQuestRewardsText & .Rewards.Exp & " de Experiencia" & vbCrLf
        End If
        
        If .Rewards.gold > 0 Then
            GetQuestRewardsText = GetQuestRewardsText & .Rewards.gold & " de Oro" & vbCrLf
        End If
        
        If .Rewards.ObjsQty > 0 Then
        
            Dim I As Integer
            
            For I = 1 To .Rewards.ObjsQty
                GetQuestRewardsText = GetQuestRewardsText & .Rewards.Objs(I).ObjQty & " " & GameMetadata.Objs(.Rewards.Objs(I).ObjIndex).Name & vbCrLf
            Next I
            
        End If
        
    End With

End Function

Private Sub LoadControls()
    Set UserInventory = New clsGraphicalInventory
    Set RewardsInventory = New clsGraphicalInventory
    Set cButtonCancelMission = New clsGraphicalButton
    Set LastButtonPressed = New clsGraphicalButton
    Set cButtonDepositItem = New clsGraphicalButton
    'Set cButtonDepositAllItems = New clsGraphicalButton
    Set cButtonShowInventory = New clsGraphicalButton
    
        
    Call UserInventory.Initialize(PicInventory, Inventario.MaxObjs, , , , , , , , , True, _
                                  eMoveType.None)
                                  
    Call RewardsInventory.Initialize(PicRewardsInventory, 5, , , , , , , , , True, _
                                  eMoveType.None)
    
    Dim GrhPath As String
    GrhPath = DirInterfaces & SELECTED_UI
    
    Me.Picture = LoadPicture(GrhPath & "VentanaGuildActiveQuest.jpg")
    
    Call cButtonCancelMission.Initialize(ImgCancelQuest, GrhPath & "BotonCancelar.jpg", _
                                         GrhPath & "BotonCancelar.jpg", _
                                         GrhPath & "BotonCancelar.jpg", Me)
                                    
                                    
    Call cButtonDepositItem.Initialize(ImgDepositItem, GrhPath & "BotonFlechaDerecha.jpg", _
                                       GrhPath & "BotonFlechaDerecha.jpg", _
                                       GrhPath & "BotonFlechaDerecha.jpg", Me, _
                                       GrhPath & "BotonFlechaDerecha_Disabled.jpg")
                                       
    Call cButtonShowInventory.Initialize(ImgShowInventory, GrhPath & "BotonIconoGold.jpg", _
                                       GrhPath & "BotonIconoGold.jpg", _
                                       GrhPath & "BotonIconoGold.jpg", Me)
                                       
    'Call cButtonDepositAllItems.Initialize(ImgDepositIAllItems, GrhPath & "BotonFlechaDobleDerecha.jpg", _
                                       GrhPath & "BotonFlechaDobleDerecha.jpg", _
                                       GrhPath & "BotonFlechaDobleDerecha.jpg", Me)
                                        
    
    Call cButtonCancelMission.EnableButton(PlayerData.Guild.IdRolOwn = ID_ROLE_LEADER)
  
End Sub

'Private Sub ImgDepositIAllItems_Click()
'    Dim Slot As Integer
'    Dim ObjectIndex As Integer
'    Dim ItemIndex As Integer
'    For Slot = 1 To Inventario.MaxObjs
'        ObjectIndex = Inventario.ObjIndex(Slot)
'        If ObjectIndex > 0 Then
'            For ItemIndex = 0 To PlayerData.Guild.Quest.CurrentStageProgress.ObjsCollected.ItemsCount - 1
'                If PlayerData.Guild.Quest.CurrentStageProgress.ObjsCollected.Items(ItemIndex).ObjIndex = ObjectIndex Then
'                    Call Protocol.WriteGuildQuestAddObject(Slot, Inventario.Amount(Slot))
'                End If
'            Next ItemIndex
'        End If
'    Next Slot
'End Sub

Private Sub ImgDepositItem_Click()
    Dim Slot As Byte
    
    Slot = UserInventory.SelectedItem
    
    If Slot = 0 Or Val(TxtDepositQuantity.text) < 1 Then Exit Sub
    
    Call Protocol.WriteGuildQuestAddObject(Slot, Val(TxtDepositQuantity.text))
End Sub

Private Sub ImgShowInventory_Click()
    Call ShowOrHideInventory(InventoryOpen = False)
End Sub

Private Sub tmrRemainingTime_Timer()
    Call CalculateRemainingTime
End Sub
Private Sub PicInventory_Click()
    On Error GoTo ErrHandler

    If UserInventory.SelectedItem = 0 Then Exit Sub
   
    TxtDepositQuantity.text = 0
    
    Call cButtonDepositItem.EnableButton(UserInventory.CanUse(UserInventory.SelectedItem))
    
    Exit Sub
  
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub PicInventory_Click de frmGuildQuestActive.frm")
End Sub

Private Sub PicInventory_DblClick()
    On Error GoTo ErrHandler
        
    If UserInventory.SelectedItem = 0 Then Exit Sub
    If UserInventory.CanUse(UserInventory.SelectedItem) = False Then Exit Sub
    
    Call Protocol.WriteGuildQuestAddObject(UserInventory.SelectedItem, 1)
    
    Exit Sub
  
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub PicInventory_DblClick de frmGuildQuestActive.frm")
End Sub
Public Sub LoadUserInventory()
    Dim I As Long
                                
    For I = 1 To Inventario.MaxObjs
        With Inventario
            Call UserInventory.SetItem(I, .ObjIndex(I), _
                                       .Amount(I), .Equipped(I), .GrhIndex(I), _
                                       .OBJType(I), .MaxHit(I), .MinHit(I), .MaxDef(I), .MinDef(I), _
                                       .Valor(I), .ItemName(I), 0, modQuests.IsQuestObject(.ObjIndex(I)))
        End With
    Next I
End Sub

Private Function IsItemRequired(ByVal ObjNumber As Integer) As Boolean
    
    Dim I As Integer
    
    If PlayerData.Guild.Quest.Id = 0 Or PlayerData.Guild.Quest.CurrentStage = 0 Then Exit Function
     
    With GameMetadata.GuildQuests(PlayerData.Guild.Quest.Id).Stages(PlayerData.Guild.Quest.CurrentStage)
        For I = 0 To .ObjsCollectQuantity - 1
            If .ObjsCollect(I).ObjIndex = ObjNumber Then
                IsItemRequired = True
                Exit Function
            End If
        Next I
    End With
    
    IsItemRequired = False

End Function


Private Sub TxtDepositQuantity_Change()
On Error GoTo ErrHandler

    Dim DepositQuantity As Long
    
    DepositQuantity = Val(TxtDepositQuantity.text)
  
    If DepositQuantity < 1 Then
        TxtDepositQuantity.text = 1
    End If
    
    If DepositQuantity > MAX_INVENTORY_OBJS Then
        TxtDepositQuantity.text = MAX_INVENTORY_OBJS
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub TxtDepositQuantity_Change de frmGuildQuestActive.frm")
End Sub

Private Sub TxtDepositQuantity_KeyPress(KeyAscii As Integer)
On Error GoTo ErrHandler

    Call modHelperFunctions.IsNumericInputKeyPressValid(KeyAscii, True)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub TxtDepositQuantity_KeyPress de frmGuildQuestActive.frm")
End Sub

Private Sub DrawRewardsInventory()
    Dim I As Integer
    Dim SlotToUse As Integer
    Dim GuildId As Integer

    Call RewardsInventory.ClearAllSlots
    
    With GameMetadata.GuildQuests(PlayerData.Guild.Quest.Id)
        If .Rewards.ObjsQty > 0 Then
            For I = 1 To .Rewards.ObjsQty
                Call RewardsInventory.SetItem(I, .Rewards.Objs(I).ObjIndex, _
                                       .Rewards.Objs(I).ObjQty, False, _
                                       GameMetadata.Objs(.Rewards.Objs(I).ObjIndex).GrhIndex, 0, _
                                       0, 0, 0, 0, 0, _
                                       GameMetadata.Objs(.Rewards.Objs(I).ObjIndex).Name, 0, True)
                SlotToUse = SlotToUse + 1
            Next I
        End If
        
        ' Add Gold
        If .Rewards.gold > 0 Then
            SlotToUse = SlotToUse + 1
            Call RewardsInventory.SetItem(SlotToUse, 0, _
                                               .Rewards.gold, False, _
                                               GameMetadata.Objs(1).GrhIndex, 1, _
                                               0, 0, 0, 0, 0, _
                                               "Monedas de Oro", 0, True)
        End If
        
        ' Add Exp
        If .Rewards.Exp > 0 Then
            SlotToUse = SlotToUse + 1
            Call RewardsInventory.SetItem(SlotToUse, 0, _
                                           .Rewards.Exp, False, _
                                           GameMetadata.Objs(1).GrhIndex, 0, _
                                           0, 0, 0, 0, 0, _
                                           "Puntos de Experiencia", 0, True)
        End If
        
        ' Add Contribution
        If .ContributionEarned > 0 Or .ContributionEarnedFirstTime Then
            SlotToUse = SlotToUse + 1
            Call RewardsInventory.SetItem(SlotToUse, 0, _
                                       .ContributionEarnedFirstTime, False, _
                                       GameMetadata.Objs(1).GrhIndex, 0, _
                                       0, 0, 0, 0, 0, _
                                       "Puntos de Contribución", 0, True)
                                       
        End If
        
        
    End With

End Sub
