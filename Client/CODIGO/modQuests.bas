Attribute VB_Name = "modQuests"
'@Folder("Quest")
Option Explicit

Private Enum eTimeValueReaderAs
    Hours
    Minutes
    Seconds
End Enum

Public Const QUEST_OVERHEADICON As Integer = 25152

'''''''
' QUEST
'''''''

Public Type tQuestObj
    ObjIndex As Integer
    ObjQty As Long
    NpcIndex As Integer
End Type

Public Type tQuestRewards
    gold As Long
    Exp As Long
    ObjsQty As Byte
    Objs() As tQuestObj
End Type

Public Type tQuestNpc
    NpcIndex As Integer
    Quantity As Integer
    Desc As String
End Type

Public Type tQuestFragAlign
    Qty As Integer
    MinLevel As Byte
End Type

Public Type tQuestFrags
    Neutral As tQuestFragAlign
    Ciuda As tQuestFragAlign
    criminal As tQuestFragAlign
    Army As tQuestFragAlign
    Legion As tQuestFragAlign
    MinLevel As Integer
End Type

Public Type tQuestStage

    StarterNpc As tQuestNpc
    EndNpc As tQuestNpc

    ObjsCollectQuantity As Integer
    ObjsCollect() As RequiredObjectListItem
    
    NpcsKillsQuantity As Integer
    NpcKill() As tQuestNpc

    Frags As tQuestFrags
    Rewards As tQuestRewards

End Type

Public Type tQuestStageProgressObjectiveText
    Text As String
    Color As Long
End Type

Public Type tQuestStageProgress
    ObjsCollected As RequiredObjectList
    NpcKilledQty As Integer
    NpcKilled() As Integer
    FragsArmyQty As Integer
    FragsLegionQty As Integer
    FragsNeutralQty As Integer
    RequirementsCompleted As Boolean
    PreCalculatedScreenTextLineQty As Integer
    PreCalculatedScreenText() As tQuestStageProgressObjectiveText
    EndStageNpc As Integer
End Type

Public Type tGuildQuestsStatus
    Id As Integer
    
    'Title As String
    'Desc As String
    
    'Alignment As Integer

    CompletedQuantiy As Integer
    Completed() As Integer
    
    'CurrentQuest As Integer
    CurrentStage As Byte
    
    'CurrentStageProgress As tQuestStage
    CurrentStageProgress As tQuestStageProgress
    
    StartedDateTime As Date
   
End Type

Public Type tCorrelativeQuest
    IdQuest As Integer
End Type


Public Type tQuest
    Id As Integer
    Title As String
    Desc As String
    
    Alignment As Integer
    
    Active As Boolean
    ContributionEarnedFirstTime As Long
    ContributionEarned As Long

    MinLevel As Byte
    MaxLevel As Byte
    
    RepetitionQuantity As Integer
    Duration As Long
    Cooldown As Long
    MinMembers As Byte

    Rewards As tQuestRewards
    StageQuantity As Integer
    Stages() As tQuestStage

    CorrelativesQuantity As Integer
    Correlatives() As tCorrelativeQuest
End Type

Public Sub LoadGuildQuests()
    Call LoadQuests("GuildQuests.dat", GameMetadata.GuildQuests, GameMetadata.GuildQuestsQty)
End Sub

Private Function ReadAsTimeStamp(ByRef TimeValue As String, ByVal TimeValueReaderAs As eTimeValueReaderAs) As Long
    Dim Segments() As String
    Dim SegmentIndex As Integer
    Dim Multiply As Long
    Dim TempValue As Long
    If LenB(TimeValue) = 0 Then
        Err.Raise -1, , "Argument TimeValue is empty on function ReadAsTimeStamp on module: modQuestSystem"
    End If
    
    Segments = Split(TimeValue, ":")
    
    If UBound(Segments) = 0 Then
        ReadAsTimeStamp = Val(TimeValue)
        Exit Function
    End If
    
    ReadAsTimeStamp = 0
    For SegmentIndex = 0 To UBound(Segments)
        TempValue = Val(Segments(SegmentIndex))
        
        Select Case SegmentIndex
            Case 0
                Multiply = 3600
            Case 1
                Multiply = 60
            Case 2
                Multiply = 1
            Case Else
                Multiply = 0
        End Select
        
        ReadAsTimeStamp = ReadAsTimeStamp + _
            (TempValue * Multiply)
    Next

    Select Case TimeValueReaderAs
        Case eTimeValueReaderAs.Hours
            ReadAsTimeStamp = ReadAsTimeStamp / 3600
        Case eTimeValueReaderAs.Minutes
            ReadAsTimeStamp = ReadAsTimeStamp / 60
    End Select

End Function

Public Sub LoadQuests(ByVal fileName As String, ByRef QuestArray() As tQuest, ByRef QuestCount As Integer)
'***************************************************
On Error GoTo ErrorHandler

    Dim Reader As clsIniManager
    Dim QuestsQty As Integer
    Dim I As Integer, J As Integer, k As Integer
    Dim GuildQuestI As Integer
    Dim QuestHeader As String
    Dim ObjDef As String
    Dim StageHeader As String
    Dim aux() As String
    Dim ReadValue As String

    'Cargamos el clsIniReader en memoria
    Set Reader = New clsIniManager

    Call Reader.Initialize(App.path & DAT_PATH & fileName)

    QuestCount = Val(Reader.GetValue("INIT", "QuestsQty"))
        
    If QuestCount = 0 Then
        Exit Sub
    End If
    ReDim QuestArray(1 To QuestCount)

    GuildQuestI = 1
    For I = 1 To QuestCount
        With QuestArray(I)
            QuestHeader = "QUEST" & I
            
            .Id = I
            .Title = Reader.GetValue(QuestHeader, "Title")
            .Desc = Reader.GetValue(QuestHeader, "Desc")
            .ContributionEarned = CLng(Reader.GetValue(QuestHeader, "ContributionEarned"))
            .ContributionEarnedFirstTime = CLng(Reader.GetValue(QuestHeader, "ContributionEarnedFirstTime"))
            .MinLevel = CByte(Val(Reader.GetValue(QuestHeader, "MinLevel")))
            .MaxLevel = CByte(Val(Reader.GetValue(QuestHeader, "MaxLevel")))
            .Active = (Reader.GetValue(QuestHeader, "Active") = "1")
            .RepetitionQuantity = CInt(Val(Reader.GetValue(QuestHeader, "RepetitionQuantity")))
            
            .Duration = ReadAsTimeStamp(Reader.GetValue(QuestHeader, "Time"), Seconds)
            .Cooldown = ReadAsTimeStamp(Reader.GetValue(QuestHeader, "Cooldown"), Seconds)
            
            .MinMembers = CByte(Val(Reader.GetValue(QuestHeader, "MinMembers")))
    
            .Alignment = CInt(Val(Reader.GetValue(QuestHeader, "Alignment")))
            
            .Rewards = modQuests.ReadReward(Reader, QuestHeader)

            .StageQuantity = CInt(Val(Reader.GetValue(QuestHeader, "StageQuantity")))

            .CorrelativesQuantity = CInt(Val(Reader.GetValue(QuestHeader, "CorrelativeQuestsQuantity")))

            If .CorrelativesQuantity > 0 Then
                ReDim .Correlatives(1 To .CorrelativesQuantity)

                For J = 1 To .CorrelativesQuantity
                    .Correlatives(J).IdQuest = CInt(Val(Reader.GetValue(QuestHeader, "CorrelativeQuest" & J)))
                Next J
            End If

            If .StageQuantity > 0 Then
                ReDim .Stages(1 To .StageQuantity)

                For J = 1 To .StageQuantity

                    StageHeader = QuestHeader & "-S" & J

                    With .Stages(J)
                        .StarterNpc.Desc = Reader.GetValue(StageHeader, "StarterNpcDesc")
                        .StarterNpc.NpcIndex = CInt(Val(Reader.GetValue(StageHeader, "StarterNpcIndex")))

                        .EndNpc.Desc = Reader.GetValue(StageHeader, "EndNpcDesc")
                        .EndNpc.NpcIndex = CInt(Val(Reader.GetValue(StageHeader, "EndNpcIndex")))

                        .ObjsCollectQuantity = CInt(Val(Reader.GetValue(StageHeader, "ObjsCollectQuantity")))
                        
                        If .ObjsCollectQuantity > 0 Then
                            ReDim .ObjsCollect(.ObjsCollectQuantity)
                            For k = 0 To .ObjsCollectQuantity - 1
                                aux = Split(Reader.GetValue(StageHeader, "ObjCollect" & (k + 1)), "-")
                                
                                .ObjsCollect(k).ObjIndex = CInt(aux(0))
                                .ObjsCollect(k).RequiredQuantity = CLng(aux(1))

                            Next k
                        End If
                     
                        .NpcsKillsQuantity = CInt(Val(Reader.GetValue(StageHeader, "NpcsKillsQuantity")))
                  
                        If .NpcsKillsQuantity > 0 Then
                            ReDim .NpcKill(1 To .NpcsKillsQuantity)
                            For k = 1 To .NpcsKillsQuantity
                                aux = Split(Reader.GetValue(StageHeader, "NpcKill" & k), "-")
                                .NpcKill(k).NpcIndex = CInt(aux(0))
                                .NpcKill(k).Quantity = CInt(aux(1))
                            Next k
                        End If
                      
                        .Frags.Neutral.Qty = 0
                        .Frags.Ciuda.Qty = 0
                        .Frags.criminal.Qty = 0
                        .Frags.Army.Qty = 0
                        .Frags.Legion.Qty = 0
                        
                        .Frags.MinLevel = 0
                        
                        Call ReadFrag(Reader.GetValue(StageHeader, "Frags"), .Frags.Neutral)
                        
                        Call ReadFrag(Reader.GetValue(StageHeader, "CriminalFrags"), .Frags.criminal)
                        
                        Call ReadFrag(Reader.GetValue(StageHeader, "ArmyFrags"), .Frags.Army)
                        
                        Call ReadFrag(Reader.GetValue(StageHeader, "LegionFrags"), .Frags.Legion)
                        
                        Call ReadFrag(Reader.GetValue(StageHeader, "CiudaFrags"), .Frags.Ciuda)
                        
                        .Frags.MinLevel = CInt(Val(Reader.GetValue(StageHeader, "MinFragLevel")))
                        
                        .Rewards = ReadReward(Reader, StageHeader)
                    End With
                Next J
            End If
        End With
    Next I

    Set Reader = Nothing

Exit Sub

ErrorHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadQuests de modQuestSystem.bas")
End Sub

Public Sub ReadFrag(ByVal ReadValue As String, ByRef QuestFrag As tQuestFragAlign)
    On Error GoTo ErrHandler
    
    If ReadValue = "" Then
        QuestFrag.Qty = 0
        QuestFrag.MinLevel = 0
        Exit Sub
    End If
    
    Dim aux() As String
    
    aux = Split(ReadValue, "-")
    
    QuestFrag.Qty = CInt(aux(0))
    
    If UBound(aux) = 1 Then
        QuestFrag.MinLevel = CByte(aux(1))
    End If
    Exit Sub
    
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ReadFrag de modQuestSystem.bas")
End Sub

Public Function ReadReward(ByRef Reader As clsIniManager, ByVal Header As String) As tQuestRewards
    Dim J As Integer
    Dim ObjDef() As String
    With ReadReward
        .ObjsQty = CInt(Val(Reader.GetValue(Header, "RewardObjs")))
        .gold = CInt(Val(Reader.GetValue(Header, "RewardGold")))
        .Exp = CInt(Val(Reader.GetValue(Header, "RewardExp")))
        If .ObjsQty > 0 Then
            ReDim .Objs(1 To .ObjsQty)
            For J = 1 To .ObjsQty
                ObjDef = Split(Reader.GetValue(Header, "RewardObj" & J), "-")
                .Objs(J).ObjIndex = ObjDef(0)
                .Objs(J).ObjQty = ObjDef(1)
            Next J
        End If
    End With
End Function
Public Function GetQuestRemainingTime() As String
    Dim SecondsRemaining As Long
    SecondsRemaining = DateDiff("s", Now, PlayerData.Guild.Quest.StartedDateTime)
    If SecondsRemaining <= 0 Then
        GetQuestRemainingTime = "00:00:00"
        Exit Function
    End If
    
    GetQuestRemainingTime = modHelperFunctions.SecondsToTimeString(SecondsRemaining)
End Function
Public Function IsQuestObject(ByVal ObjIndex As Integer)
    IsQuestObject = False
    If PlayerData.Guild.Quest.Id <= 0 Then Exit Function
    If PlayerData.Guild.Quest.CurrentStage <= 0 Then Exit Function
    
    Dim I As Integer
    
    For I = 0 To PlayerData.Guild.Quest.CurrentStageProgress.ObjsCollected.ItemsCount - 1
        If PlayerData.Guild.Quest.CurrentStageProgress.ObjsCollected.Items(I).ObjIndex = ObjIndex Then
            IsQuestObject = True
            Exit Function
        End If
    Next I
End Function
Public Sub RefreshObjectives()
    Dim I As Integer
    Dim FontColor As Long
    Dim ObjectiveCompleted As Boolean
    Dim FontCompleted As Long
    Dim FontTimeColor As Long
    Dim FontPending As Long
    Dim LineCount As Long
    
    LineCount = 0

    FontCompleted = RGBA(70, 200, 68, 255)
    FontPending = RGBA(228, 113, 131, 255)
    FontTimeColor = RGBA(210, 227, 122, 255)
        
    If PlayerData.Guild.Quest.Id <= 0 Then Exit Sub
    If PlayerData.Guild.Quest.CurrentStage <= 0 Then Exit Sub
    
    Erase PlayerData.Guild.Quest.CurrentStageProgress.PreCalculatedScreenText
    PlayerData.Guild.Quest.CurrentStageProgress.PreCalculatedScreenTextLineQty = 0
    
    Dim CurrentRequirementCompleted As Boolean
    CurrentRequirementCompleted = True
    
    Dim UpdateFormGuildQuestObjective As Boolean
    
    UpdateFormGuildQuestObjective = frmGuildQuestActive.Visible
    
    If UpdateFormGuildQuestObjective Then
        Call frmGuildQuestActive.QuestObjectives.Clear
    End If
    
    
    With GameMetadata.GuildQuests(PlayerData.Guild.Quest.Id).Stages(PlayerData.Guild.Quest.CurrentStage)
        If .NpcsKillsQuantity > 0 Then
            For I = 1 To PlayerData.Guild.Quest.CurrentStageProgress.NpcKilledQty
                FontColor = IIf(PlayerData.Guild.Quest.CurrentStageProgress.NpcKilled(I) < .NpcKill(I).Quantity, FontPending, FontCompleted)
                Call AddLineToQuestObjectiveText(PlayerData.Guild.Quest.CurrentStageProgress.NpcKilled(I) & "/" & .NpcKill(I).Quantity & " - " & GameMetadata.Npcs(.NpcKill(I).NpcIndex).Name, FontColor)
                
                If UpdateFormGuildQuestObjective Then
                     Call frmGuildQuestActive.QuestObjectives.SetNpcKill(.NpcKill(I).NpcIndex, PlayerData.Guild.Quest.CurrentStageProgress.NpcKilled(I), .NpcKill(I).Quantity, GameMetadata.Npcs(.NpcKill(I).NpcIndex).MiniatureFileName)
                End If
                
                If PlayerData.Guild.Quest.CurrentStageProgress.NpcKilled(I) < .NpcKill(I).Quantity Then CurrentRequirementCompleted = False
            Next I
        End If
        
        If .ObjsCollectQuantity > 0 Then
            'Exceptional for
            For I = 0 To PlayerData.Guild.Quest.CurrentStageProgress.ObjsCollected.ItemsCount - 1
                FontColor = IIf(PlayerData.Guild.Quest.CurrentStageProgress.ObjsCollected.Items(I).Quantity < PlayerData.Guild.Quest.CurrentStageProgress.ObjsCollected.Items(I).RequiredQuantity, FontPending, FontCompleted)
                Call AddLineToQuestObjectiveText(PlayerData.Guild.Quest.CurrentStageProgress.ObjsCollected.Items(I).Quantity & "/" & PlayerData.Guild.Quest.CurrentStageProgress.ObjsCollected.Items(I).RequiredQuantity & " - " & GameMetadata.Objs(PlayerData.Guild.Quest.CurrentStageProgress.ObjsCollected.Items(I).ObjIndex).Name, FontColor)
                
                If UpdateFormGuildQuestObjective Then
                    Call frmGuildQuestActive.QuestObjectives.SetItem(PlayerData.Guild.Quest.CurrentStageProgress.ObjsCollected.Items(I).ObjIndex, PlayerData.Guild.Quest.CurrentStageProgress.ObjsCollected.Items(I).Quantity, PlayerData.Guild.Quest.CurrentStageProgress.ObjsCollected.Items(I).RequiredQuantity)
                End If
            Next I
            If Not modRequiredObjectList.RequiredObjectListIsComplete(PlayerData.Guild.Quest.CurrentStageProgress.ObjsCollected) Then
                CurrentRequirementCompleted = False
            End If
        End If
        
        If .Frags.Army.Qty > 0 Then
            FontColor = IIf(PlayerData.Guild.Quest.CurrentStageProgress.FragsArmyQty < .Frags.Army.Qty, FontPending, FontCompleted)
            Call AddLineToQuestObjectiveText(PlayerData.Guild.Quest.CurrentStageProgress.FragsArmyQty & "/" & .Frags.Army.Qty & " - Armada Real", FontColor)
            If UpdateFormGuildQuestObjective Then
                Call frmGuildQuestActive.QuestObjectives.SetUserKill(eUserKillType.Army, PlayerData.Guild.Quest.CurrentStageProgress.FragsArmyQty, .Frags.Army.Qty)
            End If
            If .Frags.Army.Qty < PlayerData.Guild.Quest.CurrentStageProgress.FragsArmyQty Then CurrentRequirementCompleted = False
        End If
        
        If .Frags.Legion.Qty > 0 Then
            FontColor = IIf(PlayerData.Guild.Quest.CurrentStageProgress.FragsLegionQty < .Frags.Legion.Qty, FontPending, FontCompleted)
            Call AddLineToQuestObjectiveText(PlayerData.Guild.Quest.CurrentStageProgress.FragsLegionQty & "/" & .Frags.Legion.Qty & " - Legionarios", FontColor)
            
            If UpdateFormGuildQuestObjective Then
                Call frmGuildQuestActive.QuestObjectives.SetUserKill(eUserKillType.Legion, PlayerData.Guild.Quest.CurrentStageProgress.FragsLegionQty, .Frags.Legion.Qty)
            End If
            
            If .Frags.Legion.Qty < PlayerData.Guild.Quest.CurrentStageProgress.FragsLegionQty Then CurrentRequirementCompleted = False
        End If
        
        If .Frags.Neutral.Qty > 0 Then
            FontColor = IIf(PlayerData.Guild.Quest.CurrentStageProgress.FragsNeutralQty < .Frags.Neutral.Qty, FontPending, FontCompleted)
            Call AddLineToQuestObjectiveText(PlayerData.Guild.Quest.CurrentStageProgress.FragsNeutralQty & "/" & .Frags.Neutral.Qty & " - Neutrales", FontColor)
            
            If UpdateFormGuildQuestObjective Then
                Call frmGuildQuestActive.QuestObjectives.SetUserKill(eUserKillType.Neutral, PlayerData.Guild.Quest.CurrentStageProgress.FragsNeutralQty, .Frags.Neutral.Qty)
            End If
            If .Frags.Neutral.Qty < PlayerData.Guild.Quest.CurrentStageProgress.FragsNeutralQty Then CurrentRequirementCompleted = False
        End If
        
        PlayerData.Guild.Quest.CurrentStageProgress.RequirementsCompleted = CurrentRequirementCompleted
          
        If .EndNpc.NpcIndex > 0 And CurrentRequirementCompleted Then
            Call AddLineToQuestObjectiveText("Hablar con " & GameMetadata.Npcs(.EndNpc.NpcIndex).Name, FontPending)
            
            If UpdateFormGuildQuestObjective Then
                Call frmGuildQuestActive.QuestObjectives.SetTalk(.EndNpc.NpcIndex, False)
            End If
            
        End If
        
        Call AddLineToQuestObjectiveText("Tiempo: " & modQuests.GetQuestRemainingTime(), FontTimeColor)
        
    End With
    
End Sub

Private Sub AddLineToQuestObjectiveText(ByRef Text As String, ByVal FontColor As Long)
    Dim CurrentIndex As Integer
    With PlayerData.Guild.Quest.CurrentStageProgress
        CurrentIndex = .PreCalculatedScreenTextLineQty + 1
        ReDim Preserve .PreCalculatedScreenText(1 To CurrentIndex)
        
        .PreCalculatedScreenText(CurrentIndex).Text = Text
        .PreCalculatedScreenText(CurrentIndex).Color = FontColor
        
        .PreCalculatedScreenTextLineQty = CurrentIndex
    End With
End Sub


Public Sub AddNewCompletedQuest(ByVal QuestNumber As Integer)
    Dim I As Integer
    Dim Exists As Boolean
    
    If QuestNumber <= 0 Then Exit Sub
    
    For I = 1 To PlayerData.Guild.Quest.CompletedQuantiy
        If PlayerData.Guild.Quest.Completed(I) = QuestNumber Then
            Exists = True
            Exit For
        End If
    Next I
    
    If Not Exists Then
        PlayerData.Guild.Quest.CompletedQuantiy = PlayerData.Guild.Quest.CompletedQuantiy + 1
        ReDim Preserve PlayerData.Guild.Quest.Completed(1 To PlayerData.Guild.Quest.CompletedQuantiy)
        
        PlayerData.Guild.Quest.Completed(PlayerData.Guild.Quest.CompletedQuantiy) = QuestNumber
    End If
    
End Sub

Public Sub CleanCurrentQuestData()
    With PlayerData.Guild.Quest
        .Id = 0
        .CurrentStage = 0
            
        Erase .CurrentStageProgress.NpcKilled
        
        Erase .CurrentStageProgress.ObjsCollected.Items
        .CurrentStageProgress.ObjsCollected.ItemsCount = 0
        .CurrentStageProgress.ObjsCollected.IsComplete = False
        
        .CurrentStageProgress.NpcKilledQty = 0
        
        .CurrentStageProgress.FragsArmyQty = 0
        .CurrentStageProgress.FragsLegionQty = 0
        .CurrentStageProgress.FragsNeutralQty = 0
        .CurrentStageProgress.RequirementsCompleted = False
        
        .CurrentStageProgress.EndStageNpc = 0
        
        .CurrentStageProgress.PreCalculatedScreenTextLineQty = 0
        
        Erase .CurrentStageProgress.PreCalculatedScreenText
    End With
End Sub
