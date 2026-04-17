Attribute VB_Name = "modQuestSystem"
'@Folder("Quest")
Option Explicit
Private Enum eTimeValueReaderAs
    Hours
    Minutes
    Seconds
End Enum
Public Sub ReloadCurrentQuests()
On Error GoTo ErrHandler

    Dim I As Integer
    Dim J As Integer
    Dim k As Integer
    Dim L As Integer
    Dim Rest As Long
    Dim QuestOldState As tCurrentQuest
    
    For I = 1 To MaxGuildQty
        'for each guild
        If modQuestSystem.GuildHasQuest(I) Then
            With GuildList(I)
                QuestOldState = .CurrentQuest
                
                'look for quest
                For J = 1 To UBound(GuildQuestList)
                    If GuildQuestList(J).Id = .CurrentQuest.IdQuest Then
                    
                        'set stage initial state for update NPC and OBJ arrays
                        Call modQuestSystem.ChangeGuildQuestStage(I, J, .CurrentQuest.CurrentStage, False)
                        
                        'and then restore from old state
                        .CurrentQuest.CurrentFrags.Army = QuestOldState.CurrentFrags.Army
                        .CurrentQuest.CurrentFrags.Legion = QuestOldState.CurrentFrags.Legion
                        .CurrentQuest.CurrentFrags.Neutral = QuestOldState.CurrentFrags.Neutral
                        
                        For k = 1 To .CurrentQuest.CurrentNpcKillsQuantity
                            For L = 1 To QuestOldState.CurrentNpcKillsQuantity
                                If .CurrentQuest.CurrentNpcKills(k).NpcIndex = QuestOldState.CurrentNpcKills(L).NpcIndex Then
                                    .CurrentQuest.CurrentNpcKills(k).Quantity = QuestOldState.CurrentNpcKills(L).Quantity
                                End If
                            Next L
                        Next k
                        
                        'Excepcion aca, la lista de objetos solo deberia iterarse en el modulo modRequiredObjecList
                        For k = 0 To QuestOldState.CurrentObjectList.ItemsCount - 1
                            Call modRequiredObjectList.RequiredObjectListTryAdd(.CurrentQuest.CurrentObjectList _
                                , QuestOldState.CurrentObjectList.Items(k).ObjIndex, _
                                QuestOldState.CurrentObjectList.Items(k).Quantity, Rest)
                        Next k
                    End If
                Next J
            
            End With
            
            Call TryGuildQuestStageFinished(I)
        End If
    Next I

Exit Sub
ErrHandler:
     Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ReloadCurrentQuests de modQuestSystem.bas")
End Sub
Public Sub LoadGuildQuests()
    Call modQuestSystem.LoadQuests("GuildQuests.dat", GuildQuestList)
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

Public Sub LoadQuests(ByVal FileName As String, ByRef QuestArray() As tQuest)
'***************************************************
On Error GoTo ErrorHandler

    Dim Reader As clsIniManager
    Dim QuestsQty As Integer
    Dim I As Integer, J As Integer, k As Integer
    Dim GuildQuestI As Integer
    Dim QuestHeader As String
    Dim ObjDef As String
    Dim StageHeader As String
    Dim Aux() As String
    Dim ReadValue As String

    'Cargamos el clsIniReader en memoria
    Set Reader = New clsIniManager

    Call Reader.Initialize(DatPath & FileName)

    QuestsQty = Val(Reader.GetValue("INIT", "QuestsQty"))
    
    If QuestsQty = 0 Then
        Exit Sub
    End If
    ReDim QuestArray(1 To QuestsQty)

    GuildQuestI = 1
    For I = 1 To QuestsQty
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
            
            .Rewards = modQuestSystem.ReadReward(Reader, QuestHeader)

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
                                
                                .ObjsCollect(k).ObjIndex = CInt(Aux(0))
                                .ObjsCollect(k).RequiredQuantity = CLng(aux(1))
                            Next k
                        End If
                     

                        .NpcsKillsQuantity = CInt(Val(Reader.GetValue(StageHeader, "NpcsKillsQuantity")))
                  
                        If .NpcsKillsQuantity > 0 Then
                            ReDim .NpcKill(1 To .NpcsKillsQuantity)
                            For k = 1 To .NpcsKillsQuantity
                                Aux = Split(Reader.GetValue(StageHeader, "NpcKill" & k), "-")
                                .NpcKill(k).NpcIndex = CInt(Aux(0))
                                .NpcKill(k).Quantity = CInt(Aux(1))
                            Next k
                        End If
                      
                        .Frags.Neutral.Qty = 0
                        .Frags.Army.Qty = 0
                        .Frags.Legion.Qty = 0
                        
                        .Frags.MinLevel = 0
                        
                        Call ReadFrag(Reader.GetValue(StageHeader, "Frags"), .Frags.Neutral)
                                            
                        Call ReadFrag(Reader.GetValue(StageHeader, "ArmyFrags"), .Frags.Army)
                        
                        Call ReadFrag(Reader.GetValue(StageHeader, "LegionFrags"), .Frags.Legion)
                        
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
    
    Dim Aux() As String
    
    Aux = Split(ReadValue, "-")
    
    QuestFrag.Qty = CInt(Aux(0))
    
    If UBound(Aux) = 1 Then
        QuestFrag.MinLevel = CByte(Aux(1))
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
        .Gold = CInt(Val(Reader.GetValue(Header, "RewardGold")))
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

Public Function GuildHasQuest(ByVal GuildIndex As Integer) As Boolean
    On Error GoTo ErrHandler

    If GuildIndex = 0 Then
        GuildHasQuest = False
        Exit Function
    End If

    If GuildList(GuildIndex).CurrentQuest.IdQuest = 0 Then
        GuildHasQuest = False
        Exit Function
    End If
    
    If GuildList(GuildIndex).CurrentQuest.CurrentStage = 0 Then
        GuildHasQuest = False
        Exit Function
    End If

    GuildHasQuest = True

    Exit Function

ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GuildHasQuest de modQuestSystem.bas")
End Function

Public Sub GuildQuestUpdateStatus(ByVal GuildIndex As Integer, ByVal UserIndex As Integer, ByVal TargetIndex As Integer, Requirement As eQuestRequirement, ByVal RequirementIndex As Integer, Optional ByVal Quantity As Integer = 1)
On Error GoTo ErrHandler

    If GuildIndex = 0 Then Exit Sub

    Dim I As Integer
    Dim CurrentStage As Integer
    Dim QuestId As Integer
    Dim ShouldCheckStageStatus As Boolean

    QuestId = GuildList(GuildIndex).CurrentQuest.IdQuest
    
    If GuildList(GuildIndex).CurrentQuest.StageIsCompleted Then
        Exit Sub
    End If
    
    If modQuestSystem.IsQuestOverdue(GuildIndex) Then
        Call modQuestSystem.CancelCurrentGuildQuest(GuildIndex, False)
        Exit Sub
    End If

    CurrentStage = GuildList(GuildIndex).CurrentQuest.CurrentStage
    
    ShouldCheckStageStatus = False
    With GuildList(GuildIndex).CurrentQuest

        Select Case Requirement
            Case eQuestRequirement.NpcKill
                Dim NpcNumber As Integer
                
                For I = 1 To .CurrentNpcKillsQuantity
                    If .CurrentNpcKills(I).NpcIndex = RequirementIndex And GuildQuestList(QuestId).Stages(CurrentStage).NpcKill(I).Quantity > .CurrentNpcKills(I).Quantity Then
                        ShouldCheckStageStatus = True
                        NpcNumber = .CurrentNpcKills(I).NpcIndex
                        .CurrentNpcKills(I).Quantity = .CurrentNpcKills(I).Quantity + Quantity
                        
                        Exit For
                    End If
                Next I
                             
                If ShouldCheckStageStatus = True And I > 0 Then
                    Call NotifiyGuildMembersQuestUpdateStatus_NpcKill(GuildIndex, QuestId, CurrentStage, NpcNumber, I, .CurrentNpcKills(I).Quantity)
                End If

            Case eQuestRequirement.UserKill
                Dim TargetLevel As Byte

                ' We can't count this death
                If UserList(TargetIndex).Guild.GuildIndex = GuildIndex Then
                    Exit Sub
                End If

                TargetLevel = UserList(TargetIndex).Stats.ELV

                If GuildQuestList(QuestId).Stages(CurrentStage).Frags.Neutral.Qty > 0 Then
                    If CheckFragLevel(eQuestUserAlign.Neutral, QuestId, CurrentStage, TargetLevel) Then
                        If .CurrentFrags.Neutral.Qty < GuildQuestList(QuestId).Stages(CurrentStage).Frags.Neutral.Qty Then
                            .CurrentFrags.Neutral.Qty = .CurrentFrags.Neutral.Qty + 1
                            ShouldCheckStageStatus = True
                        End If
                    End If
                End If

                If GuildQuestList(QuestId).Stages(CurrentStage).Frags.Army.Qty > 0 Then
                    If UserList(TargetIndex).Faccion.ArmadaReal = 1 Then
                        If CheckFragLevel(eQuestUserAlign.Army, QuestId, CurrentStage, TargetLevel) Then
                            If .CurrentFrags.Army.Qty < GuildQuestList(QuestId).Stages(CurrentStage).Frags.Army.Qty Then
                                .CurrentFrags.Army.Qty = .CurrentFrags.Army.Qty + 1
                                ShouldCheckStageStatus = True
                            End If
                        End If
                    End If
                End If

                If GuildQuestList(QuestId).Stages(CurrentStage).Frags.Legion.Qty > 0 Then
                    If UserList(TargetIndex).Faccion.FuerzasCaos = 1 Then
                        If CheckFragLevel(eQuestUserAlign.Legion, QuestId, CurrentStage, TargetLevel) Then
                            If .CurrentFrags.Legion.Qty < GuildQuestList(QuestId).Stages(CurrentStage).Frags.Legion.Qty Then
                                .CurrentFrags.Legion.Qty = .CurrentFrags.Legion.Qty + 1
                                ShouldCheckStageStatus = True
                            End If
                        End If
                    End If
                End If
                
                If ShouldCheckStageStatus = True Then
                    Call NotifiyGuildMembersQuestUpdateStatus_UserKill(GuildIndex, QuestId, CurrentStage, .CurrentFrags.Neutral.Qty, .CurrentFrags.Army.Qty, .CurrentFrags.Legion.Qty)
                End If
        End Select

    End With
    
    If Not ShouldCheckStageStatus Then
        Exit Sub
    End If
    
    If TryGuildQuestStageFinished(GuildIndex) Then
        Call NotifyGuildQuestIsFinished(GuildIndex)
    End If

    Exit Sub
      
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GuildQuestUpdateStatus de modQuestSystem.bas")
End Sub

Public Function CheckFragLevel(ByVal QuestAlign As eQuestUserAlign, ByVal QuestIndex As Integer, ByVal StageIndex As Integer, ByVal UserLevel As Byte) As Boolean
    On Error GoTo ErrHandler
    
    Dim FragMinLevel As Integer
    
    Select Case QuestAlign
        Case eQuestUserAlign.Army
            FragMinLevel = GuildQuestList(QuestIndex).Stages(StageIndex).Frags.Army.MinLevel
        Case eQuestUserAlign.Legion
            FragMinLevel = GuildQuestList(QuestIndex).Stages(StageIndex).Frags.Legion.MinLevel
        Case eQuestUserAlign.Neutral
            FragMinLevel = GuildQuestList(QuestIndex).Stages(StageIndex).Frags.Neutral.MinLevel
    End Select
    
    If FragMinLevel = 0 Then
        FragMinLevel = GuildQuestList(QuestIndex).Stages(StageIndex).Frags.MinLevel
    End If
    
    If FragMinLevel = 0 Then
        CheckFragLevel = True
    Else
        CheckFragLevel = (UserLevel >= FragMinLevel)
    End If
    
     Exit Function
     
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function CheckFragLevel de modQuestSystem.bas")
End Function
Public Sub GuildQuestAddObject(ByVal UserIndex As Integer, ByVal InventorySlot As Byte, ByVal Quantity As Long)
    On Error GoTo ErrHandler
    Dim GuildIndex As Integer
    GuildIndex = UserList(UserIndex).Guild.GuildIndex
    
    If Not modQuestSystem.GuildHasQuest(GuildIndex) Then
        Exit Sub
    End If
    
    If modQuestSystem.IsQuestOverdue(GuildIndex) Then
        Call modQuestSystem.CancelCurrentGuildQuest(GuildIndex, False)
        Exit Sub
    End If
    
    Dim InventoryObject As UserOBJ
    
    InventoryObject = InvUsuario.GetUserInvItem(UserIndex, InventorySlot)
    
    If InventoryObject.ObjIndex <= 0 Then
        Exit Sub
    End If
    
    If Quantity < 1 Then Exit Sub
    
    If InventoryObject.Amount < Quantity Then
        Quantity = InventoryObject.Amount
    End If
    
    Dim Rest As Long
    Dim ItemsToRemove As Long
    
    With GuildList(GuildIndex).CurrentQuest
        'the selected item is not required
        If Not modRequiredObjectList.RequiredObjectListTryAdd(.CurrentObjectList, InventoryObject.ObjIndex, Quantity, Rest) Then
            Exit Sub
        End If
        
        If Rest > 0 Then
            ItemsToRemove = Quantity - Rest
        Else
            ItemsToRemove = Quantity
        End If
        
        Call InvUsuario.QuitarUserInvItem(UserIndex, InventorySlot, ItemsToRemove)
        Call NotifiyGuildMembersQuestUpdateStatus_ObjCollect(GuildIndex, .IdQuest, .CurrentStage, InventoryObject.ObjIndex, ItemsToRemove)
    End With
    
    If TryGuildQuestStageFinished(GuildIndex) Then
        Call NotifyGuildQuestIsFinished(GuildIndex)
    End If
    
    Exit Sub
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GuildQuestAddObject de modQuestSystem.bas")
End Sub

Public Sub NotifiyGuildMembersQuestUpdateStatus_NpcKill(ByVal GuildIndex As Integer, ByVal QuestId As Integer, StageNumber As Integer, ByVal NpcNumber As Integer, ByVal RequirementIndex As Integer, ByVal Amount As Integer)
    On Error GoTo ErrHandler
    
    Dim I As Integer
    
    With GuildList(GuildIndex)
        If GuildLastMemberOnline(GuildIndex) > 0 Then
            For I = 1 To GuildLastMemberOnline(GuildIndex)
                Call WriteGuildQuestUpdateReqStatus_NpcKill(.OnlineMembers(I).MemberUserIndex, QuestId, StageNumber, NpcNumber, RequirementIndex, Amount)
            Next I
        End If
    End With
    
    Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function NotifiyGuildMembersQuestUpdateStatus de modQuestSystem.bas")
End Sub

Public Sub NotifiyGuildMembersQuestUpdateStatus_ObjCollect(ByVal GuildIndex As Integer, ByVal QuestId As Integer, StageNumber As Integer, ByVal ObjectIndex As Integer, ByVal Quantity As Integer)
    On Error GoTo ErrHandler
    
    Dim I As Integer
    
    With GuildList(GuildIndex)
        If GuildLastMemberOnline(GuildIndex) > 0 Then
            For I = 1 To GuildLastMemberOnline(GuildIndex)
                Call WriteGuildQuestUpdateReqStatus_ObjCollect(.OnlineMembers(I).MemberUserIndex, QuestId, StageNumber, ObjectIndex, Quantity)
            Next I
        End If
    End With
    
    Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function NotifiyGuildMembersQuestUpdateStatus de modQuestSystem.bas")
End Sub

Public Sub NotifiyGuildMembersQuestUpdateStatus_UserKill(ByVal GuildIndex As Integer, ByVal QuestId As Integer, StageNumber As Integer, ByVal NeutralFrags As Integer, ByVal ArmyFrags As Integer, ByVal LegionFrags As Integer)
    On Error GoTo ErrHandler
    
    Dim I As Integer
    
    With GuildList(GuildIndex)
        If GuildLastMemberOnline(GuildIndex) > 0 Then
            For I = 1 To GuildLastMemberOnline(GuildIndex)
                Call WriteGuildQuestUpdateReqStatus_UserKill(.OnlineMembers(I).MemberUserIndex, QuestId, StageNumber, NeutralFrags, ArmyFrags, LegionFrags)
            Next I
        End If
    End With
    
    Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function NotifiyGuildMembersQuestUpdateStatus de modQuestSystem.bas")
End Sub

Public Sub NotifiyGuildMembersQuestUpdateStatus(ByVal GuildIndex As Integer, ByVal QuestId As Integer, StageNumber As Integer, ByVal Requirement As Integer, ByVal ExtraInfo As Integer, ByVal Quantity As Integer)
    On Error GoTo ErrHandler
    
    Dim I As Integer
    
    With GuildList(GuildIndex)
        If GuildLastMemberOnline(GuildIndex) > 0 Then
            For I = 1 To GuildLastMemberOnline(GuildIndex)
                Call WriteGuildQuestUpdateStatus(.OnlineMembers(I).MemberUserIndex, QuestId, StageNumber, Requirement, ExtraInfo, Quantity)
            Next I
        End If
    End With
    
    Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function NotifiyGuildMembersQuestUpdateStatus de modQuestSystem.bas")
End Sub

Public Function TryGuildQuestStageFinished(ByVal GuildIndex As Integer) As Boolean
    On Error GoTo ErrHandler

    Dim I As Integer, J As Integer

    With GuildList(GuildIndex).CurrentQuest

        For I = 1 To .CurrentNpcKillsQuantity
            For J = 1 To GuildQuestList(.IdQuest).Stages(.CurrentStage).NpcsKillsQuantity
                If .CurrentNpcKills(I).NpcIndex = GuildQuestList(.IdQuest).Stages(.CurrentStage).NpcKill(J).NpcIndex Then
                    If .CurrentNpcKills(I).Quantity < GuildQuestList(.IdQuest).Stages(.CurrentStage).NpcKill(J).Quantity Then
                        TryGuildQuestStageFinished = False
                        Exit Function
                    End If
                End If
            Next J
        Next I
        
        If GuildQuestList(.IdQuest).Stages(.CurrentStage).ObjsCollectQuantity > 0 Then
            If Not modRequiredObjectList.RequiredObjectListIsComplete(.CurrentObjectList) Then
                    TryGuildQuestStageFinished = False
                Exit Function
            End If
        End If

        If .CurrentFrags.Army.Qty < GuildQuestList(.IdQuest).Stages(.CurrentStage).Frags.Army.Qty Then
            TryGuildQuestStageFinished = False
            Exit Function
        End If
        
        If .CurrentFrags.Legion.Qty < GuildQuestList(.IdQuest).Stages(.CurrentStage).Frags.Legion.Qty Then
            TryGuildQuestStageFinished = False
            Exit Function
        End If

        If .CurrentFrags.Neutral.Qty < GuildQuestList(.IdQuest).Stages(.CurrentStage).Frags.Neutral.Qty Then
            TryGuildQuestStageFinished = False
            Exit Function
        End If

        .StageIsCompleted = True
        
        TryGuildQuestStageFinished = True

    End With

    Exit Function

ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function IsGuildQuestStageFinished de modQuestSystem.bas")
End Function

Public Function CheckIfCurrentStageIsCompleted(ByVal GuildIndex As Integer) As Boolean
On Error GoTo ErrHandler

    If GuildHasQuest(GuildIndex) Then
        If GuildList(GuildIndex).CurrentQuest.StageIsCompleted Then
            CheckIfCurrentStageIsCompleted = True
            Exit Function
        End If
    End If

    CheckIfCurrentStageIsCompleted = False

    Exit Function

ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function CheckIfCurrentStageIsCompleted de modQuestSystem.bas")
End Function

Public Function CheckIfCanStartNextStage(ByVal GuildIndex As Integer) As Boolean
On Error GoTo ErrHandler

    If GuildHasQuest(GuildIndex) Then
        If GuildList(GuildIndex).CurrentQuest.CanStartNextStage Then
            CheckIfCanStartNextStage = True
            Exit Function
        End If
    End If

    CheckIfCanStartNextStage = False
    Exit Function

ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function CheckIfCanStartNextStage de modQuestSystem.bas")
End Function

Public Function GetQuestNpcEndIndex(ByVal GuildIndex As Integer) As Integer
    GetQuestNpcEndIndex = GuildQuestList(GuildList(GuildIndex).CurrentQuest.IdQuest).Stages(GuildList(GuildIndex).CurrentQuest.CurrentStage).EndNpc.NpcIndex
End Function

Public Sub NotifyGuildQuestIsFinished(ByVal GuildIndex As Integer)
On Error GoTo ErrHandler

    Dim Message As String
    Dim I As Integer
    With GuildList(GuildIndex).CurrentQuest
        If .CurrentStage = GuildQuestList(.IdQuest).StageQuantity Then
            Message = "Tu clan ha finalizado la misión '" & GuildQuestList(.IdQuest).Title & "'. Dirigete al Maestro de Clanes para reclamar tu recomensa"
        Else
            Message = "Tu clan ha finalizado una etapa de la misión '" & GuildQuestList(.IdQuest).Title & "'. Dirigete al Maestro de Clanes para reclamar tu recomensa y comenzar la siguiente."
        End If

        If GuildLastMemberOnline(GuildIndex) > 0 Then
            For I = 1 To GuildLastMemberOnline(GuildIndex)
                Call WriteConsoleMsg(GuildList(GuildIndex).OnlineMembers(I).MemberUserIndex, Message, FontTypeNames.FONTTYPE_GUILD, eMessageType.Guild)
            Next I
        End If
        
    End With
    
    Exit Sub
    
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function NotifyGuildQuestIsFinished de modQuestSystem.bas")
End Sub

Public Sub NotifyGuildQuestFinished(ByVal GuildIndex As Integer, ByVal Failed As Boolean, ByVal QuestId As Integer, ByVal StageNumber As Integer)
On Error GoTo ErrHandler
    Dim I As Integer
    
    If GuildLastMemberOnline(GuildIndex) > 0 Then
        For I = 1 To GuildLastMemberOnline(GuildIndex)
            With GuildList(GuildIndex).OnlineMembers(I)
            
                Call WriteGuildQuestUpdate_Finished(.MemberUserIndex, Failed, QuestId, StageNumber)
                Call WriteGuildInfoChange(.MemberUserIndex, eChangeGuildInfo.ContributionAvailableChange, 0, GuildList(GuildIndex).ContributionAvailable)
                
            End With
            
        Next I
    End If
    
    Exit Sub
    
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function NotifyGuildQuestFinished de modQuestSystem.bas")
End Sub

Public Function IsLastStage(ByVal GuildIndex As Integer) As Boolean

    IsLastStage = GuildList(GuildIndex).CurrentQuest.CurrentStage = GuildQuestList(GuildList(GuildIndex).CurrentQuest.IdQuest).StageQuantity

End Function

Public Sub UpdateCurrentQuestInfoToOnlineMembers(ByVal GuildIndex As Integer)
On Error GoTo ErrHandler

    Dim I As Integer

    With GuildList(GuildIndex)
        If GuildLastMemberOnline(GuildIndex) > 0 Then
            For I = 1 To GuildLastMemberOnline(GuildIndex)
                Call WriteGuildCurrentQuestInfo(.OnlineMembers(I).MemberUserIndex)
            Next I
        End If
    End With
    
    Exit Sub
    
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function UpdateCurrentQuestInfoToOnlineMembers de modQuestSystem.bas")
End Sub

Public Sub UpdateCurrentQuestInfoToOnlineMembersButIndex(ByVal GuildIndex As Integer, ByVal UserIndex As Integer)
On Error GoTo ErrHandler

    Dim I As Integer

    With GuildList(GuildIndex)
        If GuildLastMemberOnline(GuildIndex) > 0 Then
            For I = 1 To GuildLastMemberOnline(GuildIndex)
                If .OnlineMembers(I).MemberUserIndex <> UserIndex Then
                    Call WriteGuildCurrentQuestInfo(.OnlineMembers(I).MemberUserIndex)
                End If
            Next I
        End If
    End With
    
    Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function UpdateCurrentQuestInfoToOnlineMembersButIndex de modQuestSystem.bas")
End Sub

Public Function FinishGuildQuestStage(ByVal GuildIndex As Integer, ByVal UserIndex As Integer) As Boolean
    On Error GoTo ErrHandler

    Dim I As Integer
    Dim ContributionGained As Long
    Dim QuestId As Integer
    
    If IsLastStage(GuildIndex) And GuildList(GuildIndex).IdRightHand = 0 Then
        Call WriteConsoleMsg(UserIndex, "Tu clan debe tener una mano derecha para finalizar la misión", FontTypeNames.FONTTYPE_GUILD, eMessageType.Guild)
        FinishGuildQuestStage = False
        Exit Function
    End If
    
    If Not GiveGuildQuestReward(GuildIndex, UserIndex) Then
        FinishGuildQuestStage = False
        Exit Function
    End If

    If IsLastStage(GuildIndex) Then
        'FinishQuest_DB also call sp_DeleteGuildCurrentQuest in the procedure
        Call modGuild_DB.FinishQuest_DB(GuildIndex)
        Call modGuild_DB.UpdateGuildStats(GuildIndex)
        
        With GuildList(GuildIndex)
            .QuestCompletedCount = .QuestCompletedCount + 1
            ReDim Preserve .QuestCompleted(1 To .QuestCompletedCount)
            .QuestCompleted(.QuestCompletedCount).CompletedDate = Now
            .QuestCompleted(.QuestCompletedCount).IdQuest = .CurrentQuest.IdQuest
            .QuestCompleted(.QuestCompletedCount).MembersContributed = .MemberCount
            
            .QuestCompleted(.QuestCompletedCount).ContributionGained = 0

            If .CurrentQuest.IsFirstTime Then
                ContributionGained = GuildQuestList(.CurrentQuest.IdQuest).ContributionEarnedFirstTime
            Else
                ContributionGained = GuildQuestList(.CurrentQuest.IdQuest).ContributionEarned
            End If
            
             .QuestCompleted(.QuestCompletedCount).ContributionGained = ContributionGained
             
            .ContributionAvailable = .ContributionAvailable + ContributionGained
            .ContributionEarned = .ContributionEarned + ContributionGained
            
            Call NotifyGuildQuestFinished(GuildIndex, False, .CurrentQuest.IdQuest, .CurrentQuest.CurrentStage)
            
            .CurrentQuest.IdQuest = 0
            .CurrentQuest.StageIsCompleted = False

        End With
    Else
        Call ChangeGuildQuestStage(GuildIndex, GuildList(GuildIndex).CurrentQuest.IdQuest, GuildList(GuildIndex).CurrentQuest.CurrentStage + 1)
        Call UpdateCurrentQuestInfoToOnlineMembers(GuildIndex)
    End If
    
    FinishGuildQuestStage = True

    Exit Function

ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function FinishGuildQuestStage de modQuestSystem.bas")
End Function

Public Function GiveGuildQuestReward(ByVal GuildIndex As Integer, ByVal UserIndex As Integer) As Boolean
    On Error GoTo ErrHandler

    With GuildList(GuildIndex).CurrentQuest

        If Not .StageIsCompleted Then Exit Function

        If Not GiveQuestReward(UserIndex, GuildQuestList(.IdQuest).Stages(.CurrentStage).Rewards) Then
            GiveGuildQuestReward = False
            Exit Function
        End If

        If IsLastStage(GuildIndex) Then

            If Not GiveQuestReward(UserIndex, GuildQuestList(.IdQuest).Rewards) Then
                GiveGuildQuestReward = False
                Exit Function
            End If

            If GuildList(GuildIndex).CurrentQuest.IsFirstTime Then
                Call GuildEarnContributionPoints(GuildIndex, GuildQuestList(.IdQuest).ContributionEarnedFirstTime)
            Else
                Call GuildEarnContributionPoints(GuildIndex, GuildQuestList(.IdQuest).ContributionEarned)
            End If
        End If

        GiveGuildQuestReward = True
    End With
    
    Exit Function
    
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GiveGuildQuestReward de modQuestSystem.bas")
End Function

Public Function GiveQuestReward(ByVal UserIndex As Integer, ByRef Rewards As tQuestRewards) As Boolean
    On Error GoTo ErrHandler

    Dim I As Integer
    Dim TieneEspacio As Boolean
    Dim RewardOBJ As Obj
    Dim Aux() As tQuestObj

    With UserList(UserIndex)
        
        If Rewards.ObjsQty > 0 Then
            ReDim Aux(1 To Rewards.ObjsQty) As tQuestObj
            TieneEspacio = True
    
            For I = 1 To Rewards.ObjsQty
                RewardOBJ.Amount = Rewards.Objs(I).ObjQty
                RewardOBJ.ObjIndex = Rewards.Objs(I).ObjIndex
    
                If MeterItemEnInventario(UserIndex, RewardOBJ) Then
                    Aux(I).ObjIndex = Rewards.Objs(I).ObjIndex
                    Aux(I).ObjQty = Rewards.Objs(I).ObjQty
                Else
                    TieneEspacio = False
                    Exit For
                End If
            Next I
    
             If Not TieneEspacio Then
                Call WriteConsoleMsg(UserIndex, "No posees espacio en tu inventario para la recomensa", FontTypeNames.FONTTYPE_GUILD, eMessageType.Guild)
    
                For I = 1 To Rewards.ObjsQty
                    If Aux(I).ObjIndex > 0 Then
                        Call QuitarObjetos(Aux(I).ObjIndex, Aux(I).ObjQty, UserIndex)
                    Else
                        Exit For
                    End If
                Next I
    
                Exit Function
            End If
            
            'TODO-QUESTS gold and exp
            
        End If

        GiveQuestReward = True
    End With
    Exit Function

ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GiveQuestReward de modQuestSystem.bas")
End Function

Public Sub ChangeGuildQuestStage(ByVal GuildIndex As Integer, ByVal QuestId As Integer, ByVal StageNumber As Integer, Optional ByVal UpdateDb As Boolean = True)
    On Error GoTo ErrHandler

    Dim I As Integer

    With GuildList(GuildIndex).CurrentQuest

        .IdQuest = QuestId
        .CurrentStage = StageNumber
        .CanStartNextStage = False
        .StageIsCompleted = False

        .CurrentFrags.Army.Qty = 0
        .CurrentFrags.Legion.Qty = 0
        .CurrentFrags.Neutral.Qty = 0
        
        .CurrentFrags.Army.MinLevel = GuildQuestList(QuestId).Stages(StageNumber).Frags.Army.MinLevel
        .CurrentFrags.Legion.MinLevel = GuildQuestList(QuestId).Stages(StageNumber).Frags.Legion.MinLevel
        .CurrentFrags.Neutral.MinLevel = GuildQuestList(QuestId).Stages(StageNumber).Frags.Neutral.MinLevel

        
        If GuildQuestList(QuestId).Stages(StageNumber).NpcsKillsQuantity <> 0 Then
            ReDim .CurrentNpcKills(1 To GuildQuestList(QuestId).Stages(StageNumber).NpcsKillsQuantity)
            
            .CurrentNpcKillsQuantity = GuildQuestList(QuestId).Stages(StageNumber).NpcsKillsQuantity
            
             For I = 1 To GuildQuestList(QuestId).Stages(StageNumber).NpcsKillsQuantity
                .CurrentNpcKills(I).NpcIndex = GuildQuestList(QuestId).Stages(StageNumber).NpcKill(I).NpcIndex
                .CurrentNpcKills(I).Quantity = 0
            Next I
        End If
        
        If GuildQuestList(QuestId).Stages(StageNumber).ObjsCollectQuantity > 0 Then
            .CurrentObjectList = modRequiredObjectList.RequiredObjectListCreate(GuildQuestList(QuestId).Stages(StageNumber).ObjsCollect, GuildQuestList(QuestId).Stages(StageNumber).ObjsCollectQuantity)
        Else
            .CurrentObjectList = modRequiredObjectList.RequiredObjectListCreateCompleted()
        End If
        
        If UpdateDb Then
            Call modGuild_DB.UpdateCurrentQuestStage(GuildList(GuildIndex).IdGuild, StageNumber)
            Call modGuild_DB.DeleteCurrentQuestStatus(GuildList(GuildIndex).IdGuild)
        End If

    End With

    Exit Sub

ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function ChangeGuildQuestStage de modQuestSystem.bas")
End Sub

Public Function GetQuestGuildCompletedIndexOf(ByVal GuildIndex As Integer, ByVal QuestId As Integer) As Integer
    On Error GoTo ErrHandler

    Dim I As Integer
    
    With GuildList(GuildIndex)
        For I = 1 To .QuestCompletedCount
            If .QuestCompleted(I).IdQuest = QuestId Then
                GetQuestGuildCompletedIndexOf = I
                Exit Function
            End If
        Next I
    End With
    GetQuestGuildCompletedIndexOf = 0
    Exit Function

ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetQuestGuildCompletedIndexOf de modQuestSystem.bas")
End Function
Public Sub SaveQuestStatus()
    On Error GoTo ErrHandler
    
    If MaxGuildQty = 0 Then Exit Sub
    
    Dim I As Integer

    For I = 1 To MaxGuildQty
        With GuildList(I)
            If .CurrentQuest.IdQuest <> 0 Then
                Call DeleteCurrentQuestStatus(.IdGuild)
                Call modGuild_DB.SaveCurrentQuestStatus(I)
            End If
            
        End With
    Next I

    Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function SaveQuestStatus de modQuestSystem.bas")
End Sub


Public Function GuildCanStartQuest(ByVal QuestId As Integer, ByVal GuildIndex As Integer) As Boolean
    On Error GoTo ErrHandler

    Dim QuestTimesCompleted As Integer
    Dim I As Integer
    Dim AuxDate As Date

    With GuildList(GuildIndex)
        
        If Not GuildQuestList(QuestId).Active Then
            GuildCanStartQuest = False
            Exit Function
        End If
        
        If GuildQuestList(QuestId).RepetitionQuantity <> 0 Then
            QuestTimesCompleted = 0

            For I = 1 To .QuestCompletedCount
                If .QuestCompleted(I).IdQuest = QuestId Then
                    QuestTimesCompleted = QuestTimesCompleted + 1
                End If
            Next I

            If QuestTimesCompleted > GuildQuestList(QuestId).RepetitionQuantity Then
                 GuildCanStartQuest = False
                 Exit Function
            End If
        End If

        If .MemberCount < GuildQuestList(QuestId).MinMembers Then
            GuildCanStartQuest = False
            Exit Function
        End If

        If GuildQuestList(QuestId).CorrelativesQuantity > 0 Then
            For I = 1 To GuildQuestList(QuestId).CorrelativesQuantity
                If GetQuestGuildCompletedIndexOf(GuildIndex, GuildQuestList(QuestId).Correlatives(I).IdQuest) = 0 Then
                    GuildCanStartQuest = False
                    Exit Function
                End If
            Next I
        End If
        
        If .QuestCompletedCount > 0 Then
            For I = 1 To .QuestCompletedCount
                If .QuestCompleted(I).IdQuest = QuestId Then
                    AuxDate = DateAdd("s", GuildQuestList(QuestId).Cooldown, GuildList(GuildIndex).QuestCompleted(I).CompletedDate)
                    If DateDiff("s", Now, AuxDate) > 0 Then
                        GuildCanStartQuest = False
                        Exit Function
                    End If
                End If
            Next I
        End If
        
        If GuildQuestList(QuestId).Alignment > 0 Then
            If .Alignment <> GuildQuestList(QuestId).Alignment Then
                GuildCanStartQuest = False
                Exit Function
            End If
        End If

    End With

    GuildCanStartQuest = True
    Exit Function
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GuildCanStartQuest de modQuestSystem.bas")
End Function
Public Function GetQuestStageRemainingTime(ByVal GuildIndex As Integer) As Long
    Dim QuestId As Integer
    Dim TotalSeconds As Long
    
    If Not GuildHasQuest(GuildIndex) Then
        GetQuestStageRemainingTime = 0
        Exit Function
    End If
    
    QuestId = GuildList(GuildIndex).CurrentQuest.IdQuest
    
    TotalSeconds = GuildList(GuildIndex).CurrentQuest.SecondsLeft - DateDiff("s", GuildList(GuildIndex).CurrentQuest.ServerStartedDate, Now())
    
    If TotalSeconds < 0 Then
        TotalSeconds = 0
    End If
    
    GetQuestStageRemainingTime = TotalSeconds
    
End Function
Public Function IsQuestOverdue(ByVal GuildIndex As Integer) As Boolean
 On Error GoTo ErrHandler
 
    Dim TotalSeconds As Long
    
    If Not GuildHasQuest(GuildIndex) Then
        IsQuestOverdue = False
        Exit Function
    End If
    
    TotalSeconds = GetQuestStageRemainingTime(GuildIndex)
    
    If TotalSeconds <= 0 Then
        IsQuestOverdue = True
        Exit Function
    End If
    
    IsQuestOverdue = False
    Exit Function
    
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function IsQuestOverdue de modQuestSystem.bas")
End Function

Public Function RemoveQuestItemFromInv(ByVal UserIndex As Integer) As Boolean
On Error GoTo ErrHandler
    Dim ItemIndex As Integer
    Dim I As Integer
    Dim InvSlots As Integer
    
    RemoveQuestItemFromInv = False
    
    With UserList(UserIndex)
        InvSlots = .CurrentInventorySlots
        For I = 1 To InvSlots
            
            ItemIndex = .Invent.Object(I).ObjIndex
            If ItemIndex > 0 Then
                If ObjData(ItemIndex).ObjType = otQuest Then
                    Call QuitarObjetos(ItemIndex, .Invent.Object(I).Amount, UserIndex)
                    RemoveQuestItemFromInv = True
                End If
            End If
        Next I
    End With
    
    Exit Function
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function RemoveQuestItemFromInv de modQuestSystem.bas")
End Function

Public Sub SetQuestTimeOnStateServer(ByVal GuildIndex As Integer)
On Error GoTo ErrHandler
    Dim QuestSecondsLeft As Long
    
    QuestSecondsLeft = GetQuestStageRemainingTime(GuildIndex)
    
    Call modStateServer.SendAddGuildQuest(GuildIndex, QuestSecondsLeft)
    
    Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function SetQuestTimeOnStateServer de modQuestSystem.bas")
End Sub
Public Sub SendQuestsTimes()
    On Error GoTo ErrHandler
    
    If MaxGuildQty = 0 Then Exit Sub

    Dim I As Integer

    For I = 1 To MaxGuildQty
        With GuildList(I)
            If .CurrentQuest.IdQuest <> 0 Then
                Call SetQuestTimeOnStateServer(I)
            End If
        End With
    Next I
    
    Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SendQuestsTimes de modQuestSystem.bas")
End Sub

Public Sub StartQuest(ByVal UserIndex As Integer, ByVal QuestId As Integer)
On Error GoTo ErrHandler
    Dim GuildIndex As Integer
    GuildIndex = UserList(UserIndex).Guild.GuildIndex
    
    With GuildList(GuildIndex)
    
        If .IdRightHand = 0 Then
            Call WriteConsoleMsg(UserIndex, "El Clan debe tener una Mano derecha asignada para poder iniciar una misión.", FontTypeNames.FONTTYPE_GUILD, eMessageType.Guild)
            Exit Sub
        End If
        If Not GuildCanStartQuest(QuestId, GuildIndex) Then
            Call WriteConsoleMsg(UserIndex, "No cumples los requisitos necesarios para comenzar esta misión.", FontTypeNames.FONTTYPE_GUILD, eMessageType.Guild)
            Exit Sub
        End If
        
        .CurrentQuest.IsFirstTime = (GetQuestGuildCompletedIndexOf(GuildIndex, QuestId) = 0)
        .CurrentQuest.StartedDate = Now()
                    
        Call ChangeGuildQuestStage(GuildIndex, QuestId, 1, False)
                
        .CurrentQuest.SecondsLeft = GuildQuestList(QuestId).Duration
        .CurrentQuest.ServerStartedDate = Now()
                
        Call modGuild_DB.AcceptQuest(GuildIndex, .CurrentQuest.IdQuest, .CurrentQuest.CurrentStage, .CurrentQuest.StartedDate, .CurrentQuest.SecondsLeft)
        Call UpdateCurrentQuestInfoToOnlineMembers(GuildIndex)
        Call SetQuestTimeOnStateServer(GuildIndex)
            
    End With
    
    Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub StartQuest de modQuestSystem.bas")

End Sub


Public Sub CleanCurrentQuestInfo(ByVal GuildIndex As Integer)
On Error GoTo ErrHandler

    If GuildList(GuildIndex).CurrentQuest.IdQuest = 0 Then Exit Sub
    
    With GuildList(GuildIndex)
        .CurrentQuest.IdQuest = 0
        .CurrentQuest.CurrentStage = 0
        .CurrentQuest.SecondsLeft = 0
        .CurrentQuest.StartedDate = 0
    End With
    
    Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CleanCurrentQuestInfo de modQuestSystem.bas")

End Sub

Public Sub CancelCurrentGuildQuest(ByVal GuildIndex As Integer, ByVal CancelFromUser As Boolean)
On Error GoTo ErrHandler
    
    Dim I As Integer
    
    If GuildIndex > 0 Then
        If GuildList(GuildIndex).CurrentQuest.IdQuest <> 0 Then
            With GuildList(GuildIndex)
                Call modGuild_DB.DeleteCurrentQuest(GuildIndex)
                Call NotifyGuildQuestFinished(GuildIndex, True, .CurrentQuest.IdQuest, .CurrentQuest.CurrentStage)
                
                Call CleanCurrentQuestInfo(GuildIndex)
                Call UpdateCurrentQuestInfoToOnlineMembers(GuildIndex)

                If CancelFromUser Then
                    Call modStateServer.SendRemoveGuildQuest(GuildIndex)
                End If
            End With
        End If
    End If
    
    Exit Sub
    
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CancelCurrentGuildQuest de modQuestSystem.bas")
End Sub


Public Sub RemoveQuestItemFromOnlineMembers(ByVal GuildIndex As Integer)
On Error GoTo ErrHandler

    Dim I As Integer
    If GuildLastMemberOnline(GuildIndex) > 0 Then
        For I = 1 To GuildLastMemberOnline(GuildIndex)
            Call RemoveQuestItemFromInv(I)
        Next I
    End If
    Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub RemoveQuestItemFromOnlineMembers de modQuestSystem.bas")
End Sub

Public Function GetGuildQuestTalkNpc(ByVal GuildIndex As Integer) As Integer
On Error GoTo ErrHandler
    
    Dim QuestId As Integer
    
    QuestId = GuildList(GuildIndex).CurrentQuest.IdQuest
    
    With GuildList(GuildIndex).CurrentQuest
        
        If .CurrentStage = 0 Then
            GetGuildQuestTalkNpc = GuildQuestList(QuestId).Stages(1).StarterNpc.NpcIndex
            Exit Function
        End If
        
        If .StageIsCompleted Then
            If GuildQuestList(QuestId).Stages(.CurrentStage).EndNpc.NpcIndex = 0 Then
                If Not IsLastStage(GuildIndex) Then
                    GetGuildQuestTalkNpc = GuildQuestList(QuestId).Stages(.CurrentStage + 1).StarterNpc.NpcIndex
                    Exit Function
                End If
            Else
                GetGuildQuestTalkNpc = GuildQuestList(QuestId).Stages(.CurrentStage).EndNpc.NpcIndex
                Exit Function
            End If
        End If
        
    End With
    
    GetGuildQuestTalkNpc = 0
    Exit Function
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub GetGuildQuestTalkNpc de modQuestSystem.bas")
End Function

Public Function QuestLastIndex() As Integer

On Error GoTo ErrHandler

    If ((Not GuildQuestList) = -1) Then
        QuestLastIndex = 0
    Else
        QuestLastIndex = UBound(GuildQuestList)
    End If

  Exit Function

ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function QuestLastIndex de modQuestSystem.bas")
End Function
