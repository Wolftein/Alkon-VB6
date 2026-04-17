Attribute VB_Name = "modStateServer"
Option Explicit

Public InboundByteQueue As New clsByteQueue
Public OutboundByteQueue As New clsByteQueue

Public Enum eStateServerRequestCodes
    ADD_CAMPFIRE = 0
    DOT_ADD
    DOT_REMOVE
    DOT_LOAD
    DOT_PERSIST
    RESOURCE
    GUILD_QUEST_ADD_TIMEOUT
    GUILD_QUEST_REMOVE_TIMEOUT
End Enum

Public Enum eStateServerResponseCodes
    REMOVE_CAMPFIRE = 0
    DOT_TICK = 1
    RESOURCE_RESPAWN = 5
    GUILD_QUEST_TIMEOUT = 6
End Enum

Public Function HandleStateServerMessage(ByRef buffer As clsByteQueue) As Boolean
On Error GoTo ErrHandler

    Select Case buffer.PeekByte
        Case eStateServerResponseCodes.REMOVE_CAMPFIRE ' Remove Campfire
            If buffer.length < 5 Then
                Exit Function
            End If
                
            Call buffer.ReadByte
            
            Call RemoveCampfire(buffer.ReadInteger(), buffer.ReadByte(), buffer.ReadByte())
       Case eStateServerResponseCodes.DOT_TICK  ' Damage over time tick
            If buffer.length < 9 Then
                Exit Function
            End If
       
            Call buffer.ReadByte
            
            Call DamageOverTimeTick(buffer.ReadByte, buffer.ReadLong(), buffer.ReadLong(), buffer.ReadByte, buffer.ReadLong(), buffer.ReadLong(), buffer.ReadInteger(), buffer.ReadBoolean())
        Case eStateServerResponseCodes.RESOURCE_RESPAWN 'Resource respawn
            Call buffer.ReadByte

            Call ResourceRespawn(buffer.ReadInteger(), buffer.ReadByte(), buffer.ReadByte(), buffer.ReadInteger())

        Case eStateServerResponseCodes.GUILD_QUEST_TIMEOUT 'Resource respawn
            Call buffer.ReadByte
            
            Call modQuestSystem.CancelCurrentGuildQuest(buffer.ReadLong(), False)
    End Select
    
    HandleStateServerMessage = buffer.length <> 0
    
    Exit Function
    
ErrHandler:
    Call LogError("Error en HandleStateServerMessage. " & Err.Number & " - " & Err.Description)
End Function
Public Sub OnConnected()
    modQuestSystem.SendQuestsTimes
End Sub
Public Sub SendStateServerData(ByRef buffer As clsByteQueue)
On Error GoTo ErrHandler:

    Dim blockToSend() As Byte
    ReDim blockToSend(0 To buffer.length - 1)
    Call buffer.ReadBlock(blockToSend, buffer.length)

    If IsStateServerOnline() Then
        Call frmMain.sckStateServer.SendData(blockToSend)
    End If
    Exit Sub
    
ErrHandler:
    Call LogError("Error en SendStateServerData sending " & buffer.length & " bytes. Error: : " & Err.Number & ": " & Err.Description)
End Sub

Public Function IsStateServerOnline() As Boolean
    On Error GoTo ErrHandler:

    IsStateServerOnline = frmMain.sckStateServer.State = sckConnected
    Exit Function
    
ErrHandler:
    Call LogError("Error en IsStateServerOnline. Error: : " & Err.Number & ": " & Err.Description)
End Function

Public Sub SendDamageOverTimeMessage(ByVal OriginType As eTypeTarget, ByVal OriginId As Long, ByVal OriginInstanceId As Long, ByRef OriginName As String, _
                                     ByVal TargetType As eTypeTarget, ByVal TargetId As Long, ByVal TargetInstanceId As Long, ByRef TargetName As String, _
                                     ByVal TickCount As Integer, ByVal TickInterval As Integer, ByVal SpellNumber As Integer, ByVal MaxStackEffect As Integer)
    On Error GoTo ErrHandler:

    
    Call OutboundByteQueue.WriteByte(eStateServerRequestCodes.DOT_ADD) ' Packet ID: 0=CAMPFIRE
    'origin dot
    Call OutboundByteQueue.WriteByte(OriginType)
    Call OutboundByteQueue.WriteLong(OriginId)
    Call OutboundByteQueue.WriteLong(OriginInstanceId)
    Call OutboundByteQueue.WriteASCIIString(OriginName)
    'target dot
    Call OutboundByteQueue.WriteByte(TargetType)
    Call OutboundByteQueue.WriteLong(TargetId)
    Call OutboundByteQueue.WriteLong(TargetInstanceId)
    Call OutboundByteQueue.WriteASCIIString(TargetName)
    'spell info
    Call OutboundByteQueue.WriteInteger(TickCount)
    Call OutboundByteQueue.WriteInteger(TickInterval)
    Call OutboundByteQueue.WriteInteger(SpellNumber)
    Call OutboundByteQueue.WriteInteger(MaxStackEffect)
    
    Call SendStateServerData(OutboundByteQueue)
    Exit Sub
    
ErrHandler:
    Call LogError("Error en SendDamageOverTimeMessage. User / Spell:( " & TargetName & " / " & SpellNumber & "). Error: : " & Err.Number & ": " & Err.Description)
End Sub

Public Sub SendDamageOverTimeRemoveMessage(ByVal TargetType As eTypeTarget, ByVal TargetId As Long, ByVal TargetInstanceId As Long)
    On Error GoTo ErrHandler:
    
    Call OutboundByteQueue.WriteByte(eStateServerRequestCodes.DOT_REMOVE)
    Call OutboundByteQueue.WriteByte(TargetType)
    Call OutboundByteQueue.WriteLong(TargetId)
    Call OutboundByteQueue.WriteLong(TargetInstanceId)
    Call SendStateServerData(OutboundByteQueue)
    Exit Sub
    
ErrHandler:
    Call LogError("Error en SendDamageOverTimeRemoveMessage. IdUSer: " & TargetId & "). Error: : " & Err.Number & ": " & Err.Description)
End Sub

Public Sub SendDamageOverTimePersist(ByVal UserId As Long)
    
    Call OutboundByteQueue.WriteByte(eStateServerRequestCodes.DOT_PERSIST)
    Call OutboundByteQueue.WriteByte(eTypeTarget.isUser)
    Call OutboundByteQueue.WriteLong(UserId)
    Call SendStateServerData(OutboundByteQueue)
    
    Exit Sub
End Sub


Public Sub SendDamageOverTimeLoad(ByVal UserId As Long)
    
    Call OutboundByteQueue.WriteByte(eStateServerRequestCodes.DOT_LOAD)
    Call OutboundByteQueue.WriteByte(eTypeTarget.isUser)
    Call OutboundByteQueue.WriteLong(UserId)
    Call SendStateServerData(OutboundByteQueue)
    
    Exit Sub
End Sub


Public Function SendAddGuildQuest(ByVal GuildIndex As Long, ByVal DurationInSeconds As Long)
    
    Call OutboundByteQueue.WriteByte(eStateServerRequestCodes.GUILD_QUEST_ADD_TIMEOUT)
    Call OutboundByteQueue.WriteLong(GuildIndex)
    Call OutboundByteQueue.WriteLong(DurationInSeconds)
    
    Call SendStateServerData(OutboundByteQueue)
   
End Function

Public Function SendRemoveGuildQuest(ByVal GuildIndex As Long)
    
    Call OutboundByteQueue.WriteByte(eStateServerRequestCodes.GUILD_QUEST_REMOVE_TIMEOUT)
    Call OutboundByteQueue.WriteLong(GuildIndex)
    
    Call SendStateServerData(OutboundByteQueue)
   
End Function

Public Function SendAddCampfire(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal duration As Long)
On Error GoTo ErrHandler:
    
    Call OutboundByteQueue.WriteByte(eStateServerRequestCodes.ADD_CAMPFIRE) ' Packet ID: 0=CAMPFIRE
    Call OutboundByteQueue.WriteInteger(Map)
    Call OutboundByteQueue.WriteByte(X)
    Call OutboundByteQueue.WriteByte(Y)
    Call OutboundByteQueue.WriteLong(duration)
    
    Call SendStateServerData(OutboundByteQueue)
   
    Exit Function
ErrHandler:
    Call LogError("Error en SendAddCampfire. MAP-X-Y:( " & Map & "-" & X & "-" & Y & "). Error: : " & Err.Number & ": " & Err.Description)
End Function


Public Sub DamageOverTimeTick(ByVal OriginType As eTypeTarget, ByVal OriginId As Long, ByVal OriginInstanceId As Long, ByVal TargetType As eTypeTarget, ByVal TargetId As Long, ByVal TargetInstanceId As Long, ByVal SpellNumber As Integer, ByVal isLastTick As Boolean)
On Error GoTo ErrHandler:
    
    ' TODO: Check if user index passed is online. If not, then send a message to the state server
    ' to stop the DoT effect.
    
    If TargetType = eTypeTarget.isUser Then
        Call TickUser(TargetId, GetUserIndexFromUserId(TargetId), SpellNumber, GetUserIndexFromUserId(OriginId))
    Else
        If Npclist(TargetId).InstanceId <> TargetInstanceId Then
            Call SendDamageOverTimeRemoveMessage(TargetType, TargetId, TargetInstanceId)
        Else
            Call TickNpc(TargetId, GetUserIndexFromUserId(OriginId), SpellNumber, isLastTick)
        End If
    End If

    Exit Sub
            
ErrHandler:
    'Call LogError("Error en RemoveCampfire. MAP-X-Y:( " & Map & "-" & X & "-" & Y & "). Error: : " & Err.Number & ": " & Err.Description)
End Sub

Public Sub RemoveCampfire(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
On Error GoTo ErrHandler:
    If MapData(Map, X, Y).ObjInfo.ObjIndex = ConstantesItems.FogataElfica Or MapData(Map, X, Y).ObjInfo.ObjIndex = ConstantesItems.Fogata Then
        Call EraseObj(ConstantesItems.FogataElfica, Map, X, Y)
        
        Dim campfirePoss As WorldPos
        Dim I As Integer
        
        With campfirePoss
            .Map = Map
            .X = X
            .Y = Y
        End With
                
        ' TODO: Replace this to only check in a square/romboid around the campfire. Players outside this space won't be resting
        ' because of the validatios to the /descansar command.
        For I = 1 To LastUser
            With UserList(I)
                If .RestObjectCoords.Map = campfirePoss.Map And .RestObjectCoords.X = campfirePoss.X And .RestObjectCoords.Y = campfirePoss.Y Then
                    If .flags.Descansar Then
                        Call UserRest(I)
                    Else
                        .RestObjectCoords.Map = 0
                        .RestObjectCoords.X = 0
                        .RestObjectCoords.Y = 0
                    End If
                End If
            End With
        Next I
        
        
    End If
    
    Exit Sub
            
ErrHandler:
    Call LogError("Error en RemoveCampfire. MAP-X-Y:( " & Map & "-" & X & "-" & Y & "). Error: : " & Err.Number & ": " & Err.Description)
End Sub

Public Sub SendResourceToSpawn(ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal ObjNumber As Integer, ByVal CooldownTime As Long)
On Error GoTo ErrHandler:

    Call OutboundByteQueue.WriteByte(eStateServerRequestCodes.RESOURCE)
    Call OutboundByteQueue.WriteInteger(Map)
    Call OutboundByteQueue.WriteByte(X)
    Call OutboundByteQueue.WriteByte(Y)
    Call OutboundByteQueue.WriteInteger(ObjNumber)
    Call OutboundByteQueue.WriteLong(CooldownTime)
    
    Call SendStateServerData(OutboundByteQueue)
    Exit Sub
    
ErrHandler:
    Call LogError("Error en SendResourceToSpawn. MAP-X-Y:( " & Map & "-" & X & "-" & Y & "), Object: " & ObjNumber & "Error: : " & Err.Number & ": " & Err.Description)
End Sub

Public Sub ResourceRespawn(ByVal Map As Integer, ByVal tX As Byte, ByVal tY As Byte, ByVal ObjIndex As Integer)
On Error GoTo ErrHandler:

    'Check the resource type to respawn
    If ObjData(ObjIndex).ObjType = otResource Then

        Dim TreeToReplace As Integer
        Dim X As Byte
        Dim Y As Byte
        Dim ind As Integer
        Dim tempArray() As WorldPos
        Dim tempIndex As Integer
        
        TreeToReplace = 0
        
        'Check if there are any empty slots
        If MapInfo(Map).MapResources.ResourceGroupQty = 0 Then
            Exit Sub
        End If
        
        'Try to respawn in a random empty slot, avoiding to respawn in the same original position but checking if its the only one
        Dim ValidSlotFound As Boolean
        ValidSlotFound = False
        
        Dim I As Integer
        Dim IterationCut As Integer
      
        For I = 1 To MapInfo(Map).MapResources.ResourceGroupQty
            With MapInfo(Map).MapResources
                If .ResourceGroup(I).ObjNumber = ObjIndex Then
                    ' Get a random element from the list. If there's no empty resources, then quit.
                    ' This should never happen if the system is working correctly.
                    If .ResourceGroup(I).EmptyResourceQty = 0 Then
                        Exit Sub
                    End If
                    
                    TreeToReplace = RandomNumber(1, .ResourceGroup(I).EmptyResourceQty)
                    
                    X = .ResourceGroup(I).EmptyResourcePositions(TreeToReplace).X
                    Y = .ResourceGroup(I).EmptyResourcePositions(TreeToReplace).Y
                    
                    ValidSlotFound = True
                    
                    ' Remove the "Empty" element from the list in that group
                    .ResourceGroup(I).EmptyResourcePositions = RemoveItemFromArray(TreeToReplace, .ResourceGroup(I).EmptyResourcePositions)
                    .ResourceGroup(I).EmptyResourceQty = .ResourceGroup(I).EmptyResourceQty - 1
                    Exit For
                End If
            End With
        Next I
        
        ' If the slot we were looking for is not available, then we exit, as there was no "empty" space for creating the resource.
        If Not ValidSlotFound Then Exit Sub
    
        With MapData(Map, X, Y)
            ' Replace the empty slot with the full resource
            .ObjInfo.ObjIndex = ObjIndex
            .ObjInfo.PendingQty = ObjData(ObjIndex).MaxExtractedQuantity
            .ObjInfo.CurrentGrhIndex = ObjData(ObjIndex).GrhIndex
            
            ReDim .ObjInfo.Resources(1 To ObjData(ObjIndex).NumResources)
            For ind = 1 To ObjData(ObjIndex).NumResources
                .ObjInfo.Resources(ind) = Resources(ObjData(ObjIndex).Resources(ind).ResourceNumber)
            Next ind
            
            'Send the Create Object to the client with the data of the new tree
            If .ObjInfo.ObjIndex > 0 Then
                If .Trigger <> eTrigger.zonaOscura Then
                    Call SendToItemArea(Map, X, Y, PrepareMessageObjectCreate(ObjData(ObjIndex).GrhIndex, X, Y, ObjData(ObjIndex).ObjType, 0, ObjData(ObjIndex).Luminous, ObjData(ObjIndex).LightOffsetX, ObjData(ObjIndex).LightOffsetY, ObjData(ObjIndex).LightSize, ObjData(ObjIndex).CanBeTransparent))
                Else
                    Call SendToItemAreaButCounselors(Map, X, Y, PrepareMessageObjectCreate(ObjData(ObjIndex).GrhIndex, X, Y, ObjData(ObjIndex).ObjType, 0, ObjData(ObjIndex).Luminous, ObjData(ObjIndex).LightOffsetX, ObjData(ObjIndex).LightOffsetY, ObjData(ObjIndex).LightSize, ObjData(ObjIndex).CanBeTransparent))
                End If
            End If
                    
        End With
    
    End If
    
    Exit Sub
            
ErrHandler:
    Call LogError("Error en ResourceRespawn. MAP-X-Y:( " & Map & "-" & X & "-" & Y & ") for ObjNumber: " & ObjIndex & "(" & ObjData(ObjIndex).Name & "). Error: : " & Err.Number & ": " & Err.Description)
End Sub

Private Function RemoveItemFromArray(ByVal indexToRemove As Integer, intSrc() As WorldPos) As WorldPos()
  Dim intIndex As Integer
  Dim intDest() As WorldPos
  Dim intLBound As Integer, intUBound As Integer
  'find the boundaries of the source array
  intLBound = LBound(intSrc)
  intUBound = UBound(intSrc)
  
  If intUBound = 1 Then
    ' Return the empty array
    Exit Function
  End If
  'set boundaries for the resulting array
  ReDim intDest(intLBound To intUBound - 1) As WorldPos
  'copy items which remain
  For intIndex = intLBound To indexToRemove - 1
    intDest(intIndex) = intSrc(intIndex)
  Next intIndex
  'skip the removed item
  'and copy the remaining items, with destination index-1
  For intIndex = indexToRemove + 1 To intUBound
    intDest(intIndex - 1) = intSrc(intIndex)
  Next intIndex
  'return the result
  RemoveItemFromArray = intDest
End Function

Public Sub TickNpc(ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByVal SpellNumber As Integer, ByVal isLastTick As Boolean)
On Error GoTo ErrHandler:
        

    Npclist(NpcIndex).flags.isAffectedByDOT = Not isLastTick

    If Npclist(NpcIndex).Stats.MinHp <= 0 Then
        Call SendDamageOverTimeRemoveMessage(eTypeTarget.IsNPC, NpcIndex, Npclist(NpcIndex).InstanceId)
        Exit Sub
    End If
    
    With Hechizos(SpellNumber)
        ' Sound & fx
        Call modTriggers.SendSpellEffects(0, NpcIndex, .WAV, .FXgrh, .Loops, 0, 0)
        ' Cast and damage Npc.
        Call modTriggers.CastSpellNpc(NpcIndex, SpellNumber, 0, 10, 15, False, UserIndex)

    End With

    Exit Sub

ErrHandler:
    Call LogError("Error en TickNpc. Error: : " & Err.Number & ": " & Err.Description)
End Sub

Public Sub TickUser(ByVal UserId As Integer, ByVal UserIndex As Integer, ByVal SpellNumber As Integer, ByVal OriginIndex As Integer)
On Error GoTo ErrHandler:
    Dim IsDead As Boolean
    With Hechizos(SpellNumber)
        If UserList(UserIndex).ID <> UserId Then
            Exit Sub
        End If
        ' check if user is dead
        If UserList(UserIndex).Stats.MinHp <= 0 Then
            Call SendDamageOverTimeRemoveMessage(eTypeTarget.isUser, UserId, UserList(UserIndex).InstanceId)
            Exit Sub
        End If
        
        Call CancelExit(UserIndex)
        ' Sound & fx
        Call modTriggers.SendSpellEffects(UserIndex, 0, .WAV, .FXgrh, .Loops, 0, 0)
        ' Cast and damage user.
        Call CastSpellUser(UserIndex, SpellNumber, 0, 10, 15, IsDead, OriginIndex)
        
        ' check if trap's damage kills user
        If IsDead Then
            Call SendDamageOverTimeRemoveMessage(eTypeTarget.isUser, UserId, UserList(UserIndex).InstanceId)
            Exit Sub
        End If
        
    End With

    Exit Sub
    
ErrHandler:
    Call LogError("Error en TickUser. Error: : " & Err.Number & ": " & Err.Description)
End Sub


 
            
