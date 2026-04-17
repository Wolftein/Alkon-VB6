Attribute VB_Name = "modQuests_Guild"
'@Folder("Quest")
Public Function UserGuildHasQuest() As Boolean
On Error GoTo ErrHandler
    
  If PlayerData.Guild.Quest.Id = 0 Then
    UserGuildHasQuest = False
    Exit Function
  End If
  
  UserGuildHasQuest = True
Exit Function
ErrHandler:
     Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub UserGuildHasQuest de modQuests.bas")
End Function

' Used
Public Function IsQuestNpc(ByVal NpcNumber As Integer) As Boolean
On Error GoTo ErrHandler
    
    Dim I As Integer
    
    If NpcNumber = 0 Then
        IsQuestNpc = False
        Exit Function
    End If
    
    If Not UserGuildHasQuest() Then
        IsQuestNpc = False
        Exit Function
    End If

    With GameMetadata.GuildQuests(PlayerData.Guild.Quest.Id).Stages(PlayerData.Guild.Quest.CurrentStage)
        If NpcNumber = .EndNpc.NpcIndex Then
            IsQuestNpc = True
            Exit Function
        End If
    End With
    
    IsQuestNpc = False
Exit Function
ErrHandler:
     Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub IsQuestNpc de modQuests.bas")
End Function

