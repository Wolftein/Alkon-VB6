Attribute VB_Name = "modGuild_DB"
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'@Folder("Guild")
Option Explicit

Public Function CreateGuildFromDB(ByVal GuildName As String, ByVal LeaderId As Long, ByVal Alignment As Double, ByVal RankingPoints As Integer) As Integer
On Error GoTo ErrHandler
  
       
    Dim Rs As Recordset
    Dim Cmd As ADODB.Command
    Dim Size As Integer, GuildIndex As Integer
    
    Set Cmd = New ADODB.Command
    
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "sp_CreateGuild"
    
    Cmd.Parameters.Append Cmd.CreateParameter("GuildName", DataTypeEnum.adBSTR, ParameterDirectionEnum.adParamInput, 1, GuildName)
    Cmd.Parameters.Append Cmd.CreateParameter("LeaderId", DataTypeEnum.adInteger, ParameterDirectionEnum.adParamInput, 1, LeaderId)
    Cmd.Parameters.Append Cmd.CreateParameter("Alignment", DataTypeEnum.adTinyInt, ParameterDirectionEnum.adParamInput, 1, Alignment)
    Cmd.Parameters.Append Cmd.CreateParameter("RankingPoints", DataTypeEnum.adInteger, ParameterDirectionEnum.adParamInput, 10, RankingPoints)
    
    
    Set Rs = ExecuteSqlCommand(Cmd)
    
    Size = SizeRS(Rs)
    
    If Size = 0 Then
        CreateGuildFromDB = -1
        Exit Function
    End If
    
    MaxGuildQty = MaxGuildQty + 1
    
    ReDim Preserve GuildList(1 To MaxGuildQty) As GuildType
    
    Call LoadGuildFromRS(MaxGuildQty, Rs)

    Rs.Close
    Set Rs = Nothing
    Set Cmd = Nothing
    
    CreateGuildFromDB = MaxGuildQty
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CreateGuildFromDB de modDB_Guild.bas")
End Function

Public Sub LoadGuildDB(ByVal GuildIndex As Integer, ByVal GuildId As Long)
'***************************************************
'Author:
'Last Modification: 17/09/2020
'Loads Guild info info from DB
'***************************************************
On Error GoTo ErrHandler
  
    Dim Rs As Recordset
    Dim Cmd As ADODB.Command
    
    Set Cmd = New ADODB.Command
    
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "sp_LoadGuild"
    
    Cmd.Parameters.Append Cmd.CreateParameter("GuildId", DataTypeEnum.adInteger, ParameterDirectionEnum.adParamInput, 1, GuildId)
        
    Set Rs = ExecuteSqlCommand(Cmd)
    
    Call LoadGuildFromRS(GuildIndex, Rs)
    
    Rs.Close
    Set Rs = Nothing
    Set Cmd = Nothing
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadGuildDB de modDB_Guild.bas")
End Sub
    
Public Sub LoadGuildFromRS(ByVal GuildIndex As Integer, ByRef Rs As Recordset)

On Error GoTo ErrHandler


    With GuildList(GuildIndex)
        If Not Rs.EOF Then
            'guild info
            .Name = CStr(Rs.Fields("NAME"))
            .Alignment = CInt(Rs.Fields("ALIGNMENT"))
            .IdLeader = CLng(Rs.Fields("ID_LEADER"))
            .IdRightHand = CLng(IIf(IsNull(Rs.Fields("ID_RIGHTHAND")), 0, Rs.Fields("ID_RIGHTHAND")))
            .MemberCount = CInt(Rs.Fields("MEMBER_COUNT"))
            .ContributionAvailable = CLng(Rs.Fields("CONTRIBUTION_EARNED"))
            .ContributionEarned = CLng(Rs.Fields("CONTRIBUTION_AVAILABLE"))
            .CreationTime = CDate(Rs.Fields("CREATION_DATE"))
            .Description = CStr(Rs.Fields("DESCRIPTION"))
            .CurrentQuest.IdQuest = CInt(IIf(IsNull(Rs.Fields("ID_CURRENT_QUEST")), 0, Rs.Fields("ID_CURRENT_QUEST")))
            .CurrentQuest.CurrentStage = CInt(IIf(IsNull(Rs.Fields("CURRENT_QUEST_STAGE")), 0, Rs.Fields("CURRENT_QUEST_STAGE")))
            .IdGuild = CLng(Rs.Fields("ID_GUILD"))
            .CurrentQuest.StartedDate = CDate(IIf(IsNull(Rs.Fields("QUEST_STARTED_DATE")), 0, Rs.Fields("QUEST_STARTED_DATE")))
            .Status = CInt(Rs.Fields("STATUS"))
            .BankGold = CLng(Rs.Fields("BANK_GOLD"))
            .IdDefaultRole = CLng(Rs.Fields("ID_ROLE_NEW_MEMBERS"))
            .RankingPoints = CInt(Rs.Fields("RANKING_POINTS"))
            
            'CURRENT_QUEST_SECONDS_LEFT
            .CurrentQuest.SecondsLeft = CLng(Rs.Fields("CURRENT_QUEST_SECONDS_LEFT"))

            ' Load roles
            Set Rs = Rs.NextRecordset()
            Call LoadGuildRolesFromRS(GuildIndex, Rs)

            'members
            Set Rs = Rs.NextRecordset()
            Call LoadGuildMembersFromRS(GuildIndex, Rs)
            
            'role permission
            Set Rs = Rs.NextRecordset()
            Call LoadGuildRolesPermissionFromRS(GuildIndex, Rs)
            
            'quest
            Set Rs = Rs.NextRecordset()
            Call LoadGuildQuestCompletedFromRS(GuildIndex, Rs)
            
            'upgrade
            Set Rs = Rs.NextRecordset()
            Call LoadGuildUpgradeFromRS(GuildIndex, Rs)
            
            'bank
            Set Rs = Rs.NextRecordset()
            Call LoadGuildBankFromRS(GuildIndex, Rs)
            
            'quest status
            Set Rs = Rs.NextRecordset()
            Call LoadGuildQuestStatus(GuildIndex, Rs)
            
            
            ReDim .Invitations(modGuild_Functions.INVITATION_MAX_COUNT)
        End If
        
        
    End With

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadGuildFromRS de modDB_Guild.bas")
End Sub

Public Sub LoadGuildMembersFromRS(ByVal GuildIndex As Integer, ByRef Rs As Recordset)

On Error GoTo ErrHandler

    Dim I, Size As Integer
    
    Size = SizeRS(Rs)
    
    If Size = 0 Then Exit Sub
    
    I = 1
    ReDim Preserve GuildList(GuildIndex).Members(1 To Size) As GuildMemberType
    
    
    While Not Rs.EOF
        With GuildList(GuildIndex).Members(I)
            .IdUser = CLng(Rs.Fields("ID_USER"))
            .NameUser = CStr(Rs.Fields("NAME"))
            .ContributionEarner = CLng(Rs.Fields("CONTRIBUTION_EARNED"))
            .JoinDate = CDate(Rs.Fields("JOIN_DATE"))
            .IdRole = CInt(Rs.Fields("ID_ROLE"))

            .RoleIndex = GetRoleIndexFromRoleId(GuildIndex, .IdRole)
            
            .RoleAssignedBy = CLng(Rs.Fields("ROLE_ASSIGNED_BY"))
            
            .IsDirty = False

            I = I + 1
            Rs.MoveNext
        End With
    Wend
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadGuildMembersFromRS de modDB_Guild.bas")
End Sub
    
Public Sub InitializeGuildRolePermissions(ByVal GuildIndex As Integer, ByVal RoleIndex As Integer)
    Dim J As Integer
    With GuildList(GuildIndex).Roles(RoleIndex)
        ReDim Preserve .RolePermission(1 To UBound(PermissionConfig)) As PermissionType
        
        For J = 1 To UBound(PermissionConfig)
            .PermissionCount = 0
            .RolePermission(J).IdPermission = PermissionConfig(J).IdPermission
            .RolePermission(J).IsEnabled = False
            .RolePermission(J).Key = PermissionConfig(J).Key
            .RolePermission(J).PermissionName = PermissionConfig(J).PermissionName
        Next J
    End With
End Sub

Public Sub LoadGuildRolesFromRS(ByVal GuildIndex As Integer, ByRef Rs As Recordset)

On Error GoTo ErrHandler
    
    Dim I As Integer
    Dim J As Integer
    Dim Size As Integer
    
    Size = SizeRS(Rs)
    
    If Size = 0 Then Exit Sub
    
    ReDim Preserve GuildList(GuildIndex).Roles(1 To Size) As GuildRoleType
    
    I = 1
    While Not Rs.EOF
        With GuildList(GuildIndex).Roles(I)
            .IdRole = CInt(Rs.Fields("ID_ROLE"))
            .RoleName = CStr(Rs.Fields("ROLE_NAME"))
            .IsDeleteable = CByte(Rs.Fields("DELETABLE"))
            
            Call InitializeGuildRolePermissions(GuildIndex, I)
            
            .IsDirty = False
            
            I = I + 1
            Rs.MoveNext
        End With
    Wend
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadGuildRolesFromRS de modDB_Guild.bas")
End Sub

Public Sub LoadGuildPermissionDB()

On Error GoTo ErrHandler
  
    
    Dim Sql As String
    Dim Rs As Recordset
    Dim I, Size As Integer

    Sql = _
        "SELECT " & _
            "ID_PERMISSION, " & _
            "GUILD_PERMISSION.KEY, " & _
            "PERMISSION_NAME " & _
        "FROM " & _
            "GUILD_PERMISSION "
    
    Set Rs = ExecuteSql(Sql)
        
    Size = SizeRS(Rs)
    
    If Size = 0 Then Exit Sub
    
    
    ReDim Preserve PermissionConfig(1 To Size) As PermissionType
    I = 1
    While Not Rs.EOF
        With PermissionConfig(I)
                .IdPermission = CLng(Rs.Fields("ID_PERMISSION"))
                .Key = CStr(Rs.Fields("KEY"))
                .PermissionName = CStr(Rs.Fields("PERMISSION_NAME"))
        End With
        I = I + 1
        Rs.MoveNext
    Wend
    
    Rs.Close
    Set Rs = Nothing
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadGuildPermissionDB de modDB_Guild.bas")
End Sub

Function SizeRS(ByRef Rs As Recordset) As Integer

On Error GoTo ErrHandler
  
    If Rs.EOF Then
      SizeRS = 0
    Else
      Rs.MoveLast
      SizeRS = Rs.RecordCount
      Rs.MoveFirst
    End If
    
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function SizeRS de modDB_Guild.bas")
End Function


Function GuildIdListDB() As Recordset

On Error GoTo ErrHandler
    
    Dim Sql As String
    Dim Rs As Recordset

    Sql = _
        "SELECT " & _
            "ID_GUILD " & _
        "FROM " & _
            "GUILD_INFO ORDER BY ID_GUILD ASC"
    
    Set GuildIdListDB = ExecuteSql(Sql)
    
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub GuildIdListDB de modDB_Guild.bas")
End Function

Public Sub LoadGuildBankFromRS(ByVal GuildIndex As Integer, ByRef Rs As Recordset)

On Error GoTo ErrHandler

    Dim I, Size As Integer
    

    ReDim Preserve GuildList(GuildIndex).Bank(1 To GetLimitOfGuildBankSlots(GuildIndex)) As GuildBankType
    
    I = 1
    While Not Rs.EOF
        With GuildList(GuildIndex).Bank(I)
            .Slot = CInt(Rs.Fields("SLOT"))
            .IdObject = CInt(Rs.Fields("ID_OBJ"))
            .Amount = CInt(Rs.Fields("AMOUNT"))
            .Box = CInt(Rs.Fields("ID_BANKBOX"))
            
            I = I + 1
            Rs.MoveNext
        End With
    Wend
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadGuildBankFromRS de modDB_Guild.bas")
End Sub
    
Public Sub LoadGuildQuestCompletedFromRS(ByVal GuildIndex As Integer, ByRef Rs As Recordset)

On Error GoTo ErrHandler

    Dim I, Size As Integer
    
    Size = SizeRS(Rs)
    
    GuildList(GuildIndex).QuestCompletedCount = Size
    
    If Size = 0 Then Exit Sub
    
    ReDim Preserve GuildList(GuildIndex).QuestCompleted(1 To Size) As GuildQuestCompletedType
    
    I = 1
    While Not Rs.EOF
        With GuildList(GuildIndex).QuestCompleted(I)
            .IdContribution = CLng(Rs.Fields("ID_CONTRIBUTION"))
            .IdQuest = CInt(Rs.Fields("ID_QUEST"))
            .CompletedDate = CDate(Rs.Fields("COMPLETED_DATE"))
            .MembersContributed = CLng(Rs.Fields("MEMBERS_CONTRIBUTED"))
            .ContributionGained = CLng(Rs.Fields("CONTRIBUTION_GAINED"))
            
            I = I + 1
            Rs.MoveNext
        End With
    Wend
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadGuildQuestCompletedFromRS de modDB_Guild.bas")
End Sub

Public Sub LoadGuildUpgradeFromRS(ByVal GuildIndex As Integer, ByRef Rs As Recordset)

On Error GoTo ErrHandler

    Dim I, Size As Integer
    Dim EnabledUpgrade As Integer
    
    Size = SizeRS(Rs)
    
    If Size = 0 Then Exit Sub
    
    ReDim Preserve GuildList(GuildIndex).Upgrades(1 To Size) As GuildUpgradeType
    
    I = 1
    While Not Rs.EOF
        With GuildList(GuildIndex)
            EnabledUpgrade = CLng(Rs.Fields("ID_UPGRADE"))
            
            If GuildConfiguration.GuildUpgradesList(EnabledUpgrade).IsEnabled = True Then
            
                .Upgrades(I).IdUpgrade = EnabledUpgrade
                .Upgrades(I).UpgradeLevel = CInt(Rs.Fields("UPGRADE_LEVEL"))
                .Upgrades(I).UpgradeDate = CDate(Rs.Fields("UPGRADE_DATE"))
                .Upgrades(I).UpgradeBy = CInt(Rs.Fields("UPGRADED_BY"))
                .Upgrades(I).IsEnabled = CBool(Rs.Fields("ENABLED"))
                
                Call AddGuildUpgradeEffect(GuildIndex, .Upgrades(I).IdUpgrade)
                
                I = I + 1
                
            End If
            
            Rs.MoveNext
        End With
    Wend
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadGuildUpgradeFromRS de modDB_Guild.bas")
End Sub

Public Sub LoadGuildQuestStatus(ByVal GuildIndex As Integer, ByRef Rs As Recordset)

On Error GoTo ErrHandler

    If GuildList(GuildIndex).CurrentQuest.IdQuest = 0 Then Exit Sub
    Dim I As Integer, J As Integer
    Dim Size As Integer
    Dim RequirementType As Integer
    Dim RequirementIndex As Integer
    Dim Quantity As Integer

    If GuildList(GuildIndex).CurrentQuest.CurrentStage = 0 Then
        'an error ocurred while saving data, we must set current stage = 1
        GuildList(GuildIndex).CurrentQuest.CurrentStage = 1
        Call LogError("Error: GuildId (" & GuildList(GuildIndex).IdGuild & ") tiene quest activa pero no stage. Nro stage seteado a 1")
    End If

    'initialize quest stage data
    Call ChangeGuildQuestStage(GuildIndex, GuildList(GuildIndex).CurrentQuest.IdQuest, GuildList(GuildIndex).CurrentQuest.CurrentStage, False)
    
    GuildList(GuildIndex).CurrentQuest.ServerStartedDate = Now()
    
    Size = SizeRS(Rs)
    If Size = 0 Then Exit Sub

    I = 1
    While Not Rs.EOF
        With GuildList(GuildIndex).CurrentQuest
            
            RequirementType = CInt(Rs.Fields("REQUIREMENT_TYPE"))
            Quantity = CInt(Rs.Fields("QUANTITY_COMPLETED"))
            RequirementIndex = CInt(IIf(IsNull(Rs.Fields("REQUIREMENT_INDEX")), 0, Rs.Fields("REQUIREMENT_INDEX")))

            Select Case RequirementType
                Case eQuestRequirementDb.NpcKill
                    For J = 1 To .CurrentNpcKillsQuantity
                        If RequirementIndex = .CurrentNpcKills(J).NpcIndex Then
                            .CurrentNpcKills(J).Quantity = .CurrentNpcKills(J).Quantity + Quantity
                        End If
                    Next J

                Case eQuestRequirementDb.ObjCollect
                    Dim Rest As Long
                    Call modRequiredObjectList.RequiredObjectListTryAdd(.CurrentObjectList, RequirementIndex, Quantity, Rest)

                Case eQuestRequirementDb.FragNeutral
                    .CurrentFrags.Neutral.Qty = Quantity

                Case eQuestRequirementDb.FragArmada
                    .CurrentFrags.Army.Qty = Quantity

                Case eQuestRequirementDb.FragLegion
                    .CurrentFrags.Legion.Qty = Quantity

            End Select

            I = I + 1
            Rs.MoveNext
        End With
    Wend

  Exit Sub

ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadGuildQuestStatus de modDB_Guild.bas")
End Sub
Function LoadRolePermissionQuery(ByVal RoleId As Integer) As Recordset

On Error GoTo ErrHandler
  
    Dim Sql As String
    Dim Rs As Recordset

    Sql = _
        "SELECT " & _
            "ID_ROLE, " & _
            "ID_PERMISSION " & _
        "FROM " & _
            "GUILD_ROLE_PERMISSION " & _
        "WHERE ID_ROLE = '" & CStr(RoleId) & "' "
    
    Set LoadRolePermissionQuery = ExecuteSql(Sql)
    
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function LoadRolePermissionQuery de modDB_Guild.bas")
End Function


Public Sub LoadGuildRolesPermissionFromRS(ByVal GuildIndex As Integer, ByRef Rs As Recordset)
    
On Error GoTo ErrHandler
    
    Dim I, J, Size As Integer
    
    Size = SizeRS(Rs)
    
    If Size = 0 Then Exit Sub
    
    I = 1
    While Not Rs.EOF
        With GuildList(GuildIndex)
            For J = 1 To UBound(.Roles)
                If .Roles(J).IdRole = CInt(Rs.Fields("ID_ROLE")) Then
                    .Roles(J).RolePermission(CInt(Rs.Fields("ID_PERMISSION"))).IsEnabled = True
                    .Roles(J).RolePermission(CInt(Rs.Fields("ID_PERMISSION"))).Key = Rs.Fields("KEY")
                    .Roles(J).PermissionCount = .Roles(J).PermissionCount + 1
                End If
            Next J
            .IsDirty = False
            
            I = I + 1
            Rs.MoveNext
        End With
    Wend
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadGuildRolesPermissionFromRS de modDB_Guild.bas")
End Sub

Public Sub GuildMemberAddDB(ByVal GuildId As Integer, ByVal MemberRequestId As Long, ByVal TargetRequestId As Long, ByVal RolMemberId As Integer)
On Error GoTo ErrHandler
    
    Dim Rs As Recordset
    Dim Size As Integer
    Dim GuildIndex As Integer
    Dim Cmd As ADODB.Command
    
    Set Cmd = New ADODB.Command
    
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "sp_GuildMemberAdd"
    
    Cmd.Parameters.Append Cmd.CreateParameter("GuildId", DataTypeEnum.adInteger, ParameterDirectionEnum.adParamInput, 1, GuildId)
    Cmd.Parameters.Append Cmd.CreateParameter("MemberRequestId", DataTypeEnum.adInteger, ParameterDirectionEnum.adParamInput, 1, MemberRequestId)
    Cmd.Parameters.Append Cmd.CreateParameter("TargetRequestId", DataTypeEnum.adInteger, ParameterDirectionEnum.adParamInput, 1, TargetRequestId)
    Cmd.Parameters.Append Cmd.CreateParameter("RolMemberId", DataTypeEnum.adTinyInt, ParameterDirectionEnum.adParamInput, 1, RolMemberId)
    
    Set Rs = ExecuteSqlCommand(Cmd)
    
    Size = SizeRS(Rs)
    
    If Size = 0 Then Exit Sub
    
    GuildIndex = GuildIndexOf(GuildId)
    
    Size = UBound(GuildList(GuildIndex).Members) + 1
    
    ReDim Preserve GuildList(GuildIndex).Members(1 To Size) As GuildMemberType
    
    GuildList(GuildIndex).MemberCount = Size
    
    With GuildList(GuildIndex).Members(Size)
        .IdUser = CLng(Rs.Fields("ID_USER"))
        .NameUser = CStr(Rs.Fields("NAME"))
        .ContributionEarner = CLng(Rs.Fields("CONTRIBUTION_EARNED"))
        .JoinDate = CDate(Rs.Fields("JOIN_DATE"))
        .IdRole = CInt(Rs.Fields("ID_ROLE"))
        .RoleIndex = Size
        .RoleAssignedBy = CLng(Rs.Fields("ROLE_ASSIGNED_BY"))
        
        .IsDirty = False
    End With
    
    Rs.Close
    
    Set Rs = Nothing
    Set Cmd = Nothing
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub GuildMemberAddDB de modDB_Guild.bas")
End Sub

Public Sub GuildMemberUpdateDB(ByVal GuildId As Integer, ByVal MemberRequestId As Long, ByVal TargetRequestId As Long, ByVal RolMemberId As Integer, ByVal Contribution As Integer)
On Error GoTo ErrHandler
       
    Dim Rs As Recordset
    Dim Cmd As ADODB.Command
    Dim GuildIndex As Integer
    
    Set Cmd = New ADODB.Command
    
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "sp_GuildMemberUpdate"
    
    Cmd.Parameters.Append Cmd.CreateParameter("GuildId", DataTypeEnum.adInteger, ParameterDirectionEnum.adParamInput, 1, GuildId)
    Cmd.Parameters.Append Cmd.CreateParameter("MemberRequestId", DataTypeEnum.adInteger, ParameterDirectionEnum.adParamInput, 1, MemberRequestId)
    Cmd.Parameters.Append Cmd.CreateParameter("TargetRequestId", DataTypeEnum.adInteger, ParameterDirectionEnum.adParamInput, 1, TargetRequestId)
    Cmd.Parameters.Append Cmd.CreateParameter("RolMemberId", DataTypeEnum.adTinyInt, ParameterDirectionEnum.adParamInput, 1, RolMemberId)
    Cmd.Parameters.Append Cmd.CreateParameter("Contribution", DataTypeEnum.adInteger, ParameterDirectionEnum.adParamInput, 1, Contribution)
    
    Set Rs = ExecuteSqlCommand(Cmd)
    
    Set Rs = Nothing
    Set Cmd = Nothing
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub GuildMemberUpdateDB de modDB_Guild.bas")
End Sub

Public Sub GuildMemberDeleteDB(ByVal GuildId As Integer, ByVal TargetRequestId As Long)
On Error GoTo ErrHandler
       
    Dim Rs As Recordset
    Dim Cmd As ADODB.Command
    Dim GuildIndex As Integer
    
    Set Cmd = New ADODB.Command
    
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "sp_GuildMemberDelete"
    
    Cmd.Parameters.Append Cmd.CreateParameter("GuildId", DataTypeEnum.adInteger, ParameterDirectionEnum.adParamInput, 1, GuildId)
    Cmd.Parameters.Append Cmd.CreateParameter("TargetRequestId", DataTypeEnum.adInteger, ParameterDirectionEnum.adParamInput, 1, TargetRequestId)
    
    Set Rs = ExecuteSqlCommand(Cmd)
    
    Set Rs = Nothing
    Set Cmd = Nothing
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub GuildMemberDeleteDB de modDB_Guild.bas")
End Sub

Public Sub SaveGuildBankDB(ByVal GuildIndex As Integer)
On Error GoTo ErrHandler:

    Dim query As String
    Dim Slot As Long
    Dim Box As Integer

    Box = 1 ' until the logic of the boxes is done

    With GuildList(GuildIndex)
        query = _
            "DELETE FROM GUILD_BANK " & _
            "WHERE ID_GUILD='" & .IdGuild & "' AND  ID_BANKBOX= '" & CStr(Box) & "' "
            
        Call ExecuteSql(query)
        
        ' Save the Guild Gold and password
        query = "UPDATE GUILD_INFO SET BANK_GOLD = " & CStr(.BankGold) & _
                " WHERE ID_GUILD = " & CStr(.IdGuild)

        Call ExecuteSql(query)

        ' Save the Guild Bank items
        For Slot = 1 To GetLimitOfGuildBankSlots(GuildIndex)
                If .Bank(Slot).IdObject <> 0 Then
                    query = _
                        "INSERT INTO GUILD_BANK (ID_GUILD, SLOT, ID_BANKBOX, ID_OBJ, AMOUNT) VALUES ('" & _
                            CStr(.IdGuild) & "','" & _
                            CStr(Slot) & "','" & _
                            CStr(Box) & "','" & _
                            CStr(.Bank(Slot).IdObject) & "','" & _
                            CStr(.Bank(Slot).Amount) & "' " & _
                        ")"
                    
                    Call ExecuteSql(query)
                End If
        Next Slot
    End With
    
    Exit Sub
    
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function SaveGuildBankDB de modGuild_Functions.bas")
End Sub

Public Sub RoleModifyFromDB(ByVal RoleId As Integer, ByVal RoleName As String, ByVal Permissions As String)
    On Error GoTo ErrHandler
      
    Dim Cmd As ADODB.Command
    
    Set Cmd = New ADODB.Command
    
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "sp_ModifyRole"

    Cmd.Parameters.Append Cmd.CreateParameter("RoleId", DataTypeEnum.adInteger, ParameterDirectionEnum.adParamInput, 1, RoleId)
    Cmd.Parameters.Append Cmd.CreateParameter("RoleName", DataTypeEnum.adBSTR, ParameterDirectionEnum.adParamInput, 1, RoleName)
    Cmd.Parameters.Append Cmd.CreateParameter("Permissions", DataTypeEnum.adBSTR, ParameterDirectionEnum.adParamInput, 1, Permissions)
        
    Call ExecuteSqlCommand(Cmd)
    
    Set Cmd = Nothing
    
    Exit Sub
    
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function RoleModifyFromDB de modGuild_DB.bas")
End Sub

Public Sub RoleDeleteFromDB(ByVal RoleId As Integer, ByVal GuildId As String)
    On Error GoTo ErrHandler
      
    Dim Cmd As ADODB.Command
    
    Set Cmd = New ADODB.Command
    
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "sp_GuildRole_Delete"

    Cmd.Parameters.Append Cmd.CreateParameter("RoleId", DataTypeEnum.adInteger, ParameterDirectionEnum.adParamInput, 1, RoleId)
    Cmd.Parameters.Append Cmd.CreateParameter("GuildId", DataTypeEnum.adInteger, ParameterDirectionEnum.adParamInput, 1, GuildId)
        
    Call ExecuteSqlCommand(Cmd)
    
    Set Cmd = Nothing
    
    Exit Sub
    
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function RoleDeleteFromDB de modGuild_DB.bas")
End Sub

Public Function CreateRoleFromDB(ByVal GuildId As Integer, ByVal RoleName As String, ByVal Permissions As String) As Integer
    On Error GoTo ErrHandler
      
    Dim Cmd As ADODB.Command
    Dim Rs As ADODB.Recordset
    Dim Size As Integer
    Set Cmd = New ADODB.Command
    
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "sp_CreateRole"

    Cmd.Parameters.Append Cmd.CreateParameter("GuildId", DataTypeEnum.adInteger, ParameterDirectionEnum.adParamInput, 1, GuildId)
    Cmd.Parameters.Append Cmd.CreateParameter("RoleName", DataTypeEnum.adBSTR, ParameterDirectionEnum.adParamInput, 1, RoleName)
    Cmd.Parameters.Append Cmd.CreateParameter("Permissions", DataTypeEnum.adBSTR, ParameterDirectionEnum.adParamInput, 1, Permissions)
        
    Set Rs = ExecuteSqlCommand(Cmd)
    Size = SizeRS(Rs)
    
    If Size = 0 Then
        CreateRoleFromDB = -1
        Exit Function
    End If
    
    CreateRoleFromDB = CInt(Rs.Fields("ID_ROLE"))

    Rs.Close
    Set Rs = Nothing
    Set Cmd = Nothing
    
    Exit Function
    
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function CreateRoleFromDB de modGuild_DB.bas")
End Function

Public Sub LoadRolePermissionsFromDB(ByVal GuildId As Integer, ByVal RoleId As String)
    On Error GoTo ErrHandler

    Dim Sql As String
    Dim Rs As Recordset
    Dim I As Integer
    Dim Size As Integer
    Dim GuildIndex As Integer
    
    GuildIndex = GuildIndexOf(GuildId)

    Sql = "" & _
        "SELECT " & _
            "GUILD_ROLE_PERMISSION.ID_ROLE, " & _
            "GUILD_ROLE_PERMISSION.ID_PERMISSION, " & _
            "GUILD_PERMISSION.KEY " & _
        "FROM GUILD_ROLE_PERMISSION " & _
            "INNER JOIN GUILD_PERMISSION ON GUILD_PERMISSION.ID_PERMISSION = GUILD_ROLE_PERMISSION.ID_PERMISSION " & _
        "WHERE GUILD_ROLE_PERMISSION.ID_ROLE = " & RoleId
        
    Set Rs = ExecuteSql(Sql)
    Size = SizeRS(Rs)
    If Size = 0 Then Exit Sub
    Call LoadGuildRolesPermissionFromRS(GuildIndex, Rs)
        
    Rs.Close
    Set Rs = Nothing

    Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function LoadRoleFromDB de modGuild_DB.bas")

End Sub


Public Sub AssignRoleFromDB(ByVal UserIdAssigner As Long, ByVal GuildId As Integer, ByVal TargetUserId As Long, ByVal RoleId As Integer)
    On Error GoTo ErrHandler

    Dim Sql As String
    Dim I As Integer
    
    Sql = "" & _
            "UPDATE GUILD_MEMBER " & _
            "  SET ID_ROLE = " & RoleId & _
            ", ROLE_ASSIGNED_BY = " & UserIdAssigner & _
            " WHERE ID_USER = " & TargetUserId & _
            " AND ID_GUILD = " & GuildId

    Call ExecuteSql(Sql)
    
    Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function AssignRoleFromDB de modGuild_DB.bas")

End Sub


Public Sub UpdateGuildAlignment(ByVal GuildId As Integer, ByVal Alignment As eGuildAlignment)
    On Error GoTo ErrHandler

    Dim Sql As String
    Dim I As Integer
    
    Sql = "" & _
            "UPDATE GUILD_INFO " & _
            "  SET ALIGNMENT = " & Alignment & _
            " WHERE ID_GUILD = " & GuildId

    Call ExecuteSql(Sql)
    
    Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function UpdateGuildAlignment de modGuild_DB.bas")

End Sub


Public Sub UpdateGuildLeadership(ByVal GuildId As Integer, ByVal LeaderId As Long, ByVal RightHand As Long)
    On Error GoTo ErrHandler

    Dim Sql As String
    Dim I As Integer
    
    Sql = "" & _
            "UPDATE GUILD_INFO " & _
            "  SET ID_LEADER = " & LeaderId & _
            ", ID_RIGHTHAND = " & RightHand & _
            " WHERE ID_GUILD = " & GuildId

    Call ExecuteSql(Sql)
    
    Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function AssignRoleFromDB de modGuild_DB.bas")

End Sub

Public Sub AddGuildUpgradeDB(ByVal GuildId As Integer, ByVal UserId As Long, ByVal UpgradeId As Integer)
On Error GoTo ErrHandler

    Dim Sql As String
    Dim I As Integer
    
    Sql = "INSERT INTO GUILD_UPGRADE" & _
            "(ID_GUILD,ID_UPGRADE,UPGRADE_LEVEL,UPGRADE_DATE,UPGRADED_BY,ENABLED) " & _
            "  VALUES ( " & GuildId & _
            ", " & UpgradeId & _
            ", " & str(0) & _
            ", NOW() " & _
            ", " & UserId & _
            ", " & 1 & ")"

    Call ExecuteSql(Sql)
    
    Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function AddGuildUpgradeDB de modGuild_DB.bas")
End Sub


Function LoadGuildUpgradeDB(ByVal GuildId As Integer, ByVal UpgradeId As Integer) As Recordset

On Error GoTo ErrHandler
    
    Dim Sql As String

    Sql = _
        "SELECT " & _
            "ID_GUILD,ID_UPGRADE,UPGRADE_LEVEL,UPGRADE_DATE,UPGRADED_BY,ENABLED " & _
        "FROM GUILD_UPGRADE " & _
        "WHERE ID_GUILD = " & GuildId & " AND ID_UPGRADE = " & UpgradeId
    
    Set LoadGuildUpgradeDB = ExecuteSql(Sql)
    
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadGuildUpgradeDB de modDB_Guild.bas")
End Function

Public Sub AcceptQuest(ByVal GuildId As Long, ByVal QuestId As Long, ByVal StageNumber As Integer, ByVal QuestStartDate As Date, ByVal SecondsLeft As Long)
    On Error GoTo ErrHandler

    Dim Cmd As ADODB.Command

    Set Cmd = New ADODB.Command

    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "sp_AcceptGuildQuest"

    Cmd.Parameters.Append Cmd.CreateParameter("GuildId", DataTypeEnum.adInteger, ParameterDirectionEnum.adParamInput, 1, GuildId)
    Cmd.Parameters.Append Cmd.CreateParameter("QuestId", DataTypeEnum.adInteger, ParameterDirectionEnum.adParamInput, 1, QuestId)
    Cmd.Parameters.Append Cmd.CreateParameter("StageId", DataTypeEnum.adInteger, ParameterDirectionEnum.adParamInput, 1, StageNumber)
    Cmd.Parameters.Append Cmd.CreateParameter("StartDate", DataTypeEnum.adVarChar, ParameterDirectionEnum.adParamInput, 50, FormatDateDB(QuestStartDate))
    Cmd.Parameters.Append Cmd.CreateParameter("SecondsLeft", DataTypeEnum.adInteger, ParameterDirectionEnum.adParamInput, 1, SecondsLeft)
    
    Call ExecuteSqlCommand(Cmd)
    
    Set Cmd = Nothing

    Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function AcceptQuest de modGuild_DB.bas")

End Sub

Public Sub DeleteCurrentQuestStatus(ByVal GuildId As Long)
On Error GoTo ErrHandler

    Dim Sql As String

    Sql = "DELETE FROM GUILD_CURRENT_QUEST_STAGE WHERE ID_GUILD = " & GuildId

    Call ExecuteSql(Sql)

 Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function DeleteCurrentQuestStatus de modGuild_DB.bas")

End Sub

Public Sub DeleteCurrentQuest(ByVal GuildId As Long)
    On Error GoTo ErrHandler

    Dim Cmd As ADODB.Command

    Set Cmd = New ADODB.Command

    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "sp_DeleteGuildCurrentQuest"

    Cmd.Parameters.Append Cmd.CreateParameter("GuildId", DataTypeEnum.adInteger, ParameterDirectionEnum.adParamInput, 1, GuildId)

    Call ExecuteSqlCommand(Cmd)

    Set Cmd = Nothing

    Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function DeleteCurrentQuest de modGuild_DB.bas")

End Sub

Public Sub UpdateCurrentQuestStage(ByVal GuildId As Long, ByVal StageNumber As Integer)
    On Error GoTo ErrHandler

    Dim Sql As String

    Sql = "" & _
            "UPDATE GUILD_INFO " & _
            "  SET CURRENT_QUEST_STAGE = " & StageNumber & _
            " WHERE ID_GUILD = " & GuildId

    Call ExecuteSql(Sql)

    Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function UpdateCurrentQuestStage de modGuild_DB.bas")
End Sub

Public Sub UpdateGuildStats(ByVal GuildIndex As Integer)
 On Error GoTo ErrHandler
    Dim Sql As String
    With GuildList(GuildIndex)
        
         Sql = "" & _
            "UPDATE GUILD_INFO " & _
            "  SET CONTRIBUTION_EARNED = " & .ContributionEarned & _
            " , CONTRIBUTION_AVAILABLE = " & .ContributionAvailable & _
            " , BANK_GOLD = " & .BankGold & _
             " , MEMBER_COUNT = " & .MemberCount & _
             " , ID_LEADER = " & .IdLeader & _
             " , ID_RIGHTHAND = " & .IdRightHand & _
             " , ALIGNMENT = " & .Alignment & _
             " , STATUS = " & .Status & _
             " , ID_CURRENT_QUEST = " & .CurrentQuest.IdQuest & _
             " , CURRENT_QUEST_STAGE = " & .CurrentQuest.CurrentStage & _
             " , QUEST_STARTED_DATE = '" & Format(.CurrentQuest.StartedDate, "yyyy-mm-dd ttttt") & "'" & _
             " , CURRENT_QUEST_SECONDS_LEFT = " & modQuestSystem.GetQuestStageRemainingTime(GuildIndex) & _
             " , QUEST_STARTED_DATE = '" & Format(.CurrentQuest.StartedDate, "yyyy-mm-dd ttttt") & "'" & _
             " , ID_ROLE_NEW_MEMBERS = " & .IdDefaultRole & _
             " , RANKING_POINTS = " & .RankingPoints & _
            " WHERE ID_GUILD = " & .IdGuild
            
    Call ExecuteSql(Sql)
        
    Exit Sub
        
    End With
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function UpdateGuildStats de modGuild_DB.bas")
End Sub
Public Sub FinishQuest_DB(ByVal GuildIndex As Integer)
 On Error GoTo ErrHandler

    Dim QuestId As Integer
    Dim Members As Integer
    Dim Contribution As Long
    Dim GuildId As Long
    Dim Sql As String
    Dim TotalSeconds As Long
    Dim TimeLeft As Long
    Dim Cmd As ADODB.Command
    
    Set Cmd = New ADODB.Command

    QuestId = GuildList(GuildIndex).CurrentQuest.IdQuest
    Members = GuildList(GuildIndex).MemberCount
    GuildId = GuildList(GuildIndex).IdGuild
    
    TimeLeft = modQuestSystem.GetQuestStageRemainingTime(GuildIndex)
    
    TotalSeconds = GuildQuestList(GuildList(GuildIndex).CurrentQuest.IdQuest).duration - TimeLeft
    
    Contribution = 0
    
    If GuildList(GuildIndex).CurrentQuest.IsFirstTime Then
        Contribution = GuildQuestList(GuildList(GuildIndex).CurrentQuest.IdQuest).ContributionEarnedFirstTime
    Else
        Contribution = GuildQuestList(GuildList(GuildIndex).CurrentQuest.IdQuest).ContributionEarned
    End If

    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "sp_FinishGuildCurrentQuest"

    Cmd.Parameters.Append Cmd.CreateParameter("GuildId", DataTypeEnum.adInteger, ParameterDirectionEnum.adParamInput, 1, GuildId)
    Cmd.Parameters.Append Cmd.CreateParameter("QuestId", DataTypeEnum.adInteger, ParameterDirectionEnum.adParamInput, 1, QuestId)
    Cmd.Parameters.Append Cmd.CreateParameter("Members", DataTypeEnum.adInteger, ParameterDirectionEnum.adParamInput, 1, Members)
    Cmd.Parameters.Append Cmd.CreateParameter("Contribution", DataTypeEnum.adInteger, ParameterDirectionEnum.adParamInput, 1, Contribution)
    Cmd.Parameters.Append Cmd.CreateParameter("TotalSeconds", DataTypeEnum.adInteger, ParameterDirectionEnum.adParamInput, 1, TotalSeconds)

    Call ExecuteSqlCommand(Cmd)

    Set Cmd = Nothing

    Exit Sub

ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function FinishQuest_DB de modGuild_DB.bas")
End Sub
'TargetIndex should be a NpcIndex or ObjectIndex dependes on QuestRequierementType
Public Sub UpdateCurrentQuestStatus(ByVal GuildIndex As Long, ByVal QuestRequierementType As eQuestRequirementDb, ByVal Quantity As Integer, Optional ByVal TargetIndex As Integer)
    Dim MainSql As String
    Dim GuildId As Long
    Dim RequirementIndex As String
    
    GuildId = GuildList(GuildIndex).IdGuild
    
    If QuestRequierementType = eQuestRequirementDb.NpcKill _
        Or QuestRequierementType = eQuestRequirementDb.ObjCollect Then
        RequirementIndex = TargetIndex
    Else
        RequirementIndex = "NULL"
    End If
    
    MainSql = "" & _
        "UPDATE GUILD_CURRENT_QUEST_STAGE " & _
        "SET QUANTITY_COMPLETED=" & Quantity & " " & _
        "WHERE ID_GUILD=" & GuildId & _
        "AND REQUIREMENT_TYPE=" & QuestRequierementType & _
        "AND REQUIREMENT_INDEX=" & RequirementIndex
        
    

End Sub

Public Sub SaveCurrentQuestStatus(ByVal GuildIndex As Long)
    On Error GoTo ErrHandler

    Dim MainSql As String
    Dim InfoSql As String
    Dim ValuesSql As String
    Dim TimeLeft As Long
    Dim I As Integer
    Dim GuildId As Long

    MainSql = "" & _
        "INSERT INTO GUILD_CURRENT_QUEST_STAGE (" & _
            "ID_GUILD, " & _
            "REQUIREMENT_TYPE, " & _
            "QUANTITY_COMPLETED, " & _
            "REQUIREMENT_INDEX " & _
            ") " & _
            "VALUES "
    InfoSql = "UPDATE GUILD_INFO SET CURRENT_QUEST_SECONDS_LEFT = "
    
    GuildId = GuildList(GuildIndex).IdGuild

    With GuildList(GuildIndex).CurrentQuest
        
        TimeLeft = modQuestSystem.GetQuestStageRemainingTime(GuildIndex)
        
        InfoSql = InfoSql & TimeLeft & _
        " WHERE ID_GUILD = " & GuildId
    
        For I = 1 To .CurrentNpcKillsQuantity
            If .CurrentNpcKills(I).Quantity > 0 Then
                ValuesSql = ValuesSql & "( " & _
                        GuildId & ", " & _
                        eQuestRequirementDb.NpcKill & "," & _
                        .CurrentNpcKills(I).Quantity & "," & _
                        .CurrentNpcKills(I).NpcIndex & "),"
            End If
        Next I
        
        If .CurrentObjectList.ItemsCount > 0 Then
            For I = 0 To .CurrentObjectList.ItemsCount - 1
                If .CurrentObjectList.Items(I).Quantity > 0 Then
                ValuesSql = ValuesSql & "( " & _
                    GuildId & ", " & _
                    eQuestRequirementDb.ObjCollect & "," & _
                    .CurrentObjectList.Items(I).Quantity & "," & _
                    .CurrentObjectList.Items(I).ObjIndex & "),"
                End If
            Next I
        End If

        If .CurrentFrags.Neutral.Qty > 0 Then
            ValuesSql = ValuesSql & "( " & _
                GuildId & "," & _
                eQuestRequirementDb.FragNeutral & "," & _
                .CurrentFrags.Neutral.Qty & "," & _
                "NULL),"
        End If
        
        If .CurrentFrags.Army.Qty > 0 Then
            ValuesSql = ValuesSql & "( " & _
                GuildId & "," & _
                eQuestRequirementDb.FragArmada & "," & _
                .CurrentFrags.Army.Qty & "," & _
                "NULL),"
        End If

        If .CurrentFrags.Legion.Qty > 0 Then
            ValuesSql = ValuesSql & "( " & _
                GuildId & "," & _
                eQuestRequirementDb.FragLegion & "," & _
                .CurrentFrags.Legion.Qty & "," & _
                "NULL)"
        End If

    End With
        
    Call ExecuteSql(InfoSql)
    
    If Len(ValuesSql) > 0 Then
        ' Remove the trailing "," character from the ValuesSql string if present
        If mid$(ValuesSql, Len(ValuesSql), 1) = "," Then ValuesSql = mid$(ValuesSql, 1, Len(ValuesSql) - 1)
        
        Call ExecuteSql(MainSql & ValuesSql)
    End If

    Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function SaveCurrentQuestStatus de modGuild_DB.bas")
End Sub

Public Sub GetGuildInformationFromUserName(ByRef UserName As String, ByRef UserId As Long, ByRef GuildId As Long, ByRef GuildName As String)
    On Error GoTo ErrHandler
        
    Dim Sql As String
    Dim Rs As Recordset
    
    Sql = "SELECT ID_USER, UI.NAME AS USER_NAME, IFNULL(UI.GUILD_ID,0) AS GUILD_ID, IFNULL(GI.NAME, '') AS GUILD_NAME " & _
        "FROM user_info UI " & _
        "LEFT JOIN guild_info GI " & _
        "ON UI.GUILD_ID = GI.ID_GUILD " & _
        "WHERE UI.NAME = '" & EscapeString(UserName) & "'"
    
    Set Rs = ExecuteSql(Sql)
    
    UserId = Rs.Fields("ID_USER")
    GuildId = Rs.Fields("GUILD_ID")
    GuildName = Rs.Fields("GUILD_NAME")
    
    
    Exit Sub
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetGuildInformationFromUserName de modGuild_DB.bas")
End Sub

