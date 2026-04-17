Attribute VB_Name = "modGuild_Functions"
'**************************************************************
' modGuilds.bas
'
' Implemented by Mariano Barrou (El Oso)
' Rewritten by ZaMa. Also apply DB.
'**************************************************************

'**************************************************************************
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
'**************************************************************************
'@Folder("Guild")
Option Explicit

Public Const NOPERMISSIONOFGUILD As String = "No posees permiso para esta accion"
Public Const INVITATION_MAX_COUNT = 200
Private Const INVITATION_DEFAULT_LIFE_TIME_IN_MINUTES = 10
Public Const MAX_GUILD_NAME_LEN As Byte = 25

Sub LoadGuildConfiguration()

On Error GoTo ErrHandler
  
    Dim UpgradesQty As Integer, I As Integer
    Dim X As Integer, J As Integer
    Dim TempUpgradeQty As Integer

      
    With GuildConfiguration
        .RolsQty = Val(GetVar(DatPath & "Guild.dat", "INIT", "RolQty"))
        .MaxGuilds = Val(GetVar(DatPath & "Guild.dat", "INIT", "MaxGuildsQty"))
        
        .MemberQty = Val(GetVar(DatPath & "Guild.dat", "INIT", "MemberQty"))
        .BankSlotQty = Val(GetVar(DatPath & "Guild.dat", "INIT", "BankSlotQty"))
        .MaxGold = Val(GetVar(DatPath & "Guild.dat", "INIT", "MaxGold"))
        .BankBoxesQty = Val(GetVar(DatPath & "Guild.dat", "INIT", "BankBoxQty"))
        .MaxContribution = Val(GetVar(DatPath & "Guild.dat", "INIT", "MaxContribution"))
        
        .CreationEnabled = Val(GetVar(DatPath & "Guild.dat", "CREATION", "Enabled"))
        .CreationLeaderRequiredLevel = Val(GetVar(DatPath & "Guild.dat", "CREATION", "LeaderRequiredLevel"))
        .CreationRigthHandRequiredLevel = Val(GetVar(DatPath & "Guild.dat", "CREATION", "RigthHandRequiredLevel"))
        .CreationRequiredGold = Val(GetVar(DatPath & "Guild.dat", "CREATION", "RequiredGold"))
        
        .InvitationLifeTimeInMinutes = Val(GetVar(DatPath & "Guild.dat", "INIT", "InvitationLifeTimeInMinutes"))
        
        If .InvitationLifeTimeInMinutes <= 0 Then
            .InvitationLifeTimeInMinutes = INVITATION_DEFAULT_LIFE_TIME_IN_MINUTES
        End If
        
        .UpgradesQty = Val(GetVar(DatPath & "GuildUpgrades.dat", "INIT", "Upgrades"))


        If .UpgradesQty > 0 Then
         
            ReDim Preserve GuildConfiguration.GuildUpgradesList(1 To .UpgradesQty) As GuildConfUpgradeType
            
            For I = 1 To .UpgradesQty
                .GuildUpgradesList(I).IsEnabled = Val(GetVar(DatPath & "GuildUpgrades.dat", "UPGRADE" & I, "Enabled"))
                .GuildUpgradesList(I).Name = GetVar(DatPath & "GuildUpgrades.dat", "UPGRADE" & I, "NameUpgrade")
                .GuildUpgradesList(I).Description = GetVar(DatPath & "GuildUpgrades.dat", "UPGRADE" & I, "Description")
                .GuildUpgradesList(I).IconGraph = Val(GetVar(DatPath & "GuildUpgrades.dat", "UPGRADE" & I, "IconGrh"))
                .GuildUpgradesList(I).GoldCost = Val(GetVar(DatPath & "GuildUpgrades.dat", "UPGRADE" & I, "GoldCost"))
                .GuildUpgradesList(I).ContributionCost = Val(GetVar(DatPath & "GuildUpgrades.dat", "UPGRADE" & I, "ContributionCost"))
                Call SlitRequired(.GuildUpgradesList(I).UpgradeRequired, GetVar(DatPath & "GuildUpgrades.dat", "UPGRADE" & I, "UpdateRequired"))
                Call SlitQuestRequired(.GuildUpgradesList(I).QuestRequired, GetVar(DatPath & "GuildUpgrades.dat", "UPGRADE" & I, "QuestRequired"))
 
                'upgrade effect
                .GuildUpgradesList(I).UpgradeEffect.AddBankBox = Val(GetVar(DatPath & "GuildUpgrades.dat", "UPGRADE" & I, "AddBankBox"))
                .GuildUpgradesList(I).UpgradeEffect.AddBankSlot = Val(GetVar(DatPath & "GuildUpgrades.dat", "UPGRADE" & I, "AddBankSlot"))
                .GuildUpgradesList(I).UpgradeEffect.AddMaxContribution = Val(GetVar(DatPath & "GuildUpgrades.dat", "UPGRADE" & I, "AddMaxContribution"))
                .GuildUpgradesList(I).UpgradeEffect.AddMemberLimit = Val(GetVar(DatPath & "GuildUpgrades.dat", "UPGRADE" & I, "AddMemberLimit"))
                .GuildUpgradesList(I).UpgradeEffect.AddRolesGuild = Val(GetVar(DatPath & "GuildUpgrades.dat", "UPGRADE" & I, "AddRolesGuild"))
                .GuildUpgradesList(I).UpgradeEffect.IsChatOverHead = Val(GetVar(DatPath & "GuildUpgrades.dat", "UPGRADE" & I, "ChatOverHead"))
                .GuildUpgradesList(I).UpgradeEffect.IsFriendlyFireProtection = Val(GetVar(DatPath & "GuildUpgrades.dat", "UPGRADE" & I, "FriendlyFireProtection"))
                .GuildUpgradesList(I).UpgradeEffect.IsGuildBank = Val(GetVar(DatPath & "GuildUpgrades.dat", "UPGRADE" & I, "GuildBank"))
                .GuildUpgradesList(I).UpgradeEffect.IsSeeInvisibleGuildMember = Val(GetVar(DatPath & "GuildUpgrades.dat", "UPGRADE" & I, "SeeInvisibleGuildMember"))
            Next I
            
        End If
        
        .UpgradesGroupsQty = Val(GetVar(DatPath & "Guild.dat", "CREATION", "UpgradeGroups"))
        
        If .UpgradesGroupsQty > 0 Then
        
        ReDim .GuildUpgradeGroup(1 To .UpgradesGroupsQty)
    
        Dim UnsplitUpgrades As String
        Dim SplitUpgrades() As String
        
        For X = 1 To .UpgradesGroupsQty
            
            UnsplitUpgrades = GetVar(DatPath & "Guild.dat", "CREATION", "UpgradeGroup" & X)
            
            If UnsplitUpgrades <> vbNullString Then
                
                SplitUpgrades = Split(UnsplitUpgrades, "-")
                UpgradesQty = UBound(SplitUpgrades) + 1
                TempUpgradeQty = 1
                For J = 1 To UpgradesQty
                    If .GuildUpgradesList(CInt(SplitUpgrades(J - 1))).IsEnabled = True Then
                        ReDim Preserve .GuildUpgradeGroup(X).Upgrades(1 To TempUpgradeQty)
                        .GuildUpgradeGroup(X).Upgrades(TempUpgradeQty) = CInt(SplitUpgrades(J - 1))
                        .GuildUpgradeGroup(X).UpgradeQty = TempUpgradeQty
                        TempUpgradeQty = TempUpgradeQty + 1
                    End If
                Next J
            
            Else
                .GuildUpgradeGroup(X).UpgradeQty = 0
            End If
        
        Next X
    
    End If
    
    
    ' Load guild invalid names
    Call LoadInvalidNames
    
    ' Load guild reserved names
    Call LoadReservedNames
        
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadGuildConst de modGuild_Functions.bas")
End Sub

Public Sub LoadInvalidNames()
On Error GoTo ErrHandler
  
    With GuildConfiguration
        
        Dim TmpArray() As String
        Dim TmpString As String
        Dim I As Long, N As Integer
        Dim FileName As String
        
        FileName = DatPath & "GuildInvalidNames.txt"
        
        If Not FileExist(FileName, vbNormal) Then
            .InvalidNamesQty = 0
            Erase .InvalidNames
            Exit Sub
        End If
        
        N = FreeFile(1)
        
        Open FileName For Binary As #N
        
        TmpString = Space$(LOF(N))
        Get #N, , TmpString
        Close #N
        
        TmpArray() = Split(TmpString, vbCrLf)
        
        If UBound(TmpArray) < 0 Then
            .InvalidNamesQty = 0
            Erase .InvalidNames
            Exit Sub
        End If
        
        .InvalidNamesQty = UBound(TmpArray) + 1
        
        ReDim .InvalidNames(1 To .InvalidNamesQty)
        
        For I = 1 To .InvalidNamesQty
            .InvalidNames(I) = Trim$(TmpArray(I - 1))
        Next

    End With
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadInvalidNames de modGuild_Functions.bas")
    
End Sub

Public Sub LoadReservedNames()
On Error GoTo ErrHandler
  
    With GuildConfiguration
        
        Dim TmpArray() As String
        Dim TempCsv() As String
        Dim TmpString As String
        Dim I As Long, N As Integer
        Dim FileName As String
        
        FileName = DatPath & "GuildReservedNames.txt"
        
        If Not FileExist(FileName, vbNormal) Then
             .ReservedNamesQty = 0
             Erase .ReservedNames
            Exit Sub
        End If
        
        N = FreeFile(1)
        
        Open FileName For Binary As #N
        
        TmpString = Space$(LOF(N))
        Get #N, , TmpString
        Close #N
        
        TmpArray() = Split(TmpString, vbCrLf)
        
        If UBound(TmpArray) < 0 Then
            .ReservedNamesQty = 0
            Erase .ReservedNames
            Exit Sub
        End If
        
        .ReservedNamesQty = UBound(TmpArray) + 1
        
        ReDim .ReservedNames(1 To .ReservedNamesQty)
        
        For I = 1 To .ReservedNamesQty
            TmpString = Trim$(TmpArray(I - 1))
            If InStr(1, TmpString, ",") > 0 Then
                TempCsv = Split(TmpString, ",")
                
                .ReservedNames(I).AccountEmail = UCase(Trim$(TempCsv(0)))
                .ReservedNames(I).GuildName = UCase(Trim$(TempCsv(1)))
                
            Else
                .ReservedNames(I).AccountEmail = vbNullString
                .ReservedNames(I).GuildName = vbNullString
            End If
        Next

    End With
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadInvalidNames de modGuild_Functions.bas")
    
End Sub

Sub LoadGuilds()

On Error GoTo ErrHandler
  
    Dim GuildIndex As Integer
    Dim Rs As Recordset
    
    Set Rs = GuildIdListDB()
    MaxGuildQty = SizeRS(Rs)
    
    If MaxGuildQty = 0 Then Exit Sub
    
    ReDim Preserve GuildList(1 To MaxGuildQty) As GuildType
    GuildIndex = 1
    
    While Not Rs.EOF
        With GuildList(GuildIndex)
            Call LoadGuildDB(GuildIndex, CInt(Rs.Fields("ID_GUILD")))
            
            ReDim .Invitations(INVITATION_MAX_COUNT)

            If GuildList(GuildIndex).CurrentQuest.IdQuest > 0 Then
                If GuildList(GuildIndex).CurrentQuest.SecondsLeft = 0 Then
                    Call modQuestSystem.CancelCurrentGuildQuest(GuildIndex, False)
                Else
                    Call modQuestSystem.SetQuestTimeOnStateServer(GuildIndex)
                End If
            End If
        End With
        GuildIndex = GuildIndex + 1
        Rs.MoveNext
    Wend
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadGuilds de modGuild_Functions.bas")
End Sub

Public Function HasPermission(ByVal UserIndex As Integer, ByVal Perm As EGuildPermission) As Boolean

On Error GoTo ErrHandler
  
    Dim I As Integer
    Dim GuildIndex As Integer
    Dim RoleIndex As Integer
    
    HasPermission = False
    
    GuildIndex = UserList(UserIndex).Guild.GuildIndex
    
    If GuildIndex = 0 Then Exit Function
    If UserIndex = 0 Then Exit Function
    If UserList(UserIndex).Guild.GuildMemberIndex = 0 Then Exit Function
    
    With GuildList(GuildIndex)
    
        RoleIndex = .Members(UserList(UserIndex).Guild.GuildMemberIndex).RoleIndex
    
        If RoleIndex = 0 Then Exit Function
    
        HasPermission = GuildList(GuildIndex).Roles(RoleIndex).RolePermission(Perm).IsEnabled
        
    End With
    
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function HasPermission de modGuild_Functions.bas")
End Function

Public Function GuildIndexOf(ByVal GuildId As Integer) As Integer

On Error GoTo ErrHandler

    Dim I As Integer
    GuildIndexOf = 0
    
    For I = 1 To MaxGuildQty
        If (GuildList(I).IdGuild = GuildId) Then
            GuildIndexOf = I
            Exit Function
        End If
    Next I
   
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GuildIndexOf de modGuild_Functions.bas")
End Function

Public Function GuildNameIsUsed(ByVal GuildName As String) As Boolean

On Error GoTo ErrHandler

    Dim I As Integer
    
    GuildNameIsUsed = False
    
    If MaxGuildQty = 0 Then
        Exit Function
    End If
    
    For I = 1 To MaxGuildQty
        If (GuildList(I).Name = GuildName) Then
            GuildNameIsUsed = True
            Exit Function
        End If
    Next I
   
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GuildNameIsUsed de modGuild_Functions.bas")
End Function
Public Function TryGetGuildIndexByGuildId(ByVal GuildId As Long, ByRef ReturnIndex As Long) As Boolean

On Error GoTo ErrHandler

    Dim I As Integer
    
    TryGetGuildIndexByGuildId = False
    
    If GuildId = 0 Then
        Exit Function
    End If
        
    For I = 1 To MaxGuildQty
        If GuildList(I).IdGuild = GuildId Then
            ReturnIndex = I
            TryGetGuildIndexByGuildId = True
            Exit Function
        End If
    Next I
   
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function TryGetGuildIndexByGuildId de modGuild_Functions.bas")
End Function

Public Function GetGuildIndexByGuildId(ByVal GuildId As Long) As Long

On Error GoTo ErrHandler

    Dim I As Integer
    
    If GuildId = 0 Then
        GetGuildIndexByGuildId = 0
        Exit Function
    End If
        
    For I = 1 To MaxGuildQty
        If GuildList(I).IdGuild = GuildId Then
            GetGuildIndexByGuildId = I
            Exit Function
        End If
    Next I
   
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetGuildIndex de modGuild_Functions.bas")
End Function

Public Function TryGetGuildIndexByName(ByVal GuildName As String, ByRef returnIndex As Long) As Boolean

On Error GoTo ErrHandler

    Dim I As Integer
        
    For I = 1 To MaxGuildQty
        If (GuildList(I).Name = GuildName) Then
            returnIndex = I
            TryGetGuildIndexByName = True
            
            Exit Function
        End If
    Next I
    
    TryGetGuildIndexByName = False
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function TryGetGuildIndexByName de modGuild_Functions.bas")
End Function


Public Function GetGuildIndex(ByVal GuildName As String) As Long

On Error GoTo ErrHandler

    Dim I As Integer
        
    For I = 1 To MaxGuildQty
        If (GuildList(I).Name = GuildName) Then
            GetGuildIndex = I
            Exit Function
        End If
    Next I
   
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetGuildIndex de modGuild_Functions.bas")
End Function


Public Sub AddOnlineMember(ByVal UserIndex As Long)
On Error GoTo ErrHandler

    Dim GuildIndex As Integer, I As Integer, UserId As Long
    
    If UserList(UserIndex).Guild.IdGuild <= 0 Then Exit Sub
    
    
    GuildIndex = UserList(UserIndex).Guild.GuildIndex
    
    With GuildList(GuildIndex)
        .OnlineMemberCount = .OnlineMemberCount + 1
        
        ReDim Preserve .OnlineMembers(1 To .OnlineMemberCount) As GuildMembersOnline
        UserId = UserList(UserIndex).ID
        
        .OnlineMembers(.OnlineMemberCount).IdUser = UserId
        .OnlineMembers(.OnlineMemberCount).MemberUserIndex = UserIndex
        
        For I = 1 To UBound(.Members)
            If .Members(I).IdUser = UserId Then
                .OnlineMembers(.OnlineMemberCount).IdRole = .Members(I).IdRole
                Exit For
            End If
        Next I
    End With

    Exit Sub

ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function AddOnlineMember de modGuild_Functions.bas")

End Sub

Public Sub NotifyMemberConnection(ByVal UserIndex As Long, ByVal Connected As Boolean)
    Dim I As Integer, GuildIndex As Integer
    Dim Message As String
    
    If UserList(UserIndex).Guild.IdGuild <= 0 Then Exit Sub
    
    GuildIndex = UserList(UserIndex).Guild.GuildIndex
    
    For I = 1 To GuildLastMemberOnline(GuildIndex)
        With GuildList(GuildIndex).OnlineMembers(I)
            If Connected Then
                Message = UserList(UserIndex).Name & " se ha conectado"
                Call WriteGuildMemberStatusChange(.MemberUserIndex, UserList(UserIndex).ID, eChangeMember.OnlineChange, 1, 0)
            Else
                Message = UserList(UserIndex).Name & " se ha desconectado"
                Call WriteGuildMemberStatusChange(.MemberUserIndex, UserList(UserIndex).ID, eChangeMember.OnlineChange, 0, 0)
            End If
            
            If .MemberUserIndex <> UserIndex Then
                Call WriteConsoleMsg(.MemberUserIndex, Message, FontTypeNames.FONTTYPE_GUILD, eMessageType.Guild)
            End If
        End With
    Next I
    
    Exit Sub

ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function NotifyMembers de modGuild_Functions.bas")
End Sub

Public Sub RemoveOnlineMember(ByVal UserIndex As Long)
On Error GoTo ErrHandler

    Dim GuildIndex As Integer, I As Integer
    
    If UserList(UserIndex).Guild.IdGuild > 0 Then
        GuildIndex = UserList(UserIndex).Guild.GuildIndex
        With GuildList(GuildIndex)
            If .OnlineMemberCount > 1 Then
        
                For I = 1 To .OnlineMemberCount
                    If .OnlineMembers(I).IdUser = UserList(UserIndex).ID Then
                        .OnlineMembers(I).IdUser = 0
                        .OnlineMembers(I).MemberUserIndex = 0
                        .OnlineMembers(I).IdRole = 0
                        Exit For
                    End If
                Next I
                
                If I <> .OnlineMemberCount Then
                    For I = I To .OnlineMemberCount - 1
                        .OnlineMembers(I).IdUser = .OnlineMembers(I + 1).IdUser
                        .OnlineMembers(I).MemberUserIndex = .OnlineMembers(I + 1).MemberUserIndex
                        .OnlineMembers(I).IdRole = .OnlineMembers(I + 1).IdRole
                    Next I
                End If
            
                .OnlineMemberCount = .OnlineMemberCount - 1
                ReDim Preserve GuildList(GuildIndex).OnlineMembers(1 To .OnlineMemberCount) As GuildMembersOnline
            Else
            
                 Erase GuildList(GuildIndex).OnlineMembers
                 .OnlineMemberCount = 0
            
            End If
           
        End With
            
        Call ClearInvitationByUserId(GuildIndex, UserList(UserIndex).Id)
    End If
    
    Exit Sub

ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function RemoveOnlineMember de modGuild_Functions.bas")

End Sub

Public Function GuildLastMemberOnline(ByVal GuildIndex As Integer) As Integer

On Error GoTo ErrHandler

    If ((Not GuildList(GuildIndex).OnlineMembers) = -1) Then
        GuildLastMemberOnline = 0
    Else
        GuildLastMemberOnline = UBound(GuildList(GuildIndex).OnlineMembers)
    End If

  Exit Function

ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GuildLastMemberOnline de modGuild_Functions.bas")
End Function

Public Function IsMemberOnline(ByVal GuildIndex As Integer, ByVal UserId As Long) As Boolean
On Error GoTo ErrHandler

    Dim I As Integer
    Dim ret As Boolean
    
    ret = False
    
    For I = 1 To GuildLastMemberOnline(GuildIndex)
        If GuildList(GuildIndex).OnlineMembers(I).IdUser = UserId Then
            ret = True
        End If
    Next I
    
    IsMemberOnline = ret
    
    Exit Function

ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function IsMemberOnline de modGuild_Functions.bas")
End Function
Sub UserWithdrawItemFromGuildBank(ByVal UserIndex As Integer, ByVal SlotIndex As Integer, ByVal Quantity As Integer, ByVal Box As Integer, ByVal GuildIndex As Integer)

On Error GoTo ErrHandler

    Dim ObjIndex As Integer
    
    With UserList(UserIndex)
        'Dead people can't operate with the guild bank
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteErrorMsg(UserIndex, "¡¡Estás muerto!!", False)
            Exit Sub
        End If
        
        '¿El target es un NPC valido?
        If .flags.TargetNpc < 1 Then Exit Sub
        
        '¿El NPC puede comerciar?
        If Npclist(.flags.TargetNpc).NPCtype <> eNPCType.GuildMaster Then
            Call WriteErrorMsg(UserIndex, "Solo se puede retirar objetos interactuando con el Maestro de Clanes", False)
            Exit Sub
        End If
        
        If GuildList(GuildIndex).Bank(SlotIndex).IdObject = 0 Then Exit Sub
        
        If ObjData(GuildList(GuildIndex).Bank(SlotIndex).IdObject).ObjType = otQuest Then
            Call WriteErrorMsg(UserIndex, "No puedes retirar este objeto", False)
            Exit Sub
        End If
        
        If ObjData(GuildList(GuildIndex).Bank(SlotIndex).IdObject).Newbie = 1 Then
            Call WriteErrorMsg(UserIndex, "No puedes retirar este objeto", False)
            Exit Sub
        End If
    End With
        
    If Quantity < 1 Then Exit Sub

    If GuildList(GuildIndex).Bank(SlotIndex).Amount > 0 Then
    
        If Quantity > GuildList(GuildIndex).Bank(SlotIndex).Amount Then _
            Quantity = GuildList(GuildIndex).Bank(SlotIndex).Amount
            
        ObjIndex = GuildList(GuildIndex).Bank(SlotIndex).IdObject
        
        Call UserReciveObjFromGB(UserIndex, CInt(SlotIndex), Quantity, Box, GuildIndex)
        
        If ObjData(ObjIndex).Log = 1 Then
            Call LogDesarrollo(UserList(UserIndex).Name & " retiró del Guildbank de " & GuildList(GuildIndex).Name & " la cantidad " & Quantity & " " & _
                ObjData(ObjIndex).Name & "[" & ObjIndex & "]")
        End If
    End If
   

ErrHandler:

End Sub


Sub UserReciveObjFromGB(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, ByVal Quantity As Integer, ByVal Box As Integer, ByVal GuildIndex As Integer)
On Error GoTo ErrHandler
  

Dim Slot As Integer
Dim ObjectIndex As Integer

With UserList(UserIndex)
    If GuildList(GuildIndex).Bank(ObjIndex).Amount <= 0 Then Exit Sub
    
    ObjectIndex = GuildList(GuildIndex).Bank(ObjIndex).IdObject
    
    '¿Ya tiene un objeto de este tipo?
    Slot = 1
    Do Until .Invent.Object(Slot).ObjIndex = ObjectIndex And _
       .Invent.Object(Slot).Amount + Quantity <= MAX_INVENTORY_OBJS
        
        Slot = Slot + 1
        If Slot > .CurrentInventorySlots Then
            Exit Do
        End If
    Loop
    
    'Sino se fija por un slot vacio
    If Slot > .CurrentInventorySlots Then
        Slot = 1
        Do Until .Invent.Object(Slot).ObjIndex = 0
            Slot = Slot + 1

            If Slot > .CurrentInventorySlots Then
                Call WriteConsoleMsg(UserIndex, "No podés tener mas objetos.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        Loop
        .Invent.NroItems = .Invent.NroItems + 1
    End If
    
    If .Invent.Object(Slot).Amount + Quantity <= MAX_INVENTORY_OBJS Then
        .Invent.Object(Slot).ObjIndex = ObjectIndex
        .Invent.Object(Slot).Amount = .Invent.Object(Slot).Amount + Quantity
        
        Call WithdrawItemFromGB(UserIndex, CByte(ObjIndex), Quantity, Box, GuildIndex)
        
        Call UpdateUserInv(False, UserIndex, Slot)
    Else
        Call WriteConsoleMsg(UserIndex, "No podés tener mas objetos.", FontTypeNames.FONTTYPE_INFO)
    End If
End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub UserReciveObjFromGB de modGuild_Functions.bas")
End Sub

Sub WithdrawItemFromGB(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Quantity As Integer, ByVal Box As Integer, ByVal GuildIndex As Integer)
On Error GoTo ErrHandler

    Dim ObjIndex As Integer, I As Integer

    With GuildList(GuildIndex)
        ObjIndex = .Bank(Slot).IdObject
    
        .Bank(Slot).Amount = .Bank(Slot).Amount - Quantity
        
        If .Bank(Slot).Amount <= 0 Then
            .Bank(Slot).IdObject = 0
            .Bank(Slot).Amount = 0
        End If
        
        For I = 1 To GuildLastMemberOnline(GuildIndex)
            With GuildList(GuildIndex).OnlineMembers(I)
                If UserList(.MemberUserIndex).ID = .IdUser Then
                    Call UpdateGuildBankInv(False, .MemberUserIndex, Slot, GuildIndex)
                End If
            End With
        Next I
        
    End With
        
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WithdrawItemFromGB de modBanco.bas")
End Sub

Sub UpdateGuildBankInv(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal GuildIndex As Integer)
On Error GoTo ErrHandler
  

    Dim LoopC As Long
    Dim ObjIndex As Integer
    Dim Box As Integer
    Dim Amount As Integer
    Dim CanUse As Boolean
    
    With UserList(UserIndex)
        'Actualiza un solo slot
        If Not UpdateAll Then
            'Actualiza el inventario
            ObjIndex = GuildList(GuildIndex).Bank(Slot).IdObject
            Box = GuildList(GuildIndex).Bank(Slot).Box
            If ObjIndex > 0 Then
                Amount = GuildList(GuildIndex).Bank(Slot).Amount
                CanUse = General.checkCanUseItem(UserIndex, ObjIndex)
            End If
            Call WriteGuildBankChangeSlot(UserIndex, Slot, ObjIndex, Amount, Box, CanUse)
        Else
            ' Limpio todos en el cliente
            Call WriteGuildBankChangeSlot(UserIndex, 0, 0, 0, 0, True)
            
            'Actualiza todos los slots
            Call WriteGuildBankList(UserIndex)
        End If
    End With

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub UpdateBanUserInv de modBanco.bas")
End Sub

Sub UserDepositItemInGuildBank(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Quantity As Integer, ByVal Box As Integer, ByVal GuildIndex As Integer)
On Error GoTo ErrHandler

    With UserList(UserIndex)

        'Dead people can't operate with the guild bank
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteErrorMsg(UserIndex, "¡¡Estás muerto!!", False)
            Exit Sub
        End If
        
        '¿El target es un NPC valido?
        If .flags.TargetNpc < 1 Then Exit Sub
        
        '¿El NPC puede comerciar?
        If Npclist(.flags.TargetNpc).NPCtype <> eNPCType.GuildMaster Then
            Call WriteErrorMsg(UserIndex, "Solo se puede depositar objetos interactuando con el Maestro de Clanes", False)
            Exit Sub
        End If
        
        If .Invent.Object(Slot).ObjIndex = 0 Then Exit Sub
        
        If ObjData(.Invent.Object(Slot).ObjIndex).ObjType = otQuest Then
            Call WriteErrorMsg(UserIndex, "No puedes depositar este objeto")
            Exit Sub
        End If
        
        If ObjData(.Invent.Object(Slot).ObjIndex).Newbie = 1 Then
            Call WriteErrorMsg(UserIndex, "No puedes depositar este objeto")
            Exit Sub
        End If
        
        'User deposita el item del slot
        Call UserDepositItemGB(UserIndex, Slot, Quantity, Box, GuildIndex)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub UserDepositItemInGuildBank de modGuild_Functions.bas. UserIndex: " & UserIndex & ", Slot: " & Slot & ", Quantity: " & Quantity & ", Box: " & Box & ", GuildIndex: " & GuildIndex)
End Sub


Sub UserDepositItemGB(ByVal UserIndex As Integer, ByVal SlotIndex As Integer, ByVal Quantity As Integer, ByVal Box As Integer, ByVal GuildIndex As Integer)
On Error GoTo ErrHandler:

    Dim ObjIndex As Integer
    With UserList(UserIndex)
        If .Invent.Object(SlotIndex).Amount > 0 And Quantity > 0 Then
        
            If Quantity > .Invent.Object(SlotIndex).Amount Then _
                Quantity = .Invent.Object(SlotIndex).Amount
            
            ObjIndex = .Invent.Object(SlotIndex).ObjIndex
            
            'Agregamos el obj que deposita al banco
            Call UserDepositObjInGB(UserIndex, CInt(SlotIndex), Quantity, GuildIndex)
            
            If ObjData(ObjIndex).Log = 1 Then
                Call LogDesarrollo(.Name & " depositó en Guildbank de " & GuildList(GuildIndex).Name & " la cantidad " & Quantity & " " & _
                    ObjData(ObjIndex).Name & "[" & ObjIndex & "]")
            End If
        End If
    End With
    Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub UserDepositItemGB de modGuild_Functions.bas. UserIndex: " & UserIndex & ", Slot: " & SlotIndex & ", Quantity: " & Quantity & ", Box: " & Box & ", GuildIndex: " & GuildIndex)
End Sub

Sub UserDepositObjInGB(ByVal UserIndex As Integer, ByVal SlotIndex As Integer, ByVal Quantity As Integer, ByVal GuildIndex As Integer)
On Error GoTo ErrHandler
  

    Dim Slot As Integer
    Dim obji As Integer
    Dim I As Integer
    
    If Quantity < 1 Then Exit Sub
    
    With UserList(UserIndex)
        obji = .Invent.Object(SlotIndex).ObjIndex
        
        '¿Ya tiene un objeto de este tipo?
        Slot = 1
        Do Until GuildList(GuildIndex).Bank(Slot).IdObject = obji And _
            GuildList(GuildIndex).Bank(Slot).Amount + Quantity <= MAX_INVENTORY_OBJS
            Slot = Slot + 1
            
            If Slot > GetLimitOfGuildBankSlots(GuildIndex) Then
                Exit Do
            End If
        Loop
        
        'Sino se fija por un slot vacio antes del slot devuelto
        If Slot > GetLimitOfGuildBankSlots(GuildIndex) Then
            Slot = 1
            Do Until GuildList(GuildIndex).Bank(Slot).IdObject = 0
                Slot = Slot + 1
                ' aca explota
                If Slot > GetLimitOfGuildBankSlots(GuildIndex) Then
                    Call WriteConsoleMsg(UserIndex, "No tienes mas espacio en el banco!!", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            Loop
        End If
        
        If Slot > GetLimitOfGuildBankSlots(GuildIndex) Then 'Slot valido
            Call WriteConsoleMsg(UserIndex, "No tienes más espacio en el el banco de clan.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Mete el obj en el slot
        If GuildList(GuildIndex).Bank(Slot).Amount + Quantity <= MAX_INVENTORY_OBJS Then
            Call WriteConsoleMsg(UserIndex, "El banco de clan no puede cargar tantos objetos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
            
        'Menor que MAX_INV_OBJS
        GuildList(GuildIndex).Bank(Slot).IdObject = obji
        GuildList(GuildIndex).Bank(Slot).Amount = GuildList(GuildIndex).Bank(Slot).Amount + Quantity
        
        Call QuitarUserInvItem(UserIndex, CByte(SlotIndex), Quantity)
        
        For I = 1 To GuildLastMemberOnline(GuildIndex)
            With GuildList(GuildIndex).OnlineMembers(I)
                If UserList(.MemberUserIndex).Id = .IdUser Then
                    Call UpdateGuildBankInv(False, .MemberUserIndex, Slot, GuildIndex)
                End If
            End With
        Next I
                
      
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub UserDepositObjInGB de modBanco.bas")
End Sub

Public Function GetOldestMemberIndex(ByVal GuildIndex As Integer, ByVal SkipLeader As Boolean, ByVal SkipRightHand As Boolean) As Long

On Error GoTo ErrHandler
    ' If no guild provided, it returns 0
    If GuildIndex = 0 Then Exit Function
    
    Dim I As Integer
    Dim OldestIndex As Integer
    
    Dim OldestDate As Date
    OldestDate = Now()
    
    For I = 1 To GuildList(GuildIndex).MemberCount
        If GuildList(GuildIndex).Members(I).JoinDate < OldestDate Then
            If Not (SkipLeader And GuildList(GuildIndex).Members(I).IdUser = GuildList(GuildIndex).IdLeader) And Not (SkipRightHand And GuildList(GuildIndex).Members(I).IdUser = GuildList(GuildIndex).IdLeader) Then
                OldestDate = GuildList(GuildIndex).Members(I).JoinDate
                OldestIndex = I
            End If
        End If
                      
    Next I
    
    GetOldestMemberIndex = OldestIndex
Exit Function
  
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetOldestMemberIndex de modGuild_Functions.bas")
End Function

Public Function GetMemberNameByUserId(ByVal GuildIndex As Integer, ByVal UserId As Long) As String
On Error GoTo ErrHandler

    Dim I As Integer
    
    For I = 1 To GuildList(GuildIndex).MemberCount
        If GuildList(GuildIndex).Members(I).IdUser = UserId Then
            GetMemberNameByUserId = GuildList(GuildIndex).Members(I).NameUser
            Exit Function
        End If
    Next I
    
Exit Function
  
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetMemberNameByUserId de modGuild_Functions.bas")
End Function

Public Function GetMemberIndexOf(ByVal UserIndex As Integer) As Integer
On Error GoTo ErrHandler

    Dim I As Integer, IndexMember As Integer, GuildIndex As Integer
    
    GuildIndex = UserList(UserIndex).Guild.GuildIndex
        
    For I = 1 To GuildList(GuildIndex).MemberCount
        If GuildList(GuildIndex).Members(I).IdUser = UserList(UserIndex).ID Then
            GetMemberIndexOf = I
            Exit Function
        End If
    Next I
    
    GetMemberIndexOf = 0
    
    Exit Function
    
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub GetMemberIndexOf de modBanco.bas")
End Function


Public Sub GuildMemberAdd(ByVal GuildIndex As Long, ByVal MemberRequestIndex As Long, ByVal TargetRequestIndex As Integer, ByVal RoleId As Integer)
On Error GoTo ErrHandler

    Dim I  As Integer
    
    Call GuildMemberAddDB(GuildList(GuildIndex).IdGuild, UserList(MemberRequestIndex).ID, UserList(TargetRequestIndex).ID, RoleId)

    UserList(TargetRequestIndex).Guild.IdGuild = GuildList(GuildIndex).IdGuild
    UserList(TargetRequestIndex).Guild.GuildIndex = GuildIndex
    UserList(TargetRequestIndex).Guild.GuildMemberIndex = GuildList(GuildIndex).MemberCount
    
    For I = 1 To GuildList(GuildIndex).MemberCount
        Debug.Print GuildList(GuildIndex).Members(I).NameUser
    Next I
    
    Call WriteGuildInfo(TargetRequestIndex)
    Call WriteGuildRolesList(TargetRequestIndex)
    Call WriteGuildMembersList(TargetRequestIndex)
    Call WriteGuildBankList(TargetRequestIndex)
    Call WriteGuildUpgradesList(TargetRequestIndex)
    Call WriteGuildUpgradesAcquired(TargetRequestIndex)
    Call WriteGuildQuestsCompletedList(TargetRequestIndex)
    Call WriteGuildCurrentQuestInfo(TargetRequestIndex)

    Call AddOnlineMember(TargetRequestIndex)
    
    Call NotifyMemberJoin(TargetRequestIndex)
    
    Call RefreshCharStatus(TargetRequestIndex, False)
    
    With GuildList(GuildIndex)
        If GuildLastMemberOnline(GuildIndex) > 0 Then
            For I = 1 To GuildLastMemberOnline(GuildIndex)
                Call WriteGuildMembersList(.OnlineMembers(I).MemberUserIndex)
            Next I
        End If
    End With
    
    Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub GuildMemberAdd de modGuild_Functions.bas")
End Sub

Public Sub GuildRemoveMember(ByVal GuildIndex As Long, ByVal UserId As Long, Optional ByRef KickedUsername As String, Optional ByVal NotifyUser As Boolean)
On Error GoTo ErrHandler
    
    Dim I As Integer, Size As Integer, KickedIndex As Integer, ExMemberIndex As Integer
    Dim Flag As Boolean
    Dim KickedName As String
    Dim CurrentLeaderId As Integer
    Dim NewLeaderId As Integer
    Dim UserName As String
    Dim RoleName As String
    Dim Message As String
    Dim MemberIndex As Integer
    
    Call GuildMemberDeleteDB(GuildList(GuildIndex).IdGuild, UserId)
    
    Flag = False
    Size = GuildList(GuildIndex).MemberCount - 1
    MemberIndex = 0
    
    If GuildList(GuildIndex).IdRightHand = UserId Then
        GuildList(GuildIndex).IdRightHand = 0
        GuildList(GuildIndex).IsDirty = True
        Call CleanCurrentQuestInfo(GuildIndex)
    End If
    
    If GuildList(GuildIndex).IdLeader = UserId Then
        
        Call CleanCurrentQuestInfo(GuildIndex)
        CurrentLeaderId = GuildList(GuildIndex).IdLeader

        ' If the new Leader Id is 0, that means the guild didn't had a right hand, so
        ' we need to find a new leader from the list of players. It will be the oldest
        ' member, without counting the leader or righthand
        If GuildList(GuildIndex).IdRightHand = 0 Then
            NewLeaderId = GuildList(GuildIndex).Members(GetOldestMemberIndex(GuildIndex, True, True)).IdUser
        Else
            NewLeaderId = GuildList(GuildIndex).IdRightHand
        End If
        
         ' Clean the right hand role
        GuildList(GuildIndex).IdLeader = NewLeaderId
        GuildList(GuildIndex).IdRightHand = 0
        
        ' Is the new leader online?
        Dim NewLeaderIndex As Integer
        NewLeaderIndex = GetUserIndexFromUserId(GuildList(GuildIndex).IdLeader)
        
        If NewLeaderIndex > 0 Then
            UserList(NewLeaderIndex).Guild.IdGuild = GuildList(GuildIndex).IdGuild
            UserList(NewLeaderIndex).Guild.GuildIndex = GuildIndex
            UserList(NewLeaderIndex).Guild.RoleId = ID_ROLE_LEADER
            UserList(NewLeaderIndex).Guild.RoleIndex = 1
            
            MemberIndex = UserList(NewLeaderIndex).Guild.GuildMemberIndex
            
        End If
        
        If MemberIndex = 0 Then
             For I = 1 To GuildList(GuildIndex).MemberCount
                If GuildList(GuildIndex).Members(I).IdUser = GuildList(GuildIndex).IdLeader Then
                    MemberIndex = I
                    Exit For
                End If
            Next I
        End If
        
        UserName = GuildList(GuildIndex).Members(MemberIndex).NameUser
        GuildList(GuildIndex).Members(MemberIndex).IdRole = ID_ROLE_LEADER
        GuildList(GuildIndex).Members(MemberIndex).RoleIndex = GetRoleIndexFromRoleId(GuildIndex, ID_ROLE_LEADER)
        
        RoleName = GuildList(GuildIndex).Roles(ID_ROLE_LEADER).RoleName
        
        ' Assign the leader role to the current righthand
        Call modGuild_DB.AssignRoleFromDB(CurrentLeaderId, GuildList(GuildIndex).IdGuild, NewLeaderId, ID_ROLE_LEADER)
        Call modGuild_Functions.ChangeUserRole(GuildIndex, NewLeaderId, ID_ROLE_LEADER, CurrentLeaderId)
        
        ' Update guild leadership again
        Call UpdateGuildLeadership(GuildList(GuildIndex).IdGuild, NewLeaderId, GuildList(GuildIndex).IdRightHand)
        
        If GuildLastMemberOnline(GuildIndex) > 0 Then
            For I = 1 To GuildLastMemberOnline(GuildIndex)
                Message = UserName & " ha sido cambiado al rol " & RoleName & "."
                If GuildList(GuildIndex).OnlineMembers(I).IdUser = NewLeaderId Then
                    'affected user is online, update his info
                    Message = "Has sido cambiado al rol " & RoleName & "."
                End If
                Call WriteGuildMemberStatusChange(GuildList(GuildIndex).OnlineMembers(I).MemberUserIndex, NewLeaderId, eChangeMember.RoleChange, ID_ROLE_LEADER, 0)
                Call WriteConsoleMsg(GuildList(GuildIndex).OnlineMembers(I).MemberUserIndex, Message, FontTypeNames.FONTTYPE_GUILD, eMessageType.Guild)
            Next I
        End If
    End If
        
    For I = 1 To GuildList(GuildIndex).MemberCount
        If GuildList(GuildIndex).Members(I).IdUser = UserId Then
            Flag = True
            KickedName = GuildList(GuildIndex).Members(I).NameUser
            KickedUsername = KickedName
        End If
        
        If Flag = True And I >= 1 And I <= Size Then
            GuildList(GuildIndex).Members(I) = GuildList(GuildIndex).Members(I + 1)
        End If
    Next I
    
    GuildList(GuildIndex).IsDirty = True
    ReDim Preserve GuildList(GuildIndex).Members(1 To Size) As GuildMemberType
    GuildList(GuildIndex).MemberCount = Size
    
    ' if he's online when he's kicked out
    KickedIndex = NameIndex(KickedName)
    If KickedIndex > 0 Then
        Call RemoveOnlineMember(KickedIndex)
        UserList(KickedIndex).Guild.IdGuild = 0
        UserList(KickedIndex).Guild.GuildMemberIndex = 0
        
        If NotifyUser Then
            Call WriteGuildMemberKicked(KickedIndex)
            Call RefreshCharStatus(KickedIndex, False)
        End If
    End If
    
    If GuildLastMemberOnline(GuildIndex) > 0 Then
        For I = 1 To GuildLastMemberOnline(GuildIndex)
            're-send guild info at all members
            UserList(GuildList(GuildIndex).OnlineMembers(I).MemberUserIndex).Guild.GuildMemberIndex = GetMemberIndexOf(GuildList(GuildIndex).OnlineMembers(I).MemberUserIndex)
            Call WriteGuildInfo(GuildList(GuildIndex).OnlineMembers(I).MemberUserIndex)
            Call WriteGuildMembersList(GuildList(GuildIndex).OnlineMembers(I).MemberUserIndex)
        Next I
    End If
    Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub GuildRemoveMember de modGuild_Functions.bas")
End Sub

Public Function GetGuildRoleId(ByVal GuildIndex As Integer, ByVal UserIndex As Integer)
On Error GoTo ErrHandler
    
    If GuildIndex > 0 And UserList(UserIndex).Guild.GuildMemberIndex > 0 Then
        GetGuildRoleId = GuildList(GuildIndex).Members(UserList(UserIndex).Guild.GuildMemberIndex).IdRole
    End If
    
    Exit Function
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub GetGuildRoleId de modGuild_Functions.bas")
End Function

Public Function ChangeUserRole(ByVal GuildIndex As Integer, ByVal UserId As Long, ByVal UserRole As Integer, ByVal AssignedBy As Long)

On Error GoTo ErrHandler

      Dim I As Integer
      
      With GuildList(GuildIndex)
      
        For I = 1 To .MemberCount
            If .Members(I).IdUser = UserId Then
                .Members(I).IdRole = UserRole
                .Members(I).RoleAssignedBy = AssignedBy
                .Members(I).IsDirty = True
                Exit Function
            End If
        Next I
      
      End With
Exit Function
  
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function ChangeUserRole de modGuild_Functions.bas")
End Function

Public Sub SendToDiosesYclan(ByVal GuildIndex As Integer, ByRef sndData As String, Optional ByVal IsUrgent As Boolean = False)
On Error GoTo ErrHandler
  
    
    ' Send to members
    Call SendToGuildMembers(GuildIndex, sndData, IsUrgent)
    ' Send to admins
    Call SendToOnlineAdmins(GuildIndex, sndData, IsUrgent)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SendToDiosesYclan de modGuilds.bas")
End Sub

Public Sub SendToGuildMembers(ByVal GuildIndex As Integer, ByRef sndData As String, Optional ByVal IsUrgent As Boolean = False)
On Error GoTo ErrHandler

    Dim I As Integer
    Dim UserIndex As Integer
    
    If GuildLastMemberOnline(GuildIndex) > 0 Then
        For I = 1 To GuildLastMemberOnline(GuildIndex)
            UserIndex = GuildList(GuildIndex).OnlineMembers(I).MemberUserIndex
            
            If (UserList(UserIndex).ConnIDValida) Then
                TCP.Send UserList(UserIndex).Connection, IsUrgent
            End If
        Next I
    End If
    Exit Sub
    
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function SendToGuildMembers de modGuild_Functions.bas")
End Sub

Private Sub SendToOnlineAdmins(ByVal GuildIndex As Integer, ByRef sndData As String, Optional ByVal IsUrgent As Boolean = False)
On Error GoTo ErrHandler

    Dim I As Integer
    Dim UserIndex As Integer
    
    With GuildList(GuildIndex)
        For I = 1 To .ListeningAdminsCount
            UserIndex = .ListeningAdmins(I)
            If (UserList(UserIndex).ConnIDValida) Then
                TCP.Send UserList(UserIndex).Connection, IsUrgent
            End If
        Next I
    End With
    
    Exit Sub
    
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function SendToOnlineAdmins de modGuild_Functions.bas")
End Sub
Public Sub NotifyMemberJoin(ByVal UserIndex As Long)
    Dim I As Integer, GuildIndex As Integer
    Dim Message As String
    
    If UserList(UserIndex).Guild.IdGuild <= 0 Then Exit Sub
    
    GuildIndex = UserList(UserIndex).Guild.GuildIndex
    
    If GuildLastMemberOnline(GuildIndex) > 0 Then
        For I = 1 To GuildLastMemberOnline(GuildIndex)
            With GuildList(GuildIndex).OnlineMembers(I)
                    Message = UserList(UserIndex).Name & " se unio al clan."
                    Call WriteGuildMemberStatusChange(.MemberUserIndex, UserList(UserIndex).ID, eChangeMember.OnlineChange, 1, 0)
                
                If .MemberUserIndex <> UserIndex Then
                    Call WriteConsoleMsg(.MemberUserIndex, Message, FontTypeNames.FONTTYPE_GUILD, eMessageType.Guild)
                Else
                    Call WriteConsoleMsg(.MemberUserIndex, "Te has unido al clan " & GuildList(GuildIndex).Name, FontTypeNames.FONTTYPE_GUILD, eMessageType.Guild)
                End If
            End With
        Next I
    End If
    Exit Sub

ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function NotifyMembers de modGuild_Functions.bas")
End Sub

Public Function GetLimitOfGuildMember(ByVal GuildIndex As Integer) As Integer
On Error GoTo ErrHandler
    
    GetLimitOfGuildMember = GuildConfiguration.MemberQty + GuildList(GuildIndex).UpgradeEffect.AddMemberLimit

    Exit Function
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetLimitOfGuildMember de modGuild_Functions.bas")
End Function
Public Sub BuyGuildUpgrade(ByVal UserIndex As Integer, ByVal UpgradeId As Integer)
On Error GoTo ErrHandler
    Dim RoleId As Integer
    
    RoleId = UserList(UserIndex).Guild.RoleId
    
    If RoleId <> ID_ROLE_LEADER And RoleId <> ID_ROLE_RIGHTHAND Then
         Call WriteConsoleMsg(UserIndex, NOPERMISSIONOFGUILD, FontTypeNames.FONTTYPE_INFO, info)
         Exit Sub
    End If
    
    With GuildList(UserList(UserIndex).Guild.GuildIndex)
    
        If .IdRightHand = 0 Then
            Call WriteConsoleMsg(UserIndex, "El Clan debe tener una Mano derecha asignada para poder comprar una mejora.", FontTypeNames.FONTTYPE_GUILD, eMessageType.Guild)
            Exit Sub
        End If
        
        If .ContributionAvailable < GuildConfiguration.GuildUpgradesList(UpgradeId).ContributionCost Then
            Call WriteConsoleMsg(UserIndex, "No posees suficientes puntos de contribución para esta Mejora.", FontTypeNames.FONTTYPE_GUILD, eMessageType.Guild)
            Exit Sub
        End If
        
        If UserList(UserIndex).Stats.GLD < GuildConfiguration.GuildUpgradesList(UpgradeId).GoldCost Then
            Call WriteConsoleMsg(UserIndex, "No posees suficientes Oro para esta Mejora.", FontTypeNames.FONTTYPE_GUILD, eMessageType.Guild)
            Exit Sub
        End If
        
        If Not GuildUpgradeRequirement(UserIndex, UpgradeId) Then
            Call WriteConsoleMsg(UserIndex, "No Cumples los requerimientos de esta Mejora.", FontTypeNames.FONTTYPE_GUILD, eMessageType.Guild)
            Exit Sub
        End If
        
        .ContributionAvailable = .ContributionAvailable - GuildConfiguration.GuildUpgradesList(UpgradeId).ContributionCost
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - GuildConfiguration.GuildUpgradesList(UpgradeId).GoldCost
    End With
    
    Call AcquireGuildUpgrade(UserIndex, UpgradeId)
    
    Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub BuyGuildUpgrade de modGuild_Functions.bas")
End Sub

Private Sub AcquireGuildUpgrade(ByVal UserIndex As Integer, ByVal UpgradeId As Integer)
On Error GoTo ErrHandler:
    Dim Rs As Recordset
    Dim CountUpgrade As Integer, GuildIndex As Integer, MaxSlot As Integer
    
    ' Insert the guildUpgrade in the database.
    Call AddGuildUpgradeDB(UserList(UserIndex).Guild.IdGuild, UserList(UserIndex).ID, UpgradeId)
    
    Set Rs = LoadGuildUpgradeDB(UserList(UserIndex).Guild.IdGuild, UpgradeId)
    
    If SizeRS(Rs) = 0 Then Exit Sub
    
    GuildIndex = UserList(UserIndex).Guild.GuildIndex
    
    CountUpgrade = GuildLastUpgrade(GuildIndex) + 1
    
    ReDim Preserve GuildList(GuildIndex).Upgrades(1 To CountUpgrade) As GuildUpgradeType
    
    MaxSlot = GetLimitOfGuildBankSlots(GuildIndex)
    
    With GuildList(GuildIndex)
        .Upgrades(CountUpgrade).IdUpgrade = CInt(Rs.Fields("ID_UPGRADE"))
        .Upgrades(CountUpgrade).IsEnabled = CBool(Rs.Fields("ENABLED"))
        .Upgrades(CountUpgrade).UpgradeBy = CInt(Rs.Fields("UPGRADED_BY"))
        .Upgrades(CountUpgrade).UpgradeDate = CDate(Rs.Fields("UPGRADE_DATE"))
        .Upgrades(CountUpgrade).UpgradeLevel = CInt(Rs.Fields("UPGRADE_LEVEL"))
        
        Call AddGuildUpgradeEffect(GuildIndex, .Upgrades(CountUpgrade).IdUpgrade)
    End With
    
    ' if qty slots change, resize
    If (MaxSlot <> GetLimitOfGuildBankSlots(GuildIndex)) Then
        ReDim Preserve GuildList(GuildIndex).Bank(1 To GetLimitOfGuildBankSlots(GuildIndex)) As GuildBankType
    End If
    
    Call NotifyUpgradeBought(UserIndex, UpgradeId)
    
    Exit Sub
    
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub AcquireGuildUpgrade de modGuild_function.bas")
End Sub

Private Sub NotifyUpgradeBought(ByVal UserIndex As Long, ByVal UpgradeId As Integer)
On Error GoTo ErrHandler
    Dim I As Integer, J As Integer, GuildIndex As Integer
    Dim TypeOfUpgrade As eChangeGuildInfo
    Dim ValueToSend As Integer, QtyNotification As Byte
    Dim ValueToSendLong As Long
    
    If UserList(UserIndex).Guild.IdGuild <= 0 Then Exit Sub
    
    GuildIndex = UserList(UserIndex).Guild.GuildIndex
    
    QtyNotification = UpgradeNotificationQuantity(GuildIndex, UpgradeId)
        
    If GuildLastMemberOnline(GuildIndex) > 0 Then
        For I = 1 To GuildLastMemberOnline(GuildIndex)
            With GuildList(GuildIndex).OnlineMembers(I)
                Call WriteGuildUpgradesAcquired(.MemberUserIndex, UpgradeId)
                For J = 1 To QtyNotification
                    Call GetValueEffectNumber(J, QtyNotification, GuildIndex, UpgradeId, TypeOfUpgrade, ValueToSend, ValueToSendLong)
                    Call WriteGuildInfoChange(.MemberUserIndex, TypeOfUpgrade, ValueToSend, ValueToSendLong)
                    ValueToSend = 0
                    ValueToSendLong = 0
                Next J
                If QtyNotification = 0 Then
                    Call WriteGuildInfoChange(.MemberUserIndex, eChangeGuildInfo.EnableBank, GuildList(GuildIndex).UpgradeEffect.IsGuildBank, 0)
                End If
                
                ValueToSendLong = GuildList(GuildIndex).ContributionAvailable
                
                Call WriteGuildInfoChange(.MemberUserIndex, eChangeGuildInfo.ContributionAvailableChange, ValueToSend, ValueToSendLong)
                
                ValueToSendLong = GuildList(GuildIndex).BankGold
                
                Call WriteGuildInfoChange(.MemberUserIndex, eChangeGuildInfo.BankGoldChange, ValueToSend, ValueToSendLong)
                
                Call WriteConsoleMsg(.MemberUserIndex, "Se ha adquirido la mejora de clan " & GuildConfiguration.GuildUpgradesList(UpgradeId).Name, FontTypeNames.FONTTYPE_GUILD, eMessageType.Guild)
            End With
        Next I
    End If

    Exit Sub

ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function NotifyUpgradeBought de modGuild_Functions.bas")
End Sub

Private Function UpgradeNotificationQuantity(ByVal GuildIndex As Integer, ByVal UpgradeId As Integer) As Byte
On Error GoTo ErrHandler
    Dim Qty As Byte

    With GuildConfiguration.GuildUpgradesList(UpgradeId)

        If .UpgradeEffect.AddMemberLimit > 0 Then
            Qty = Qty + 1
        End If
        If .UpgradeEffect.AddRolesGuild > 0 Then
            Qty = Qty + 1
        End If
        If .UpgradeEffect.AddBankSlot > 0 Then
            Qty = Qty + 1
        End If
        If .UpgradeEffect.AddBankBox > 0 Then
            Qty = Qty + 1
        End If
        If .UpgradeEffect.AddMaxContribution > 0 Then
            Qty = Qty + 1
        End If

    End With
    
    UpgradeNotificationQuantity = Qty

Exit Function
  
ErrHandler:
Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function UpgradeNotificationQuantity de modGuild_Functions.bas")
End Function


Private Sub GetValueEffectNumber(ByVal Index As Integer, ByVal TotalEffect As Integer, ByVal GuildIndex As Integer, ByVal UpgradeId As Integer, ByRef TypeOfUpgrade As eChangeGuildInfo, ByRef ValueToSend As Integer, ByRef ValueToSendLong As Long)
On Error GoTo ErrHandler
    Dim I As Integer, NumberEffect As Integer
    
    With GuildConfiguration.GuildUpgradesList(UpgradeId)
        For I = 1 To TotalEffect
            If .UpgradeEffect.AddMemberLimit > 0 Then
                NumberEffect = NumberEffect + 1
                If NumberEffect = Index Then
                    TypeOfUpgrade = MaxMembersQtyChange
                    ValueToSend = GetLimitOfGuildMember(GuildIndex)
                    Exit Sub
                End If
            End If
            
            If .UpgradeEffect.AddRolesGuild > 0 Then
                NumberEffect = NumberEffect + 1
                If NumberEffect = Index Then
                    TypeOfUpgrade = MaxRolesQtyChange
                    ValueToSend = GetLimitOfGuildRoles(GuildIndex)
                    Exit Sub
                End If
            End If
            If .UpgradeEffect.AddBankSlot > 0 Then
                NumberEffect = NumberEffect + 1
                If NumberEffect = Index Then
                    TypeOfUpgrade = MaxSlotsBankQtyChange
                    ValueToSend = GetLimitOfGuildBankSlots(GuildIndex)
                    Exit Sub
                End If
            End If
            If .UpgradeEffect.AddBankBox > 0 Then
                NumberEffect = NumberEffect + 1
                If NumberEffect = Index Then
                    TypeOfUpgrade = MaxBoxesBankQtyChange
                    ValueToSend = GetLimitOfGuildBankBoxes(GuildIndex)
                    Exit Sub
                End If
            End If
            
            If .UpgradeEffect.AddMaxContribution > 0 Then
                NumberEffect = NumberEffect + 1
                If NumberEffect = Index Then
                    TypeOfUpgrade = MaxContributionChange
                    ValueToSendLong = GetLimitOfGuildContribution(GuildIndex)
                    Exit Sub
                End If
            End If
    
        Next I
    
    End With
   
    Exit Sub
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub GetValueEffectNumber de modGuild_Functions.bas")
End Sub

Public Function GuildLastUpgrade(ByVal GuildIndex As Integer) As Integer

On Error GoTo ErrHandler

    If ((Not GuildList(GuildIndex).Upgrades) = -1) Then
        GuildLastUpgrade = 0
    Else
        GuildLastUpgrade = UBound(GuildList(GuildIndex).Upgrades)
    End If

  Exit Function

ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GuildLastUpgrade de modGuild_Functions.bas")
End Function


Public Sub AddGuildUpgradeEffect(ByVal GuildIndex As Integer, ByVal UpgradeId As Integer)
On Error GoTo ErrHandler
    With GuildList(GuildIndex)
    
        ' add upgrade effects
        .UpgradeEffect.AddBankBox = .UpgradeEffect.AddBankBox + GuildConfiguration.GuildUpgradesList(UpgradeId).UpgradeEffect.AddBankBox
        .UpgradeEffect.AddBankSlot = .UpgradeEffect.AddBankSlot + GuildConfiguration.GuildUpgradesList(UpgradeId).UpgradeEffect.AddBankSlot
        .UpgradeEffect.AddMaxContribution = .UpgradeEffect.AddMaxContribution + GuildConfiguration.GuildUpgradesList(UpgradeId).UpgradeEffect.AddMaxContribution
        .UpgradeEffect.AddMemberLimit = .UpgradeEffect.AddMemberLimit + GuildConfiguration.GuildUpgradesList(UpgradeId).UpgradeEffect.AddMemberLimit
        .UpgradeEffect.AddRolesGuild = .UpgradeEffect.AddRolesGuild + GuildConfiguration.GuildUpgradesList(UpgradeId).UpgradeEffect.AddRolesGuild
        .UpgradeEffect.IsChatOverHead = .UpgradeEffect.IsChatOverHead + GuildConfiguration.GuildUpgradesList(UpgradeId).UpgradeEffect.IsChatOverHead
        .UpgradeEffect.IsFriendlyFireProtection = .UpgradeEffect.IsFriendlyFireProtection + GuildConfiguration.GuildUpgradesList(UpgradeId).UpgradeEffect.IsFriendlyFireProtection
        .UpgradeEffect.IsGuildBank = .UpgradeEffect.IsGuildBank + GuildConfiguration.GuildUpgradesList(UpgradeId).UpgradeEffect.IsGuildBank
        .UpgradeEffect.IsSeeInvisibleGuildMember = .UpgradeEffect.IsSeeInvisibleGuildMember + GuildConfiguration.GuildUpgradesList(UpgradeId).UpgradeEffect.IsSeeInvisibleGuildMember
    End With
Exit Sub
  
ErrHandler:
Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub AddGuildUpgradeEffect de modGuild_Functions.bas")
End Sub

Public Function GetLimitOfGuildRoles(ByVal GuildIndex As Integer) As Byte
On Error GoTo ErrHandler
    GetLimitOfGuildRoles = GuildConfiguration.RolsQty + GuildList(GuildIndex).UpgradeEffect.AddRolesGuild

    Exit Function
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetLimitOfGuildRoles de modGuild_Functions.bas")
End Function

Public Function GetLimitOfGuildBankSlots(ByVal GuildIndex As Integer) As Byte
On Error GoTo ErrHandler
    GetLimitOfGuildBankSlots = GuildConfiguration.BankSlotQty + GuildList(GuildIndex).UpgradeEffect.AddBankSlot

    Exit Function
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetLimitOfGuildBankSlots de modGuild_Functions.bas")
End Function


Public Function GetLimitOfGuildBankBoxes(ByVal GuildIndex As Integer) As Byte
On Error GoTo ErrHandler
    GetLimitOfGuildBankBoxes = GuildConfiguration.BankBoxesQty + GuildList(GuildIndex).UpgradeEffect.AddBankSlot

    Exit Function
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetLimitOfGuildBankBoxes de modGuild_Functions.bas")
End Function

Public Function GetLimitOfGuildContribution(ByVal GuildIndex As Integer) As Long
On Error GoTo ErrHandler
    GetLimitOfGuildContribution = GuildConfiguration.MaxContribution + GuildList(GuildIndex).UpgradeEffect.AddMaxContribution

    Exit Function
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetLimitOfGuildContribution de modGuild_Functions.bas")
End Function


Public Function GetTransparencyAllie(ByVal InstigatorUserIndex As Integer, ByVal ObserverUserIndex As Integer) As Boolean
On Error GoTo ErrHandler
    If (UserList(InstigatorUserIndex).Guild.GuildIndex = 0 Or UserList(ObserverUserIndex).Guild.GuildIndex = 0) Then
        GetTransparencyAllie = False
         Exit Function
    End If

    GetTransparencyAllie = UserList(InstigatorUserIndex).Guild.GuildIndex = UserList(ObserverUserIndex).Guild.GuildIndex And GuildList(UserList(InstigatorUserIndex).Guild.GuildIndex).UpgradeEffect.IsSeeInvisibleGuildMember
    
    Exit Function
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetTransparencyAllie de modGuild_Functions.bas")
End Function

Public Function FriendlyFireProtectionEnabled(ByVal UserIndex As Integer, ByVal TargetUserIndex As Integer) As Boolean
On Error GoTo ErrHandler
    If UserList(UserIndex).Guild.IdGuild > 0 Then
        If GuildList(UserList(UserIndex).Guild.GuildIndex).UpgradeEffect.IsFriendlyFireProtection And UserList(UserIndex).Guild.IdGuild = UserList(TargetUserIndex).Guild.IdGuild Then
            FriendlyFireProtectionEnabled = True
            Exit Function
        End If
    End If
    
    FriendlyFireProtectionEnabled = False
    Exit Function
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function FriendlyFireProtectionEnabled de modGuild_Functions.bas")
End Function

Public Sub GuildBanChar(ByVal GuildIndex As Long, ByVal UserId As Long)
On Error GoTo ErrHandler
    Dim NewMemberRole As Integer
    Dim NewMemberRoleName As String
    
    Dim ChangeMemberRole As Boolean
    Dim FindNewLeader As Boolean
    Dim CurrentRightHand As Long
    
    Dim NewLeaderId As Long
    Dim NewLeaderIndex As Integer
    Dim UserIndex As Integer
    
    Dim NewLeaderName As String
    Dim CurrentUserName As String
    Dim I As Integer
    
    Dim CurrentLeaderId As Long
    
    Dim NewMemberRoleIndex As Integer
    Dim LeaderRoleIndex As Integer
    
    ' If the guild has only one member, we just exit as there's nothing to do.
    ' The leader will be banned and the clan won't be usable anymore.
    ' TODO: When banning guilds is implemented, the guild should be marked as banned if there's no one
    ' else to support it's functioning.
    If GuildList(GuildIndex).MemberCount = 1 Then Exit Sub
    
    CurrentLeaderId = GuildList(GuildIndex).IdLeader
    
    ' The user index of the player we're banning if
    UserIndex = GetUserIndexFromUserId(UserId)
    
    ' If the user is the leader or the righthand, it will be moved to a normal role (Reclut)
    If GuildList(GuildIndex).IdLeader = UserId Or GuildList(GuildIndex).IdRightHand = UserId Then
        ChangeMemberRole = True
        NewMemberRole = GuildList(GuildIndex).Roles(3).IdRole
        NewMemberRoleName = GuildList(GuildIndex).Roles(3).RoleName
        
        NewMemberRoleIndex = 3
        LeaderRoleIndex = 1
        
        If GuildList(GuildIndex).IdLeader = UserId Then
            
            If GuildList(GuildIndex).IdRightHand > 0 Then
                NewLeaderId = GuildList(GuildIndex).IdRightHand
                NewLeaderName = GetMemberNameByUserId(GuildIndex, NewLeaderId)

                GuildList(GuildIndex).IdRightHand = 0
            Else
                Dim MemberIndex As Integer
                MemberIndex = GetOldestMemberIndex(GuildIndex, True, True)
                ' Find the oldest member as it will be assigned as a leader
                NewLeaderId = GuildList(GuildIndex).Members(MemberIndex).IdUser
                NewLeaderName = GuildList(GuildIndex).Members(MemberIndex).NameUser
            End If
            GuildList(GuildIndex).IdLeader = NewLeaderId
        Else
            GuildList(GuildIndex).IdRightHand = 0
        End If
    Else
        ChangeMemberRole = False
    End If
    
    ' Move users to the new roles
    If ChangeMemberRole Then
    
        ' Let's get the banned user name.. If the user is only we also update the user's metadata
        If UserIndex > 0 Then
            UserList(UserIndex).Guild.RoleId = NewMemberRole
            CurrentUserName = UserList(UserIndex).Name
        Else
            For I = 1 To GuildList(GuildIndex).MemberCount
                If GuildList(GuildIndex).Members(I).IdUser = UserId Then
                    CurrentUserName = GuildList(GuildIndex).Members(I).NameUser
                    Exit For
                End If
            Next I
        End If
    
        ' Change the role of the banned player
        Call modGuild_Functions.ChangeUserRole(GuildIndex, UserId, NewMemberRole, UserId)
        Call modGuild_DB.AssignRoleFromDB(CurrentLeaderId, GuildList(GuildIndex).IdGuild, UserId, NewMemberRole)
        
        ' Change the role of the new leader if needed
        If NewLeaderId > 0 Then
            NewLeaderIndex = GetUserIndexFromUserId(NewLeaderId)
                        
            ' Let's get the new leader name. If the user is only we also update the user's metadata
            If NewLeaderIndex > 0 Then
                UserList(NewLeaderIndex).Guild.RoleId = ID_ROLE_LEADER
                UserList(NewLeaderIndex).Guild.RoleIndex = GetRoleIndexFromRoleId(GuildIndex, ID_ROLE_LEADER)
                NewLeaderName = UserList(NewLeaderIndex).Name
            End If
            
            For I = 1 To GuildList(GuildIndex).MemberCount
                If GuildList(GuildIndex).Members(I).IdUser = GuildList(GuildIndex).IdLeader Then
                    NewLeaderName = GuildList(GuildIndex).Members(I).NameUser
                    GuildList(GuildIndex).Members(I).IdRole = ID_ROLE_LEADER
                    GuildList(GuildIndex).Members(I).RoleIndex = GetRoleIndexFromRoleId(GuildIndex, ID_ROLE_LEADER)
                    Exit For
                End If
            Next I
            
            Call modGuild_DB.AssignRoleFromDB(CurrentLeaderId, GuildList(GuildIndex).IdGuild, NewLeaderId, ID_ROLE_LEADER)
            Call modGuild_Functions.ChangeUserRole(GuildIndex, NewLeaderId, ID_ROLE_LEADER, UserId)
            
            GuildList(GuildIndex).IdLeader = NewLeaderId

        End If
        
        ' Update guild leadership again
        Call UpdateGuildLeadership(GuildList(GuildIndex).IdGuild, NewLeaderId, GuildList(GuildIndex).IdRightHand)
        
        ' Notify online members
        If GuildList(GuildIndex).OnlineMemberCount > 0 Then
            
            Dim MessageBannedUser As String
            Dim MessageDemotedUser As String
            Dim MessagePromotedUser As String
            
            MessageDemotedUser = CurrentUserName & " ha sido cambiado al rol " & GuildList(GuildIndex).Roles(NewMemberRoleIndex).RoleName & "."
            MessagePromotedUser = NewLeaderName & " ha sido cambiado al rol " & GuildList(GuildIndex).Roles(LeaderRoleIndex).RoleName & "."
            
            For I = 1 To GuildList(GuildIndex).OnlineMemberCount
                If ChangeMemberRole Then
                    ' Notify about the change of role of the current user.
                    Call WriteGuildMemberStatusChange(GuildList(GuildIndex).OnlineMembers(I).MemberUserIndex, UserId, eChangeMember.RoleChange, NewMemberRole, 0)
                    Call WriteConsoleMsg(GuildList(GuildIndex).OnlineMembers(I).MemberUserIndex, MessageDemotedUser, FontTypeNames.FONTTYPE_GUILD, eMessageType.Guild)
                
                    ' if the kicked user was the leader, that means there's a new leader being assigned so we need to notify that.
                    If NewLeaderId > 0 Then
                        Call WriteGuildMemberStatusChange(GuildList(GuildIndex).OnlineMembers(I).MemberUserIndex, NewLeaderId, eChangeMember.RoleChange, ID_ROLE_LEADER, 0)
                        Call WriteConsoleMsg(GuildList(GuildIndex).OnlineMembers(I).MemberUserIndex, MessagePromotedUser, FontTypeNames.FONTTYPE_GUILD, eMessageType.Guild)
                    End If
                End If
                
                Call WriteGuildInfo(GuildList(GuildIndex).OnlineMembers(I).MemberUserIndex)
            Next I
        End If
                
                
    End If
    
    GuildList(GuildIndex).IsDirty = True
Exit Sub
  
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub GuildBanChar de modGuild_Functions.bas")
End Sub

Public Sub GuildBanOnlineChar(ByVal UserIndex As Integer, ByVal UserId As Integer)
    On Error GoTo ErrHandler
    
    If UserList(UserIndex).Guild.GuildIndex > 0 Then
        Call GuildRemoveMember(UserList(UserIndex).Guild.GuildIndex, UserId)
    End If
    
    Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GuildBanOnlineChar de modGuild_Functions.bas")
End Sub

Public Sub GuildBanOfflineChar(ByVal UserId As Integer)
    On Error GoTo ErrHandler

    Dim I As Integer
    Dim J As Integer
    
    For I = 1 To MaxGuildQty
        With GuildList(I)
            For J = 1 To .MemberCount
                If .Members(J).IdUser = UserId Then
                    Call GuildRemoveMember(I, UserId)
                    Exit Sub
                End If
            Next J
        End With
    Next I
    
    Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GuildBanOfflineChar de modGuild_Functions.bas")
End Sub
Private Sub SlitRequired(ByRef SplitedRequest() As Integer, ByVal UnSplitString As String)
On Error GoTo ErrHandler
    Dim J As Integer, RequestQty As Integer
    Dim SplitRequestString() As String
       
    If UnSplitString <> vbNullString And UnSplitString <> "0" Then
        
        SplitRequestString = Split(UnSplitString, "-")
        RequestQty = UBound(SplitRequestString) + 1
        ReDim SplitedRequest(1 To RequestQty) As Integer
        For J = 1 To RequestQty
            SplitedRequest(J) = CInt(SplitRequestString(J - 1))
        Next J
    End If
    
    Exit Sub
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SlitRequired de modGuild_Functions.bas")
End Sub

Private Sub SlitQuestRequired(ByRef SplitedQuest() As GuildQuestReq, ByVal UnSplitString As String)
On Error GoTo ErrHandler
    Dim J As Integer, I As Integer, RequestQty As Integer
    Dim QuestQty As Integer, TempID As Integer
    Dim SplitRequestString() As String
    
    QuestQty = UBound(GuildQuestList)
    
    If UnSplitString <> vbNullString And UnSplitString <> "0" Then
        
        SplitRequestString = Split(UnSplitString, "-")
        RequestQty = UBound(SplitRequestString) + 1
        
        ReDim SplitedQuest(1 To RequestQty) As GuildQuestReq
        For J = 1 To RequestQty
            
            TempID = CInt(SplitRequestString(J - 1))
            SplitedQuest(J).ID = TempID
            
            For I = 1 To QuestQty
                If GuildQuestList(I).ID = TempID Then
                    SplitedQuest(J).Title = GuildQuestList(I).Title
                    
                End If
            Next I
        Next J
    End If
    
    Exit Sub
  
ErrHandler:
Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SplitedQuest de modGuild_Functions.bas")
End Sub

Public Function UpgradeRequireSize(ByRef VectorInt() As Integer) As Integer

On Error GoTo ErrHandler

    If Utility.IsArrayNull(VectorInt) Then
        UpgradeRequireSize = 0
    Else
        UpgradeRequireSize = UBound(VectorInt)
    End If

  Exit Function

ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function UpgradeRequireSize de modGuild_Functions.bas")
End Function


Public Function GuildUpgradeRequirement(UserIndex, Upgrade) As Boolean
On Error GoTo ErrHandler
    Dim ret As Boolean, RetTemp As Boolean
    Dim QuestsRequired As Integer, UpgradesRequired As Integer, I As Integer, J As Integer
    Dim UpgradesBuyed As Integer, QuestCompleted As Integer

    UpgradesRequired = UpgradeRequireSize(GuildConfiguration.GuildUpgradesList(Upgrade).UpgradeRequired)
    QuestsRequired = GetQuestQtyReq(Upgrade)
    
    With UserList(UserIndex)
    
        UpgradesBuyed = GuildLastUpgrade(.Guild.GuildIndex)
        
        ret = True
        
        If UpgradesRequired > 0 Then
            For I = 1 To UpgradesRequired
                RetTemp = False
                For J = 1 To UpgradesBuyed
                     If (GuildConfiguration.GuildUpgradesList(Upgrade).UpgradeRequired(I) = GuildList(.Guild.GuildIndex).Upgrades(J).IdUpgrade) Then
                        RetTemp = True
                     End If
                Next
                ret = ret And RetTemp
            Next I
        End If
        
        If QuestsRequired > 0 Then
           For I = 1 To QuestsRequired
               RetTemp = False
                For J = 1 To GuildList(.Guild.GuildIndex).QuestCompletedCount
                     If GuildConfiguration.GuildUpgradesList(Upgrade).QuestRequired(I).ID = GuildList(.Guild.GuildIndex).QuestCompleted(J).IdQuest Then
                        RetTemp = True
                     End If
                Next
                ret = ret And RetTemp
           Next I
        End If
        
    End With
    
    GuildUpgradeRequirement = ret
    
    Exit Function
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GuildUpgradeRequirement de modGuild_Functions.bas")
End Function

Public Function GetQuestQtyReq(ByVal Index As Integer) As Integer

On Error GoTo ErrHandler

    If Utility.IsArrayNull(GuildConfiguration.GuildUpgradesList(Index).QuestRequired) Then
        GetQuestQtyReq = 0
    Else
        GetQuestQtyReq = UBound(GuildConfiguration.GuildUpgradesList(Index).QuestRequired)
    End If

  Exit Function

ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetQuestQtyReq de modGuild_Functions.bas")
End Function

Public Sub GuildEarnContributionPoints(ByVal GuildIndex As Integer, ByVal ContributionPoints As Integer)
On Error GoTo ErrHandler

    GuildList(GuildIndex).ContributionEarned = GuildList(GuildIndex).ContributionEarned + ContributionPoints

    Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function EarnContributionPoints de modGuild_Functions.bas")
End Sub


Public Sub AdminListenToGuild(ByVal UserIndex As Integer, ByVal GuildName As String)
'***************************************************
'Autor: ZaMa
'Last Modification: 27/11/2009
'Adds an admin to the given guild, letting him to listen to members chat.
'***************************************************
On Error GoTo ErrHandler


    Dim GuildIndex As Integer

    ' Valid Guild?
    GuildIndex = GetGuildIndex(GuildName)
    
    If GuildIndex <= 0 Then
        Call WriteConsoleMsg(UserIndex, "El clan no existe.", FontTypeNames.FONTTYPE_GUILD)
        Exit Sub
    End If
    
    
    With UserList(UserIndex)
        ' Admin is already listening the chats of a guild.
        If .ListeningGuild <> 0 Then
            Call WriteConsoleMsg(UserIndex, "Dejas de escuchar el clan " & GuildList(.ListeningGuild).Name, FontTypeNames.FONTTYPE_GUILD)
            Call modGuild_Functions.RemoveListeningAdmin(.ListeningGuild, UserIndex)
            .ListeningGuild = 0
            Exit Sub
        End If
            
        Call modGuild_Functions.AddListeningAdmin(GuildIndex, UserIndex)
        .ListeningGuild = GuildIndex
        Call WriteConsoleMsg(UserIndex, "Comienzas a escuchar el clan " & GuildList(.ListeningGuild).Name, FontTypeNames.FONTTYPE_GUILD)
        
           
    End With
    
    Exit Sub

ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function AddAdminToGuild de modGuilds.bas")
End Sub


Public Sub AddListeningAdmin(ByVal GuildIndex As Integer, ByVal UserIndex As Integer)
On Error GoTo ErrHandler
    Dim UserExists As Boolean
    Dim I As Integer
    
    With GuildList(GuildIndex)
        For I = 1 To .ListeningAdminsCount
            If .ListeningAdmins(I) = UserIndex Then
                UserExists = True
                Exit For
            End If
        Next I
        
        If UserExists Then Exit Sub
        
        ' Add user to the listening admins list
        ReDim Preserve .ListeningAdmins(1 To .ListeningAdminsCount + 1)
        .ListeningAdmins(.ListeningAdminsCount + 1) = UserIndex
        .ListeningAdminsCount = .ListeningAdminsCount + 1
            
    End With
Exit Sub
  
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub AddListeningAdmin de modGuild_Functions.bas")
End Sub

Public Sub RemoveListeningAdmin(ByVal GuildIndex As Integer, ByVal UserIndex As Integer)
On Error GoTo ErrHandler
    If GuildList(GuildIndex).ListeningAdminsCount <= 0 Then Exit Sub
    
    Dim UserExists As Boolean
    Dim I As Integer
    Dim ListeningAdminIndex As Integer
    
    With GuildList(GuildIndex)
        For I = 1 To .ListeningAdminsCount
            If .ListeningAdmins(I) = UserIndex Then
                UserExists = True
                ListeningAdminIndex = I
                Exit For
            End If
        Next I
        
        ' User doesnt exist, so we don't do anything
        If Not UserExists Then Exit Sub
        
        ' If there's only one admin listening, then we need to clear the array
        ' as VB6 doesnt have any smart way of having 0 elements in an array
        If .ListeningAdminsCount = 1 Then
            Erase .ListeningAdmins
            .ListeningAdminsCount = 0
            Exit Sub
        End If
        
        Dim TmpListeningAdmins() As Integer
        ReDim TmpListeningAdmins(1 To .ListeningAdminsCount - 1)
        
        Dim LowerBound As Integer
        Dim UpperBound As Integer
        LowerBound = LBound(.ListeningAdmins)
        
        For I = 1 To ListeningAdminIndex - 1
            TmpListeningAdmins(I) = .ListeningAdmins(I)
        Next I
        
        For I = ListeningAdminIndex + 1 To .ListeningAdminsCount
            TmpListeningAdmins(I - 1) = .ListeningAdmins(I)
        Next I
        
        .ListeningAdminsCount = .ListeningAdminsCount - 1
        
        .ListeningAdmins = TmpListeningAdmins
                    
    End With
    
Exit Sub
  
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub RemoveListeningAdmin de modGuild_Functions.bas")
End Sub

Public Sub UserInvitationResponse(ByVal TargetUserIndex As Long, ByVal GuildIndex As Integer, ByVal InvitationIndex As Integer, ByVal IsAccepted As Boolean)
On Error GoTo ErrHandler
    If Not ValidateInvitation(GuildIndex, InvitationIndex, TargetUserIndex, UserList(TargetUserIndex).Id) Then
        Call WriteConsoleMsg(TargetUserIndex, "Invitacion caducada", FontTypeNames.FONTTYPE_INFO, info)
        Exit Sub
    End If
    With GuildList(GuildIndex)

        If Not IsAccepted Then
            Call WriteConsoleMsg(.Invitations(InvitationIndex).InvitedByUserIndex, UserList(TargetUserIndex).Name & " ha rechazado la invitacion al clan.", FontTypeNames.FONTTYPE_GUILD, eMessageType.Guild)
            Exit Sub
        End If
        
        If .MemberCount >= GetLimitOfGuildMember(GuildIndex) Then
            Call WriteConsoleMsg(TargetUserIndex, "El clan ha llegado al limite de miembros.", FontTypeNames.FONTTYPE_INFO, info)
            Exit Sub
        End If
            
        Call GuildMemberAdd(GuildIndex, .Invitations(InvitationIndex).InvitedByUserIndex, TargetUserIndex, .IdDefaultRole)
            
        Call ClearInvitation(GuildIndex, InvitationIndex)
    
    End With

    Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub UserInvitationResponse de modGuilds.bas")

End Sub

Public Sub InviteUser(ByVal InvitedByUserIndex as Long, ByRef TargetUserName as String)
On Error GoTo ErrHandler
    Dim GuildIndex as Long
    Dim TargetUserIndex as Long
    Dim MemberAlignment As eGuildAlignment
    Dim InvitationIndex as Integer

    If Not HasPermission(InvitedByUserIndex, EGuildPermission.MEMBER_ACCEPT) Then
        Call MessageManager.Prepare(InvitedByUserIndex, eMessageId.Guild_No_Permission)
        Call MessageManager.SendToMessageBox
        Exit Sub
    End If

    GuildIndex = UserList(InvitedByUserIndex).Guild.GuildIndex

    If GuildList(GuildIndex).MemberCount >= GetLimitOfGuildMember(GuildIndex) Then
        Call MessageManager.Prepare(InvitedByUserIndex, eMessageId.Guild_Invitation_Limit_Reached)
        Call MessageManager.SendToMessageBox
        Exit Sub
    End If

    TargetUserIndex = NameIndex(TargetUserName)

    If TargetUserIndex = 0 Then
        Call MessageManager.Prepare(InvitedByUserIndex, eMessageId.Guild_Invitation_User_Offline)
        Call MessageManager.SendToMessageBox
        Exit Sub
    End If

    If UserList(TargetUserIndex).Guild.IdGuild > 0 Then
        Call MessageManager.Prepare(InvitedByUserIndex, eMessageId.Guild_Invitation_User_Already_Has_Guild)
        Call MessageManager.AddParameterAsText(UserList(TargetUserIndex).Name)
        Call MessageManager.SendToMessageBox
        Exit Sub
    End If
    
            
    If UserList(TargetUserIndex).Faccion.ArmadaReal = 1 Then
        MemberAlignment = eGuildAlignment.Real
    ElseIf UserList(TargetUserIndex).Faccion.FuerzasCaos = 1 Then
        MemberAlignment = eGuildAlignment.Evil
    Else
        MemberAlignment = eGuildAlignment.Neutral
    End If

    If UserList(TargetUserIndex).Faccion.Alignment <> GuildList(GuildIndex).Alignment Then
        Call MessageManager.Prepare(InvitedByUserIndex, eMessageId.Guild_Invitation_User_In_Other_Faction)
        Call MessageManager.AddParameterAsText(UserList(TargetUserIndex).Name)
        Call MessageManager.SendToMessageBox
        Exit Sub
    End If

    If Not TryAddInvitation(GuildIndex, InvitedByUserIndex, TargetUserIndex, InvitationIndex) Then
        Call MessageManager.Prepare(InvitedByUserIndex, eMessageId.Guild_Invitation_User_Already_Invitated)
        Call MessageManager.AddParameterAsText(UserList(TargetUserIndex).Name)
        Call MessageManager.SendToMessageBox
        Exit Sub
    End If
    
    Call MessageManager.Prepare(InvitedByUserIndex, eMessageId.Guild_Invitation_Sent)
    Call MessageManager.AddParameterAsText(UserList(TargetUserIndex).Name)
    Call MessageManager.SendToMessageBox
    
    Call WriteGuildInvitation(GuildIndex, InvitedByUserIndex, TargetUserIndex, InvitationIndex, GuildConfiguration.InvitationLifeTimeInMinutes)

    Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function InviteUser de modGuilds.bas")
End Sub

Private Function TryAddInvitation(ByVal GuildIndex As Long, ByVal InvitedByUserIndex As Long, ByVal TargetUserIndex As Long, ByRef InvitationResultIndex) As Boolean
On Error GoTo ErrHandler
    Dim InvitationIndex as Integer
    Dim TargetUserId As Long
    Dim NextFree As Long
    Dim InvitationLifeTime As Long

    TargetUserId = UserList(TargetUserIndex).Id
    NextFree = -1
    
    TryAddInvitation = False

    With GuildList(GuildIndex)
        For InvitationIndex = 0 To INVITATION_MAX_COUNT - 1
            If .Invitations(InvitationIndex).InvitedByUserIndex > 0 Then
                If InvitationDateIsExpired(.Invitations(InvitationIndex).InvitationDate) Then
                    Call ClearInvitation(GuildIndex, InvitationIndex)
                    if NextFree = -1 Then
                        NextFree = InvitationIndex
                    end if
                End If

                If .Invitations(InvitationIndex).TargetUserId = TargetUserId then
                    exit Function
                end If
            Else
                if NextFree = -1 Then
                    NextFree = InvitationIndex
                end if
            End If
        Next InvitationIndex

        if NextFree = -1 Then
             Call LogError("Error Can't add more invitations en Function TryAddInvitation de modGuilds.bas")
             exit Function
        end if

        .Invitations(NextFree).InvitedByUserIndex = InvitedByUserIndex
        .Invitations(NextFree).InvitedByUserId = UserList(InvitedByUserIndex).Id
        .Invitations(NextFree).TargetUserIndex = TargetUserIndex
        .Invitations(NextFree).TargetUserId = TargetUserId
        .Invitations(NextFree).InvitationDate = Now()
        
        UserList(TargetUserIndex).InvitationGuildIndex = GuildIndex
        
        InvitationResultIndex = NextFree

        TryAddInvitation = true
    End With
Exit Function
  
ErrHandler:
Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function TryAddInvitation de modGuild_Functions.bas")
End Function

Private Sub ClearInvitation(ByVal GuildIndex As Integer, ByVal InvitationIndex As Integer)
On Error GoTo ErrHandler
    With GuildList(GuildIndex).Invitations(InvitationIndex)
        If .InvitedByUserIndex > 0 Then
            UserList(.TargetUserIndex).InvitationGuildIndex = 0
        End If
        .InvitedByUserIndex = 0
        .InvitedByUserId = 0
        .TargetUserIndex = 0
        .TargetUserId = 0
        .InvitationDate = vbDate
    End With
Exit Sub
  
ErrHandler:
Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ClearInvitation de modGuild_Functions.bas")
End Sub

Private Function InvitationDateIsExpired(ByRef InvitationDate As Date) As Boolean
On Error GoTo ErrHandler
    Dim InvitationLifeTime As Long
    InvitationLifeTime = DateDiff("n", InvitationDate, Now())
    InvitationDateIsExpired = InvitationLifeTime >= GuildConfiguration.InvitationLifeTimeInMinutes
Exit Function
  
ErrHandler:
Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function InvitationDateIsExpired de modGuild_Functions.bas")
End Function

Public Function ValidateInvitation(ByVal GuildIndex As Long, ByVal InvitationIndex As Integer, ByVal TargetUserIndex As Long, ByVal TargetUserId As Long) As Boolean
On Error GoTo ErrHandler
    ValidateInvitation = False

    With GuildList(GuildIndex)
        If .Invitations(InvitationIndex).InvitedByUserIndex > 0 Then
            If Not InvitationDateIsExpired(.Invitations(InvitationIndex).InvitationDate) Then
                If .Invitations(InvitationIndex).TargetUserIndex = TargetUserIndex And _
                    .Invitations(InvitationIndex).TargetUserId = TargetUserId And _
                    UserList(.Invitations(InvitationIndex).InvitedByUserIndex).Id = .Invitations(InvitationIndex).InvitedByUserId _
                    Then
                        ValidateInvitation = True
                End If
            End If
        End If
    End With
Exit Function

ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function ValidateInvitation de modGuild_Functions.bas")
End Function

Public Sub ClearInvitationByUserId(ByVal GuildIndex As Long, ByVal UserId As Long)
    Dim InvitationIndex as Integer
    
    With GuildList(GuildIndex)
        For InvitationIndex = 0 To INVITATION_MAX_COUNT - 1
            If .Invitations(InvitationIndex).InvitedByUserId = UserId Or _
                .Invitations(InvitationIndex).TargetUserId = UserId Then
                    Call ClearInvitation(GuildIndex, InvitationIndex)
            End If
        Next
    End With
End Sub

Public Sub RoleUpsert(ByVal GuildIndex As Long, ByVal RoleId As Long, ByRef RoleName As String, ByVal UserIndex As Integer, ByRef Permissions As String)
On Error GoTo ErrHandler
    Dim RoleIndex As Long
    Dim J As Integer
    Dim I As Integer
    
    With UserList(UserIndex)
        If RoleId = 0 And Not HasPermission(UserIndex, ROLE_CREATE_DELETE) Then
            Call WriteErrorMsg(UserIndex, NOPERMISSIONOFGUILD)
            Exit Sub
        End If
        
        If RoleId <> 0 And Not HasPermission(UserIndex, ROLE_MODIFY) Then
            Call WriteErrorMsg(UserIndex, NOPERMISSIONOFGUILD)
            Exit Sub
        End If
        
        GuildIndex = .Guild.GuildIndex
                  
        ' Only check the limit if we are creating a new role
        If RoleId = 0 Then
            If UBound(GuildList(GuildIndex).Roles) >= GetLimitOfGuildRoles(GuildIndex) Then
                Call WriteErrorMsg(UserIndex, "El clan no posee mas roles disponibles")
                Exit Sub
            End If
        End If
    
        With GuildList(.Guild.GuildIndex)
        
            If RoleId = 0 Then
                'create role on DB
                RoleId = CreateRoleFromDB(.IdGuild, RoleName, Permissions)
                RoleIndex = UBound(.Roles) + 1
                ReDim Preserve .Roles(1 To RoleIndex) As GuildRoleType
                .Roles(RoleIndex).IdRole = RoleId
                .Roles(RoleIndex).IsDeleteable = True
                .Roles(RoleIndex).RoleName = RoleName
                
                Call InitializeGuildRolePermissions(GuildIndex, RoleIndex)
                Call LoadRolePermissionsFromDB(.IdGuild, RoleId)
                If GuildLastMemberOnline(GuildIndex) > 0 Then
                    For J = 1 To GuildLastMemberOnline(GuildIndex)
                        Call WriteGuildRolesList(.OnlineMembers(J).MemberUserIndex)
                    Next J
                End If
            Else
                For I = 1 To UBound(.Roles)
                    If .Roles(I).IdRole = RoleId Then
                        'modify role on DB
                        Call modGuild_DB.RoleModifyFromDB(RoleId, RoleName, Permissions)
                        
                        .Roles(I).RoleName = RoleName
                        Call InitializeGuildRolePermissions(GuildIndex, I)
                        Call LoadRolePermissionsFromDB(.IdGuild, RoleId)
                        
                        If GuildLastMemberOnline(GuildIndex) > 0 Then
                            For J = 1 To GuildLastMemberOnline(GuildIndex)
                                Call WriteGuildRolesList(.OnlineMembers(J).MemberUserIndex)
                            Next J
                        End If

                        Exit For
                    End If
                Next I
            End If
        End With
    End With
Exit Sub
  
ErrHandler:
Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub RoleUpsert de modGuild_Functions.bas")
End Sub

Public Sub RoleDelete(ByVal GuildIndex As Long, RoleId As Long, UserIndex As Integer)
On Error GoTo ErrHandler
    Dim J As Integer
    Dim RoleIndex As Integer
    
    If Not HasPermission(UserIndex, ROLE_MODIFY) Then
        Call WriteConsoleMsg(UserIndex, NOPERMISSIONOFGUILD, FontTypeNames.FONTTYPE_INFO, info)
        Exit Sub
    End If
    
    ' Check if the role can be deleted
    For J = 1 To UBound(GuildList(GuildIndex).Roles)
        If GuildList(GuildIndex).Roles(J).IdRole = RoleId Then
            If Not GuildList(GuildIndex).Roles(J).IsDeleteable Then
                Call WriteErrorMsg(UserIndex, "El rol " & GuildList(GuildIndex).Roles(J).RoleName & " no puede ser eliminado.")
                Exit Sub
            End If
                        
            RoleIndex = J
            Exit For
        End If
    Next J
    
    ' If we haven't found the role, then we exit and do nothing.
    If RoleIndex = 0 Then Exit Sub
           
    Dim I As Integer
    With GuildList(GuildIndex)
      
        For J = 1 To .MemberCount
            If .Members(J).IdRole = RoleId Then
                Call WriteErrorMsg(UserIndex, "El rol " & .Roles(RoleIndex).RoleName & " no puede ser eliminado porque hay miembros con ese rol")
                Exit Sub
            End If
        Next J
      
      
    
        ' Remove the role from the list
        Dim NextRole As Integer
        NextRole = RoleIndex + 1
        
        If NextRole < UBound(GuildList(GuildIndex).Roles) Then
        For J = RoleIndex To UBound(GuildList(GuildIndex).Roles) - 1
            .Roles(J).IdRole = .Roles(NextRole).IdRole
            .Roles(J).IsDeleteable = .Roles(NextRole).IsDeleteable
            .Roles(J).IsDirty = .Roles(NextRole).IsDirty
            .Roles(J).RoleName = .Roles(NextRole).RoleName
            .Roles(J).PermissionCount = .Roles(NextRole).PermissionCount
            
            ReDim .Roles(J).RolePermission(1 To .Roles(NextRole).PermissionCount)
            For I = 1 To .Roles(NextRole).PermissionCount
                .Roles(J).RolePermission(I).IdPermission = .Roles(NextRole).RolePermission(I).IdPermission
                .Roles(J).RolePermission(I).IsEnabled = .Roles(NextRole).RolePermission(I).IsEnabled
                .Roles(J).RolePermission(I).Key = .Roles(NextRole).RolePermission(I).Key
                .Roles(J).RolePermission(I).PermissionName = .Roles(NextRole).RolePermission(I).PermissionName
            Next I
        Next J
        End If
        
        ReDim Preserve GuildList(GuildIndex).Roles(1 To UBound(GuildList(GuildIndex).Roles) - 1)
        
        ' Delete the role from the DB.
        Call modGuild_DB.RoleDeleteFromDB(RoleId, .IdGuild)
   
        ' Send the list again.
        If GuildLastMemberOnline(GuildIndex) > 0 Then
            For J = 1 To GuildLastMemberOnline(GuildIndex)
                Call WriteGuildRolesList(.OnlineMembers(J).MemberUserIndex)
            Next J
        End If
    End With
Exit Sub
  
ErrHandler:
Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub RoleDelete de modGuild_Functions.bas")
End Sub

Public Function GetRoleIndexFromRoleId(ByVal GuildIndex, ByVal RoleId As Long) As Integer
On Error GoTo ErrHandler
    Dim I As Integer
    
    For I = 1 To UBound(GuildList(GuildIndex).Roles)
        If GuildList(GuildIndex).Roles(I).IdRole = RoleId Then
            GetRoleIndexFromRoleId = I
            Exit Function
        End If
    Next I
Exit Function
  
ErrHandler:
Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetRoleIndexFromRoleId de modGuild_Functions.bas")
End Function


Public Function CanUseGuildNameByReservation(ByVal AccountEmail As String, ByVal GuildName As String)
On Error GoTo ErrHandler
    Dim I As Integer
    
    AccountEmail = UCase(Trim$(AccountEmail))
    GuildName = UCase(Trim$(GuildName))
    
    For I = 1 To GuildConfiguration.ReservedNamesQty
        With GuildConfiguration.ReservedNames(I)
            If .GuildName = GuildName Then
                CanUseGuildNameByReservation = (.AccountEmail = AccountEmail)
                Exit Function
            End If
        End With
        
    Next I
    
    CanUseGuildNameByReservation = True
Exit Function
  
ErrHandler:
Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function CanUseGuildNameByReservation de modGuild_Functions.bas")
End Function

Public Sub SaveAllGuilds()
On Error GoTo ErrHandler:

    Dim I As Integer
    
    For I = 1 To MaxGuildQty
        With GuildList(I)
            ' Save guild header (GUILD_INFO table)
            Call modGuild_DB.UpdateGuildStats(I)
            
            ' Save the current quest status
            If GuildList(I).CurrentQuest.IdQuest > 0 Then
                Call DeleteCurrentQuestStatus(.IdGuild)
                Call modGuild_DB.SaveCurrentQuestStatus(I)
            End If
        End With
    Next I

    Exit Sub
ErrHandler:
Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function SaveAllGuilds de modGuild_Functions.bas")
End Sub
