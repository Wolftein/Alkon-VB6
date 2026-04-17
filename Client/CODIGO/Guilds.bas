Attribute VB_Name = "Guilds"
'@Folder("Guild")

Option Explicit
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Public Const GP_EDIT_GUILD_DESC As String = "EDIT_GUILD_DESC"
Public Const GP_RIGHT_HAND_ASSIGN As String = "RIGHT_HAND_ASSIGN"
Public Const GP_ROLE_ASSIGN As String = "ROLE_ASSIGN"
Public Const GP_ROLE_CREATE_DELETE As String = "ROLE_CREATE_DELETE"
Public Const GP_ROLE_MODIFY As String = "ROLE_MODIFY"
Public Const GP_BANK_DEPOSIT_ITEM As String = "BANK_DEPOSIT_ITEM"
Public Const GP_BANK_WITHDRAW_ITEM As String = "BANK_WITHDRAW_ITEM"
Public Const GP_BANK_DEPOSIT_GOLD As String = "BANK_DEPOSIT_GOLD"
Public Const GP_BANK_WITHDRAW_GOLD As String = "BANK_WITHDRAW_GOLD"
Public Const GP_MEMBER_ACCEPT As String = "MEMBER_ACCEPT"
Public Const GP_MEMBER_KICK As String = "MEMBER_KICK"

Public Const MAX_GUILD_NAME_LEN As Byte = 25

Public Const ID_ROLE_LEADER As Integer = 1
Public Const ID_ROLE_RIGHTHAND As Integer = 2

Public Enum eRoleAction
    Assign = 1
    Create
    Delete
End Enum

Public Type tGuildUserMember
    UserName    As String
    UserId      As Long
    RoleId      As Integer
    IsOnline    As Boolean
    IsRightHand As Boolean
End Type

Public Type tGuildRolePermission
    PermissionId    As Integer
    Key             As String
    
End Type
Public Type tGuildRole
    RoleName    As String
    RoleId      As String
    DeleteEnabled As Boolean
    RenameEnabled As Boolean
    UpdatePermissionsEnabled As Boolean
    PermissionsQty As Integer
    Permissions() As tGuildRolePermission
End Type

Public Type tGuildBank
    Box As Integer
    Slot As Integer
    IdObject As Integer
    Amount As Long
    CanUse As Boolean
End Type

Public Enum eQuestUserAlign
    Ciuda = 1
    Crimi
    Army
    Legion
    Neutral
End Enum

Public Enum eQuestRequirement
    NpcKill = 1
    ObjCollect
    UserKill
End Enum

Public Enum eQuestUpdateEvent
    EventNpcKill = 1
    EventObjectCollect
    EventUserKill
    EventQuestFinished
End Enum

Public Enum eGuildAlignment
    IsNeutral = 1
    IsReal
    IsEvil
End Enum

Public Type GuildQuestReq
    Id As Integer
    Title As String
    Obtained As Boolean
End Type

Public Type GuildReqUpgradeType
    Name As String
    Description As String
    IconGraph As Integer
    GoldCost As Long
    ContributionCost As Long
    UpgradeRequired() As Integer
    QuestRequired() As GuildQuestReq
End Type

Public Type GuildUpgradeGroupConfig
    UpgradeQty As Integer
    Upgrades() As Integer
End Type

Public GuildUpgradesGroup() As GuildUpgradeGroupConfig
Public GuildUpgrades() As GuildReqUpgradeType

Public Type GuildUpgradeType
    IdUpgrade As Integer
    UpgradeLevel As Integer
    UpgradeDate As Date
    UpgradeBy As Long
    IsEnabled As Boolean
End Type

Public Enum eChangeGuildInfo
    MaxMembersQtyChange = 1
    MaxRolesQtyChange
    MaxSlotsBankQtyChange
    MaxBoxesBankQtyChange
    MaxContributionChange
    ContributionAvailableChange
    EnableBank
    BankGoldChange
    CompletedQuestAdded
End Enum

Public Type tGuildInfo

    Name                        As String
    Alignment                   As eGuildAlignment
    IdLeader                    As Long
    IdRightHand                 As Long
    MemberCount                 As Integer
    MaxMemberQty                As Integer
    ContributionAvailable       As Long
    MaxContributionAvailable    As Long
    ContributionEarned          As Long
    CreationTime                As Date
    Description                 As String
    IdGuild                     As Long
    Status                      As Integer
    BankGold                    As Long
    IdRolOwn                    As Integer
    MaxSlotBank                 As Integer
    MaxBoxesBank                As Integer
    MaxRoles                    As Integer
    CurrentQuest            As tQuest
    IsFullFormOpen          As Boolean
    BankAvalaible           As Boolean
    
    Members()               As tGuildUserMember
    Roles()                 As tGuildRole
    
    Quest As tGuildQuestsStatus
    
    Bank()                  As tGuildBank
    Upgrades()              As GuildUpgradeType
End Type

Public GuildCreation As tGuildInfo

Public Enum eExchangeType
    IsGold = 1
    IsObject
End Enum

Public Enum eExchangeAction
    Withdraw = 1
    Deposit
End Enum

Public Enum eChangeMember
    OnlineChange = 1
    RoleChange
    GoldGBChange
End Enum

Public Enum eMemberAction
    KickMember = 1
    SendInvitation
    LeaveGuild
End Enum

Public Type tGuildTempInvitacion
    InvitedByUserName     As String
    GuildIndex    As Integer
    GuildName    As String
    InvitationIndex           As Integer
    ExpirationDate As Date
End Type

Private GuildCurrentInvitation As tGuildTempInvitacion
Private GuildInvitationMessageBoxHandler As New clsGuildMsgBoxHandler

Public Sub KickGuildMember(ByVal MemberName As String)
    Dim Resp As Byte
    Dim KickedId As Long
    Dim I As Integer
    Dim Messege As String, Title As String
    Dim ActionToSend As eMemberAction
    Dim IsRightHand As Boolean
    
    KickedId = 0
    
    If MemberName = UserName Then
        ActionToSend = LeaveGuild
        Title = "Salir de Clan"
        Messege = "¿Estas seguro que quieres salir del clan?"
    Else
        ActionToSend = KickMember
        Title = "Expulsion de Clan"
        Messege = "¿Esta seguro que desea expulsar a " & MemberName & " del clan?"
    End If
    
    For I = 1 To PlayerData.Guild.MemberCount
        If PlayerData.Guild.Members(I).UserName = MemberName Then
            KickedId = PlayerData.Guild.Members(I).UserId
            IsRightHand = PlayerData.Guild.Members(I).RoleId = ID_ROLE_RIGHTHAND
        End If
    Next I
    
    If KickedId = 0 Then Exit Sub
    
    Resp = MsgBox(Messege, vbYesNo, Title)

    If Resp = vbYes Then
        Call WriteGuildMember(PlayerData.Guild.IdGuild, KickedId, ActionToSend)
    End If
    
    Exit Sub
End Sub

Public Function HasPermission(ByVal CheckKey As String) As Boolean

    Dim I As Integer, J As Integer
    
    If PlayerData.Guild.IdGuild <= 0 Then Exit Function
    
    HasPermission = False
    For I = 1 To UBound(PlayerData.Guild.Roles)
        If (PlayerData.Guild.Roles(I).RoleId = PlayerData.Guild.IdRolOwn) Then
            For J = 1 To PlayerData.Guild.Roles(I).PermissionsQty
                    If PlayerData.Guild.Roles(I).Permissions(J).Key = CheckKey Then
                        HasPermission = True
                        Exit For
                    End If
            Next J
            Exit For
        End If
    Next I
    
    Exit Function
End Function

Public Sub UpdateFormInfo()

    If frmGuildMain.Visible Then
        Call frmGuildMain.DisableOptions
    End If
    
    If frmGuildEditRoles.Visible Or frmGuildRolesList.Visible Then
        Call frmGuildRolesList.CleanUCs
        Call frmGuildRolesList.DrawRoles
        Call frmGuildRolesList.UpdateRoleList
    End If
    
    If frmGuildBank.Visible Then
        Call frmGuildBank.EnableButtons
    End If
    
    If frmGuildMembers.Visible Then
        Call frmGuildMembers.CleanUCs
        Call frmGuildMembers.InitializeUCs
        
    End If
    
    If frmGuildQuests.Visible Then
        If PlayerData.Guild.Quest.Id <> 0 Then
            Call frmGuildMain.LoadForm(frmGuildQuestActive, "Mision actual")
            Call frmGuildQuestActive.ShowData
            
        End If
    End If
    
      If frmGuildQuestActive.Visible Then
        If PlayerData.Guild.Quest.Id = 0 Then
            Call frmGuildMain.LoadForm(frmGuildQuests, "Misiones")
            Call frmGuildQuests.ShowData
        End If
    End If
        
    If frmGuildUpgrades.Visible Then
        frmGuildUpgrades.LoadUpgrades
    End If
    
End Sub

Public Function GetQtyGuildBankObjects() As Integer
On Error GoTo ErrHandler

    If ((Not PlayerData.Guild.Bank) = -1) Then
        GetQtyGuildBankObjects = 0
    Else
        GetQtyGuildBankObjects = UBound(PlayerData.Guild.Bank)
    End If

  Exit Function

ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetGuildBankQtyObjects de Guilds.bas")

End Function

Public Sub ResetGuildInfo()
On Error GoTo ErrHandler
    With PlayerData.Guild
        .Alignment = 0
                
        .BankAvalaible = False
        .BankGold = 0
        .ContributionAvailable = 0
        .ContributionEarned = 0
        .CreationTime = 0
        .Description = ""
        .IdGuild = 0
        .IdLeader = 0
        .IdRightHand = 0
        .IdRolOwn = 0
        .IsFullFormOpen = False
        .MaxBoxesBank = 0
        .MaxContributionAvailable = 0
        .MaxMemberQty = 0
        .MaxRoles = 0
        .MaxSlotBank = 0
        .MemberCount = 0
        .Name = ""
        .Quest.CompletedQuantiy = 0
        .Quest.CurrentStage = 0
        .Quest.Id = 0
        '.Quest.StartedDateTime = ""
        .Status = 0
        
        .BankAvalaible = False
        Erase .Bank
        Erase .Upgrades
        Erase .Members
        Erase .Roles
        Erase .Quest.Completed
        Call modQuests.CleanCurrentQuestData
    End With
    
    Exit Sub
    
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function ResetGuildInfo de Guilds.bas")
End Sub

Public Function GetNameOfAlignment(ByVal Align As eGuildAlignment)
    
    Select Case Align
        Case IsNeutral
            GetNameOfAlignment = "Neutral"
        Case IsReal
            GetNameOfAlignment = "Real"
        Case IsEvil
            GetNameOfAlignment = "Del Caos"
    End Select
    Exit Function
End Function

Public Sub GuildInvitation(ByRef InvitedByUserName As String, ByVal GuildIndex As Integer, ByRef GuildName As String, _
                                ByVal InvitationIndex As Integer, ByVal InvitationLifeTimeInMinutes As Integer)
    With GuildCurrentInvitation
        .InvitedByUserName = InvitedByUserName
        .GuildIndex = GuildIndex
        .GuildName = GuildName
        .InvitationIndex = InvitationIndex
        .ExpirationDate = DateAdd("n", InvitationLifeTimeInMinutes, Now)
    End With
    
    Call modMessages.ShowConsoleMessage(eMessageId.Guild_Invitation_User_Invitated, InvitedByUserName, GuildName)
End Sub

Public Sub GuildPendingInvitation()
    Call frmMessageBoxYesNo.ShowMessage(GuildCurrentInvitation.InvitedByUserName & " te ha enviado una invitación para unirte al clan " & GuildCurrentInvitation.GuildName & " ¿Deseas aceptarla?", GuildInvitationMessageBoxHandler)
End Sub

Public Sub GuildCleanInvitation()
    With GuildCurrentInvitation
        .InvitedByUserName = ""
        .GuildName = ""
        .GuildIndex = -1
        .InvitationIndex = -1
        .ExpirationDate = vbDate
    End With
End Sub


Public Function InvitationEmpty() As Boolean
    InvitationEmpty = GuildCurrentInvitation.InvitationIndex = -1
End Function

Public Function GetQtyGuildUpgrades() As Integer
On Error GoTo ErrHandler

    If ((Not PlayerData.Guild.Upgrades) = -1) Then
        GetQtyGuildUpgrades = 0
    Else
        GetQtyGuildUpgrades = UBound(PlayerData.Guild.Upgrades)
    End If

  Exit Function

ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetQtyGuildUpgrades de Guilds.bas")

End Function


Public Function GetQtyGuildUpgradesList() As Integer
On Error GoTo ErrHandler

    If ((Not GuildUpgrades) = -1) Then
        GetQtyGuildUpgradesList = 0
    Else
        GetQtyGuildUpgradesList = UBound(GuildUpgrades)
    End If

  Exit Function

ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetQtyGuildUpgradesList de Guilds.bas")

End Function

Public Function GetQtyGuildUpgradesGroup() As Integer
On Error GoTo ErrHandler

    If ((Not GuildUpgradesGroup) = -1) Then
        GetQtyGuildUpgradesGroup = 0
    Else
        GetQtyGuildUpgradesGroup = UBound(GuildUpgradesGroup)
    End If

  Exit Function

ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetQtyGuildUpgradesGroup de Guilds.bas")

End Function

Public Sub GuildAcceptInvitation()
    Call WriteGuildInvitationResponse(GuildCurrentInvitation.GuildIndex, GuildCurrentInvitation.InvitationIndex, True)
    Call GuildCleanInvitation
End Sub

Public Sub GuildRejectInvitation()
    Call WriteGuildInvitationResponse(GuildCurrentInvitation.GuildIndex, GuildCurrentInvitation.InvitationIndex, False)
    Call GuildCleanInvitation
End Sub
