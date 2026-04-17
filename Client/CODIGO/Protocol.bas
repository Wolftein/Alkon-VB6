Attribute VB_Name = "Protocol"
'**************************************************************
' Protocol.bas - Handles all incoming / outgoing messages for client-server communications.
' Uses a binary protocol designed by myself.
'
' Designed and implemented by Juan Martín Sotuyo Dodero (Maraxus)
' (juansotuyo@gmail.com)
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

''
'Handles all incoming / outgoing packets for client - server communications
'The binary prtocol here used was designed by Juan Martín Sotuyo Dodero.
'This is the first time it's used in Alkon, though the second time it's coded.
'This implementation has several enhacements from the first design.
'
' @file     Protocol.bas
' @author   Juan Martín Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version  1.0.0
' @date     20060517
'@Folder("Protocol")
Option Explicit

''
' TODO : /BANIP y /UNBANIP ya no trabajan con nicks. Esto lo puede mentir en forma local el cliente con un paquete a NickToIp

''
'When we have a list of strings, we use this to separate them and prevent
'having too many string lengths in the queue. Yes, each string is NULL-terminated :P
Private Const SEPARATOR As String * 1 = vbNullChar

Public Const CustomPath As String = "\INIT\custom.dat"

Public AdminMsg As Byte
Public InfoMsg As Byte
Public GuildMsg As Byte
Public PartyMsg As Byte
Public CombateMsg As Byte
Public TrabajoMsg As Byte

Public Enum eMessageType
    None = -1
    Info = 0
    Admin = 1
    Guild = 2
    Party = 3
    combate = 4
    Trabajo = 5
    m_MOTD = 6
End Enum

Public Enum eConsoleType
    General '0
    Acciones
    Agrupaciones
    Custom
    Last '4 (Used for redim)
End Enum

Public Enum eWorkerStoreAction
    WorkerStoreGetRecipes
    WorkerStoreCreate
    WorkerStoreClose
    WorkerStoreCraftItem
End Enum

Public Enum eWorkerStoreServerSubAction
    ShowStore = 1
    OpenFormForCreation = 2
    OpenStore = 3
    ItemCrafted = 4
End Enum

Private Type UpgradeType
      IdUpgrade As Integer
      UpgradeDate As String
      UpgradeBy As String
      IsEnabled As Boolean
      UpgradeLevel As Byte
End Type


#If EnableSecurity = 0 Then

Private Enum ServerPacketID
    Connected               ' CONNECTED
    Logged                  ' LOGGED
    RemoveDialogs           ' QTDL
    RemoveCharDialog        ' QDL
    NavigateChange          ' NAVEG
    Disconnect              ' FINOK
    CommerceEnd             ' FINCOMOK
    BankEnd                 ' FINBANOK
    CommerceInit            ' INITCOM
    BankInit                ' INITBANCO
    UserCommerceInit        ' INITCOMUSU
    UserCommerceEnd         ' FINCOMUSUOK
    UserOfferConfirm
    CommerceChat
    ShowCraftForm
    UpdateSta               ' ASS
    UpdateMana              ' ASM
    UpdateHP                ' ASH
    UpdateGold              ' ASG
    UpdateBankGold
    UpdateExp               ' ASE
    UpdateChallenge
    UpdateChallengeStat
    ChangeMap               ' CM
    PosUpdate               ' PU
    ChatOverHead            ' ||
    ConsoleMsg              ' || - Beware!! its the same as above, but it was properly splitted
    ConsoleFormattedMessage
    GuildChat               ' |+
    ShowMessageBox          ' !!
    ShowFormattedMessageBox
    UserIndexInServer       ' IU
    UserCharIndexInServer   ' IP
    CharacterCreate         ' CC
    CharacterRemove         ' BP
    CharacterChangeNick
    CharacterMove           ' MP, +, * and _ '
    CharacterAttackMovement
    ForceCharMove
    CharacterChange         ' CP
    ObjectCreate            ' HO
    ObjectDelete            ' BO
    ObjectUpdate            ' BU
    BlockPosition           ' BQ
    PlayMusic               ' TM
    PlayEffect              ' TW
    PauseToggle             ' BKW
    RainToggle              ' LLU
    CreateFX                ' CFX
    UpdateUserStats         ' EST
    WorkRequestTarget       ' T01
    SpellAttackResult
    AttackResult
    ChangeInventorySlot     ' CSI
    ChangeBankSlot          ' SBO
    ChangeSpellSlot         ' SHS
    Atributes               ' ATR
    CraftableRecipes
    RestOK                  ' DOK
    ErrorMsg                ' ERR
    Blind                   ' CEGU
    Dumb                    ' DUMB
    ShowSignal              ' MCAR
    ChangeNPCInventorySlot  ' NPCI
    SetUserSalePrice
    UpdateHungerAndThirst   ' EHYS
    MiniStats               ' MEST
    LevelUp                 ' SUNI
    AddForumMsg             ' FMSG
    ShowForumForm           ' MFOR
    SetInvisible            ' NOVER
    MeditateToggle          ' MEDOK
    BlindNoMore             ' NSEGUE
    DumbNoMore              ' NESTUP
    SendSkills              ' SKILLS
    TrainerCreatureList     ' LSTCRI
    CharacterInfo           ' CHRINFO
    ParalizeOK              ' PARADOK
    ShowUserRequest         ' PETICIO
    TradeOK                 ' TRANSOK
    BankOK                  ' BANCOOK
    ChangeUserTradeSlot     ' COMUSUINV
    ChangeUserTradeGold
    
    Pong
    UpdateTagAndStatus
    UpdateUserSpellCooldown
    
    'GM messages
    SpawnList               ' SPL
    ShowSOSForm             ' MSOS
    ShowMOTDEditionForm     ' ZMOTD
    ShowGMPanelForm         ' ABPANEL
    UserNameList            ' LISTUSU
    ShowDenounces
    RecordList
    RecordDetails

    ShowPartyForm
    PartyInvitation
    UpdateStrenghtAndDexterity
    UpdateStrenght
    UpdateDexterity
    UpdateCharacterInfo
    AddSlots
    MultiMessage
    StopWorking
    CancelOfferItem
    ShowMenu
    StrDextRunningOut
    ChatPersonalizado
    
    ' Tournament
    TournamentCompetitorList
    TournamentConfig
    
    ' Account
    AccountRemoveChar
    AccountShow
    AccountPersonaje
    AccountQuestion
    
    'Punishment
    PunishmentTypeList
    
    LoginScreenShow
    CloseForm
    EnableBerserker

    ' Duelos
    Retar
    MensajeDuelo
    OkDueloPublico
    
    'Boveda de cuenta
    AccBankChangeSlot
    AccBankInit
    AccBankUpdateGold
    AccBankEnd
    AccBankRequestPass

    StartPresentEffect
    
    BlacksmithUpgrades
    CarpenterUpgrades
    SendPetSelection
    SendPetList
    
    SendSessionToken
    
    SendMasteries
    
    ShowGuildCreate
    ShowGuildForm
    GuildCreated
    GuildInfo
    GuildRolesList
    GuildMembersList
    GuildUpgradesList
    GuildMemberStatusChange
    GuildBankList
    GuildBankChangeSlot
    GuildSendInvitation
    GuildMemberKicked
    GuildUpgradesAcquired
    GuildInfoChange
    GuildQuestsCompletedList
    GuildCurrentQuestInfo
    GuildQuestUpdateStatus

    WorkerStore
    
    SetIntervals
    
    'Put new packets before this one. LastServerPacketId should be the last element of the enum
    LastServerPacketId
End Enum

'The last existing client packet id.
Private LAST_CLIENT_PACKET_ID As Byte

Private Enum ClientPacketID
    LoginExistingChar       'OLOGIN
    Talk                    ';
    Yell                    '-
    Whisper                 '\
    Walk                    'M
    RequestPositionUpdate   'RPU
    Attack                  'AT
    PickUp                  'AG
    SafeToggle              '/SEG & SEG  (SEG's behaviour has to be coded in the client)
    ResuscitationSafeToggle
    RequestSkills           'ESKI
    RequestStadictis
    CommerceEnd             'FINCOM
    UserCommerceEnd         'FINCOMUSU
    UserCommerceConfirm
    CommerceChat
    BankEnd                 'FINBAN
    UserCommerceOk          'COMUSUOK
    UserCommerceReject      'COMUSUNO
    Drop                    'TI
    DropXY                  'TIXY
    CastSpell               'LH
    LeftClick               'LC
    DoubleClick             'RC
    Work                    'UK
    UseSpellMacro           'UMH
    UseItem                 'USA
    WorkLeftClick           'WLC
    SpellInfo               'INFS
    EquipItem               'EQUI
    ChangeHeading           'CHEA
    ModifySkills            'SKSE
    Train                   'ENTR
    CommerceBuy             'COMP
    BankExtractItem         'RETI
    CommerceSell            'VEND
    BankDeposit             'DEPO
    ForumPost               'DEMSG
    MoveSpell               'DESPHE
    MoveBank
    UserCommerceOffer       'OFRECER
    UserCommerceOfferGold
    Online                  '/ONLINE
    Quit                    '/SALIR
    RequestAccountState     '/BALANCE
    PetStand                '/QUIETO
    PetFollow               '/ACOMPAÑAR
    ReleasePet              '/LIBERAR
    TrainList               '/ENTRENAR
    Rest                    '/DESCANSAR
    Meditate                '/MEDITAR
    Resucitate              '/RESUCITAR
    Heal                    '/CURAR
    Help                    '/AYUDA
    RequestStats            '/EST
    CommerceStart           '/COMERCIAR
    BankStart               '/BOVEDA
    Enlist                  '/ENLISTAR
    Information             '/INFORMACION
    Reward                  '/RECOMPENSA
    UpTime                  '/UPTIME
    PartyLeave              '/SALIRPARTY
    PartyCreate             '/CREARPARTY
    Inquiry                 '/ENCUESTA ( with no params )
    GuildMessage            '/CMSG
    PartyMessage            '/PMSG
    PartyOnline             '/ONLINEPARTY
    CouncilMessage          '/BMSG
    RoleMasterRequest       '/ROL
    GMRequest               '/GM
    bugReport               '/_BUG
    ChangeDescription       '/DESC
    Punishments             '/PENAS
    Gamble                  '/APOSTAR
    InquiryVote             '/ENCUESTA ( with parameters )
    LeaveFaction            '/RETIRAR ( with no arguments )
    BankExtractGold         '/RETIRAR ( with arguments )
    BankDepositGold         '/DEPOSITAR
    Denounce                '/DENUNCIAR
    PartyKick               '/ECHARPARTY
    PartySetLeader          '/PARTYLIDER
    PartyInviteMember       '/INVITAR
    PartyAcceptInvitation
    Ping                    '/PING
    RequestPartyForm
    ItemUpgrade
    GMCommands
    InitCrafting
    Home
    Consultation
    moveItem
    RightClick
    
    PMDeleteList
    PMList

    MenuAction
    
    ' Torneos
    Participar              '/PARTICIPAR
    
    ' Account
    AccountLogin
    AccountLoginChar
    AccountCreateChar
    AccountCreate
    AccountDeleteChar
    AccountRecover
    AccountChangePassword

'desafios - Mithrandir: Se podrï¿½a crear 1 paquete, en vez de 2 (para aceptar y cancelar)
    Chat_desafio
    Cancel_desafio
    Accept_desafio
    Enviardatos_desafio

    ' Duelos
    Retar
    Duelos
    AceptarDuelo
    RechazarDuelo
    CancelarElDuelo
    DueloPublico
    CancelarEspera
    
    'Boveda de cuenta
    AccBankExtractItem
    AccBankDepositItem
    AccBankExtractGold
    AccBankDepositGold
    AccBankStart
    AccBankEnd
    AccBankChangePass

    CraftItem
    SelectPet
    RequestPetSelection
    
    MasteryAssign
    ' New Guild System
    GuildCreate
    GuildMember
    GuildExchange
    GuildRole
    GuildUpgrade
    GuildBankEnd
    GuildQuest
    GuildQuestAddObject
    GuildUserInvitationResponse
    ' Crafting Store
    WorkerStore
    
    'Put new packets before this one. LastClientPacketId should be the last element of the enum
    LastClientPacketId
End Enum
#End If

Public Reader   As BinaryReader ' @NOTE: need refactoring, this is so its easier to use in every handler
Public Writer   As BinaryWriter ' @NOTE: same as above
Public Client   As Network_Client
Private Protocol As Network_Protocol
Private bConnected As Boolean

Public Sub OnNetworkConnect(ByVal Client As Network_Client)
    bConnected = True

    Set Writer = New BinaryWriter
End Sub

Public Sub OnNetworkClose(ByVal Client As Network_Client)
    bConnected = False

    Call ResetAllInfo(True)
End Sub

Public Sub OnNetworkSend(ByVal Client As Network_Client, ByVal Message As BinaryReader)
    ' #SECURITY
End Sub

Public Sub OnNetworkRecv(ByVal Client As Network_Client, ByVal Message As BinaryReader)
   
    Set Reader = Message

    While (Message.GetAvailable() > 0)
        Call HandleIncomingData
    Wend

    Set Reader = Nothing
    
End Sub

Public Sub OnNetworkError(ByVal Client As Network_Client, ByVal Error As Long, ByVal Description As String)
        ' #SECURITY
End Sub


Public Sub Connect(ByVal Address As String, ByVal Service As String)

    If (Protocol Is Nothing) Then
        #If EnableSecurity = 0 Then
            Set Protocol = New Network_Protocol
        
            Call Protocol.Attach(AddressOf OnNetworkConnect, AddressOf OnNetworkClose, AddressOf OnNetworkRecv, AddressOf OnNetworkSend, AddressOf OnNetworkError)
        #Else
            Set Protocol = GetSecureProtocol()
        #End If
    End If
    
    If (Not Client Is Nothing) Then
        Call Client.Close(True)
    End If
    
    Set Client = Aurora_Network.Connect(Address, Service)
    Call Client.SetProtocol(Protocol)
    
End Sub

Public Sub Shutdown(Optional ByVal Forcibly As Boolean = False)
    Call Client.Close(Forcibly)
End Sub

Public Sub Send(ByVal Urgent As Boolean)

    ' While is paused prevent flooding the server
    If (Not pausa) Then
        
        Call Client.Write(Writer)
        
        If (Urgent) Then
            Call Client.Flush
        End If
    
    End If
    
    Call Writer.Clear
End Sub

Public Function IsConnected() As Boolean

    IsConnected = bConnected
    
End Function

Public Function HandleLogin() As Boolean
    If EstadoLogin = E_MODO.Normal Then
        Call WriteLoginExistingChar
    ElseIf EstadoLogin = E_MODO.AccountCreateChar Then
        Call WriteAccountCreateChar
    ElseIf EstadoLogin = E_MODO.AccountCreate Then
        Call WriteAccountCreate
    ElseIf EstadoLogin = E_MODO.AccountLogin Then
        Call WriteAccountLogin
    ElseIf EstadoLogin = E_MODO.AccountLoginChar Then
        Call WriteAccountLoginChar
    ElseIf EstadoLogin = E_MODO.AccountDeleteChar Then
        Call WriteAccountDeleteChar
    ElseIf EstadoLogin = E_MODO.AccountRecover Then
        Call WriteAccountRecover
    ElseIf EstadoLogin = E_MODO.AccountChangePassword Then
        Call WriteAccountChangePassword
    
    ElseIf EstadoLogin = E_MODO.Dados Then
        Call Engine_Audio.PlayMusic("7.mp3")
            
        If frmAccount.Visible Then
            frmAccount.Visible = False
        End If
        
        frmCrearPersonaje.Show
#If EnableSecurity Then
        Call ProtectForm(frmCrearPersonaje)
#End If
        
    End If
End Function

''
' Handles incoming data.

Public Function HandleIncomingData() As Boolean
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
On Error GoTo ErrHandler

    Dim PacketID  As Byte
    PacketID = Reader.ReadInt8
    
    Select Case PacketID
        Case ServerPacketID.Connected               ' CONNECTED
            Call HandleConnected
        
        Case ServerPacketID.Logged                  ' LOGGED
            Call HandleLogged
        
        Case ServerPacketID.RemoveDialogs           ' QTDL
            Call HandleRemoveDialogs
        
        Case ServerPacketID.RemoveCharDialog        ' QDL
            Call HandleRemoveCharDialog
        
        Case ServerPacketID.NavigateChange          ' NAVEG
            Call HandleNavigateChange
        
        Case ServerPacketID.Disconnect              ' FINOK
            Call HandleDisconnect
        
        Case ServerPacketID.CommerceEnd             ' FINCOMOK
            Call HandleCommerceEnd
            
        Case ServerPacketID.CommerceChat
            Call HandleCommerceChat
            
        Case ServerPacketID.SpellAttackResult
            Call HandleSpellAttackResult
            
        Case ServerPacketID.AttackResult
            Call HandleAttackResult
        
        Case ServerPacketID.BankEnd                 ' FINBANOK
            Call HandleBankEnd
        
        Case ServerPacketID.CommerceInit            ' INITCOM
            Call HandleCommerceInit
        
        Case ServerPacketID.BankInit                ' INITBANCO
            Call HandleBankInit
        
        Case ServerPacketID.UserCommerceInit        ' INITCOMUSU
            Call HandleUserCommerceInit
        
        Case ServerPacketID.UserCommerceEnd         ' FINCOMUSUOK
            Call HandleUserCommerceEnd
            
        Case ServerPacketID.UserOfferConfirm
            Call HandleUserOfferConfirm
        
        Case ServerPacketID.ShowCraftForm      ' SFH
            Call HandleShowCraftForm
        
        Case ServerPacketID.UpdateSta               ' ASS
            Call HandleUpdateSta
        
        Case ServerPacketID.UpdateMana              ' ASM
            Call HandleUpdateMana
        
        Case ServerPacketID.UpdateHP                ' ASH
            Call HandleUpdateHP
        
        Case ServerPacketID.UpdateGold              ' ASG
            Call HandleUpdateGold
            
        Case ServerPacketID.UpdateBankGold
            Call HandleUpdateBankGold

        Case ServerPacketID.UpdateExp               ' ASE
            Call HandleUpdateExp
        
        Case ServerPacketID.UpdateChallenge
            Call HandleUpdateChallenge
            
        Case ServerPacketID.UpdateChallengeStat
            Call HandleUpdateChallengeStat
            
        Case ServerPacketID.ChangeMap               ' CM
            Call HandleChangeMap
        
        Case ServerPacketID.PosUpdate               ' PU
            Call HandlePosUpdate
        
        Case ServerPacketID.ChatOverHead            ' ||
            Call HandleChatOverHead
        
        Case ServerPacketID.ConsoleMsg              ' || - Beware!! its the same as above, but it was properly splitted
            Call HandleConsoleMessage
            
        Case ServerPacketID.ConsoleFormattedMessage
            Call HandleConsoleFormattedMessage
            
        Case ServerPacketID.GuildChat
            Call HandleGuildChat
        
        Case ServerPacketID.ShowMessageBox          ' !!
            Call HandleShowMessageBox

        Case ServerPacketID.ShowFormattedMessageBox
            Call HandleShowFormattedMessageBox
        
        Case ServerPacketID.UserIndexInServer       ' IU
            Call HandleUserIndexInServer
        
        Case ServerPacketID.UserCharIndexInServer   ' IP
            Call HandleUserCharIndexInServer
        
        Case ServerPacketID.CharacterCreate         ' CC
            Call HandleCharacterCreate
        
        Case ServerPacketID.CharacterRemove         ' BP
            Call HandleCharacterRemove
        
        Case ServerPacketID.CharacterChangeNick
            Call HandleCharacterChangeNick
            
        Case ServerPacketID.CharacterMove           ' MP, +, * and _ '
            Call HandleCharacterMove
        
        Case ServerPacketID.CharacterAttackMovement
            Call HandleCharacterAttackMovement
            
        Case ServerPacketID.ForceCharMove
            Call HandleForceCharMove
        
        Case ServerPacketID.CharacterChange         ' CP
            Call HandleCharacterChange
        
        Case ServerPacketID.ObjectCreate            ' HO
            Call HandleObjectCreate
        
        Case ServerPacketID.ObjectDelete            ' BO
            Call HandleObjectDelete
                
        Case ServerPacketID.ObjectUpdate            ' BU
            Call HandleObjectUpdate
        
        Case ServerPacketID.BlockPosition           ' BQ
            Call HandleBlockPosition
        
        Case ServerPacketID.PlayMusic               ' TM
            Call HandlePlayMusic
        
        Case ServerPacketID.PlayEffect                ' TW
            Call HandlePlayEffect

        Case ServerPacketID.PauseToggle             ' BKW
            Call HandlePauseToggle
        
        Case ServerPacketID.RainToggle              ' LLU
            Call HandleRainToggle
        
        Case ServerPacketID.CreateFX                ' CFX
            Call HandleCreateFX
        
        Case ServerPacketID.UpdateUserStats         ' EST
            Call HandleUpdateUserStats
        
        Case ServerPacketID.WorkRequestTarget       ' T01
            Call HandleWorkRequestTarget
        
        Case ServerPacketID.ChangeInventorySlot     ' CSI
            Call HandleChangeInventorySlot
        
        Case ServerPacketID.ChangeBankSlot          ' SBO
            Call HandleChangeBankSlot
        
        Case ServerPacketID.ChangeSpellSlot         ' SHS
            Call HandleChangeSpellSlot
        
        Case ServerPacketID.Atributes               ' ATR
            Call HandleAtributes
            
        Case ServerPacketID.CraftableRecipes
            Call HandleCraftableRecipes
        
        Case ServerPacketID.RestOK                  ' DOK
            Call HandleRestOK
        
        Case ServerPacketID.ErrorMsg                ' ERR
            Call HandleErrorMessage
        
        Case ServerPacketID.Blind                   ' CEGU
            Call HandleBlind
        
        Case ServerPacketID.Dumb                    ' DUMB
            Call HandleDumb
        
        Case ServerPacketID.ShowSignal              ' MCAR
            Call HandleShowSignal
        
        Case ServerPacketID.ChangeNPCInventorySlot  ' NPCI
            Call HandleChangeNPCInventorySlot
            
        Case ServerPacketID.SetUserSalePrice
            Call HandleSetUserSalePrice
        
        Case ServerPacketID.UpdateHungerAndThirst   ' EHYS
            Call HandleUpdateHungerAndThirst
        
        Case ServerPacketID.MiniStats               ' MEST
            Call HandleMiniStats
        
        Case ServerPacketID.LevelUp                 ' SUNI
            Call HandleLevelUp
        
        Case ServerPacketID.AddForumMsg             ' FMSG
            Call HandleAddForumMessage
        
        Case ServerPacketID.ShowForumForm           ' MFOR
            Call HandleShowForumForm
        
        Case ServerPacketID.SetInvisible            ' NOVER
            Call HandleSetInvisible
        
        Case ServerPacketID.MeditateToggle          ' MEDOK
            Call HandleMeditateToggle
        
        Case ServerPacketID.BlindNoMore             ' NSEGUE
            Call HandleBlindNoMore
        
        Case ServerPacketID.DumbNoMore              ' NESTUP
            Call HandleDumbNoMore
        
        Case ServerPacketID.SendSkills              ' SKILLS
            Call HandleSendSkills
        
        Case ServerPacketID.TrainerCreatureList     ' LSTCRI
            Call HandleTrainerCreatureList
        
       ' Case ServerPacketID.OfferDetails            ' PEACEDE and ALLIEDE
       '     Call HandleOfferDetails
        
        Case ServerPacketID.ParalizeOK              ' PARADOK
            Call HandleParalizeOK
        
        Case ServerPacketID.ShowUserRequest         ' PETICIO
            Call HandleShowUserRequest
        
        Case ServerPacketID.TradeOK                 ' TRANSOK
            Call HandleTradeOK
        
        Case ServerPacketID.BankOK                  ' BANCOOK
            Call HandleBankOK
        
        Case ServerPacketID.ChangeUserTradeSlot     ' COMUSUINV
            Call HandleChangeUserTradeSlot
            
        Case ServerPacketID.ChangeUserTradeGold
            Call HandleChangeUserTradeGold

        Case ServerPacketID.Pong
            Call HandlePong
        
        Case ServerPacketID.UpdateTagAndStatus
            Call HandleUpdateTagAndStatus
            
        Case ServerPacketID.UpdateUserSpellCooldown
            Call HandleUpdateUserSpellCooldown
         
        '*******************
        'GM messages
        '*******************
        Case ServerPacketID.SpawnList               ' SPL
            Call HandleSpawnList
        
        Case ServerPacketID.ShowSOSForm             ' RSOS and MSOS
            Call HandleShowSOSForm
            
        Case ServerPacketID.ShowDenounces
            Call HandleShowDenounces
            
        Case ServerPacketID.RecordDetails
            Call HandleRecordDetails
            
        Case ServerPacketID.RecordList
            Call HandleRecordList
            
        Case ServerPacketID.ShowMOTDEditionForm     ' ZMOTD
            Call HandleShowMOTDEditionForm
        
        Case ServerPacketID.ShowGMPanelForm         ' ABPANEL
            Call HandleShowGMPanelForm
        
        Case ServerPacketID.UserNameList            ' LISTUSU
            Call HandleUserNameList
        
        Case ServerPacketID.ShowPartyForm
            Call HandleShowPartyForm
        
        Case ServerPacketID.PartyInvitation
            Call HandlePartyInvitation
        
        Case ServerPacketID.UpdateStrenghtAndDexterity
            Call HandleUpdateStrenghtAndDexterity
            
        Case ServerPacketID.UpdateStrenght
            Call HandleUpdateStrenght
            
        Case ServerPacketID.UpdateDexterity
            Call HandleUpdateDexterity
            
        Case ServerPacketID.UpdateCharacterInfo
            Call HandleUpdateCharacterInfo
            
        Case ServerPacketID.AddSlots
            Call HandleAddSlots

        Case ServerPacketID.MultiMessage
            Call HandleMultiMessage
        
        Case ServerPacketID.StopWorking
            Call HandleStopWorking
            
        Case ServerPacketID.CancelOfferItem
            Call HandleCancelOfferItem
            
        Case ServerPacketID.ShowMenu
            Call HandleShowMenu
            
        Case ServerPacketID.StrDextRunningOut
            Call HandleStrDextRunningOut
            
        Case ServerPacketID.ChatPersonalizado
            Call HandleChatPersonalizado

        Case ServerPacketID.TournamentCompetitorList
            Call HandleTournamentCompetitorList
            
        Case ServerPacketID.TournamentConfig
            Call HandleTournamentConfig
        
        Case ServerPacketID.AccountRemoveChar
            Call HandleAccountRemoveChar
        
        Case ServerPacketID.AccountShow
            Call HandleAccountShow
       
        Case ServerPacketID.AccountPersonaje
            Call HandleAccountPersonaje
       
        Case ServerPacketID.AccountQuestion
            Call HandleAccountQuestion
        
        Case ServerPacketID.PunishmentTypeList
            Call HandleGetPunishmentTypeList
            
        Case ServerPacketID.LoginScreenShow
            Call HandleShowLoginScreen
        
        Case ServerPacketID.CloseForm
            Call HandleCloseForm
            
        Case ServerPacketID.EnableBerserker
            Call HandleEnableBerserker

        Case ServerPacketID.OkDueloPublico
            Call HandleOkDueloPublico
            
        Case ServerPacketID.MensajeDuelo
            Call HandleMensajeDuelo
            
        Case ServerPacketID.Retar
            Call HandleRetar
            
        Case ServerPacketID.AccBankChangeSlot
            Call HandleAccBankChangeSlot
            
        Case ServerPacketID.AccBankInit
            Call HandleAccBankInit
            
        Case ServerPacketID.AccBankUpdateGold
            Call HandleAccBankUpdateGold
            
        Case ServerPacketID.AccBankEnd
            Call HandleAccBankEnd
            
        Case ServerPacketID.AccBankRequestPass
            Call HandleAccBankRequestPass
 
        Case ServerPacketID.StartPresentEffect
            Call HandleStartPresentEffect

        Case ServerPacketID.BlacksmithUpgrades
            Call HandleBlacksmithUpgrades
            
        Case ServerPacketID.CarpenterUpgrades
            Call HandleCarpenterUpgrades
            
        Case ServerPacketID.SendPetList
            Call HandleSendPetList
            
        Case ServerPacketID.SendSessionToken
            Call HandleSendSessionToken
        Case ServerPacketID.SendMasteries
            Call HandleSendMasteries

        Case ServerPacketID.GuildInfo
            Call HandleGuildInfo
    
        Case ServerPacketID.GuildRolesList
            Call HandleGuildRolesList
            
        Case ServerPacketID.GuildMembersList
            Call HandleGuildMembersList
            
        Case ServerPacketID.GuildUpgradesList
            Call HandleGuildUpgradesList
            
        Case ServerPacketID.GuildMemberStatusChange
            Call HandleGuildMemberStatusChange
            
        Case ServerPacketID.GuildMemberKicked
            Call HandleGuildMemberKicked
            
        Case ServerPacketID.ShowGuildCreate
            Call HandleShowGuildCreate
            
        Case ServerPacketID.ShowGuildForm
            Call HandleShowGuildForm
            
        Case ServerPacketID.GuildCreated
            Call HandleGuildCreated
        
        Case ServerPacketID.GuildBankList
            Call HandleGuildBankList
            
        Case ServerPacketID.GuildBankChangeSlot
            Call HandleGuildBankChangeSlot
        
        Case ServerPacketID.GuildSendInvitation
            Call HandleGuildSendInvitation
        
        Case ServerPacketID.GuildUpgradesAcquired
            Call HandleGuildUpgradesAcquired
            
        Case ServerPacketID.GuildInfoChange
            Call HandleGuildInfoChange

        Case ServerPacketID.WorkerStore
            Call HandleWorkerStore
        
        Case ServerPacketID.GuildQuestsCompletedList
            Call HandleGuildQuestsCompletedList
            
        Case ServerPacketID.GuildCurrentQuestInfo
            Call HandleGuildCurrentQuestInfo
            
        Case ServerPacketID.GuildQuestUpdateStatus
            Call HandleGuildQuestUpdateStatus
            
        Case ServerPacketID.WorkerStore
            Call HandleWorkerStore
            
        Case ServerPacketID.SetIntervals
            Call HandleSetIntervals

            
#If EnableSecurity Then
        Case Else
            Call HandleIncomingDataEx(PacketID)
#Else
        Case Else
            GoTo ErrHandler
#End If
    End Select
    
    #If EnableSecurity Then
        Call SockReady
    #End If

    Exit Function

ErrHandler:

    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en " & "HandleIncomingData: " & PacketID & " de Protocol.bas")

End Function

Private Sub HandleStartPresentEffect()


    If (Reader.ReadInt8 = 0) Then
       Reader.ReadString8
    End If

End Sub
Private Sub HandleMultiMessage()
'***************************************************
'Author: Unknown
'Last Modification: 11/16/2010
' 09/28/2010: C4b3z0n - Ahora se le saco la "," a los minutos de distancia del /hogar, ya que a veces quedaba "12,5 minutos y 30segundos"
'***************************************************

  
    Dim BodyPart As Byte
    Dim Damage As Integer


    Select Case Reader.ReadInt8
        Case eMessages.DontSeeAnything
            Call AddtoRichTextBox(frmMain.RecTxt(0), MENSAJE_NO_VES_NADA_INTERESANTE, 65, 190, 156, False, False, True, eMessageType.Info)
        
        Case eMessages.NPCSwing
            Call AddtoRichTextBox(frmMain.RecTxt(0), MENSAJE_CRIATURA_FALLA_GOLPE, 255, 0, 0, True, False, True, eMessageType.combate)
        
        Case eMessages.NPCKillUser
            Call AddtoRichTextBox(frmMain.RecTxt(0), MENSAJE_CRIATURA_MATADO, 255, 0, 0, True, False, True, eMessageType.combate)
        
        Case eMessages.BlockedWithShieldUser
            Call AddtoRichTextBox(frmMain.RecTxt(0), MENSAJE_RECHAZO_ATAQUE_ESCUDO, 255, 0, 0, True, False, True, eMessageType.combate)
        
        Case eMessages.BlockedWithShieldOther
            Call AddtoRichTextBox(frmMain.RecTxt(0), MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO, 255, 0, 0, True, False, True, eMessageType.combate)
        
        Case eMessages.UserSwing
            Call AddtoRichTextBox(frmMain.RecTxt(0), MENSAJE_FALLADO_GOLPE, 255, 0, 0, True, False, True, eMessageType.combate)
        
        Case eMessages.SafeModeOn
            Call frmMain.ControlSM(eSMType.sSafemode, True)
        
        Case eMessages.SafeModeOff
            Call frmMain.ControlSM(eSMType.sSafemode, False)
        
        Case eMessages.ResuscitationSafeOff
            Call frmMain.ControlSM(eSMType.sResucitation, False)
         
        Case eMessages.ResuscitationSafeOn
            Call frmMain.ControlSM(eSMType.sResucitation, True)
        
        Case eMessages.NobilityLost
            Call AddtoRichTextBox(frmMain.RecTxt(0), MENSAJE_PIERDE_NOBLEZA, 255, 0, 0, False, False, True, eMessageType.Info)
        
        Case eMessages.CantUseWhileMeditating
            Call AddtoRichTextBox(frmMain.RecTxt(0), MENSAJE_USAR_MEDITANDO, 255, 0, 0, False, False, True, eMessageType.Info)
        
        Case eMessages.NPCHitUser
            Select Case Reader.ReadInt8()
                Case bCabeza
                    Call AddtoRichTextBox(frmMain.RecTxt(0), MENSAJE_GOLPE_CABEZA & CStr(Reader.ReadInt16()) & "!!", 255, 0, 0, True, False, True, eMessageType.combate)
                
                Case bBrazoIzquierdo
                    Call AddtoRichTextBox(frmMain.RecTxt(0), MENSAJE_GOLPE_BRAZO_IZQ & CStr(Reader.ReadInt16()) & "!!", 255, 0, 0, True, False, True, eMessageType.combate)
                
                Case bBrazoDerecho
                    Call AddtoRichTextBox(frmMain.RecTxt(0), MENSAJE_GOLPE_BRAZO_DER & CStr(Reader.ReadInt16()) & "!!", 255, 0, 0, True, False, True, eMessageType.combate)
                
                Case bPiernaIzquierda
                    Call AddtoRichTextBox(frmMain.RecTxt(0), MENSAJE_GOLPE_PIERNA_IZQ & CStr(Reader.ReadInt16()) & "!!", 255, 0, 0, True, False, True, eMessageType.combate)
                
                Case bPiernaDerecha
                    Call AddtoRichTextBox(frmMain.RecTxt(0), MENSAJE_GOLPE_PIERNA_DER & CStr(Reader.ReadInt16()) & "!!", 255, 0, 0, True, False, True, eMessageType.combate)
                
                Case bTorso
                    Call AddtoRichTextBox(frmMain.RecTxt(0), MENSAJE_GOLPE_TORSO & CStr(Reader.ReadInt16() & "!!"), 255, 0, 0, True, False, True, eMessageType.combate)
            End Select
        
        Case eMessages.UserHitNPC
            Call AddtoRichTextBox(frmMain.RecTxt(0), MENSAJE_GOLPE_CRIATURA_1 & CStr(Reader.ReadInt32()) & MENSAJE_2, 255, 0, 0, True, False, True, eMessageType.combate)
        
        Case eMessages.UserAttackedSwing
            Call AddtoRichTextBox(frmMain.RecTxt(0), MENSAJE_1 & charlist(Reader.ReadInt16()).Nombre & MENSAJE_ATAQUE_FALLO, 255, 0, 0, True, False, True, eMessageType.combate)
        
        Case eMessages.UserHittedByUser
            Dim AttackerName As String
            
            AttackerName = GetRawName(charlist(Reader.ReadInt16()).Nombre)
            BodyPart = Reader.ReadInt8()
            Damage = Reader.ReadInt16()
            
            Select Case BodyPart
                Case bCabeza
                    Call AddtoRichTextBox(frmMain.RecTxt(0), MENSAJE_1 & AttackerName & MENSAJE_RECIVE_IMPACTO_CABEZA & Damage & MENSAJE_2, 255, 0, 0, True, False, True, eMessageType.combate)
                
                Case bBrazoIzquierdo
                    Call AddtoRichTextBox(frmMain.RecTxt(0), MENSAJE_1 & AttackerName & MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ & Damage & MENSAJE_2, 255, 0, 0, True, False, True, eMessageType.combate)
                
                Case bBrazoDerecho
                    Call AddtoRichTextBox(frmMain.RecTxt(0), MENSAJE_1 & AttackerName & MENSAJE_RECIVE_IMPACTO_BRAZO_DER & Damage & MENSAJE_2, 255, 0, 0, True, False, True, eMessageType.combate)
                
                Case bPiernaIzquierda
                    Call AddtoRichTextBox(frmMain.RecTxt(0), MENSAJE_1 & AttackerName & MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ & Damage & MENSAJE_2, 255, 0, 0, True, False, True, eMessageType.combate)
                
                Case bPiernaDerecha
                    Call AddtoRichTextBox(frmMain.RecTxt(0), MENSAJE_1 & AttackerName & MENSAJE_RECIVE_IMPACTO_PIERNA_DER & Damage & MENSAJE_2, 255, 0, 0, True, False, True, eMessageType.combate)
                
                Case bTorso
                    Call AddtoRichTextBox(frmMain.RecTxt(0), MENSAJE_1 & AttackerName & MENSAJE_RECIVE_IMPACTO_TORSO & Damage & MENSAJE_2, 255, 0, 0, True, False, True, eMessageType.combate)
            End Select
        
        Case eMessages.UserHittedUser

            Dim VictimName As String
            
            VictimName = GetRawName(charlist(Reader.ReadInt16()).Nombre)
            BodyPart = Reader.ReadInt8()
            Damage = Reader.ReadInt16()
            
            Select Case BodyPart
                Case bCabeza
                    Call AddtoRichTextBox(frmMain.RecTxt(0), MENSAJE_PRODUCE_IMPACTO_1 & VictimName & MENSAJE_PRODUCE_IMPACTO_CABEZA & Damage & MENSAJE_2, 255, 0, 0, True, False, True, eMessageType.combate)
                
                Case bBrazoIzquierdo
                    Call AddtoRichTextBox(frmMain.RecTxt(0), MENSAJE_PRODUCE_IMPACTO_1 & VictimName & MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ & Damage & MENSAJE_2, 255, 0, 0, True, False, True, eMessageType.combate)
                
                Case bBrazoDerecho
                    Call AddtoRichTextBox(frmMain.RecTxt(0), MENSAJE_PRODUCE_IMPACTO_1 & VictimName & MENSAJE_PRODUCE_IMPACTO_BRAZO_DER & Damage & MENSAJE_2, 255, 0, 0, True, False, True, eMessageType.combate)
                
                Case bPiernaIzquierda
                    Call AddtoRichTextBox(frmMain.RecTxt(0), MENSAJE_PRODUCE_IMPACTO_1 & VictimName & MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ & Damage & MENSAJE_2, 255, 0, 0, True, False, True, eMessageType.combate)
                
                Case bPiernaDerecha
                    Call AddtoRichTextBox(frmMain.RecTxt(0), MENSAJE_PRODUCE_IMPACTO_1 & VictimName & MENSAJE_PRODUCE_IMPACTO_PIERNA_DER & Damage & MENSAJE_2, 255, 0, 0, True, False, True, eMessageType.combate)
                
                Case bTorso
                    Call AddtoRichTextBox(frmMain.RecTxt(0), MENSAJE_PRODUCE_IMPACTO_1 & VictimName & MENSAJE_PRODUCE_IMPACTO_TORSO & Damage & MENSAJE_2, 255, 0, 0, True, False, True, eMessageType.combate)
            End Select
        
        Case eMessages.WorkRequestTarget
            UsingSkill = Reader.ReadInt8()
            
            frmMain.MousePointer = 2
            
            Select Case UsingSkill
                Case Pesca
                    Call AddtoRichTextBox(frmMain.RecTxt(0), MENSAJE_TRABAJO_PESCA, 100, 100, 120, 0, 0, , eMessageType.Trabajo)
                
                Case Robar
                    Call AddtoRichTextBox(frmMain.RecTxt(0), MENSAJE_TRABAJO_ROBAR, 100, 100, 120, 0, 0, , eMessageType.Trabajo)
                
                Case Talar
                    Call AddtoRichTextBox(frmMain.RecTxt(0), MENSAJE_TRABAJO_TALAR, 100, 100, 120, 0, 0, , eMessageType.Trabajo)
                
                Case Mineria
                    Call AddtoRichTextBox(frmMain.RecTxt(0), MENSAJE_TRABAJO_MINERIA, 100, 100, 120, 0, 0, , eMessageType.Trabajo)
                
                Case FundirMetal
                    Call AddtoRichTextBox(frmMain.RecTxt(0), MENSAJE_TRABAJO_FUNDIRMETAL, 100, 100, 120, 0, 0, , eMessageType.Trabajo)
                
                Case Proyectiles
                    Call AddtoRichTextBox(frmMain.RecTxt(0), MENSAJE_TRABAJO_PROYECTILES, 100, 100, 120, 0, 0, , eMessageType.Trabajo)
            End Select
        
        Case eMessages.SpellCastRequestTarget
            CastedSpellIndex = Reader.ReadInt
            CastedSpellNumber = Reader.ReadInt
                        
            UsingSkill = eSkill.Magia
            frmMain.MousePointer = MousePointerConstants.vbCustom
            
            Call frmMain.RecalculateMousePointerForSpell
            
            Call AddtoRichTextBox(frmMain.RecTxt(0), MENSAJE_TRABAJO_MAGIA, 100, 100, 120, 0, 0, , eMessageType.Trabajo)

        Case eMessages.HaveKilledUser
            Dim KilledUser As Integer
            Dim ELV As Long
            
            KilledUser = Reader.ReadInt16
            ELV = Reader.ReadInt32
            
            Call AddtoRichTextBox(frmMain.RecTxt(0), MENSAJE_HAS_MATADO_A & charlist(KilledUser).Nombre & MENSAJE_22, 255, 0, 0, True, False, True, eMessageType.combate)
            Call AddtoRichTextBox(frmMain.RecTxt(0), charlist(KilledUser).Nombre & MENSAJE_ERA_NIVEL & ELV & ".", 255, 0, 0, True, False, True, eMessageType.combate)
 
        Case eMessages.UserKill
            Dim KillerUser As Integer
            
            KillerUser = Reader.ReadInt16
            
            Call ShowConsoleMsg(charlist(KillerUser).Nombre & MENSAJE_TE_HA_MATADO, 255, 0, 0, True, False)

        Case eMessages.EarnExp
            'Call ShowConsoleMsg(MENSAJE_HAS_GANADO_EXPE_1 & .ReadInt32 & MENSAJE_HAS_GANADO_EXPE_2, 255, 0, 0, True, False)
        
        Case eMessages.GoHome
            Dim Hogar As String
            Dim Tiempo As Integer
            Dim msg As String
            
            Tiempo = Reader.ReadInt16
            Hogar = Reader.ReadString8
            
            If Tiempo >= 60 Then
                If Tiempo Mod 60 = 0 Then
                    msg = Tiempo / 60 & " minuto" & IIf(Tiempo = 60, ".", "s.")
                Else
                    msg = CInt(Tiempo \ 60) & " minuto" & IIf(Tiempo = 60, "s", "") & " y " & Tiempo Mod 60 & " segundos."  'Agregado el CInt() asi el número no es con , [C4b3z0n - 09/28/2010]
                End If
            Else
                msg = Tiempo & " segundos."
            End If
            
            Call ShowConsoleMsg("Tu viaje de vuelta a tu hogar durará " & msg, 255, 0, 0, True)
            Traveling = True
        
        Case eMessages.FinishHome
            Call ShowConsoleMsg(MENSAJE_HOGAR, 255, 255, 255)
            Traveling = False
        
        Case eMessages.CancelGoHome
            Dim Cancelled As Boolean
            Cancelled = Reader.ReadBool
            
            If (Cancelled) Then
                Call ShowConsoleMsg(MENSAJE_HOGAR_CANCEL, 255, 0, 0, True)
            End If
            
            Traveling = False
    End Select

End Sub

Private Sub HandleConnected()

    Call HandleLogin
    
End Sub

''
' Handles the Logged message.

Private Sub HandleLogged()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 20/06/2014 (D'Artagnan)
'20/06/2014: D'Artagnan - Clean account session token.
'***************************************************
    EngineRun = True
    Nombres = GameConfig.Extras.NameStyle
    bRain = False
    
    ' Clean account session token.
    Acc_Data.Acc_Token = vbNullString

    'Set connected state
    Call SetConnected
    
    If bShowTutorial Then frmTutorial.Show vbModeless
    
End Sub

''
' Handles the RemoveDialogs message.

Private Sub HandleRemoveDialogs()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    Call Dialogos.RemoveAllDialogs

End Sub

''
' Handles the RemoveCharDialog message.

Private Sub HandleRemoveCharDialog()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    Call Dialogos.RemoveDialog(Reader.ReadInt16())

End Sub

''
' Handles the NavigateChange message.

Private Sub HandleNavigateChange()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    'change IsSaling flag with new value
    charlist(UserCharIndex).IsSailing = Reader.ReadBool()
End Sub

''
' Handles the Disconnect message.

Private Sub HandleDisconnect()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    Call ResetAllInfo(True)
End Sub

''
' Handles the CommerceEnd message.

Private Sub HandleCommerceEnd()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    Set InvComUsu = Nothing
    Set InvComNpc = Nothing
    
    'Hide form
    Unload frmComerciar
    
    'Reset vars
    Comerciando = False

End Sub

''
' Handles the BankEnd message.

Private Sub HandleBankEnd()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    Set InvBanco(0) = Nothing
    Set InvBanco(1) = Nothing
    
    Unload frmBancoObj
    Comerciando = False

End Sub

''
' Handles the CommerceInit message.

Private Sub HandleCommerceInit()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

  
    Dim I As Long

    Set InvComUsu = New clsGraphicalInventory
    Set InvComNpc = New clsGraphicalInventory
    
    Load frmComerciar
    
    ' Initialize commerce inventories
    Call InvComUsu.Initialize(frmComerciar.picInvUser, Inventario.MaxObjs, , , , , , , , , True, _
                              eMoveType.Inventory, InvComNpc)
    Call InvComNpc.Initialize(frmComerciar.picInvNpc, MAX_NPC_INVENTORY_SLOTS, , , , , , , , , True, _
                              eMoveType.Target, InvComUsu)

    Set InvComUsu.dropInventory = InvComNpc
    Set InvComNpc.dropInventory = InvComUsu

    'Fill user inventory
    For I = 1 To MAX_INVENTORY_SLOTS
        If Inventario.ObjIndex(I) <> 0 Then
            With Inventario
                Call InvComUsu.SetItem(I, .ObjIndex(I), _
                    .Amount(I), .Equipped(I), .GrhIndex(I), _
                    .OBJType(I), .MaxHit(I), .MinHit(I), .MaxDef(I), .MinDef(I), _
                    .Valor(I), .ItemName(I), 0, .CanUse(I))
            End With
        End If
    Next I
    
    ' Fill Npc inventory
    For I = 1 To MAX_NPC_INVENTORY_SLOTS
        If NPCInventory(I).ObjIndex <> 0 Then
            With NPCInventory(I)
                Call InvComNpc.SetItem(I, .ObjIndex, _
                    .Amount, 0, .GrhIndex, _
                    .OBJType, .MaxHit, .MinHit, .MaxDef, .MinDef, _
                    .Valor, .Name, 0, .CanUse)
            End With
        End If
    Next I
    
    'Set state and show form
    Comerciando = True
    frmComerciar.Show , frmMain

End Sub

''
' Handles the BankInit message.

Private Sub HandleBankInit()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    Dim I As Long
    Dim BankGold As Long
    
    Set InvBanco(0) = New clsGraphicalInventory
    Set InvBanco(1) = New clsGraphicalInventory
    
    BankGold = Reader.ReadInt32
    
    Load frmBancoObj
    
    Call InvBanco(0).Initialize(frmBancoObj.PicBancoInv, MAX_BANCOINVENTORY_SLOTS, , , , , , , , , True, _
                                eMoveType.Target, InvBanco(1))

    Call InvBanco(1).Initialize(frmBancoObj.PicInv, Inventario.MaxObjs, , , , , , , , , True, _
                                eMoveType.Inventory, InvBanco(0))
    
    For I = 1 To Inventario.MaxObjs
        With Inventario
            Call InvBanco(1).SetItem(I, .ObjIndex(I), _
                .Amount(I), .Equipped(I), .GrhIndex(I), _
                .OBJType(I), .MaxHit(I), .MinHit(I), .MaxDef(I), .MinDef(I), _
                .Valor(I), .ItemName(I), 0, .CanUse(I))
        End With
    Next I
    
    For I = 1 To MAX_BANCOINVENTORY_SLOTS
        With UserBancoInventory(I)
            Call InvBanco(0).SetItem(I, .ObjIndex, _
                .Amount, .Equipped, .GrhIndex, _
                .OBJType, .MaxHit, .MinHit, .MaxDef, .MinDef, _
                .Valor, .Name, 0, .CanUse)
        End With
    Next I
    
    frmBancoObj.lblUserGld.Caption = BankGold
    
    'Set state and show form
    Comerciando = True
        
    frmBancoObj.Show , frmMain
    
End Sub

Private Sub HandleShowGuildCreate()
    Dim Faccion As eGuildAlignment
    
    Faccion = Reader.ReadInt8()
    Call frmGuildCreateStep1.SetAlignment(Faccion)
    frmGuildCreateStep1.Show , frmMain
End Sub

Private Sub HandleShowGuildForm()
    Call frmGuildMain.ShowFull
End Sub

Private Sub HandleGuildCreated()
    frmGuildCreateStep3.Show , frmMain
    frmGuildBank.LblUserGold = UserGLD
End Sub

''
' Handles the UserCommerceInit message.

Private Sub HandleUserCommerceInit()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

  
    Dim I As Long

    TradingUserName = Reader.ReadString8
    

    Set InvComUsu = New clsGraphicalInventory
    Set InvOfferComUsu(0) = New clsGraphicalInventory
    Set InvOfferComUsu(1) = New clsGraphicalInventory
    
    Load frmComerciarUsu
    
    ' Initialize commerce inventories
    Call InvComUsu.Initialize(frmComerciarUsu.picInvComercio, Inventario.MaxObjs)
    Call InvOfferComUsu(0).Initialize(frmComerciarUsu.picInvOfertaProp, INV_OFFER_SLOTS)
    
    Call InvOfferComUsu(1).Initialize(frmComerciarUsu.picInvOfertaOtro, INV_OFFER_SLOTS)

    'Fill user inventory
    For I = 1 To MAX_INVENTORY_SLOTS
        If Inventario.ObjIndex(I) <> 0 Then
            With Inventario
                Call InvComUsu.SetItem(I, .ObjIndex(I), _
                    .Amount(I), .Equipped(I), .GrhIndex(I), _
                    .OBJType(I), .MaxHit(I), .MinHit(I), .MaxDef(I), .MinDef(I), _
                    .Valor(I), .ItemName(I), 0, .CanUse(I))
            End With
        End If
    Next I

    frmComerciarUsu.LblUserGold.Caption = UserGLD
    frmComerciarUsu.lblUserOfferedGold.Caption = 0
    frmComerciarUsu.lblOtherUserGold.Caption = 0
    
    ' Inventarios de oro

    'Set state and show form
    Comerciando = True
    Call frmComerciarUsu.Show(vbModeless, frmMain)
    
End Sub

''
' Handles the UserCommerceEnd message.

Private Sub HandleUserCommerceEnd()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    Set InvComUsu = Nothing
    Set InvOfferComUsu(0) = Nothing
    Set InvOfferComUsu(1) = Nothing
    
    'Destroy the form and reset the state
    Unload frmComerciarUsu
    Comerciando = False

End Sub

''
' Handles the UserOfferConfirm message.
Private Sub HandleUserOfferConfirm()

    Call frmComerciarUsu.UserConfirmedOffer

End Sub

''
' Handles the ShowBlacksmithForm message.

Private Sub HandleShowCraftForm()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    Call frmCraft.Initialize
    frmCraft.Show , frmMain
    frmCraft.Visible = True
    MirandoHerreria = True
 

End Sub

''
' Handles the UpdateSta message.

Private Sub HandleUpdateSta()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    'Get data and update form
    UserMinSTA = Reader.ReadInt16()
    
    frmMain.lblEnergia = UserMinSTA & "/" & UserMaxSTA
    Call frmMain.UpdateStaBar

End Sub

''
' Handles the UpdateMana message.

Private Sub HandleUpdateMana()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    'Get data and update form
    UserMinMAN = Reader.ReadInt16()
    
    frmMain.lblMana = UserMinMAN & "/" & UserMaxMAN
    Call frmMain.UpdateManBar

End Sub

''
' Handles the UpdateHP message.

Private Sub HandleUpdateHP()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    'Get data and update form
    UserMinHP = Reader.ReadInt16()
    
    frmMain.lblVida = UserMinHP & "/" & UserMaxHP
    Call frmMain.UpdateHPBar
    
    'Is the user alive??
    If UserMinHP = 0 Then
        
        
        ' Start the grayscale transition
        If UserEstado <> 1 Then
            'Call GoGhost
        End If
        
        UserEstado = 1
        
        If frmMain.TrainingMacro Then Call frmMain.DesactivarMacroHechizos
        If frmMain.macrotrabajo Then Call frmMain.DesactivarMacroTrabajo
    Else
        
        ' Disable the grayscale
        'If DeathEffect.State <> DeathEffectStates.ToAlive And UserEstado <> 0 Then
            'Call GoAlive
        'End If
    
        UserEstado = 0
    End If

End Sub

Private Sub HandleUpdateChallenge()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    Dim Amount_gold As Long
    Dim Maxim_dead As Byte
    
    Dim Event_time As Byte
    Dim Time_start As Byte
    Dim Event_map As Byte
    
    Dim Invisibility As Byte
    Dim Resucitar As Byte
    Dim Elementary As Byte
    
    
    'Get data and update form
    Amount_gold = Reader.ReadInt32()
    Maxim_dead = Reader.ReadInt8()
    
    Event_time = Reader.ReadInt8()
    Time_start = Reader.ReadInt8()
    Event_map = Reader.ReadInt8()
    
    Invisibility = Reader.ReadInt8()
    Resucitar = Reader.ReadInt8()
    Elementary = Reader.ReadInt8()

End Sub

' TODO: Nightw - This functtion is part of the old Challenge system that never got launched
' This should be removed.
'paquete para todos los usuarios del desafio, se muestran en pantalla
Private Sub HandleUpdateChallengeStat()

    ' Uncomment this
    'With Desafio
    '    .Arena = Reader.ReadInt8()
    '    .Oro = Reader.ReadInt32()
    '
    '    .Maxim_dead = Reader.ReadInt8()
    '    .Event_time = Reader.ReadInt8()
    '
    '    .Time_start = Reader.ReadInt8()
    '    .Event_map = Reader.ReadInt8()
    '
    '    .Invisibility = Reader.ReadInt8()
    '    .Resucitar = Reader.ReadInt8()
    '    .Elementary = Reader.ReadInt8()
    '
    '    .Team = Reader.ReadInt8()
    '    .Point(1) = Reader.ReadInt8()
    '    .Point(2) = Reader.ReadInt8()
    'End With

End Sub

''
' Handles the UpdateGold message.

Private Sub HandleUpdateGold()
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 10/01/2016
'Last Modified By: Anagrama
'- 08/14/07: Tavo - Added GldLbl color variation depending on User Gold and Level
'- 09/21/10: C4b3z0n - Modified color change of gold ONLY if the player's level is greater than 12 (NOT newbie).
'- 10/01/2016: Anagrama - Now the color changes even if the user is newbie.
'***************************************************

    'Get data and update form
    UserGLD = Reader.ReadInt32()
   
    frmMain.GldLbl.ForeColor = &HFFFF& 'Yellow
    frmMain.GldLbl.Caption = UserGLD
    
    If PlayerData.Guild.IdGuild > 0 Then
        frmGuildBank.LblUserGold.Caption = UserGLD
    End If

End Sub

''
' Handles the UpdateBankGold message.

Private Sub HandleUpdateBankGold()
'***************************************************
'Autor: ZaMa
'Last Modification: 14/12/2009
'
'***************************************************

    frmBancoObj.lblUserGld.Caption = Reader.ReadInt32
    If frmBancoObj.Visible = False Then
        Unload frmBancoObj
    End If

End Sub

''
' Handles the UpdateExp message.

Private Sub HandleUpdateExp()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    'Get data and update form
    UserExp = Reader.ReadInt32()

    Call ShowLevelCompletionPerc

    Dim bWidth As Byte

    If UserPasarNivel > 0 Then _
        bWidth = (((UserExp / 100) / (UserPasarNivel / 100)) * 138)
        
    frmMain.shpExp.Width = 138 - bWidth
    frmMain.shpExp.Left = 611 + (138 - frmMain.shpExp.Width)
    
    frmMain.shpExp.Visible = (bWidth <> 138)

End Sub

''
' Handles the UpdateStrenghtAndDexterity message.

Private Sub HandleUpdateStrenghtAndDexterity()
'***************************************************
'Author: Budi
'Last Modification: 11/26/09
'***************************************************

    'Get data and update form
    UserFuerza = Reader.ReadInt8
    UserAgilidad = Reader.ReadInt8
    frmMain.lblStrg.Caption = UserFuerza
    frmMain.lblDext.Caption = UserAgilidad
    
    frmMain.tmrBlink.Enabled = False
    
    frmMain.lblStrg.ForeColor = getStrenghtColor(UserFuerza)
    frmMain.lblDext.ForeColor = getDexterityColor(UserAgilidad)

End Sub

' Handles the UpdateStrenghtAndDexterity message.

Private Sub HandleUpdateStrenght()
'***************************************************
'Author: Budi
'Last Modification: 11/26/09
'***************************************************

    'Get data and update form
    UserFuerza = Reader.ReadInt8
    frmMain.lblStrg.Caption = UserFuerza
    frmMain.lblStrg.ForeColor = getStrenghtColor(UserFuerza)
    
    If frmMain.tmrBlink.Enabled Then
        frmMain.lblDext.ForeColor = getDexterityColor(UserAgilidad)
    End If
    
    frmMain.tmrBlink.Enabled = False

End Sub

' Handles the HandleUpdateCharacterInfo message.

Private Sub HandleUpdateCharacterInfo()
'***************************************************
'Author: Budi
'Last Modification: 11/26/09
'***************************************************
    Dim Class As Byte, Race As Byte, Gender As Byte

    Class = Reader.ReadInt
    Race = Reader.ReadInt
    Gender = Reader.ReadInt
    
    
    If Class <> PlayerData.Class Then
        PlayerData.Class = Class
        Call frmMain.ShowStoreButton(PlayerData.CurrentMap.CraftingStoreAllowed)
    End If
    
    If Race <> PlayerData.Race Then
        PlayerData.Race = Race
    End If
    
    If Gender <> PlayerData.Gender Then
        PlayerData.Gender = Gender
    End If
    
End Sub

' Handles the UpdateStrenghtAndDexterity message.

Private Sub HandleUpdateDexterity()
'***************************************************
'Author: Budi
'Last Modification: 11/26/09
'***************************************************

    'Get data and update form
    UserAgilidad = Reader.ReadInt8
    frmMain.lblDext.Caption = UserAgilidad
    frmMain.lblDext.ForeColor = getDexterityColor(UserAgilidad)
    
    If frmMain.tmrBlink.Enabled Then
        frmMain.lblStrg.ForeColor = getStrenghtColor(UserFuerza)
    End If
    
    frmMain.tmrBlink.Enabled = False
    
End Sub

''
' Handles the ChangeMap message.
Private Sub HandleChangeMap()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    Dim CanOpenCraftStore As Boolean
    
    PlayerData.CurrentMap.Number = Reader.ReadInt16()
    Call Reader.ReadInt16 'TODO: Once on-the-fly editor is implemented check for map version before loading....
    Reader.ReadInt8 ' TODO: Wolftein REMOVE
    PlayerData.CurrentMap.CraftingStoreAllowed = Reader.ReadBool
    
    UserMap = PlayerData.CurrentMap.Number
    
    
#If EnableSecurity Then
    Call InitMI
#End If
    
    If (Not SwitchMap(PlayerData.CurrentMap.Number)) Then
        'no encontramos el mapa en el hd
        MsgBox "Error en los mapas, algún archivo ha sido modificado o esta dañado."
        
        Call CloseClient
        Exit Sub
    End If
    
    If UserCharIndex > 0 Then
        g_Last_OffsetX = charlist(UserCharIndex).MoveOffsetX
        g_Last_OffsetY = charlist(UserCharIndex).MoveOffsetY
    End If

    ' Is it raining?
    If bRain Then
        ' Is it an underground map? If yes, can't rain there!
        If bLluvia(PlayerData.CurrentMap.Number) = 0 Then
            Call Engine_Audio.DisableSound(RainBufferIndex)

            frmMain.IsPlaying = PlayLoop.plNone
            'Call Mod_Weather.ResetWeather
        Else
            'Call Mod_Weather.SetWeather(RAIN)
        End If
    End If
    
    Call frmMain.ShowStoreButton(PlayerData.CurrentMap.CraftingStoreAllowed)
    
End Sub

''
' Handles the PosUpdate message.

Private Sub HandlePosUpdate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    'Remove char from old position
    If MapData(UserPos.X, UserPos.Y).CharIndex = UserCharIndex Then
        MapData(UserPos.X, UserPos.Y).CharIndex = 0
    End If
    
    'Set new pos
    UserPos.X = Reader.ReadInt8()
    UserPos.Y = Reader.ReadInt8()
    
    'Set char
    MapData(UserPos.X, UserPos.Y).CharIndex = UserCharIndex
    charlist(UserCharIndex).Pos = UserPos

    With charlist(UserCharIndex)
        Call UpdateNodeSceneChar(UserCharIndex)
    
        Call Aurora_Scene.Update(.Node)
    End With
    
    Call Engine_Audio.UpdateEmitter(charlist(UserCharIndex).SoundSource, UserPos.X, UserPos.Y)
    
    'Are we under a roof?
    bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)
                
    'Update pos label
    frmMain.Coord.Caption = PlayerData.CurrentMap.Number & " X: " & UserPos.X & " Y: " & UserPos.Y

End Sub

''
' Handles the ChatOverHead message.

Private Sub HandleChatOverHead()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    Dim chat As String
    Dim CharIndex As Integer
    Dim r As Byte
    Dim g As Byte
    Dim b As Byte
    
    chat = Trim$(Reader.ReadString8())
    CharIndex = Reader.ReadInt16()
    
    r = Reader.ReadInt8()
    g = Reader.ReadInt8()
    b = Reader.ReadInt8()

    Call Dialogos.CreateDialog(chat, CharIndex, RGBA(r, g, b, 255))

End Sub

''
' Handles the ChatPersonalizado message.

Private Sub HandleChatPersonalizado()
'***************************************************
'Author: Juan Dalmasso (CHOTS)
'Last Modification: 11/06/2011
'***************************************************

    Dim chat As String
    Dim CharIndex As Integer
    Dim TIPO As Byte

    
    chat = Trim$(Reader.ReadString8())
    CharIndex = Reader.ReadInt16()
    
    TIPO = Reader.ReadInt8()

    Call Dialogos.CreateDialog(chat, CharIndex, RGBA(ColoresDialogos(TIPO).r, ColoresDialogos(TIPO).g, ColoresDialogos(TIPO).b, 255))
   
End Sub
Private Sub HandleConsoleMessage()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/05/11
'D'Artagnan: Agrego la división de consolas
'***************************************************

    Dim chat As String
    Dim FontIndex As Integer
    Dim MessageType As eMessageType
    Dim str As String
    Dim r As Byte
    Dim g As Byte
    Dim b As Byte
    
    chat = Reader.ReadString8()
    FontIndex = Reader.ReadInt8()
    MessageType = Reader.ReadInt8()

    If InStr(1, chat, "~") Then
        str = ReadField(2, chat, 126)
            If Val(str) > 255 Then
                r = 255
            Else
                r = Val(str)
            End If
            
            str = ReadField(3, chat, 126)
            If Val(str) > 255 Then
                g = 255
            Else
                g = Val(str)
            End If
            
            str = ReadField(4, chat, 126)
            If Val(str) > 255 Then
                b = 255
            Else
                b = Val(str)
            End If
        
        Call AddtoRichTextBox(frmMain.RecTxt(0), Left$(chat, InStr(1, chat, "~") - 1), r, g, b, Val(ReadField(5, chat, 126)) <> 0, Val(ReadField(6, chat, 126)) <> 0, , MessageType)
    Else
        If FontIndex = FontTypeNames.FONTTYPE_PARTY Then 'CHOTS | Mensajes personalizados de Party
            Call AddtoRichTextBox(frmMain.RecTxt(0), chat, ColoresDialogos(3).r, ColoresDialogos(3).g, ColoresDialogos(3).b, False, False, , MessageType)
        Else
            With FontTypes(FontIndex)
                Call AddtoRichTextBox(frmMain.RecTxt(0), chat, .red, .green, .blue, .bold, .italic, , MessageType)
            End With
        End If
        
        ' Para no perder el foco cuando chatea por party
        If FontIndex = FontTypeNames.FONTTYPE_PARTY Then
            If MirandoParty Then frmParty.SendTxt.SetFocus
        End If
    End If

End Sub

Private Sub HandleConsoleFormattedMessage()
    Dim I As Integer
    Dim MessageId  As Integer
    Dim ParametersCount As Integer
    Dim ParameterType As eMessageParameterType
    
    Dim Result As FormattedMessage
    
    MessageId = Reader.ReadInt
    ParametersCount = Reader.ReadInt
    
    Call MessageManager.Prepare(MessageId)
    
    For I = 1 To ParametersCount
        ParameterType = Reader.ReadInt
        Select Case ParameterType
            Case eMessageParameterType.IsNumber
                Call MessageManager.AddParameterAsNumber(Reader.ReadInt())
            Case eMessageParameterType.IsText
                Call MessageManager.AddParameterAsText(Reader.ReadString8())
        End Select
    Next I
    
    Result = MessageManager.Format()
    
    With FontTypes(Result.FontType)
        Call ShowConsoleMsg(Result.Message, .red, .green, .blue, .bold, .italic, Result.MessageType)
    End With
End Sub

Private Sub HandleShowFormattedMessageBox()
    Dim I As Integer
    Dim MessageId  As Integer
    Dim ParametersCount As Integer
    Dim ParameterType As eMessageParameterType
    
    Dim Result As FormattedMessage
    
    MessageId = Reader.ReadInt
    ParametersCount = Reader.ReadInt
    
    Call MessageManager.Prepare(MessageId)
    
    For I = 1 To ParametersCount
        ParameterType = Reader.ReadInt
        Select Case ParameterType
            Case eMessageParameterType.IsNumber
                Call MessageManager.AddParameterAsNumber(Reader.ReadInt())
            Case eMessageParameterType.IsText
                Call MessageManager.AddParameterAsText(Reader.ReadString8())
        End Select
    Next I
    
    Result = MessageManager.Format()
    
    Call frmMessageBox.ShowMessage(Result.Message)
    
End Sub


''
' Handles the GuildChat message.

Private Sub HandleGuildChat()

    Dim chat As String
    Dim IsMOTD As Boolean
    Dim str As String
    Dim r As Byte
    Dim g As Byte
    Dim b As Byte
    
    chat = Reader.ReadString8()
    IsMOTD = Reader.ReadBool()

    If Not DialogosClanes.Activo Then
        If InStr(1, chat, "~") Then
            str = ReadField(2, chat, 126)
            If Val(str) > 255 Then
                r = 255
            Else
                r = Val(str)
            End If
            
            str = ReadField(3, chat, 126)
            If Val(str) > 255 Then
                g = 255
            Else
                g = Val(str)
            End If
            
            str = ReadField(4, chat, 126)
            If Val(str) > 255 Then
                b = 255
            Else
                b = Val(str)
            End If
            
            If IsMOTD = True Then
                Call AddtoRichTextBox(frmMain.RecTxt(0), Left$(chat, InStr(1, chat, "~") - 1), r, g, b, Val(ReadField(5, chat, 126)) <> 0, Val(ReadField(6, chat, 126)) <> 0, , eMessageType.m_MOTD)
            Else
                Call AddtoRichTextBox(frmMain.RecTxt(0), Left$(chat, InStr(1, chat, "~") - 1), r, g, b, Val(ReadField(5, chat, 126)) <> 0, Val(ReadField(6, chat, 126)) <> 0, , eMessageType.Guild)
            End If
        Else
            With FontTypes(FontTypeNames.FONTTYPE_GUILDMSG)
                If IsMOTD = True Then
                    Call AddtoRichTextBox(frmMain.RecTxt(0), chat, ColoresDialogos(2).r, ColoresDialogos(2).g, ColoresDialogos(2).b, .bold, .italic, , eMessageType.m_MOTD)
                Else
                    Call AddtoRichTextBox(frmMain.RecTxt(0), chat, ColoresDialogos(2).r, ColoresDialogos(2).g, ColoresDialogos(2).b, .bold, .italic, , eMessageType.Guild)
                End If
            End With
        End If
    Else
        Call DialogosClanes.PushBackText(ReadField(1, chat, 126))
    End If
    
End Sub

''
' Handles the ConsoleMessage message.

Private Sub HandleCommerceChat()
'***************************************************
'Author: ZaMa
'Last Modification: 03/12/2009
'
'***************************************************

    Dim chat As String
    Dim FontIndex As Integer
    Dim str As String
    Dim r As Byte
    Dim g As Byte
    Dim b As Byte
    
    chat = Reader.ReadString8()
    FontIndex = Reader.ReadInt8()

    If InStr(1, chat, "~") Then
        str = ReadField(2, chat, 126)
            If Val(str) > 255 Then
                r = 255
            Else
                r = Val(str)
            End If
            
            str = ReadField(3, chat, 126)
            If Val(str) > 255 Then
                g = 255
            Else
                g = Val(str)
            End If
            
            str = ReadField(4, chat, 126)
            If Val(str) > 255 Then
                b = 255
            Else
                b = Val(str)
            End If
            
        Call AddtoRichTextBox(frmComerciarUsu.CommerceConsole, Left$(chat, InStr(1, chat, "~") - 1), r, g, b, Val(ReadField(5, chat, 126)) <> 0, Val(ReadField(6, chat, 126)) <> 0)
    Else
        With FontTypes(FontIndex)
            Call AddtoRichTextBox(frmComerciarUsu.CommerceConsole, chat, .red, .green, .blue, .bold, .italic)
        End With
    End If

End Sub

''
' Handles the ShowMessageBox message.

Private Sub HandleShowMessageBox()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    Dim tmpString As String
    tmpString = Reader.ReadString8()

    frmMensaje.msg.Caption = tmpString
    frmMensaje.Show
    
End Sub

''
' Handles the UserIndexInServer message.

Private Sub HandleUserIndexInServer()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    UserIndex = Reader.ReadInt16()

End Sub

''
' Handles the UserCharIndexInServer message.

Private Sub HandleUserCharIndexInServer()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

     UserCharIndex = Reader.ReadInt16()
     UserPos = charlist(UserCharIndex).Pos

    'Are we under a roof?
     bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)

    frmMain.Coord.Caption = PlayerData.CurrentMap.Number & " X: " & UserPos.X & " Y: " & UserPos.Y
End Sub

''
' Handles the CharacterCreate message.

Private Sub HandleCharacterCreate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 17/09/14
'
'17/09/14: D'Artagnan - Hostile and merchant attributes.
'***************************************************


    Dim CharIndex As Integer
    Dim Body As Integer
    Dim Head As Integer
    Dim Heading As E_Heading
    Dim X As Byte
    Dim Y As Byte
    Dim weapon As Integer
    Dim shield As Integer
    Dim helmet As Integer
    Dim privs As Integer
    Dim NickColor As Byte
    Dim NpcNumber As Integer
    Dim OverheadIcon As Integer
    Dim WithTransparency As Boolean
    Dim Fx As Integer
    Dim Loops As Integer
    Dim Nombre As String
    Dim IsHostile As Boolean
    Dim IsMerchant As Boolean
    Dim IsSailing As Boolean
    Dim Alignment As eCharacterAlignment
    
    CharIndex = Reader.ReadInt16()
    Body = Reader.ReadInt16()
    Head = Reader.ReadInt16()
    Heading = Reader.ReadInt8()
    X = Reader.ReadInt8()
    Y = Reader.ReadInt8()
    weapon = Reader.ReadInt16()
    shield = Reader.ReadInt16()
    helmet = Reader.ReadInt16()
    

    
    Fx = Reader.ReadInt16()
    Loops = Reader.ReadInt16()
    Nombre = Reader.ReadString8()
    NickColor = Reader.ReadInt8()
    Alignment = Reader.ReadInt8()
    privs = Reader.ReadInt8()
    
    IsHostile = Reader.ReadBool()
    IsMerchant = Reader.ReadBool()
    IsSailing = Reader.ReadBool()
    NpcNumber = Reader.ReadInt16()
    OverheadIcon = Reader.ReadInt
    WithTransparency = Reader.ReadBool

    With charlist(CharIndex)
        Call SetCharacterFx(CharIndex, Fx, Loops)
        
        .Nombre = Nombre
        NickColor = NickColor
        
        If (NickColor And eNickColor.ieCriminal) <> 0 Then
            .criminal = 1
        Else
            .criminal = 0
        End If
        
        .Atacable = (NickColor And eNickColor.ieAtacable) <> 0
        
        If privs <> 0 Then
            'If the player belongs to a council AND is an admin, only whos as an admin
            If (privs And PlayerType.ChaosCouncil) <> 0 And (privs And PlayerType.User) = 0 Then
                privs = privs Xor PlayerType.ChaosCouncil
            End If
            
            If (privs And PlayerType.RoyalCouncil) <> 0 And (privs And PlayerType.User) = 0 Then
                privs = privs Xor PlayerType.RoyalCouncil
            End If
            
            'If the player is a RM, ignore other flags
            If privs And PlayerType.RoleMaster Then
                privs = PlayerType.RoleMaster
            End If
            
            'Log2 of the bit flags sent by the server gives our numbers ^^
            .priv = Log(privs) / Log(2)
        Else
            .priv = 0
        End If
        
        .bHostile = IsHostile
        .bMerchant = IsMerchant
        .IsSailing = IsSailing
        .NpcNumber = NpcNumber
        .OverheadIcon = OverheadIcon
        .Alignment = Alignment
        .UseInvisibilityAlpha = WithTransparency
    End With
        
    Call MakeChar(CharIndex, Body, Head, Heading, X, Y, weapon, shield, helmet)

End Sub

Private Sub HandleCharacterChangeNick()
'***************************************************
'Author: Budi
'Last Modification: 07/23/09
'
'***************************************************


    Dim CharIndex As Integer
    Dim tmpString As String
    CharIndex = Reader.ReadInt16
    tmpString = Reader.ReadString8()

    charlist(CharIndex).Nombre = tmpString

End Sub

''
' Handles the CharacterRemove message.

Private Sub HandleCharacterRemove()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    Dim CharIndex As Integer
    
    CharIndex = Reader.ReadInt16()

    Call EraseChar(CharIndex)

End Sub

''
' Handles the CharacterMove message.

Private Sub HandleCharacterMove()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************


    Dim CharIndex As Integer
    Dim X As Byte
    Dim Y As Byte
    Dim Warped As Boolean
    
    CharIndex = Reader.ReadInt16()
    X = Reader.ReadInt8()
    Y = Reader.ReadInt8()
    Warped = Reader.ReadBool
    
    With charlist(CharIndex)
        If .Fx(0).FxIndex >= 40 And .Fx(0).FxIndex <= 49 Then   'If it's meditating, we remove the FX
            .Fx(0).FxIndex = 0
        End If
        
        ' Play steps sounds if the user is not an admin of any kind
        If (Not Warped) Then
            
            If .priv <> 1 And .priv <> 2 And .priv <> 3 And .priv <> 5 And .priv <> 25 Then
                Call DoPasosFx(CharIndex)
            End If
            
            Call MoveCharbyPos(CharIndex, X, Y)
    
        Else
            
            Call MoveCharByTelep(CharIndex, X, Y)
        End If
        
    End With
    

End Sub

''
' Handles the ForceCharMove message.

Private Sub HandleForceCharMove()

    Dim Direccion As Byte
    
    Direccion = Reader.ReadInt8()

    Call MoveCharbyHead(UserCharIndex, Direccion)
    Call MoveScreen(Direccion)

End Sub

''
' Handles the CharacterChange message.
 
Private Sub HandleCharacterChange()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 21/09/2010 - C4b3z0n
'25/08/2009: ZaMa - Changed a variable used incorrectly.
'28/09/2016: Loopzer - Try to read all data then execute
'***************************************************

    Dim CharIndex As Integer
    'Dim tempint As Integer
    Dim Body As Integer
    
    Dim headIndex As Integer
    Dim FxIndex As Integer
    Dim FxLoops As Integer
    Dim Heading As Byte
    Dim WeaponAnim As Integer
    Dim ShieldAnim As Integer
    Dim CascoAnim As Integer
    Dim IsSailing As Integer
    Dim IsDead As Byte
    Dim OverheadIcon As Integer
    Dim Alignment As Byte
    
    CharIndex = Reader.ReadInt16()
    Body = Reader.ReadInt16()
    headIndex = Reader.ReadInt16()
    Heading = Reader.ReadInt8()
    WeaponAnim = Reader.ReadInt16()
    ShieldAnim = Reader.ReadInt16()
    CascoAnim = Reader.ReadInt16()
    FxIndex = Reader.ReadInt16()
    FxLoops = Reader.ReadInt16()
    IsSailing = Reader.ReadBool()
    IsDead = Reader.ReadBool()
    OverheadIcon = Reader.ReadInt
    Alignment = Reader.ReadInt8
    
    With charlist(CharIndex)
        
        If Body < LBound(BodyData()) Or Body > UBound(BodyData()) Then
            Body = Not Body
            InitGrh .Body.Walk(1), Body
            InitGrh .Body.Walk(2), Body
            InitGrh .Body.Walk(3), Body
            InitGrh .Body.Walk(4), Body
        Else
            .Body = BodyData(Body)
            .iBody = Body
        End If
        
        If headIndex < LBound(HeadData()) Or headIndex > UBound(HeadData()) Then
            .Head = HeadData(0)
            .iHead = 0
        Else
            .Head = HeadData(headIndex)
            .iHead = headIndex
        End If
        
        .muerto = IsDead
        
        .Heading = Heading
        
        .OverheadIcon = OverheadIcon
        .Alignment = Alignment
        
        If WeaponAnim <> 0 Then .Arma = WeaponAnimData(WeaponAnim)
        
        If ShieldAnim <> 0 Then .Escudo = ShieldAnimData(ShieldAnim)
        
        If CascoAnim <> 0 Then .Casco = CascoAnimData(CascoAnim)
        

        .IsSailing = IsSailing
        'Call SetCharacterFx(CharIndex, FxIndex, FxLoops)

        Call UpdateNodeSceneChar(CharIndex)
        Call Aurora_Scene.Update(.Node)

    End With

End Sub

''
' Handles the ObjectCreate message.

Private Sub HandleObjectCreate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    Dim X As Byte
    Dim Y As Byte
    Dim Luminous As Byte
    Dim LightSize As Integer
    Dim GrhIndex As Integer
    Dim OffsetX As Integer
    Dim OffsetY As Integer
    Dim CanBeTransparent As Boolean
    Dim OBJType As Byte
    Dim ObjMetadata As Long
    
    X = Reader.ReadInt8()
    Y = Reader.ReadInt8()
    GrhIndex = Reader.ReadInt16()
    Luminous = Reader.ReadInt8()
    OffsetX = Reader.ReadInt16
    OffsetY = Reader.ReadInt16
    LightSize = Reader.ReadInt16()
    CanBeTransparent = Reader.ReadBool()
    OBJType = Reader.ReadInt8()
    ObjMetadata = Reader.ReadInt
    
    With MapData(X, Y)

        .ObjGrh.GrhIndex = GrhIndex
        .OBJInfo.Luminous = Luminous
        .OBJInfo.LightOffsetX = OffsetX
        .OBJInfo.LightOffsetY = OffsetY
        .OBJInfo.LightSize = LightSize
        .OBJInfo.CanBeTransparent = CanBeTransparent
        
        
        Call InitGrh(.ObjGrh, .ObjGrh.GrhIndex)
        Call InitGrhDepth(.ObjGrh, 3, X, Y, 1)
        
        If (GrhIndex <> 0) Then
            With GrhData(GrhIndex)
                Call UpdateNodeScene(MapData(X, Y).OBJInfo.ObjNode, -1, X, Y, 5, .TileWidth, .TileHeight)
            End With
  
            Call Aurora_Scene.Insert(.OBJInfo.ObjNode)
        End If

        Select Case OBJType
            Case eObjType.otPuertas
                ' Should we block the tile?
                MapData(X, Y).Blocked = ObjMetadata
                MapData(X - 1, Y).Blocked = ObjMetadata
            
                Set .OBJInfo.SoundSource = Engine_Audio.CreateEmitter(X, Y)
            Case eObjType.otFogata
                Set .OBJInfo.SoundSource = Engine_Audio.CreateEmitter(X, Y)
                   
                Call Engine_Audio.PlayEffect("fuego.wav", .OBJInfo.SoundSource, True)
        End Select
        
    End With
    
End Sub

''
' Handles the ObjectDelete message.

Private Sub HandleObjectDelete()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    Dim X As Byte
    Dim Y As Byte

    X = Reader.ReadInt8()
    Y = Reader.ReadInt8()

    With MapData(X, Y)
        Call Engine_Audio.DeleteEmitter(.OBJInfo.SoundSource, True)
    
        Call Aurora_Scene.Remove(.OBJInfo.ObjNode)
        
        .ObjGrh.GrhIndex = 0
    End With

End Sub

''
' Handles the ObjectUpdate message.

Private Sub HandleObjectUpdate()
    Dim X As Byte
    Dim Y As Byte
    Dim GrhIndex As Integer
    Dim OBJType As Long
    Dim ObjMetadata As Long
    
    X = Reader.ReadInt8()
    Y = Reader.ReadInt8()
    GrhIndex = Reader.ReadInt()
    OBJType = Reader.ReadInt
    ObjMetadata = Reader.ReadInt
   
    With MapData(X, Y)

        .ObjGrh.GrhIndex = GrhIndex

        Call InitGrh(.ObjGrh, .ObjGrh.GrhIndex)
        Call InitGrhDepth(.ObjGrh, 3, X, Y, 1)

        Select Case OBJType
            Case eObjType.otPuertas
                ' Should we block the tile?
                MapData(X, Y).Blocked = ObjMetadata
                MapData(X - 1, Y).Blocked = ObjMetadata
                
            Case eObjType.otFogata
                If (.OBJInfo.SoundSource Is Nothing) Then
                    Set .OBJInfo.SoundSource = Engine_Audio.CreateEmitter(X, Y)
                End If
                Call Engine_Audio.PlayEffect("fuego.wav", .OBJInfo.SoundSource, True)
        End Select
        
    End With
    
End Sub

''
' Handles the BlockPosition message.

Private Sub HandleBlockPosition()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    Dim X As Byte
    Dim Y As Byte
    
    X = Reader.ReadInt8()
    Y = Reader.ReadInt8()
    
    If Reader.ReadBool() Then
        MapData(X, Y).Blocked = 1
    Else
        MapData(X, Y).Blocked = 0
    End If
    
End Sub

''
' Handles the PlayMusic message.

Private Sub HandlePlayMusic()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'14/07/2016: Anagrama - Ahora recibe una lista de musica para el mapa.
'***************************************************

    Dim Playlist() As Long
    Call Reader.ReadSafeArrayInt32(Playlist)    ' TODO: Remove this
    
    Call Engine_Audio.PlayMusic(Playlist(1) & ".mp3", True)
    
End Sub

''
' Handles the PlayEffect message.

Private Sub HandlePlayEffect()
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 08/14/07
'Last Modified by: Rapsodius
'Added support for 3D Sounds.
'***************************************************

    Dim wave As Integer
    Dim srcX As Byte
    Dim srcY As Byte
    Dim Entity As Long
    
    wave = Reader.ReadInt16()
    srcX = Reader.ReadInt8()
    srcY = Reader.ReadInt8()
    Entity = Reader.ReadInt

    If (Entity > 0) Then
        Call Engine_Audio.PlayEffect(wave & ".wav", charlist(Entity).SoundSource)
    Else
        Call Engine_Audio.PlayEffect(CStr(wave) & ".wav", Engine_Audio.CreateEmitter(srcX, srcY))
    End If

End Sub

''
' Handles the PauseToggle message.

Private Sub HandlePauseToggle()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    pausa = Not pausa

End Sub

''
' Handles the RainToggle message.

Private Sub HandleRainToggle()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    If Not InMapBounds(UserPos.X, UserPos.Y) Then Exit Sub
    
    bTecho = (MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 4)
    If bRain Then
        If bLluvia(PlayerData.CurrentMap.Number) Then
            'Stop playing the rain sound
            Call Engine_Audio.DisableSound(RainBufferIndex)

            If bTecho Then
                Call Engine_Audio.PlayEffect("lluviainend.wav")
            Else
                Call Engine_Audio.PlayEffect("lluviaoutend.wav")
            End If
            
            frmMain.IsPlaying = PlayLoop.plNone
        End If
    End If
    
    bRain = Not bRain
End Sub

''
' Handles the CreateFX message.

Private Sub HandleCreateFX()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    Dim CharIndex As Integer
    Dim Fx As Integer
    Dim Loops As Integer
    Dim Slot As Byte
    
    CharIndex = Reader.ReadInt16()
    Fx = Reader.ReadInt16()
    Loops = Reader.ReadInt16()
    Slot = Reader.ReadInt8()
    
    ' Reads if the effect received is a glow effect or not
    If Reader.ReadBool Then
        Call SetCharacterGlowFx(CharIndex, Fx, Loops)
    Else
        Call SetCharacterFx(CharIndex, Fx, Loops, Slot)
    End If

End Sub

''
' Handles the UpdateUserStats message.

Private Sub HandleUpdateUserStats()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
On Error GoTo ErrHandler:

    UserMaxHP = Reader.ReadInt16()
    UserMinHP = Reader.ReadInt16()
    UserMaxMAN = Reader.ReadInt16()
    UserMinMAN = Reader.ReadInt16()
    UserMaxSTA = Reader.ReadInt16()
    UserMinSTA = Reader.ReadInt16()
    UserGLD = Reader.ReadInt32()
    UserLvl = Reader.ReadInt8()
    UserPasarNivel = Reader.ReadInt32()
    UserExp = Reader.ReadInt32()
    IsMaxLevel = Reader.ReadBool()
    UserMasteryPoints = Reader.ReadInt16()

    Call ShowLevelCompletionPerc
    
    frmMain.GldLbl.Caption = UserGLD
    
    If PlayerData.Guild.IdGuild > 0 Then
        frmGuildBank.LblUserGold = UserGLD
    End If
    
    If IsMaxLevel Then
        frmMain.lblLvl.FontSize = 10
        frmMain.lblLvl.Top = 58
        frmMain.lblLvl.Caption = "MAX"
        frmMain.lblMasteryPoints.Visible = True
        frmMain.lblMasteryPoints.Caption = UserMasteryPoints
    Else
        frmMain.lblMasteryPoints.Visible = False
        frmMain.lblLvl.Top = 60
        frmMain.lblLvl.Caption = UserLvl
        frmMain.lblLvl.FontSize = 12
        
    End If
   
    Dim bWidth As Integer
    
    
    'Exp
    
    '***************************
    If UserPasarNivel > 0 Then _
        bWidth = (((UserExp / 100) / (UserPasarNivel / 100)) * 138)
        
    If bWidth > 138 Then bWidth = 138
        
    frmMain.shpExp.Width = 138 - bWidth
    frmMain.shpExp.Left = 611 + (138 - frmMain.shpExp.Width)
    
    frmMain.shpExp.Visible = (bWidth <> 138)
    
    
    'Stats
    frmMain.lblMana = UserMinMAN & "/" & UserMaxMAN
    frmMain.lblVida = UserMinHP & "/" & UserMaxHP
    frmMain.lblEnergia = UserMinSTA & "/" & UserMaxSTA
    
    Call frmMain.UpdateManBar
    Call frmMain.UpdateHPBar
    Call frmMain.UpdateStaBar
    
    If UserMinHP = 0 Then
        
        
         ' Disable the grayscale
        If UserEstado <> 1 Then
            'Call GoGhost
        End If
        
        UserEstado = 1
        
        If frmMain.TrainingMacro Then Call frmMain.DesactivarMacroHechizos
        If frmMain.macrotrabajo Then Call frmMain.DesactivarMacroTrabajo
    Else
    
         ' Disable the grayscale
        'If DeathEffect.State <> DeathEffectStates.ToAlive And UserEstado <> 0 Then
            'Call GoAlive
        'End If
    
        UserEstado = 0
    End If
    
    frmMain.GldLbl.ForeColor = &HFFFF& 'Yellow
    Exit Sub
    
    
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleUpdateUserStats de Protocol.bas")

End Sub

''
' Handles the WorkRequestTarget message.

Private Sub HandleWorkRequestTarget()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    UsingSkill = Reader.ReadInt8()

    frmMain.MousePointer = 2
    
    Select Case UsingSkill
        Case Magia
            Call AddtoRichTextBox(frmMain.RecTxt(0), MENSAJE_TRABAJO_MAGIA, 100, 100, 120, 0, 0, , eMessageType.Trabajo)
        Case Pesca
            Call AddtoRichTextBox(frmMain.RecTxt(0), MENSAJE_TRABAJO_PESCA, 100, 100, 120, 0, 0, , eMessageType.Trabajo)
        Case Robar
            Call AddtoRichTextBox(frmMain.RecTxt(0), MENSAJE_TRABAJO_ROBAR, 100, 100, 120, 0, 0, , eMessageType.Trabajo)
        Case Talar
            Call AddtoRichTextBox(frmMain.RecTxt(0), MENSAJE_TRABAJO_TALAR, 100, 100, 120, 0, 0, , eMessageType.Trabajo)
        Case Mineria
            Call AddtoRichTextBox(frmMain.RecTxt(0), MENSAJE_TRABAJO_MINERIA, 100, 100, 120, 0, 0, , eMessageType.Trabajo)
        Case FundirMetal
            Call AddtoRichTextBox(frmMain.RecTxt(0), MENSAJE_TRABAJO_FUNDIRMETAL, 100, 100, 120, 0, 0, , eMessageType.Trabajo)
        Case Proyectiles
            Call AddtoRichTextBox(frmMain.RecTxt(0), MENSAJE_TRABAJO_PROYECTILES, 100, 100, 120, 0, 0, , eMessageType.Trabajo)
    End Select

End Sub

''
' Handles the ChangeInventorySlot message.

Private Sub HandleChangeInventorySlot()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 03/02/2015
'03/02/2015: D'Artagnan - Defense modifiers.
'***************************************************

    Dim Slot As Byte
    Dim ObjIndex As Integer
    Dim oldObjIndex As Integer
    Dim Name As String
    Dim OldName As String
    Dim Amount As Integer
    Dim Equipped As Boolean
    Dim GrhIndex As Integer
    Dim OBJType As Byte
    Dim MaxHit As Integer
    Dim MinHit As Integer
    Dim MaxDef As Integer
    Dim MinDef As Integer
    Dim nPrimaryArmourMinDef As Integer
    Dim nPrimaryArmourMaxDef As Integer
    Dim CanUse As Boolean
    Dim SalePrice As Long
    
    Dim tempInventory As clsGraphicalInventory

    Slot = Reader.ReadInt8()
    ObjIndex = Reader.ReadInt16()
    Amount = Reader.ReadInt16()
    Equipped = Reader.ReadBool()
    CanUse = Reader.ReadBool()
    
    ' Limpio todo
    If Slot = 0 Then
        Inventario.ClearAllSlots
        frmMain.lblWeapon = "0/0"
        UserWeaponEqpSlot = 0
        frmMain.lblArmor = "0/0"
        UserArmourEqpSlot = 0
        UserSecArmourEqpSlot = 0
        frmMain.lblShielder = "0/0"
        UserHelmEqpSlot = 0
        frmMain.lblHelm = "0/0"
        UserShieldEqpSlot = 0
                    
    ' Actualiza
    Else
        If ObjIndex > 0 Then
            
            With GameMetadata.Objs(ObjIndex)
                Name = .Name
                GrhIndex = .GrhIndex
                OBJType = .OBJType
                MaxHit = .MaxHit
                MinHit = .MinHit
                MaxDef = .MaxDef
                MinDef = .MinDef
                
                If OBJType = eObjType.otArmadura And IsSecondaryArmour(ObjIndex) And _
                   UserArmourEqpSlot > 0 Then
                    Dim ArmourObjIndex As Integer
                    
                    ArmourObjIndex = Inventario.ObjIndex(UserArmourEqpSlot)
                    nPrimaryArmourMinDef = GameMetadata.Objs(ArmourObjIndex).MinDef
                    nPrimaryArmourMaxDef = GameMetadata.Objs(ArmourObjIndex).MaxDef
                End If
            End With
        End If
        
        If Equipped Then
            Select Case OBJType
                Case eObjType.otWeapon, eObjType.otTool
                    frmMain.lblWeapon = MinHit & "/" & MaxHit
                    UserWeaponEqpSlot = Slot
                Case eObjType.otArmadura
                    If IsSecondaryArmour(ObjIndex) Then
                        ' Apply defense modifier.
                        MaxDef = nPrimaryArmourMaxDef * MOD_DEF_SEG_JERARQUIA
                        MinDef = nPrimaryArmourMinDef * MOD_DEF_SEG_JERARQUIA
                        
                        UserSecArmourEqpSlot = Slot
                    Else
                        If UserSecArmourEqpSlot > 0 Then
                            MaxDef = MaxDef * MOD_DEF_SEG_JERARQUIA
                            MinDef = MinDef * MOD_DEF_SEG_JERARQUIA
                        End If
                        UserArmourEqpSlot = Slot
                    End If
                    frmMain.lblArmor = MinDef & "/" & MaxDef
                Case eObjType.otEscudo
                    frmMain.lblShielder = MinDef & "/" & MaxDef
                    UserHelmEqpSlot = Slot
                Case eObjType.otCasco
                    frmMain.lblHelm = MinDef & "/" & MaxDef
                    UserShieldEqpSlot = Slot
            End Select
        Else
            Select Case Slot
                Case UserWeaponEqpSlot
                    frmMain.lblWeapon = "0/0"
                    UserWeaponEqpSlot = 0
                Case UserArmourEqpSlot
                    frmMain.lblArmor = "0/0"
                    UserArmourEqpSlot = 0
                Case UserSecArmourEqpSlot
                    frmMain.lblArmor = nPrimaryArmourMinDef & "/" & nPrimaryArmourMaxDef
                    UserSecArmourEqpSlot = 0
                Case UserHelmEqpSlot
                    frmMain.lblShielder = "0/0"
                    UserHelmEqpSlot = 0
                Case UserShieldEqpSlot
                    frmMain.lblHelm = "0/0"
                    UserShieldEqpSlot = 0
            End Select
        End If
        
        OldName = Inventario.ItemName(Slot)
        oldObjIndex = Inventario.ObjIndex(Slot)
        SalePrice = Inventario.Valor(Slot)
        
        Call Inventario.SetItem(Slot, ObjIndex, Amount, Equipped, GrhIndex, OBJType, _
            MaxHit, MinHit, MaxDef, MinDef, SalePrice, Name, 0, CanUse)
    
        If Not (InvBanco(1) Is Nothing) Then
            Set tempInventory = InvBanco(1)
        ElseIf Not (InvComUsu Is Nothing) Then
            Set tempInventory = InvComUsu
        ElseIf Not (AccBank(1) Is Nothing) Then
            Set tempInventory = AccBank(1)
        ElseIf frmGuildMain.ActiveForm Is frmGuildQuestActive Then
            Set tempInventory = frmGuildQuestActive.UserInventory
        ElseIf frmGuildMain.ActiveForm Is frmGuildBank Then
            Set tempInventory = frmGuildBank.GMemberInv
        End If

        If Not (tempInventory Is Nothing) Then
            Call tempInventory.SetItem(Slot, ObjIndex, Amount, Equipped, GrhIndex, OBJType, _
                MaxHit, MinHit, MaxDef, MinDef, 0, Name, 0, CanUse)
        End If

    End If

End Sub

' Handles the AddSlots message.
Private Sub HandleAddSlots()
'***************************************************
'Author: Budi
'Last Modification: 12/01/09
'
'***************************************************

    MaxInventorySlots = Reader.ReadInt8

End Sub

' Handles the StopWorking message.
Private Sub HandleStopWorking()
'***************************************************
'Author: Budi
'Last Modification: 12/01/09
'
'***************************************************

    With FontTypes(FontTypeNames.FONTTYPE_INFO)
        Call ShowConsoleMsg("¡Has terminado de trabajar!", .red, .green, .blue, .bold, .italic)
    End With
    
    If frmMain.macrotrabajo.Enabled Then Call frmMain.DesactivarMacroTrabajo

End Sub

' Handles the CancelOfferItem message.

Private Sub HandleCancelOfferItem()
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 05/03/10
'
'***************************************************

    Dim Slot As Byte
    Dim Amount As Long

    Slot = Reader.ReadInt8
    
    With InvOfferComUsu(0)
        Amount = .Amount(Slot)
        
        ' No tiene sentido que se quiten 0 unidades
        If Amount <> 0 Then
            ' Actualizo el inventario general
            Call frmComerciarUsu.UpdateInvCom(.ObjIndex(Slot), Amount)
            
            ' Borro el item
            Call .SetItem(Slot, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", 0, True)
        End If
    End With
    
    ' Si era el único ítem de la oferta, no puede confirmarla
    If Not frmComerciarUsu.HasAnyItem(InvOfferComUsu(0)) And _
        frmComerciarUsu.GetOtherPlayerOfferedGold() <= 0 Then Call frmComerciarUsu.HabilitarConfirmar(False)
    
    With FontTypes(FontTypeNames.FONTTYPE_INFO)
        Call frmComerciarUsu.PrintCommerceMsg("¡No puedes comerciar ese objeto!", FontTypeNames.FONTTYPE_INFO)
    End With

End Sub

''
' Handles the ChangeBankSlot message.

Private Sub HandleChangeBankSlot()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    Dim Slot As Byte
    Slot = Reader.ReadInt8()
    
    Dim ObjIndex As Integer
    ObjIndex = Reader.ReadInt16()
    
    Dim Amount As Integer
    Amount = Reader.ReadInt16()
    
    Dim CanUse As Boolean
    CanUse = Reader.ReadBool()
    
    If Slot = 0 Then
        ReDim UserBancoInventory(1 To MAX_BANCOINVENTORY_SLOTS) As Inventory
    Else
        Dim sMain As String
        sMain = "OBJ" & CStr(ObjIndex)
        
        With UserBancoInventory(Slot)
        
            If ObjIndex > 0 Then
                .ObjIndex = ObjIndex
                .Name = GameMetadata.Objs(ObjIndex).Name
                .Amount = Amount
                .GrhIndex = GameMetadata.Objs(ObjIndex).GrhIndex
                .OBJType = GameMetadata.Objs(ObjIndex).OBJType
                .MaxHit = GameMetadata.Objs(ObjIndex).MaxHit
                .MinHit = GameMetadata.Objs(ObjIndex).MinHit
                .MaxDef = GameMetadata.Objs(ObjIndex).MaxDef
                .MinDef = GameMetadata.Objs(ObjIndex).MinDef
                .Valor = GameMetadata.Objs(ObjIndex).Valor
                .CanUse = CanUse
            Else
                .ObjIndex = 0
                .Name = "Nada"
                .Amount = 0
                .GrhIndex = 0
                .OBJType = 0
                .MaxHit = 0
                .MinHit = 0
                .MaxDef = 0
                .MinDef = 0
                .Valor = 0
                .CanUse = True
            End If
            
            If Comerciando Then
                Call InvBanco(0).SetItem(Slot, .ObjIndex, .Amount, _
                    .Equipped, .GrhIndex, .OBJType, .MaxHit, _
                    .MinHit, .MaxDef, .MinDef, .Valor, .Name, 0, .CanUse)
            End If
        End With
    End If

End Sub

''
' Handles the ChangeSpellSlot message.

Private Sub HandleChangeSpellSlot()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    Dim Slot As Byte
    Dim Interval As Double
    Dim tmpString As String
    
    Slot = Reader.ReadInt()
    UserHechizos(Slot) = Reader.ReadInt()
    
    tmpString = Reader.ReadString8()
    Interval = Reader.ReadInt()
    
    'If Interval < PlayerData.Intervals.PlayerCastSpell Then Interval = PlayerData.Intervals.PlayerCastSpell

    If Slot <= frmMain.hlst.ListCount Then
        Call frmMain.hlst.SetItem(Slot, tmpString, Interval)
    Else
        Call frmMain.hlst.Add(tmpString, Interval)
    End If

End Sub

''
' Handles the Attributes message.

Private Sub HandleAtributes()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    Dim I As Long
    
    For I = 1 To NUMATRIBUTES
        UserAtributos(I) = Reader.ReadInt8()
    Next I

End Sub

''
' Handles the craftable recipes message.

Private Sub HandleCraftableRecipes()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim GroupsQty As Byte
    Dim I As Integer, J As Integer, k As Integer
    
    GroupsQty = Reader.ReadInt
    
    With PlayerData
        .CraftingRecipeGroupsQty = GroupsQty
        
        If GroupsQty <= 0 Then
            Erase PlayerData.CraftingRecipeGroups
            Exit Sub
        End If
    
        ' Recipe groups
        ReDim PlayerData.CraftingRecipeGroups(1 To GroupsQty)
        For I = 1 To .CraftingRecipeGroupsQty
            .CraftingRecipeGroups(I).TabTitle = Reader.ReadString16
            .CraftingRecipeGroups(I).TabImage = Reader.ReadString16
            .CraftingRecipeGroups(I).ProfessionType = Reader.ReadInt
            .CraftingRecipeGroups(I).RecipesQty = Reader.ReadInt
            
            If .CraftingRecipeGroups(I).RecipesQty > 0 Then
                ReDim .CraftingRecipeGroups(I).Recipes(1 To .CraftingRecipeGroups(I).RecipesQty)
                
                ' Recipes for the group
                For J = 1 To .CraftingRecipeGroups(I).RecipesQty
                    .CraftingRecipeGroups(I).Recipes(J).RecipeIndex = Reader.ReadInt
                    .CraftingRecipeGroups(I).Recipes(J).ObjNumber = Reader.ReadInt

                    ' Materials
                    .CraftingRecipeGroups(I).Recipes(J).MaterialsQty = Reader.ReadInt
                    If .CraftingRecipeGroups(I).Recipes(J).MaterialsQty > 0 Then
                        ReDim .CraftingRecipeGroups(I).Recipes(J).Materials(1 To .CraftingRecipeGroups(I).Recipes(J).MaterialsQty)
                        
                        For k = 1 To .CraftingRecipeGroups(I).Recipes(J).MaterialsQty
                            .CraftingRecipeGroups(I).Recipes(J).Materials(k).ObjNumber = Reader.ReadInt
                            .CraftingRecipeGroups(I).Recipes(J).Materials(k).Amount = Reader.ReadInt
                        Next k
                    End If
                Next J
            End If
            
        Next I

    End With

End Sub


''
' Handles the RestOK message.

Private Sub HandleRestOK()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    UserDescansar = Not UserDescansar

End Sub

''
' Handles the ErrorMessage message.

Private Sub HandleErrorMessage()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    Dim bCloseConnection As Boolean
    
    Call frmMessageBox.ShowMessage(Reader.ReadString8())

    bCloseConnection = Reader.ReadBool()
    
End Sub

''
' Handles the Blind message.

Private Sub HandleBlind()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    UserCiego = True
  
End Sub

''
' Handles the Dumb message.

Private Sub HandleDumb()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    UserEstupido = True

End Sub

''
' Handles the ShowSignal message.

Private Sub HandleShowSignal()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    Dim tmp As String
    tmp = Reader.ReadString8()
    
    Call InitCartel(tmp, Reader.ReadInt16())

End Sub

''
' Handles the ChangeNPCInventorySlot message.

Private Sub HandleChangeNPCInventorySlot()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    Dim Slot As Byte
    Slot = Reader.ReadInt8()
    
    Dim Amount As Integer
    Amount = Reader.ReadInt16()
    
    Dim Price As Single
    Price = Reader.ReadReal32()
    
    Dim ObjIndex As Integer
    ObjIndex = Reader.ReadInt16()
    
    Dim CanUse As Boolean
    CanUse = Reader.ReadBool()
    
    ' Clear
    If Slot = 0 Then
        ReDim NPCInventory(1 To MAX_NPC_INVENTORY_SLOTS) As NpCinV
    Else
        Dim sMain As String
        sMain = "OBJ" & CStr(ObjIndex)
        
        With NPCInventory(Slot)
            If ObjIndex > 0 Then
                .ObjIndex = ObjIndex
                .Name = GameMetadata.Objs(ObjIndex).Name
                .Amount = Amount
                .Valor = Price
                .GrhIndex = GameMetadata.Objs(ObjIndex).GrhIndex
                .OBJType = GameMetadata.Objs(ObjIndex).OBJType
                .MaxHit = GameMetadata.Objs(ObjIndex).MaxHit
                .MinHit = GameMetadata.Objs(ObjIndex).MinHit
                .MaxDef = GameMetadata.Objs(ObjIndex).MaxDef
                .MinDef = GameMetadata.Objs(ObjIndex).MinDef
                .CanUse = CanUse
            Else
                .ObjIndex = 0
                .Name = "Nada"
                .Amount = 0
                .Valor = 0
                .GrhIndex = 0
                .OBJType = 0
                .MaxHit = 0
                .MinHit = 0
                .MaxDef = 0
                .MinDef = 0
                .CanUse = True
            End If
            
            
            If Comerciando Then
                ' Compraron la totalidad de un item, o vendieron un item que el npc no tenia
                If .ObjIndex <> InvComNpc.ObjIndex(Slot) Then
                    Call InvComNpc.SetItem(Slot, .ObjIndex, _
                        .Amount, 0, .GrhIndex, _
                        .OBJType, .MaxHit, .MinHit, .MaxDef, .MinDef, _
                        .Valor, .Name, 0, .CanUse)
                ' Compraron o vendieron cierta cantidad (no su totalidad)
                ElseIf .Amount <> InvComNpc.Amount(Slot) Then
                    Call InvComNpc.ChangeSlotItemAmount(Slot, .Amount)
                End If
            End If
        End With
    End If

End Sub

''
' Handles the SetUserSalePrice message.
Private Sub HandleSetUserSalePrice()

    Dim Slot As Byte
    Dim SalePrice As Long
    
    Slot = Reader.ReadInt8()
    SalePrice = Reader.ReadInt32()
    
    Call Inventario.ChangeSlotItemPrice(Slot, SalePrice)

End Sub

''
' Handles the UpdateHungerAndThirst message.

Private Sub HandleUpdateHungerAndThirst()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    UserMaxAGU = Reader.ReadInt8()
    UserMinAGU = Reader.ReadInt8()
    UserMaxHAM = Reader.ReadInt8()
    UserMinHAM = Reader.ReadInt8()
    frmMain.lblHambre = UserMinHAM & "/" & UserMaxHAM
    frmMain.lblSed = UserMinAGU & "/" & UserMaxAGU

    Dim bWidth As Integer
    bWidth = (((UserMinHAM / 100) / (UserMaxHAM / 100)) * 88)
    
    frmMain.shpHambre.Width = 88 - bWidth
    frmMain.shpHambre.Left = 573 + (88 - frmMain.shpHambre.Width)
    
    frmMain.shpHambre.Visible = (bWidth <> 88)
    
    '*********************************
    
    bWidth = (((UserMinAGU / 100) / (UserMaxAGU / 100)) * 88)
    
    frmMain.shpSed.Width = 88 - bWidth
    frmMain.shpSed.Left = 573 + (88 - frmMain.shpSed.Width)
    
    frmMain.shpSed.Visible = (bWidth <> 88)

End Sub

''
' Handles the MiniStats message.

Private Sub HandleMiniStats()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    Dim ClassType As Byte

    With UserEstadisticas
        .CiudadanosMatados = Reader.ReadInt32()
        .CriminalesMatados = Reader.ReadInt32()
        .UsuariosMatados = Reader.ReadInt32()
        .NpcsMatados = Reader.ReadInt32()
        
        ClassType = Reader.ReadInt8()
        .PenaCarcel = Reader.ReadInt32()
        
        .Clase = ListaClases(ClassType)
    End With

End Sub

''
' Handles the LevelUp message.

Private Sub HandleLevelUp()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    SkillPoints = Reader.ReadInt16()
    
    Call frmMain.LightSkillStar(True)

End Sub

''
' Handles the AddForumMessage message.

Private Sub HandleAddForumMessage()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    Dim ForumType As eForumMsgType
    Dim Title As String
    Dim Message As String
    Dim Author As String
    
    ForumType = Reader.ReadInt8
    
    Title = Reader.ReadString8()
    Author = Reader.ReadString8()
    Message = Reader.ReadString8()

    If Not frmForo.ForoLimpio Then
        clsForos.ClearForums
        frmForo.ForoLimpio = True
    End If

    Call clsForos.AddPost(ForumAlignment(ForumType), Title, Author, Message, EsAnuncio(ForumType))

End Sub

''
' Handles the ShowForumForm message.

Private Sub HandleShowForumForm()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    frmForo.Privilegios = Reader.ReadInt8
    frmForo.CanPostSticky = Reader.ReadInt8
    
    If Not MirandoForo Then
        frmForo.Show , frmMain
    End If

End Sub

''
' Handles the SetInvisible message.

Private Sub HandleSetInvisible()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    Dim CharIndex As Integer

    CharIndex = Reader.ReadInt16()
    
#If EnableSecurity Then
    If Reader.ReadBool() Then
        Call MI(CualMI).SetInvisible(CharIndex)
    Else
        Call MI(CualMI).ResetInvisible(CharIndex)
    End If
#Else
    charlist(CharIndex).invisible = Reader.ReadBool()
#End If

    charlist(CharIndex).UseInvisibilityAlpha = Reader.ReadBool()

End Sub

''
' Handles the MeditateToggle message.

Private Sub HandleMeditateToggle()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    UserMeditar = Not UserMeditar

    WaitInput = False
    
End Sub

''
' Handles the BlindNoMore message.

Private Sub HandleBlindNoMore()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    UserCiego = False

End Sub

''
' Handles the DumbNoMore message.

Private Sub HandleDumbNoMore()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    UserEstupido = False

End Sub

''
' Handles the SendSkills message.

Private Sub HandleSendSkills()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 28/04/2015
'11/19/09: Pato - Now the server send the percentage of progress of the skills.
'28/04/2015: D'Artagnan - Read both natural and assigned skills.
'***************************************************

    Dim I As Long

    For I = 1 To NUMSKILLS
        UserSkills(I).Natural = Reader.ReadInt8()
        UserSkills(I).Assigned = Reader.ReadInt8()
        PorcentajeSkills(I) = Reader.ReadInt8()
    Next I
    
    If OrigenSkills = eOrigenSkills.ieAsignacion Then
        frmSkills3.Show , frmMain
        
    ElseIf OrigenSkills = eOrigenSkills.ieEstadisticas Then
        frmEstadisticas.Show , frmMain
    End If

End Sub

''
' Handles the TrainerCreatureList message.

Private Sub HandleTrainerCreatureList()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************


    Dim creatures() As String
    Dim I As Long
    
    creatures = Split(Reader.ReadString8(), SEPARATOR)

    For I = 0 To UBound(creatures())
        Call frmEntrenador.lstCriaturas.AddItem(creatures(I))
    Next I
    frmEntrenador.Show , frmMain

End Sub

''
' Handles the OfferDetails message.

Private Sub HandleOfferDetails()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    Call frmUserRequest.recievePeticion(Reader.ReadString8())

End Sub

''
' Handles the ParalizeOK message.

Private Sub HandleParalizeOK()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    UserParalizado = Not UserParalizado

End Sub

''
' Handles the ShowUserRequest message.

Private Sub HandleShowUserRequest()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    Call frmUserRequest.recievePeticion(Reader.ReadString8())
    Call frmUserRequest.Show(vbModeless, frmMain)

End Sub

''
' Handles the TradeOK message.

Private Sub HandleTradeOK()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    If frmComerciar.Visible Then
        Dim I As Long
        
        'Update user inventory
        For I = 1 To MAX_INVENTORY_SLOTS
            ' Agrego o quito un item en su totalidad
            If Inventario.ObjIndex(I) <> InvComUsu.ObjIndex(I) Then
                With Inventario
                    Call InvComUsu.SetItem(I, .ObjIndex(I), _
                        .Amount(I), .Equipped(I), .GrhIndex(I), _
                        .OBJType(I), .MaxHit(I), .MinHit(I), .MaxDef(I), .MinDef(I), _
                        .Valor(I), .ItemName(I), 0, .CanUse(I))
                End With
            ' Vendio o compro cierta cantidad de un item que ya tenia
            ElseIf Inventario.Amount(I) <> InvComUsu.Amount(I) Then
                Call InvComUsu.ChangeSlotItemAmount(I, Inventario.Amount(I))
            End If
        Next I
        
        ' Fill Npc inventory
        For I = 1 To 20
            ' Compraron la totalidad de un item, o vendieron un item que el npc no tenia
            If NPCInventory(I).ObjIndex <> InvComNpc.ObjIndex(I) Then
                With NPCInventory(I)
                    Call InvComNpc.SetItem(I, .ObjIndex, _
                        .Amount, 0, .GrhIndex, _
                        .OBJType, .MaxHit, .MinHit, .MaxDef, .MinDef, _
                        .Valor, .Name, 0, .CanUse)
                End With
            ' Compraron o vendieron cierta cantidad (no su totalidad)
            ElseIf NPCInventory(I).Amount <> InvComNpc.Amount(I) Then
                Call InvComNpc.ChangeSlotItemAmount(I, NPCInventory(I).Amount)
            End If
        Next I
    
    End If

End Sub

''
' Handles the BankOK message.

Private Sub HandleBankOK()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    Dim I As Long
    
    If frmBancoObj.Visible Then
        For I = 1 To Inventario.MaxObjs
            With Inventario
                Call InvBanco(1).SetItem(I, .ObjIndex(I), .Amount(I), _
                    .Equipped(I), .GrhIndex(I), .OBJType(I), .MaxHit(I), _
                    .MinHit(I), .MaxDef(I), .MinDef(I), .Valor(I), .ItemName(I), 0, .CanUse(I))
            End With
        Next I

        frmBancoObj.NoPuedeMover = False
    End If

End Sub

Public Sub HandleChangeUserTradeGold()
    Dim Amount As Long
    Amount = Reader.ReadInt32
    
    Call frmComerciarUsu.SetOtherPlayerOfferedGold(Amount)
    Call frmComerciarUsu.PrintCommerceMsg(TradingUserName & " ha modificado su oferta.", FontTypeNames.FONTTYPE_VENENO)

End Sub

''
' Handles the ChangeUserTradeSlot message.

Private Sub HandleChangeUserTradeSlot()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    Dim OfferSlot As Byte
    OfferSlot = Reader.ReadInt8
    
    Dim ObjIndex As Integer
    ObjIndex = Reader.ReadInt16
    
    Dim Amount As Long
    Amount = Reader.ReadInt32
    
    Dim CanUse As Boolean
    CanUse = Reader.ReadBool
    
    Dim tempInventory As clsGraphicalInventory
    
    Call frmComerciarUsu.PrintCommerceMsg(TradingUserName & " ha modificado su oferta.", FontTypeNames.FONTTYPE_VENENO)
    
    
    Set tempInventory = InvOfferComUsu(1)
    
    ' The user removed the last item from this slot, so the ObjIndex will be 0
    If ObjIndex <= 0 Then
        If tempInventory.SelectedItem = OfferSlot Then
            tempInventory.DeselectItem
        End If
        Call tempInventory.SetItem(OfferSlot, ObjIndex, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", 0, True)
    Else
        With GameMetadata.Objs(ObjIndex)
            Call tempInventory.SetItem(OfferSlot, ObjIndex, Amount, 0, _
                .GrhIndex, _
                .OBJType, _
                .MaxHit, _
                .MinHit, _
                .MaxDef, _
                .MinDef, _
                0, _
                .Name, 0, CanUse)
        End With
    End If
    
End Sub

''
' Handles the SpawnList message.

Private Sub HandleSpawnList()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    Dim creatureList() As String
    Dim I As Long
    
    creatureList = Split(Reader.ReadString8(), SEPARATOR)
    
    For I = 0 To UBound(creatureList())
        Call frmSpawnList.lstCriaturas.AddItem(creatureList(I))
    Next I
    frmSpawnList.Show , frmMain

End Sub

''
' Handles the ShowSOSForm message.

Private Sub HandleShowSOSForm()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    Dim sosList() As String
    Dim I As Long
    
    sosList = Split(Reader.ReadString8(), SEPARATOR)
    
    For I = 0 To UBound(sosList())
        Call frmMSG.List1.AddItem(sosList(I))
    Next I
    
    frmMSG.Show , frmMain

End Sub

''
' Handles the ShowDenounces message.

Private Sub HandleShowDenounces()
'***************************************************
'Author: ZaMa
'Last Modification: 14/11/2010
'
'***************************************************

    Dim DenounceList() As String
    Dim DenounceIndex As Long
    
    DenounceList = Split(Reader.ReadString8(), SEPARATOR)

    With FontTypes(FontTypeNames.FONTTYPE_GUILDMSG)
        For DenounceIndex = 0 To UBound(DenounceList())
            Call AddtoRichTextBox(frmMain.RecTxt(0), DenounceList(DenounceIndex), .red, .green, .blue, .bold, .italic, True, eMessageType.Info)
        Next DenounceIndex
    End With

End Sub

''
' Handles the ShowSOSForm message.

Private Sub HandleShowPartyForm()
'***************************************************
'Author: Budi
'Last Modification: 11/26/09
'
'***************************************************


    
    Dim MembersStr As String
    Dim TotalExp As Long
    Dim Members() As String
    Dim I As Long

    EsPartyLeader = CBool(Reader.ReadInt8())
    MembersStr = Reader.ReadString8()
    TotalExp = Reader.ReadInt32

    Members = Split(MembersStr, SEPARATOR)
    
    frmParty.lstMembers.Clear
    For I = 0 To UBound(Members())
        Call frmParty.lstMembers.AddItem(Members(I))
    Next I
    
    frmParty.lblTotalExp.Caption = TotalExp
    frmParty.Show , frmMain
    


End Sub



''
' Handles the ShowMOTDEditionForm message.

Private Sub HandleShowMOTDEditionForm()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'*************************************Su**************

    frmCambiaMotd.txtMotd.Text = Reader.ReadString8()
    frmCambiaMotd.Show , frmMain

End Sub

''
' Handles the ShowGMPanelForm message.

Private Sub HandleShowGMPanelForm()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    frmPanelGm.Show vbModeless, frmMain

End Sub

''
' Handles the UserNameList message.

Private Sub HandleUserNameList()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    Dim userList() As String
    Dim I As Long
    
    userList = Split(Reader.ReadString8(), SEPARATOR)
    
    If frmPanelGm.Visible Then
        frmPanelGm.cboListaUsus.Clear
        For I = 0 To UBound(userList())
            Call frmPanelGm.cboListaUsus.AddItem(userList(I))
        Next I
        If frmPanelGm.cboListaUsus.ListCount > 0 Then frmPanelGm.cboListaUsus.ListIndex = 0
    End If

End Sub

''
' Handles the Pong message.

Private Sub HandlePong()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    
    currentPingTime = timeGetTime - Reader.ReadInt32()

End Sub

''
' Handles the UpdateTag message.

Private Sub HandleUpdateTagAndStatus()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    Dim CharIndex As Integer
    Dim NickColor As Byte
    Dim UserTag As String
    
    CharIndex = Reader.ReadInt16()
    NickColor = Reader.ReadInt8()
    UserTag = Reader.ReadString8()

    'Update char status adn tag!
    With charlist(CharIndex)
        If (NickColor And eNickColor.ieCriminal) <> 0 Then
            .criminal = 1
        Else
            .criminal = 0
        End If
        
        .Atacable = (NickColor And eNickColor.ieAtacable) <> 0
        
        .Nombre = UserTag
    End With

End Sub

Private Sub HandleUpdateUserSpellCooldown()

    Dim Slot As Integer
    Dim Cooldown As Long
    Dim UpdateAll As Boolean
    Dim I As Long
    
    UpdateAll = Reader.ReadBool()
    
    If UpdateAll = True Then
        For I = 1 To MAXHECHI
            If (Not frmMain.hlst.List(I - 1) = "(None)") Then
                Cooldown = Reader.ReadInt()
                Call frmMain.hlst.SetItem(I, frmMain.hlst.List(I - 1), Cooldown)
            End If
        Next I
    Else
        Slot = Reader.ReadInt()
        Cooldown = Reader.ReadInt()
        
        Call frmMain.hlst.SetItem(Slot, frmMain.hlst.List(Slot - 1), Cooldown)
    End If

    
End Sub

''
' Handles the RecordList message.

Private Sub HandleRecordList()
'***************************************************
'Author: Amraphen
'Last Modification: 29/11/2010
'
'***************************************************

    Dim NumRecords As Byte
    Dim I As Long
    
    NumRecords = Reader.ReadInt8
    
    'Se limpia el ListBox y se agregan los usuarios
    frmPanelGm.lstUsers.Clear
    For I = 1 To NumRecords
        frmPanelGm.lstUsers.AddItem Reader.ReadString8
    Next I

End Sub

''
' Handles the RecordDetails message.

Private Sub HandleRecordDetails()
'***************************************************
'Author: Amraphen
'Last Modification: 29/11/2010
'
'***************************************************

    Dim tmpStr As String

    With frmPanelGm
        .txtCreador.Text = Reader.ReadString8
        .txtDescrip.Text = Reader.ReadString8
        
        'Status del pj
        If Reader.ReadBool Then
            .lblEstado.ForeColor = vbGreen
            .lblEstado.Caption = "ONLINE"
        Else
            .lblEstado.ForeColor = vbRed
            .lblEstado.Caption = "OFFLINE"
        End If
        
        'IP del personaje
        tmpStr = Reader.ReadString8
        If LenB(tmpStr) Then
            .txtIP.Text = tmpStr
        Else
            .txtIP.Text = "Usuario offline"
        End If
        
        'Tiempo online
        tmpStr = Reader.ReadString8
        If LenB(tmpStr) Then
            .txtTimeOn.Text = tmpStr
        Else
            .txtTimeOn.Text = "Usuario offline"
        End If
        
        'Observaciones
        tmpStr = Reader.ReadString8
        If LenB(tmpStr) Then
            .txtObs.Text = tmpStr
        Else
            .txtObs.Text = "Sin observaciones"
        End If
    End With

End Sub

''
' Handles the ShowMenu message.

Private Sub HandleShowMenu()
'***************************************************
'Author: ZaMa
'Last Modification: 15/05/2010
'
'***************************************************
'Check if the packet is complete

    
    ' [WGL -> TODO] REMOVE THIS PACKET PLEASE

    Reader.ReadInt8

End Sub

''
' Handles the StrDextRunningOut message.

Private Sub HandleStrDextRunningOut()
'***************************************************
'Author: CHOTS
'Last Modification: 01/12/2014
'01/12/2014: D'Artagnan - Text changed.
'***************************************************

    Call ShowConsoleMsg("¡Tus atributos están próximos a regresar a su estado original!", 255, 130, 170, True)
    frmMain.tmrBlink.Enabled = True

End Sub

''
' Handles the CharacterAttackMovement

Private Sub HandleCharacterAttackMovement()
'***************************************************
'Author: Amraphen
'Last Modification: 24/05/2010
'
'***************************************************

Dim CharIndex As Integer


        CharIndex = Reader.ReadInt16
        
        With charlist(CharIndex)
            Call InitGrh(.Arma.WeaponWalk(.Heading), .Arma.WeaponWalk(.Heading).GrhIndex, , False)
            
            .UsandoArma = True
        End With
        

End Sub

''
' Handles the TournamentCompetitorList message.

Private Sub HandleTournamentCompetitorList()
'***************************************************
'Author: ZaMa
'Last Modification: 07/06/2012
'
'***************************************************

    Tournament.CompetitorsList = Split(Reader.ReadString8(), SEPARATOR)
    
    ' Update list
    frmTournament.Load_lstCompetitors

End Sub

''
' Handles the TournamentConfig message.

Private Sub HandleTournamentConfig()
'***************************************************
'Author: ZaMa
'Last Modification: 07/06/2012
'
'***************************************************

    Dim lTemp As Long
    With Tournament
        ' General
        .MinLevel = Reader.ReadInt8
        .MaxLevel = Reader.ReadInt8
        .MaxCompetitors = Reader.ReadInt8
        .NumRoundsToWin = Reader.ReadInt8
        .RequiredGold = Reader.ReadInt32
        .KillAfterLoose = Reader.ReadInt8
        
        ' Classes
        .NumPermitedClass = Reader.ReadInt8
        
        If .NumPermitedClass <> 0 Then ReDim .PermitedClass(1 To .NumPermitedClass)
        For lTemp = 1 To .NumPermitedClass
            .PermitedClass(lTemp) = Reader.ReadInt8
        Next lTemp
        
        ' Items
        .NumForbiddenItems = Reader.ReadInt8
        
        If .NumForbiddenItems <> 0 Then ReDim .ForbiddenItem(1 To .NumForbiddenItems)
        For lTemp = 1 To .NumForbiddenItems
            .ForbiddenItem(lTemp) = Reader.ReadInt16
        Next lTemp
    
        ' Maps
        .WaitingMap.Map = Reader.ReadInt16
        .WaitingMap.X = Reader.ReadInt8
        .WaitingMap.Y = Reader.ReadInt8
        
        .FinalMap.Map = Reader.ReadInt16
        .FinalMap.X = Reader.ReadInt8
        .FinalMap.Y = Reader.ReadInt8
        
        For lTemp = 1 To MAX_ARENAS
            .Arenas(lTemp).Map = Reader.ReadInt16
            
            .Arenas(lTemp).UserPos1.X = Reader.ReadInt8
            .Arenas(lTemp).UserPos1.X = Reader.ReadInt8
            
            .Arenas(lTemp).UserPos2.X = Reader.ReadInt8
            .Arenas(lTemp).UserPos2.X = Reader.ReadInt8
        Next lTemp
    End With

    ' Update screen
    frmTournament.UpdateConfig

End Sub
 
Private Sub HandleAccountPersonaje()
    
    '
    ' @ Agrega un personaje.
     

    Dim i_Slot      As Byte
    Dim temp_Data   As modAccount.tAccChars

    With temp_Data

        'Cargo el slot.
        i_Slot = Reader.ReadInt8()
        
        'Busco el nombre.
        .Char_Name = Reader.ReadString8()
        
        .Char_Map_Name = Reader.ReadString8()
        .Char_Muerto = Reader.ReadBool()
        .bSailing = Reader.ReadBool()
        .Char_Nivel = Reader.ReadInt8()
        
        .Alignment = Reader.ReadInt8()
        .IdGuild = Reader.ReadInt()
        .GuildName = Reader.ReadString8()
        .JailRemainingTime = Reader.ReadInt
        .Banned = Reader.ReadBool
        
        'Cargo los datos del char.
        With .Char_Character
            .Body = Reader.ReadInt16()
            .Head = Reader.ReadInt16()
            .Arma = Reader.ReadInt16()
            .Escudo = Reader.ReadInt16()
            .Casco = Reader.ReadInt16()
        End With
        
        If .Char_Name <> vbNullString Then
            ' Increase character counter.
            If i_Slot > Acc_Data.nCharCount Then
                Acc_Data.nCharCount = i_Slot
            End If
            
            'Agrego el personaje a la lista.
            Call modAccount.Agregar_Personaje(i_Slot, temp_Data)
        End If
        
    End With

End Sub

Private Sub HandleAccountRemoveChar()
'***************************************************
'Author: D'Artagnan
'Date: 20/06/2014
'Last Modification: 20/06/2014
'Remove a character from the account form.
'***************************************************
'Check if the packet is complete

        Call AccountRemoveCharacter(Reader.ReadInt8())

    If (frmAccount.Visible) Then frmAccount.Refresh

End Sub

Private Sub HandleAccountShow()
'
' @ Muestra la cuenta.
'

    ' Logging in account.
    If frmAccountCreate.Visible Then
        Unload frmAccountCreate
    ' Just finished with character creation.
    ElseIf frmCrearPersonaje.Visible Then
        Unload frmCrearPersonaje
    End If
    
    frmConnect.txtPasswd = vbNullString
    frmConnect.Visible = False
    
    frmAccount.Caption = Acc_Data.Acc_Name & " - " & App.Title
    'frmAccount.cmdLog.Enabled = CBool(Acc_Data.Acc_Char_Selected)
    frmAccount.Visible = True
    frmAccount.Show , frmConnect
    frmAccount.SetFocus

End Sub
 
Private Sub HandleAccountQuestion()

    ' TODO: Nightw - I think this function can be removed. I'll leave it there just for now.

    'Guarda la pregunta.
    Reader.ReadString8

End Sub

Private Sub HandleGetPunishmentTypeList()

  
    Dim amountOfTypes As Integer
    Dim amountOfRules As Integer
    
    Dim I As Integer, J As Integer

        frmPunishmentAdm.punishmentType = Reader.ReadInt8()
        frmPunishmentAdm.userToPunish = Reader.ReadString8()
        
        'Get the amount of punishment types
        amountOfTypes = Reader.ReadInt16
                
        ReDim ModPunishments.punishmentList(amountOfTypes)
        
        For I = 0 To amountOfTypes
            ModPunishments.punishmentList(I).Id = Reader.ReadInt16()
            ModPunishments.punishmentList(I).Name = Reader.ReadString8()
            ModPunishments.punishmentList(I).BaseType = frmPunishmentAdm.punishmentType
        Next I

        frmPunishmentAdm.Show

End Sub


Private Sub HandleShowLoginScreen()

  
    frmConnect.Show
    

End Sub

Private Sub HandleCloseForm()
'***************************************************
'Author: D'Artagnan
'Creation Date: 07/11/2014
'Last Modification: 07/11/2014
'Close the specified form.
'***************************************************

  
    Dim sFormName As String
    Dim frmForm As Form

        sFormName = Reader.ReadString8()

        ' Find target.
        For Each frmForm In Forms
            If frmForm.Name = sFormName Then
                ' Close it.
                Unload frmForm
            End If
        Next
        
        ' Specific behavior.
        If sFormName = "frmAccount" Then
            frmConnect.Visible = True
        End If

End Sub


Private Sub HandleEnableBerserker()

  
    ' [WGL -> TODO] Remove this packet PLEASE!!!!

    Reader.ReadBool

End Sub

Private Sub HandleMensajeDuelo()

    Dim Mode As Byte
    Dim MyTeam(1 To 3) As String
    Dim EnemyTeam(1 To 4) As String
    Dim Drop As Boolean
    Dim Resucitar As Boolean
    Dim Bet As Long
    Dim TeamMate As Boolean

    Mode = Reader.ReadInt8
    Select Case Mode
        Case 0 'vs1
            EnemyTeam(1) = Reader.ReadString8
            
            Bet = Reader.ReadInt32
            Drop = Reader.ReadBool

            Call AddtoRichTextBox(frmMain.RecTxt(0), EnemyTeam(1) & " te ha desafio a un duelo por " & Bet & " monedas de oro", 65, 190, 156, 1, 0)
        Case 1 'vs2
            MyTeam(1) = Reader.ReadString8
            EnemyTeam(1) = Reader.ReadString8
            EnemyTeam(2) = Reader.ReadString8
            
            Bet = Reader.ReadInt32
            Drop = Reader.ReadBool
            Resucitar = Reader.ReadBool
            TeamMate = Reader.ReadBool
            
            If TeamMate Then
                Call AddtoRichTextBox(frmMain.RecTxt(0), MyTeam(1) & " te ha invitado a pelear con el en un duelo 2vs2 contra " & EnemyTeam(1) & " y " & EnemyTeam(2) & " por " & Bet & " monedas de oro", 65, 190, 156, 1, 0)
            Else
                Call AddtoRichTextBox(frmMain.RecTxt(0), EnemyTeam(1) & " y " & EnemyTeam(2) & " los han desafiado a ti y a " & MyTeam(1) & " a un duelo 2vs2 por " & Bet & " monedas de oro", 65, 190, 156, 1, 0)
            End If
        Case 2 'vs3
            MyTeam(1) = Reader.ReadString8
            MyTeam(2) = Reader.ReadString8
            EnemyTeam(1) = Reader.ReadString8
            EnemyTeam(2) = Reader.ReadString8
            EnemyTeam(3) = Reader.ReadString8
            
            Bet = Reader.ReadInt32
            Drop = Reader.ReadBool
            Resucitar = Reader.ReadBool
            TeamMate = Reader.ReadBool
            
            If TeamMate Then
                Call AddtoRichTextBox(frmMain.RecTxt(0), MyTeam(1) & " te ha invitado a pelear con el y con " & MyTeam(2) & " en un duelo 3vs3 contra " & EnemyTeam(1) & ", " & EnemyTeam(2) & " y " & EnemyTeam(3) & " por " & Bet & " monedas de oro", 65, 190, 156, 1, 0)
            Else
                Call AddtoRichTextBox(frmMain.RecTxt(0), EnemyTeam(1) & ", " & EnemyTeam(2) & " y " & EnemyTeam(3) & " los han desafiado a ti, a " & MyTeam(1) & " y a " & MyTeam(2) & " a un duelo 3vs3 por " & Bet & " monedas de oro", 65, 190, 156, 1, 0)
            End If
        Case 3 'vs4
            MyTeam(1) = Reader.ReadString8
            MyTeam(2) = Reader.ReadString8
            MyTeam(3) = Reader.ReadString8
            EnemyTeam(1) = Reader.ReadString8
            EnemyTeam(2) = Reader.ReadString8
            EnemyTeam(3) = Reader.ReadString8
            EnemyTeam(4) = Reader.ReadString8
            
            Bet = Reader.ReadInt32
            Drop = Reader.ReadBool
            Resucitar = Reader.ReadBool
            TeamMate = Reader.ReadBool
            
            If TeamMate Then
                Call AddtoRichTextBox(frmMain.RecTxt(0), MyTeam(1) & " te ha invitado a pelear con el, con " & MyTeam(2) & " y con " & MyTeam(3) & " en un duelo 4vs4 contra " & EnemyTeam(1) & ", " & EnemyTeam(2) & ", " & EnemyTeam(3) & " y " & EnemyTeam(4) & " por " & Bet & " monedas de oro", 65, 190, 156, 1, 0)
            Else
                Call AddtoRichTextBox(frmMain.RecTxt(0), EnemyTeam(1) & ", " & EnemyTeam(2) & ", " & EnemyTeam(3) & " y " & EnemyTeam(4) & " los han desafiado a ti, a " & MyTeam(1) & ", a " & MyTeam(2) & " y a " & MyTeam(3) & " a un duelo 4vs4 por " & Bet & " monedas de oro", 65, 190, 156, 1, 0)
            End If
    End Select
    

    If Mode = 2 Or Mode = 3 Then
        If Drop Then
            Call AddtoRichTextBox(frmMain.RecTxt(0), " con drop, ", 255, 0, 0, 1, 0, 0)
        Else
            Call AddtoRichTextBox(frmMain.RecTxt(0), " sin drop, ", 65, 190, 156, 1, 0, 0)
        End If
        Call AddtoRichTextBox(frmMain.RecTxt(0), IIf(Resucitar, "y est?ermitido resucitar.", "y no est?ermitido resucitar."), 65, 190, 156, 1, 0, 0)
    Else
        If Drop Then
            Call AddtoRichTextBox(frmMain.RecTxt(0), " con drop.", 255, 0, 0, 1, 0, 0)
        Else
            Call AddtoRichTextBox(frmMain.RecTxt(0), " sin drop.", 65, 190, 156, 1, 0, 0)
        End If
    End If
    Call AddtoRichTextBox(frmMain.RecTxt(0), "Escribe /AceptarDuelo para aceptar el duelo o /RechazarDuelo para rechazar el duelo.", 65, 190, 156, 1, 0)

    

End Sub

Private Sub HandleRetar()

    frmDuelos.Show , frmMain

End Sub

Private Sub HandleOkDueloPublico()
'---------------------------------------------------------------------------------------
' Module    : Protocol
' Author    : Anagrama
' Date      : 20/08/2016
' Purpose   : Abre la ventana de espera del duelo.
'---------------------------------------------------------------------------------------

    If frmEsperandoDuelo.Visible = False Then
        frmEsperandoDuelo.Show vbModeless, frmMain
    Else
        Unload frmEsperandoDuelo
    End If

End Sub


Private Sub HandleAccBankChangeSlot()
'***************************************************
'Author: Anagrama
'Creation Date: 21/08/2016
'Actualiza el slot de la boveda de cuenta.
'***************************************************

    Dim Slot As Byte
    Slot = Reader.ReadInt8()
    
    Dim ObjIndex As Integer
    ObjIndex = Reader.ReadInt16()
    
    Dim Amount As Integer
    Amount = Reader.ReadInt16()
    
    Dim CanUse As Boolean
    CanUse = Reader.ReadBool()
    
    If Slot = 0 Then
        ReDim AccBankInventory(1 To MAX_BANCOINVENTORY_SLOTS) As Inventory
    Else
        Dim sMain As String
        sMain = "OBJ" & CStr(ObjIndex)
        
        With AccBankInventory(Slot)
            If ObjIndex > 0 Then
                .ObjIndex = ObjIndex
                .Name = GameMetadata.Objs(ObjIndex).Name
                .Amount = Amount
                .GrhIndex = GameMetadata.Objs(ObjIndex).GrhIndex
                .OBJType = GameMetadata.Objs(ObjIndex).OBJType
                .MaxHit = GameMetadata.Objs(ObjIndex).MaxHit
                .MinHit = GameMetadata.Objs(ObjIndex).MinHit
                .MaxDef = GameMetadata.Objs(ObjIndex).MaxDef
                .MinDef = GameMetadata.Objs(ObjIndex).MinDef
                .Valor = GameMetadata.Objs(ObjIndex).Valor
                .CanUse = CanUse
            
            Else
                .ObjIndex = 0
                .Name = "Nada"
                .Amount = 0
                .GrhIndex = 0
                .OBJType = 0
                .MaxHit = 0
                .MinHit = 0
                .MaxDef = 0
                .MinDef = 0
                .Valor = 0
                .CanUse = True
            End If
            
            
            If Comerciando Then
                Call AccBank(0).SetItem(Slot, .ObjIndex, .Amount, _
                    .Equipped, .GrhIndex, .OBJType, .MaxHit, _
                    .MinHit, .MaxDef, .MinDef, .Valor, .Name, 0, .CanUse)
            End If
        End With
    End If

End Sub

Private Sub HandleAccBankInit()
'***************************************************
'Author: Anagrama
'Last Modification: 21/08/06
'Inicia la transaccion con la boveda de cuenta
'***************************************************

  
    Dim I As Long
    Dim AccBankGold As Long

    Set AccBank(0) = New clsGraphicalInventory
    Set AccBank(1) = New clsGraphicalInventory
    
    Load frmBancoAcc
    
    AccBankGold = Reader.ReadInt32
    Call AccBank(0).Initialize(frmBancoAcc.PicBancoInv, MAX_BANCOINVENTORY_SLOTS, , , , , , , , , True, _
                                eMoveType.Target, AccBank(1))

    Call AccBank(1).Initialize(frmBancoAcc.PicInv, Inventario.MaxObjs, , , , , , , , , True, _
                                eMoveType.Inventory, AccBank(0))
    
    For I = 1 To Inventario.MaxObjs
        With Inventario
            Call AccBank(1).SetItem(I, .ObjIndex(I), _
                .Amount(I), .Equipped(I), .GrhIndex(I), _
                .OBJType(I), .MaxHit(I), .MinHit(I), .MaxDef(I), .MinDef(I), _
                .Valor(I), .ItemName(I), 0, .CanUse(I))
        End With
    Next I
    
    For I = 1 To MAX_BANCOINVENTORY_SLOTS
        With AccBankInventory(I)
            Call AccBank(0).SetItem(I, .ObjIndex, _
                .Amount, .Equipped, .GrhIndex, _
                .OBJType, .MaxHit, .MinHit, .MaxDef, .MinDef, _
                .Valor, .Name, 0, .CanUse)
        End With
    Next I
    frmBancoAcc.lblUserGld.Caption = AccBankGold
    
    'Set state and show form
    Comerciando = True
        
    frmBancoAcc.Show , frmMain
    If frmBancoAccPass.Visible = True Then Unload frmBancoAccPass
 
End Sub

Private Sub HandleAccBankUpdateGold()
'***************************************************
'Author: Anagrama
'Last Modification: 21/08/06
'Actualiza el oro de la boveda de cuenta
'***************************************************
 
    frmBancoAcc.lblUserGld.Caption = Reader.ReadInt32
    If frmBancoAcc.Visible = False Then
        Unload frmBancoAcc
    End If
    
End Sub

Private Sub HandleAccBankEnd()
'***************************************************
'Author: Anagrama
'Last Modification: 21/08/06
'Termina la transaccion con la boveda de cuenta
'***************************************************

    Set AccBank(0) = Nothing
    Set AccBank(1) = Nothing
    
    Unload frmBancoAcc
    Unload frmBancoAccPass
    Unload frmBancoAccChangePass
    Comerciando = False
    
End Sub

Private Sub HandleAccBankRequestPass()
'***************************************************
'Author: Anagrama
'Last Modification: 21/08/06
'Solicita la contraseña de la boveda
'***************************************************

    frmBancoAccPass.Show , frmMain

End Sub

Private Sub HandleSendPetList()
'***************************************************
'Author: Nightw
'Last Modification: 18/09/2016
'Receive the list of pets tammed by the user
'***************************************************

    Dim I As Integer

    HasPets = False
    PetListQty = Reader.ReadInt8
    
    If PetListQty > 0 Then
        ReDim PetList(1 To PetListQty)
    
        For I = 1 To PetListQty
            PetList(I) = Reader.ReadInt16
    
            If PetList(I) > 0 Then
                HasPets = True
            End If
        Next I
    End If
    
    If Not HasPets Then
        PetSelectedIndex = 0
    End If
    
    If Not frmMascotas.Visible Then
        Unload frmMascotas
    Else
        Call frmMascotas.RefreshBoxes
    End If
    
    

End Sub


Private Sub HandleBlacksmithUpgrades()
'***************************************************
'Author: Anagrama
'Last Modification: 07/04/2017
'
'***************************************************


    Dim Count As Integer
    Dim MatCount As Integer
    Dim I As Long
    Dim J As Long
    Dim k As Long
    
    Count = Reader.ReadInt16()
    
    ReDim UpgradeHerrero(Count) As tItemsConstruibles
    
    For I = 1 To Count
        With UpgradeHerrero(I)
            .StationRecipeIndex = Reader.ReadInt16()
            .ObjIndex = Reader.ReadInt16()
            .GrhIndex = GameMetadata.Objs(.ObjIndex).GrhIndex
            .Name = GameMetadata.Objs(.ObjIndex).Name
            MatCount = Reader.ReadInt16()
            ReDim .CraftItem(1 To MatCount)
            For k = 1 To MatCount
                .CraftItem(k).ObjIndex = Reader.ReadInt16()
                .CraftItem(k).Amount = Reader.ReadInt16()
            Next k
        End With
    Next I

End Sub


Private Sub HandleCarpenterUpgrades()
'***************************************************
'Author: Anagrama
'Last Modification: 07/04/2017
'
'***************************************************

    Dim Count As Integer
    Dim MatCount As Integer
    Dim I As Long
    Dim J As Long
    Dim k As Long
    
    Count = Reader.ReadInt16()
    
    ReDim UpgradeCarpintero(Count) As tItemsConstruibles
    
    For I = 1 To Count
        With UpgradeCarpintero(I)
            .StationRecipeIndex = Reader.ReadInt16()
            .ObjIndex = Reader.ReadInt16()
            .GrhIndex = GameMetadata.Objs(.ObjIndex).GrhIndex
            .Name = GameMetadata.Objs(.ObjIndex).Name
            MatCount = Reader.ReadInt16()
            ReDim .CraftItem(1 To MatCount)
            For k = 1 To MatCount
                .CraftItem(k).ObjIndex = Reader.ReadInt16()
                .CraftItem(k).Amount = Reader.ReadInt16()
            Next k
        End With
    Next I

End Sub

Public Sub HandleSendSessionToken()

    Acc_Data.Acc_Token = Reader.ReadString8

End Sub

Public Sub HandleSendMasteries()
On Error GoTo ErrHandler
    Dim I As Integer
    Dim J As Integer
    
    With PlayerData
        If Reader.ReadInt8 = 1 Then
             Erase .ClassMasteryGroups
             
             .ClassMasteryGroupsQty = Reader.ReadInt16
             
             If .ClassMasteryGroupsQty > 0 Then
                ReDim .ClassMasteryGroups(1 To .ClassMasteryGroupsQty)
                
                For I = 1 To .ClassMasteryGroupsQty
                    .ClassMasteryGroups(I).MasteriesQty = Reader.ReadInt16
                    
                    If .ClassMasteryGroups(I).MasteriesQty > 0 Then
                        ReDim .ClassMasteryGroups(I).Masteries(1 To .ClassMasteryGroups(I).MasteriesQty)
                        
                        For J = 1 To .ClassMasteryGroups(I).MasteriesQty
                            .ClassMasteryGroups(I).Masteries(J) = Reader.ReadInt16
                        Next J
                    End If
                Next I
             End If
            
        Else
             Erase .MasteryGroups
             
             .MasteryGroupsQty = Reader.ReadInt16
             
             If .MasteryGroupsQty > 0 Then
                ReDim .MasteryGroups(1 To .MasteryGroupsQty)
                
                For I = 1 To .MasteryGroupsQty
                    .MasteryGroups(I).GroupId = Reader.ReadInt16
                    .MasteryGroups(I).MasteriesQty = Reader.ReadInt16
                    
                    If .MasteryGroups(I).MasteriesQty > 0 Then
                        ReDim .MasteryGroups(I).Masteries(1 To .MasteryGroups(I).MasteriesQty)
                        
                        For J = 1 To .MasteryGroups(I).MasteriesQty
                            .MasteryGroups(I).Masteries(J) = Reader.ReadInt16
                        Next J
                    End If
                Next I
             End If
             
            Call frmMasteries.Show(, frmMain)
            
            Call frmMasteries.DrawAllMasteries
            Call frmMasteries.SetFocus
        
        End If
    End With
    
    Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleSendMasteries de Protocol.bas")
End Sub


''
' Writes the "LoginExistingChar" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLoginExistingChar()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "LoginExistingChar" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.LoginExistingChar)
        
    If UserName <> uName Then UserName = UserName & vbNullChar & uName
    Call Writer.WriteString8(UserName)
    Call Writer.WriteString8(UserPassword)
        
    Call Writer.WriteInt8(App.Major)
    Call Writer.WriteInt8(App.Minor)
    Call Writer.WriteInt8(App.Revision)

#If EnableSecurity Then
    Call Writer.WriteString8(MD5HushYo)
#End If

#If SeguridadTesteo Then
    Call Writer.WriteString8(getCode)
#End If

    Call Send(False)
    
    Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteLoginExistingChar de Protocol.bas")
End Sub

''
' Writes the "Talk" message to the outgoing data buffer.
'
' @param    chat The chat text to be sent.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTalk(ByVal chat As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Talk" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ClientPacketID.Talk)
        
        Call Writer.WriteString8(chat)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteTalk de Protocol.bas")
End Sub

''
' Writes the "Yell" message to the outgoing data buffer.
'
' @param    chat The chat text to be sent.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteYell(ByVal chat As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Yell" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ClientPacketID.Yell)
        
        Call Writer.WriteString8(chat)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteYell de Protocol.bas")
End Sub

''
' Writes the "Whisper" message to the outgoing data buffer.
'
' @param    charIndex The index of the char to whom to whisper.
' @param    chat The chat text to be sent to the user.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWhisper(ByVal CharName As String, ByVal chat As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 03/12/10
'Writes the "Whisper" message to the outgoing data buffer
'03/12/10: Enanoh - Ahora se envía el nick y no el charindex.
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ClientPacketID.Whisper)
        
        Call Writer.WriteString8(CharName)
        
        Call Writer.WriteString8(chat)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteWhisper de Protocol.bas")
End Sub

''
' Writes the "Walk" message to the outgoing data buffer.
'
' @param    heading The direction in wich the user is moving.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWalk(ByVal Heading As E_Heading)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Walk" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ClientPacketID.Walk)
        
        Call Writer.WriteInt8(Heading)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteWalk de Protocol.bas")
End Sub

''
' Writes the "RequestPositionUpdate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestPositionUpdate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestPositionUpdate" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.RequestPositionUpdate)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteRequestPositionUpdate de Protocol.bas")
End Sub

''
' Writes the "Attack" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAttack()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Attack" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.Attack)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteAttack de Protocol.bas")
End Sub

''
' Writes the "PickUp" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePickUp()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PickUp" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.PickUp)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WritePickUp de Protocol.bas")
End Sub

''
' Writes the "SafeToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSafeToggle()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SafeToggle" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.SafeToggle)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteSafeToggle de Protocol.bas")
End Sub

''
' Writes the "ResuscitationSafeToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResuscitationToggle()
'**************************************************************
'Author: Rapsodius
'Creation Date: 10/10/07
'Writes the Resuscitation safe toggle packet to the outgoing data buffer.
'**************************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.ResuscitationSafeToggle)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteResuscitationToggle de Protocol.bas")
End Sub

Public Sub WriteRequestPartyForm()
'***************************************************
'Author: Budi
'Last Modification: 11/26/09
'Writes the "RequestPartyForm" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.RequestPartyForm)
    
    Call Send(False)
   
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteRequestPartyForm de Protocol.bas")
End Sub

''
' Writes the "ItemUpgrade" message to the outgoing data buffer.
'
' @param    ItemIndex The index to the item to upgrade.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteItemUpgrade(ByVal ItemIndex As Integer)
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 12/09/09
'Writes the "ItemUpgrade" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.ItemUpgrade)
    Call Writer.WriteInt16(ItemIndex)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteItemUpgrade de Protocol.bas")
End Sub

''
' Writes the "RequestStadictis" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestStadictis()
'***************************************************
'Author: ZaMa
'Last Modification: 24/05/2012
'Writes the "RequestStadictis" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.RequestStadictis)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteRequestStadictis de Protocol.bas")
End Sub

''
' Writes the "RequestSkills" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestSkills()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestSkills" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.RequestSkills)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteRequestSkills de Protocol.bas")
End Sub

''
' Writes the "CommerceEnd" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceEnd()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CommerceEnd" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.CommerceEnd)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteCommerceEnd de Protocol.bas")
End Sub

''
' Writes the "UserCommerceEnd" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceEnd()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserCommerceEnd" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.UserCommerceEnd)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteUserCommerceEnd de Protocol.bas")
End Sub

''
' Writes the "UserCommerceConfirm" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceConfirm()
'***************************************************
'Author: ZaMa
'Last Modification: 14/12/2009
'Writes the "UserCommerceConfirm" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.UserCommerceConfirm)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteUserCommerceConfirm de Protocol.bas")
End Sub

''
' Writes the "BankEnd" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankEnd()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankEnd" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.BankEnd)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteBankEnd de Protocol.bas")
End Sub

''
' Writes the "UserCommerceOk" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceOk()
'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 01/10/07
'Writes the "UserCommerceOk" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.UserCommerceOk)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteUserCommerceOk de Protocol.bas")
End Sub

''
' Writes the "UserCommerceReject" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceReject()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserCommerceReject" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.UserCommerceReject)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteUserCommerceReject de Protocol.bas")
End Sub

''
' Writes the "Drop" message to the outgoing data buffer.
'
' @param    slot Inventory slot where the item to drop is.
' @param    amount Number of items to drop.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDropXY(ByVal Slot As Byte, ByVal Amount As Integer, X As Byte, Y As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Drop" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ClientPacketID.DropXY)
        
        Call Writer.WriteInt8(Slot)
        Call Writer.WriteInt16(Amount)
        Call Writer.WriteInt8(X)
        Call Writer.WriteInt8(Y)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteDropXY de Protocol.bas")
End Sub

Public Sub WriteDrop(ByVal Slot As Byte, ByVal Amount As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Drop" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ClientPacketID.Drop)
        
        Call Writer.WriteInt8(Slot)
        Call Writer.WriteInt16(Amount)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteDrop de Protocol.bas")
End Sub

''
' Writes the "CastSpell" message to the outgoing data buffer.
'
' @param    slot Spell List slot where the spell to cast is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCastSpell(ByVal Slot As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CastSpell" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ClientPacketID.CastSpell)
        
        Call Writer.WriteInt8(Slot)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteCastSpell de Protocol.bas")
End Sub

''
' Writes the "LeftClick" message to the outgoing data buffer.
'
' @param    X Tile coord in the x-axis in which the user clicked.
' @param    y Tile coord in the y-axis in which the user clicked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLeftClick(ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "LeftClick" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ClientPacketID.LeftClick)
        
        Call Writer.WriteInt8(X)
        Call Writer.WriteInt8(Y)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteLeftClick de Protocol.bas")
End Sub

''
' Writes the "RightClick" message to the outgoing data buffer.
'
' @param    x Tile coord in the x-axis in which the user clicked.
' @param    y Tile coord in the y-axis in which the user clicked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRightClick(ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: ZaMa
'Last Modification: 15/05/2011
'Writes the "RightClick" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ClientPacketID.RightClick)
        
        Call Writer.WriteInt8(X)
        Call Writer.WriteInt8(Y)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteRightClick de Protocol.bas")
End Sub

''
' Writes the "DoubleClick" message to the outgoing data buffer.
'
' @param    x Tile coord in the x-axis in which the user clicked.
' @param    y Tile coord in the y-axis in which the user clicked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDoubleClick(ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "DoubleClick" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ClientPacketID.DoubleClick)
        
        Call Writer.WriteInt8(X)
        Call Writer.WriteInt8(Y)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteDoubleClick de Protocol.bas")
End Sub

''
' Writes the "Work" message to the outgoing data buffer.
'
' @param    skill The skill which the user attempts to use.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWork(ByVal Skill As eSkill)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Work" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ClientPacketID.Work)
        
        Call Writer.WriteInt8(Skill)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteWork de Protocol.bas")
End Sub

''
' Writes the "UseSpellMacro" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUseSpellMacro()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UseSpellMacro" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.UseSpellMacro)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteUseSpellMacro de Protocol.bas")
End Sub

''
' Writes the "UseItem" message to the outgoing data buffer.
'
' @param    slot Invetory slot where the item to use is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUseItem(ByVal Slot As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UseItem" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ClientPacketID.UseItem)
        
        Call Writer.WriteInt8(Slot)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteUseItem de Protocol.bas")
End Sub

''
' Writes the "WorkLeftClick" message to the outgoing data buffer.
'
' @param    x Tile coord in the x-axis in which the user clicked.
' @param    y Tile coord in the y-axis in which the user clicked.
' @param    skill The skill which the user attempts to use.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWorkLeftClick(ByVal X As Byte, ByVal Y As Byte, ByVal Skill As eSkill)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "WorkLeftClick" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ClientPacketID.WorkLeftClick)
        
        Call Writer.WriteInt8(X)
        Call Writer.WriteInt8(Y)
        
        Call Writer.WriteInt8(Skill)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteWorkLeftClick de Protocol.bas")
End Sub

''
' Writes the "SpellInfo" message to the outgoing data buffer.
'
' @param    slot Spell List slot where the spell which's info is requested is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpellInfo(ByVal Slot As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SpellInfo" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.SpellInfo)
        
        Call Writer.WriteInt8(Slot)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteSpellInfo de Protocol.bas")
End Sub

''
' Writes the "EquipItem" message to the outgoing data buffer.
'
' @param    slot Invetory slot where the item to equip is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteEquipItem(ByVal Slot As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "EquipItem" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.EquipItem)
        
        Call Writer.WriteInt8(Slot)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteEquipItem de Protocol.bas")
End Sub

''
' Writes the "ChangeHeading" message to the outgoing data buffer.
'
' @param    heading The direction in wich the user is moving.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeHeading(ByVal Heading As E_Heading)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeHeading" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.ChangeHeading)
        
        Call Writer.WriteInt8(Heading)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteChangeHeading de Protocol.bas")
End Sub

''
' Writes the "ModifySkills" message to the outgoing data buffer.
'
' @param    skillEdt a-based array containing for each skill the number of points to add to it.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteModifySkills(ByRef skillEdt() As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ModifySkills" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Dim I As Long
    
    
        Call Writer.WriteInt8(ClientPacketID.ModifySkills)
        
        For I = 1 To NUMSKILLS
            Call Writer.WriteInt8(skillEdt(I))
        Next I
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteModifySkills de Protocol.bas")
End Sub

''
' Writes the "Train" message to the outgoing data buffer.
'
' @param    creature Position within the list provided by the server of the creature to train against.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTrain(ByVal creature As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Train" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.Train)
        
        Call Writer.WriteInt8(creature)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteTrain de Protocol.bas")
End Sub

''
' Writes the "CommerceBuy" message to the outgoing data buffer.
'
' @param    slot Position within the NPC's inventory in which the desired item is.
' @param    amount Number of items to buy.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceBuy(ByVal Slot As Byte, ByVal Amount As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CommerceBuy" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.CommerceBuy)
        
        Call Writer.WriteInt8(Slot)
        Call Writer.WriteInt16(Amount)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteCommerceBuy de Protocol.bas")
End Sub

''
' Writes the "BankExtractItem" message to the outgoing data buffer.
'
' @param    slot Position within the bank in which the desired item is.
' @param    amount Number of items to extract.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankExtractItem(ByVal Slot As Byte, ByVal Amount As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankExtractItem" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.BankExtractItem)
        
        Call Writer.WriteInt8(Slot)
        Call Writer.WriteInt16(Amount)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteBankExtractItem de Protocol.bas")
End Sub

''
' Writes the "CommerceSell" message to the outgoing data buffer.
'
' @param    slot Position within user inventory in which the desired item is.
' @param    amount Number of items to sell.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceSell(ByVal Slot As Byte, ByVal Amount As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CommerceSell" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.CommerceSell)
        
        Call Writer.WriteInt8(Slot)
        Call Writer.WriteInt16(Amount)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteCommerceSell de Protocol.bas")
End Sub

''
' Writes the "BankDeposit" message to the outgoing data buffer.
'
' @param    slot Position within the user inventory in which the desired item is.
' @param    amount Number of items to deposit.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankDeposit(ByVal Slot As Byte, ByVal Amount As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankDeposit" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.BankDeposit)
        
        Call Writer.WriteInt8(Slot)
        Call Writer.WriteInt16(Amount)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteBankDeposit de Protocol.bas")
End Sub

''
' Writes the "ForumPost" message to the outgoing data buffer.
'
' @param    title The message's title.
' @param    message The body of the message.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForumPost(ByVal Title As String, ByVal Message As String, ByVal ForumMsgType As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ForumPost" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.ForumPost)
        
        Call Writer.WriteInt8(ForumMsgType)
        Call Writer.WriteString8(Title)
        Call Writer.WriteString8(Message)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteForumPost de Protocol.bas")
End Sub

''
' Writes the "MoveSpell" message to the outgoing data buffer.
'
' @param    upwards True if the spell will be moved up in the list, False if it will be moved downwards.
' @param    slot Spell List slot where the spell which's info is requested is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMoveSpell(ByVal upwards As Boolean, ByVal Slot As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "MoveSpell" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.MoveSpell)
        
        Call Writer.WriteBool(upwards)
        Call Writer.WriteInt8(Slot)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteMoveSpell de Protocol.bas")
End Sub

''
' Writes the "MoveBank" message to the outgoing data buffer.
'
' @param    upwards True if the item will be moved up in the list, False if it will be moved downwards.
' @param    slot Bank List slot where the item which's info is requested is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMoveBank(ByVal upwards As Boolean, ByVal Slot As Byte)
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 06/14/09
'Writes the "MoveBank" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.MoveBank)
        
        Call Writer.WriteBool(upwards)
        Call Writer.WriteInt8(Slot)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteMoveBank de Protocol.bas")
End Sub

''
' Writes the "UserCommerceOffer" message to the outgoing data buffer.
'
' @param    slot Position within user inventory in which the desired item is.
' @param    amount Number of items to offer.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceOffer(ByVal Slot As Byte, ByVal Amount As Long, ByVal OfferSlot As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserCommerceOffer" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ClientPacketID.UserCommerceOffer)
        
        Call Writer.WriteInt8(Slot)
        Call Writer.WriteInt32(Amount)
        Call Writer.WriteInt8(OfferSlot)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteUserCommerceOffer de Protocol.bas")
End Sub

''
' Writes the "WriteUserCommerceOfferGold" message to the outgoing data buffer.
'
' @param    amount Number of items to offer.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceOfferGold(ByVal Amount As Long)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserCommerceOffer" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ClientPacketID.UserCommerceOfferGold)
        Call Writer.WriteInt32(Amount)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteUserCommerceOffer de Protocol.bas")
End Sub

Public Sub WriteCommerceChat(ByVal chat As String)
'***************************************************
'Author: ZaMa
'Last Modification: 03/12/2009
'Writes the "CommerceChat" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ClientPacketID.CommerceChat)
        
        Call Writer.WriteString8(chat)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteCommerceChat de Protocol.bas")
End Sub

''
' Writes the "Online" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnline()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Online" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.Online)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteOnline de Protocol.bas")
End Sub

''
' Writes the "Quit" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteQuit()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 08/16/08
'Writes the "Quit" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.Quit)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteQuit de Protocol.bas")
End Sub

''
' Writes the "RequestAccountState" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestAccountState()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestAccountState" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.RequestAccountState)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteRequestAccountState de Protocol.bas")
End Sub

''
' Writes the "PetStand" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePetStand()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PetStand" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.PetStand)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WritePetStand de Protocol.bas")
End Sub

''
' Writes the "PetFollow" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePetFollow()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PetFollow" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.PetFollow)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WritePetFollow de Protocol.bas")
End Sub

''
' Writes the "ReleasePet" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReleasePet(ByVal fromForm As Boolean, Optional ByVal petSlot As Byte = 0)
'***************************************************
'Author: ZaMa
'Last Modification: 18/11/2009
'Writes the "ReleasePet" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.ReleasePet)
    Call Writer.WriteBool(fromForm)
    Call Writer.WriteInt8(petSlot)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteReleasePet de Protocol.bas")
End Sub

''
' Writes the "ReleasePet" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReleaseTammedPet(ByVal PetIndex As Byte)
'***************************************************
'Author: ZaMa
'Last Modification: 18/11/2009
'Writes the "ReleasePet" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.ReleasePet)
        
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteReleasePet de Protocol.bas")
End Sub


''
' Writes the "TrainList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTrainList()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TrainList" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.TrainList)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteTrainList de Protocol.bas")
End Sub

''
' Writes the "Rest" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRest()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Rest" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.Rest)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteRest de Protocol.bas")
End Sub

''
' Writes the "Meditate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMeditate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Meditate" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.Meditate)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteMeditate de Protocol.bas")
End Sub

''
' Writes the "Resucitate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResucitate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Resucitate" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.Resucitate)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteResucitate de Protocol.bas")
End Sub

''
' Writes the "Consultation" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteConsultation()
'***************************************************
'Author: ZaMa
'Last Modification: 01/05/2010
'Writes the "Consultation" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.Consultation)
    
    Call Send(False)
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteConsultation de Protocol.bas")
End Sub

''
' Writes the "Heal" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteHeal()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Heal" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.Heal)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteHeal de Protocol.bas")
End Sub

''
' Writes the "Help" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteHelp()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Help" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.Help)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteHelp de Protocol.bas")
End Sub

''
' Writes the "RequestStats" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestStats()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestStats" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.RequestStats)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteRequestStats de Protocol.bas")
End Sub

''
' Writes the "CommerceStart" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceStart()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CommerceStart" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.CommerceStart)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteCommerceStart de Protocol.bas")
End Sub

''
' Writes the "BankStart" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankStart()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankStart" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.BankStart)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteBankStart de Protocol.bas")
End Sub

''
' Writes the "Enlist" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteEnlist()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Enlist" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.Enlist)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteEnlist de Protocol.bas")
End Sub

''
' Writes the "Information" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteInformation()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Information" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.Information)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteInformation de Protocol.bas")
End Sub

''
' Writes the "Reward" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReward()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Reward" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.Reward)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteReward de Protocol.bas")
End Sub

''
' Writes the "UpTime" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpTime()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpTime" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.UpTime)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteUpTime de Protocol.bas")
End Sub

''
' Writes the "PartyLeave" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyLeave()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PartyLeave" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.PartyLeave)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WritePartyLeave de Protocol.bas")
End Sub

''
' Writes the "PartyCreate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyCreate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PartyCreate" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.PartyCreate)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WritePartyCreate de Protocol.bas")
End Sub


''
' Writes the "Inquiry" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteInquiry()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Inquiry" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.Inquiry)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteInquiry de Protocol.bas")
End Sub

''
' Writes the "GuildMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the guild.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildMessage(ByVal Message As String)
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ClientPacketID.GuildMessage)
        
        Call Writer.WriteString8(Message)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteGuildMessage de Protocol.bas")
End Sub



''
' Writes the "PartyMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the party.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PartyMessage" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ClientPacketID.PartyMessage)
        
        Call Writer.WriteString8(Message)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WritePartyMessage de Protocol.bas")
End Sub

''
' Writes the "PartyOnline" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyOnline()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PartyOnline" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.PartyOnline)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WritePartyOnline de Protocol.bas")
End Sub

''
' Writes the "CouncilMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the other council members.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCouncilMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CouncilMessage" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.CouncilMessage)
        
        Call Writer.WriteString8(Message)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteCouncilMessage de Protocol.bas")
End Sub

''
' Writes the "RoleMasterRequest" message to the outgoing data buffer.
'
' @param    message The message to send to the role masters.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRoleMasterRequest(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RoleMasterRequest" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.RoleMasterRequest)
        
        Call Writer.WriteString8(Message)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteRoleMasterRequest de Protocol.bas")
End Sub

''
' Writes the "GMRequest" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGMRequest()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GMRequest" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.GMRequest)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteGMRequest de Protocol.bas")
End Sub

''
' Writes the "BugReport" message to the outgoing data buffer.
'
' @param    message The message explaining the reported bug.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBugReport(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BugReport" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ClientPacketID.bugReport)
        
        Call Writer.WriteString8(Message)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteBugReport de Protocol.bas")
End Sub

''
' Writes the "ChangeDescription" message to the outgoing data buffer.
'
' @param    desc The new description of the user's character.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeDescription(ByVal Desc As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeDescription" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.ChangeDescription)
        
        Call Writer.WriteString8(Desc)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteChangeDescription de Protocol.bas")
End Sub

''
' Writes the "Punishments" message to the outgoing data buffer.
'
' @param    username The user whose's  punishments are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePunishments(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Punishments" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  

    UserName = complexNameToSimple(UserName, True)

    
        Call Writer.WriteInt8(ClientPacketID.Punishments)
        
        Call Writer.WriteString8(UserName)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WritePunishments de Protocol.bas")
End Sub

''
' Writes the "Gamble" message to the outgoing data buffer.
'
' @param    amount The amount to gamble.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGamble(ByVal Amount As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Gamble" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.Gamble)
        
        Call Writer.WriteInt16(Amount)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteGamble de Protocol.bas")
End Sub

''
' Writes the "InquiryVote" message to the outgoing data buffer.
'
' @param    opt The chosen option to vote for.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteInquiryVote(ByVal opt As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "InquiryVote" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.InquiryVote)
        
        Call Writer.WriteInt8(opt)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteInquiryVote de Protocol.bas")
End Sub

''
' Writes the "LeaveFaction" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLeaveFaction()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "LeaveFaction" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.LeaveFaction)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteLeaveFaction de Protocol.bas")
End Sub

''
' Writes the "BankExtractGold" message to the outgoing data buffer.
'
' @param    amount The amount of money to extract from the bank.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankExtractGold(ByVal Amount As Long)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankExtractGold" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.BankExtractGold)
        
        Call Writer.WriteInt32(Amount)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteBankExtractGold de Protocol.bas")
End Sub

''
' Writes the "BankDepositGold" message to the outgoing data buffer.
'
' @param    amount The amount of money to deposit in the bank.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankDepositGold(ByVal Amount As Long)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankDepositGold" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.BankDepositGold)
        
        Call Writer.WriteInt32(Amount)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteBankDepositGold de Protocol.bas")
End Sub

''
' Writes the "Denounce" message to the outgoing data buffer.
'
' @param    message The message to send with the denounce.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDenounce(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Denounce" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.Denounce)
        
        Call Writer.WriteString8(Message)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteDenounce de Protocol.bas")
End Sub

''
' Writes the "PartyKick" message to the outgoing data buffer.
'
' @param    username The user to kick fro mthe party.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyKick(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PartyKick" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    UserName = complexNameToSimple(UserName, True)
    
    
        Call Writer.WriteInt8(ClientPacketID.PartyKick)
            
        Call Writer.WriteString8(UserName)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WritePartyKick de Protocol.bas")
End Sub

''
' Writes the "PartySetLeader" message to the outgoing data buffer.
'
' @param    username The user to set as the party's leader.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartySetLeader(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PartySetLeader" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    UserName = complexNameToSimple(UserName, True)
    
    
        Call Writer.WriteInt8(ClientPacketID.PartySetLeader)
        
        Call Writer.WriteString8(UserName)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WritePartySetLeader de Protocol.bas")
End Sub


''
' Writes the "GuildMemberList" message to the outgoing data buffer.
'
' @param    guild The guild whose member list is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildMemberList(ByVal Guild As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildMemberList" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
    Call Writer.WriteInt8(ClientPacketID.GMCommands)
    Call Writer.WriteInt8(eGMCommands.GuildMemberList)
    
    Call Writer.WriteString8(Guild)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteGuildMemberList de Protocol.bas")
End Sub

Public Sub WriteAdminChangeGuildAlign(ByVal GuildName As String, ByVal NewAlignment As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildMemberList" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.AdminChangeGuildAlign)
        
        Call Writer.WriteString8(GuildName)
        Call Writer.WriteInt8(NewAlignment)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteAdminChangeGuildAlign de Protocol.bas")
End Sub

''
' Writes the "InitCrafting" message to the outgoing data buffer.
'
' @param    Cantidad The final aumont of item to craft.

Public Sub WriteInitCrafting(ByVal cantidad As Long)
'***************************************************
'Author: ZaMa
'Last Modification: 29/01/2010
'Writes the "InitCrafting" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.InitCrafting)
        Call Writer.WriteInt32(cantidad)
            
    Call Send(False)
    

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteInitCrafting de Protocol.bas")
End Sub

''
' Writes the "Home" message to the outgoing data buffer.
'
Public Sub WriteHome()
'***************************************************
'Author: Budi
'Last Modification: 01/06/10
'Writes the "Home" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ClientPacketID.Home)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteHome de Protocol.bas")
End Sub



''
' Writes the "GMMessage" message to the outgoing data buffer.
'
' @param    message The message to be sent to the other GMs online.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGMMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GMMessage" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.GMMessage)
        Call Writer.WriteString8(Message)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteGMMessage de Protocol.bas")
End Sub

''
' Writes the "ShowName" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowName()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowName" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.GMCommands)
    Call Writer.WriteInt8(eGMCommands.showName)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteShowName de Protocol.bas")
End Sub

''
' Writes the "OnlineRoyalArmy" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnlineRoyalArmy()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "OnlineRoyalArmy" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.GMCommands)
    Call Writer.WriteInt8(eGMCommands.OnlineRoyalArmy)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteOnlineRoyalArmy de Protocol.bas")
End Sub

''
' Writes the "OnlineChaosLegion" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnlineChaosLegion()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "OnlineChaosLegion" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.GMCommands)
    Call Writer.WriteInt8(eGMCommands.OnlineChaosLegion)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteOnlineChaosLegion de Protocol.bas")
End Sub

''
' Writes the "GoNearby" message to the outgoing data buffer.
'
' @param    username The suer to approach.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGoNearby(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GoNearby" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    UserName = complexNameToSimple(UserName, True)
    
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.GoNearby)
        
        Call Writer.WriteString8(UserName)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteGoNearby de Protocol.bas")
End Sub

''
' Writes the "Comment" message to the outgoing data buffer.
'
' @param    message The message to leave in the log as a comment.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteComment(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Comment" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.Comment)
        
        Call Writer.WriteString8(Message)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteComment de Protocol.bas")
End Sub

''
' Writes the "ServerTime" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteServerTime()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ServerTime" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.GMCommands)
    Call Writer.WriteInt8(eGMCommands.serverTime)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteServerTime de Protocol.bas")
End Sub

''
' Writes the "Where" message to the outgoing data buffer.
'
' @param    username The user whose position is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWhere(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Where" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    UserName = complexNameToSimple(UserName, True)
    
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.Where)
        
        Call Writer.WriteString8(UserName)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteWhere de Protocol.bas")
End Sub

''
' Writes the "CreaturesInMap" message to the outgoing data buffer.
'
' @param    map The map in which to check for the existing creatures.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreaturesInMap(ByVal Map As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CreaturesInMap" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.CreaturesInMap)
        
        Call Writer.WriteInt16(Map)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteCreaturesInMap de Protocol.bas")
End Sub

''
' Writes the "WarpMeToTarget" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWarpMeToTarget()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "WarpMeToTarget" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.GMCommands)
    Call Writer.WriteInt8(eGMCommands.WarpMeToTarget)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteWarpMeToTarget de Protocol.bas")
End Sub

''
' Writes the "WarpChar" message to the outgoing data buffer.
'
' @param    username The user to be warped. "YO" represent's the user's char.
' @param    map The map to which to warp the character.
' @param    x The x position in the map to which to waro the character.
' @param    y The y position in the map to which to waro the character.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWarpChar(ByVal UserName As String, ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "WarpChar" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    UserName = complexNameToSimple(UserName, True)
    
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.WarpChar)
        
        Call Writer.WriteString8(UserName)
        
        Call Writer.WriteInt16(Map)
        
        Call Writer.WriteInt8(X)
        Call Writer.WriteInt8(Y)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteWarpChar de Protocol.bas")
End Sub

''
' Writes the "Silence" message to the outgoing data buffer.
'
' @param    username The user to silence.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSilence(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Silence" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    UserName = complexNameToSimple(UserName, True)
    
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.Silence)
        
        Call Writer.WriteString8(UserName)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteSilence de Protocol.bas")
End Sub

''
' Writes the "SOSShowList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSOSShowList()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SOSShowList" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.GMCommands)
    Call Writer.WriteInt8(eGMCommands.SOSShowList)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteSOSShowList de Protocol.bas")
End Sub

''
' Writes the "SOSRemove" message to the outgoing data buffer.
'
' @param    username The user whose SOS call has been already attended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSOSRemove(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SOSRemove" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.SOSRemove)
        
        Call Writer.WriteString8(UserName)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteSOSRemove de Protocol.bas")
End Sub

''
' Writes the "GoToChar" message to the outgoing data buffer.
'
' @param    username The user to be approached.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGoToChar(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GoToChar" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    UserName = complexNameToSimple(UserName, True)
    
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.GoToChar)
        
        Call Writer.WriteString8(UserName)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteGoToChar de Protocol.bas")
End Sub

''
' Writes the "invisible" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteInvisible()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "invisible" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.GMCommands)
    Call Writer.WriteInt8(eGMCommands.invisible)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteInvisible de Protocol.bas")
End Sub

''
' Writes the "GMPanel" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGMPanel()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GMPanel" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.GMCommands)
    Call Writer.WriteInt8(eGMCommands.GMPanel)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteGMPanel de Protocol.bas")
End Sub

''
' Writes the "RequestUserList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestUserList()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestUserList" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.GMCommands)
    Call Writer.WriteInt8(eGMCommands.RequestUserList)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteRequestUserList de Protocol.bas")
End Sub


''
' Writes the "Jail" message to the outgoing data buffer.
'
' @param    username The user to be sent to jail.
' @param    reason The reason for which to send him to jail.
' @param    time The time (in minutes) the user will have to spend there.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteJail(ByVal UserName As String, ByRef Reason As String, ByRef AdminNotes As String, ByVal PunishmentID As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Jail" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    UserName = complexNameToSimple(UserName, True)
    
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.Jail)
        
        Call Writer.WriteString8(UserName)
        Call Writer.WriteString8(Reason)
        
        'Call Writer.WriteInt8(Time)
        
        Call Writer.WriteInt16(PunishmentID) ' Punishment ID
        
        Call Writer.WriteString8(AdminNotes) ' AdminNotes
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteJail de Protocol.bas")
End Sub

''
' Writes the "KillNPC" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKillNPC()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "KillNPC" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.GMCommands)
    Call Writer.WriteInt8(eGMCommands.KillNPC)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteKillNPC de Protocol.bas")
End Sub

''
' Writes the "WarnUser" message to the outgoing data buffer.
'
' @param    username The user to be warned.
' @param    reason Reason for the warning.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWarnUser(ByRef UserName As String, ByRef Reason As String, ByRef AdminNotes As String, ByVal PunishmentID As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "WarnUser" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    UserName = complexNameToSimple(UserName, True)
    
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.WarnUser)
        
        Call Writer.WriteString8(UserName)
        Call Writer.WriteString8(Reason)
        Call Writer.WriteString8(AdminNotes)
        Call Writer.WriteInt16(PunishmentID)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteWarnUser de Protocol.bas")
End Sub

''
' Writes the "EditChar" message to the outgoing data buffer.
'
' @param    UserName    The user to be edited.
' @param    editOption  Indicates what to edit in the char.
' @param    arg1        Additional argument 1. Contents depend on editoption.
' @param    arg2        Additional argument 2. Contents depend on editoption.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteEditChar(ByVal UserName As String, ByVal EditOption As eEditOptions, ByVal Arg1 As String, ByVal arg2 As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "EditChar" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    UserName = complexNameToSimple(UserName, True)
    
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.EditChar)
        
        Call Writer.WriteString8(UserName)
        
        Call Writer.WriteInt8(EditOption)
        
        Call Writer.WriteString8(Arg1)
        Call Writer.WriteString8(arg2)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteEditChar de Protocol.bas")
End Sub

Public Sub WriteRequestStatsBosses()
    Call Writer.WriteInt8(ClientPacketID.GMCommands)
    Call Writer.WriteInt8(eGMCommands.RequestStatsBosses)
        
    Call Send(False)
End Sub

''
' Writes the "RequestCharInfo" message to the outgoing data buffer.
'
' @param    username The user whose information is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharInfo(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestCharInfo" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    UserName = complexNameToSimple(UserName, True)
    
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.RequestCharInfo)
        
        Call Writer.WriteString8(UserName)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteRequestCharInfo de Protocol.bas")
End Sub

''
' Writes the "RequestCharStats" message to the outgoing data buffer.
'
' @param    username The user whose stats are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharStats(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestCharStats" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    UserName = complexNameToSimple(UserName, True)
    
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.RequestCharStats)
        
        Call Writer.WriteString8(UserName)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteRequestCharStats de Protocol.bas")
End Sub

''
' Writes the "RequestCharGold" message to the outgoing data buffer.
'
' @param    username The user whose gold is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharGold(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestCharGold" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    UserName = complexNameToSimple(UserName, True)
    
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.RequestCharGold)
        
        Call Writer.WriteString8(UserName)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteRequestCharGold de Protocol.bas")
End Sub
    
''
' Writes the "RequestCharInventory" message to the outgoing data buffer.
'
' @param    username The user whose inventory is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharInventory(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestCharInventory" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    UserName = complexNameToSimple(UserName, True)
    
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.RequestCharInventory)
        
        Call Writer.WriteString8(UserName)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteRequestCharInventory de Protocol.bas")
End Sub

''
' Writes the "RequestCharBank" message to the outgoing data buffer.
'
' @param    username The user whose banking information is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharBank(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestCharBank" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    UserName = complexNameToSimple(UserName, True)
    
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.RequestCharBank)
        
        Call Writer.WriteString8(UserName)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteRequestCharBank de Protocol.bas")
End Sub

''
' Writes the "RequestCharSkills" message to the outgoing data buffer.
'
' @param    username The user whose skills are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharSkills(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestCharSkills" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    UserName = complexNameToSimple(UserName, True)
    
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.RequestCharSkills)
        
        Call Writer.WriteString8(UserName)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteRequestCharSkills de Protocol.bas")
End Sub

''
' Writes the "ReviveChar" message to the outgoing data buffer.
'
' @param    username The user to eb revived.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReviveChar(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ReviveChar" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    UserName = complexNameToSimple(UserName, True)
    
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.ReviveChar)
        
        Call Writer.WriteString8(UserName)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteReviveChar de Protocol.bas")
End Sub

''
' Writes the "OnlineGM" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnlineGM()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "OnlineGM" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.GMCommands)
    Call Writer.WriteInt8(eGMCommands.OnlineGM)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteOnlineGM de Protocol.bas")
End Sub

''
' Writes the "OnlineMap" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnlineMap(ByVal Map As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 26/03/2009
'Writes the "OnlineMap" message to the outgoing data buffer
'26/03/2009: Now you don't need to be in the map to use the comand, so you send the map to server
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.OnlineMap)
        
        Call Writer.WriteInt16(Map)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteOnlineMap de Protocol.bas")
End Sub

''
' Writes the "Kick" message to the outgoing data buffer.
'
' @param    username The user to be kicked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKick(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Kick" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    UserName = complexNameToSimple(UserName, True)
    
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.Kick)
        
        Call Writer.WriteString8(UserName)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteKick de Protocol.bas")
End Sub

''
' Writes the "Execute" message to the outgoing data buffer.
'
' @param    username The user to be executed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteExecute(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Execute" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    UserName = complexNameToSimple(UserName, True)
    
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.Execute)
        
        Call Writer.WriteString8(UserName)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteExecute de Protocol.bas")
End Sub

''
' Writes the "BanChar" message to the outgoing data buffer.
'
' @param    username The user to be banned.
' @param    reason The reason for which the user is to be banned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBanChar(ByVal UserName As String, ByVal Reason As String, ByRef AdminNotes As String, ByVal PunishmentTypeId As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BanChar" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    UserName = complexNameToSimple(UserName, True)
    
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.banChar)
        Call Writer.WriteString8(UserName)
        Call Writer.WriteString8(Reason)
        Call Writer.WriteString8(AdminNotes)
        Call Writer.WriteInt16(PunishmentTypeId)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteBanChar de Protocol.bas")
End Sub

''
' Writes the "UnbanChar" message to the outgoing data buffer.
'
' @param    username The user to be unbanned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUnbanChar(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UnbanChar" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    UserName = complexNameToSimple(UserName, True)
    
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.UnbanChar)
        
        Call Writer.WriteString8(UserName)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteUnbanChar de Protocol.bas")
End Sub

''
' Writes the "NPCFollow" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNPCFollow()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "NPCFollow" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.GMCommands)
    Call Writer.WriteInt8(eGMCommands.NPCFollow)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteNPCFollow de Protocol.bas")
End Sub

''
' Writes the "SummonChar" message to the outgoing data buffer.
'
' @param    username The user to be summoned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSummonChar(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SummonChar" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    UserName = complexNameToSimple(UserName, False)
    
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.SummonChar)
        
        Call Writer.WriteString8(UserName)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteSummonChar de Protocol.bas")
End Sub

''
' Writes the "SpawnListRequest" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpawnListRequest()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SpawnListRequest" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.GMCommands)
    Call Writer.WriteInt8(eGMCommands.SpawnListRequest)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteSpawnListRequest de Protocol.bas")
End Sub

''
' Writes the "SpawnCreature" message to the outgoing data buffer.
'
' @param    creatureIndex The index of the creature in the spawn list to be spawned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpawnCreature(ByVal creatureIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SpawnCreature" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.SpawnCreature)
        
        Call Writer.WriteInt16(creatureIndex)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteSpawnCreature de Protocol.bas")
End Sub

''
' Writes the "ResetNPCInventory" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResetNPCInventory()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ResetNPCInventory" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.GMCommands)
    Call Writer.WriteInt8(eGMCommands.ResetNPCInventory)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteResetNPCInventory de Protocol.bas")
End Sub

''
' Writes the "CleanWorld" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCleanWorld()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CleanWorld" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.GMCommands)
    Call Writer.WriteInt8(eGMCommands.CleanWorld)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteCleanWorld de Protocol.bas")
End Sub

''
' Writes the "ServerMessage" message to the outgoing data buffer.
'
' @param    message The message to be sent to players.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteServerMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ServerMessage" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.ServerMessage)
        
        Call Writer.WriteString8(Message)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteServerMessage de Protocol.bas")
End Sub
''
' Writes the "MapMessage" message to the outgoing data buffer.
'
' @param    message The message to be sent to players.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMapMessage(ByVal Message As String)
'***************************************************
'Author: ZaMa
'Last Modification: 14/11/2010
'Writes the "MapMessage" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.MapMessage)
        
        Call Writer.WriteString8(Message)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteMapMessage de Protocol.bas")
End Sub

''
' Writes the "NickToIP" message to the outgoing data buffer.
'
' @param    username The user whose IP is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNickToIP(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "NickToIP" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    UserName = complexNameToSimple(UserName, True)
    
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.nickToIP)
        
        Call Writer.WriteString8(UserName)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteNickToIP de Protocol.bas")
End Sub

''
' Writes the "IPToNick" message to the outgoing data buffer.
'
' @param    IP The IP for which to search for players. Must be an array of 4 elements with the 4 components of the IP.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteIPToNick(ByRef Ip() As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "IPToNick" message to the outgoing data buffer
'***************************************************
    If UBound(Ip()) - LBound(Ip()) + 1 <> 4 Then Exit Sub   'Invalid IP
On Error GoTo ErrHandler
  
    
    Dim I As Long
    
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.IPToNick)
        
        For I = LBound(Ip()) To UBound(Ip())
            Call Writer.WriteInt8(Ip(I))
        Next I
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteIPToNick de Protocol.bas")
End Sub

''
' Writes the "GuildOnlineMembers" message to the outgoing data buffer.
'
' @param    guild The guild whose online player list is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOnlineMembers(ByVal Guild As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildOnlineMembers" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
    Call Writer.WriteInt8(ClientPacketID.GMCommands)
    Call Writer.WriteInt8(eGMCommands.GuildOnlineMembers)
    
    Call Writer.WriteString8(Guild)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteGuildOnlineMembers de Protocol.bas")
End Sub

''
' Writes the "TeleportCreate" message to the outgoing data buffer.
'
' @param    map the map to which the teleport will lead.
' @param    x The position in the x axis to which the teleport will lead.
' @param    y The position in the y axis to which the teleport will lead.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTeleportCreate(ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte, Optional ByVal Radio As Byte = 0)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TeleportCreate" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
            Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.TeleportCreate)
        
        Call Writer.WriteInt16(Map)
        
        Call Writer.WriteInt8(X)
        Call Writer.WriteInt8(Y)
        
        Call Writer.WriteInt8(Radio)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteTeleportCreate de Protocol.bas")
End Sub

''
' Writes the "TeleportDestroy" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTeleportDestroy()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TeleportDestroy" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.GMCommands)
    Call Writer.WriteInt8(eGMCommands.TeleportDestroy)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteTeleportDestroy de Protocol.bas")
End Sub

''
' Writes the "RainToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRainToggle()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RainToggle" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.GMCommands)
    Call Writer.WriteInt8(eGMCommands.RainToggle)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteRainToggle de Protocol.bas")
End Sub

''
' Writes the "SetCharDescription" message to the outgoing data buffer.
'
' @param    desc The description to set to players.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetCharDescription(ByVal Desc As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SetCharDescription" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.SetCharDescription)
        
        Call Writer.WriteString8(Desc)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteSetCharDescription de Protocol.bas")
End Sub

''
' Writes the "ForceMIDIToMap" message to the outgoing data buffer.
'
' @param    midiID The ID of the midi file to play.
' @param    map The map in which to play the given midi.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForceMIDIToMap(ByVal midiID As Byte, ByVal Map As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ForceMIDIToMap" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.ForceMIDIToMap)
        
        Call Writer.WriteInt8(midiID)
        
        Call Writer.WriteInt16(Map)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteForceMIDIToMap de Protocol.bas")
End Sub

''
' Writes the "ForceWAVEToMap" message to the outgoing data buffer.
'
' @param    waveID  The ID of the wave file to play.
' @param    Map     The map into which to play the given wave.
' @param    x       The position in the x axis in which to play the given wave.
' @param    y       The position in the y axis in which to play the given wave.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForceWAVEToMap(ByVal waveID As Byte, ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ForceWAVEToMap" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.ForceWAVEToMap)
        
        Call Writer.WriteInt8(waveID)
        
        Call Writer.WriteInt16(Map)
        
        Call Writer.WriteInt8(X)
        Call Writer.WriteInt8(Y)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteForceWAVEToMap de Protocol.bas")
End Sub

''
' Writes the "RoyalArmyMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the royal army members.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRoyalArmyMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RoyalArmyMessage" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.RoyalArmyMessage)
        
        Call Writer.WriteString8(Message)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteRoyalArmyMessage de Protocol.bas")
End Sub

''
' Writes the "ChaosLegionMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the chaos legion member.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChaosLegionMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChaosLegionMessage" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.ChaosLegionMessage)
        
        Call Writer.WriteString8(Message)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteChaosLegionMessage de Protocol.bas")
End Sub


''
' Writes the "TalkAsNPC" message to the outgoing data buffer.
'
' @param    message The message to send to the royal army members.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTalkAsNPC(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TalkAsNPC" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.TalkAsNPC)
        
        Call Writer.WriteString8(Message)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteTalkAsNPC de Protocol.bas")
End Sub

''
' Writes the "DestroyAllItemsInArea" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDestroyAllItemsInArea()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "DestroyAllItemsInArea" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.GMCommands)
    Call Writer.WriteInt8(eGMCommands.DestroyAllItemsInArea)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteDestroyAllItemsInArea de Protocol.bas")
End Sub

''
' Writes the "AcceptRoyalCouncilMember" message to the outgoing data buffer.
'
' @param    username The name of the user to be accepted into the royal army council.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAcceptRoyalCouncilMember(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "AcceptRoyalCouncilMember" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    UserName = complexNameToSimple(UserName, True)
    
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.AcceptRoyalCouncilMember)
        
        Call Writer.WriteString8(UserName)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteAcceptRoyalCouncilMember de Protocol.bas")
End Sub

''
' Writes the "AcceptChaosCouncilMember" message to the outgoing data buffer.
'
' @param    username The name of the user to be accepted as a chaos council member.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAcceptChaosCouncilMember(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "AcceptChaosCouncilMember" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    UserName = complexNameToSimple(UserName, True)
    
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.AcceptChaosCouncilMember)
        
        Call Writer.WriteString8(UserName)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteAcceptChaosCouncilMember de Protocol.bas")
End Sub

''
' Writes the "ItemsInTheFloor" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteItemsInTheFloor()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ItemsInTheFloor" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.GMCommands)
    Call Writer.WriteInt8(eGMCommands.ItemsInTheFloor)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteItemsInTheFloor de Protocol.bas")
End Sub

''
' Writes the "MakeDumb" message to the outgoing data buffer.
'
' @param    username The name of the user to be made dumb.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMakeDumb(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "MakeDumb" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    UserName = complexNameToSimple(UserName, True)
    
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.MakeDumb)
        
        Call Writer.WriteString8(UserName)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteMakeDumb de Protocol.bas")
End Sub

''
' Writes the "MakeDumbNoMore" message to the outgoing data buffer.
'
' @param    username The name of the user who will no longer be dumb.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMakeDumbNoMore(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "MakeDumbNoMore" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    UserName = complexNameToSimple(UserName, True)
    
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.MakeDumbNoMore)
        
        Call Writer.WriteString8(UserName)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteMakeDumbNoMore de Protocol.bas")
End Sub

''
' Writes the "DumpIPTables" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDumpIPTables()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "DumpIPTables" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.GMCommands)
    Call Writer.WriteInt8(eGMCommands.dumpIPTables)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteDumpIPTables de Protocol.bas")
End Sub

''
' Writes the "CouncilKick" message to the outgoing data buffer.
'
' @param    username The name of the user to be kicked from the council.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCouncilKick(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CouncilKick" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    UserName = complexNameToSimple(UserName, True)
    
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.CouncilKick)
        
        Call Writer.WriteString8(UserName)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteCouncilKick de Protocol.bas")
End Sub

''
' Writes the "SetTrigger" message to the outgoing data buffer.
'
' @param    trigger The type of trigger to be set to the tile.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetTrigger(ByVal Trigger As eTrigger)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SetTrigger" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.SetTrigger)
        
        Call Writer.WriteInt8(Trigger)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteSetTrigger de Protocol.bas")
End Sub

''
' Writes the "AskTrigger" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAskTrigger()
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 04/13/07
'Writes the "AskTrigger" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.GMCommands)
    Call Writer.WriteInt8(eGMCommands.AskTrigger)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteAskTrigger de Protocol.bas")
End Sub

''
' Writes the "BannedIPList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBannedIPList()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BannedIPList" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.GMCommands)
    Call Writer.WriteInt8(eGMCommands.BannedIPList)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteBannedIPList de Protocol.bas")
End Sub

''
' Writes the "BannedIPReload" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBannedIPReload()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BannedIPReload" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.GMCommands)
    Call Writer.WriteInt8(eGMCommands.BannedIPReload)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteBannedIPReload de Protocol.bas")
End Sub

''
' Writes the "GuildBan" message to the outgoing data buffer.
'
' @param    guild The guild whose members will be banned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildBan(ByVal Guild As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildBan" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.GuildBan)
        
        Call Writer.WriteString8(Guild)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteGuildBan de Protocol.bas")
End Sub

''
' Writes the "BanIP" message to the outgoing data buffer.
'
' @param    byIp    If set to true, we are banning by IP, otherwise the ip of a given character.
' @param    IP      The IP for which to search for players. Must be an array of 4 elements with the 4 components of the IP.
' @param    nick    The nick of the player whose ip will be banned.
' @param    reason  The reason for the ban.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBanIP(ByVal byIp As Boolean, ByRef Ip() As Byte, ByVal Nick As String, ByVal Reason As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BanIP" message to the outgoing data buffer
'***************************************************
    If byIp And UBound(Ip()) - LBound(Ip()) + 1 <> 4 Then Exit Sub   'Invalid IP
On Error GoTo ErrHandler
  
    
    Dim I As Long
    
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.BanIP)
        
        Call Writer.WriteBool(byIp)
        
        If byIp Then
            For I = LBound(Ip()) To UBound(Ip())
                Call Writer.WriteInt8(Ip(I))
            Next I
        Else
            Call Writer.WriteString8(Nick)
        End If
        
        Call Writer.WriteString8(Reason)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteBanIP de Protocol.bas")
End Sub

''
' Writes the "UnbanIP" message to the outgoing data buffer.
'
' @param    IP The IP for which to search for players. Must be an array of 4 elements with the 4 components of the IP.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUnbanIP(ByRef Ip() As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UnbanIP" message to the outgoing data buffer
'***************************************************
    If UBound(Ip()) - LBound(Ip()) + 1 <> 4 Then Exit Sub   'Invalid IP
On Error GoTo ErrHandler
  
    
    Dim I As Long
    
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.UnbanIP)
        
        For I = LBound(Ip()) To UBound(Ip())
            Call Writer.WriteInt8(Ip(I))
        Next I
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteUnbanIP de Protocol.bas")
End Sub

''
' Writes the "CreateItem" message to the outgoing data buffer.
'
' @param    itemIndex The index of the item to be created.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateItem(ByVal ItemIndex As Long)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CreateItem" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.CreateItem)
        Call Writer.WriteInt16(ItemIndex)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteCreateItem de Protocol.bas")
End Sub

''
' Writes the "DestroyItems" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDestroyItems()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "DestroyItems" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.GMCommands)
    Call Writer.WriteInt8(eGMCommands.DestroyItems)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteDestroyItems de Protocol.bas")
End Sub

''
' Writes the "FactionKick" message to the outgoing data buffer.
'
' @param    username The name of the user to be kicked from the faction.
Public Sub WriteFactionKick(ByVal UserName As String)
On Error GoTo ErrHandler
  
    UserName = complexNameToSimple(UserName, True)
    
    Call Writer.WriteInt8(ClientPacketID.GMCommands)
    Call Writer.WriteInt8(eGMCommands.FactionKick)
    
    Call Writer.WriteString8(UserName)
    
    Call Send(False)
    
    Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteChaosLegionKick de Protocol.bas")
End Sub


''
' Writes the "ForceMIDIAll" message to the outgoing data buffer.
'
' @param    midiID The id of the midi file to play.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForceMIDIAll(ByVal midiID As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ForceMIDIAll" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.ForceMIDIAll)
        
        Call Writer.WriteInt8(midiID)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteForceMIDIAll de Protocol.bas")
End Sub

''
' Writes the "ForceWAVEAll" message to the outgoing data buffer.
'
' @param    waveID The id of the wave file to play.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForceWAVEAll(ByVal waveID As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ForceWAVEAll" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.ForceWAVEAll)
        
        Call Writer.WriteInt8(waveID)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteForceWAVEAll de Protocol.bas")
End Sub

''
' Writes the "RemovePunishment" message to the outgoing data buffer.
'
' @param    username The user whose punishments will be altered.
' @param    punishment The id of the punishment to be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemovePunishment(ByVal UserName As String, ByVal punishment As Byte, ByVal NewText As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RemovePunishment" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    UserName = complexNameToSimple(UserName, True)
    
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.RemovePunishment)
        
        Call Writer.WriteString8(UserName)
        Call Writer.WriteInt8(punishment)
        Call Writer.WriteString8(NewText)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteRemovePunishment de Protocol.bas")
End Sub

''
' Writes the "TileBlockedToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTileBlockedToggle()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TileBlockedToggle" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.GMCommands)
    Call Writer.WriteInt8(eGMCommands.TileBlockedToggle)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteTileBlockedToggle de Protocol.bas")
End Sub

''
' Writes the "KillNPCNoRespawn" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKillNPCNoRespawn()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "KillNPCNoRespawn" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.GMCommands)
    Call Writer.WriteInt8(eGMCommands.KillNPCNoRespawn)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteKillNPCNoRespawn de Protocol.bas")
End Sub

''
' Writes the "KillAllNearbyNPCs" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKillAllNearbyNPCs()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "KillAllNearbyNPCs" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.GMCommands)
    Call Writer.WriteInt8(eGMCommands.KillAllNearbyNPCs)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteKillAllNearbyNPCs de Protocol.bas")
End Sub

''
' Writes the "LastIP" message to the outgoing data buffer.
'
' @param    username The user whose last IPs are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLastIP(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "LastIP" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    UserName = complexNameToSimple(UserName, True)
    
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.LastIP)
        
        Call Writer.WriteString8(UserName)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteLastIP de Protocol.bas")
End Sub

''
' Writes the "ChangeMOTD" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMOTD()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeMOTD" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.GMCommands)
    Call Writer.WriteInt8(eGMCommands.ChangeMOTD)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteChangeMOTD de Protocol.bas")
End Sub

''
' Writes the "SetMOTD" message to the outgoing data buffer.
'
' @param    message The message to be set as the new MOTD.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetMOTD(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SetMOTD" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.SetMOTD)
        
        Call Writer.WriteString8(Message)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteSetMOTD de Protocol.bas")
End Sub

''
' Writes the "SystemMessage" message to the outgoing data buffer.
'
' @param    message The message to be sent to all players.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSystemMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SystemMessage" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.SystemMessage)
        
        Call Writer.WriteString8(Message)
        
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteSystemMessage de Protocol.bas")
End Sub

''
' Writes the "CreateNPC" message to the outgoing data buffer.
'
' @param    npcIndex The index of the NPC to be created.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateNPC(ByVal NpcIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CreateNPC" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.CreateNPC)
        
        Call Writer.WriteInt16(NpcIndex)
        
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteCreateNPC de Protocol.bas")
End Sub

''
' Writes the "CreateNPCWithRespawn" message to the outgoing data buffer.
'
' @param    npcIndex The index of the NPC to be created.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateNPCWithRespawn(ByVal NpcIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CreateNPCWithRespawn" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.CreateNPCWithRespawn)
        
        Call Writer.WriteInt16(NpcIndex)
        
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteCreateNPCWithRespawn de Protocol.bas")
End Sub

''
' Writes the "ImperialArmour" message to the outgoing data buffer.
'
' @param    armourIndex The index of imperial armour to be altered.
' @param    objectIndex The index of the new object to be set as the imperial armour.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteImperialArmour(ByVal armourIndex As Byte, ByVal ObjectIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ImperialArmour" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.ImperialArmour)
        
        Call Writer.WriteInt8(armourIndex)
        
        Call Writer.WriteInt16(ObjectIndex)
        
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteImperialArmour de Protocol.bas")
End Sub

''
' Writes the "ChaosArmour" message to the outgoing data buffer.
'
' @param    armourIndex The index of chaos armour to be altered.
' @param    objectIndex The index of the new object to be set as the chaos armour.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChaosArmour(ByVal armourIndex As Byte, ByVal ObjectIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChaosArmour" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.ChaosArmour)
        
        Call Writer.WriteInt8(armourIndex)
        
        Call Writer.WriteInt16(ObjectIndex)
        
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteChaosArmour de Protocol.bas")
End Sub

''
' Writes the "NavigateToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNavigateToggle()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "NavigateToggle" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.GMCommands)
    Call Writer.WriteInt8(eGMCommands.NavigateToggle)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteNavigateToggle de Protocol.bas")
End Sub

''
' Writes the "ServerOpenToUsersToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteServerOpenToUsersToggle()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ServerOpenToUsersToggle" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.GMCommands)
    Call Writer.WriteInt8(eGMCommands.ServerOpenToUsersToggle)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteServerOpenToUsersToggle de Protocol.bas")
End Sub

''
' Writes the "ResetFactions" message to the outgoing data buffer.
'
' @param    username The name of the user who will be removed from any faction.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResetFactions(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ResetFactions" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    UserName = complexNameToSimple(UserName, True)
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.ResetFactions)
        
        Call Writer.WriteString8(UserName)
    
    Call Send(False)
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteResetFactions de Protocol.bas")
End Sub

''
' Writes the "RemoveCharFromGuild" message to the outgoing data buffer.
'
' @param    username The name of the user who will be removed from any guild.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemoveCharFromGuild(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RemoveCharFromGuild" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    UserName = complexNameToSimple(UserName, True)
    
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.RemoveCharFromGuild)
        
        Call Writer.WriteString8(UserName)
    
    Call Send(False)
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteRemoveCharFromGuild de Protocol.bas")
End Sub

''
' Writes the "WriteModGuildContribution" message to the outgoing data buffer.
'
' @param    GuildName The name of the guild to which we will be adding or substracting contribution points
' @param    Amount The amount of contribution points that will be added or substracted
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteModGuildContribution(ByVal GuildName As String, ByVal Amount As Long)
On Error GoTo ErrHandler
      
    Call Writer.WriteInt8(ClientPacketID.GMCommands)
    Call Writer.WriteInt8(eGMCommands.ModGuildContribution)
    
    Call Writer.WriteString8(GuildName)
    Call Writer.WriteInt32(Amount)
    
    Call Send(False)
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteRemoveCharFromGuild de Protocol.bas")
End Sub

''
' Writes the "RequestCharMail" message to the outgoing data buffer.
'
' @param    username The name of the user whose mail is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharMail(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestCharMail" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    UserName = complexNameToSimple(UserName, True)
    
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.RequestCharMail)
        
        Call Writer.WriteString8(UserName)
    
    Call Send(False)
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteRequestCharMail de Protocol.bas")
End Sub

''
' Writes the "AlterPassword" message to the outgoing data buffer.
'
' @param    username The name of the user whose mail is requested.
' @param    copyFrom The name of the user from which to copy the password.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAlterPassword(ByVal UserName As String, ByVal sNewPassword As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "AlterPassword" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    UserName = complexNameToSimple(UserName, True)
    
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.AlterPassword)
        
        Call Writer.WriteString8(UserName)
        Call Writer.WriteString8(MD5.GetMD5String(sNewPassword))
        Call MD5.MD5Reset
    
    Call Send(False)
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteAlterPassword de Protocol.bas")
End Sub

''
' Writes the "AlterMail" message to the outgoing data buffer.
'
' @param    username The name of the user whose mail is requested.
' @param    newMail The new email of the player.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAlterMail(ByVal UserName As String, ByVal newMail As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "AlterMail" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    UserName = complexNameToSimple(UserName, True)
    
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.AlterMail)
        
        Call Writer.WriteString8(UserName)
        Call Writer.WriteString8(newMail)
    
    Call Send(False)
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteAlterMail de Protocol.bas")
End Sub

''
' Writes the "AlterName" message to the outgoing data buffer.
'
' @param    username The name of the user whose mail is requested.
' @param    newName The new user name.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAlterName(ByVal UserName As String, ByVal newName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "AlterName" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    UserName = complexNameToSimple(UserName, True)
    
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.AlterName)
        
        Call Writer.WriteString8(UserName)
        Call Writer.WriteString8(newName)

    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteAlterName de Protocol.bas")
End Sub

Public Sub WriteAlterGuildName(ByVal GuildName As String, ByVal newGuildName As String)
'***************************************************
'Author: Lex!
'Last Modification: 14/05/12
'Writes the "AlterGuildName" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.AlterGuildName)
        
        Call Writer.WriteString8(GuildName)
        Call Writer.WriteString8(newGuildName)
    
    Call Send(False)
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteAlterGuildName de Protocol.bas")
End Sub

''
' Writes the "DoBackup" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDoBackup()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "DoBackup" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.GMCommands)
    Call Writer.WriteInt8(eGMCommands.DoBackUp)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteDoBackup de Protocol.bas")
End Sub

''
' Writes the "ShowGuildMessages" message to the outgoing data buffer.
'
' @param    guild The guild to listen to.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowGuildMessages(ByVal Guild As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowGuildMessages" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.ShowGuildMessages)
        
        Call Writer.WriteString8(Guild)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteShowGuildMessages de Protocol.bas")
End Sub

''
' Writes the "SaveMap" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSaveMap()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SaveMap" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.GMCommands)
    Call Writer.WriteInt8(eGMCommands.SaveMap)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteSaveMap de Protocol.bas")
End Sub

''
' Writes the "ChangeMapInfoPK" message to the outgoing data buffer.
'
' @param    isPK True if the map is PK, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoPK(ByVal isPK As Boolean)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeMapInfoPK" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.ChangeMapInfoPK)
        
        Call Writer.WriteBool(isPK)

      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteChangeMapInfoPK de Protocol.bas")
End Sub

''
' Writes the "ChangeMapInfoNoOcultar" message to the outgoing data buffer.
'
' @param    PermitirOcultar True if the map permits to hide, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoNoOcultar(ByVal PermitirOcultar As Boolean)
'***************************************************
'Author: ZaMa
'Last Modification: 19/09/2010
'Writes the "ChangeMapInfoNoOcultar" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.ChangeMapInfoNoOcultar)
        
        Call Writer.WriteBool(PermitirOcultar)

      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteChangeMapInfoNoOcultar de Protocol.bas")
End Sub

''
' Writes the "ChangeMapInfoNoInvocar" message to the outgoing data buffer.
'
' @param    PermitirInvocar True if the map permits to invoke, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoNoInvocar(ByVal PermitirInvocar As Boolean)
'***************************************************
'Author: ZaMa
'Last Modification: 18/09/2010
'Writes the "ChangeMapInfoNoInvocar" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.ChangeMapInfoNoInvocar)
        
        Call Writer.WriteBool(PermitirInvocar)
    
    Call Send(False)
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteChangeMapInfoNoInvocar de Protocol.bas")
End Sub

''
' Writes the "ChangeMapInfoBackup" message to the outgoing data buffer.
'
' @param    backup True if the map is to be backuped, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoBackup(ByVal backup As Boolean)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeMapInfoBackup" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.ChangeMapInfoBackup)
        
        Call Writer.WriteBool(backup)
    
    Call Send(False)
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteChangeMapInfoBackup de Protocol.bas")
End Sub

''
' Writes the "ChangeMapInfoRestricted" message to the outgoing data buffer.
'
' @param    restrict NEWBIES (only newbies), NO (everyone), ARMADA (just Armadas), CAOS (just caos) or FACCION (Armadas & caos only)
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoRestricted(ByVal restrict As String)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Writes the "ChangeMapInfoRestricted" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.ChangeMapInfoRestricted)
        
        Call Writer.WriteString8(restrict)
    
    Call Send(False)
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteChangeMapInfoRestricted de Protocol.bas")
End Sub

''
' Writes the "ChangeMapInfoNoMagic" message to the outgoing data buffer.
'
' @param    nomagic TRUE if no magic is to be allowed in the map.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoNoMagic(ByVal nomagic As Boolean)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Writes the "ChangeMapInfoNoMagic" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.ChangeMapInfoNoMagic)
        
        Call Writer.WriteBool(nomagic)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteChangeMapInfoNoMagic de Protocol.bas")
End Sub

''
' Writes the "ChangeMapInfoNoInvi" message to the outgoing data buffer.
'
' @param    noinvi TRUE if invisibility is not to be allowed in the map.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoNoInvi(ByVal noinvi As Boolean)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Writes the "ChangeMapInfoNoInvi" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.ChangeMapInfoNoInvi)
        
        Call Writer.WriteBool(noinvi)
    
    Call Send(False)
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteChangeMapInfoNoInvi de Protocol.bas")
End Sub
                            
''
' Writes the "ChangeMapInfoNoResu" message to the outgoing data buffer.
'
' @param    noresu TRUE if resurection is not to be allowed in the map.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoNoResu(ByVal noresu As Boolean)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Writes the "ChangeMapInfoNoResu" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler

        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.ChangeMapInfoNoResu)
        
        Call Writer.WriteBool(noresu)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteChangeMapInfoNoResu de Protocol.bas")
End Sub
                        
''
' Writes the "ChangeMapInfoLand" message to the outgoing data buffer.
'
' @param    land options: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoLand(ByVal land As String)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Writes the "ChangeMapInfoLand" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.ChangeMapInfoLand)
        
        Call Writer.WriteString8(land)
       
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteChangeMapInfoLand de Protocol.bas")
End Sub
                        
''
' Writes the "ChangeMapInfoZone" message to the outgoing data buffer.
'
' @param    zone options: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoZone(ByVal zone As String)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Writes the "ChangeMapInfoZone" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.ChangeMapInfoZone)
        
        Call Writer.WriteString8(zone)
        
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteChangeMapInfoZone de Protocol.bas")
End Sub

''
' Writes the "ChangeMapInfoStealNpc" message to the outgoing data buffer.
'
' @param    forbid TRUE if stealNpc forbiden.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoStealNpc(ByVal forbid As Boolean)
'***************************************************
'Author: ZaMa
'Last Modification: 25/07/2010
'Writes the "ChangeMapInfoStealNpc" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.ChangeMapInfoStealNpc)
        
        Call Writer.WriteBool(forbid)
        
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteChangeMapInfoStealNpc de Protocol.bas")
End Sub

''
' Writes the "SaveChars" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSaveChars()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SaveChars" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.GMCommands)
    Call Writer.WriteInt8(eGMCommands.SaveChars)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteSaveChars de Protocol.bas")
End Sub

''
' Writes the "CleanSOS" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCleanSOS()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CleanSOS" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.GMCommands)
    Call Writer.WriteInt8(eGMCommands.CleanSOS)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteCleanSOS de Protocol.bas")
End Sub

''
' Writes the "ShowServerForm" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowServerForm()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowServerForm" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.GMCommands)
    Call Writer.WriteInt8(eGMCommands.ShowServerForm)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteShowServerForm de Protocol.bas")
End Sub

''
' Writes the "ShowDenouncesList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowDenouncesList()
'***************************************************
'Author: ZaMa
'Last Modification: 14/11/2010
'Writes the "ShowDenouncesList" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.GMCommands)
    Call Writer.WriteInt8(eGMCommands.ShowDenouncesList)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteShowDenouncesList de Protocol.bas")
End Sub

''
' Writes the "EnableDenounces" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteEnableDenounces()
'***************************************************
'Author: ZaMa
'Last Modification: 14/11/2010
'Writes the "EnableDenounces" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.GMCommands)
    Call Writer.WriteInt8(eGMCommands.EnableDenounces)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteEnableDenounces de Protocol.bas")
End Sub

''
' Writes the "KickAllChars" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKickAllChars()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "KickAllChars" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.GMCommands)
    Call Writer.WriteInt8(eGMCommands.KickAllChars)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteKickAllChars de Protocol.bas")
End Sub

''
' Writes the "ReloadNPCs" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReloadNPCs()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ReloadNPCs" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.GMCommands)
    Call Writer.WriteInt8(eGMCommands.ReloadNPCs)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteReloadNPCs de Protocol.bas")
End Sub

''
' Writes the "ReloadServerIni" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReloadServerIni()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ReloadServerIni" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.GMCommands)
    Call Writer.WriteInt8(eGMCommands.ReloadServerIni)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteReloadServerIni de Protocol.bas")
End Sub

''
' Writes the "ReloadSpells" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReloadSpells()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ReloadSpells" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.GMCommands)
    Call Writer.WriteInt8(eGMCommands.ReloadSpells)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteReloadSpells de Protocol.bas")
End Sub

''
' Writes the "ReloadObjects" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReloadObjects()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ReloadObjects" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.GMCommands)
    Call Writer.WriteInt8(eGMCommands.ReloadObjects)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteReloadObjects de Protocol.bas")
End Sub

''
' Writes the "Restart" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRestart()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Restart" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.GMCommands)
    Call Writer.WriteInt8(eGMCommands.Restart)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteRestart de Protocol.bas")
End Sub

''
' Writes the "ResetAutoUpdate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResetAutoUpdate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ResetAutoUpdate" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.GMCommands)
    Call Writer.WriteInt8(eGMCommands.ResetAutoUpdate)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteResetAutoUpdate de Protocol.bas")
End Sub

''
' Writes the "ChatColor" message to the outgoing data buffer.
'
' @param    r The red component of the new chat color.
' @param    g The green component of the new chat color.
' @param    b The blue component of the new chat color.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChatColor(ByVal r As Byte, ByVal g As Byte, ByVal b As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChatColor" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.ChatColor)
        
        Call Writer.WriteInt8(r)
        Call Writer.WriteInt8(g)
        Call Writer.WriteInt8(b)
        
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteChatColor de Protocol.bas")
End Sub

''
' Writes the "Ignored" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteIgnored()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Ignored" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.GMCommands)
    Call Writer.WriteInt8(eGMCommands.Ignored)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteIgnored de Protocol.bas")
End Sub

''
' Writes the "CheckSlot" message to the outgoing data buffer.
'
' @param    UserName    The name of the char whose slot will be checked.
' @param    slot        The slot to be checked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCheckSlot(ByVal UserName As String, ByVal Slot As Byte)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Writes the "CheckSlot" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    UserName = complexNameToSimple(UserName, True)
    
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.CheckSlot)
        Call Writer.WriteString8(UserName)
        Call Writer.WriteInt8(Slot)
        
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteCheckSlot de Protocol.bas")
End Sub

''
' Writes the "Ping" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePing()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 26/01/2007
    'Writes the "Ping" message to the outgoing data buffer
    '***************************************************
    
    'Prevent the timer from being cut
    On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.Ping)
    Call Writer.WriteInt32(timeGetTime)
        
    Call Send(True)

    Exit Sub
  
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WritePing de Protocol.bas")

End Sub

''
' Writes the "SetIniVar" message to the outgoing data buffer.
'
' @param    sLlave the name of the key which contains the value to edit
' @param    sClave the name of the value to edit
' @param    sValor the new value to set to sClave
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetIniVar(ByRef sLlave As String, ByRef sClave As String, ByRef sValor As String)
'***************************************************
'Author: Brian Chaia (BrianPr)
'Last Modification: 21/06/2009
'Writes the "SetIniVar" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.SetIniVar)
        
        Call Writer.WriteString8(sLlave)
        Call Writer.WriteString8(sClave)
        Call Writer.WriteString8(sValor)
        
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteSetIniVar de Protocol.bas")
End Sub

''
' Writes the "CreatePretorianClan" message to the outgoing data buffer.
'
' @param    Map         The map in which create the pretorian clan.
' @param    X           The x pos where the king is settled.
' @param    Y           The y pos where the king is settled.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreatePretorianClan(ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: ZaMa
'Last Modification: 29/10/2010
'Writes the "CreatePretorianClan" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.CreatePretorianClan)
        Call Writer.WriteInt16(Map)
        Call Writer.WriteInt8(X)
        Call Writer.WriteInt8(Y)
        
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteCreatePretorianClan de Protocol.bas")
End Sub

''
' Writes the "DeletePretorianClan" message to the outgoing data buffer.
'
' @param    Map         The map which contains the pretorian clan to be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDeletePretorianClan(ByVal Map As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 29/10/2010
'Writes the "DeletePretorianClan" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.RemovePretorianClan)
        Call Writer.WriteInt16(Map)
        
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteDeletePretorianClan de Protocol.bas")
End Sub

''
' Writes the "MapMessage" message to the outgoing data buffer.
'
' @param    Dialog The new dialog of the NPC.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetDialog(ByVal dialog As String)
'***************************************************
'Author: Amraphen
'Last Modification: 18/11/2010
'Writes the "SetDialog" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.SetDialog)
        
        Call Writer.WriteString8(dialog)
    
    Call Send(False)
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteSetDialog de Protocol.bas")
End Sub

''
' Writes the "Impersonate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteImpersonate()
'***************************************************
'Author: ZaMa
'Last Modification: 20/11/2010
'Writes the "Impersonate" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.GMCommands)
    Call Writer.WriteInt8(eGMCommands.Impersonate)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteImpersonate de Protocol.bas")
End Sub

''
' Writes the "Imitate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteImitate()
'***************************************************
'Author: ZaMa
'Last Modification: 20/11/2010
'Writes the "Imitate" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.GMCommands)
    Call Writer.WriteInt8(eGMCommands.Imitate)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteImitate de Protocol.bas")
End Sub

''
' Writes the "RecordAddObs" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRecordAddObs(ByVal RecordIndex As Byte, ByVal Observation As String)
'***************************************************
'Author: Amraphen
'Last Modification: 29/11/2010
'Writes the "RecordAddObs" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.RecordAddObs)
        
        Call Writer.WriteInt8(RecordIndex)
        Call Writer.WriteString8(Observation)
    
    Call Send(False)
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteRecordAddObs de Protocol.bas")
End Sub

''
' Writes the "RecordAdd" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRecordAdd(ByVal Nickname As String, ByVal Reason As String)
'***************************************************
'Author: Amraphen
'Last Modification: 29/11/2010
'Writes the "RecordAdd" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.RecordAdd)
        
        Call Writer.WriteString8(Nickname)
        Call Writer.WriteString8(Reason)

      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteRecordAdd de Protocol.bas")
End Sub

''
' Writes the "RecordRemove" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRecordRemove(ByVal RecordIndex As Byte)
'***************************************************
'Author: Amraphen
'Last Modification: 29/11/2010
'Writes the "RecordRemove" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.RecordRemove)
        
        Call Writer.WriteInt8(RecordIndex)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteRecordRemove de Protocol.bas")
End Sub

''
' Writes the "RecordListRequest" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRecordListRequest()
'***************************************************
'Author: Amraphen
'Last Modification: 29/11/2010
'Writes the "RecordListRequest" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.GMCommands)
    Call Writer.WriteInt8(eGMCommands.RecordListRequest)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteRecordListRequest de Protocol.bas")
End Sub

''
' Writes the "RecordDetailsRequest" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRecordDetailsRequest(ByVal RecordIndex As Byte)
'***************************************************
'Author: Amraphen
'Last Modification: 29/11/2010
'Writes the "RecordDetailsRequest" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.RecordDetailsRequest)
        
        Call Writer.WriteInt8(RecordIndex)
    
    Call Send(False)
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteRecordDetailsRequest de Protocol.bas")
End Sub


''
' Writes the "Moveitem" message to the outgoing data buffer.
'
Public Sub WriteMoveItem(ByVal originalSlot As Integer, ByVal newSlot As Integer, ByVal moveType As eMoveType)
'***************************************************
'Author: Budi
'Last Modification: 05/01/2011
'Writes the "MoveItem" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ClientPacketID.moveItem)
        Call Writer.WriteInt8(originalSlot)
        Call Writer.WriteInt8(newSlot)
        Call Writer.WriteInt8(moveType)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteMoveItem de Protocol.bas")
End Sub

''
' Writes the "PMSend" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePMSend(ByVal UserName As String, ByVal Message As String)
'***************************************************
'Author: Amraphen
'Last Modification: 05/08/2011
'Writes the "PMSend" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    UserName = complexNameToSimple(UserName, True)
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.PMSend)

        If (InStrB(UserName, "+") <> 0) Then
            UserName = Replace$(UserName, "+", " ")
        End If
        
        UserName = UCase$(UserName)
        
        Call Writer.WriteString8(UserName)
        Call Writer.WriteString8(Message)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WritePMSend de Protocol.bas")
End Sub

''
' Writes the "PMList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePMList()
'***************************************************
'Author: Amraphen
'Last Modification: 05/08/2011
'Writes the "PMList" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  

    Call Writer.WriteInt8(ClientPacketID.PMList)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WritePMList de Protocol.bas")
End Sub

''
' Writes the "PMDeleteList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePMDeleteList()
'***************************************************
'Author: Amraphen
'Last Modification: 05/08/2011
'Writes the "PMDeleteList" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  

    Call Writer.WriteInt8(ClientPacketID.PMDeleteList)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WritePMDeleteList de Protocol.bas")
End Sub

''
' Writes the "PMDeleteUser" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePMDeleteUser(ByVal UserName As String, ByVal PMIndex As Byte)
'***************************************************
'Author: Amraphen
'Last Modification: 05/08/2011
'Writes the "PMDeleteUser" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    UserName = complexNameToSimple(UserName, True)
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.PMDeleteUser)
        
        If (InStrB(UserName, "+") <> 0) Then
            UserName = Replace$(UserName, "+", " ")
        End If
        
        UserName = UCase$(UserName)

        Call Writer.WriteString8(UserName)
        Call Writer.WriteInt8(PMIndex)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WritePMDeleteUser de Protocol.bas")
End Sub

''
' Writes the "PMListUser" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePMListUser(ByVal UserName As String, ByVal PMIndex As Byte)
'***************************************************
'Author: Amraphen
'Last Modification: 05/08/2011
'Writes the "PMListUser" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
    UserName = complexNameToSimple(UserName, True)
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.PMListUser)
        
        If (InStrB(UserName, "+") <> 0) Then
            UserName = Replace$(UserName, "+", " ")
        End If
        
        UserName = UCase$(UserName)

        Call Writer.WriteString8(UserName)
        Call Writer.WriteInt8(PMIndex)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WritePMListUser de Protocol.bas")
End Sub

''
' Writes the "MenuAction" message to the outgoing data buffer.
'
Public Sub WriteMenuAction(ByVal iAction As Integer, ByVal Slot As Byte)
'***************************************************
'Author: ZaMa
'Last Modification: 22/03/2012
'Writes the "MenuAction" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ClientPacketID.MenuAction)
        Call Writer.WriteInt16(iAction)
        Call Writer.WriteInt8(Slot)
    
    Call Send(False)
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteMenuAction de Protocol.bas")
End Sub

Public Sub WriteTournamentParticipate()
'***************************************************
'Author: ZaMa
'Last Modification: 31/05/2012
'
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ClientPacketID.Participar)

      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteTournamentParticipate de Protocol.bas")
End Sub
 
Public Sub WriteRequestTournamentCompetitors()
'***************************************************
'Author: ZaMa
'Last Modification: 31/05/2012
'
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.RequestTournamentCompetitors)
    
    Call Send(False)
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteRequestTournamentCompetitors de Protocol.bas")
End Sub
 
Public Sub WriteTournamentDisqualify(ByRef Participante As String)
'***************************************************
'Author: ZaMa
'Last Modification: 31/05/2012
'
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.Descalificar)
        Call Writer.WriteString8(Participante)
    
    Call Send(False)
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteTournamentDisqualify de Protocol.bas")
End Sub
 
Public Sub WriteTournamentFight(ByRef Participante1 As String, ByRef Participante2 As String, ByVal ArenaIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 31/05/2012
'
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.Pelea)
        Call Writer.WriteString8(Participante1)
        Call Writer.WriteString8(Participante2)
        Call Writer.WriteInt16(ArenaIndex)
    
    Call Send(False)
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteTournamentFight de Protocol.bas")
End Sub
 
Public Sub WriteTorunamentCancel()
'***************************************************
'Author: ZaMa
'Last Modification: 31/05/2012
'
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.CerrarTorneo)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteTorunamentCancel de Protocol.bas")
End Sub
 
Public Sub WriteTorunamentBegin()
'***************************************************
'Author: ZaMa
'Last Modification: 31/05/2012
'
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.IniciarTorneo)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteTorunamentBegin de Protocol.bas")
End Sub

Public Sub WriteTorunamentEdit_Flags(ByVal EditOption As Byte, ByVal NewValue As Byte)
'***************************************************
'Author: ZaMa
'Last Modification: 01/06/2012
'Modify byte-base flags of tournament configuration
'***************************************************
On Error GoTo ErrHandler
  

'TODO_TORNEO:
'ieMaxCompetitor
'ieMinLevel
'ieMaxLevel
'ieNumRoundsToWin
'ieKillAfterLoose

        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.TorunamentEdit)
        Call Writer.WriteInt8(EditOption)
        Call Writer.WriteInt8(NewValue)
    
    Call Send(False)
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteTorunamentEdit_Flags de Protocol.bas")
End Sub

Public Sub WriteTorunamentEdit_SaveConfig()
'***************************************************
'Author: ZaMa
'Last Modification: 01/06/2012
'Saves current tournament configuration
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.TorunamentEdit)
        Call Writer.WriteInt8(eTournamentEdit.ieSaveConfig)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteTorunamentEdit_SaveConfig de Protocol.bas")
End Sub

Public Sub WriteTorunamentEdit_RequiredGold(ByVal gold As Long)
'***************************************************
'Author: ZaMa
'Last Modification: 01/06/2012
'Updates current required gold
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.TorunamentEdit)
        Call Writer.WriteInt8(eTournamentEdit.ieSaveConfig)
        Call Writer.WriteInt32(gold)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteTorunamentEdit_RequiredGold de Protocol.bas")
End Sub

Public Sub WriteTorunamentEdit_ForbiddenItems(ByVal NumItems As Byte, ByRef ItemList() As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 01/06/2012
'Updates current forbidden items
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.TorunamentEdit)
        Call Writer.WriteInt8(eTournamentEdit.ieForbiddenItems)
        
        Call Writer.WriteInt8(NumItems)
        
        Dim lItem As Long
        For lItem = 1 To NumItems
            Call Writer.WriteInt16(ItemList(lItem))
        Next lItem
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteTorunamentEdit_ForbiddenItems de Protocol.bas")
End Sub

Public Sub WriteTorunamentEdit_PermitedClass(ByVal NumClases As Byte, ByRef ClassList() As Byte)
'***************************************************
'Author: ZaMa
'Last Modification: 01/06/2012
'Updates current Permited Classes
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.TorunamentEdit)
        Call Writer.WriteInt8(eTournamentEdit.iePermitedClass)
        
        Call Writer.WriteInt8(NumClases)
        
        Dim lClass As Long
        For lClass = 1 To NumClases
            Call Writer.WriteInt8(ClassList(lClass))
        Next lClass
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteTorunamentEdit_PermitedClass de Protocol.bas")
End Sub

Public Sub WriteTorunamentEdit_Map(ByVal EditOption As Byte, ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: ZaMa
'Last Modification: 01/06/2012
'Updates current maps
'***************************************************
On Error GoTo ErrHandler
  

'TODO_TORNEO: EditOption
'ieWaitingMap
'ieFinalMap

        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.TorunamentEdit)
        Call Writer.WriteInt8(EditOption)
        
        Call Writer.WriteInt16(Map)
        Call Writer.WriteInt8(X)
        Call Writer.WriteInt8(Y)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteTorunamentEdit_Map de Protocol.bas")
End Sub

Public Sub WriteTorunamentEdit_Arena(ByVal ArenaIndex As Byte, ByVal Map As Integer, _
    ByVal User1_X As Byte, ByVal User1_Y As Byte, ByVal User2_X As Byte, ByVal User2_Y As Byte)
'***************************************************
'Author: ZaMa
'Last Modification: 01/06/2012
'Updates current maps
'***************************************************
On Error GoTo ErrHandler
  

'TODO_TORNEO: EditOption
'ieWaitingMap
'ieFinalMap

        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.TorunamentEdit)
        Call Writer.WriteInt8(eTournamentEdit.ieArenaPosition)
        
        Call Writer.WriteInt8(ArenaIndex)
        Call Writer.WriteInt16(Map)
        Call Writer.WriteInt8(User1_X)
        Call Writer.WriteInt8(User1_Y)
        Call Writer.WriteInt8(User2_X)
        Call Writer.WriteInt8(User2_Y)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteTorunamentEdit_Arena de Protocol.bas")
End Sub


Public Sub WriteRequestTournamentConfig()
'***************************************************
'Author: ZaMa
'Last Modification: 07/06/2012
'
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.RequestTournamentConfig)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteRequestTournamentConfig de Protocol.bas")
End Sub

''
' Writes the "HigherAdminsMessage" message to the outgoing data buffer.
'
' @param    message The message to be sent to the other higher admins online.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteHigherAdminsMessage(ByVal Message As String)
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 03/30/12
'Writes the "HigherAdminsMessage" message to the outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.HigherAdminsMessage)
        Call Writer.WriteString8(Message)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteHigherAdminsMessage de Protocol.bas")
End Sub
Public Sub WriteAccountLogin()
On Error GoTo ErrHandler
  
 
'
' @ Conecta la cuenta


        
        Call Writer.WriteInt8(ClientPacketID.AccountLogin)
        Call Writer.WriteString8(Acc_Data.Acc_Name)
    
        Call Writer.WriteString8(Acc_Data.Acc_Password)

        Call Writer.WriteString8(AccountGetToken())
        Call Writer.WriteString8(RandomClientToken)
        
        Call Writer.WriteInt8(App.Major)
        Call Writer.WriteInt8(App.Minor)
        Call Writer.WriteInt8(App.Revision)
        
    #If EnableSecurity Then
        Call Writer.WriteString8(MD5HushYo)
    #End If
    
    Call Send(False)
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteAccountLogin de Protocol.bas")
End Sub
 
Public Sub WriteAccountLoginChar()
On Error GoTo ErrHandler
  
 
'
' @ Conecta un personaje de la cuenta.
 

         Call Writer.WriteInt8(ClientPacketID.AccountLoginChar)
         Call Writer.WriteInt8(Acc_Data.Acc_Char_Selected)
         Call Writer.WriteInt8(App.Major)
         Call Writer.WriteInt8(App.Minor)
         Call Writer.WriteInt8(App.Revision)
         Call Writer.WriteString8(AccountGetToken())
         Call Writer.WriteString8(RandomClientToken)
    
    Call Send(False)
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteAccountLoginChar de Protocol.bas")
End Sub
 
Public Sub WriteAccountCreateChar()
On Error GoTo ErrHandler
'
' @ Crea un personaje .
        Call Writer.WriteInt8(ClientPacketID.AccountCreateChar)
       
        Call Writer.WriteString8(Trim(UserName))
        Call Writer.WriteInt8(AccountConnecting.UserRace)
        Call Writer.WriteInt8(AccountConnecting.UserGender)
        Call Writer.WriteInt8(AccountConnecting.UserClass)
        Call Writer.WriteInt16(UserHead)
        Call Writer.WriteInt8(UserHogar)
        Call Writer.WriteInt8(App.Major)
        Call Writer.WriteInt8(App.Minor)
        Call Writer.WriteInt8(App.Revision)
        Call Writer.WriteString8(AccountGetToken())
        Call Writer.WriteString8(RandomClientToken)
        
        Acc_Data.Acc_Waiting_CharName = UserName
    
    Call Send(False)
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteAccountCreateChar de Protocol.bas")
End Sub
 
Public Sub WriteAccountCreate()
On Error GoTo ErrHandler
  
 
'
' @ Crea la cuenta.
 

        Call Writer.WriteInt8(ClientPacketID.AccountCreate)
        Call Writer.WriteString8(Acc_Data.Acc_Name)

        Call Writer.WriteString8(Acc_Data.Acc_Password)

        Call Writer.WriteString8(Acc_Data.Acc_Email)
        Call Writer.WriteString8(Acc_Data.Acc_Pregunta)
        Call Writer.WriteString8(Acc_Data.Acc_Respuesta)
        Call Writer.WriteInt8(App.Major)
        Call Writer.WriteInt8(App.Minor)
        Call Writer.WriteInt8(App.Revision)
        Call Writer.WriteString8(RandomClientToken)
            
    Call Send(False)
    

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteAccountCreate de Protocol.bas")
End Sub
 
Public Sub WriteAccountDeleteChar()
On Error GoTo ErrHandler
  
 
'
' @ Borra un personaje.
 

         Call Writer.WriteInt8(ClientPacketID.AccountDeleteChar)
         Call Writer.WriteInt8(Acc_Data.Acc_Char_Selected)
         Call Writer.WriteInt8(App.Major)
         Call Writer.WriteInt8(App.Minor)
         Call Writer.WriteInt8(App.Revision)
         Call Writer.WriteString8(AccountGetToken())
         Call Writer.WriteString8(RandomClientToken)
         Call Writer.WriteString8(Acc_Data.Acc_Respuesta)
    
    Call Send(False)
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteAccountDeleteChar de Protocol.bas")
End Sub
 
Public Sub WriteAccountRecover()
On Error GoTo ErrHandler
  
 
'
' @ Recupera cuenta *
 

        Call Writer.WriteInt8(ClientPacketID.AccountRecover)
        Call Writer.WriteString8(Acc_Data.Acc_Name)
        Call Writer.WriteString8(Acc_Data.Acc_Token)
        
        Call Writer.WriteInt8(App.Major)
        Call Writer.WriteInt8(App.Minor)
        Call Writer.WriteInt8(App.Revision)
        
        Call Writer.WriteString8(RandomClientToken)
    
    Call Send(False)
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteAccountRecover de Protocol.bas")
End Sub

Public Sub WriteAccountChangePassword()
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ClientPacketID.AccountChangePassword)

        Call Writer.WriteString8(Acc_Data.Acc_Password)
        Call Writer.WriteString8(Acc_Data.Acc_New_Password)
        
        Call Writer.WriteString8(AccountGetToken())
        Call Writer.WriteString8(RandomClientToken)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteAccountChangePassword de Protocol.bas")
End Sub


Public Sub WriteChatDesafio(ByVal Texto As String)
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ClientPacketID.Chat_desafio)
        Call Writer.WriteString8(Texto)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteChatDesafio de Protocol.bas")
End Sub
  
Public Sub WriteCancel_desafio()
On Error GoTo ErrHandler
    Call Writer.WriteInt8(ClientPacketID.Cancel_desafio)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteCancel_desafio de Protocol.bas")
End Sub

Public Sub WriteAccept_desafio()

On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ClientPacketID.Accept_desafio)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteAccept_desafio de Protocol.bas")
End Sub

Public Sub WriteEnviardatos_desafio(ByVal Oro As Long, ByVal Dead As Byte, ByVal Time As Byte, ByVal StarTime As Byte, _
ByVal Mapa As Byte, ByVal Invisibilidad As Byte, ByVal Resucitar As Byte, ByVal Elementales As Byte)
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ClientPacketID.Enviardatos_desafio)
        
        Call Writer.WriteInt32(Oro)
        Call Writer.WriteInt8(Dead)
        Call Writer.WriteInt8(Time)
        Call Writer.WriteInt8(StarTime)
        Call Writer.WriteInt8(Mapa)
        Call Writer.WriteInt8(Invisibilidad)
        Call Writer.WriteInt8(Resucitar)
        Call Writer.WriteInt8(Elementales)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteEnviardatos_desafio de Protocol.bas")
End Sub

Public Sub WriteGetPunishmentTypeList(ByRef userToPunish As String, ByVal punishmentType As Byte)
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.GetPunishmentTypeList)
        Call Writer.WriteInt8(punishmentType)
        Call Writer.WriteString8(userToPunish)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteGetPunishmentTypeList de Protocol.bas")
End Sub

Public Sub WriteCancelarDuelo()
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.CancelarElDuelo)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteCancelarDuelo de Protocol.bas")
End Sub

Public Sub WriteRetar(ByVal vs As Byte, ByVal Oro As Long, ByVal Drop As Boolean, ByVal Nick1 As String, _
                        Optional ByVal Nick2 As String, Optional ByVal Nick3 As String, _
                        Optional ByVal Resucitar As Boolean, Optional ByVal Nick4 As String, _
                        Optional ByVal Nick5 As String, Optional ByVal Nick6 As String, Optional ByVal Nick7 As String)
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.Retar)
    Call Writer.WriteInt8(vs)
    Call Writer.WriteInt32(Oro)
    Call Writer.WriteBool(Drop)
    Call Writer.WriteString8(Nick1)
    If Not vs = 1 Then
        Call Writer.WriteString8(Nick2)
        Call Writer.WriteString8(Nick3)
        Call Writer.WriteBool(Resucitar)
        If Not vs = 2 Then
            Call Writer.WriteString8(Nick4)
            Call Writer.WriteString8(Nick5)
            If Not vs = 3 Then
                Call Writer.WriteString8(Nick6)
                Call Writer.WriteString8(Nick7)
            End If
        End If
    End If
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteRetar de Protocol.bas")
End Sub

Public Sub WriteDuelos()
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.Duelos)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteDuelos de Protocol.bas")
End Sub

Public Sub WriteAceptarDuelo()
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.AceptarDuelo)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteAceptarDuelo de Protocol.bas")
End Sub

Public Sub WriteRechazarDuelo()
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.RechazarDuelo)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteRechazarDuelo de Protocol.bas")
End Sub

Public Sub WriteDueloPublico()
'---------------------------------------------------------------------------------------
' Module    : Protocol
' Author    : Anagrama
' Date      : 20/08/2016
' Purpose   : Ingresa a la lista de duelos publicos.
'---------------------------------------------------------------------------------------
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ClientPacketID.DueloPublico)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteDueloPublico de Protocol.bas")
End Sub

Public Sub WriteCancelarEspera()
'---------------------------------------------------------------------------------------
' Module    : Protocol
' Author    : Anagrama
' Date      : 20/08/2016
' Purpose   : Cancela la espera de duelo.
'---------------------------------------------------------------------------------------
On Error GoTo ErrHandler

        Call Writer.WriteInt8(ClientPacketID.CancelarEspera)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteCancelarEspera de Protocol.bas")
End Sub

Public Sub WriteAccBankExtractGold(ByVal Amount As Long)
'***************************************************
'Author: Anagrama
'Last Modification: 21/08/06
'Solicita extraer Amount cantidad de oro de la boveda de cuenta
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ClientPacketID.AccBankExtractGold)
        
        Call Writer.WriteInt32(Amount)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteAccBankExtractGold de Protocol.bas")
End Sub


Public Sub WriteAccBankDepositGold(ByVal Amount As Long)
'***************************************************
'Author: Anagrama
'Last Modification: 21/08/06
'Solicita depositar Amount cantidad de oro de la boveda de cuenta
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ClientPacketID.AccBankDepositGold)
        
        Call Writer.WriteInt32(Amount)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteAccBankDepositGold de Protocol.bas")
End Sub

Public Sub WriteAccBankExtractItem(ByVal Slot As Byte, ByVal Amount As Integer)
'***************************************************
'Author: Anagrama
'Last Modification: 21/08/06
'Solicita extraer uno o mas items de la boveda de cuenta
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ClientPacketID.AccBankExtractItem)
        
        Call Writer.WriteInt8(Slot)
        Call Writer.WriteInt16(Amount)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteAccBankExtractItem de Protocol.bas")
End Sub

Public Sub WriteAccBankDeposit(ByVal Slot As Byte, ByVal Amount As Integer)
'***************************************************
'Author: Anagrama
'Last Modification: 21/08/06
'Solicita depositar uno o mas items de la boveda de cuenta
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ClientPacketID.AccBankDepositItem)
        
        Call Writer.WriteInt8(Slot)
        Call Writer.WriteInt16(Amount)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteAccBankDeposit de Protocol.bas")
End Sub

Public Sub WriteAccBankStart(ByVal Password As String)
'***************************************************
'Author: Anagrama
'Last Modification: 21/08/06
'Solicita abrir la boveda de cuenta y envia la contraseña
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.AccBankStart)
    Call Writer.WriteString8(Password)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteAccBankStart de Protocol.bas")
End Sub

Public Sub WriteAccBankEnd()
'***************************************************
'Author: Anagrama
'Last Modification: 21/08/06
'Envia la terminacion la transaccion con la boveda de cuenta
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.AccBankEnd)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteAccBankEnd de Protocol.bas")
End Sub

Public Sub WriteAccBankChangePass(ByVal Token As String, ByVal Password As String)
'***************************************************
'Author: Anagrama
'Last Modification: 21/08/06
'Cambia la contraseña de la boveda
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.AccBankChangePass)
    Call Writer.WriteString8(Token)
    Call Writer.WriteString8(Password)
      
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteAccBankChangePass de Protocol.bas")
End Sub

Public Sub WriteChangeMapInfoNoInmo(ByVal NoInmo As Boolean)
'***************************************************
'Author: Anagrama
'Last Modification: 09/09/2016
'Envia un cambio a la posibilidad de usar inmovilizar en el mapa.
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.ChangeMapInfoNoInmo)
        
        Call Writer.WriteBool(NoInmo)
    
    Call Send(False)
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteChangeMapInfoNoInmo de Protocol.bas")
End Sub

Public Sub WriteChangeMapInfoMismoBando(ByVal MismoBando As Boolean)
'***************************************************
'Author: Anagrama
'Last Modification: 09/09/2016
'Envia un cambio a la posibilidad de que ciudadanos no ataquen armadas y criminales no ataquen caos en el mapa.
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.ChangeMapInfoMismoBando)
        
        Call Writer.WriteBool(MismoBando)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteChangeMapInfoMismoBando de Protocol.bas")
End Sub

Public Sub WriteSpawnBoss(ByVal BossID As String)
'***************************************************
'Author: Anagrama
'Last Modification: 09/12/2016
'Envia el número del boss a invocar.
'***************************************************
On Error GoTo ErrHandler
  
    If Val(BossID) < 1 Or Val(BossID) > 255 Then Exit Sub
    
        Call Writer.WriteInt8(ClientPacketID.GMCommands)
        Call Writer.WriteInt8(eGMCommands.SpawnBoss)
        Call Writer.WriteInt8(Val(BossID))
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteSpawnBoss de Protocol.bas")
End Sub

Public Sub WriteForgive(ByRef TargetUserName As String)
On Error GoTo ErrHandler

    Call Writer.WriteInt8(ClientPacketID.GMCommands)
    Call Writer.WriteInt8(eGMCommands.Forgive)
    Call Writer.WriteString8(TargetUserName)
    
    Call Send(False)
    
    Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteSpawnBoss de Protocol.bas")
End Sub

Public Sub WriteCraftItem(ByVal CraftingGroup As Byte, ByVal GroupRecipeIndex As Integer, ByVal FromMacro As Boolean)
'***************************************************
'Author: Anagrama
'Last Modification: 07/04/2017
'Envía la estación y la receta a construir.
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ClientPacketID.CraftItem)
        Call Writer.WriteInt8(CraftingGroup)
        Call Writer.WriteInt(GroupRecipeIndex)
        Call Writer.WriteBool(FromMacro)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteCraftItem de Protocol.bas")
End Sub

Public Sub WriteSelectPet(ByVal PetIndex As Byte)
'***************************************************
'Author: Anagrama
'Last Modification: 12/08/2017
'Envía la mascota a seleccionar.
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.SelectPet)
    Call Writer.WriteInt8(PetIndex)
        
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteSelectPet de Protocol.bas")
End Sub

Public Sub WriteRequestPetSelection()
'***************************************************
'Author: Anagrama
'Last Modification: 12/08/2017
'Solicita el formulario de selección de mascotas.
'***************************************************
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.RequestPetSelection)
        
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteRequestPetSelection de Protocol.bas")
End Sub


Public Sub WriteAssignMastery(ByVal MasteryGroup As Integer, ByVal MasteryId As Integer)
On Error GoTo ErrHandler
  
    Call Writer.WriteInt8(ClientPacketID.MasteryAssign)
    Call Writer.WriteInt16(MasteryGroup)
    Call Writer.WriteInt16(MasteryId)
        
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteAssignMastery de Protocol.bas")
End Sub

Public Sub InitProtocol()
    LAST_CLIENT_PACKET_ID = ClientPacketID.LastClientPacketId - 1
End Sub


Public Sub WriteGuildCreate(ByVal GuildName As String)
'***************************************************
On Error GoTo ErrHandler

    Call Writer.WriteInt8(ClientPacketID.GuildCreate)
    Call Writer.WriteString16(GuildName)
    
    Call Send(False)
    
    Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteGuildCreation de Protocol.bas")
End Sub

Public Sub WriteGuildQuest(ByVal QuestId As Integer)
'***************************************************
On Error GoTo ErrHandler

    Call Writer.WriteInt8(ClientPacketID.GuildQuest)
    
    Call Writer.WriteInt(QuestId)
    
    Call Send(False)
    
    Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteGuildQuest de Protocol.bas")
End Sub

Public Sub WriteGuildQuestAddObject(ByVal InventorySlot As Integer, ByVal Quantity As Long)
'***************************************************
On Error GoTo ErrHandler

    Call Writer.WriteInt8(ClientPacketID.GuildQuestAddObject)
    
    Call Writer.WriteInt(InventorySlot)
    Call Writer.WriteInt(Quantity)
    
    Call Send(False)
    
    Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteGuildQuestAddObject de Protocol.bas")
End Sub

Public Sub WriteGuildInvitationResponse(ByVal GuildId As Long, ByVal InvitationIndex As Long, ByVal Accepted As Boolean)
On Error GoTo ErrHandler
    
    Call Writer.WriteInt8(ClientPacketID.GuildUserInvitationResponse)
    Call Writer.WriteInt(GuildId)
    Call Writer.WriteInt(InvitationIndex)
    Call Writer.WriteBool(Accepted)

    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteGuildInvitationResponse de Protocol.bas")
End Sub

Public Sub WriteGuildMember(ByVal GuildId As Long, ByVal KickedUserId As Long, ByVal Action As Integer, Optional ByVal NamePlayer As String, Optional ByVal Accepted As Boolean)
On Error GoTo ErrHandler
    
    'Member Action: Accept/Reject Invitation = 1   Kick Member = 2  send invitation =3
    
    Call Writer.WriteInt8(ClientPacketID.GuildMember)
    
    Call Writer.WriteInt8(Action)
    Call Writer.WriteInt(KickedUserId)
    Call Writer.WriteInt16(GuildId)
    If Action = eMemberAction.SendInvitation Then
        Call Writer.WriteString16(NamePlayer)
    End If
    Call Writer.WriteBool(Accepted)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteGuildMember de Protocol.bas")
End Sub

Public Sub WriteGuildExchange(ByVal ExchangeType As Byte, ByVal ExchangeAction As Byte, ByVal Quantity As Long, Optional ByVal Slot As Integer = 0, Optional ByVal Box As Integer = 0)
On Error GoTo ErrHandler

    'ExchangeType: IsGold = 1 IsObject =2
    'ExchangeAction: Withdraw = 1 Deposit=2
    
    Call Writer.WriteInt8(ClientPacketID.GuildExchange)
    
    Call Writer.WriteInt8(ExchangeType)
    Call Writer.WriteInt8(ExchangeAction)
    Call Writer.WriteInt32(Quantity)
    Call Writer.WriteInt16(Slot)
    Call Writer.WriteInt16(Box)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteGuildExchange de Protocol.bas")
End Sub

Public Sub WriteGuildRole_Create(ByVal RoleAction As Integer, ByVal RoleId As Integer, ByVal RoleName As String, ByRef Permissions() As String)
On Error GoTo ErrHandler

    ' Role: Assign = 1  Create = 2  Delete = 3
    Dim I As Integer, QtyPermissions As Integer
    
    If ((Not Permissions) = -1) Then
        QtyPermissions = 0
    Else
        QtyPermissions = UBound(Permissions)
    End If
    
    Call Writer.WriteInt8(ClientPacketID.GuildRole)

    Call Writer.WriteInt8(RoleAction)
    Call Writer.WriteInt32(RoleId)
    Call Writer.WriteString8(RoleName)
    Call Writer.WriteInt8(QtyPermissions)
    
    For I = 1 To QtyPermissions
        Call Writer.WriteString8(Permissions(I))
    Next I
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteGuildRole_Create de Protocol.bas")
End Sub

Public Sub WriteGuildRole_Assign(ByVal RoleAction As Integer, ByVal RoleId As Integer, ByVal TargetUserId As Long)
On Error GoTo ErrHandler

    ' Role: Assign = 1  Create = 2  Delete = 3
    Call Writer.WriteInt8(ClientPacketID.GuildRole)
    Call Writer.WriteInt8(RoleAction)
    Call Writer.WriteInt32(RoleId)
    Call Writer.WriteInt(TargetUserId)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteGuildRole_Assign de Protocol.bas")
End Sub

Public Sub WriteGuildRole_Delete(ByVal RoleId As Integer)
On Error GoTo ErrHandler

    ' Role: Assign = 1  Create = 2  Delete = 3
    Call Writer.WriteInt8(ClientPacketID.GuildRole)
    Call Writer.WriteInt8(eRoleAction.Delete)
    Call Writer.WriteInt32(RoleId)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteGuildRole_Assign de Protocol.bas")
End Sub


Public Sub WriteGuildUpgrade(ByVal UpgradeNumber As Integer)
On Error GoTo ErrHandler
    
    Call Writer.WriteInt8(ClientPacketID.GuildUpgrade)
    
    Call Writer.WriteInt8(UpgradeNumber)
    
    Call Send(False)
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteGuildUpgrade de Protocol.bas")
End Sub

Private Sub HandleGuildInfo()

    Dim IdGuild As Integer
    Dim IdCurrentQuest As Integer
    Dim Alignment As Byte
    Dim Status As Byte
    Dim MemberCount As Byte
    Dim Name As String
    Dim Description As String
    Dim IdLeader As Long
    Dim IdRightHand As Long
    Dim CreationTime As Date
    Dim QuestStartedDate As Date
    Dim ContributionEarned As Long
    Dim ContributionAvailable As Long
    Dim BankGold As Long, MaxContribution As Long
    Dim IdRolOwn As Integer
    Dim MaxMemberQty As Byte, MaxRolesQty As Byte, MaxSlotBankQty As Byte, MaxBoxesQty As Byte
    Dim BankUpg As Boolean
    
    IdGuild = Reader.ReadInt32
    Name = Reader.ReadString16
    Description = Reader.ReadString16
    Alignment = Reader.ReadInt8
    CreationTime = Reader.ReadString16
    Status = Reader.ReadInt8
    IdLeader = Reader.ReadInt32
    IdRightHand = Reader.ReadInt32
    MemberCount = Reader.ReadInt8
    IdCurrentQuest = Reader.ReadInt16
    QuestStartedDate = Reader.ReadString16
    ContributionEarned = Reader.ReadInt32
    ContributionAvailable = Reader.ReadInt32
    BankGold = Reader.ReadInt32
    IdRolOwn = Reader.ReadInt32
    MaxMemberQty = Reader.ReadInt8
    MaxRolesQty = Reader.ReadInt8
    MaxSlotBankQty = Reader.ReadInt8
    MaxBoxesQty = Reader.ReadInt8
    MaxContribution = Reader.ReadInt32
    BankUpg = Reader.ReadBool
    
    With PlayerData.Guild
        .Name = Name
        .IdGuild = IdGuild
        .Alignment = Alignment
        .Description = Description
        .CreationTime = CreationTime
        .Status = Status
        .IdLeader = IdLeader
        .IdRightHand = IdRightHand
        .MemberCount = MemberCount
        .ContributionEarned = ContributionEarned
        .ContributionAvailable = ContributionAvailable
        .BankGold = BankGold
        .IdRolOwn = IdRolOwn
        .MaxMemberQty = MaxMemberQty
        .MaxRoles = MaxRolesQty
        .MaxSlotBank = MaxSlotBankQty
        .MaxBoxesBank = MaxBoxesQty
        .MaxContributionAvailable = MaxContribution
        .BankAvalaible = BankUpg
    End With
       
    
End Sub


Private Sub HandleGuildRolesList()
    Dim CountRoles As Integer
    Dim I As Integer
    Dim J As Integer
    Dim QtyPermissions As Integer
    
    CountRoles = Reader.ReadInt16()
    
    ReDim PlayerData.Guild.Roles(1 To CountRoles) As tGuildRole
    
    For I = 1 To CountRoles
        With PlayerData.Guild.Roles(I)
            .RoleId = Reader.ReadInt32()
            .RoleName = Reader.ReadString16()
            .DeleteEnabled = Reader.ReadBool()
            .UpdatePermissionsEnabled = Reader.ReadBool()
            .RenameEnabled = Reader.ReadBool()

            .PermissionsQty = Reader.ReadInt16()
            If .PermissionsQty > 0 Then
                ReDim .Permissions(1 To .PermissionsQty) As tGuildRolePermission
            
                For J = 1 To .PermissionsQty
                    .Permissions(J).PermissionId = Reader.ReadInt32()
                    .Permissions(J).Key = Reader.ReadString16()
                Next J
            End If
            
        End With
    Next I
        
    Call Guilds.UpdateFormInfo
End Sub

Private Sub HandleGuildMembersList()
    Dim CountMembers As Integer
    Dim I As Integer
    
    CountMembers = Reader.ReadInt8()
   
    ReDim PlayerData.Guild.Members(1 To CountMembers) As tGuildUserMember

    PlayerData.Guild.MemberCount = CountMembers
    
    For I = 1 To CountMembers
        With PlayerData.Guild.Members(I)
            .UserId = Reader.ReadInt32()
            .RoleId = Reader.ReadInt32()
            .UserName = Reader.ReadString16()
            .IsOnline = Reader.ReadBool()
        End With
    Next I
    
    Call Guilds.UpdateFormInfo

End Sub

Private Sub HandleGuildUpgradesAcquired()

    Dim CountUpgrade As Integer, ActualQtyUpgrade As Integer
    Dim Upgrades() As UpgradeType
    Dim I As Integer
    
    CountUpgrade = Reader.ReadInt8()
    
    If CountUpgrade = 0 Then Exit Sub
    
    ActualQtyUpgrade = GetQtyGuildUpgrades()
    
    If CountUpgrade = 1 And ActualQtyUpgrade > 0 Then
        CountUpgrade = 1 + ActualQtyUpgrade
    End If
    
    ReDim Preserve PlayerData.Guild.Upgrades(1 To CountUpgrade) As GuildUpgradeType

    With PlayerData.Guild
        For I = 1 + ActualQtyUpgrade To CountUpgrade
                .Upgrades(I).IdUpgrade = Reader.ReadInt32()
                .Upgrades(I).IsEnabled = Reader.ReadBool()
                .Upgrades(I).UpgradeBy = Reader.ReadString16()
                .Upgrades(I).UpgradeDate = Reader.ReadString16()
                .Upgrades(I).UpgradeLevel = Reader.ReadInt8()
            
        Next I
    End With
    
    If ActualQtyUpgrade = 0 Then
        frmGuildUpgrades.LoadUpgrades
    End If

    Call Guilds.UpdateFormInfo

End Sub

Private Sub HandleGuildMemberStatusChange()
    Dim IdUser As Long
    Dim TypeChanged As Byte
    Dim ValueChanged As Byte
    Dim ValueChangedLong As Long
    Dim ObjForm As Form
    Dim I As Integer


    IdUser = Reader.ReadInt32()
    TypeChanged = Reader.ReadInt8() '1= conexion, 2= rolchange, 3=others        TypeChanged = Reader.ReadInt8() '1= conexion, 2= rolchange, 3=others
    ValueChanged = Reader.ReadInt8()
    ValueChangedLong = Reader.ReadInt32()

    Select Case TypeChanged
        Case 1 ' online / offline member
            For I = 1 To UBound(PlayerData.Guild.Members)
                If (PlayerData.Guild.Members(I).UserId = IdUser) Then
                    PlayerData.Guild.Members(I).IsOnline = CBool(ValueChanged)
                    If frmGuildMembers.Visible Then
                        Call frmGuildMembers.MemberListUpdate(PlayerData.Guild.Members(I).UserId, PlayerData.Guild.Members(I).IsOnline)
                    End If
                End If
            Next I
        Case 2 ' role changed
            PlayerData.Guild.IdRolOwn = ValueChanged
            For I = 1 To UBound(PlayerData.Guild.Members)
                If PlayerData.Guild.Members(I).UserId = IdUser Then
                    PlayerData.Guild.Members(I).RoleId = ValueChanged
                End If
            Next I
            Call frmGuildRolesList.AddRoleButtonEnable
            Call frmGuildMembers.InviteButtonEnable
            Call UpdateFormInfo
        Case 3 ' GoldBank changed
            PlayerData.Guild.BankGold = ValueChangedLong

            If frmGuildBank.Visible Then
                frmGuildBank.LblBankGold.Caption = ValueChangedLong
            End If
    End Select
End Sub

Private Sub HandleGuildMemberKicked()
    PlayerData.Guild.IdGuild = 0
    If frmGuildMain.Visible Then
        Call frmGuildMain.CloseForm
    End If
    Unload frmGuildMain
    Call ResetGuildInfo
End Sub

Private Sub HandleGuildQuestUpdateStatus()

    Dim SubpacketId As Integer
    Dim ExtraInfo As Integer
    Dim I As Integer
    Dim QuestId As Integer
    Dim StageNumber As Integer
    Dim RequirementIndex As Integer
    Dim Quantity As Long
    
    'cw
    SubpacketId = Reader.ReadInt8
    QuestId = Reader.ReadInt
    StageNumber = Reader.ReadInt
    
    If QuestId <> PlayerData.Guild.Quest.Id Then Exit Sub
    If StageNumber <> PlayerData.Guild.Quest.CurrentStage Then Exit Sub
        
    With PlayerData.Guild.Quest.CurrentStageProgress
        Select Case SubpacketId
            Case eQuestUpdateEvent.EventNpcKill
            
                Dim NpcNumber As Integer
                NpcNumber = Reader.ReadInt
                RequirementIndex = Reader.ReadInt
                Quantity = Reader.ReadInt
                
                If RequirementIndex > GameMetadata.GuildQuests(QuestId).Stages(StageNumber).NpcsKillsQuantity Then Exit Sub
                If NpcNumber <> GameMetadata.GuildQuests(QuestId).Stages(StageNumber).NpcKill(RequirementIndex).NpcIndex Then Exit Sub
                
                PlayerData.Guild.Quest.CurrentStageProgress.NpcKilled(RequirementIndex) = Quantity
                                
            Case eQuestUpdateEvent.EventObjectCollect
                Dim ObjectIndex As Integer
                Dim Rest As Long
                
                ObjectIndex = Reader.ReadInt
                Quantity = Reader.ReadInt
                
                Call modRequiredObjectList.RequiredObjectListTryAdd(PlayerData.Guild.Quest.CurrentStageProgress.ObjsCollected, ObjectIndex, Quantity, Rest)
                
            Case eQuestUpdateEvent.EventUserKill
                       
                .FragsNeutralQty = Reader.ReadInt
                .FragsArmyQty = Reader.ReadInt
                .FragsLegionQty = Reader.ReadInt
                
            Case eQuestUpdateEvent.EventQuestFinished
                Dim Failed As Boolean
                Dim Message As String
                Failed = Reader.ReadBool

                If Failed Then
                    Message = "Tu clan ha fallado la misión"
                Else
                    Message = "Tu clan ha completado la misión."
                    Call AddNewCompletedQuest(QuestId)
                End If
                
                With FontTypes(FONTTYPE_GUILD)
                    Call AddtoRichTextBox(frmMain.RecTxt(eConsoleType.Agrupaciones), Message, .red, .green, .blue, .bold, .italic, True, eMessageType.Guild)
                End With
                
                modQuests.CleanCurrentQuestData
                

        End Select
    End With
    
    Call modQuests.RefreshObjectives
    Call UpdateFormInfo

End Sub

Private Sub HandleGuildQuestsCompletedList()
    Dim I As Integer
    Dim QuestsQty As Integer
        
    With PlayerData.Guild.Quest
        .CompletedQuantiy = Reader.ReadInt
        
        If .CompletedQuantiy > 0 Then
            ReDim .Completed(1 To .CompletedQuantiy)
            
            For I = 1 To .CompletedQuantiy
                .Completed(I) = Reader.ReadInt
            Next I
        Else
            Erase .Completed
        End If
    End With
        
End Sub

Private Sub HandleGuildCurrentQuestInfo()
     
    Dim Quantity As Long
    Dim I As Integer
    Dim TalkToNpc As Integer
        
    With PlayerData.Guild.Quest
        .Id = Reader.ReadInt
        
        .StartedDateTime = DateAdd("s", Reader.ReadInt, Now())
        .CurrentStage = Reader.ReadInt
        
        .CurrentStageProgress.EndStageNpc = Reader.ReadInt
        
        .CurrentStageProgress.NpcKilledQty = Reader.ReadInt
        If .CurrentStageProgress.NpcKilledQty > 0 Then
            ReDim .CurrentStageProgress.NpcKilled(1 To .CurrentStageProgress.NpcKilledQty)
            For I = 1 To .CurrentStageProgress.NpcKilledQty
                .CurrentStageProgress.NpcKilled(I) = Reader.ReadInt
            Next I
        End If
        
        If GameMetadata.GuildQuests(.Id).Stages(.CurrentStage).ObjsCollectQuantity > 0 Then
            .CurrentStageProgress.ObjsCollected = modRequiredObjectList.RequiredObjectListCreate(GameMetadata.GuildQuests(.Id).Stages(.CurrentStage).ObjsCollect, GameMetadata.GuildQuests(.Id).Stages(.CurrentStage).ObjsCollectQuantity)
        Else
            .CurrentStageProgress.ObjsCollected = modRequiredObjectList.RequiredObjectListCreateCompleted()
        End If
        
        Dim ObjectsQuantity As Integer
        ObjectsQuantity = Reader.ReadInt
        
        If ObjectsQuantity > 0 Then
            Dim ObjectIndex As Integer
            Dim Rest As Long
            
            For I = 1 To ObjectsQuantity
                ObjectIndex = Reader.ReadInt
                Quantity = Reader.ReadInt
                Call modRequiredObjectList.RequiredObjectListTryAdd(.CurrentStageProgress.ObjsCollected, ObjectIndex, Quantity, Rest)
            Next I
        End If
        
        .CurrentStageProgress.FragsArmyQty = Reader.ReadInt
        .CurrentStageProgress.FragsLegionQty = Reader.ReadInt
        .CurrentStageProgress.FragsNeutralQty = Reader.ReadInt
        
        .CurrentStageProgress.RequirementsCompleted = Reader.ReadBool
 
        PlayerData.Guild.CurrentQuest.Duration = GameMetadata.GuildQuests(.Id).Duration
        
        Call modQuests.RefreshObjectives
        Call UpdateFormInfo
    End With

End Sub

Private Sub HandleGuildSendInvitation()

    Dim InvitedByUserName As String
    Dim GuildIndex As Integer
    Dim GuildName As String
    Dim InvitationIndex As Integer
    Dim InvitationLifeTimeInMinutes As Integer

    InvitedByUserName = Reader.ReadString16()
    GuildIndex = Reader.ReadInt()
    GuildName = Reader.ReadString16()
    InvitationIndex = Reader.ReadInt()
    InvitationLifeTimeInMinutes = Reader.ReadInt()

    Call Guilds.GuildInvitation(InvitedByUserName, GuildIndex, GuildName, InvitationIndex, InvitationLifeTimeInMinutes)
    
End Sub

Private Sub HandleGuildBankList()
    Dim CountBankItems As Integer
    Dim I As Integer
    
    CountBankItems = Reader.ReadInt16()
   
    ReDim PlayerData.Guild.Bank(1 To CountBankItems) As tGuildBank

    For I = 1 To CountBankItems
        With PlayerData.Guild.Bank(I)
            .IdObject = Reader.ReadInt32()
            .Box = Reader.ReadInt16()
            .Slot = Reader.ReadInt16()
            .Amount = Reader.ReadInt16()
            .CanUse = Reader.ReadBool()
        End With
    Next I

    Call frmGuildBank.FillGuildBankInv
    Call frmGuildBank.FillMemberInv

    Exit Sub
End Sub


Private Sub HandleGuildBankChangeSlot()
    Dim Slot As Byte
    Dim CanUse As Boolean
    Dim Amount As Integer, Box As Integer, ObjIndex As Integer
    Dim GrhIndex As Integer, OBJType As Integer
    Dim NameObj As String
    
    Slot = Reader.ReadInt8()
    Box = Reader.ReadInt16()
    ObjIndex = Reader.ReadInt16()
    Amount = Reader.ReadInt16()
    CanUse = Reader.ReadBool()
    
    If Slot = 0 Then
        Erase PlayerData.Guild.Bank
    Else
        With PlayerData.Guild.Bank(Slot)
        
            If ObjIndex > 0 Then
                .IdObject = ObjIndex
                NameObj = GameMetadata.Objs(ObjIndex).Name
                .Amount = Amount
                GrhIndex = GameMetadata.Objs(ObjIndex).GrhIndex
                OBJType = GameMetadata.Objs(ObjIndex).OBJType
                .CanUse = CanUse
                
            Else
                .IdObject = 0
                NameObj = "Nada"
                .Amount = 0
                GrhIndex = 0
                OBJType = 0
                .CanUse = 0
            End If
        
            
            If Not (frmGuildBank.GBankInv Is Nothing) Then
                If frmGuildBank.Visible Then
                    Call frmGuildBank.GBankInv.SetItem(Slot, .IdObject, .Amount, _
                        0, GrhIndex, OBJType, 0, _
                        0, 0, 0, 0, NameObj, 0, .CanUse)
                End If
            End If
            
        End With
    End If

End Sub


Public Sub WriteGuildBankEnd()
On Error GoTo ErrHandler
    
    Call Writer.WriteInt8(ClientPacketID.GuildBankEnd)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteGuildBankEnd de Protocol.bas")
End Sub


Public Sub WriteWorkerStore_Create(ByRef ItemsToCraft As tCurrentOpenStore)
    
    Dim I As Integer
    
    Call Writer.WriteInt8(ClientPacketID.WorkerStore)
    
    ' Packet Subtype
    Call Writer.WriteInt8(eWorkerStoreAction.WorkerStoreCreate)
    
    ' Qty Items
    Call Writer.WriteInt16(ItemsToCraft.ItemsQty)
    
    For I = 1 To ItemsToCraft.ItemsQty
        With ItemsToCraft.Items(I)
            
            Call Writer.WriteInt(.RecipeNumber)
            Call Writer.WriteInt(.RecipeIndex)
            Call Writer.WriteInt(.ConstructionPrice)
            Call Writer.WriteInt(.MaterialsPrice)
            Call Writer.WriteInt8(.SelectedCraftingGroup)
            
            
        End With
        
    Next I

    Call Send(False)
    
End Sub

Public Sub WriteWorkerStore_Close()
    
    Call Writer.WriteInt8(ClientPacketID.WorkerStore)
    
    ' Packet Subtype
    Call Writer.WriteInt8(eWorkerStoreAction.WorkerStoreClose)
    
    Call Send(False)
    
End Sub

Public Sub HandleWorkerStore()
    
    Select Case Reader.ReadInt8
        Case eWorkerStoreServerSubAction.ShowStore
            Call HandleWorkerStore_Show
        Case eWorkerStoreServerSubAction.OpenFormForCreation
            Call HandleWorkerStore_OpenFormForCreation
        Case eWorkerStoreServerSubAction.OpenStore
            Call frmCraftingStore.SetStoreStatus(True, True)
        Case eWorkerStoreServerSubAction.ItemCrafted
            Call HandleWorkerStore_ItemCraftedNotification
        
    End Select
    
End Sub

Public Sub HandleWorkerStore_Show()
    Dim WorkerName As String
    Dim StoreType As Integer
    Dim ItemsQty As Integer
    Dim StoreInstanceId As String
    Dim I As Integer
    
    Dim ItemNumber As Integer
    Dim SelectedCraftingGroup As Byte
    
    'CurrentOpenStore.OwnerName = Reader.ReadString16
    WorkerName = Reader.ReadString16
    StoreInstanceId = Reader.ReadString16

    'CurrentOpenStore.ItemsQty = Reader.ReadInt16
    ItemsQty = Reader.ReadInt16
    
    If ItemsQty = 0 Then
        PlayerData.CraftingRecipeGroupsQty = 0
        Erase PlayerData.CraftingRecipeGroups
        Exit Sub
    End If
    
    ReDim PlayerData.CraftingRecipeGroups(1 To 1)
    PlayerData.CraftingRecipeGroupsQty = 1
    PlayerData.CraftingRecipeGroups(1).TabTitle = "Objetos"
    PlayerData.CraftingRecipeGroups(1).TabImage = "BotonItems"
    
    ReDim PlayerData.CraftingRecipeGroups(1).Recipes(1 To ItemsQty)
    PlayerData.CraftingRecipeGroups(1).RecipesQty = ItemsQty
    
    For I = 1 To ItemsQty
        With PlayerData.CraftingRecipeGroups(1).Recipes(I)
            .ObjNumber = Reader.ReadInt16
            .RecipeIndex = Reader.ReadInt16
            .ConstructionPrice = Reader.ReadInt16
            .MaterialsPrice = Reader.ReadInt16
            .SelectedCraftingGroup = Reader.ReadInt8
            
            .MaterialsQty = Reader.ReadInt16
            If .MaterialsQty > 0 Then
                ReDim .Materials(1 To .MaterialsQty)
                Dim J As Integer
                
                For J = 1 To .MaterialsQty
                    .Materials(J).ObjNumber = Reader.ReadInt16
                    .Materials(J).Amount = Reader.ReadInt16
                Next J
            End If
        End With
    Next I
    
    Load frmCraftingStore
    Call frmCraftingStore.Show(vbModeless, frmMain)
    Call frmCraftingStore.SetStoreMode(True, WorkerName, StoreInstanceId)
    Call frmCraftingStore.SelectFirstGroup
    Call frmCraftingStore_History.CleanControls

    
    
    
End Sub

Public Sub HandleWorkerStore_OpenFormForCreation() 'GetWorkerRecipes()
    Dim WorkerName As String
    Dim StoreType As Integer
    Dim ItemsQty As Integer
    Dim I As Integer
    Dim J As Integer
    
    Load frmCraftingStore
    Call frmCraftingStore.Show(vbModeless, frmMain)
    Call frmCraftingStore.SetStoreMode(False, vbNullString)
    Call frmCraftingStore.SelectFirstGroup
    Call frmCraftingStore_History.CleanControls

End Sub


Public Sub HandleWorkerStore_ItemCraftedNotification()
    Dim BuyerName As String
    Dim ItemNumber As Integer
    Dim ConstructionPrice As Double
    Dim ItemQuantity As Integer
    
    BuyerName = Reader.ReadString16()
    ItemNumber = Reader.ReadInt16()
    ItemQuantity = Reader.ReadInt16()
    ConstructionPrice = Reader.ReadInt32()
    
    Call frmCraftingStore_History.AddCraftedItemLog(ItemNumber, ItemQuantity, ConstructionPrice, BuyerName)
  
End Sub



Public Sub WriteWorkerStore_CraftItem(ByVal SelectedItemIndex As Integer, ByRef InstanceId As String)
    
    Call Writer.WriteInt8(ClientPacketID.WorkerStore)
    Call Writer.WriteInt8(eWorkerStoreAction.WorkerStoreCraftItem)
    Call Writer.WriteInt16(SelectedItemIndex)
    Call Writer.WriteString16(InstanceId)
    
    Call Send(False)
    
End Sub

Public Sub WriteWorkerStore_WorkerStoreGetRecipes()

    Call Writer.WriteInt8(ClientPacketID.WorkerStore)
    Call Writer.WriteInt8(eWorkerStoreAction.WorkerStoreGetRecipes)
    
    Call Send(False)
    
End Sub
Private Sub HandleGuildUpgradesList()

    Dim CountUpgrade As Integer, CountUpgradeGroup As Integer
    Dim I As Integer, J As Integer
    Dim UpgReqQty As Integer, QstReqQty As Integer
    
    CountUpgrade = Reader.ReadInt16()
    CountUpgradeGroup = Reader.ReadInt16()
    
    ReDim GuildUpgrades(1 To CountUpgrade) As GuildReqUpgradeType
    ReDim GuildUpgradesGroup(1 To CountUpgradeGroup) As GuildUpgradeGroupConfig
    
    
    For I = 1 To CountUpgradeGroup
        GuildUpgradesGroup(I).UpgradeQty = Reader.ReadInt16()
        
        If GuildUpgradesGroup(I).UpgradeQty > 0 Then
            ReDim GuildUpgradesGroup(I).Upgrades(1 To GuildUpgradesGroup(I).UpgradeQty)
              
            For J = 1 To GuildUpgradesGroup(I).UpgradeQty
                GuildUpgradesGroup(I).Upgrades(J) = Reader.ReadInt16()
            Next J
        End If
    Next I
    
    For I = 1 To CountUpgrade
        With GuildUpgrades(I)
            .Name = Reader.ReadString16()
            .Description = Reader.ReadString16()
            .IconGraph = Reader.ReadInt32()
            .GoldCost = Reader.ReadInt32()
            .ContributionCost = Reader.ReadInt32()
            
            UpgReqQty = Reader.ReadInt16()
            If UpgReqQty > 0 Then
                ReDim .UpgradeRequired(1 To UpgReqQty) As Integer
                For J = 1 To UpgReqQty
                     .UpgradeRequired(J) = Reader.ReadInt16()
                Next J
            End If
            
            QstReqQty = Reader.ReadInt16()
            If QstReqQty > 0 Then
                ReDim .QuestRequired(1 To QstReqQty) As GuildQuestReq
                For J = 1 To QstReqQty
                    .QuestRequired(J).Id = Reader.ReadInt16()
                    .QuestRequired(J).Title = Reader.ReadString16()
                    .QuestRequired(J).Obtained = Reader.ReadBool()
                Next J
            End If
        End With
    Next I
    
    frmGuildUpgrades.LoadUpgrades

End Sub

Private Sub HandleGuildInfoChange()
    Dim TypeChanged As Byte
    Dim ValueChanged As Byte
    Dim ValueChangedLong As Long

    TypeChanged = Reader.ReadInt8()
    ValueChanged = Reader.ReadInt8()
    ValueChangedLong = Reader.ReadInt32()

    Select Case TypeChanged

        Case eChangeGuildInfo.MaxMembersQtyChange 'maximum amount of the member
            PlayerData.Guild.MaxMemberQty = ValueChanged
            
            If frmGuildInformation.Visible Then
               frmGuildInformation.lblMemberCount.Caption = PlayerData.Guild.MemberCount & "/" & PlayerData.Guild.MaxMemberQty
            End If
            
            If frmGuildMembers.Visible Then
               frmGuildMembers.lblGuildMemberQty.Caption = PlayerData.Guild.MemberCount & "/" & PlayerData.Guild.MaxMemberQty
            End If
        Case eChangeGuildInfo.MaxRolesQtyChange 'maximum amount of the roles
            PlayerData.Guild.MaxRoles = ValueChanged
            
            If frmGuildRolesList.Visible Then
               frmGuildRolesList.lblGuildRolesQty.Caption = UBound(PlayerData.Guild.Roles) & "/" & PlayerData.Guild.MaxRoles
            End If
        Case eChangeGuildInfo.MaxSlotsBankQtyChange 'maximum amount of the bank's slots
            PlayerData.Guild.MaxSlotBank = ValueChanged
            
            ReDim Preserve PlayerData.Guild.Bank(1 To PlayerData.Guild.MaxSlotBank) As tGuildBank
            Call frmGuildBank.Reload
            
            If frmGuildBank.Visible Then
            '   frmGuildBank.Lbl .Caption = ValueChanged
            End If
        Case eChangeGuildInfo.MaxBoxesBankQtyChange 'maximum amount of the bank's boxes
            PlayerData.Guild.MaxBoxesBank = ValueChanged
            
            If frmGuildBank.Visible Then
            '   frmGuildBank.Lbl .Caption = ValueChanged
            End If
        Case eChangeGuildInfo.MaxContributionChange 'maximum amount of contribution point
            PlayerData.Guild.MaxContributionAvailable = ValueChangedLong
            
            If frmGuildInformation.Visible Then
               frmGuildInformation.lblContribution.Caption = PlayerData.Guild.ContributionAvailable & "/" & PlayerData.Guild.MaxContributionAvailable
            End If
        Case eChangeGuildInfo.ContributionAvailableChange
            PlayerData.Guild.ContributionAvailable = ValueChangedLong
            
            If frmGuildInformation.Visible Then
               frmGuildInformation.lblContribution.Caption = PlayerData.Guild.ContributionAvailable & "/" & PlayerData.Guild.MaxContributionAvailable
            End If
        Case eChangeGuildInfo.BankGoldChange
            PlayerData.Guild.BankGold = ValueChangedLong
            
            If frmGuildBank.Visible Then
               frmGuildBank.LblBankGold.Caption = PlayerData.Guild.BankGold
            End If
            
        Case eChangeGuildInfo.EnableBank
            PlayerData.Guild.BankAvalaible = ValueChanged
            
            If frmGuildMain.Visible Then
                frmGuildMain.ImgBlockBank.Visible = False
            End If
        Case eChangeGuildInfo.CompletedQuestAdded
            
            Call modQuests.AddNewCompletedQuest(ValueChanged)
    End Select
        
    Exit Sub
    
End Sub

Private Sub UpgradeQuestReq(ByVal QuestId As Integer)
    Dim I As Integer, J As Integer
    Dim UpgradeQty As Integer, QuestReqQty As Integer
    
    UpgradeQty = GetQtyGuildUpgradesList()
    
    For I = 1 To UpgradeQty
        If ((Not GuildUpgrades(I).QuestRequired) = -1) Then
             QuestReqQty = 0
        Else
            QuestReqQty = UBound(GuildUpgrades(I).QuestRequired)
        End If
        
        For J = 1 To QuestReqQty
            If GuildUpgrades(I).QuestRequired(J).Id = QuestId Then
                GuildUpgrades(I).QuestRequired(J).Obtained = True
            End If
        Next J
    Next I


End Sub

Public Sub WritePartyInviteMember(ByVal UserName As String)
On Error GoTo ErrHandler
  
    UserName = complexNameToSimple(UserName, True)
    
        Call Writer.WriteInt8(ClientPacketID.PartyInviteMember)
        
        Call Writer.WriteString8(UserName)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WritePartyInviteMember de Protocol.bas")
End Sub

Private Sub HandlePartyInvitation()

    Dim UserNameRequest As String
    Dim UserIndexRequest As Integer
    
    UserNameRequest = Reader.ReadString16()
    UserIndexRequest = Reader.ReadInt8()
    
    Call PartyTempInviSave(UserNameRequest, UserIndexRequest)
    
End Sub

Public Sub WritePartyAcceptInvitation(ByVal UserIndex As Integer, ByVal NamePlayer As String, ByVal Accepted As Boolean)
On Error GoTo ErrHandler
       
    Call Writer.WriteInt8(ClientPacketID.PartyAcceptInvitation)
    
    Call Writer.WriteString16(NamePlayer)
    Call Writer.WriteBool(Accepted)
    
    Call Send(False)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WritePartyAcceptInvitation de Protocol.bas")
End Sub

Public Sub HandleSetIntervals()
     
    With PlayerData.Intervals
        .SpellCastMacro = Reader.ReadInt
        .WorkMacro = Reader.ReadInt
        .Actions = Reader.ReadInt
        .PlayerAttack = Reader.ReadInt
        .PlayerAttackArrow = Reader.ReadInt
        .PlayerCastSpell = Reader.ReadInt
        .PlayerAttackAfterSpell = Reader.ReadInt
        .PlayerCastSpellAfterAttack = Reader.ReadInt
        .Work = Reader.ReadInt
        .UseItemWithKey = Reader.ReadInt
        .UseItemDoubleClick = Reader.ReadInt
        .RequestPositionUpdate = Reader.ReadInt
        .Meditate = Reader.ReadInt
        
        
        Call frmMain.hlst.SetDefaultCooldown(.PlayerCastSpell)
        Call frmMain.hlst.SetSpellAfterMeleeCooldown(.PlayerCastSpellAfterAttack)
    End With
    
    Call LoadTimerIntervals

End Sub

Public Sub HandleAttackResult()
    Dim AttackResult As Boolean
    
    AttackResult = Reader.ReadBool
    
    If esGM(UserCharIndex) Then Exit Sub
    
    ' If the attack was successful, then we need to update the attack timer.
    If AttackResult = True Then Call MainTimer.Check(TimersIndex.Attack)

End Sub


Public Sub HandleSpellAttackResult()
    
    Dim SpellAttackResult As Boolean
    Dim SpellIndex As Integer
    
    SpellAttackResult = Reader.ReadBool
    SpellIndex = Reader.ReadInt()
    
    If esGM(UserCharIndex) Then Exit Sub
    
    If SpellAttackResult Then charlist(UserCharIndex).LastSpellCast = SpellIndex - 1
    
    ' If the attack was successful, then we need to start the timer and the cooldown progressbar
    Call frmMain.hlst.Start(SpellIndex - 1, Not SpellAttackResult)
    
End Sub
