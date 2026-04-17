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
' @author Juan Martín Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version 1.0.0
' @date 20060517
'@Folder("Protocol")
Option Explicit

''
'When we have a list of strings, we use this to separate them and prevent
'having too many string lengths in the queue. Yes, each string is NULL-terminated :P
Private Const SEPARATOR As String * 1 = vbNullChar

Public Enum eMessageType
    info = 0
    Admin = 1
    Guild = 2
    Party = 3
    Combate = 4
    Trabajo = 5
End Enum

''
'Auxiliar ByteQueue used as buffer to generate messages not intended to be sent right away.
'Specially usefull to create a message once and send it over to several clients.
Public Writer  As BinaryWriter
Public Reader  As BinaryReader

#If EnableSecurity = 0 Then
Private Enum ServerPacketID
    Connected               ' CONNECTED
    logged                  ' LOGGED
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
    PlayMusic                ' TM
    PlayEffect                ' TW
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
    DropXY                    'TIXY
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
    LeaveFaction            '/RETIRARFACCION ( with no arguments )
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

'desafios - Mithrandir: Se podría crear 1 paquete, en vez de 2 (para aceptar y cancelar)
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
    
    WorkerStore
    
    'Put new packets before this one. LastClientPacketId should be the last element of the enum
    LastClientPacketId
End Enum
#End If
 
''
'The last existing client packet id.
Private LAST_CLIENT_PACKET_ID As Byte

Public Enum FontTypeNames
    FONTTYPE_TALK
    FONTTYPE_FIGHT
    FONTTYPE_WARNING
    FONTTYPE_INFO
    FONTTYPE_INFOBOLD
    FONTTYPE_EJECUCION
    FONTTYPE_PARTY
    FONTTYPE_VENENO
    FONTTYPE_GUILD
    FONTTYPE_SERVER
    FONTTYPE_GUILDMSG
    FONTTYPE_CONSEJO
    FONTTYPE_CONSEJOCAOS
    FONTTYPE_CONSEJOVesA
    FONTTYPE_CONSEJOCAOSVesA
    FONTTYPE_CENTINELA
    FONTTYPE_GMMSG
    FONTTYPE_GM
    FONTTYPE_NEWBIE
    FONTTYPE_NEUTRAL
    FONTTYPE_CITIZEN
    FONTTYPE_CONSE
    FONTTYPE_DIOS
    FONTTYPE_NPCNAME
End Enum

Public Enum eEditOptions
    eo_Gold = 1
    eo_Experience
    eo_Body
    eo_Head
    eo_CiticensKilled
    eo_CriminalsKilled
    eo_Level
    eo_Class
    eo_Skills
    eo_SkillPointsLeft
    eo_Sex
    eo_Raza
    eo_addGold
    eo_Vida
    eo_Poss
    eo_PlayerPoints
End Enum

' Enumerators for Guild purposes
Public Enum eExchangeType
    IsGold = 1
    IsObject
End Enum

Public Enum eExchangeAction
    Withdraw = 1
    Deposit
End Enum

Public Enum eMemberAction
    KickMember = 1
    SendInvitation
    LeaveGuild
End Enum

Public Enum eRoleAction
    Assign = 1
    Create
    Delete
    Update
End Enum

Public Enum eWorkerStoreServerSubAction
    ShowStore = 1
    OpenFormForCreation = 2
    OpenStore = 3
    ItemCrafted = 4
End Enum

' End enumerators for Guild purposes

Public Sub InitAuxiliarBuffer()
'***************************************************
'Author: ZaMa
'Last Modification: 15/03/2011
'Initializaes Auxiliar Buffer
'***************************************************
On Error GoTo ErrHandler
  
    Set Writer = New BinaryWriter
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub InitAuxiliarBuffer de Protocol.bas")
End Sub

''
' Handles incoming data.
'
' @param    userIndex The index of the user sending the message.

Public Function HandleIncomingData(ByVal UserIndex As Integer) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/09/07
'
'***************************************************
On Error GoTo ErrHandler
  
On Error Resume Next

    Dim PacketID As Long
    Dim IsLoggedUserRequired As Boolean
    Dim IsLoggedUser As Boolean
    
    PacketID = Reader.ReadInt8
    
    IsLoggedUser = UserList(UserIndex).flags.UserLogged
    IsLoggedUserRequired = Not (PacketID = ClientPacketID.LoginExistingChar _
        Or PacketID = ClientPacketID.AccountCreate _
        Or PacketID = ClientPacketID.AccountLogin Or PacketID = ClientPacketID.AccountLoginChar Or _
        PacketID = ClientPacketID.AccountDeleteChar Or PacketID = ClientPacketID.AccountRecover Or _
        PacketID = ClientPacketID.AccountCreateChar Or PacketID = ClientPacketID.AccountChangePassword)
        
    'Does the packet requires a logged user??
    If IsLoggedUserRequired And IsLoggedUser = False Or _
       IsLoggedUserRequired = False And IsLoggedUser Then
        Call CloseSocket(UserIndex)
        Exit Function
    End If
    
    If Not PacketID = ClientPacketID.Ping Then
        'Reset idle counter if id is valid.
        UserList(UserIndex).Counters.IdleCount = SetIntervalEnd(ServerConfiguration.Intervals.IntervaloIdleKick)
        
        ' Pierde la proteccion contra ataques
        UserList(UserIndex).flags.NoPuedeSerAtacado = False
    End If

    Select Case PacketID
        Case ClientPacketID.LoginExistingChar       'OLOGIN
            Call HandleLoginExistingChar(UserIndex)
        
        Case ClientPacketID.Talk                    ';
            Call HandleTalk(UserIndex)
        
        Case ClientPacketID.Yell                    '-
            Call HandleYell(UserIndex)
        
        Case ClientPacketID.Whisper                 '\
            Call HandleWhisper(UserIndex)
        
        Case ClientPacketID.Walk                    'M
            Call HandleWalk(UserIndex)
            
        Case ClientPacketID.RequestPositionUpdate   'RPU
            Call HandleRequestPositionUpdate(UserIndex)
        
        Case ClientPacketID.Attack                  'AT
            Call HandleAttack(UserIndex)
        
        Case ClientPacketID.PickUp                  'AG
            Call HandlePickUp(UserIndex)
        
        Case ClientPacketID.SafeToggle              '/SEG & SEG  (SEG's behaviour has to be coded in the client)
            Call HandleSafeToggle(UserIndex)
        
        Case ClientPacketID.ResuscitationSafeToggle
            Call HandleResuscitationToggle(UserIndex)
        
        Case ClientPacketID.RequestSkills           'ESKI
            Call HandleRequestSkills(UserIndex)
        
        Case ClientPacketID.RequestStadictis
            Call HandleRequestStadictics(UserIndex)
        
        Case ClientPacketID.CommerceEnd             'FINCOM
            Call HandleCommerceEnd(UserIndex)
            
        Case ClientPacketID.CommerceChat
            Call HandleCommerceChat(UserIndex)
        
        Case ClientPacketID.UserCommerceEnd         'FINCOMUSU
            Call HandleUserCommerceEnd(UserIndex)
            
        Case ClientPacketID.UserCommerceConfirm
            Call HandleUserCommerceConfirm(UserIndex)
        
        Case ClientPacketID.BankEnd                 'FINBAN
            Call HandleBankEnd(UserIndex)
        
        Case ClientPacketID.UserCommerceOk          'COMUSUOK
            Call HandleUserCommerceOk(UserIndex)
        
        Case ClientPacketID.UserCommerceReject      'COMUSUNO
            Call HandleUserCommerceReject(UserIndex)
        
        Case ClientPacketID.Drop                    'TI
            Call HandleDrop(UserIndex)
        
        Case ClientPacketID.DropXY                    'TIXY
            Call HandleDropXY(UserIndex)
        
        Case ClientPacketID.CastSpell               'LH
            Call HandleCastSpell(UserIndex)
        
        Case ClientPacketID.LeftClick               'LC
            Call HandleLeftClick(UserIndex)
        
        Case ClientPacketID.DoubleClick             'RC
            Call HandleDoubleClick(UserIndex)
        
        Case ClientPacketID.Work                    'UK
            Call HandleWork(UserIndex)
        
        Case ClientPacketID.UseSpellMacro           'UMH
            Call HandleUseSpellMacro(UserIndex)
        
        Case ClientPacketID.UseItem                 'USA
            Call HandleUseItem(UserIndex)

        Case ClientPacketID.WorkLeftClick           'WLC
            Call HandleWorkLeftClick(UserIndex)
        
        Case ClientPacketID.SpellInfo               'INFS
            Call HandleSpellInfo(UserIndex)
        
        Case ClientPacketID.EquipItem               'EQUI
            Call HandleEquipItem(UserIndex)
        
        Case ClientPacketID.ChangeHeading           'CHEA
            Call HandleChangeHeading(UserIndex)
        
        Case ClientPacketID.ModifySkills            'SKSE
            Call HandleModifySkills(UserIndex)
        
        Case ClientPacketID.Train                   'ENTR
            Call HandleTrain(UserIndex)
        
        Case ClientPacketID.CommerceBuy             'COMP
            Call HandleCommerceBuy(UserIndex)
        
        Case ClientPacketID.BankExtractItem         'RETI
            Call HandleBankExtractItem(UserIndex)
        
        Case ClientPacketID.CommerceSell            'VEND
            Call HandleCommerceSell(UserIndex)
        
        Case ClientPacketID.BankDeposit             'DEPO
            Call HandleBankDeposit(UserIndex)
        
        Case ClientPacketID.ForumPost               'DEMSG
            Call HandleForumPost(UserIndex)
        
        Case ClientPacketID.MoveSpell               'DESPHE
            Call HandleMoveSpell(UserIndex)
            
        Case ClientPacketID.MoveBank
            Call HandleMoveBank(UserIndex)
        
        Case ClientPacketID.UserCommerceOffer       'OFRECER
            Call HandleUserCommerceOffer(UserIndex)
                  
        Case ClientPacketID.UserCommerceOfferGold
            Call HandleUserCommerceOfferGold(UserIndex)
                  
        Case ClientPacketID.Online                  '/ONLINE
            Call HandleOnline(UserIndex)
        
        Case ClientPacketID.Quit                    '/SALIR
            Call HandleQuit(UserIndex)
        
        Case ClientPacketID.RequestAccountState     '/BALANCE
            Call HandleRequestAccountState(UserIndex)
        
        Case ClientPacketID.PetStand                '/QUIETO
            Call HandlePetStand(UserIndex)
        
        Case ClientPacketID.PetFollow               '/ACOMPAÑAR
            Call HandlePetFollow(UserIndex)
            
        Case ClientPacketID.ReleasePet              '/LIBERAR
            Call HandleReleasePet(UserIndex)
        
        Case ClientPacketID.TrainList               '/ENTRENAR
            Call HandleTrainList(UserIndex)
        
        Case ClientPacketID.Rest                    '/DESCANSAR
            Call HandleRest(UserIndex)
        
        Case ClientPacketID.Meditate                '/MEDITAR
            Call HandleMeditate(UserIndex)
        
        Case ClientPacketID.Resucitate              '/RESUCITAR
            Call HandleResucitate(UserIndex)
        
        Case ClientPacketID.Heal                    '/CURAR
            Call HandleHeal(UserIndex)
        
        Case ClientPacketID.Help                    '/AYUDA
            Call HandleHelp(UserIndex)
        
        Case ClientPacketID.RequestStats            '/EST
            Call HandleRequestStats(UserIndex)
        
        Case ClientPacketID.CommerceStart           '/COMERCIAR
            Call HandleCommerceStart(UserIndex)
        
        Case ClientPacketID.BankStart               '/BOVEDA
            Call HandleBankStart(UserIndex)
        
        Case ClientPacketID.Enlist                  '/ENLISTAR
            Call HandleEnlist(UserIndex)
        
        Case ClientPacketID.Information             '/INFORMACION
            Call HandleInformation(UserIndex)
        
        Case ClientPacketID.Reward                  '/RECOMPENSA
            Call HandleReward(UserIndex)
        
        Case ClientPacketID.UpTime                  '/UPTIME
            Call HandleUpTime(UserIndex)
        
        Case ClientPacketID.PartyLeave              '/SALIRPARTY
            Call HandlePartyLeave(UserIndex)
        
        Case ClientPacketID.PartyCreate             '/CREARPARTY
            Call HandlePartyCreate(UserIndex)
        
        Case ClientPacketID.Inquiry                 '/ENCUESTA ( with no params )
            Call HandleInquiry(UserIndex)
            
        Case ClientPacketID.GuildMessage            '/CMSG
            Call HandleGuildMessage(UserIndex)
        
        Case ClientPacketID.PartyMessage            '/PMSG
            Call HandlePartyMessage(UserIndex)
        
        Case ClientPacketID.PartyOnline             '/ONLINEPARTY
            Call HandlePartyOnline(UserIndex)
        
        Case ClientPacketID.CouncilMessage          '/BMSG
            Call HandleCouncilMessage(UserIndex)
        
        Case ClientPacketID.RoleMasterRequest       '/ROL
            Call HandleRoleMasterRequest(UserIndex)
        
        Case ClientPacketID.GMRequest               '/GM
            Call HandleGMRequest(UserIndex)
        
        Case ClientPacketID.bugReport               '/_BUG
            Call HandleBugReport(UserIndex)
        
        Case ClientPacketID.ChangeDescription       '/DESC
            Call HandleChangeDescription(UserIndex)
        
        Case ClientPacketID.Punishments             '/PENAS
            Call HandlePunishments(UserIndex)
        
        Case ClientPacketID.Gamble                  '/APOSTAR
            Call HandleGamble(UserIndex)
        
        Case ClientPacketID.InquiryVote             '/ENCUESTA ( with parameters )
            Call HandleInquiryVote(UserIndex)
        
        Case ClientPacketID.LeaveFaction            '/RETIRARFACCION ( with no arguments )
            Call HandleLeaveFaction(UserIndex)
        
        Case ClientPacketID.BankExtractGold         '/RETIRAR ( with arguments )
            Call HandleBankExtractGold(UserIndex)
        
        Case ClientPacketID.BankDepositGold         '/DEPOSITAR
            Call HandleBankDepositGold(UserIndex)
        
        Case ClientPacketID.Denounce                '/DENUNCIAR
            Call HandleDenounce(UserIndex)
        
        Case ClientPacketID.PartyKick               '/ECHARPARTY
            Call HandlePartyKick(UserIndex)
        
        Case ClientPacketID.PartySetLeader          '/PARTYLIDER
            Call HandlePartySetLeader(UserIndex)
        
        Case ClientPacketID.PartyInviteMember       '
            Call HandlePartyInviteMember(UserIndex)
        
        Case ClientPacketID.PartyAcceptInvitation
            Call HandlePartyAcceptInvitation(UserIndex)
        
        Case ClientPacketID.Ping                    '/PING
            Call HandlePing(UserIndex)
            
        Case ClientPacketID.RequestPartyForm
            Call HandlePartyForm(UserIndex)
            
        Case ClientPacketID.ItemUpgrade
            Call HandleItemUpgrade(UserIndex)
        
        Case ClientPacketID.GMCommands              'GM Messages
            Call HandleGMCommands(UserIndex)
            
        Case ClientPacketID.InitCrafting
            Call HandleInitCrafting(UserIndex)
        
        Case ClientPacketID.Home
            Call HandleHome(UserIndex)
            
        Case ClientPacketID.Consultation
            Call HandleConsultation(UserIndex)
        
        Case ClientPacketID.moveItem
            Call HandleMoveItem(UserIndex)
            
        Case ClientPacketID.RightClick
            Call HandleRightClick(UserIndex)

        Case ClientPacketID.PMList
            Call HandlePMList(UserIndex)
            
        Case ClientPacketID.PMDeleteList
            Call HandlePMDeleteList(UserIndex)
        
        Case ClientPacketID.MenuAction
            Call HandleMenuAction(UserIndex)
            
        Case ClientPacketID.Participar
           Call HandleTournamentParticipate(UserIndex)
           
        Case ClientPacketID.AccountCreate
            Call HandleAccountCreate(UserIndex)
             
        Case ClientPacketID.AccountLogin
            Call HandleAccountLogin(UserIndex)
             
        Case ClientPacketID.AccountLoginChar
            Call HandleAccountLoginChar(UserIndex)
            
        Case ClientPacketID.AccountCreateChar
            Call HandleAccountCreateChar(UserIndex)
           
        Case ClientPacketID.AccountDeleteChar
            Call HandleAccountDeleteChar(UserIndex)
           
        Case ClientPacketID.AccountRecover
            Call HandleAccountRecover(UserIndex)
        
        Case ClientPacketID.AccountChangePassword
            Call HandleAccountChangePassword(UserIndex)
            
        'Desafio
        Case ClientPacketID.Chat_desafio
            Call HandleChatDesafio(UserIndex)
            
        Case ClientPacketID.Cancel_desafio
            Call HandleCancelDesafio(UserIndex)
        
        Case ClientPacketID.Accept_desafio
            Call HandleAceptDesafio(UserIndex)
            
        Case ClientPacketID.Enviardatos_desafio
            Call HandleDatosDesafio(UserIndex)

        Case ClientPacketID.Retar
            Call HandleRetar(UserIndex)
            
        Case ClientPacketID.AccBankExtractItem
            Call HandleAccBankExtractItem(UserIndex)
        
        Case ClientPacketID.Duelos
            Call HandleDuelos(UserIndex)
            
        Case ClientPacketID.AceptarDuelo
            Call HandleAceptarDuelo(UserIndex)
            
        Case ClientPacketID.RechazarDuelo
            Call HandleRechazarDuelo(UserIndex)
        
        Case ClientPacketID.DueloPublico
            Call HandleDueloPublico(UserIndex)
            
        Case ClientPacketID.CancelarEspera
            Call HandleCancelarEspera(UserIndex)
            
        Case ClientPacketID.CancelarElDuelo
            Call HandleCancelarElDuelo(UserIndex)
            
        Case ClientPacketID.AccBankDepositItem
            Call HandleAccBankDepositItem(UserIndex)
            
        Case ClientPacketID.AccBankExtractGold
            Call HandleAccBankExtractGold(UserIndex)
        
        Case ClientPacketID.AccBankDepositGold
            Call HandleAccBankDepositGold(UserIndex)
            
        Case ClientPacketID.AccBankStart
            Call HandleAccBankStart(UserIndex)
            
        Case ClientPacketID.AccBankEnd
            Call HandleAccBankEnd(UserIndex)
        
        Case ClientPacketID.AccBankChangePass
            Call HandleAccBankChangePass(UserIndex)

        Case ClientPacketID.CraftItem
            Call HandleCraftItem(UserIndex)
            
        Case ClientPacketID.SelectPet
            Call HandleSelectPet(UserIndex)
            
        Case ClientPacketID.MasteryAssign
            Call HandleMasteryAssign(UserIndex)
         
        Case ClientPacketID.GuildCreate
            Call HandleGuildCreate(UserIndex)
            
        Case ClientPacketID.GuildExchange
            Call HandleGuildExchange(UserIndex)
            
        Case ClientPacketID.GuildMember
            Call HandleGuildMember(UserIndex)
        
        Case ClientPacketID.GuildRole
            Call HandleGuildRole(UserIndex)
        
        Case ClientPacketID.GuildUpgrade
            Call HandleGuildUpgrade(UserIndex)
        
        Case ClientPacketID.GuildBankEnd
            Call HandleGuildBankEnd(UserIndex)
            
        Case ClientPacketID.GuildQuest
            Call HandleGuildQuest(UserIndex)
        
        Case ClientPacketID.GuildQuestAddObject
            Call HandleGuildQuestAddObject(UserIndex)
            
        Case ClientPacketID.GuildUserInvitationResponse
            Call HandleGuildUserInvitationResponse(UserIndex)
            
        Case ClientPacketID.WorkerStore
            Call HandleWorkerStore(UserIndex)
            
        
        Case Else
#If EnableSecurity Then
            Call HandleIncomingDataEx(UserIndex)
#Else
            'ERROR : Abort!
            Call CloseSocket(UserIndex)
#End If
    End Select

    If Err.Number <> 0 Then
        'An error ocurred, log it and kick player.
        Call LogError("Error: " & Err.Number & " [" & Err.Description & "] " & " Source: " & Err.Source & _
                        vbTab & " HelpFile: " & Err.HelpFile & vbTab & " HelpContext: " & Err.HelpContext & _
                        vbTab & " LastDllError: " & Err.LastDllError & vbTab & _
                        " - UserIndex: " & UserIndex & "(" & UserList(UserIndex).Name & ") (LastPacket: " & UserList(UserIndex).LastCompletedPacket & ") - producido al manejar el paquete: " & CStr(PacketID))
        Call CloseSocket(UserIndex)
    End If
    
    If Err.Number <> 0 Then
        Err.Clear
        HandleIncomingData = 100
    End If
    
    UserList(UserIndex).LastCompletedPacket = PacketID
    
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function HandleIncomingData de Protocol.bas")
End Function

Public Sub WriteMultiMessage(ByVal UserIndex As Integer, ByVal MessageIndex As Integer, Optional ByVal Arg1 As Long, Optional ByVal Arg2 As Long, Optional ByVal Arg3 As Long, Optional ByVal StringArg1 As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

        Call Writer.WriteInt8(ServerPacketID.MultiMessage)
        Call Writer.WriteInt8(MessageIndex)
        
        Select Case MessageIndex
            Case eMessages.DontSeeAnything, eMessages.NPCSwing, eMessages.NPCKillUser, eMessages.BlockedWithShieldUser, _
                eMessages.BlockedWithShieldother, eMessages.UserSwing, eMessages.SafeModeOn, eMessages.SafeModeOff, _
                eMessages.ResuscitationSafeOff, eMessages.ResuscitationSafeOn, eMessages.NobilityLost, _
                eMessages.CantUseWhileMeditating, eMessages.FinishHome
            
            Case eMessages.CancelHome
                Call Writer.WriteBool(Arg1) ' Cancelled
                
            Case eMessages.NPCHitUser
                Call Writer.WriteInt8(Arg1) 'Target
                Call Writer.WriteInt16(Arg2) 'damage
                
            Case eMessages.UserHitNPC
                Call Writer.WriteInt32(Arg1) 'damage
                
            Case eMessages.UserAttackedSwing
                Call Writer.WriteInt16(UserList(Arg1).Char.CharIndex)
                
            Case eMessages.UserHittedByUser
                Call Writer.WriteInt16(Arg1) 'AttackerIndex
                Call Writer.WriteInt8(Arg2) 'Target
                Call Writer.WriteInt16(Arg3) 'damage
                
            Case eMessages.UserHittedUser
                Call Writer.WriteInt16(Arg1) 'AttackerIndex
                Call Writer.WriteInt8(Arg2) 'Target
                Call Writer.WriteInt16(Arg3) 'damage
                
            Case eMessages.WorkRequestTarget
                Call Writer.WriteInt8(Arg1) 'skill
                
            Case eMessages.SpellCastRequestTarget
                Call Writer.WriteInt(Arg1) 'Spell Index
                Call Writer.WriteInt(Arg2) 'Spell Number
            
            Case eMessages.HaveKilledUser '"Has matado a " & UserList(VictimIndex).name & "!" "UserList(VictimIndex).name" & " era nivel " & VictimELV & "."
                Call Writer.WriteInt16(UserList(Arg1).Char.CharIndex) 'VictimIndex
                Call Writer.WriteInt32(Arg2) 'Level
            
            Case eMessages.UserKill '"¡" & .name & " te ha matado!"
                Call Writer.WriteInt16(UserList(Arg1).Char.CharIndex) 'AttackerIndex
            
            Case eMessages.EarnExp
            
            Case eMessages.Home
                Call Writer.WriteInt16(Arg1)
                'El cliente no conoce nada sobre nombre de mapas y hogares, por lo tanto _
                 hasta que no se pasen los dats e .INFs al cliente, esto queda así.
                Call Writer.WriteString8(StringArg1) 'Call Writer.WriteInt8(CByte(Arg2))
                
        End Select
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

Private Sub HandleGMCommands(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  
Dim Command As Long

With UserList(UserIndex)

    Command = Reader.ReadInt8
    
    Select Case Command
        Case eGMCommands.SpawnBoss
            Call HandleSpawnBoss(UserIndex)
            
        Case eGMCommands.GMMessage                '/GMSG
            Call HandleGMMessage(UserIndex)
        
        Case eGMCommands.ShowName                '/SHOWNAME
            Call HandleShowName(UserIndex)
        
        Case eGMCommands.OnlineRoyalArmy
            Call HandleOnlineRoyalArmy(UserIndex)
        
        Case eGMCommands.OnlineChaosLegion       '/ONLINECAOS
            Call HandleOnlineChaosLegion(UserIndex)
        
        Case eGMCommands.GoNearby                '/IRCERCA
            Call HandleGoNearby(UserIndex)
        
        Case eGMCommands.comment                 '/REM
            Call HandleComment(UserIndex)
        
        Case eGMCommands.serverTime              '/HORA
            Call HandleServerTime(UserIndex)
        
        Case eGMCommands.Where                   '/DONDE
            Call HandleWhere(UserIndex)
        
        Case eGMCommands.CreaturesInMap          '/NENE
            Call HandleCreaturesInMap(UserIndex)
        
        Case eGMCommands.WarpMeToTarget          '/TELEPLOC
            Call HandleWarpMeToTarget(UserIndex)
        
        Case eGMCommands.WarpChar                '/TELEP
            Call HandleWarpChar(UserIndex)
        
        Case eGMCommands.Silence                 '/SILENCIAR
            Call HandleSilence(UserIndex)
        
        Case eGMCommands.SOSShowList             '/SHOW SOS
            Call HandleSOSShowList(UserIndex)
            
        Case eGMCommands.SOSRemove               'SOSDONE
            Call HandleSOSRemove(UserIndex)
        
        Case eGMCommands.GoToChar                '/IRA
            Call HandleGoToChar(UserIndex)
        
        Case eGMCommands.invisible               '/INVISIBLE
            Call HandleInvisible(UserIndex)
        
        Case eGMCommands.GMPanel                 '/PANELGM
            Call HandleGMPanel(UserIndex)
        
        Case eGMCommands.RequestUserList         'LISTUSU
            Call HandleRequestUserList(UserIndex)

        Case eGMCommands.Jail                    '/CARCEL
            Call HandleJail(UserIndex)
        
        Case eGMCommands.KillNPC                 '/RMATA
            Call HandleKillNPC(UserIndex)
        
        Case eGMCommands.WarnUser                '/ADVERTENCIA
            Call HandleWarnUser(UserIndex)
        
        Case eGMCommands.EditChar                '/MOD
            Call HandleEditChar(UserIndex)
        
        Case eGMCommands.RequestStatsBosses
            Call HandleRequestStatsBosses(UserIndex)
        
        Case eGMCommands.RequestCharInfo         '/INFO
            Call HandleRequestCharInfo(UserIndex)
        
        Case eGMCommands.RequestCharStats        '/STAT
            Call HandleRequestCharStats(UserIndex)
        
        Case eGMCommands.RequestCharGold         '/BAL
            Call HandleRequestCharGold(UserIndex)
        
        Case eGMCommands.RequestCharInventory    '/INV
            Call HandleRequestCharInventory(UserIndex)
        
        Case eGMCommands.RequestCharBank         '/BOV
            Call HandleRequestCharBank(UserIndex)
        
        Case eGMCommands.RequestCharSkills       '/SKILLS
            Call HandleRequestCharSkills(UserIndex)
        
        Case eGMCommands.ReviveChar              '/REVIVIR
            Call HandleReviveChar(UserIndex)
        
        Case eGMCommands.OnlineGM                '/ONLINEGM
            Call HandleOnlineGM(UserIndex)
        
        Case eGMCommands.OnlineMap               '/ONLINEMAP
            Call HandleOnlineMap(UserIndex)
        
        Case eGMCommands.Kick                    '/ECHAR
            Call HandleKick(UserIndex)
        
        Case eGMCommands.Execute                 '/EJECUTAR
            Call HandleExecute(UserIndex)
        
        Case eGMCommands.BanChar                 '/BAN
            Call HandleBanChar(UserIndex)
        
        Case eGMCommands.UnBanChar               '/UNBAN
            Call HandleUnbanChar(UserIndex)
        
        Case eGMCommands.NPCFollow               '/SEGUIR
            Call HandleNPCFollow(UserIndex)
        
        Case eGMCommands.SummonChar              '/SUM
            Call HandleSummonChar(UserIndex)
        
        Case eGMCommands.SpawnListRequest        '/CC
            Call HandleSpawnListRequest(UserIndex)
        
        Case eGMCommands.SpawnCreature           'SPA
            Call HandleSpawnCreature(UserIndex)
        
        Case eGMCommands.ResetNPCInventory       '/RESETINV
            Call HandleResetNPCInventory(UserIndex)
        
        Case eGMCommands.CleanWorld              '/LIMPIAR
            Call HandleCleanWorld(UserIndex)
        
        Case eGMCommands.ServerMessage           '/RMSG
            Call HandleServerMessage(UserIndex)
        
        Case eGMCommands.MapMessage              '/MAPMSG
            Call HandleMapMessage(UserIndex)
            
        Case eGMCommands.NickToIP                '/NICK2IP
            Call HandleNickToIP(UserIndex)
        
        Case eGMCommands.IPToNick                '/IP2NICK
            Call HandleIPToNick(UserIndex)
        
        Case eGMCommands.TeleportCreate          '/CT
            Call HandleTeleportCreate(UserIndex)
        
        Case eGMCommands.TeleportDestroy         '/DT
            Call HandleTeleportDestroy(UserIndex)
        
        Case eGMCommands.RainToggle              '/LLUVIA
            Call HandleRainToggle(UserIndex)
        
        Case eGMCommands.SetCharDescription      '/SETDESC
            Call HandleSetCharDescription(UserIndex)
        
        Case eGMCommands.ForceMIDIToMap          '/FORCEMIDIMAP
            Call HanldeForceMIDIToMap(UserIndex)
        
        Case eGMCommands.ForceWAVEToMap          '/FORCEWAVMAP
            Call HandleForceWAVEToMap(UserIndex)
        
        Case eGMCommands.RoyalArmyMessage        '/REALMSG
            Call HandleRoyalArmyMessage(UserIndex)
        
        Case eGMCommands.ChaosLegionMessage      '/CAOSMSG
            Call HandleChaosLegionMessage(UserIndex)
        
        Case eGMCommands.TalkAsNPC               '/TALKAS
            Call HandleTalkAsNPC(UserIndex)
        
        Case eGMCommands.DestroyAllItemsInArea   '/MASSDEST
            Call HandleDestroyAllItemsInArea(UserIndex)
        
        Case eGMCommands.AcceptRoyalCouncilMember '/ACEPTCONSE
            Call HandleAcceptRoyalCouncilMember(UserIndex)
        
        Case eGMCommands.AcceptChaosCouncilMember '/ACEPTCONSECAOS
            Call HandleAcceptChaosCouncilMember(UserIndex)
        
        Case eGMCommands.ItemsInTheFloor         '/PISO
            Call HandleItemsInTheFloor(UserIndex)
        
        Case eGMCommands.MakeDumb                '/ESTUPIDO
            Call HandleMakeDumb(UserIndex)
        
        Case eGMCommands.MakeDumbNoMore          '/NOESTUPIDO
            Call HandleMakeDumbNoMore(UserIndex)
        
        Case eGMCommands.DumpIPTables            '/DUMPSECURITY
            Call HandleDumpIPTables(UserIndex)
        
        Case eGMCommands.CouncilKick             '/KICKCONSE
            Call HandleCouncilKick(UserIndex)
        
        Case eGMCommands.SetTrigger              '/TRIGGER
            Call HandleSetTrigger(UserIndex)
        
        Case eGMCommands.AskTrigger              '/TRIGGER with no args
            Call HandleAskTrigger(UserIndex)
        
        Case eGMCommands.BannedIPList            '/BANIPLIST
            Call HandleBannedIPList(UserIndex)
        
        Case eGMCommands.BannedIPReload          '/BANIPRELOAD
            Call HandleBannedIPReload(UserIndex)
        
        Case eGMCommands.BanIP                   '/BANIP
            Call HandleBanIP(UserIndex)
        
        Case eGMCommands.UnbanIP                 '/UNBANIP
            Call HandleUnbanIP(UserIndex)
        
        Case eGMCommands.CreateItem              '/CI
            Call HandleCreateItem(UserIndex)
        
        Case eGMCommands.DestroyItems            '/DEST
            Call HandleDestroyItems(UserIndex)
        
        Case eGMCommands.FactionKick         '/ECHARFACCION
            Call HandleFactionKick(UserIndex)
        
        Case eGMCommands.ForceMIDIAll            '/FORCEMIDI
            Call HandleForceMIDIAll(UserIndex)
        
        Case eGMCommands.ForceWAVEAll            '/FORCEWAV
            Call HandleForceWAVEAll(UserIndex)
        
        Case eGMCommands.RemovePunishment        '/BORRARPENA
            Call HandleRemovePunishment(UserIndex)
        
        Case eGMCommands.TileBlockedToggle       '/BLOQ
            Call HandleTileBlockedToggle(UserIndex)
        
        Case eGMCommands.KillNPCNoRespawn        '/MATA
            Call HandleKillNPCNoRespawn(UserIndex)
        
        Case eGMCommands.KillAllNearbyNPCs       '/MASSKILL
            Call HandleKillAllNearbyNPCs(UserIndex)
        
        Case eGMCommands.LastIP                  '/LASTIP
            Call HandleLastIP(UserIndex)
        
        Case eGMCommands.ChangeMOTD              '/MOTDCAMBIA
            Call HandleChangeMOTD(UserIndex)
        
        Case eGMCommands.SetMOTD                 'ZMOTD
            Call HandleSetMOTD(UserIndex)
        
        Case eGMCommands.SystemMessage           '/SMSG
            Call HandleSystemMessage(UserIndex)
        
        Case eGMCommands.CreateNPC               '/ACC
            Call HandleCreateNPC(UserIndex)
        
        Case eGMCommands.CreateNPCWithRespawn    '/RACC
            Call HandleCreateNPCWithRespawn(UserIndex)
        
        Case eGMCommands.ImperialArmour          '/AI1 - 4
            Call HandleImperialArmour(UserIndex)
        
        Case eGMCommands.ChaosArmour             '/AC1 - 4
            Call HandleChaosArmour(UserIndex)
        
        Case eGMCommands.NavigateToggle          '/NAVE
            Call HandleNavigateToggle(UserIndex)
        
        Case eGMCommands.ServerOpenToUsersToggle '/HABILITAR
            Call HandleServerOpenToUsersToggle(UserIndex)
        
        Case eGMCommands.ResetFactions           '/RAJAR
            Call HandleResetFactions(UserIndex)
       
       Case eGMCommands.AdminChangeGuildAlign
            Call HandleAdminChangeGuildAlign(UserIndex)
            
        Case eGMCommands.GuildMemberList
            Call HandleAdminGuildMembers(UserIndex)
            
        Case eGMCommands.GuildOnlineMembers
            Call HandleAdminGuildOnlineMembers(UserIndex)
            
        Case eGMCommands.RemoveCharFromGuild
            Call HandleRemoveCharFromGuild(UserIndex)
            
        Case eGMCommands.ModGuildContribution    '/MODCLANCONTRI
            Call HandleModGuildContribution(UserIndex)
       
        Case eGMCommands.RequestCharMail         '/LASTEMAIL
            Call HandleRequestCharMail(UserIndex)
        
        Case eGMCommands.AlterPassword           '/APASS
            Call HandleAlterPassword(UserIndex)
        
        Case eGMCommands.AlterMail               '/AEMAIL
            Call HandleAlterMail(UserIndex)
        
        Case eGMCommands.AlterName               '/ANAME
            Call HandleAlterName(UserIndex)
        
        Case Declaraciones.eGMCommands.DoBackUp               '/DOBACKUP
            Call HandleDoBackUp(UserIndex)
        
        Case eGMCommands.SaveMap                 '/GUARDAMAPA
            Call HandleSaveMap(UserIndex)
        
        Case eGMCommands.ChangeMapInfoPK         '/MODMAPINFO PK
            Call HandleChangeMapInfoPK(UserIndex)
            
        Case eGMCommands.ChangeMapInfoBackup     '/MODMAPINFO BACKUP
            Call HandleChangeMapInfoBackup(UserIndex)
        
        Case eGMCommands.ChangeMapInfoRestricted '/MODMAPINFO RESTRINGIR
            Call HandleChangeMapInfoRestricted(UserIndex)
        
        Case eGMCommands.ChangeMapInfoNoMagic    '/MODMAPINFO MAGIASINEFECTO
            Call HandleChangeMapInfoNoMagic(UserIndex)
        
        Case eGMCommands.ChangeMapInfoNoInvi     '/MODMAPINFO INVISINEFECTO
            Call HandleChangeMapInfoNoInvi(UserIndex)
        
        Case eGMCommands.ChangeMapInfoNoResu     '/MODMAPINFO RESUSINEFECTO
            Call HandleChangeMapInfoNoResu(UserIndex)
        
        Case eGMCommands.ChangeMapInfoLand       '/MODMAPINFO TERRENO
            Call HandleChangeMapInfoLand(UserIndex)
        
        Case eGMCommands.ChangeMapInfoZone       '/MODMAPINFO ZONA
            Call HandleChangeMapInfoZone(UserIndex)
        
        Case eGMCommands.ChangeMapInfoStealNpc   '/MODMAPINFO ROBONPC
            Call HandleChangeMapInfoStealNpc(UserIndex)
            
        Case eGMCommands.ChangeMapInfoNoOcultar  '/MODMAPINFO OCULTARSINEFECTO
            Call HandleChangeMapInfoNoOcultar(UserIndex)
            
        Case eGMCommands.ChangeMapInfoNoInvocar  '/MODMAPINFO INVOCARSINEFECTO
            Call HandleChangeMapInfoNoInvocar(UserIndex)
            
        Case eGMCommands.SaveChars               '/GRABAR
            Call HandleSaveChars(UserIndex)
        
        Case eGMCommands.CleanSOS                '/BORRAR SOS
            Call HandleCleanSOS(UserIndex)
        
        Case eGMCommands.ShowServerForm          '/SHOW INT
            Call HandleShowServerForm(UserIndex)

        Case eGMCommands.KickAllChars            '/ECHARTODOSPJS
            Call HandleKickAllChars(UserIndex)
        
        Case eGMCommands.ReloadNPCs              '/RELOADNPCS
            Call HandleReloadNPCs(UserIndex)
        
        Case eGMCommands.ReloadServerIni         '/RELOADSINI
            Call HandleReloadServerIni(UserIndex)
        
        Case eGMCommands.ReloadSpells            '/RELOADHECHIZOS
            Call HandleReloadSpells(UserIndex)
        
        Case eGMCommands.ReloadObjects           '/RELOADOBJ
            Call HandleReloadObjects(UserIndex)
        
        Case eGMCommands.Restart                 '/REINICIAR
            Call HandleRestart(UserIndex)
        
        Case eGMCommands.ResetAutoUpdate         '/AUTOUPDATE
            Call HandleResetAutoUpdate(UserIndex)
        
        Case eGMCommands.ChatColor               '/CHATCOLOR
            Call HandleChatColor(UserIndex)
        
        Case eGMCommands.Ignored                 '/IGNORADO
            Call HandleIgnored(UserIndex)
        
        Case eGMCommands.CheckSlot               '/SLOT
            Call HandleCheckSlot(UserIndex)
        
        Case eGMCommands.SetIniVar               '/SETINIVAR LLAVE CLAVE VALOR
            Call HandleSetIniVar(UserIndex)
            
        Case eGMCommands.CreatePretorianClan     '/CREARPRETORIANOS
            Call HandleCreatePretorianClan(UserIndex)
         
        Case eGMCommands.RemovePretorianClan     '/ELIMINARPRETORIANOS
            Call HandleDeletePretorianClan(UserIndex)
                
        Case eGMCommands.EnableDenounces         '/DENUNCIAS
            Call HandleEnableDenounces(UserIndex)
            
        Case eGMCommands.ShowDenouncesList       '/SHOW DENUNCIAS
            Call HandleShowDenouncesList(UserIndex)
        
        Case eGMCommands.SetDialog               '/SETDIALOG
            Call HandleSetDialog(UserIndex)
            
        Case eGMCommands.Impersonate             '/IMPERSONAR
            Call HandleImpersonate(UserIndex)
            
        Case eGMCommands.Imitate                 '/MIMETIZAR
            Call HandleImitate(UserIndex)
            
        Case eGMCommands.RecordAdd
            Call HandleRecordAdd(UserIndex)
            
        Case eGMCommands.RecordAddObs
            Call HandleRecordAddObs(UserIndex)
            
        Case eGMCommands.RecordRemove
            Call HandleRecordRemove(UserIndex)
            
        Case eGMCommands.RecordListRequest
            Call HandleRecordListRequest(UserIndex)
            
        Case eGMCommands.RecordDetailsRequest
            Call HandleRecordDetailsRequest(UserIndex)
        
        Case eGMCommands.HigherAdminsMessage
            Call HandleHigherAdminsMessage(UserIndex)
            
        Case eGMCommands.PMSend
            Call HandlePMSend(UserIndex)
            
        Case eGMCommands.PMDeleteUser
            Call HandlePMDeleteUser(UserIndex)
            
        Case eGMCommands.PMListUser
            Call HandlePMListUser(UserIndex)
        
        Case eGMCommands.RequestTournamentCompetitors
            Call HandleRequestTournamentCompetitors(UserIndex)
        
        Case eGMCommands.Descalificar
            Call HandleTournamentDisqualify(UserIndex)
        
        Case eGMCommands.Pelea
            Call HandleTournamentFight(UserIndex)
        
        Case eGMCommands.CerrarTorneo
            Call HandleTorunamentCancel(UserIndex)
        
        Case eGMCommands.IniciarTorneo
            Call HandleTorunamentBegin(UserIndex)
        
        Case eGMCommands.RequestTournamentConfig
            Call HandleRequestTournamentConfig(UserIndex)
        
        Case eGMCommands.GetPunishmenttypelist
            Call HandleGetPunishmentList(UserIndex)
            
        Case eGMCommands.ChangeMapInfoNoInmo
            Call HandleChangeMapInfoNoInmo(UserIndex)
            
        Case eGMCommands.ChangeMapInfoMismoBando
            Call HandleChangeMapInfoMismoBando(UserIndex)
            
        Case eGMCommands.ShowGuildMessages
            Call HandleShowGuildMessages(UserIndex)
            
        Case eGMCommands.Forgive
            Call HandleForgive(UserIndex)
    End Select
End With

Exit Sub


    Call LogError("Error en GmCommands. Error: " & Err.Number & " - " & Err.Description & _
                  ". Paquete: " & Command)

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleGMCommands de Protocol.bas")
End Sub

''
' Handles the "Home" message.
'
' @param    userIndex The index of the user sending the message.
Private Sub HandleHome(ByVal UserIndex As Integer)
'***************************************************
'Author: Budi
'Creation Date: 06/01/2010
'Last Modification: 05/06/10
'Pato - 05/06/10: Add the Ucase$ to prevent problems.
'***************************************************
On Error GoTo ErrHandler
  
With UserList(UserIndex)
    
    Dim Cancelled As Boolean
    Cancelled = False
    
    ' Setting home.
    If .flags.TargetNpcTipo = eNPCType.Gobernador Then
        Call EndTravel(UserIndex, False)
        Call setHome(UserIndex, Npclist(.flags.TargetNpc).Ciudad, .flags.TargetNpc)
        Exit Sub
    End If
    
    If .flags.Muerto = 0 Then
        Call EndTravel(UserIndex, False)
        Call WriteConsoleMsg(UserIndex, "Debes estar muerto para utilizar este comando.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If .Hogar <= 0 Then
        Call EndTravel(UserIndex, False)
        Call WriteConsoleMsg(UserIndex, "No tienes ningún hogar guardado. Habla con un Gobernador para guardar tu hogar.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If .flags.DueloIndex > 0 Then
        Call EndTravel(UserIndex, False)
        Call WriteConsoleMsg(UserIndex, "No puedes usar este comando durante un duelo.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
        
    'Si es un mapa común y no está en cana
    If (MapInfo(.Pos.Map).Restringir <> eRestrict.restrict_no) Or (.Counters.Pena > 0) Then
        Call EndTravel(UserIndex, False)
        Call WriteConsoleMsg(UserIndex, "No puedes usar este comando aquí.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If .flags.Traveling = 1 Then
        Call EndTravel(UserIndex, True)
        Exit Sub
    End If
    
    If Ciudades(.Hogar).Map = .Pos.Map Then
        Call EndTravel(UserIndex, False)
        Call WriteConsoleMsg(UserIndex, "Ya te encuentras en tu hogar.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    Call goHome(UserIndex)
      
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleHome de Protocol.bas")
End Sub

''
' Handles the "LoginExistingChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLoginExistingChar(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

On Error GoTo ErrHandler

    Dim UserName As String
    Dim uName As String
    Dim Password As String
    Dim version As String
    
    With UserList(UserIndex)
        
        UserName = Reader.ReadString8
        
        If InStr(1, UserName, vbNullChar) > 0 Then
            uName = ReadField(1, UserName, Asc(vbNullChar))
            UserName = ReadField(2, UserName, Asc(vbNullChar))
        Else
            uName = UserName
        End If
        
        Password = Reader.ReadString8
        
        'Convert version number to string
        version = CStr(Reader.ReadInt8()) & "." & CStr(Reader.ReadInt8()) & "." & CStr(Reader.ReadInt8())
        
        If Not AsciiValidos(UserName, False) Then
            Call DisconnectWithMessage(UserIndex, "Nombre inválido.")
            Exit Sub
        End If
        
        If Not PersonajeExiste(UserName) Then
            Call DisconnectWithMessage(UserIndex, "El personaje no existe.")
            Exit Sub
        End If
        
        Dim bConFailed As Boolean

#If SeguridadTesteo Then
        If Not PermiteIngresarUser(UserIndex, Reader.ReadString8()) Then
            Call WriteErrorMsg(UserIndex, "No estas autorizado para ingresar!")
        Else
#ElseIf EnableSecurity Then
        If Not MD5ok(Reader.ReadString8()) Then
            Call WriteErrorMsg(UserIndex, "El cliente está dañado, por favor descarguelo nuevamente desde www.argentumonline.com.ar")
        Else
#End If
        If Not VersionOK(version) Then
                Call WriteErrorMsg(UserIndex, "Esta versión del juego es obsoleta, la versión correcta es la " & ULTIMAVERSION & ". La misma se encuentra disponible en www.argentumonline.com.ar")
            Else
                bConFailed = Not ConnectUser(UserIndex, UserName, Password, uName, False)
            End If
    #If EnableSecurity Or SeguridadTesteo Then
        End If
    #End If
    
    End With
    
    Exit Sub
    
ErrHandler:

    Err.Raise Err.Number

End Sub

''
' Handles the "Talk" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTalk(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 13/01/2010
'15/07/2009: ZaMa - Now invisible admins talk by console.
'23/09/2009: ZaMa - Now invisible admins can't send empty chat.
'13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim Chat As String
        
        Chat = Reader.ReadString8()

        '[Consejeros & GMs]
        If .flags.Privilegios And (PlayerType.Consejero Or PlayerType.SemiDios) Then
            Call LogGM(.Name, "Dijo: " & Chat)
        End If
        
        'I see you....
        If .flags.Oculto > 0 Then
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0
            
            If Not ThiefRestoreBoatAppearance(UserIndex) And .flags.invisible = 0 Then
                Call UsUaRiOs.SetInvisible(UserIndex, UserList(UserIndex).Char.CharIndex, False)
                Call WriteConsoleMsg(UserIndex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
                
                ' Enable the Berzerk if needed
                If BerzerkConditionMet(UserIndex) Then
                    Call ActivatePassive(UserIndex, ePassiveSpells.Berserk, True)
                    Call SendBerserkEffect(UserIndex, ePassiveSpells.Berserk, True)
                End If
                
            End If
        End If
        
        If LenB(Chat) <> 0 Then
            'Analize chat...
            Call Statistics.ParseChat(Chat)
            
            If Not (.flags.AdminInvisible = 1) Then
                If .flags.Muerto = 1 Then
                    Call SendData(SendTarget.ToDeadArea, UserIndex, PrepareMessageChatOverHead(Chat, .Char.CharIndex, CHAT_COLOR_DEAD_CHAR))
                Else
                    If EsGm(UserIndex) Then
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(Chat, .Char.CharIndex, .flags.ChatColor))
                    Else
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatPersonalizado(Chat, .Char.CharIndex, 1))
                    End If
                End If
            End If
        End If
    End With
    
    Exit Sub
    
ErrHandler:

    Err.Raise Err.Number
End Sub

''
' Handles the "Yell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleYell(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 13/01/2010 (ZaMa)
'15/07/2009: ZaMa - Now invisible admins yell by console.
'13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)
    
        Dim Chat As String
        
        Chat = Reader.ReadString8()

        '[Consejeros & GMs]
        If .flags.Privilegios And (PlayerType.Consejero Or PlayerType.SemiDios) Then
            Call LogGM(.Name, "Grito: " & Chat)
        End If
            
        'I see you....
        If .flags.Oculto > 0 Then
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0
            
            If .flags.Navegando = 1 Then
                If .clase = eClass.Thief Then
                    ' Pierde la apariencia de fragata fantasmal
                    Call ToggleBoatBody(UserIndex)
                    Call WriteConsoleMsg(UserIndex, "¡Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                    Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.heading, ConstantesGRH.NingunArma, _
                                        ConstantesGRH.NingunEscudo, ConstantesGRH.NingunCasco)
                End If
            Else
                If .flags.invisible = 0 Then
                    Call UsUaRiOs.SetInvisible(UserIndex, .Char.CharIndex, False)
                    Call WriteConsoleMsg(UserIndex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
            
        If LenB(Chat) <> 0 Then
            'Analize chat...
            Call Statistics.ParseChat(Chat)
                
            If .flags.Privilegios And PlayerType.User Then
                If UserList(UserIndex).flags.Muerto = 1 Then
                    Call SendData(SendTarget.ToDeadArea, UserIndex, PrepareMessageChatOverHead(Chat, .Char.CharIndex, CHAT_COLOR_DEAD_CHAR))
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatPersonalizado(Chat, .Char.CharIndex, 4))
                End If
            Else
                If Not (.flags.AdminInvisible = 1) Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(Chat, .Char.CharIndex, CHAT_COLOR_GM_YELL))
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageConsoleMsg("Gm> " & Chat, FontTypeNames.FONTTYPE_GM))
                End If
            End If
        End If
    
    End With
    
    Exit Sub
    
ErrHandler:

    Err.Raise Err.Number
End Sub

''
' Handles the "Whisper" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWhisper(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 03/12/2010
'28/05/2009: ZaMa - Now it doesn't appear any message when private talking to an invisible admin
'15/07/2009: ZaMa - Now invisible admins wisper by console.
'03/12/2010: Enanoh - Agregué susurro a Admins en modo consulta y Los Dioses pueden susurrar en ciertos casos.
'***************************************************

On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim Chat As String
        Dim TargetUserIndex As Integer
        Dim TargetPriv As PlayerType
        Dim UserPriv As PlayerType
        Dim TargetName As String
        
        TargetName = Reader.ReadString8()
        Chat = Reader.ReadString8()

        UserPriv = .flags.Privilegios
        
        If .flags.Muerto Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Los muertos no pueden comunicarse con el mundo de los vivos. ", FontTypeNames.FONTTYPE_INFO)
        Else
            ' Offline?
            TargetUserIndex = NameIndex(TargetName)
            If TargetUserIndex = INVALID_INDEX Then
                ' Admin?
                If EsGmChar(TargetName) Then
                    Call WriteConsoleMsg(UserIndex, "No puedes susurrarle a los Administradores.", FontTypeNames.FONTTYPE_INFO)
                ' Whisperer admin? (Else say nothing)
                ElseIf (UserPriv And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Then
                    Call WriteConsoleMsg(UserIndex, "Usuario inexistente.", FontTypeNames.FONTTYPE_INFO)
                End If
                
            ' Online
            Else
                ' Privilegios
                TargetPriv = UserList(TargetUserIndex).flags.Privilegios
                
                ' Consejeros, semis y usuarios no pueden susurrar a dioses (Salvo en consulta)
                If (TargetPriv And (PlayerType.Dios Or PlayerType.Admin)) <> 0 And _
                   (UserPriv And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios)) <> 0 And _
                   Not .flags.HelpMode Then
                    
                    ' No puede
                    Call WriteConsoleMsg(UserIndex, "No puedes susurrarle a los Administradores.", FontTypeNames.FONTTYPE_INFO)

                ' Usuarios no pueden susurrar a semis o conses (Salvo en consulta)
                ElseIf (UserPriv And PlayerType.User) <> 0 And _
                       (Not TargetPriv And PlayerType.User) <> 0 And _
                        Not .flags.HelpMode Then
                    
                    ' No puede
                    Call WriteConsoleMsg(UserIndex, "No puedes susurrarle a los Administradores.", FontTypeNames.FONTTYPE_INFO)
                
                ' En rango? (Los dioses pueden susurrar a distancia)
                ElseIf Not EstaPCarea(UserIndex, TargetUserIndex) And _
                    (UserPriv And (PlayerType.Dios Or PlayerType.Admin)) = 0 Then
                    
                    ' No se puede susurrar a admins fuera de su rango
                    If (TargetPriv And (PlayerType.User)) = 0 And (UserPriv And (PlayerType.Dios Or PlayerType.Admin)) = 0 Then
                        Call WriteConsoleMsg(UserIndex, "No puedes susurrarle a los Administradores.", FontTypeNames.FONTTYPE_INFO)
                    
                    ' Whisperer admin? (Else say nothing)
                    ElseIf (UserPriv And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Then
                        Call WriteConsoleMsg(UserIndex, "Estás muy lejos del usuario.", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    '[Consejeros & GMs]
                    If UserPriv And (PlayerType.Consejero Or PlayerType.SemiDios) Then
                        Call LogGM(.Name, "Le susurro a '" & UserList(TargetUserIndex).Name & "' " & Chat)
                    
                    ' Usuarios a administradores
                    ElseIf (UserPriv And PlayerType.User) <> 0 And (TargetPriv And PlayerType.User) = 0 Then
                        Call LogGM(UserList(TargetUserIndex).Name, .Name & " le susurro en consulta: " & Chat)
                    End If
                    
                    If LenB(Chat) <> 0 Then
                        'Analize chat...
                        Call Statistics.ParseChat(Chat)
                        
                        ' Dios susurrando a distancia
                        If Not EstaPCarea(UserIndex, TargetUserIndex) And _
                            (UserPriv And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Then
                            
                            Call WriteConsoleMsg(UserIndex, "Susurraste> " & Chat, FontTypeNames.FONTTYPE_GM)
                            Call WriteConsoleMsg(TargetUserIndex, "Gm susurra> " & Chat, FontTypeNames.FONTTYPE_GM)
                            
                        ElseIf Not (.flags.AdminInvisible = 1) Then
                            Call WriteChatPersonalizado(UserIndex, Chat, .Char.CharIndex, 6)
                            Call WriteChatPersonalizado(TargetUserIndex, Chat, .Char.CharIndex, 6)

                            '[CDT 17-02-2004]
                            If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then
                                Call SendData(SendTarget.ToAdminsAreaButConsejeros, UserIndex, PrepareMessageChatOverHead("A " & UserList(TargetUserIndex).Name & "> " & Chat, .Char.CharIndex, vbYellow))
                            End If
                        Else
                            Call WriteConsoleMsg(UserIndex, "Susurraste> " & Chat, FontTypeNames.FONTTYPE_GM)
                            If UserIndex <> TargetUserIndex Then Call WriteConsoleMsg(TargetUserIndex, "Gm susurra> " & Chat, FontTypeNames.FONTTYPE_GM)
                            
                            If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then
                                Call SendData(SendTarget.ToAdminsAreaButConsejeros, UserIndex, PrepareMessageConsoleMsg("Gm dijo a " & UserList(TargetUserIndex).Name & "> " & Chat, FontTypeNames.FONTTYPE_GM))
                            End If
                        End If
                    End If
                End If
            End If
        End If
    
    End With
    
    Exit Sub
    
ErrHandler:

    Err.Raise Err.Number
End Sub

''
' Handles the "Walk" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWalk(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 13/01/2010 (ZaMa)
'11/19/09 Pato - Now the class bandit can walk hidden.
'13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
'***************************************************

On Error GoTo ErrHandler

    Dim dummy As Long
    Dim TempTick As Long
    Dim heading As eHeading
    
    With UserList(UserIndex)

        heading = Reader.ReadInt8()
        
        'Prevent SpeedHack
        If .flags.TimesWalk >= 30 Then
            TempTick = GetTickCount()
            dummy = GetInterval(TempTick, .flags.StartWalk)
            
            ' 5800 is actually less than what would be needed in perfect conditions to take 30 steps
            '(it's about 193 ms per step against the over 200 needed in perfect conditions)
            If dummy < 5800 Then
                If GetInterval(TempTick, .flags.CountSH) > 30000 Then
                    .flags.CountSH = 0
                End If
                
                If Not .flags.CountSH = 0 Then
                    If dummy <> 0 Then _
                        dummy = 126000 \ dummy
                    
                    Call LogHackAttemp("Tramposo SH: " & .Name & " , " & dummy)
                    Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> " & .Name & " ha sido echado por el servidor por posible uso de SH.", FontTypeNames.FONTTYPE_SERVER, eMessageType.Admin))
                    Call CloseSocket(UserIndex)
                    
                    Exit Sub
                Else
                    .flags.CountSH = TempTick
                End If
            End If
            .flags.StartWalk = TempTick
            .flags.TimesWalk = 0
        End If
        
        .flags.TimesWalk = .flags.TimesWalk + 1
        
        If .flags.TournamentState = eTournamentState.ieWaitingForFight Then
            Call WriteConsoleMsg(UserIndex, "Aún no terminó la cuenta regresiva.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        'If exiting, cancel
        Call CancelExit(UserIndex)
        
        ' Close the crafting store
        If .CraftingStore.IsOpen Then
            Call CloseWorkerStore(UserIndex)
        End If
        
        'TODO: Debería decirle por consola que no puede?
        'Esta usando el /HOGAR, no se puede mover
        If .flags.Traveling = 1 Then Exit Sub
        
        If .flags.Paralizado = 0 Then
            If .flags.Meditando Then
                'Stop meditating, next action will start movement.
                .flags.Meditando = False
                .Char.FX = 0
                .Char.Loops = 0
                
                Call WriteMeditateToggle(UserIndex)
                Call WriteConsoleMsg(UserIndex, "Dejas de meditar.", FontTypeNames.FONTTYPE_INFO, eMessageType.info)
                
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0, , 0))
            Else
                'Move user
                If MoveUserChar(UserIndex, heading) Then
                
                    'Stop resting if needed
                    If .flags.Descansar Then
                        .flags.Descansar = False
                        
                        Call WriteRestOK(UserIndex)
                        Call WriteConsoleMsg(UserIndex, "Has dejado de descansar.", FontTypeNames.FONTTYPE_INFO)
                    End If
                    
                    'Can't move while hidden except he is a thief
                    If .flags.Oculto = 1 And .flags.AdminInvisible = 0 Then
                        If .clase <> eClass.Thief Then
                            .flags.Oculto = 0
                            .Counters.TiempoOculto = 0
                            
                            If Not ThiefRestoreBoatAppearance(UserIndex) And .flags.invisible = 0 Then
                                Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
                                Call UsUaRiOs.SetInvisible(UserIndex, .Char.CharIndex, False)
                            End If
                        End If
                        
                        ' Enable the Berzerk if needed
                        If BerzerkConditionMet(UserIndex) Then
                            Call ActivatePassive(UserIndex, ePassiveSpells.Berserk, True)
                            Call SendBerserkEffect(UserIndex, ePassiveSpells.Berserk, True)
                        End If
                    End If

                    ' Step into trap?
                    Call CheckTriggerActivation(UserIndex, 0, .Pos.Map, .Pos.X, .Pos.Y, True)
                    
                End If
            End If
        Else    'paralized
            If Not .flags.UltimoMensaje = 1 Then
                .flags.UltimoMensaje = 1
                
                Call WriteConsoleMsg(UserIndex, "No puedes moverte porque estás paralizado.", FontTypeNames.FONTTYPE_INFO)
            End If
            
            .flags.CountSH = 0
        End If
    End With
    
    Exit Sub
    
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleWalk de Protocol.bas")
End Sub

''
' Handles the "RequestPositionUpdate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestPositionUpdate(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
On Error GoTo ErrHandler

    Call WritePosUpdate(UserIndex)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleRequestPositionUpdate de Protocol.bas")
End Sub

''
' Handles the "Attack" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAttack(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 13/01/2010
'Last Modified By: ZaMa
'10/01/2008: Tavo - Se cancela la salida del juego si el user esta saliendo.
'13/11/2009: ZaMa - Se cancela el estado no atacable al atcar.
'13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)
        Dim AttackSucceeded As Boolean

        'If dead, can't attack
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'If user meditates, can't attack
        If .flags.Meditando Then
            Exit Sub
        End If
        
        'Conses can't attack.
        If .flags.Privilegios And PlayerType.Consejero Then Exit Sub
      
        'If equiped weapon is ranged, can't attack this way
        If .Invent.WeaponEqpObjIndex > 0 Then
            If ObjData(.Invent.WeaponEqpObjIndex).proyectil = 1 Then
                Call WriteConsoleMsg(UserIndex, "No puedes usar así este arma.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        
        'If exiting, cancel
        Call CancelExit(UserIndex)
        
        'Attack!
        AttackSucceeded = UsuarioAtaca(UserIndex)
        
        'Now you can be atacked
        .flags.NoPuedeSerAtacado = False
        
        'I see you...
        If .flags.Oculto > 0 And .flags.AdminInvisible = 0 Then
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0
            
            If Not ThiefRestoreBoatAppearance(UserIndex) And .flags.invisible = 0 Then
                Call UsUaRiOs.SetInvisible(UserIndex, .Char.CharIndex, False)
                Call WriteConsoleMsg(UserIndex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        Call WriteAttackResult(UserIndex, AttackSucceeded)
    End With
  
  Exit Sub
  
ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleAttack de Protocol.bas")
End Sub

''
' Handles the "PickUp" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePickUp(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 07/25/09
'02/26/2006: Marco - Agregué un checkeo por si el usuario trata de agarrar un item mientras comercia.
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        'If the user is trading, he can't pickup items
        If Not DropAllowed(UserIndex) Then Exit Sub
        
        'Lower rank administrators can't pick up items
        If .flags.Privilegios And PlayerType.Consejero Then
            If Not .flags.Privilegios And PlayerType.RoleMaster Then
                Call WriteConsoleMsg(UserIndex, "No puedes tomar ningún objeto.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        
        Call GetObj(UserIndex)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandlePickUp de Protocol.bas")
End Sub

''
' Handles the "SafeToggle" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSafeToggle(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 27/12/2014
'
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        If .flags.Seguro Then
            Call WriteMultiMessage(UserIndex, eMessages.SafeModeOff) 'Call WriteSafeModeOff(UserIndex)
        Else
            Call WriteMultiMessage(UserIndex, eMessages.SafeModeOn) 'Call WriteSafeModeOn(UserIndex)
        End If
        
        .flags.Seguro = Not .flags.Seguro
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleSafeToggle de Protocol.bas")
End Sub

''
' Handles the "ResuscitationSafeToggle" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleResuscitationToggle(ByVal UserIndex As Integer)
'***************************************************
'Author: Rapsodius
'Creation Date: 10/10/07
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        .flags.SeguroResu = Not .flags.SeguroResu
        
        If .flags.SeguroResu Then
            Call WriteMultiMessage(UserIndex, eMessages.ResuscitationSafeOn) 'Call WriteResuscitationSafeOn(UserIndex)
        Else
            Call WriteMultiMessage(UserIndex, eMessages.ResuscitationSafeOff) 'Call WriteResuscitationSafeOff(UserIndex)
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleResuscitationToggle de Protocol.bas")
End Sub

''
' Handles the "RequestSkills" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestSkills(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
On Error GoTo ErrHandler

    Call WriteSendSkills(UserIndex)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleRequestSkills de Protocol.bas")
End Sub

''
' Handles the "RequestStadictics" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestStadictics(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 24/05/2012
'
'***************************************************
    'Remove packet ID
On Error GoTo ErrHandler

    Call WriteMiniStats(UserIndex)
    Call WriteAttributes(UserIndex)
    Call WriteSendSkills(UserIndex)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleRequestStadictics de Protocol.bas")
End Sub

''
' Handles the "CommerceEnd" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceEnd(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
On Error GoTo ErrHandler

    'User quits commerce mode
    UserList(UserIndex).flags.Comerciando = 0
    Call WriteCommerceEnd(UserIndex)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleCommerceEnd de Protocol.bas")
End Sub

''
' Handles the "UserCommerceEnd" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceEnd(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 11/03/2010
'11/03/2010: ZaMa - Le avisa por consola al que cencela que dejo de comerciar.
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)
        Dim tempUsu As Integer

        tempUsu = getTradingUser(UserIndex)
        
        'Quits commerce mode with user
        If tempUsu > 0 Then
            If getTradingUser(tempUsu) = UserIndex Then
                Call WriteConsoleMsg(tempUsu, .Name & " ha dejado de comerciar con vos.", FontTypeNames.FONTTYPE_TALK)
                Call FinComerciarUsu(tempUsu)
            End If
        End If
        
        Call FinComerciarUsu(UserIndex)
        Call WriteConsoleMsg(UserIndex, "Has dejado de comerciar.", FontTypeNames.FONTTYPE_TALK)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleUserCommerceEnd de Protocol.bas")
End Sub

''
' Handles the "UserCommerceConfirm" message.
'
' @param    userIndex The index of the user sending the message.
Private Sub HandleUserCommerceConfirm(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 14/12/2009
'
'***************************************************
On Error GoTo ErrHandler
  
    'Validate the commerce
    If PuedeSeguirComerciando(UserIndex) Then
        'Tell the other user the confirmation of the offer
        Call WriteUserOfferConfirm(getTradingUser(UserIndex))
        UserList(UserIndex).ComUsu.Confirmo = True
    End If
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleUserCommerceConfirm de Protocol.bas")
End Sub

Private Sub HandleCommerceChat(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 03/12/2009
'
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim Chat As String
        
        Chat = Reader.ReadString8()
  
        If LenB(Chat) <> 0 Then
            If PuedeSeguirComerciando(UserIndex) Then
                'Analize chat...
                Call Statistics.ParseChat(Chat)
                
                Chat = UserList(UserIndex).Name & "> " & Chat
                Call WriteCommerceChat(UserIndex, Chat, FontTypeNames.FONTTYPE_PARTY)
                Call WriteCommerceChat(getTradingUser(UserIndex), Chat, FontTypeNames.FONTTYPE_PARTY)
            End If
        End If
        

    End With
  
  Exit Sub
  
ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleCommerceChat de Protocol.bas")
End Sub


''
' Handles the "BankEnd" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankEnd(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        'User exits banking mode
        .flags.Comerciando = 0
        Call WriteBankEnd(UserIndex)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleBankEnd de Protocol.bas")
End Sub

''
' Handles the "UserCommerceOk" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceOk(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
On Error GoTo ErrHandler

    'Trade accepted
    Call AceptarComercioUsu(UserIndex)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleUserCommerceOk de Protocol.bas")
End Sub

''
' Handles the "UserCommerceReject" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceReject(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
On Error GoTo ErrHandler
  
    Dim OtherUser As Integer
    
    With UserList(UserIndex)

        OtherUser = getTradingUser(UserIndex)
        
        'Offer rejected
        If OtherUser > 0 Then
            If UserList(OtherUser).flags.UserLogged Then
                Call WriteConsoleMsg(OtherUser, .Name & " ha rechazado tu oferta.", FontTypeNames.FONTTYPE_TALK)
                Call FinComerciarUsu(OtherUser)
            End If
        End If
        
        Call WriteConsoleMsg(UserIndex, "Has rechazado la oferta del otro usuario.", FontTypeNames.FONTTYPE_TALK)
        Call FinComerciarUsu(UserIndex)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleUserCommerceReject de Protocol.bas")
End Sub

''
' Handles the "Drop" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDrop(ByVal UserIndex As Integer)
On Error GoTo ErrHandler
      
    Dim Slot As Byte
    Dim Amount As Integer
    
    With UserList(UserIndex)

        Slot = Reader.ReadInt8()
        Amount = Reader.ReadInt16()
        

        'low rank admins can't drop item. Neither can the dead.
        If .flags.Muerto = 1 Or _
           ((.flags.Privilegios And PlayerType.Consejero) <> 0 And (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0) Then Exit Sub

        'If the user is trading, he can't drop items
        If Not DropAllowed(UserIndex) Then Exit Sub

        'Are we dropping gold or other items??
        If Slot = FLAGORO Then
            If Amount > 10000 Then Exit Sub 'Don't drop too much gold

            Call TirarOro(Amount, UserIndex)
            
            Call WriteUpdateGold(UserIndex)
        Else
            'Only drop valid slots
            If Slot <= MAX_INVENTORY_SLOTS And Slot > 0 Then
                If .Invent.Object(Slot).ObjIndex = 0 Then
                    Exit Sub
                End If
                
                Call DropObj(UserIndex, Slot, Amount, .Pos.Map, .Pos.X, .Pos.Y, True)
            End If
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleDrop de Protocol.bas")
End Sub

Private Sub HandleDropXY(ByVal UserIndex As Integer)
On Error GoTo ErrHandler

    Dim Slot As Byte
    Dim Amount As Integer
    Dim X As Byte
    Dim Y As Byte
    
    With UserList(UserIndex)

        Slot = Reader.ReadInt8()
        Amount = Reader.ReadInt16()
        X = Reader.ReadInt8()
        Y = Reader.ReadInt8()
        
        If MapData(.Pos.Map, X, Y).Blocked Then Exit Sub
        If MapData(.Pos.Map, X, Y).TileExit.Map > 0 Then Exit Sub
        If MapData(.Pos.Map, X, Y).NpcIndex > 0 Then Exit Sub

        'low rank admins can't drop item. Neither can the dead.
        If .flags.Muerto = 1 Or _
           ((.flags.Privilegios And PlayerType.Consejero) <> 0 And (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0) Then Exit Sub

        'If the user is trading, he can't drop items
        If Not DropAllowed(UserIndex, False) Then Exit Sub

        'Are we dropping gold or other items??
        If Slot = FLAGORO Then
            If Amount > 10000 Then Exit Sub 'Don't drop too much gold

            Call TirarOro(Amount, UserIndex)
            
            Call WriteUpdateGold(UserIndex)
        Else
            'Only drop valid slots
            If Slot <= MAX_INVENTORY_SLOTS And Slot > 0 Then
                If .Invent.Object(Slot).ObjIndex = 0 Then
                    Exit Sub
                End If
                
                Call DropObjCloseUser(UserIndex, Slot, Amount, X, Y, True)
            End If
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleDropXY de Protocol.bas")
End Sub

''
' Handles the "CastSpell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCastSpell(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'13/11/2009: ZaMa - Ahora los npcs pueden atacar al usuario si quizo castear un hechizo
'***************************************************

On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        Dim SpellIndex As Byte
        
        SpellIndex = Reader.ReadInt8()
        
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Now you can be atacked
        .flags.NoPuedeSerAtacado = False
        
        If SpellIndex < 1 Then
            .flags.CastedSpellNumber = 0
            .flags.CastedSpellIndex = 0
            Exit Sub
        ElseIf SpellIndex > MAXUSERHECHIZOS Then
            .flags.CastedSpellNumber = 0
            .flags.CastedSpellIndex = 0
            Exit Sub
        End If
                    
        .flags.CastedSpellNumber = .Stats.UserHechizos(SpellIndex).SpellNumber
        .flags.CastedSpellIndex = SpellIndex
        
        ' Let the user know it can launch the selected spell
        'If exiting, cancel
        Call CancelExit(UserIndex)
        
        Call WriteMultiMessage(UserIndex, eMessages.SpellCastRequestTarget, SpellIndex, .flags.CastedSpellNumber)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleCastSpell de Protocol.bas")
End Sub

''
' Handles the "LeftClick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLeftClick(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

On Error GoTo ErrHandler

    With Reader
        Dim X As Byte
        Dim Y As Byte
        
        X = .ReadInt8()
        Y = .ReadInt8()
#If EnableSecurity Then
    Call Security.CheckClick(UserIndex, X, Y, 1)
#End If
        Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleLeftClick de Protocol.bas")
End Sub

''
' Handles the "DoubleClick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDoubleClick(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

On Error GoTo ErrHandler

    With Reader

        Dim X As Byte
        Dim Y As Byte
        
        X = .ReadInt8()
        Y = .ReadInt8()
        
        Call Accion(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleDoubleClick de Protocol.bas")
End Sub

''
' Handles the "RightClick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRightClick(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 10/05/2011
'
'***************************************************

On Error GoTo ErrHandler
      
    With Reader
        Dim X As Byte
        Dim Y As Byte
        
        X = .ReadInt8()
        Y = .ReadInt8()
        
        Call Extra.ShowMenu(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleRightClick de Protocol.bas")
End Sub

''
' Handles the "Work" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWork(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 13/04/2014 (D'Artagnan)
'13/01/2010: ZaMa - El pirata se puede ocultar en barca
'13/04/2014: D'Artagnan - Interval checking
'***************************************************

On Error GoTo ErrHandler

    With UserList(UserIndex)
        Dim skill As eSkill
        Dim dwCurrentTicks As Long
        
        skill = Reader.ReadInt8()
        
        If UserList(UserIndex).flags.Muerto = 1 Then Exit Sub
        
        'If exiting, cancel
        Call CancelExit(UserIndex)
        
        Select Case skill
        
            Case Robar, Domar
                Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, skill)
                
            Case Ocultarse
                
                ' Verifico si se peude ocultar en este mapa
                If MapInfo(.Pos.Map).OcultarSinEfecto = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡Ocultarse no funciona aquí!", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If .flags.HelpMode Then
                    Call WriteConsoleMsg(UserIndex, "No puedes ocultarte si estás en consulta.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If .flags.DueloIndex > 0 Then
                    Call WriteConsoleMsg(UserIndex, "No puedes ocultarte durante un duelo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            
                If .flags.Navegando = 1 Then
                    If .clase <> eClass.Thief Then
                        '[CDT 17-02-2004]
                        If Not .flags.UltimoMensaje = 3 Then
                            Call WriteConsoleMsg(UserIndex, "No puedes ocultarte si estás navegando.", FontTypeNames.FONTTYPE_INFO)
                            .flags.UltimoMensaje = 3
                        End If
                        '[/CDT]
                        Exit Sub
                    End If
                End If
                
                If .flags.Oculto = 1 Then
                    '[CDT 17-02-2004]
                    If Not .flags.UltimoMensaje = 2 Then
                        Call WriteConsoleMsg(UserIndex, "Ya estás oculto.", FontTypeNames.FONTTYPE_INFO)
                        .flags.UltimoMensaje = 2
                    End If
                    '[/CDT]
                    Exit Sub
                End If
                
                dwCurrentTicks = GetTickCount()
               
                ' Check if the timer is initialized
                If .Counters.TimerHide Then
                    ' Interval checking
                    If Not checkInterval(.Counters.TimerHide, dwCurrentTicks, ServerConfiguration.Intervals.IntervaloOcultar) Then _
                        Exit Sub
                End If
               
                ' Update timer
                .Counters.TimerHide = dwCurrentTicks
                
                Call DoOcultarse(UserIndex)
                
        End Select
        
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleWork de Protocol.bas")
End Sub

''
' Handles the "InitCrafting" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInitCrafting(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 29/01/2010
'
'***************************************************

On Error GoTo ErrHandler
  
    Dim TotalItems As Long
    
    With UserList(UserIndex)

        TotalItems = Reader.ReadInt32
        
        If TotalItems > 0 Then
            .Construir.Cantidad = TotalItems
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleInitCrafting de Protocol.bas")
End Sub

''
' Handles the "UseSpellMacro" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUseSpellMacro(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        Call SendData(SendTarget.ToAdmins, UserIndex, PrepareMessageConsoleMsg(.Name & " fue expulsado por Anti-macro de hechizos.", FontTypeNames.FONTTYPE_VENENO))
        Call DisconnectWithMessage(UserIndex, "Has sido expulsado por usar macro de hechizos. Recomendamos leer el reglamento sobre el tema macros.")
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleUseSpellMacro de Protocol.bas")
End Sub

''
' Handles the "UseItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUseItem(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        Dim Slot As Byte
        
        Slot = Reader.ReadInt8()
        
        If Slot <= .CurrentInventorySlots And Slot > 0 Then
            If .Invent.Object(Slot).ObjIndex = 0 Then Exit Sub
        End If
        
        If .flags.Meditando Then
            Exit Sub    'The error message should have been provided by the client.
        End If

#If EnableSecurity Then
        Call checkSecurity(UserIndex, 0)
#End If

        Call UseInvItem(UserIndex, Slot)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleUseItem de Protocol.bas")
End Sub

''
' Handles the "WorkLeftClick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWorkLeftClick(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 14/01/2010 (ZaMa)
'16/11/2009: ZaMa - Agregada la posibilidad de extraer madera elfica.
'12/01/2010: ZaMa - Ahora se admiten armas arrojadizas (proyectiles sin municiones).
'14/01/2010: ZaMa - Ya no se pierden municiones al atacar npcs con dueño.
'***************************************************

On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        Dim X As Byte
        Dim Y As Byte
        Dim skill As eSkill
        Dim DummyInt As Integer
        Dim tU As Integer   'Target user
        Dim tN As Integer   'Target NPC
        
        Dim WeaponIndex As Integer
        
        X = Reader.ReadInt8()
        Y = Reader.ReadInt8()
        
        skill = Reader.ReadInt8()
        
#If EnableSecurity Then
    Call Security.CheckClick(UserIndex, X, Y, skill)
#End If
        If .flags.Muerto = 1 Or .flags.Descansar Or .flags.Meditando _
        Or Not InMapBounds(.Pos.Map, X, Y) Then Exit Sub

        If Not InRangoVision(UserIndex, X, Y) Then
            Call WritePosUpdate(UserIndex)
            Exit Sub
        End If
        
        'If exiting, cancel
        Call CancelExit(UserIndex)
        
        Select Case skill
            Case eSkill.Proyectiles
                                
                If Not EsGm(UserIndex) Then
                    
                    'Check attack interval
                    If Not IntervaloPermiteAtacar(UserIndex, False) Then Exit Sub
                    'Check Magic interval
                    If Not IntervaloPermiteLanzarSpell(UserIndex, 0, False) Then Exit Sub
                    'Check bow's interval
                    If Not IntervaloPermiteUsarArcos(UserIndex) Then Exit Sub
                
                End If
                
                Call LanzarProyectil(UserIndex, X, Y)
                            
            Case eSkill.Magia
                'Check the map allows spells to be casted.
                If MapInfo(.Pos.Map).MagiaSinEfecto > 0 Then
                    Call WriteConsoleMsg(UserIndex, "Una fuerza oscura te impide canalizar tu energía.", FontTypeNames.FONTTYPE_FIGHT)
                    Exit Sub
                End If
                
                'Target whatever is in that tile
                Call LookatTile(UserIndex, .Pos.Map, X, Y)
                
                'If it's outside range log it and exit
                If Abs(.Pos.X - X) > RANGO_VISION_X Or Abs(.Pos.Y - Y) > RANGO_VISION_Y Then
                    Call LogCheating("Ataque fuera de rango de " & .Name & "(" & .Pos.Map & "/" & .Pos.X & "/" & .Pos.Y & ") ip: " & .IP & " a la posición (" & .Pos.Map & "/" & X & "/" & Y & ")")
                    Exit Sub
                End If
                
                If .flags.CastedSpellNumber <= 0 Then
                    Call WriteConsoleMsg(UserIndex, "¡Primero selecciona el hechizo que quieres lanzar!", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If Not EsGm(UserIndex) Then
                
                    Dim CanAttack As Boolean
                    CanAttack = IntervaloPermiteAtacar(UserIndex, False)

                    'Check bow's interval
                    If Not IntervaloPermiteUsarArcos(UserIndex, False) Then Exit Sub
        
                    If Not IntervaloPermiteLanzarSpell(UserIndex, .flags.CastedSpellIndex) Then
                        ' If the user attacked previously, and the interval allows the casting of a spell
                        'If Not CanAttack And IntervaloPermiteGolpeMagia(UserIndex, False) Then
                        '    Call LogError("Not CanAttack & IntervaloPermiteGolpeMagia")
                        '    Exit Sub
                        'End If
                        
                            Exit Sub
                        End If
                    End If
                

                Call LanzarHechizo(.flags.CastedSpellNumber, UserIndex) ', X, Y)
                .flags.CastedSpellNumber = 0
                .flags.CastedSpellIndex = 0
                
                .Counters.TimerPuedeAtacar = GetTickCount

            
            Case eSkill.Robar
                'Does the map allow us to steal here?
                If MapInfo(.Pos.Map).Pk Then
                    
                    'Check interval
                    If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                    
                    'Target whatever is in that tile
                    Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
                    
                    tU = .flags.TargetUser
                    
                    If tU > 0 And tU <> UserIndex Then
                        'Can't steal administrative players
                        If UserList(tU).flags.Privilegios And PlayerType.User Then
                            If UserList(tU).flags.Muerto = 0 Then
                                 If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                                     Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                                     Exit Sub
                                 End If
                                 
                                 '17/09/02
                                 'Check the trigger
                                 If MapData(UserList(tU).Pos.Map, X, Y).Trigger = eTrigger.ZONASEGURA Then
                                     Call WriteConsoleMsg(UserIndex, "No puedes robar aquí.", FontTypeNames.FONTTYPE_WARNING)
                                     Exit Sub
                                 End If
                                 
                                 If MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger = eTrigger.ZONASEGURA Then
                                     Call WriteConsoleMsg(UserIndex, "No puedes robar aquí.", FontTypeNames.FONTTYPE_WARNING)
                                     Exit Sub
                                 End If
                                 
                                 Call DoRobar(UserIndex, tU)
                            End If
                        End If
                    Else
                        Call WriteConsoleMsg(UserIndex, "¡No hay a quien robarle!", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    Call WriteConsoleMsg(UserIndex, "¡No puedes robar en zonas seguras!", FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case eSkill.Talar, eSkill.Mineria, eSkill.Pesca
            
                Dim ProfessionIndex As Byte
                
                'Check interval
                If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                 
                DummyInt = MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex
 
                If DummyInt = 0 Then
                    Call WriteConsoleMsg(UserIndex, "No hay ningún recurso extraíble ahí.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                ProfessionIndex = ObjData(DummyInt).ProfessionType
                If ProfessionIndex = 0 Then Exit Sub 'Bad Configuration
                
                'Che if the distance against the resource is between the configured values.
                If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > ObjData(.Invent.WeaponEqpObjIndex).MaxDistanceFromTarget Then
                    Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If Not Professions(ProfessionIndex).Enabled Then
                    Call WriteConsoleMsg(UserIndex, "La profesión se encuentra deshabilitada temporalmente.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If Not Professions(ProfessionIndex).EnabledInSafeZone And MapInfo(.Pos.Map).Pk = False Then
                    Call WriteConsoleMsg(UserIndex, "No puedes extraer este recurso en zona segura", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                WeaponIndex = .Invent.WeaponEqpObjIndex
                
                If WeaponIndex = 0 Then
                    Call WriteConsoleMsg(UserIndex, "No tienes ninguna herramienta equipada", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If ObjData(WeaponIndex).ProfessionType <> ProfessionIndex Then
                    Call WriteConsoleMsg(UserIndex, "La herramienta utilizada no es la adecuada para este recurso.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                ' Is the resource depleted?
                If MapData(.Pos.Map, X, Y).ObjInfo.PendingQty <= 0 Then
                    Call WriteConsoleMsg(UserIndex, "Este recurso se encuentra agotado.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                                
                Dim Pos As WorldPos
                Pos.X = X
                Pos.Y = Y
                Pos.Map = .Pos.Map
                
                If ObjData(WeaponIndex).SoundNumber > 0 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(ObjData(WeaponIndex).SoundNumber, Pos.X, Pos.Y, .Char.CharIndex))
                End If
                Call DoExtractResource(UserIndex, Pos, ProfessionIndex, WeaponIndex)
            
            Case eSkill.Domar
                'Modificado 25/11/02
                'Optimizado y solucionado el bug de la doma de
                'criaturas hostiles.
                
                'Target whatever is that tile
                Call LookatTile(UserIndex, .Pos.Map, X, Y)
                tN = .flags.TargetNpc
                
                If tN > 0 Then
                    If CanTameNpc(UserIndex, tN) Then
                        Call DoTameNpc(UserIndex, tN)
                    End If
                Else
                    Call WriteConsoleMsg(UserIndex, "¡No hay ninguna criatura allí!", FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case FundirMetal    'UGLY!!! This is a constant, not a skill!!
                
                If ObjData(.flags.TargetObjInvIndex).ObjType = eOBJType.otMinerales Then
                    If CanMelt(UserIndex, True) Then
                        Call FundirMineral(UserIndex)
                    End If
                'Nightw: Fundición desactivada hasta nuevo aviso
                'ElseIf ObjData(.flags.TargetObjInvIndex).ObjType = eOBJType.otWeapon Then
                '    If CanMelt(UserIndex, False) Then
                '        Call FundirArmas(UserIndex)
                '    End If
                End If
            
            Case eSkill.Herreria
                'Target wehatever is in that tile
                
                WeaponIndex = .Invent.WeaponEqpObjIndex
                
                If WeaponIndex = 0 Then
                    Call WriteConsoleMsg(UserIndex, "No tienes ninguna herramienta equipada", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If ObjData(WeaponIndex).ProfessionType <> 5 Then
                    Call WriteConsoleMsg(UserIndex, "La herramienta utilizada no es la adecuada para realizar esta acción", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If CanBlacksmith(UserIndex) Then
                    
                    Call WriteCraftableRecipes(UserIndex, ObjData(WeaponIndex).ProfessionType)
                    Call WriteShowCraftForm(UserIndex)
                End If
        End Select
    End With

        Exit Sub

ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleWorkLeftClick de Protocol.bas")
End Sub

''
' Handles the "SpellInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSpellInfo(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

On Error GoTo ErrHandler
      
    With UserList(UserIndex)

        Dim spellSlot As Byte
        Dim Spell As Integer
        
        spellSlot = Reader.ReadInt8()
        
        'Validate slot
        If spellSlot < 1 Or spellSlot > MAXUSERHECHIZOS Then
            Call WriteConsoleMsg(UserIndex, "¡Primero selecciona el hechizo!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate spell in the slot
        Spell = .Stats.UserHechizos(SpellSlot).SpellNumber
        If Spell > 0 And Spell < NumeroHechizos + 1 Then
            With Hechizos(Spell)
                'Send information
                Call WriteConsoleMsg(UserIndex, "%%%%%%%%%%%% INFO DEL HECHIZO %%%%%%%%%%%%" & vbCrLf _
                                               & "Nombre:" & .Nombre & vbCrLf _
                                               & "Descripción:" & .Desc & vbCrLf _
                                               & "Skill requerido: " & .MinSkill & " de magia." & vbCrLf _
                                               & "Nivel mínimo: " & .MinLevel & "." & vbCrLf _
                                               & "Maná necesario: " & .ManaRequerido & vbCrLf _
                                               & "Maná necesario porcentual: " & .ManaRequeridoPerc & vbCrLf _
                                               & "Energía necesaria: " & .StaRequerido & vbCrLf _
                                               & "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%", FontTypeNames.FONTTYPE_INFO)
            End With
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleSpellInfo de Protocol.bas")
End Sub

''
' Handles the "EquipItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleEquipItem(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim itemSlot As Byte
        
        itemSlot = Reader.ReadInt8()
        
        'Dead users can't equip items
        If (.flags.Muerto = 1) Then Exit Sub
        
        'Validate item slot
        If (itemSlot > .CurrentInventorySlots) Or (itemSlot < 1) Then Exit Sub
        
        If (.Invent.Object(itemSlot).ObjIndex = 0) Then Exit Sub
        
        Call EquiparInvItem(UserIndex, itemSlot)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleEquipItem de Protocol.bas")
End Sub

''
' Handles the "ChangeHeading" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChangeHeading(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 06/28/2008
'Last Modified By: NicoNZ
' 10/01/2008: Tavo - Se cancela la salida del juego si el user esta saliendo
' 06/28/2008: NicoNZ - Sólo se puede cambiar si está inmovilizado.
'***************************************************

On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim heading As eHeading
        Dim posX As Integer
        Dim posY As Integer
                
        heading = Reader.ReadInt8()
        
        If heading > 5 Then
            heading = eHeading.WEST
        End If
        
        If .flags.Paralizado = 1 And .flags.Inmovilizado = 0 Then
            Select Case heading
                Case eHeading.NORTH
                    posY = -1
                Case eHeading.EAST
                    posX = 1
                Case eHeading.SOUTH
                    posY = 1
                Case eHeading.WEST
                    posX = -1
            End Select
            
                If LegalPos(.Pos.Map, .Pos.X + posX, .Pos.Y + posY, CBool(.flags.Navegando), Not CBool(.flags.Navegando)) Then
                    Exit Sub
                End If
        End If
        
        'Validate heading (VB won't say invalid cast if not a valid index like .Net languages would do... *sigh*)
        If heading > 0 And heading < 5 Then
            .Char.heading = heading
            Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleChangeHeading de Protocol.bas")
End Sub

''
' Handles the "ModifySkills" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleModifySkills(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 11/19/09
'11/19/09: Pato - Adapting to new skills system.
'***************************************************

On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        Dim I As Long
        Dim Count As Integer
        Dim Points(1 To NUMSKILLS) As Byte
        
        'Codigo para prevenir el hackeo de los skills
        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        For I = 1 To NUMSKILLS
            Points(I) = Reader.ReadInt8()
            
            If Points(I) < 0 Then
                Call LogHackAttemp(.Name & " IP:" & .IP & " trató de hackear los skills.")
                .Stats.SkillPts = 0
                Call CloseSocket(UserIndex)
                Exit Sub
            End If
            
            Count = Count + Points(I)
        Next I
        
        If Count > .Stats.SkillPts Then
            Call LogHackAttemp(.Name & " IP:" & .IP & " trató de hackear los skills.")
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        
        With .Stats
            For I = 1 To NUMSKILLS
                If Points(I) > 0 Then
                    .SkillPts = .SkillPts - Points(I)
                    Call AddAssignedSkills(UserIndex, I, Points(I))
                    
                    'Client should prevent this, but just in case...
                    If GetSkills(UserIndex, I) > 100 Then
                        .SkillPts = .SkillPts + GetSkills(UserIndex, I) - 100
                        Call ZeroSkills(UserIndex, I)
                        Call AddNaturalSkills(UserIndex, I, 100)
                    End If
                    
                    Call CheckEluSkill(UserIndex, I, True)
                End If
            Next I
        End With
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleModifySkills de Protocol.bas")
End Sub

''
' Handles the "Train" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTrain(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

On Error GoTo ErrHandler
    
    With UserList(UserIndex)

        Dim SpawnedNpc As Integer
        Dim PetIndex As Byte

        
        PetIndex = Reader.ReadInt8()
        
        If .flags.TargetNpc = 0 Then Exit Sub
        
        If Npclist(.flags.TargetNpc).NPCtype <> eNPCType.Entrenador Then Exit Sub
        
        If Npclist(.flags.TargetNpc).Mascotas < MaxMascotasEntrenador Then
            If PetIndex > 0 And PetIndex < Npclist(.flags.TargetNpc).NroCriaturas + 1 Then
                'Create the creature
                SpawnedNpc = SpawnNpc(Npclist(.flags.TargetNpc).Criaturas(PetIndex).NpcIndex, Npclist(.flags.TargetNpc).Pos, True, False)
                
                If SpawnedNpc > 0 Then
                    Npclist(SpawnedNpc).MaestroNpc = .flags.TargetNpc
                    Npclist(.flags.TargetNpc).Mascotas = Npclist(.flags.TargetNpc).Mascotas + 1
                End If
            End If
        Else
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No puedo traer más criaturas, mata las existentes.", Npclist(.flags.TargetNpc).Char.CharIndex, vbWhite))
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleTrain de Protocol.bas")
End Sub

''
' Handles the "CommerceBuy" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceBuy(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        Dim Slot As Byte
        Dim Amount As Integer
        
        Slot = Reader.ReadInt8()
        Amount = Reader.ReadInt16()
        
        'Dead people can't commerce...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        '¿El target es un NPC valido?
        If .flags.TargetNpc < 1 Then Exit Sub
            
        '¿El NPC puede comerciar?
        If Npclist(.flags.TargetNpc).Comercia = 0 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No tengo ningún interés en comerciar.", Npclist(.flags.TargetNpc).Char.CharIndex, vbWhite))
            Exit Sub
        End If
        
        'Only if in commerce mode....
        If Not isTrading(UserIndex) Then
            Call WriteConsoleMsg(UserIndex, "No estás comerciando.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'User compra el item
        Call Comercio(eModoComercio.Compra, UserIndex, .flags.TargetNpc, Slot, Amount)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleCommerceBuy de Protocol.bas")
End Sub

''
' Handles the "BankExtractItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankExtractItem(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim Slot As Byte
        Dim Amount As Integer
        
        Slot = Reader.ReadInt8()
        Amount = Reader.ReadInt16()
        
        'Dead people can't commerce
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        '¿El target es un NPC valido?
        If .flags.TargetNpc < 1 Then Exit Sub
        
        '¿Es el banquero?
        If Npclist(.flags.TargetNpc).NPCtype <> eNPCType.Banquero Then
            Exit Sub
        End If
        
        'User retira el item del slot
        Call UserRetiraItem(UserIndex, Slot, Amount)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleBankExtractItem de Protocol.bas")
End Sub

''
' Handles the "CommerceSell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceSell(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim Slot As Byte
        Dim Amount As Integer
        
        Slot = Reader.ReadInt8()
        Amount = Reader.ReadInt16()
        
        'Dead people can't commerce...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        '¿El target es un NPC valido?
        If .flags.TargetNpc < 1 Then Exit Sub
        
        If UserList(UserIndex).Invent.Object(Slot).ObjIndex <= 0 Then Exit Sub
        
        '¿El NPC puede comerciar?
        If Npclist(.flags.TargetNpc).Comercia = 0 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No tengo ningún interés en comerciar.", Npclist(.flags.TargetNpc).Char.CharIndex, vbWhite))
            Exit Sub
        End If
        
        If ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).ObjType = otQuest Then
            Call WriteConsoleMsg(UserIndex, "No puedes vender este tipo de objeto.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'User compra el item del slot
        Call Comercio(eModoComercio.Venta, UserIndex, .flags.TargetNpc, Slot, Amount)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleCommerceSell de Protocol.bas")
End Sub

''
' Handles the "BankDeposit" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankDeposit(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim Slot As Byte
        Dim Amount As Integer
        
        Slot = Reader.ReadInt8()
        Amount = Reader.ReadInt16()
        
        'Dead people can't commerce...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        '¿El target es un NPC valido?
        If .flags.TargetNpc < 1 Then Exit Sub
        
        '¿El NPC puede comerciar?
        If Npclist(.flags.TargetNpc).NPCtype <> eNPCType.Banquero Then
            Exit Sub
        End If
        
        'User deposita el item del slot rdata
        Call UserDepositaItem(UserIndex, Slot, Amount)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleBankDeposit de Protocol.bas")
End Sub

''
' Handles the "ForumPost" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForumPost(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 02/01/2010
'02/01/2010: ZaMa - Implemento nuevo sistema de foros
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim ForumMsgType As eForumMsgType
        
        Dim Title As String
        Dim Post As String
        Dim ForumIndex As Integer
        Dim ForumType As Byte
                
        ForumMsgType = Reader.ReadInt8()
        
        Title = Reader.ReadString8()
        Post = Reader.ReadString8()

        If .flags.TargetObj > 0 Then
            ForumType = ForumAlignment(ForumMsgType)
            
            Select Case ForumType
            
                Case eForumType.ieGeneral
                    ForumIndex = GetForumIndex(ObjData(.flags.TargetObj).ForoID)
                    
                Case eForumType.ieREAL
                    ForumIndex = GetForumIndex(FORO_REAL_ID)
                    
                Case eForumType.ieCAOS
                    ForumIndex = GetForumIndex(FORO_CAOS_ID)
                    
            End Select
            
            Call AddPost(ForumIndex, Post, .Name, Title, EsAnuncio(ForumMsgType))
        End If
        
    End With
  
  Exit Sub
  
ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleForumPost de Protocol.bas")
End Sub

''
' Handles the "MoveSpell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMoveSpell(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

On Error GoTo ErrHandler
  
    With Reader
        Dim dir As Integer
        
        If .ReadBool() Then
            dir = 1
        Else
            dir = -1
        End If
        
        Call DesplazarHechizo(UserIndex, dir, .ReadInt8())
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleMoveSpell de Protocol.bas")
End Sub

''
' Handles the "MoveBank" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMoveBank(ByVal UserIndex As Integer)
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 06/14/09
'
'***************************************************

On Error GoTo ErrHandler

    With Reader

        Dim dir As Integer
        Dim Slot As Byte
        Dim TempItem As Obj
        
        If .ReadBool() Then
            dir = 1
        Else
            dir = -1
        End If
        
        Slot = .ReadInt8()
    End With
        
    With UserList(UserIndex)
        TempItem.ObjIndex = .BancoInvent.Object(Slot).ObjIndex
        TempItem.Amount = .BancoInvent.Object(Slot).Amount
        
        If dir = 1 Then 'Mover arriba
            .BancoInvent.Object(Slot) = .BancoInvent.Object(Slot - 1)
            .BancoInvent.Object(Slot - 1).ObjIndex = TempItem.ObjIndex
            .BancoInvent.Object(Slot - 1).Amount = TempItem.Amount
        Else 'mover abajo
            .BancoInvent.Object(Slot) = .BancoInvent.Object(Slot + 1)
            .BancoInvent.Object(Slot + 1).ObjIndex = TempItem.ObjIndex
            .BancoInvent.Object(Slot + 1).Amount = TempItem.Amount
        End If
    End With
    
    Call UpdateBanUserInv(False, UserIndex, Slot)
    Call UpdateBanUserInv(False, UserIndex, Slot + dir)
    
    'Call UpdateVentanaBanco(UserIndex)

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleMoveBank de Protocol.bas")
End Sub

''
' Handles the "HandleUserCommerceOfferGold" message.
'
' @param    userIndex The index of the user sending the message.
Private Sub HandleUserCommerceOfferGold(ByVal UserIndex As Integer)
    Dim Amount As Long
    Dim tUser As Integer
    
    Amount = Reader.ReadInt32()
    
    'Get the other player
    tUser = getTradingUser(UserIndex)
    
    With UserList(UserIndex)
    
        ' We shouldn't be receiving an offer after the user has confirmed it's previous one
        If .ComUsu.Confirmo = True Then
            ' Finish the trade
            Call FinComerciarUsu(UserIndex)
            Call FinComerciarUsu(tUser)
            Exit Sub
        End If
        
        
        ' Can't offer more than he has
        If Amount > (.Stats.GLD - .ComUsu.GoldAmount) Then
            Call WriteCommerceChat(UserIndex, "No tienes esa cantidad de oro para agregar a la oferta.", FontTypeNames.FONTTYPE_TALK)
            Exit Sub
        End If
    
        If Amount < 0 Then
            If Abs(Amount) > .ComUsu.GoldAmount Then
                Amount = .ComUsu.GoldAmount * (-1)
            End If
        End If
    
    End With
    Call AgregarOferta(UserIndex, 0, 0, Amount, True)

    Call EnviarOfertaOro(tUser)
    
End Sub

''
' Handles the "UserCommerceOffer" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceOffer(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 24/11/2009
'24/11/2009: ZaMa - Nuevo sistema de comercio
'***************************************************

 On Error GoTo ErrHandler
    Dim errorTracker As Integer

    With UserList(UserIndex)
        Dim Amount As Long
        Dim Slot As Byte
        Dim tUser As Integer
        Dim OfferSlot As Byte
        Dim ObjIndex As Integer
        
        Slot = Reader.ReadInt8()
        Amount = Reader.ReadInt32()
        OfferSlot = Reader.ReadInt8()

        If Not PuedeSeguirComerciando(UserIndex) Then Exit Sub

        'Get the other player
        tUser = getTradingUser(UserIndex)

        ' If he's already confirmed his offer, but now tries to change it, then he's cheating
        If UserList(UserIndex).ComUsu.Confirmo = True Then
            ' Finish the trade
            Call FinComerciarUsu(UserIndex)

            Call FinComerciarUsu(tUser)

            Exit Sub
        End If

        'If slot is invalid and it's not gold or it's not 0 (Substracting), then ignore it.
        If ((Slot < 0 Or Slot > UserList(UserIndex).CurrentInventorySlots) And Slot <> FLAGORO) Then Exit Sub

        'If OfferSlot is invalid, then ignore it.
        If OfferSlot < 1 Or OfferSlot > MAX_OFFER_SLOTS + 1 Then Exit Sub

        ' Can be negative if substracted from the offer, but never 0.
        If Amount = 0 Then Exit Sub

            'If modifing a filled offerSlot, we already got the objIndex, then we don't need to know it
            If Slot <> 0 Then ObjIndex = .Invent.Object(Slot).ObjIndex

            ' Non-Transferible or commerciable?
            If ObjIndex <> 0 Then
                If (IsSecondaryArmour(ObjIndex) Or ObjData(ObjIndex).Intransferible = 1 Or _
                    ObjData(ObjIndex).NoComerciable = 1) Then

                    Call WriteCommerceChat(UserIndex, "No puedes comerciar este ítem.", FontTypeNames.FONTTYPE_TALK)

                    Exit Sub
                End If
            End If
            
            'Don't allow to sell boats if they are equipped (you can't take them off in the water and causes trouble)
            If .flags.Navegando = 1 And .Invent.BarcoSlot = Slot Then
                Call WriteCommerceChat(UserIndex, "No puedes vender tu barco mientras lo estés usando.", FontTypeNames.FONTTYPE_TALK)
                Exit Sub
            End If

            If .Invent.MochilaEqpSlot > 0 And .Invent.MochilaEqpSlot = Slot Then
                Call WriteCommerceChat(UserIndex, "No puedes vender tu mochila mientras la estés usando.", FontTypeNames.FONTTYPE_TALK)
                Exit Sub
            End If
            
            ' Can't offer more than he has
            If Not HasEnoughItems(UserIndex, ObjIndex, _
                TotalOfferItems(ObjIndex, UserIndex) + Amount) Then

                Call WriteCommerceChat(UserIndex, "No tienes esa cantidad.", FontTypeNames.FONTTYPE_TALK)

                Exit Sub
            End If

            If Amount < 0 Then
                If Abs(Amount) > .ComUsu.cant(OfferSlot) Then
                    Amount = .ComUsu.cant(OfferSlot) * (-1)
                End If
            End If
            
            If ObjIndex > 0 Then
                If ObjData(ObjIndex).ObjType = otQuest Then
                    Call WriteCancelOfferItem(UserIndex, OfferSlot)
                    Exit Sub
                End If
                 
                If ItemNewbie(ObjIndex) Then
                    Call WriteCancelOfferItem(UserIndex, OfferSlot)
                    Exit Sub
                End If
    
                Call WriteCommerceChat(UserIndex, "¡Agregaste " & Amount & " " & ObjData(ObjIndex).Name & " a tu oferta!!", FontTypeNames.FONTTYPE_GUILD)
                
            End If
           

        Call AgregarOferta(UserIndex, OfferSlot, ObjIndex, Amount, Slot = FLAGORO)

        Call EnviarOferta(tUser, OfferSlot)

    End With
    
    Exit Sub

ErrHandler:
    Call LogError("Error en HandleUserCommerceOffer. Error: " & Err.Description & ". User: " & UserList(UserIndex).Name & "(" & UserIndex & ")" & _
        ". tUser: " & tUser & ". Slot: " & Slot & ". Amount: " & Amount & ". OfferSlot: " & OfferSlot)
End Sub

''
' Handles the "Online" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnline(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
On Error GoTo ErrHandler
  
    Dim Count As Long
    
    With UserList(UserIndex)
        Count = GetUsersCount()
        
        Call WriteConsoleMsg(UserIndex, "Número de usuarios: " & CStr(Count), FontTypeNames.FONTTYPE_INFO)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleOnline de Protocol.bas")
End Sub

''
' Handles the "Quit" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleQuit(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/15/2008 (NicoNZ)
'If user is invisible, it automatically becomes
'visible before doing the countdown to exit
'04/15/2008 - No se reseteaban lso contadores de invi ni de ocultar. (NicoNZ)
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)
    
        If .flags.Paralizado Then
            Call MessageManager.Prepare(UserIndex, eMessageId.Cant_Quit_Paralized).SendToConsole
            Exit Sub
        End If
       
        If Not EsGm(UserIndex) And .flags.invisible Then
            Call WriteConsoleMsg(UserIndex, "No puedes salir estando invisible.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Not EsGm(UserIndex) And .flags.Oculto Then
            Call WriteConsoleMsg(UserIndex, "No puedes salir estando oculto.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Call ExitSecureCommerce(UserIndex)
        Call Cerrar_Usuario(UserIndex, True)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleQuit de Protocol.bas")
End Sub

''
' Handles the "RequestAccountState" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestAccountState(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
On Error GoTo ErrHandler
  
    Dim earnings As Integer
    Dim Percentage As Integer
    
    With UserList(UserIndex)
        'Dead people can't check their accounts
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNpc = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Distancia(Npclist(.flags.TargetNpc).Pos, .Pos) > 3 Then
            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Select Case Npclist(.flags.TargetNpc).NPCtype
            Case eNPCType.Banquero
                Call WriteChatOverHead(UserIndex, "Tienes " & .Stats.Banco & " monedas de oro en tu cuenta.", Npclist(.flags.TargetNpc).Char.CharIndex, vbWhite)
            
            Case eNPCType.Timbero
                If Not .flags.Privilegios And PlayerType.User Then
                    earnings = Apuestas.Ganancias - Apuestas.Perdidas
                    
                    If earnings >= 0 And Apuestas.Ganancias <> 0 Then
                        Percentage = Int(earnings * 100 / Apuestas.Ganancias)
                    End If
                    
                    If earnings < 0 And Apuestas.Perdidas <> 0 Then
                        Percentage = Int(earnings * 100 / Apuestas.Perdidas)
                    End If
                    
                    Call WriteConsoleMsg(UserIndex, "Entradas: " & Apuestas.Ganancias & " Salida: " & Apuestas.Perdidas & " Ganancia Neta: " & earnings & " (" & Percentage & "%) Jugadas: " & Apuestas.Jugadas, FontTypeNames.FONTTYPE_INFO)
                End If
        End Select
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleRequestAccountState de Protocol.bas")
End Sub

''
' Handles the "PetStand" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePetStand(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)
        'Dead people can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNpc = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Dim NpcIndex As Integer
        NpcIndex = .flags.TargetNpc
        
        'Make sure it's close enough
        If Distancia(Npclist(NpcIndex).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make sure it's his pet
        If Npclist(NpcIndex).MaestroUser <> UserIndex Then Exit Sub
        
        'Do it!
        Npclist(NpcIndex).Movement = TipoAI.ESTATICO
        Npclist(NpcIndex).MenuIndex = eMenues.ieMascotaQuieta
        
        Call Expresar(NpcIndex, UserIndex)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandlePetStand de Protocol.bas")
End Sub

''
' Handles the "PetFollow" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePetFollow(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNpc = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Dim NpcIndex As Integer
        NpcIndex = .flags.TargetNpc
        
        'Make sure it's close enough
        If Distancia(Npclist(NpcIndex).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make usre it's the user's pet
        If Npclist(NpcIndex).MaestroUser <> UserIndex Then Exit Sub
        
        'Do it
        Call FollowAmo(NpcIndex)
        Npclist(NpcIndex).MenuIndex = eMenues.ieMascota
        
        Call Expresar(NpcIndex, UserIndex)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandlePetFollow de Protocol.bas")
End Sub


''
' Handles the "ReleasePet" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleReleasePet(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 18/11/2009
'
'***************************************************

On Error GoTo ErrHandler
    Dim fromForm As Boolean
    Dim petIndexFromForm As Byte
    
    With UserList(UserIndex)
        fromForm = Reader.ReadBool
        petIndexFromForm = Reader.ReadInt8
        
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
                
        ' If the command was generated by using the Pets form, then:
        If fromForm Then
            
            If UserList(UserIndex).TammedPetsCount = 0 Then
                Call WriteConsoleMsg(UserIndex, "No tienes ninguna mascota domada.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            If petIndexFromForm < 1 Or petIndexFromForm > Classes(.clase).ClassMods.MaxTammedPets Then
                Call WriteConsoleMsg(UserIndex, "Debes seleccionar un slot válido.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            Call QuitarPet(UserIndex, petIndexFromForm)
            
            Exit Sub
        End If
        
        If UserList(UserIndex).flags.TargetNpc = 0 Then
            Call WriteConsoleMsg(UserIndex, "Debes clickear un NPC antes de usar este comando.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        Dim invokedPet As Integer
        invokedPet = GetInvokedPetIndexByNpcIndex(UserIndex, UserList(UserIndex).flags.TargetNpc)
                   
        ' if its an invoked pet.
        If invokedPet >= 1 And invokedPet <= Classes(.clase).ClassMods.MaxInvokedPets Then
            If UserList(UserIndex).InvokedPetsCount = 0 Then
                Call WriteConsoleMsg(UserIndex, "No tienes ninguna invocación.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            Call QuitarInvocacion(UserIndex, invokedPet)
            Call WriteConsoleMsg(UserIndex, "Has liberado a tu invocación.", FontTypeNames.FONTTYPE_INFOBOLD, eMessageType.info)

            ' Exit sub, because the NPC could be or an invocation or a tammed pet, not both.
            Exit Sub
        End If
                
        Dim tammedPet As Integer
        tammedPet = GetTammedPetIndexByNpcIndex(UserIndex, .flags.TargetNpc)
                
        ' if its a tammed pet
        If tammedPet >= 1 And tammedPet <= Classes(.clase).ClassMods.MaxTammedPets Then
            If .TammedPetsCount = 0 Then
                Call WriteConsoleMsg(UserIndex, "No tienes ninguna mascota domada.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            Call QuitarPet(UserIndex, tammedPet)
            
            Exit Sub
        End If
                
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleReleasePet de Protocol.bas")
End Sub

''
' Handles the "TrainList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTrainList(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNpc = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make sure it's close enough
        If Distancia(Npclist(.flags.TargetNpc).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make sure it's the trainer
        If Npclist(.flags.TargetNpc).NPCtype <> eNPCType.Entrenador Then Exit Sub
        
        Call WriteTrainerCreatureList(UserIndex, .flags.TargetNpc)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleTrainList de Protocol.bas")
End Sub

''
' Handles the "Rest" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRest(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)
        Call UserRest(UserIndex)
        
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleRest de Protocol.bas")
End Sub

''
' Handles the "Meditate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMeditate(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/15/08 (NicoNZ)
'Arreglé un bug que mandaba un index de la meditacion diferente
'al que decia el server.
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes meditar cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Can he meditate?
        If .Stats.MaxMan = 0 Then
             Call WriteConsoleMsg(UserIndex, "Sólo las clases mágicas conocen el arte de la meditación.", FontTypeNames.FONTTYPE_INFO)
             Exit Sub
        End If
        
        'Admins don't have to wait :D
        If Not .flags.Privilegios And PlayerType.User Then
            .Stats.MinMAN = .Stats.MaxMan
            Call WriteConsoleMsg(UserIndex, "Maná restaurado.", FontTypeNames.FONTTYPE_VENENO)
            Call WriteUpdateMana(UserIndex)
            Exit Sub
        End If
        
        Call WriteMeditateToggle(UserIndex)
        
        If .flags.Meditando Then _
           Call WriteConsoleMsg(UserIndex, "Dejas de meditar.", FontTypeNames.FONTTYPE_INFO, eMessageType.info)
        
        .flags.Meditando = Not .flags.Meditando
        
        'Barrin 3/10/03 Tiempo de inicio al meditar
        If .flags.Meditando Then
            .Counters.tInicioMeditar = GetTickCount()
            
            Call WriteConsoleMsg(UserIndex, "Te estás concentrando. En breve comenzarás a meditar.", FontTypeNames.FONTTYPE_INFO)
            .Char.Loops = INFINITE_LOOPS
            
            'Show proper FX according to level if not dark zone
            If Not MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger = eTrigger.zonaOscura Then
                .Char.FX = ConstantesMeditations(.Faccion.Alignment, .Stats.ELV)
            End If
            
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, .Char.FX, INFINITE_LOOPS, , 0))
        Else
            .Counters.bPuedeMeditar = False
            
            .Char.FX = 0
            .Char.Loops = 0
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0, , 0))
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleMeditate de Protocol.bas")
End Sub

''
' Handles the "Resucitate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleResucitate(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'Se asegura que el target es un npc
        If .flags.TargetNpc = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate NPC and make sure player is dead
        If (Npclist(.flags.TargetNpc).NPCtype <> eNPCType.Revividor _
            And (Npclist(.flags.TargetNpc).NPCtype <> eNPCType.ResucitadorNewbie Or Not EsNewbie(UserIndex))) _
            Or .flags.Muerto = 0 Then Exit Sub
        
        'Make sure it's close enough
        If Distancia(.Pos, Npclist(.flags.TargetNpc).Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "El sacerdote no puede resucitarte debido a que estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Call RevivirUsuario(UserIndex, False)
        
        .Stats.MinHp = .Stats.MaxHp
        .Stats.MinMAN = .Stats.MaxMan
        .Stats.MinSta = .Stats.MaxSta
        
        Call WriteUpdateUserStats(UserIndex)
        Call WriteConsoleMsg(UserIndex, "¡¡Has sido resucitado!!", FontTypeNames.FONTTYPE_INFO)
    End With
    
    Exit Sub

ErrHandler:

    Call LogError("Error en HandleResucitate. Error: " & Err.Number & " - " & _
        Err.Description & ". Usuario: " & UserList(UserIndex).Name & "(" & UserIndex & ")")
End Sub

''
' Handles the "Consultation" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleConsultation(ByVal UserIndex As String)
'***************************************************
'Author: ZaMa
'Last Modification: 01/05/2010
'Habilita/Deshabilita el modo consulta.
'01/05/2010: ZaMa - Agrego validaciones.
'16/09/2010: ZaMa - No se hace visible en los clientes si estaba navegando (porque ya lo estaba).
'***************************************************
On Error GoTo ErrHandler
  
    
    Dim UserConsulta As Integer
    
    With UserList(UserIndex)
        ' Comando exclusivo para gms
        If Not EsGm(UserIndex) Then Exit Sub
        
        UserConsulta = .flags.TargetUser
        
        'Se asegura que el target es un usuario
        If UserConsulta = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un usuario, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        ' No podes ponerte a vos mismo en modo consulta.
        If UserConsulta = UserIndex Then Exit Sub
        
        ' No podes estra en consulta con otro gm
        If EsGm(UserConsulta) Then
            Call WriteConsoleMsg(UserIndex, "No puedes iniciar el modo consulta con otro administrador.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        ' If the current user is GM, then check if there's another player in HelpMode
        If EsGm(UserIndex) Then
            If .flags.HelpingUser > 0 And .flags.HelpingUser <> UserConsulta Then
                ' If there's another player in HelpMode, then finish it.
                Call WriteConsoleMsg(UserIndex, "Has terminado el modo consulta con " & UserList(.flags.HelpingUser).Name & ".", FontTypeNames.FONTTYPE_INFOBOLD)
                Call WriteConsoleMsg(.flags.HelpingUser, "Has terminado el modo consulta.", FontTypeNames.FONTTYPE_INFOBOLD)
                Call SetHelpModeToUser(UserIndex, .flags.HelpingUser, False)
            End If
        End If
        
        Dim UserName As String
        UserName = UserList(UserConsulta).Name
        
        ' Si ya estaba en consulta, termina la consulta
        If UserList(UserConsulta).flags.HelpMode Then
            
            Call SetHelpModeToUser(UserIndex, UserConsulta, False)
            
            Call WriteConsoleMsg(UserIndex, "Has terminado el modo consulta con " & UserName & ".", FontTypeNames.FONTTYPE_INFOBOLD)
            Call WriteConsoleMsg(UserConsulta, "Has terminado el modo consulta.", FontTypeNames.FONTTYPE_INFOBOLD)
            
            Call LogGM(.Name, "Termino consulta con " & UserName)
        
        ' Sino la inicia
        Else
            Call SetHelpModeToUser(UserIndex, UserConsulta, True)
            
            Call WriteConsoleMsg(UserIndex, "Has iniciado el modo consulta con " & UserName & ".", FontTypeNames.FONTTYPE_INFOBOLD)
            Call WriteConsoleMsg(UserConsulta, "Has iniciado el modo consulta.", FontTypeNames.FONTTYPE_INFOBOLD)
            
            Call LogGM(.Name, "Inicio consulta con " & UserName)
            
            With UserList(UserConsulta)

                
                ' Pierde invi u ocu
                If .flags.invisible = 1 Or .flags.Oculto = 1 Then
                    .flags.Oculto = 0
                    .flags.invisible = 0
                    .Counters.TiempoOculto = 0
                    .Counters.Invisibilidad = 0
                    
                    If UserList(UserConsulta).flags.Navegando = 0 Then
                        Call UsUaRiOs.SetInvisible(UserConsulta, UserList(UserConsulta).Char.CharIndex, False)
                    End If
                End If
            End With
        End If
        
    End With

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleConsultation de Protocol.bas")
End Sub

''
' Handles the "Heal" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleHeal(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)
        'Se asegura que el target es un npc
        If .flags.TargetNpc = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If (Npclist(.flags.TargetNpc).NPCtype <> eNPCType.Revividor _
            And Npclist(.flags.TargetNpc).NPCtype <> eNPCType.ResucitadorNewbie) _
            Or .flags.Muerto <> 0 Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNpc).Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "El sacerdote no puede curarte debido a que estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        .Stats.MinHp = .Stats.MaxHp
        .Stats.MinMAN = .Stats.MaxMan
        .Stats.MinSta = .Stats.MaxSta
        
        'Call WriteUpdateHP(UserIndex)
        Call WriteUpdateUserStats(UserIndex)
        
        Call WriteConsoleMsg(UserIndex, "¡¡Has sido curado!!", FontTypeNames.FONTTYPE_INFO)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleHeal de Protocol.bas")
End Sub

''
' Handles the "RequestStats" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestStats(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

On Error GoTo ErrHandler

    Call SendUserStatsTxt(UserIndex, UserIndex)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleRequestStats de Protocol.bas")
End Sub

''
' Handles the "Help" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleHelp(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

On Error GoTo ErrHandler

    Call SendHelp(UserIndex)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleHelp de Protocol.bas")
End Sub

''
' Handles the "CommerceStart" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceStart(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 28/10/2014
'28/10/2014: D'Artagnan - Can't commerce while sailing.
'***************************************************
On Error GoTo ErrHandler
  
    Dim I As Integer
    With UserList(UserIndex)

        If Not CommerceAllowed(UserIndex) Then Exit Sub
        
        'Validate target NPC
        If .flags.TargetNpc > 0 Then
            'Does the NPC want to trade??
            If Npclist(.flags.TargetNpc).Comercia = 0 Then
                If LenB(Npclist(.flags.TargetNpc).Desc) <> 0 Then
                    Call WriteChatOverHead(UserIndex, "No tengo ningún interés en comerciar.", Npclist(.flags.TargetNpc).Char.CharIndex, vbWhite)
                End If
                
                Exit Sub
            End If
            
            If Distancia(Npclist(.flags.TargetNpc).Pos, .Pos) > 3 Then
                Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'Start commerce....
            Call IniciarComercioNPC(UserIndex)
        '[Alejo]
        ElseIf .flags.TargetUser > 0 Then
            'User commerce...
            'Can he commerce??
            Dim Priv As PlayerType
            Dim Valid As Boolean
            ' Privilegios
            Priv = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
            
            If .flags.Privilegios And Priv Then
                Valid = (UserList(.flags.TargetUser).flags.Privilegios And Priv)
            Else
                Valid = (UserList(.flags.TargetUser).flags.Privilegios And Priv) = 0
            End If
            
            If Not Valid Then
                Call WriteConsoleMsg(UserIndex, "No puedes comerciar con alguien de esta jerarquía.", FontTypeNames.FONTTYPE_WARNING)
                Exit Sub
            End If
            
            'Is the other one dead??
            If UserList(.flags.TargetUser).flags.Muerto = 1 Then
                Call WriteConsoleMsg(UserIndex, "¡¡No puedes comerciar con los muertos!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'Is it me??
            If .flags.TargetUser = UserIndex Then
                Call WriteConsoleMsg(UserIndex, "¡¡No puedes comerciar con vos mismo!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'Check distance
            If Distancia(UserList(.flags.TargetUser).Pos, .Pos) > 3 Then
                Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos del usuario.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'Is he already trading?? is it with me or someone else??
            If isTrading(.flags.TargetUser) And (getTradingUser(.flags.TargetUser) <> UserIndex) Then
                Call WriteConsoleMsg(UserIndex, "No puedes comerciar con el usuario en este momento.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'Initialize some variables...
            .flags.Comerciando = (Not .flags.TargetUser)
            .ComUsu.DestNick = UserList(.flags.TargetUser).Name
            
            For I = 1 To MAX_OFFER_SLOTS
                .ComUsu.cant(I) = 0
                .ComUsu.Objeto(I) = 0
            Next I
            
            .ComUsu.GoldAmount = 0
            
            .ComUsu.Acepto = False
            .ComUsu.Confirmo = False
            
            'Rutina para comerciar con otro usuario
            Call IniciarComercioConUsuario(UserIndex, .flags.TargetUser)
        Else
            Call WriteConsoleMsg(UserIndex, "Primero haz click izquierdo sobre el personaje.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleCommerceStart de Protocol.bas")
End Sub

''
' Handles the "BankStart" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankStart(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 28/10/2014
'28/10/2014: D'Artagnan - Can't commerce while sailing.
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        If Not CommerceAllowed(UserIndex) Then Exit Sub
        
        'Validate target NPC
        If .flags.TargetNpc > 0 Then
            If Distancia(Npclist(.flags.TargetNpc).Pos, .Pos) > 3 Then
                Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'If it's the banker....
            If Npclist(.flags.TargetNpc).NPCtype = eNPCType.Banquero Then
                Call IniciarDeposito(UserIndex)
            End If
        Else
            Call WriteConsoleMsg(UserIndex, "Primero haz click izquierdo sobre el personaje.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleBankStart de Protocol.bas")
End Sub

''
' Handles the "Enlist" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleEnlist(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        'Validate target NPC
        If .flags.TargetNpc = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Npclist(.flags.TargetNpc).NPCtype <> eNPCType.Noble _
            Or .flags.Muerto <> 0 Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNpc).Pos) > 4 Then
            Call WriteConsoleMsg(UserIndex, "Debes acercarte más.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Npclist(.flags.TargetNpc).flags.Faccion = 0 Then
            Call EnlistarArmadaReal(UserIndex)
        Else
            Call EnlistarCaos(UserIndex)
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleEnlist de Protocol.bas")
End Sub

''
' Handles the "Information" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInformation(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
On Error GoTo ErrHandler
  
    Dim Matados As Integer
    Dim NextRecom As Integer
    Dim Diferencia As Integer
    
    With UserList(UserIndex)

        'Validate target NPC
        If .flags.TargetNpc = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Npclist(.flags.TargetNpc).NPCtype <> eNPCType.Noble _
                Or .flags.Muerto <> 0 Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNpc).Pos) > 4 Then
            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        
        NextRecom = .Faccion.NextRecompensa
        
        If Npclist(.flags.TargetNpc).flags.Faccion = 0 Then
            If .Faccion.ArmadaReal = 0 Then
                Call WriteChatOverHead(UserIndex, "¡¡No perteneces a las tropas reales!!", Npclist(.flags.TargetNpc).Char.CharIndex, vbWhite)
                Exit Sub
            End If
            
            Matados = .Faccion.CriminalesMatados
            Diferencia = NextRecom - Matados
            
            If Diferencia > 0 Then
                Call WriteChatOverHead(UserIndex, "Tu deber es combatir criminales, mata " & Diferencia & " criminales más y te daré una recompensa.", Npclist(.flags.TargetNpc).Char.CharIndex, vbWhite)
            Else
                Call WriteChatOverHead(UserIndex, "Tu deber es combatir criminales, y ya has matado los suficientes como para merecerte una recompensa.", Npclist(.flags.TargetNpc).Char.CharIndex, vbWhite)
            End If
        Else
            If .Faccion.FuerzasCaos = 0 Then
                Call WriteChatOverHead(UserIndex, "¡¡No perteneces a la legión oscura!!", Npclist(.flags.TargetNpc).Char.CharIndex, vbWhite)
                Exit Sub
            End If
            
            Matados = .Faccion.CiudadanosMatados
            Diferencia = NextRecom - Matados
            
            If Diferencia > 0 Then
                Call WriteChatOverHead(UserIndex, "Tu deber es sembrar el caos y la desesperanza, mata " & Diferencia & " ciudadanos más y te daré una recompensa.", Npclist(.flags.TargetNpc).Char.CharIndex, vbWhite)
            Else
                Call WriteChatOverHead(UserIndex, "Tu deber es sembrar el caos y la desesperanza, y creo que estás en condiciones de merecer una recompensa.", Npclist(.flags.TargetNpc).Char.CharIndex, vbWhite)
            End If
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleInformation de Protocol.bas")
End Sub

''
' Handles the "Reward" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleReward(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        'Validate target NPC
        If .flags.TargetNpc = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Npclist(.flags.TargetNpc).NPCtype <> eNPCType.Noble _
            Or .flags.Muerto <> 0 Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNpc).Pos) > 4 Then
            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Npclist(.flags.TargetNpc).flags.Faccion = 0 Then
             If .Faccion.ArmadaReal = 0 Then
                 Call WriteChatOverHead(UserIndex, "¡¡No perteneces a las tropas reales!!", Npclist(.flags.TargetNpc).Char.CharIndex, vbWhite)
                 Exit Sub
             End If
             Call RecompensaArmadaReal(UserIndex)
        Else
             If .Faccion.FuerzasCaos = 0 Then
                 Call WriteChatOverHead(UserIndex, "¡¡No perteneces a la legión oscura!!", Npclist(.flags.TargetNpc).Char.CharIndex, vbWhite)
                 Exit Sub
             End If
             Call RecompensaCaos(UserIndex)
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleReward de Protocol.bas")
End Sub

''
' Handles the "UpTime" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUpTime(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/10/08
'01/10/2008 - Marcos Martinez (ByVal) - Automatic restart removed from the server along with all their assignments and varibles
'***************************************************

On Error GoTo ErrHandler

    Dim Time As Long
    Dim UpTimeStr As String
    
    'Get total time in seconds
    Time = GetInterval(GetTickCount(), tInicioServer) \ 1000
    
    'Get times in dd:hh:mm:ss format
    UpTimeStr = (Time Mod 60) & " segundos."
    Time = Time \ 60
    
    UpTimeStr = (Time Mod 60) & " minutos, " & UpTimeStr
    Time = Time \ 60
    
    UpTimeStr = (Time Mod 24) & " horas, " & UpTimeStr
    Time = Time \ 24
    
    If Time = 1 Then
        UpTimeStr = Time & " día, " & UpTimeStr
    Else
        UpTimeStr = Time & " días, " & UpTimeStr
    End If
    
    Call WriteConsoleMsg(UserIndex, "Server Online: " & UpTimeStr, FontTypeNames.FONTTYPE_INFO)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleUpTime de Protocol.bas")
End Sub

''
' Handles the "PartyLeave" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyLeave(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

On Error GoTo ErrHandler

    Call mdParty.SalirDeParty(UserIndex)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandlePartyLeave de Protocol.bas")
End Sub

''
' Handles the "PartyCreate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyCreate(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

On Error GoTo ErrHandler

    If Not mdParty.PuedeCrearParty(UserIndex) Then Exit Sub
    
    Call mdParty.CrearParty(UserIndex)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandlePartyCreate de Protocol.bas")
End Sub

''
' Handles the "Inquiry" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInquiry(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

On Error GoTo ErrHandler

    ConsultaPopular.SendInfoEncuesta (UserIndex)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleInquiry de Protocol.bas")
End Sub

''
' Handles the "GuildMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildMessage(ByVal UserIndex As Integer)

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim Chat As String
        
        Chat = Reader.ReadString8()
        
        If LenB(Chat) = 0 Then Exit Sub
        
        If .Guild.GuildIndex <= 0 Then Exit Sub

        'Analize chat...
        Call Statistics.ParseChat(Chat)
            
        Call SendData(SendTarget.ToDiosesYclan, .Guild.GuildIndex, PrepareMessageGuildChat(.Name & "> " & Chat))
        
        If Not (.flags.AdminInvisible = 1) And GuildList(.Guild.GuildIndex).UpgradeEffect.IsChatOverHead Then _
            Call SendData(SendTarget.ToClanArea, UserIndex, PrepareMessageChatPersonalizado("< " & Chat & " >", .Char.CharIndex, 2))

    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleGuildMessage de Protocol.bas")
End Sub


''
' Handles the "PartyMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyMessage(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim Chat As String
        
        Chat = Reader.ReadString8()

        If LenB(Chat) <> 0 Then
            'Analize chat...
            Call Statistics.ParseChat(Chat)
            
            Call mdParty.BroadCastParty(UserIndex, Chat)
'TODO : Con la 0.12.1 se debe definir si esto vuelve o se borra (/CMSG overhead)
            'Call SendData(SendTarget.ToPartyArea, UserIndex, UserList(UserIndex).Pos.map, "||" & vbYellow & "ï¿½< " & mid$(rData, 7) & " >ï¿½" & CStr(UserList(UserIndex).Char.CharIndex))
        End If
    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandlePartyMessage de Protocol.bas")
End Sub

''
' Handles the "PartyOnline" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyOnline(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

On Error GoTo ErrHandler

    Call mdParty.OnlineParty(UserIndex)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandlePartyOnline de Protocol.bas")
End Sub

''
' Handles the "CouncilMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCouncilMessage(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim Chat As String
        
        Chat = Reader.ReadString8()

        If LenB(Chat) <> 0 Then
            'Analize chat...
            Call Statistics.ParseChat(Chat)
            
            If .flags.Privilegios And PlayerType.RoyalCouncil Then
                Call SendData(SendTarget.ToConsejo, UserIndex, PrepareMessageConsoleMsg("(Consejero) " & .Name & "> " & Chat, FontTypeNames.FONTTYPE_CONSEJO))
            ElseIf .flags.Privilegios And PlayerType.ChaosCouncil Then
                Call SendData(SendTarget.ToConsejoCaos, UserIndex, PrepareMessageConsoleMsg("(Consejero) " & .Name & "> " & Chat, FontTypeNames.FONTTYPE_CONSEJOCAOS))
            End If
        End If
        

    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleCouncilMessage de Protocol.bas")
End Sub

''
' Handles the "RoleMasterRequest" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRoleMasterRequest(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim request As String
        
        request = Reader.ReadString8()

        If LenB(request) <> 0 Then
            Call WriteConsoleMsg(UserIndex, "Su solicitud ha sido enviada.", FontTypeNames.FONTTYPE_INFO)
            Call SendData(SendTarget.ToRMsAndHigherAdmins, 0, PrepareMessageConsoleMsg(.Name & " PREGUNTA ROL: " & request, FontTypeNames.FONTTYPE_GUILDMSG))
        End If
    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleRoleMasterRequest de Protocol.bas")
End Sub

''
' Handles the "GMRequest" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGMRequest(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        If Not Ayuda.Existe(.Name) Then
            Call WriteConsoleMsg(UserIndex, "El mensaje ha sido entregado, ahora sólo debes esperar que se desocupe algún GM.", FontTypeNames.FONTTYPE_INFO)
            Call SendData(SendTarget.ToAdminsButRMs, 0, PrepareMessageConsoleMsg("Servidor> " & .Name & " ha solicitado consulta.", FontTypeNames.FONTTYPE_SERVER))
            Call Ayuda.Push(.Name)
        Else
            Call Ayuda.Quitar(.Name)
            Call Ayuda.Push(.Name)
            Call WriteConsoleMsg(UserIndex, "Ya habías mandado un mensaje, tu mensaje ha sido movido al final de la cola de mensajes.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleGMRequest de Protocol.bas")
End Sub

''
' Handles the "BugReport" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBugReport(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim N As Integer

        Dim bugReport As String
        
        bugReport = Reader.ReadString8()
        
        N = FreeFile
        Open ServerConfiguration.LogsPaths.GeneralPath & "BUGs.log" For Append Shared As N
        Print #N, "Usuario:" & .Name & "  Fecha:" & Date & "    Hora:" & Time
        Print #N, "BUG:"
        Print #N, bugReport
        Print #N, "########################################################################"
        Close #N
        

    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleBugReport de Protocol.bas")
End Sub

''
' Handles the "ChangeDescription" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChangeDescription(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim Description As String
        
        Description = Trim$(Reader.ReadString8())
        Description = Left$(Description, 50)  'Only 50 chars allowed.

        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "No puedes cambiar la descripción estando muerto.", FontTypeNames.FONTTYPE_INFO)
        Else
            If Description <> "" And Not AsciiValidos(Description, True) Then
                Call WriteConsoleMsg(UserIndex, "La descripción tiene caracteres inválidos.", FontTypeNames.FONTTYPE_INFO)
            Else
                .Desc = Trim$(Description)
                Call WriteConsoleMsg(UserIndex, "La descripción ha cambiado.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleChangeDescription de Protocol.bas")
End Sub

''
' Handles the "Punishments" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePunishments(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 25/08/2009
'25/08/2009: ZaMa - Now only admins can see other admins' punishment list
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim Name As String
        Dim Count As Integer
        
        Name = Reader.ReadString8()

        If LenB(Name) <> 0 Then
            If (InStrB(Name, "\") <> 0) Then
                Name = Replace(Name, "\", "")
            End If
            If (InStrB(Name, "/") <> 0) Then
                Name = Replace(Name, "/", "")
            End If
            If (InStrB(Name, ":") <> 0) Then
                Name = Replace(Name, ":", "")
            End If
            If (InStrB(Name, "|") <> 0) Then
                Name = Replace(Name, "|", "")
            End If
            
            If (EsAdmin(Name) Or EsDios(Name) Or EsSemiDios(Name) Or EsConsejero(Name) Or EsRolesMaster(Name)) And (UserList(UserIndex).flags.Privilegios And PlayerType.User) Then
                Call WriteConsoleMsg(UserIndex, "No puedes ver las penas de los administradores.", FontTypeNames.FONTTYPE_INFO)
            Else

                Dim UserId As Long
                UserId = GetUserID(Name)
                
                If UserId <> 0 Then
                    Call SendUserPunishments(UserIndex, UserId)
                Else
                    Call WriteConsoleMsg(UserIndex, "Personaje """ & Name & """ inexistente.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
    
    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandlePunishments de Protocol.bas")
End Sub

''
' Handles the "Gamble" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGamble(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'10/07/2010: ZaMa - Now normal npcs don't answer if asked to gamble.
'***************************************************

On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim Amount As Integer
        
        Amount = Reader.ReadInt16()
        
        ' Dead?
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
        
        'Validate target NPC
        ElseIf .flags.TargetNpc = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
        
        ' Validate Distance
        ElseIf Distancia(Npclist(.flags.TargetNpc).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
        
        ' Validate NpcType
        ElseIf Npclist(.flags.TargetNpc).NPCtype <> eNPCType.Timbero Then
            
            
            Dim TargetNpcType As eNPCType
            TargetNpcType = Npclist(.flags.TargetNpc).NPCtype
            
            ' Normal npcs don't speak
            If TargetNpcType <> eNPCType.Comun And TargetNpcType <> eNPCType.DRAGON And TargetNpcType <> eNPCType.Pretoriano Then
                Call WriteChatOverHead(UserIndex, "No tengo ningún interés en apostar.", Npclist(.flags.TargetNpc).Char.CharIndex, vbWhite)
            End If
            
        ' Validate amount
        ElseIf Amount < 1 Then
            Call WriteChatOverHead(UserIndex, "El mínimo de apuesta es 1 moneda.", Npclist(.flags.TargetNpc).Char.CharIndex, vbWhite)
        
        ' Validate amount
        ElseIf Amount > 5000 Then
            Call WriteChatOverHead(UserIndex, "El máximo de apuesta es 5000 monedas.", Npclist(.flags.TargetNpc).Char.CharIndex, vbWhite)
        
        ' Validate user gold
        ElseIf .Stats.GLD < Amount Then
            Call WriteChatOverHead(UserIndex, "No tienes esa cantidad.", Npclist(.flags.TargetNpc).Char.CharIndex, vbWhite)
        
        Else
            If RandomNumber(1, 100) <= 47 Then
                .Stats.GLD = .Stats.GLD + Amount
                Call WriteChatOverHead(UserIndex, "¡Felicidades! Has ganado " & CStr(Amount) & " monedas de oro.", Npclist(.flags.TargetNpc).Char.CharIndex, vbWhite)
                
                Apuestas.Perdidas = Apuestas.Perdidas + Amount
                Call WriteVar(DatPath & "apuestas.dat", "Main", "Perdidas", CStr(Apuestas.Perdidas))
            Else
                .Stats.GLD = .Stats.GLD - Amount
                Call WriteChatOverHead(UserIndex, "Lo siento, has perdido " & CStr(Amount) & " monedas de oro.", Npclist(.flags.TargetNpc).Char.CharIndex, vbWhite)
                
                Apuestas.Ganancias = Apuestas.Ganancias + Amount
                Call WriteVar(DatPath & "apuestas.dat", "Main", "Ganancias", CStr(Apuestas.Ganancias))
            End If
            
            Apuestas.Jugadas = Apuestas.Jugadas + 1
            
            Call WriteVar(DatPath & "apuestas.dat", "Main", "Jugadas", CStr(Apuestas.Jugadas))
            
            Call WriteUpdateGold(UserIndex)
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleGamble de Protocol.bas")
End Sub

''
' Handles the "InquiryVote" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInquiryVote(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim opt As Byte
        
        opt = Reader.ReadInt8()
        
        Call WriteConsoleMsg(UserIndex, ConsultaPopular.doVotar(UserIndex, opt), FontTypeNames.FONTTYPE_GUILD, eMessageType.Guild)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleInquiryVote de Protocol.bas")
End Sub

''
' Handles the "BankExtractGold" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankExtractGold(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim Amount As Long
        
        Amount = Reader.ReadInt32()
        
        'Dead people can't leave a faction.. they can't talk...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNpc = 0 Then
             Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
             Exit Sub
        End If
        
        If Npclist(.flags.TargetNpc).NPCtype <> eNPCType.Banquero Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNpc).Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Amount > 0 And Amount <= .Stats.Banco Then
             .Stats.Banco = .Stats.Banco - Amount
             .Stats.GLD = .Stats.GLD + Amount
             Call WriteChatOverHead(UserIndex, "Tenés " & .Stats.Banco & " monedas de oro en tu cuenta.", Npclist(.flags.TargetNpc).Char.CharIndex, vbWhite)
        Else
            Call WriteChatOverHead(UserIndex, "No tienes esa cantidad.", Npclist(.flags.TargetNpc).Char.CharIndex, vbWhite)
        End If
        
        Call WriteUpdateGold(UserIndex)
        Call WriteUpdateBankGold(UserIndex)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleBankExtractGold de Protocol.bas")
End Sub

''
' Handles the "LeaveFaction" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLeaveFaction(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 09/28/2010
' 09/28/2010 C4b3z0n - Ahora la respuesta de los NPCs sino perteneces a ninguna facción solo la hacen el Rey o el Demonio
' 05/17/06 - Maraxus
'***************************************************
On Error GoTo ErrHandler

    Dim TalkToKing As Boolean
    Dim TalkToDemon As Boolean
    Dim NpcIndex As Integer
    
    With UserList(UserIndex)

        'Dead people can't leave a faction.. they can't talk...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        ' Chequea si habla con el rey o el demonio. Puede salir sin hacerlo, pero si lo hace le reponden los npcs
        NpcIndex = .flags.TargetNpc
        
        ' Needs an NPC as a target
        If NpcIndex <= 0 Then Exit Sub
      
        ' Validate distance
        If Distancia(.Pos, Npclist(.flags.TargetNpc).Pos) > 4 Then
            Call WriteConsoleMsg(UserIndex, "Estás muy lejos del NPC.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        ' Es rey o domonio?
        If Npclist(NpcIndex).NPCtype = eNPCType.Noble Then
            'Rey?
            If Npclist(NpcIndex).flags.Faccion = 0 Then
                TalkToKing = True
            ' Demonio
            Else
                TalkToDemon = True
            End If
        End If
        
        ' Needs 2 very specific NPCs for this command to work.
        If Not TalkToKing And Not TalkToDemon Then Exit Sub
        
        ' Run some validations
        If .Faccion.Alignment = eCharacterAlignment.Neutral Or .Faccion.Alignment = eCharacterAlignment.Newbie Then
            Call WriteConsoleMsg(UserIndex, "¡No perteneces a ninguna facción!", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub
        End If
        
        If (TalkToDemon And .Faccion.Alignment <> eCharacterAlignment.FactionLegion) Or (TalkToKing And .Faccion.Alignment <> eCharacterAlignment.FactionRoyal) Then
            Call WriteChatOverHead(UserIndex, "No perteneces a nuestra facción. Si deseas unirte, di /ENLISTAR", _
                                   Npclist(NpcIndex).Char.CharIndex, vbWhite)
            Exit Sub
        End If
        
        If TalkToDemon And .Faccion.Alignment = eCharacterAlignment.FactionRoyal Then
            Call WriteChatOverHead(UserIndex, "¡¡¡Sal de aquí bufón!!!", Npclist(NpcIndex).Char.CharIndex, vbWhite)
            Exit Sub
        End If
        
        If TalkToKing And .Faccion.Alignment = eCharacterAlignment.FactionLegion Then
            Call WriteChatOverHead(UserIndex, "¡¡¡Sal de aquí maldito criminal!!!", Npclist(NpcIndex).Char.CharIndex, vbWhite)
            Exit Sub
        End If
        
        'Quit the Royal Army?
        If .Faccion.ArmadaReal = 1 Or .Faccion.Alignment = eCharacterAlignment.FactionRoyal Then
            ' Si le pidio al demonio salir de la armada, este le responde.
            If TalkToDemon Then
                Call WriteChatOverHead(UserIndex, "¡¡¡Sal de aquí bufón!!!", _
                                       Npclist(NpcIndex).Char.CharIndex, vbWhite)
            
            Else
                ' Si le pidio al rey salir de la armada, le responde.
                If TalkToKing Then
                    Call WriteChatOverHead(UserIndex, "Serás bienvenido a las fuerzas imperiales si deseas regresar.", _
                                           Npclist(NpcIndex).Char.CharIndex, vbWhite)
                End If
                
                Call ExpulsarFaccionReal(UserIndex, False)
                Call RefreshCharStatus(UserIndex, True)
            End If
        
        'Quit the Chaos Legion?
        ElseIf .Faccion.FuerzasCaos = 1 Or .Faccion.Alignment = eCharacterAlignment.FactionLegion Then
            ' Si le pidio al rey salir del caos, le responde.
            If TalkToKing Then
                Call WriteChatOverHead(UserIndex, "¡¡¡Sal de aquí maldito criminal!!!", _
                                       Npclist(NpcIndex).Char.CharIndex, vbWhite)
            Else
                ' Si le pidio al demonio salir del caos, este le responde.
                If TalkToDemon Then
                    Call WriteChatOverHead(UserIndex, "Ya volverás arrastrandote.", _
                                           Npclist(NpcIndex).Char.CharIndex, vbWhite)
                End If
                
                Call ExpulsarFaccionCaos(UserIndex, False)
                Call RefreshCharStatus(UserIndex, True)
            End If
        End If
        
    End With
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleLeaveFaction de Protocol.bas")
End Sub

''
' Handles the "BankDepositGold" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankDepositGold(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        Dim Amount As Long
        
        Amount = Reader.ReadInt32()
        
        'Dead people can't leave a faction.. they can't talk...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNpc = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Distancia(Npclist(.flags.TargetNpc).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Npclist(.flags.TargetNpc).NPCtype <> eNPCType.Banquero Then Exit Sub
        
        If Amount > 0 And Amount <= .Stats.GLD Then
            .Stats.Banco = .Stats.Banco + Amount
            .Stats.GLD = .Stats.GLD - Amount
            Call WriteChatOverHead(UserIndex, "Tenés " & .Stats.Banco & " monedas de oro en tu cuenta.", Npclist(.flags.TargetNpc).Char.CharIndex, vbWhite)
            
            Call WriteUpdateGold(UserIndex)
            Call WriteUpdateBankGold(UserIndex)
        Else
            Call WriteChatOverHead(UserIndex, "No tenés esa cantidad.", Npclist(.flags.TargetNpc).Char.CharIndex, vbWhite)
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleBankDepositGold de Protocol.bas")
End Sub

''
' Handles the "Denounce" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDenounce(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 14/11/2010
'14/11/2010: ZaMa - Now denounces can be desactivated.
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim Text As String
        Dim msg As String
        
        Text = Reader.ReadString8()

        If .flags.Silenciado = 0 Then
            'Analize chat...
            Call Statistics.ParseChat(Text)
            
            msg = LCase$(.Name) & " DENUNCIA: " & Text
            
            Call SendData(SendTarget.ToAdmins, 0, _
                PrepareMessageConsoleMsg(msg, FontTypeNames.FONTTYPE_GUILDMSG), True)
            
            Call Denuncias.Push(msg, False)
            
            Call WriteConsoleMsg(UserIndex, "Denuncia enviada, espere..", FontTypeNames.FONTTYPE_INFO)
        End If
        
    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleDenounce de Protocol.bas")
End Sub

''
' Handles the "PartyKick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyKick(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/05/09
'Last Modification by: Marco Vanotti (Marco)
'- 05/05/09: Now it uses "UserPuedeEjecutarComandos" to check if the user can use party commands
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim UserName As String
        Dim tUser As Integer
        
        UserName = Reader.ReadString8()

        If UserPuedeEjecutarComandos(UserIndex) Then
            tUser = NameIndex(UserName)
            
            If tUser > 0 Then
                Call mdParty.ExpulsarDeParty(UserIndex, tUser)
            Else
                If InStr(UserName, "+") Then
                    UserName = Replace(UserName, "+", " ")
                End If
                
                Call WriteConsoleMsg(UserIndex, LCase$(UserName) & " no pertenece a tu party.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandlePartyKick de Protocol.bas")
End Sub

''
' Handles the "PartySetLeader" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartySetLeader(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/05/09
'Last Modification by: Marco Vanotti (MarKoxX)
'- 05/05/09: Now it uses "UserPuedeEjecutarComandos" to check if the user can use party commands
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim UserName As String
        Dim tUser As Integer
        Dim rank As Integer
        
        UserName = Reader.ReadString8()

        rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero

        If UserPuedeEjecutarComandos(UserIndex) Then
            tUser = NameIndex(UserName)
            If tUser > 0 Then
                'Don't allow users to spoof online GMs
                If (UserDarPrivilegioLevel(UserName) And rank) <= (.flags.Privilegios And rank) Then
                    Call mdParty.TransformarEnLider(UserIndex, tUser)
                Else
                    Call WriteConsoleMsg(UserIndex, LCase$(UserList(tUser).Name) & " no pertenece a tu party.", FontTypeNames.FONTTYPE_INFO)
                End If
                
            Else
                If InStr(UserName, "+") Then
                    UserName = Replace(UserName, "+", " ")
                End If
                Call WriteConsoleMsg(UserIndex, LCase$(UserName) & " no pertenece a tu party.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
    End With
    

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandlePartySetLeader de Protocol.bas")
End Sub

''
' Handles the "GMMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGMMessage(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/08/07
'Last Modification by: (liquid)
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim Message As String
        
        Message = Reader.ReadString8()

        If Not .flags.Privilegios And PlayerType.User Then
            Call LogGM(.Name, "Mensaje a Gms:" & Message)
        
            If LenB(Message) <> 0 Then
                'Analize chat...
                Call Statistics.ParseChat(Message)
                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.Name & "> " & Message, FontTypeNames.FONTTYPE_GMMSG))
            End If
        End If
        
    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleGMMessage de Protocol.bas")
End Sub

''
' Handles the "ShowName" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleShowName(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            .ShowName = Not .ShowName 'Show / Hide the name
            
            Call RefreshCharStatus(UserIndex, False)
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleShowName de Protocol.bas")
End Sub

''
' Handles the "OnlineRoyalArmy" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineRoyalArmy(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 28/05/2010
'28/05/2010: ZaMa - Ahora solo dioses pueden ver otros dioses online.
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        If .flags.Privilegios And PlayerType.User Then Exit Sub
    
        Dim I As Long
        Dim list As String
        Dim Priv As PlayerType

        Priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios
        
        ' Solo dioses pueden ver otros dioses online
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
            Priv = Priv Or PlayerType.Dios Or PlayerType.Admin
        End If
     
        For I = 1 To LastUser
            If UserList(I).ConnIDValida Then
                If UserList(I).Faccion.ArmadaReal = 1 Then
                    If UserList(I).flags.Privilegios And Priv Then
                        list = list & UserList(I).Name & ", "
                    End If
                End If
            End If
        Next I
    End With
    
    If Len(list) > 0 Then
        Call WriteConsoleMsg(UserIndex, "Reales conectados: " & Left$(list, Len(list) - 2), FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(UserIndex, "No hay reales conectados.", FontTypeNames.FONTTYPE_INFO)
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleOnlineRoyalArmy de Protocol.bas")
End Sub

''
' Handles the "OnlineChaosLegion" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineChaosLegion(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 28/05/2010
'28/05/2010: ZaMa - Ahora solo dioses pueden ver otros dioses online.
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        If .flags.Privilegios And PlayerType.User Then Exit Sub
    
        Dim I As Long
        Dim list As String
        Dim Priv As PlayerType

        Priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios
        
        ' Solo dioses pueden ver otros dioses online
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
            Priv = Priv Or PlayerType.Dios Or PlayerType.Admin
        End If
     
        For I = 1 To LastUser
            If UserList(I).ConnIDValida Then
                If UserList(I).Faccion.FuerzasCaos = 1 Then
                    If UserList(I).flags.Privilegios And Priv Then
                        list = list & UserList(I).Name & ", "
                    End If
                End If
            End If
        Next I
    End With

    If Len(list) > 0 Then
        Call WriteConsoleMsg(UserIndex, "Caos conectados: " & Left$(list, Len(list) - 2), FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(UserIndex, "No hay Caos conectados.", FontTypeNames.FONTTYPE_INFO)
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleOnlineChaosLegion de Protocol.bas")
End Sub

''
' Handles the "GoNearby" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGoNearby(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/10/07
'
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim UserName As String
        
        UserName = Reader.ReadString8()

        Dim tIndex As Integer
        Dim X As Long
        Dim Y As Long
        Dim I As Long
        Dim Found As Boolean
        
        tIndex = NameIndex(UserName)
        
        'Check the user has enough powers
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then
            'Si es dios o Admins no podemos salvo que nosotros también lo seamos
            If (Not (EsDios(UserName) Or EsAdmin(UserName))) Or (((.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0) And ((.flags.Privilegios And PlayerType.RoleMaster) = 0)) Then
                If tIndex <= 0 Then 'existe el usuario destino?
                    Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
                Else
                    For I = 2 To 5 'esto for sirve ir cambiando la distancia destino
                        For X = UserList(tIndex).Pos.X - I To UserList(tIndex).Pos.X + I
                            For Y = UserList(tIndex).Pos.Y - I To UserList(tIndex).Pos.Y + I
                                If MapData(UserList(tIndex).Pos.Map, X, Y).UserIndex = 0 Then
                                    If LegalPos(UserList(tIndex).Pos.Map, X, Y, True, True) Then
                                        Call WarpUserChar(UserIndex, UserList(tIndex).Pos.Map, X, Y, True)
                                        Call LogGM(.Name, "/IRCERCA " & UserName & " Mapa:" & UserList(tIndex).Pos.Map & " X:" & UserList(tIndex).Pos.X & " Y:" & UserList(tIndex).Pos.Y)
                                        Found = True
                                        Exit For
                                    End If
                                End If
                            Next Y
                            
                            If Found Then Exit For  ' Feo, pero hay que abortar 3 fors sin usar GoTo
                        Next X
                        
                        If Found Then Exit For  ' Feo, pero hay que abortar 3 fors sin usar GoTo
                    Next I
                    
                    'No space found??
                    If Not Found Then
                        Call WriteConsoleMsg(UserIndex, "Todos los lugares están ocupados.", FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
            End If
        End If
        
    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleGoNearby de Protocol.bas")
End Sub

''
' Handles the "Comment" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleComment(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim comment As String
        comment = Reader.ReadString8()

        If Not .flags.Privilegios And PlayerType.User Then
            Call LogGM(.Name, "Comentario: " & comment)
            Call WriteConsoleMsg(UserIndex, "Comentario salvado...", FontTypeNames.FONTTYPE_INFO)
        End If

    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleComment de Protocol.bas")
End Sub

''
' Handles the "ServerTime" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleServerTime(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/08/07
'Last Modification by: (liquid)
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        If .flags.Privilegios And PlayerType.User Then Exit Sub
    
        Call LogGM(.Name, "Hora.")
    End With
    
    Call modSendData.SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Hora: " & Time & " " & Date, FontTypeNames.FONTTYPE_INFO))
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleServerTime de Protocol.bas")
End Sub

''
' Handles the "Where" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWhere(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 18/11/2010
'07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
'18/11/2010: ZaMa - Obtengo los privs del charfile antes de mostrar la posicion de un usuario offline.
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim UserName As String
        Dim tUser As Integer
        Dim MiPos As String
        
        UserName = Reader.ReadString8()

        If Not .flags.Privilegios And PlayerType.User Then
            
            tUser = NameIndex(UserName)
            If tUser <= 0 Then

                Dim UserId As Long
                UserId = GetUserID(UserName)
                
                If UserId <> 0 Then

                    Dim CharPrivs As PlayerType
                    CharPrivs = GetCharPrivs(UserName)
                    
                    If (CharPrivs And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios)) <> 0 Or ((CharPrivs And (PlayerType.Dios Or PlayerType.Admin) <> 0) And (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0) Then
                        
                        MiPos = GetCharData("USER_INFO", "LAST_POS", UserId)

                        Call WriteConsoleMsg(UserIndex, "Ubicación  " & UserName & " (Offline): " & ReadField(1, MiPos, 45) & ", " & ReadField(2, MiPos, 45) & ", " & ReadField(3, MiPos, 45) & ".", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    If Not (EsDios(UserName) Or EsAdmin(UserName)) Then
                        Call WriteConsoleMsg(UserIndex, "Usuario inexistente.", FontTypeNames.FONTTYPE_INFO)
                    ElseIf .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
                        Call WriteConsoleMsg(UserIndex, "Usuario inexistente.", FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
            Else
                If (UserList(tUser).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios)) <> 0 Or ((UserList(tUser).flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) <> 0) And (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0) Then
                    Call WriteConsoleMsg(UserIndex, "Ubicación  " & UserName & ": " & UserList(tUser).Pos.Map & ", " & UserList(tUser).Pos.X & ", " & UserList(tUser).Pos.Y & ".", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
            
            Call LogGM(.Name, "/Donde " & UserName)
        End If

    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleWhere de Protocol.bas")
End Sub

''
' Handles the "CreaturesInMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCreaturesInMap(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 30/07/06
'Pablo (ToxicWaste): modificaciones generales para simplificar la visualización.
'***************************************************

On Error GoTo ErrHandler
      
    With UserList(UserIndex)

        Dim Map As Integer
        Dim I, J As Long
        Dim NPCcount1, NPCcount2 As Integer
        Dim NPCcant1() As Integer
        Dim NPCcant2() As Integer
        Dim List1() As String
        Dim List2() As String
        
        Map = Reader.ReadInt16()
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        
        If MapaValido(Map) Then
            For I = 1 To LastNPC
                'VB isn't lazzy, so we put more restrictive condition first to speed up the process
                If Npclist(I).Pos.Map = Map Then
                    '¿esta vivo?
                    If Npclist(I).flags.NPCActive And Npclist(I).Hostile = 1 And Npclist(I).Stats.Alineacion = 2 Then
                        If NPCcount1 = 0 Then
                            ReDim List1(0) As String
                            ReDim NPCcant1(0) As Integer
                            NPCcount1 = 1
                            List1(0) = Npclist(I).Name & ": (" & Npclist(I).Pos.X & "," & Npclist(I).Pos.Y & ")"
                            NPCcant1(0) = 1
                        Else
                            For J = 0 To NPCcount1 - 1
                                If Left$(List1(J), Len(Npclist(I).Name)) = Npclist(I).Name Then
                                    List1(J) = List1(J) & ", (" & Npclist(I).Pos.X & "," & Npclist(I).Pos.Y & ")"
                                    NPCcant1(J) = NPCcant1(J) + 1
                                    Exit For
                                End If
                            Next J
                            If J = NPCcount1 Then
                                ReDim Preserve List1(0 To NPCcount1) As String
                                ReDim Preserve NPCcant1(0 To NPCcount1) As Integer
                                NPCcount1 = NPCcount1 + 1
                                List1(J) = Npclist(I).Name & ": (" & Npclist(I).Pos.X & "," & Npclist(I).Pos.Y & ")"
                                NPCcant1(J) = 1
                            End If
                        End If
                    Else
                        If NPCcount2 = 0 Then
                            ReDim List2(0) As String
                            ReDim NPCcant2(0) As Integer
                            NPCcount2 = 1
                            List2(0) = Npclist(I).Name & ": (" & Npclist(I).Pos.X & "," & Npclist(I).Pos.Y & ")"
                            NPCcant2(0) = 1
                        Else
                            For J = 0 To NPCcount2 - 1
                                If Left$(List2(J), Len(Npclist(I).Name)) = Npclist(I).Name Then
                                    List2(J) = List2(J) & ", (" & Npclist(I).Pos.X & "," & Npclist(I).Pos.Y & ")"
                                    NPCcant2(J) = NPCcant2(J) + 1
                                    Exit For
                                End If
                            Next J
                            If J = NPCcount2 Then
                                ReDim Preserve List2(0 To NPCcount2) As String
                                ReDim Preserve NPCcant2(0 To NPCcount2) As Integer
                                NPCcount2 = NPCcount2 + 1
                                List2(J) = Npclist(I).Name & ": (" & Npclist(I).Pos.X & "," & Npclist(I).Pos.Y & ")"
                                NPCcant2(J) = 1
                            End If
                        End If
                    End If
                End If
            Next I
            
            Call WriteConsoleMsg(UserIndex, "Npcs Hostiles en mapa: ", FontTypeNames.FONTTYPE_WARNING)
            If NPCcount1 = 0 Then
                Call WriteConsoleMsg(UserIndex, "No hay NPCS Hostiles.", FontTypeNames.FONTTYPE_INFO)
            Else
                For J = 0 To NPCcount1 - 1
                    Call WriteConsoleMsg(UserIndex, NPCcant1(J) & " " & List1(J), FontTypeNames.FONTTYPE_INFO)
                Next J
            End If
            Call WriteConsoleMsg(UserIndex, "Otros Npcs en mapa: ", FontTypeNames.FONTTYPE_WARNING)
            If NPCcount2 = 0 Then
                Call WriteConsoleMsg(UserIndex, "No hay más NPCS.", FontTypeNames.FONTTYPE_INFO)
            Else
                For J = 0 To NPCcount2 - 1
                    Call WriteConsoleMsg(UserIndex, NPCcant2(J) & " " & List2(J), FontTypeNames.FONTTYPE_INFO)
                Next J
            End If
            Call LogGM(.Name, "Numero enemigos en mapa " & Map)
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleCreaturesInMap de Protocol.bas")
End Sub

''
' Handles the "WarpMeToTarget" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWarpMeToTarget(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 26/03/09
'26/03/06: ZaMa - Chequeo que no se teletransporte donde haya un char o npc
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        Dim X As Integer
        Dim Y As Integer
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        
        X = .flags.TargetX
        Y = .flags.TargetY
        
        Call FindLegalPos(UserIndex, .flags.TargetMap, X, Y)
        Call WarpUserChar(UserIndex, .flags.TargetMap, X, Y, True)
        Call LogGM(.Name, "/TELEPLOC a x:" & .flags.TargetX & " Y:" & .flags.TargetY & " Map:" & .Pos.Map)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleWarpMeToTarget de Protocol.bas")
End Sub

''
' Handles the "WarpChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWarpChar(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 26/03/2009
'26/03/2009: ZaMa -  Chequeo que no se teletransporte a un tile donde haya un char o npc.
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim UserName As String
        Dim Map As Integer
        Dim X As Integer
        Dim Y As Integer
        Dim tUser As Integer
        
        UserName = Reader.ReadString8()
        Map = Reader.ReadInt16()
        X = Reader.ReadInt8()
        Y = Reader.ReadInt8()

        If Not .flags.Privilegios And PlayerType.User Then
            If MapaValido(Map) And LenB(UserName) <> 0 Then
                If UCase$(UserName) <> "YO" Then
                    If Not .flags.Privilegios And PlayerType.Consejero Then
                        tUser = NameIndex(UserName)
                    End If
                Else
                    tUser = UserIndex
                End If
            
                If tUser <= 0 Then
                    If Not (EsDios(UserName) Or EsAdmin(UserName)) Then
                        Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call WriteConsoleMsg(UserIndex, "No puedes transportar dioses o admins.", FontTypeNames.FONTTYPE_INFO)
                    End If
                    
                ElseIf Not ((UserList(tUser).flags.Privilegios And PlayerType.Dios) <> 0 Or _
                            (UserList(tUser).flags.Privilegios And PlayerType.Admin) <> 0) Or _
                           tUser = UserIndex Then
                            
                    If InMapBounds(Map, X, Y) Then
                        Call FindLegalPos(tUser, Map, X, Y)
                        Call WarpUserChar(tUser, Map, X, Y, True, True, True)
                        Call WriteConsoleMsg(UserIndex, UserList(tUser).Name & " transportado.", FontTypeNames.FONTTYPE_INFO)
                        Call LogGM(.Name, "Transportó a " & UserList(tUser).Name & " hacia " & "Mapa" & Map & " X:" & X & " Y:" & Y)
                    End If
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes transportar dioses o admins.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If

    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleWarpUser de Protocol.bas")
End Sub

''
' Handles the "Silence" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSilence(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim UserName As String
        Dim tUser As Integer
        
        UserName = Reader.ReadString8()

        If Not .flags.Privilegios And PlayerType.User Then
            tUser = NameIndex(UserName)
        
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else
                If UserList(tUser).flags.Silenciado = 0 Then
                    UserList(tUser).flags.Silenciado = 1
                    Call WriteConsoleMsg(UserIndex, "Usuario silenciado.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteShowMessageBox(tUser, "Estimado usuario, ud. ha sido silenciado por los administradores. Sus denuncias serán ignoradas por el servidor de aquí en más. Utilice /GM para contactar un administrador.")
                    Call LogGM(.Name, "/silenciar " & UserList(tUser).Name)
                Else
                    UserList(tUser).flags.Silenciado = 0
                    Call WriteConsoleMsg(UserIndex, "Usuario des silenciado.", FontTypeNames.FONTTYPE_INFO)
                    Call LogGM(.Name, "/DESsilenciar " & UserList(tUser).Name)
                End If
            End If
        End If
    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleSilence de Protocol.bas")
End Sub

''
' Handles the "SOSShowList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSOSShowList(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        If .flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster) Then Exit Sub
        Call WriteShowSOSForm(UserIndex)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleSOSShowList de Protocol.bas")
End Sub

''
' Handles the "RequestPartyForm" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyForm(ByVal UserIndex As Integer)
'***************************************************
'Author: Budi
'Last Modification: 11/26/09
'
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        If .PartyIndex > 0 Then
            Call WriteShowPartyForm(UserIndex)
            
        Else
            Call WriteConsoleMsg(UserIndex, "No perteneces a ningún grupo!", FontTypeNames.FONTTYPE_INFOBOLD)
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandlePartyForm de Protocol.bas")
End Sub

''
' Handles the "ItemUpgrade" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleItemUpgrade(ByVal UserIndex As Integer)
'***************************************************
'Author: Torres Patricio
'Last Modification: 12/09/09
'
'***************************************************

On Error GoTo ErrHandler
    
    With UserList(UserIndex)
        Dim ItemIndex As Integer

        ItemIndex = Reader.ReadInt16()
        
        If ItemIndex <= 0 Then Exit Sub
        If Not TieneObjetos(ItemIndex, 1, UserIndex) Then Exit Sub
        
        If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
        'Call DoUpgrade(UserIndex, ItemIndex)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleItemUpgrade de Protocol.bas")
End Sub

''
' Handles the "SOSRemove" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSOSRemove(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim UserName As String
        UserName = Reader.ReadString8()
        
        If Not .flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster) Then _
            Call Ayuda.Quitar(UserName)
    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleSOSRemove de Protocol.bas")
End Sub

''
' Handles the "GoToChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGoToChar(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 26/03/2009
'26/03/2009: ZaMa -  Chequeo que no se teletransporte a un tile donde haya un char o npc.
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim UserName As String
        Dim tUser As Integer
        Dim X As Integer
        Dim Y As Integer
        
        UserName = Reader.ReadString8()

        tUser = NameIndex(UserName)
        
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.SemiDios Or PlayerType.Consejero) Then
            'Si es dios o Admins no podemos salvo que nosotros también lo seamos
            If Not (EsDios(UserName) Or EsAdmin(UserName)) Or (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Then
                If tUser <= 0 Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
                Else
                    X = UserList(tUser).Pos.X
                    Y = UserList(tUser).Pos.Y + 1
                    Call FindLegalPos(UserIndex, UserList(tUser).Pos.Map, X, Y)
                    
                    Call WarpUserChar(UserIndex, UserList(tUser).Pos.Map, X, Y, True)
                    
                    If .flags.AdminInvisible = 0 Then
                        Call WriteConsoleMsg(tUser, .Name & " se ha trasportado hacia donde te encuentras.", FontTypeNames.FONTTYPE_INFO)
                    End If
                    
                    Call LogGM(.Name, "/IRA " & UserName & " Mapa:" & UserList(tUser).Pos.Map & " X:" & UserList(tUser).Pos.X & " Y:" & UserList(tUser).Pos.Y)
                End If
            End If
        End If
        
    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleGoToChar de Protocol.bas")
End Sub

''
' Handles the "Invisible" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInvisible(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        
        Call DoAdminInvisible(UserIndex)
        Call LogGM(.Name, "/INVISIBLE")
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleInvisible de Protocol.bas")
End Sub

''
' Handles the "GMPanel" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGMPanel(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        If .flags.Privilegios And PlayerType.User Then Exit Sub
        
        Call WriteShowGMPanelForm(UserIndex)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleGMPanel de Protocol.bas")
End Sub

''
' Handles the "GMPanel" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestUserList(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/09/07
'Last modified by: Lucas Tavolaro Ortiz (Tavo)
'I haven`t found a solution to split, so i make an array of names
'***************************************************
On Error GoTo ErrHandler
  
    Dim I As Long
    Dim names() As String
    Dim Count As Long
    
    With UserList(UserIndex)

        If .flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster) Then Exit Sub
        
        ReDim names(1 To LastUser) As String
        Count = 1
        
        For I = 1 To LastUser
            If (LenB(UserList(I).Name) <> 0) Then
                If UserList(I).flags.Privilegios And PlayerType.User Then
                    names(Count) = UserList(I).Name
                    Count = Count + 1
                End If
            End If
        Next I
        
        If Count > 1 Then Call WriteUserNameList(UserIndex, names(), Count - 1)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleRequestUserList de Protocol.bas")
End Sub


''
' Handles the "Jail" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleJail(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 07/06/2010
'07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim UserName As String
        Dim Reason As String
        Dim AdminNotes As String
        Dim jailTime As Byte
        Dim Count As Byte
        Dim tUser As Integer
        Dim PunishmentTypeId As Integer
        
        UserName = Reader.ReadString8()
        Reason = Reader.ReadString8()
        PunishmentTypeId = Reader.ReadInt16()
        AdminNotes = Reader.ReadString8()

        If InStr(1, UserName, "+") Then
            UserName = Replace(UserName, "+", " ")
        End If
        
        '/carcel nick
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (Not .flags.Privilegios And PlayerType.User) <> 0 Then
            If LenB(UserName) = 0 Then
                Call WriteConsoleMsg(UserIndex, "Utilice /carcel nick", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            If (EsDios(UserName) Or EsAdmin(UserName)) Then
                Call WriteConsoleMsg(UserIndex, "No puedes encarcelar a administradores.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            tUser = NameIndex(UserName)
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "El usuario no está online.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            Dim UserId As Long
            UserId = GetUserID(UserName)

            If Not UserList(tUser).flags.Privilegios And PlayerType.User Then
                Call WriteConsoleMsg(UserIndex, "No puedes encarcelar a administradores.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            If (InStrB(UserName, "\") <> 0) Then
               UserName = Replace(UserName, "\", "")
            End If
            If (InStrB(UserName, "/") <> 0) Then
               UserName = Replace(UserName, "/", "")
            End If
            

            Dim PunishmentResponse As tPunishmentDbResponse
            PunishmentResponse = AddPunishmentDB(UserId, .Id, PunishmentTypeId, Reason, AdminNotes)
           
            ' If a jail was added then we apply the punishment
            If PunishmentResponse.PunishmentBaseType = 1 Then
                Call Encarcelar(tUser, PunishmentResponse.PunishmentSeverity, .Name)
                
                Call LogGM(.Name, " encarceló a " & UserName)
            End If
        
            ' If the user got jailed several times for the same reason, the punishment can lead to a ban
            ' If that's the case, it will be considered a "forced punishment" of type 2, so we need to act
            ' accordingly
            If PunishmentResponse.ForcedPunismentBaseType = 2 Then
                UserList(tUser).flags.Ban = 1
                UserList(tUser).Punishment.Id = PunishmentResponse.LastInsertedPunishmentId
                
                Call CloseSocket(tUser, True)
                
                Call LogGM(.Name, " aplicó un ban forzado " & UserName & " por acumulación de penas.")
            End If
        
        End If

    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleJail de Protocol.bas")
End Sub

''
' Handles the "KillNPC" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKillNPC(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/22/08 (NicoNZ)
'
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        If .flags.Privilegios And PlayerType.User Then Exit Sub
        
        Dim tNPC As Integer
        Dim auxNPC As npc
        
        'Los consejeros no pueden RMATAr a nada en el mapa pretoriano
        If .flags.Privilegios And PlayerType.Consejero Then
            If .Pos.Map = MAPA_PRETORIANO Then
                Call WriteConsoleMsg(UserIndex, "Los consejeros no pueden usar este comando en el mapa pretoriano.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        
        tNPC = .flags.TargetNpc
        
        If tNPC > 0 Then
            Call WriteConsoleMsg(UserIndex, "RMatas (con posible respawn) a: " & Npclist(tNPC).Name, FontTypeNames.FONTTYPE_INFO)
            
            auxNPC = Npclist(tNPC)
            Call QuitarNPC(tNPC)
            Call ReSpawnNpc(auxNPC)
            
            .flags.TargetNpc = 0
        Else
            Call WriteConsoleMsg(UserIndex, "Antes debes hacer click sobre el NPC.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleKillNPC de Protocol.bas")
End Sub

''
' Handles the "WarnUser" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWarnUser(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/26/06
'
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim UserName As String
        Dim Reason As String
        Dim AdminNotes As String
        Dim Privs As PlayerType
        Dim Count As Byte
        Dim PunishmentTypeId As Integer
        
        UserName = Reader.ReadString8()
        Reason = Reader.ReadString8()
        AdminNotes = Reader.ReadString8()
        PunishmentTypeId = Reader.ReadInt16()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (Not .flags.Privilegios And PlayerType.User) <> 0 Then
            If LenB(UserName) = 0 Or LenB(Reason) = 0 Then
                Call WriteConsoleMsg(UserIndex, "Utilice /advertencia nick@motivo", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            Privs = UserDarPrivilegioLevel(UserName)
            
            If Not Privs And PlayerType.User Then
                Call WriteConsoleMsg(UserIndex, "No puedes advertir a administradores.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "\", "")
            End If
            
            If (InStrB(UserName, "/") <> 0) Then
                UserName = Replace(UserName, "/", "")
            End If

            Dim UserId As Long
            UserId = GetUserID(UserName)
            
            If UserId <> 0 Then
            
                Call AddPunishmentDB(UserId, .Id, PunishmentTypeId, Reason, AdminNotes)

                Call WriteConsoleMsg(UserIndex, "Has advertido a " & UCase$(UserName) & ".", FontTypeNames.FONTTYPE_INFO)
                Call LogGM(.Name, " advirtio a " & UserName)
            End If
            
            
        End If

    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleWarnUser de Protocol.bas")
End Sub

''
' Handles the "EditChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleEditChar(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 18/09/2010
'02/03/2009: ZaMa - Cuando editas nivel, chequea si el pj puede permanecer en clan faccionario
'11/06/2009: ZaMa - Todos los comandos se pueden usar aunque el pj este offline
'18/09/2010: ZaMa - Ahora se puede editar la vida del propio pj (cualquier rm o dios).
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim UserName As String
        Dim tUser As Integer
        Dim opcion As Byte
        Dim Arg1 As String
        Dim Arg2 As String
        Dim valido As Boolean
        Dim LoopC As Byte
        Dim CommandString As String
        Dim UserCharPath As String
        Dim Var As Long
        Dim NaturalSkillPoints As Byte
        Dim AssignedSkillPoints As Byte
        Dim Diff As Byte
        
        UserName = Reader.ReadString8()
        opcion = Reader.ReadInt8()
        Arg1 = Reader.ReadString8()
        Arg2 = Reader.ReadString8()

        UserName = Replace(UserName, "+", " ")
        
        If UCase$(UserName) = "YO" Then
            tUser = UserIndex
        Else
            tUser = NameIndex(UserName)
        End If
                
        If .flags.Privilegios And PlayerType.RoleMaster Then
            Select Case .flags.Privilegios And (PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero)
                Case PlayerType.Consejero
                    ' Los RMs consejeros sólo se pueden editar su head, body, level y vida
                    valido = tUser = UserIndex And _
                            (opcion = eEditOptions.eo_Body Or _
                             opcion = eEditOptions.eo_Head Or _
                             opcion = eEditOptions.eo_Level Or _
                             opcion = eEditOptions.eo_Vida)
                
                Case PlayerType.SemiDios
                    ' Los RMs sólo se pueden editar su level o vida y el head y body de cualquiera
                    valido = ((opcion = eEditOptions.eo_Level Or opcion = eEditOptions.eo_Vida) And tUser = UserIndex) Or _
                              opcion = eEditOptions.eo_Body Or _
                              opcion = eEditOptions.eo_Head
                    
                Case PlayerType.Dios
                    ' Los DRMs pueden aplicar los siguientes comandos sobre cualquiera
                    ' pero si quiere modificar el level o vida sólo lo puede hacer sobre sí mismo
                    valido = ((opcion = eEditOptions.eo_Level Or opcion = eEditOptions.eo_Vida) And tUser = UserIndex) Or _
                            opcion = eEditOptions.eo_Body Or _
                            opcion = eEditOptions.eo_Head Or _
                            opcion = eEditOptions.eo_CiticensKilled Or _
                            opcion = eEditOptions.eo_CriminalsKilled Or _
                            opcion = eEditOptions.eo_Class Or _
                            opcion = eEditOptions.eo_Skills Or _
                            opcion = eEditOptions.eo_addGold
            End Select
        
        'Si no es RM debe ser dios para poder usar este comando
        ElseIf .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
            
            If opcion = eEditOptions.eo_Vida Then
                '  Por ahora dejo para que los dioses no puedan editar la vida de otros
                valido = (tUser = UserIndex)
            Else
                valido = True
            End If
        
        ElseIf (.flags.Privilegios And PlayerType.SemiDios) Then
            
            valido = (opcion = eEditOptions.eo_Poss Or _
                     ((opcion = eEditOptions.eo_Vida) And (tUser = UserIndex)))
            
            If .flags.PrivEspecial Then
                valido = valido Or (opcion = eEditOptions.eo_CiticensKilled) Or _
                     (opcion = eEditOptions.eo_CriminalsKilled)
            End If
        
        ElseIf (.flags.Privilegios And PlayerType.Consejero) Then
            valido = ((opcion = eEditOptions.eo_Vida) And (tUser = UserIndex))
        End If
        
        If valido Then
            Dim UserId As Long
            
            If tUser <= 0 Then
                UserId = GetUserID(UserName)
            End If
            
            If tUser <= 0 And UserId = 0 Then
                Call WriteConsoleMsg(UserIndex, "Estás intentando editar un usuario inexistente.", FontTypeNames.FONTTYPE_INFO)
                Call LogGM(.Name, "Intentó editar un usuario inexistente.")
            Else
                'For making the Log
                CommandString = "/MOD "
                
                Select Case opcion
                    Case eEditOptions.eo_Gold
                        If Val(Arg1) <= MAX_ORO_EDIT Then
                            If tUser <= 0 Then ' Esta offline?
                                Call UpdateCharData("USER_STATS", "ORO", UserId, Arg1)
                                
                                Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                            Else ' Online
                                UserList(tUser).Stats.GLD = Val(Arg1)
                                Call WriteUpdateGold(tUser)
                            End If
                        Else
                            Call WriteConsoleMsg(UserIndex, "No está permitido utilizar valores mayores a " & MAX_ORO_EDIT & ". Su comando ha quedado en los logs del juego.", FontTypeNames.FONTTYPE_INFO)
                        End If
                    
                        ' Log it
                        CommandString = CommandString & "ORO "
                
                    Case eEditOptions.eo_Experience
                        If Val(Arg1) > 20000000 Then
                            Arg1 = 20000000
                        End If
                        
                        If tUser <= 0 Then ' Offline
                            Call UpdateCharData("USER_STATS", "EXP", UserId, "EXP + '" & Val(Arg1) & "'", False)

                            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else ' Online
                            UserList(tUser).Stats.Exp = UserList(tUser).Stats.Exp + Val(Arg1)
                            Call CheckUserLevel(tUser)
                        End If
                        
                        ' Log it
                        CommandString = CommandString & "EXP "
                    
                    Case eEditOptions.eo_Body
                        If tUser <= 0 Then
                            Call UpdateCharData("USER_INFO", "BODY", UserId, Arg1)
                            
                            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else
                            Call ChangeUserChar(tUser, Val(Arg1), UserList(tUser).Char.head, UserList(tUser).Char.heading, UserList(tUser).Char.WeaponAnim, UserList(tUser).Char.ShieldAnim, UserList(tUser).Char.CascoAnim)
                        End If
                        
                        ' Log it
                        CommandString = CommandString & "BODY "
                    
                    Case eEditOptions.eo_Head
                        If tUser <= 0 Then
                            Call UpdateCharData("USER_INFO", "HEAD", UserId, Arg1)

                            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else
                            Call ChangeUserChar(tUser, UserList(tUser).Char.body, Val(Arg1), UserList(tUser).Char.heading, UserList(tUser).Char.WeaponAnim, UserList(tUser).Char.ShieldAnim, UserList(tUser).Char.CascoAnim)
                        End If
                        
                        ' Log it
                        CommandString = CommandString & "HEAD "
                    
                    Case eEditOptions.eo_CriminalsKilled
                        Var = IIf(Val(Arg1) > ConstantesBalance.MaxUsersMatados, ConstantesBalance.MaxUsersMatados, Val(Arg1))
                        
                        If tUser <= 0 Then ' Offline
                            Call UpdateCharData("USER_FACTION", "CRI_KILLED", UserId, CStr(Var))

                            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else ' Online
                            UserList(tUser).Faccion.CriminalesMatados = Var
                        End If
                        
                        ' Log it
                        CommandString = CommandString & "CRI "
                    
                    Case eEditOptions.eo_CiticensKilled
                        Var = IIf(Val(Arg1) > ConstantesBalance.MaxUsersMatados, ConstantesBalance.MaxUsersMatados, Val(Arg1))
                        
                        If tUser <= 0 Then ' Offline
                            Call UpdateCharData("USER_FACTION", "CITY_KILLED", UserId, CStr(Var))

                            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else ' Online
                            UserList(tUser).Faccion.CiudadanosMatados = Var
                        End If
                        
                        ' Log it
                        CommandString = CommandString & "CIU "
                    
                    Case eEditOptions.eo_Level
                        If Val(Arg1) > ConstantesBalance.MaxLvl Or Val(Arg1) < 1 Then
                            Call WriteConsoleMsg(UserIndex, "El nivel del personaje tiene que ser un número entre 0 y " & ConstantesBalance.MaxLvl & ".", FONTTYPE_INFO)
                            Exit Sub
                        End If
                        
                        If tUser <= 0 Then ' Offline
                            Call UpdateCharData("USER_STATS", "NIVEL", UserId, Arg1)
                            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else ' Online
                            UserList(tUser).Stats.ELV = Val(Arg1)
                            Call WriteUpdateUserStats(tUser)
                        End If
                    
                        ' Log it
                        CommandString = CommandString & "LEVEL "
                    
                    Case eEditOptions.eo_Class
                        For LoopC = 1 To NUMCLASES
                            If UCase$(ListaClases(LoopC)) = UCase$(Arg1) Then Exit For
                        Next LoopC
                            
                        If LoopC > NUMCLASES Then
                            Call WriteConsoleMsg(UserIndex, "Clase desconocida. Intente nuevamente.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                        
                        If Not Classes(LoopC).Enabled Then
                            Call WriteConsoleMsg(UserIndex, "La clase seleccionada no se encuentra habilitada.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                        
                        
                        If tUser <= 0 Then ' Offline
                            Call UpdateCharData("USER_INFO", "CLASS", UserID, CStr(LoopC))
                            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else ' Online
                            UserList(tUser).clase = LoopC
                        End If
                        
                        ' Recalculate the user passives
                        Call RecalculateUserPassives(UserIndex, True)
                    
                        ' Log it
                        CommandString = CommandString & "CLASE "
                        
                    Case eEditOptions.eo_Skills
                        For LoopC = 1 To NUMSKILLS
                            If UCase$(Replace$(SkillsNames(LoopC), " ", "+")) = UCase$(Arg1) Then Exit For
                        Next LoopC
                       
                        If LoopC > NUMSKILLS Then
                            Call WriteConsoleMsg(UserIndex, "Skill Inexistente!", FontTypeNames.FONTTYPE_INFO)
                        Else
                        
                        If Val(Arg2) > MAX_SKILL_POINTS Then
                            Call WriteConsoleMsg(UserIndex, "Puedes asignar un máximo de " & MAX_SKILL_POINTS & " skills", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                        
                            If tUser <= 0 Then ' Offline
                                
                                Dim EluSkill As Long
                                If CByte(Arg2) < MAX_SKILL_POINTS Then
                                    EluSkill = ConstantesBalance.EluSkillInicial * 1.05 ^ Val(Arg2)
                                End If
                                
                                Call modDB_Functions.GetCharSkillDB(UserId, CByte(LoopC), NaturalSkillPoints, AssignedSkillPoints)
                                
                                Call CalculateNaturalAndAssignSkills(Val(Arg2), NaturalSkillPoints, AssignedSkillPoints)
                                
                                Call UpdateCharSkills(UserId, CByte(LoopC), NaturalSkillPoints, AssignedSkillPoints, EluSkill, 0)
                                Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                            Else ' Online
                                
                                NaturalSkillPoints = UserList(tUser).Stats.NaturalSkills(LoopC)
                                AssignedSkillPoints = UserList(tUser).Stats.AssignedSkills(LoopC)
                               
                                Call CalculateNaturalAndAssignSkills(Val(Arg2), NaturalSkillPoints, AssignedSkillPoints)
                      
                                UserList(tUser).Stats.NaturalSkills(LoopC) = NaturalSkillPoints
                                UserList(tUser).Stats.AssignedSkills(LoopC) = AssignedSkillPoints
                                
                                Call CheckEluSkill(tUser, LoopC, True)
                                
                                Call WriteConsoleMsg(UserIndex, "Skill " & SkillsNames(LoopC) & " editado a " & CStr(GetSkills(tUser, LoopC)) & "!", FontTypeNames.FONTTYPE_INFO)

                            End If
                        End If
                        
                        ' Log it
                        CommandString = CommandString & "SKILLS "
                    
                    Case eEditOptions.eo_SkillPointsLeft
                    
                            
                        If Val(Arg1) > MAX_SKILLS_LIBRES Then
                            Call WriteConsoleMsg(UserIndex, "Puedes asignar un máximo de " & MAX_SKILLS_LIBRES & " skills libres", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                        
                        If tUser <= 0 Then ' Offline
                            Call UpdateCharData("USER_STATS", "SKILLS", UserId, Arg1)

                            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else ' Online
                            UserList(tUser).Stats.SkillPts = Val(Arg1)
                        End If
                        
                        ' Log it
                        CommandString = CommandString & "SKILLSLIBRES "
                    
                    Case eEditOptions.eo_Sex
                        Dim Sex As Byte
                        
                        If UCase$(Arg1) = "MUJER" Then
                            Sex = eGenero.Mujer
                        ElseIf UCase$(Arg1) = "HOMBRE" Then
                            Sex = eGenero.Hombre
                        End If
                        
                        If Sex <> 0 Then ' Es Hombre o mujer?
                            If tUser <= 0 Then ' OffLine
                                Call UpdateCharData("USER_INFO", "GENDER", UserId, CStr(Sex))

                                Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                            Else ' Online
                                UserList(tUser).Genero = Sex
                            End If
                        Else
                            Call WriteConsoleMsg(UserIndex, "Genero desconocido. Intente nuevamente.", FontTypeNames.FONTTYPE_INFO)
                        End If
                        
                        ' Log it
                        CommandString = CommandString & "SEX "
                    
                    Case eEditOptions.eo_Raza
                        Dim raza As Byte
                        
                        Arg1 = UCase$(Arg1)
                        Select Case Arg1
                            Case "HUMANO"
                                raza = eRaza.Humano
                            Case "ELFO"
                                raza = eRaza.Elfo
                            Case "DROW"
                                raza = eRaza.Drow
                            Case "ENANO"
                                raza = eRaza.Enano
                            Case "GNOMO"
                                raza = eRaza.Gnomo
                            Case Else
                                raza = 0
                        End Select
                        
                            
                        If raza = 0 Then
                            Call WriteConsoleMsg(UserIndex, "Raza desconocida. Intente nuevamente.", FontTypeNames.FONTTYPE_INFO)
                        Else
                            If tUser <= 0 Then
                                Call UpdateCharData("USER_INFO", "RACE", UserId, CStr(raza))

                                Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                            Else
                                UserList(tUser).raza = raza
                            End If
                        End If
                            
                        ' Log it
                        CommandString = CommandString & "RAZA "
                        
                    Case eEditOptions.eo_addGold
                    
                        Dim BankGold As Long
                        
                        If Abs(Arg1) > MAX_ORO_EDIT Then
                            Call WriteConsoleMsg(UserIndex, "No está permitido utilizar valores mayores a " & MAX_ORO_EDIT & ".", FontTypeNames.FONTTYPE_INFO)
                        Else
                            If tUser <= 0 Then
                                Dim sConditionValue As String
                                sConditionValue = "IF(ORO_BANCO + '" & Val(Arg1) & "' <= '0', '0', ORO_BANCO + '" & Val(Arg1) & "')"
                                
                                Call UpdateCharData("USER_STATS", "ORO_BANCO", UserId, sConditionValue, False)

                                Call WriteConsoleMsg(UserIndex, "Se le ha agregado " & Arg1 & " monedas de oro a " & UserName & ".", FONTTYPE_TALK)
                            Else
                                UserList(tUser).Stats.Banco = IIf(UserList(tUser).Stats.Banco + Val(Arg1) <= 0, 0, UserList(tUser).Stats.Banco + Val(Arg1))
                                Call WriteConsoleMsg(tUser, STANDARD_BOUNTY_HUNTER_MESSAGE, FONTTYPE_TALK)
                            End If
                        End If
                        
                        ' Log it
                        CommandString = CommandString & "AGREGAR "
                    
                    Case eEditOptions.eo_Vida
                        
                        If Val(Arg1) > MAX_VIDA_EDIT Then
                            Arg1 = CStr(MAX_VIDA_EDIT)
                            Call WriteConsoleMsg(UserIndex, "No puedes tener vida superior a " & MAX_VIDA_EDIT & ".", FONTTYPE_INFO)
                        End If
                        
                        ' No valido si esta offline, porque solo se puede editar a si mismo
                        UserList(tUser).Stats.MaxHp = Val(Arg1)
                        UserList(tUser).Stats.MinHp = Val(Arg1)
                        
                        Call WriteUpdateUserStats(tUser)
                        
                        ' Log it
                        CommandString = CommandString & "VIDA "
                        
                    Case eEditOptions.eo_Poss
                    
                        Dim Map As Integer
                        Dim X As Integer
                        Dim Y As Integer
                        
                        Map = Val(ReadField(1, Arg1, 45))
                        X = Val(ReadField(2, Arg1, 45))
                        Y = Val(ReadField(3, Arg1, 45))
                        
                        If InMapBounds(Map, X, Y) Then
                            
                            If tUser <= 0 Then
                                Dim sPos As String
                                sPos = Map & "-" & X & "-" & Y
                                
                                Call UpdateCharData("USER_INFO", "LAST_POS", UserId, sPos)

                                Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                            Else
                                Call WarpUserChar(tUser, Map, X, Y, True, True)
                                Call WriteConsoleMsg(UserIndex, "Usuario teletransportado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                            End If
                        Else
                            Call WriteConsoleMsg(UserIndex, "Posición inválida", FONTTYPE_INFO)
                        End If
                        
                        ' Log it
                        CommandString = CommandString & "POSS "
                    
                    Case eEditOptions.eo_PlayerPoints
                    
                        If Val(Arg1) > MAX_PP_EDIT Then
                            Arg1 = CStr(MAX_PP_EDIT)
                            Call WriteConsoleMsg(UserIndex, "No está permitido utilizar valores mayores a " & MAX_PP_EDIT & ".", FONTTYPE_INFO)
                        
                        ElseIf Val(Arg1) < MIN_PP_EDIT Then
                            Arg1 = CStr(MIN_PP_EDIT)
                            Call WriteConsoleMsg(UserIndex, "No está permitido utilizar valores menores a " & MIN_PP_EDIT & ".", FONTTYPE_INFO)
                        End If
                        
                        If tUser <= 0 Then
                            Call UpdateCharData("USER_STATS", "RANKING_POINTS", UserId, Arg1)

                            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName & ". Puntos: " & Arg1 & ".", FONTTYPE_TALK)
                        Else
                            UserList(tUser).Stats.RankingPoints = CLng(Arg1)
                        End If
                        
                        ' Log it
                        CommandString = CommandString & "PP "
                        
                    Case Else
                        Call WriteConsoleMsg(UserIndex, "Comando no permitido.", FontTypeNames.FONTTYPE_INFO)
                        CommandString = CommandString & "UNKOWN "
                        
                End Select
                
                CommandString = CommandString & Arg1 & " " & Arg2
                Call LogGM(.Name, CommandString & " " & UserName)
                
            End If
        End If

    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleEditUser de Protocol.bas")
End Sub

Private Sub HandleRequestStatsBosses(ByVal UserIndex As Integer)
On Error GoTo ErrHandler

    If UserList(UserIndex).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
        Dim I As Integer
        Dim J As Integer
        Dim BossInfoString As String
        
        For I = 1 To UBound(BossData)
            With BossData(I)
                BossInfoString = "Boss Info: " & NpcData(BossData(I).NpcIndex).Name & IIf(BossData(I).Alive, " (Activo)", " (Inactivo)") & vbCrLf & _
                                "-   NPCs: " & CStr(.CurAmount) & "/" & CStr(.Amount)
                                
                Call WriteConsoleMsg(UserIndex, BossInfoString, FONTTYPE_INFO)
            End With
        Next I
    End If
    
    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleRequestStatsBosses de Protocol.bas")
End Sub

''
' Handles the "RequestCharInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharInfo(ByVal UserIndex As Integer)
'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 01/08/07
'Last Modification by: (liquid).. alto bug zapallo..
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)
    
        Dim TargetName As String
        Dim TargetIndex As Integer
        
        TargetName = Reader.ReadString8()

        TargetName = Replace$(TargetName, "+", " ")
        TargetIndex = NameIndex(TargetName)
        
        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios) Then
            'is the player offline?
            If TargetIndex <= 0 Then
                'don't allow to retrieve administrator's info
                If Not (EsDios(TargetName) Or EsAdmin(TargetName)) Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline, buscando en charfile.", FontTypeNames.FONTTYPE_INFO)
                    
                    Call SendUserStatsDB(UserIndex, TargetName)
                End If
            Else
                'don't allow to retrieve administrator's info
                If UserList(TargetIndex).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then
                    Call SendUserStatsTxt(UserIndex, TargetIndex)
                End If
            End If
        End If
    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleRequestCharInfo de Protocol.bas")
End Sub

''
' Handles the "RequestCharStats" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharStats(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 07/06/2010
'07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim UserName As String
        Dim tUser As Integer
        
        Dim UserIsAdmin As Boolean
        Dim OtherUserIsAdmin As Boolean

        UserName = Reader.ReadString8()

        UserIsAdmin = (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And ((.flags.Privilegios And PlayerType.SemiDios) <> 0 Or UserIsAdmin) Then
            Call LogGM(.Name, "/STAT " & UserName)
            
            tUser = NameIndex(UserName)
            
            OtherUserIsAdmin = EsDios(UserName) Or EsAdmin(UserName)
            
            If tUser <= 0 Then
                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    Call SendUserMiniStatsDB(UserIndex, UserName)
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes ver los stats de un dios o admin.", FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    Call SendUserMiniStatsTxt(UserIndex, tUser)
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes ver los stats de un dios o admin.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleRequestCharStats de Protocol.bas")
End Sub

''
' Handles the "RequestCharGold" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharGold(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 07/06/2010
'07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim UserName As String
        Dim tUser As Integer
        
        Dim UserIsAdmin As Boolean
        Dim OtherUserIsAdmin As Boolean

        UserName = Reader.ReadString8()
        
        UserIsAdmin = (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
        
        If (.flags.Privilegios And PlayerType.SemiDios) Or UserIsAdmin Then
            
            Call LogGM(.Name, "/BAL " & UserName)
            
            tUser = NameIndex(UserName)
            OtherUserIsAdmin = EsDios(UserName) Or EsAdmin(UserName)
            
            tUser = NameIndex(UserName)
            OtherUserIsAdmin = EsDios(UserName) Or EsAdmin(UserName)
            
            If tUser <= 0 Then
                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline. Leyendo charfile... ", FontTypeNames.FONTTYPE_TALK)
                    
                    Call SendUserOROTxtFromChar(UserIndex, UserName)
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes ver el oro de un dios o admin.", FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    Call WriteConsoleMsg(UserIndex, "El usuario " & UserName & " tiene " & UserList(tUser).Stats.Banco & " en el banco.", FontTypeNames.FONTTYPE_TALK)
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes ver el oro de un dios o admin.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleRequestCharGold de Protocol.bas")
End Sub

''
' Handles the "RequestCharInventory" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharInventory(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 07/06/2010
'07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim UserName As String
        Dim tUser As Integer
        
        Dim UserIsAdmin As Boolean
        Dim OtherUserIsAdmin As Boolean
        
        UserName = Reader.ReadString8()

        UserIsAdmin = (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call LogGM(.Name, "/INV " & UserName)
            
            tUser = NameIndex(UserName)
            OtherUserIsAdmin = EsDios(UserName) Or EsAdmin(UserName)
            
            tUser = NameIndex(UserName)
            OtherUserIsAdmin = EsDios(UserName) Or EsAdmin(UserName)
            
            If tUser <= 0 Then
                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline. Leyendo del charfile...", FontTypeNames.FONTTYPE_TALK)
                    
                    Call SendUserInvTxtFromChar(UserIndex, UserName)
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes ver el inventario de un dios o admin.", FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    Call SendUserInvTxt(UserIndex, tUser)
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes ver el inventario de un dios o admin.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleRequestCharInventory de Protocol.bas")
End Sub

''
' Handles the "RequestCharBank" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharBank(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 07/06/2010
'07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim UserName As String
        Dim tUser As Integer
        
        Dim UserIsAdmin As Boolean
        Dim OtherUserIsAdmin As Boolean

        UserName = Reader.ReadString8()

        UserIsAdmin = (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
        
        If (.flags.Privilegios And PlayerType.SemiDios) <> 0 Or UserIsAdmin Then
            Call LogGM(.Name, "/BOV " & UserName)
            
            tUser = NameIndex(UserName)
            OtherUserIsAdmin = EsDios(UserName) Or EsAdmin(UserName)
            
            tUser = NameIndex(UserName)
            OtherUserIsAdmin = EsDios(UserName) Or EsAdmin(UserName)
            
            If tUser <= 0 Then
                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline. Leyendo charfile... ", FontTypeNames.FONTTYPE_TALK)
                
                    Call SendUserBovedaTxtFromChar(UserIndex, UserName)
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes ver la bóveda de un dios o admin.", FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    Call SendUserBovedaTxt(UserIndex, tUser)
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes ver la bóveda de un dios o admin.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleRequestCharBank de Protocol.bas")
End Sub

''
' Handles the "RequestCharSkills" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharSkills(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim UserName As String
        Dim tUser As Integer
        Dim LoopC As Long
        Dim Message As String
        
        UserName = Reader.ReadString8()

        tUser = NameIndex(UserName)
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call LogGM(.Name, "/STATS " & UserName)
            
            If tUser <= 0 Then
                If (InStrB(UserName, "\") <> 0) Then
                        UserName = Replace(UserName, "\", "")
                End If
                If (InStrB(UserName, "/") <> 0) Then
                        UserName = Replace(UserName, "/", "")
                End If

                Call SendUserSkillsDB(UserIndex, UserName)
            Else
                Call SendUserSkillsTxt(UserIndex, tUser)
            End If
        End If
    
    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleRequestCharSkills de Protocol.bas")
End Sub

''
' Handles the "ReviveChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleReviveChar(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 11/03/2010
'11/03/2010: ZaMa - Al revivir con el comando, si esta navegando le da cuerpo e barca.
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim UserName As String
        Dim tUser As Integer
        
        UserName = Reader.ReadString8()

        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            If UCase$(UserName) <> "YO" Then
                tUser = NameIndex(UserName)
            Else
                tUser = UserIndex
            End If
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else
                With UserList(tUser)
                    'If dead, show him alive (naked).
                    If .flags.Muerto = 1 Then
                        .flags.Muerto = 0
                        
                        If .flags.Navegando = 1 Then
                            Call ToggleBoatBody(tUser)
                        Else
                            Call DarCuerpoDesnudo(tUser)
                        End If
                        
                        If .flags.Traveling = 1 Then
                            Call EndTravel(tUser, True)
                        End If
                        
                        Call ChangeUserChar(tUser, .Char.body, .OrigChar.head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                        
                        Call WriteConsoleMsg(tUser, UserList(UserIndex).Name & " te ha resucitado.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call WriteConsoleMsg(tUser, UserList(UserIndex).Name & " te ha curado.", FontTypeNames.FONTTYPE_INFO)
                    End If
                    
                    .Stats.MinHp = .Stats.MaxHp
                    .Stats.MinMAN = .Stats.MaxMan
                    .Stats.MinSta = .Stats.MaxSta
                    
                    If .flags.Traveling = 1 Then
                        Call EndTravel(tUser, True)
                    End If
                    
                End With
                
                Call WriteUpdateUserStats(tUser)

                Call LogGM(.Name, "Resucito a " & UserName)
            End If
        End If
        
    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleReviveChar de Protocol.bas")
End Sub

''
' Handles the "OnlineGM" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineGM(ByVal UserIndex As Integer)
'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 12/28/06
'
'***************************************************
On Error GoTo ErrHandler
  
    Dim I As Long
    Dim list As String
    Dim Priv As PlayerType
    Dim isRM As Boolean
    
    With UserList(UserIndex)

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub

        Priv = PlayerType.Consejero Or PlayerType.SemiDios
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then Priv = Priv Or PlayerType.Dios Or PlayerType.Admin
        
        isRM = ((.flags.Privilegios And PlayerType.RoleMaster) <> 0)
         
        For I = 1 To LastUser
            If UserList(I).flags.UserLogged Then
                If ((UserList(I).flags.Privilegios And Priv) <> 0) Then
                    If Not (isRM And (((UserList(I).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0)) And (UserList(I).flags.Privilegios And PlayerType.RoleMaster) = 0) Then
                        list = list & UserList(I).Name & ", "
                    End If
                End If
            End If
        Next I
        
        If LenB(list) <> 0 Then
            list = Left$(list, Len(list) - 2)
            Call WriteConsoleMsg(UserIndex, list & ".", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, "No hay GMs Online.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleOnlineGM de Protocol.bas")
End Sub

''
' Handles the "OnlineMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineMap(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 23/03/2009
'23/03/2009: ZaMa - Ahora no requiere estar en el mapa, sino que por defecto se toma en el que esta, pero se puede especificar otro
'***************************************************
On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim Map As Integer
        Map = Reader.ReadInt16
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub
        
        Dim LoopC As Long
        Dim list As String
        Dim Priv As PlayerType
        
        Priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then Priv = Priv + (PlayerType.Dios Or PlayerType.Admin)
        
        For LoopC = 1 To LastUser
            If LenB(UserList(LoopC).Name) <> 0 And UserList(LoopC).Pos.Map = Map Then
                If UserList(LoopC).flags.Privilegios And Priv Then _
                    list = list & UserList(LoopC).Name & ", "
            End If
        Next LoopC
        
        If Len(list) > 2 Then list = Left$(list, Len(list) - 2)
        
        Call WriteConsoleMsg(UserIndex, "Usuarios en el mapa: " & list, FontTypeNames.FONTTYPE_INFO)
        Call LogGM(.Name, "/ONLINEMAP " & Map)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleOnlineMap de Protocol.bas")
End Sub


''
' Handles the "Kick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKick(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 07/06/2010
'07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim UserName As String
        Dim tUser As Integer
        Dim rank As Integer
        Dim IsAdmin As Boolean
        
        UserName = Reader.ReadString8()

        rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
        

        IsAdmin = (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
        
        If (.flags.Privilegios And PlayerType.SemiDios) Or IsAdmin Then
            tUser = NameIndex(UserName)
            
            If tUser <= 0 Then
                If Not (EsDios(UserName) Or EsAdmin(UserName)) Or IsAdmin Then
                    Call WriteConsoleMsg(UserIndex, "El usuario no está online.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes echar a alguien con jerarquía mayor a la tuya.", FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                If (UserList(tUser).flags.Privilegios And rank) > (.flags.Privilegios And rank) Then
                    Call WriteConsoleMsg(UserIndex, "No puedes echar a alguien con jerarquía mayor a la tuya.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call CloseSocket(tUser)
                    Call LogGM(.Name, "Echó a " & UserName)
                End If
            End If
        End If
    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleKick de Protocol.bas")
End Sub

''
' Handles the "Execute" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleExecute(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 07/06/2010
'07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim UserName As String
        Dim tUser As Integer
        
        UserName = Reader.ReadString8()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            tUser = NameIndex(UserName)
            
            If tUser > 0 Then
                If Not UserList(tUser).flags.Privilegios And PlayerType.User Then
                    Call WriteConsoleMsg(UserIndex, "¿¿Estás loco?? ¿¿Cómo vas a piñatear un gm?? :@", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call UserDie(tUser)
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & " ha ejecutado a " & UserName & ".", FontTypeNames.FONTTYPE_EJECUCION))
                    Call LogGM(.Name, " ejecuto a " & UserName)
                End If
            Else
                If Not (EsDios(UserName) Or EsAdmin(UserName)) Then
                    Call WriteConsoleMsg(UserIndex, "No está online.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, "¿¿Estás loco?? ¿¿Cómo vas a piñatear un gm?? :@", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleExecute de Protocol.bas")
End Sub

''
' Handles the "BanChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBanChar(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim UserName As String
        Dim Reason As String
        Dim AdminNotes As String
        Dim PunishmentTypeId As Integer
        
        UserName = Reader.ReadString8()
        Reason = Reader.ReadString8()
        AdminNotes = Reader.ReadString8()
        PunishmentTypeId = Reader.ReadInt16()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            Call BanCharacter(UserIndex, UserName, Reason, AdminNotes, PunishmentTypeId)
        End If
    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleBanChar de Protocol.bas")
End Sub

''
' Handles the "UnbanChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUnbanChar(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim UserName As String
        Dim cantPenas As Byte
        
        UserName = Reader.ReadString8()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "\", "")
            End If
            If (InStrB(UserName, "/") <> 0) Then
                UserName = Replace(UserName, "/", "")
            End If
            
            Dim UserId As Long
            Dim bUserBanned As Boolean
            Call GetCharInfo(UserName, UserId, bUserBanned)
            
            If UserId = 0 Then
                Call WriteConsoleMsg(UserIndex, "Charfile inexistente (no use +).", FontTypeNames.FONTTYPE_INFO)
            Else

                If bUserBanned Then
                    Call UnbanCharacterDB(UserId, UserList(UserIndex).Id, PunishmentStaticIds.UnBanChar, "UNBAN manual")
                    
                    Call LogGM(.Name, "/UNBAN a " & UserName)
                    Call WriteConsoleMsg(UserIndex, UserName & " unbanned.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, UserName & " no está baneado. Imposible unbanear.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleUnbanChar de Protocol.bas")
End Sub

''
' Handles the "NPCFollow" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleNPCFollow(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub
        
        If .flags.TargetNpc > 0 Then
            Call DoFollow(.flags.TargetNpc, .Name)
            Npclist(.flags.TargetNpc).flags.Inmovilizado = 0
            Npclist(.flags.TargetNpc).flags.Paralizado = 0
            Npclist(.flags.TargetNpc).Contadores.Paralisis = 0
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleNPCFollow de Protocol.bas")
End Sub

''
' Handles the "SummonChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSummonChar(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 26/03/2009
'26/03/2009: ZaMa - Chequeo que no se teletransporte donde haya un char o npc
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim UserName As String
        Dim tUser As Integer
        Dim X As Integer
        Dim Y As Integer
        Dim I As Long
        Dim names() As String
        UserName = Reader.ReadString8()
 
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            names = Split(UserName, ",")
            For I = LBound(names) To UBound(names)
                tUser = NameIndex(names(I))
                
                If tUser <= 0 Then
                    If EsDios(names(I)) Or EsAdmin(names(I)) Then
                        Call WriteConsoleMsg(UserIndex, "Imposible invocar a " & names(I) & ". No puedes invocar a dioses y admins.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call WriteConsoleMsg(UserIndex, "El jugador " & names(I) & " no está online.", FontTypeNames.FONTTYPE_INFO)
                    End If
                    
                Else
                    If (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Or _
                      (UserList(tUser).flags.Privilegios And (PlayerType.Consejero Or PlayerType.User)) <> 0 Then
                        Call WriteConsoleMsg(tUser, .Name & " te ha trasportado.", FontTypeNames.FONTTYPE_INFO)
                        X = .Pos.X
                        Y = .Pos.Y + 1
                        Call FindLegalPos(tUser, .Pos.Map, X, Y)
                        Call WarpUserChar(tUser, .Pos.Map, X, Y, True, True, True)
                        Call LogGM(.Name, "/SUM " & names(I) & " Map:" & .Pos.Map & " X:" & .Pos.X & " Y:" & .Pos.Y)
                    Else
                        Call WriteConsoleMsg(UserIndex, "Imposible invocar a " & names(I) & ". No puedes invocar a dioses y admins.", FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
            Next I
        End If
    
    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleSummonChar de Protocol.bas")
End Sub

''
' Handles the "SpawnListRequest" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSpawnListRequest(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub
        
        Call EnviarSpawnList(UserIndex)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleSpawnListRequest de Protocol.bas")
End Sub

''
' Handles the "SpawnCreature" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSpawnCreature(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim npc As Integer
        npc = Reader.ReadInt16()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            If npc > 0 And npc <= UBound(Declaraciones.SpawnList()) Then _
              Call SpawnNpc(Declaraciones.SpawnList(npc).NpcIndex, .Pos, True, False)
            
            Call LogGM(.Name, "Sumoneo " & Declaraciones.SpawnList(npc).NpcName)
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleSpawnCreature de Protocol.bas")
End Sub

''
' Handles the "ResetNPCInventory" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleResetNPCInventory(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub
        If .flags.TargetNpc = 0 Then Exit Sub
        
        Call ResetNpcInv(.flags.TargetNpc)
        Call LogGM(.Name, "/RESETINV " & Npclist(.flags.TargetNpc).Name)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleResetNPCInventory de Protocol.bas")
End Sub

''
' Handles the "CleanWorld" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCleanWorld(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster Or PlayerType.ChaosCouncil Or PlayerType.SemiDios Or PlayerType.RoyalCouncil) Then Exit Sub
        
        Call LimpiarMundo
        
        Call SecurityIp.IpTableSecurityCleanIpCount
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleCleanWorld de Protocol.bas")
End Sub

''
' Handles the "ServerMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleServerMessage(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 28/05/2010
'28/05/2010: ZaMa - Ahora no dice el nombre del gm que lo dice.
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim Message As String
        Message = Reader.ReadString8()

        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            If LenB(Message) <> 0 Then
                Call LogGM(.Name, "Mensaje Broadcast:" & Message)
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(Message, FontTypeNames.FONTTYPE_TALK, eMessageType.Admin))
            End If
        End If
    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleServerMessage de Protocol.bas")
End Sub

''
' Handles the "MapMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMapMessage(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 14/11/2010
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim Message As String
        Message = Reader.ReadString8()

        If ((.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.RoleMaster)) <> 0) Then
            If LenB(Message) <> 0 Then
                
                Dim mapa As Integer
                mapa = .Pos.Map
                
                Call LogGM(.Name, "Mensaje a mapa " & mapa & ":" & Message)
                Call SendData(SendTarget.toMap, mapa, PrepareMessageConsoleMsg(Message, FontTypeNames.FONTTYPE_TALK, eMessageType.Admin))
            End If
        End If
    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleMapMessage de Protocol.bas")
End Sub

''
' Handles the "NickToIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleNickToIP(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 07/06/2010
'Pablo (ToxicWaste): Agrego para que el /nick2ip tambien diga los nicks en esa ip por pedido de la DGM.
'07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim UserName As String
        Dim tUser As Integer
        Dim Priv As PlayerType
        Dim IsAdmin As Boolean
        
        UserName = Reader.ReadString8()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            tUser = NameIndex(UserName)
            Call LogGM(.Name, "NICK2IP Solicito la IP de " & UserName)
            
            IsAdmin = (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0
            If IsAdmin Then
                Priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin
            Else
                Priv = PlayerType.User
            End If
            
            If tUser > 0 Then
                If UserList(tUser).flags.Privilegios And Priv Then
                    Call WriteConsoleMsg(UserIndex, "El ip de " & UserName & " es " & UserList(tUser).IP, FontTypeNames.FONTTYPE_INFO)
                    Dim IP As String
                    Dim Lista As String
                    Dim LoopC As Long
                    IP = UserList(tUser).IP
                    For LoopC = 1 To LastUser
                        If UserList(LoopC).IP = IP Then
                            If LenB(UserList(LoopC).Name) <> 0 And UserList(LoopC).flags.UserLogged Then
                                If UserList(LoopC).flags.Privilegios And Priv Then
                                    Lista = Lista & UserList(LoopC).Name & ", "
                                End If
                            End If
                        End If
                    Next LoopC
                    If LenB(Lista) <> 0 Then Lista = Left$(Lista, Len(Lista) - 2)
                    Call WriteConsoleMsg(UserIndex, "Los personajes con ip " & IP & " son: " & Lista, FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                If Not (EsDios(UserName) Or EsAdmin(UserName)) Or IsAdmin Then
                    Call WriteConsoleMsg(UserIndex, "No hay ningún personaje con ese nick.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleNickToIP de Protocol.bas")
End Sub

''
' Handles the "IPToNick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleIPToNick(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
On Error GoTo ErrHandler

    With UserList(UserIndex)
        Dim IP As String
        Dim LoopC As Long
        Dim Lista As String
        Dim Priv As PlayerType
        
        IP = Reader.ReadInt8() & "."
        IP = IP & Reader.ReadInt8() & "."
        IP = IP & Reader.ReadInt8() & "."
        IP = IP & Reader.ReadInt8()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, "IP2NICK Solicito los Nicks de IP " & IP)
        
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
            Priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin
        Else
            Priv = PlayerType.User
        End If

        For LoopC = 1 To LastUser
            If UserList(LoopC).IP = IP Then
                If LenB(UserList(LoopC).Name) <> 0 And UserList(LoopC).flags.UserLogged Then
                    If UserList(LoopC).flags.Privilegios And Priv Then
                        Lista = Lista & UserList(LoopC).Name & ", "
                    End If
                End If
            End If
        Next LoopC
        
        If LenB(Lista) <> 0 Then Lista = Left$(Lista, Len(Lista) - 2)
        Call WriteConsoleMsg(UserIndex, "Los personajes con ip " & IP & " son: " & Lista, FontTypeNames.FONTTYPE_INFO)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleIPToNick de Protocol.bas")
End Sub

Private Sub HandleAdminGuildMembers(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    
    Dim GuildName As String
    Dim GuildIndex As Long
    
    GuildName = Reader.ReadString8()
    
     ' Only Admins are allowed to use this command.
    If (UserList(UserIndex).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) = 0 Then Exit Sub
    
    If LenB(GuildName) <= 0 Then
        Call WriteConsoleMsg(UserIndex, "El nombre del clan no puede estar vacío.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If Not modGuild_Functions.TryGetGuildIndexByName(GuildName, GuildIndex) Then
        Call WriteConsoleMsg(UserIndex, "El clan seleccionado no existe.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    Dim I As Integer
    Dim ListOfUsers As String
    Dim OnlineUserName As String
    
    For I = 1 To GuildList(GuildIndex).MemberCount
        OnlineUserName = GuildList(GuildIndex).Members(I).NameUser
        
        If LenB(ListOfUsers) > 0 Then
            ListOfUsers = ListOfUsers & "," & OnlineUserName
        Else
            ListOfUsers = OnlineUserName
        End If
        
    Next I
    
    Call WriteConsoleMsg(UserIndex, "Los usuarios del clan " & GuildList(GuildIndex).Name & " son: " & ListOfUsers, FontTypeNames.FONTTYPE_INFO)
    
    Exit Sub
  
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleAdminGuildMembers de Protocol.bas")
End Sub

Private Sub HandleRemoveCharFromGuild(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    
    Dim GuildId As Long
    Dim GuildName As String
    Dim GuildIndex As Long
    Dim UserName As String
    Dim UserId As Long
    Dim TargetUserIndex As Integer

    UserName = Reader.ReadString8()
        
     ' Only Admins are allowed to use this command.
    If (UserList(UserIndex).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) = 0 Then Exit Sub
    
    If LenB(UserName) <= 0 Then
        Call WriteConsoleMsg(UserIndex, "El nombre del usuario no puede estar vacío.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    TargetUserIndex = NameIndex(UserName)
    
    If TargetUserIndex > 0 Then
        UserId = UserList(TargetUserIndex).Id
        GuildId = UserList(TargetUserIndex).Guild.IdGuild
        GuildIndex = UserList(TargetUserIndex).Guild.GuildIndex
        GuildName = GuildList(UserList(TargetUserIndex).Guild.GuildIndex).Name
        
    Else
        Call GetGuildInformationFromUserName(UserName, UserId, GuildId, GuildName)
        
        If UserId <= 0 Then
            Call WriteConsoleMsg(UserIndex, "El usuario no existe.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If GuildId <= 0 Then
            Call WriteConsoleMsg(UserIndex, "El usuario no se encuentra en un clan.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Not modGuild_Functions.TryGetGuildIndexByGuildId(GuildId, GuildIndex) Then
            Call WriteConsoleMsg(UserIndex, "El clan seleccionado no existe.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
    End If
    
    Call modGuild_Functions.GuildRemoveMember(GuildIndex, UserId, NotifyUser:=True)
    
    Call WriteConsoleMsg(UserIndex, "El usuario " & UserName & " fue expulsado del clan " & GuildName, FontTypeNames.FONTTYPE_INFO)
    
    If TargetUserIndex > 0 Then
        Call WriteConsoleMsg(TargetUserIndex, "Fuiste expulsado del clan " & GuildName & " por " & UserList(UserIndex).Name, FontTypeNames.FONTTYPE_INFO)
    End If
    
    
    Exit Sub
  
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleAdminGuildOnlineMembers de Protocol.bas")
End Sub

Private Sub HandleModGuildContribution(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    
    Dim GuildName As String
    Dim Quantity As String
    Dim GuildIndex As Long

    GuildName = Reader.ReadString8()
    Quantity = Reader.ReadInt32()
        
     ' Only Admins are allowed to use this command.
    If (UserList(UserIndex).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) = 0 Then Exit Sub
    
    If LenB(GuildName) <= 0 Then
        Call WriteConsoleMsg(UserIndex, "El nombre del usuario no puede estar vacío.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If Not modGuild_Functions.TryGetGuildIndexByName(GuildName, GuildIndex) Then
        Call WriteConsoleMsg(UserIndex, "El clan seleccionado no existe.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    Dim PreviousPoints As Long
    Dim QuantityTemp As Long
    
    ' Add/Substract the contribution from the guild.
    With GuildList(GuildIndex)
        
        PreviousPoints = .ContributionAvailable
        QuantityTemp = Quantity
    
        ' Just checking for underflows/overflows as this could be a destructive operation
        If (.ContributionAvailable + Quantity) < 0 Then
            QuantityTemp = .ContributionAvailable
        ElseIf (.ContributionAvailable + Quantity) > 2147483647 Then
            QuantityTemp = 2147483647 - .ContributionAvailable
        Else
            QuantityTemp = Quantity
        End If
        
        .ContributionAvailable = .ContributionAvailable + QuantityTemp
        
        Quantity = QuantityTemp
        
        If (.ContributionEarned + Quantity) < 0 Then
            QuantityTemp = .ContributionEarned
        ElseIf (.ContributionEarned + Quantity) > 2147483647 Then
            QuantityTemp = 2147483647 - .ContributionEarned
        Else
            QuantityTemp = Quantity
        End If
        
        .ContributionEarned = .ContributionEarned + QuantityTemp
        
        Call modGuild_DB.UpdateGuildStats(GuildIndex)
        Call WriteGuildInfoChange(UserIndex, eChangeGuildInfo.ContributionAvailableChange, 0, .ContributionAvailable)
        Call LogGM(UserList(UserIndex).Name, "MODCLANCONTRI Modificó la contribución del clan " & GuildName & " en " & Quantity & " puntos. Antes: " & PreviousPoints & ", Despues: " & .ContributionAvailable)
        Call WriteConsoleMsg(UserIndex, "Modificaste los puntos de contribución del clan " & GuildName & " en " & Quantity & " puntos. El clan ahora tiene " & .ContributionAvailable, FontTypeNames.FONTTYPE_INFO)
        
    End With
    
    Exit Sub
  
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleAdminGuildOnlineMembers de Protocol.bas")
End Sub


Private Sub HandleAdminGuildOnlineMembers(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    
    Dim GuildName As String
    Dim GuildIndex As Long
    
    GuildName = Reader.ReadString8()
    
     ' Only Admins are allowed to use this command.
    If (UserList(UserIndex).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) = 0 Then Exit Sub
    
    If LenB(GuildName) <= 0 Then
        Call WriteConsoleMsg(UserIndex, "El nombre del clan no puede estar vacío.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If Not modGuild_Functions.TryGetGuildIndexByName(GuildName, GuildIndex) Then
        Call WriteConsoleMsg(UserIndex, "El clan seleccionado no existe.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    Dim I As Integer
    Dim ListOfUsers As String
    Dim OnlineUserName As String
    
    For I = 1 To GuildList(GuildIndex).OnlineMemberCount
        OnlineUserName = UserList(GuildList(GuildIndex).OnlineMembers(I).MemberUserIndex).Name
        
        If LenB(ListOfUsers) > 0 Then
            ListOfUsers = ListOfUsers & "," & OnlineUserName
        Else
            ListOfUsers = OnlineUserName
        End If
        
    Next I
    
    Call WriteConsoleMsg(UserIndex, "Los usuarios online del clan " & GuildList(GuildIndex).Name & " son: " & ListOfUsers, FontTypeNames.FONTTYPE_INFO)
    
    Exit Sub
  
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleAdminGuildOnlineMembers de Protocol.bas")
End Sub

''
' Handles the "AdminChangeGuildAlign" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAdminChangeGuildAlign(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler
    
    With UserList(UserIndex)

        Dim GuildName As String
        Dim NewAlignment As Integer
        Dim OldAlignment As Integer
        
        GuildName = Reader.ReadString8()
        NewAlignment = Reader.ReadInt8()

        ' Only Admins are allowed to use this command.
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) = 0 Then Exit Sub
        
        'If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
        
        Dim GuildIndex As Integer
        GuildIndex = GetGuildIndex(GuildName)
        
        If GuildIndex <= 0 Then
            Call WriteConsoleMsg(UserIndex, "El clan seleccionado no existe.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If NewAlignment = eGuildAlignment.GameMaster Then
            Call WriteConsoleMsg(UserIndex, "No se puede setear la alineación GameMaster a un clan utilizando este comando.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
                
        If NewAlignment < 1 Or NewAlignment >= eGuildAlignment.LastElement Then
            Call WriteConsoleMsg(UserIndex, "La alineación no es un valor válido. Los valores posibles son: Neutral=" & eGuildAlignment.Neutral _
                & ", Caos=" & eGuildAlignment.Evil & ", Real=" & eGuildAlignment.Real, FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If GuildList(GuildIndex).Alignment = NewAlignment Then
            Call WriteConsoleMsg(UserIndex, "El clan ya posee esa alineación.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        OldAlignment = GuildList(GuildIndex).Alignment
        GuildList(GuildIndex).Alignment = NewAlignment
        
        Call modGuild_DB.UpdateGuildAlignment(GuildList(GuildIndex).IdGuild, NewAlignment)
        
        Call LogGM(UserList(UserIndex).Name, "Cambiada alineación del clan " & GuildList(GuildIndex).Name & " de " & OldAlignment & " a " & NewAlignment)
        
        Call WriteConsoleMsg(UserIndex, "La alineación del clan " & GuildList(GuildIndex).Name & " fue cambiada correctamente.", FontTypeNames.FONTTYPE_INFO)
       
    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleAdminChangeGuildAlign de Protocol.bas")
End Sub

''
' Handles the "TeleportCreate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTeleportCreate(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 22/03/2010
'15/11/2009: ZaMa - Ahora se crea un teleport con un radio especificado.
'22/03/2010: ZaMa - Harcodeo los teleps y radios en el dat, para evitar mapas bugueados.
'***************************************************
On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim mapa As Integer
        Dim X As Byte
        Dim Y As Byte
        Dim Radio As Byte
        
        mapa = Reader.ReadInt16()
        X = Reader.ReadInt8()
        Y = Reader.ReadInt8()
        Radio = Reader.ReadInt8()
        
        Radio = MinimoInt(Radio, 6)
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Call LogGM(.Name, "/CT " & mapa & "," & X & "," & Y & "," & Radio)
        
        If Not MapaValido(mapa) Or Not InMapBounds(mapa, X, Y) Then _
            Exit Sub
        
        If MapData(.Pos.Map, .Pos.X, .Pos.Y - 1).ObjInfo.ObjIndex > 0 Then _
            Exit Sub
        
        If MapData(.Pos.Map, .Pos.X, .Pos.Y - 1).TileExit.Map > 0 Then _
            Exit Sub
        
        If MapData(mapa, X, Y).ObjInfo.ObjIndex > 0 Then
            Call WriteConsoleMsg(UserIndex, "Hay un objeto en el piso en ese lugar.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If MapData(mapa, X, Y).TileExit.Map > 0 Then
            Call WriteConsoleMsg(UserIndex, "No puedes crear un teleport que apunte a la entrada de otro.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Dim ET As Obj
        ET.Amount = 1
        ' Es el numero en el dat. El indice es el comienzo + el radio, todo harcodeado :(.
        ET.ObjIndex = ConstantesItems.Telep + Radio
        
        ET.CurrentGrhIndex = ObjData(ET.ObjIndex).GrhIndex
        
        With MapData(.Pos.Map, .Pos.X, .Pos.Y - 1)
            
            .TileExit.Map = mapa
            .TileExit.X = X
            .TileExit.Y = Y
        End With
        
        Call MakeObj(ET, .Pos.Map, .Pos.X, .Pos.Y - 1)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleTeleportCreate de Protocol.bas")
End Sub

''
' Handles the "TeleportDestroy" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTeleportDestroy(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)
        Dim mapa As Integer
        Dim X As Byte
        Dim Y As Byte

        '/dt
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        mapa = .flags.TargetMap
        X = .flags.TargetX
        Y = .flags.TargetY
        
        If Not InMapBounds(mapa, X, Y) Then Exit Sub
        
        With MapData(mapa, X, Y)
            If .ObjInfo.ObjIndex = 0 Then Exit Sub
            
            If ObjData(.ObjInfo.ObjIndex).ObjType = eOBJType.otTeleport And .TileExit.Map > 0 Then
                Call LogGM(UserList(UserIndex).Name, "/DT: " & mapa & "," & X & "," & Y)
                
                Call EraseObj(.ObjInfo.Amount, mapa, X, Y)
                
                If MapData(.TileExit.Map, .TileExit.X, .TileExit.Y).ObjInfo.ObjIndex = 651 Then
                    Call EraseObj(1, .TileExit.Map, .TileExit.X, .TileExit.Y)
                End If
                
                .TileExit.Map = 0
                .TileExit.X = 0
                .TileExit.Y = 0
            End If
        End With
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleTeleportDestroy de Protocol.bas")
End Sub

''
' Handles the "RainToggle" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRainToggle(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub
        
        Call LogGM(.Name, "/LLUVIA")
        Lloviendo = Not Lloviendo
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessageRainToggle())
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleRainToggle de Protocol.bas")
End Sub

''
' Handles the "EnableDenounces" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleEnableDenounces(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 14/11/2010
'Enables/Disables
'***************************************************
On Error GoTo ErrHandler
  

    With UserList(UserIndex)

        ' Gm?
        If Not EsGm(UserIndex) Then Exit Sub
        ' Rm?
        If (.flags.Privilegios And PlayerType.RoleMaster) <> 0 Then Exit Sub

        Dim Activado As Boolean
        Dim msg As String
        
        Activado = Not .flags.SendDenounces
        .flags.SendDenounces = Activado
        
        msg = "Denuncias por consola " & IIf(Activado, "ativadas", "desactivadas") & "."
        
        Call LogGM(.Name, msg)
        
        Call WriteConsoleMsg(UserIndex, msg, FontTypeNames.FONTTYPE_INFO)
    End With

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleEnableDenounces de Protocol.bas")
End Sub

''
' Handles the "ShowDenouncesList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleShowDenouncesList(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 14/11/2010
'
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        If (.flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster)) <> 0 Then Exit Sub
        Call WriteShowDenounces(UserIndex)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleShowDenouncesList de Protocol.bas")
End Sub

''
' Handles the "SetCharDescription" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSetCharDescription(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim tUser As Integer
        Dim Desc As String
        
        Desc = Reader.ReadString8()

        If (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Or (.flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
            tUser = .flags.TargetUser
            If tUser > 0 Then
                UserList(tUser).DescRM = Desc
            Else
                Call WriteConsoleMsg(UserIndex, "Haz click sobre un personaje antes.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
    
    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleSetCharDescription de Protocol.bas")
End Sub

''
' Handles the "ForceMIDIToMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HanldeForceMIDIToMap(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim midiID As Byte
        Dim mapa As Integer
        
        midiID = Reader.ReadInt8
        mapa = Reader.ReadInt16
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            'Si el mapa no fue enviado tomo el actual
            If Not InMapBounds(mapa, 50, 50) Then
                mapa = .Pos.Map
            End If
        
            If midiID = 0 Then
                'Ponemos el default del mapa
                Call SendData(SendTarget.toMap, mapa, PrepareMessagePlayMusic(MapInfo(.Pos.Map).Music(RandomNumber(1, MapInfo(.Pos.Map).NumMusic))))
            Else
                'Ponemos el pedido por el GM
                Call SendData(SendTarget.toMap, mapa, PrepareMessagePlayMusic(midiID))
            End If
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HanldeForceMIDIToMap de Protocol.bas")
End Sub

''
' Handles the "ForceWAVEToMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForceWAVEToMap(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************

On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        Dim waveID As Byte
        Dim mapa As Integer
        Dim X As Byte
        Dim Y As Byte
        
        waveID = Reader.ReadInt8()
        mapa = Reader.ReadInt16()
        X = Reader.ReadInt8()
        Y = Reader.ReadInt8()
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
        'Si el mapa no fue enviado tomo el actual
            If Not InMapBounds(mapa, X, Y) Then
                mapa = .Pos.Map
                X = .Pos.X
                Y = .Pos.Y
            End If
            
            'Ponemos el pedido por el GM
            Call SendData(SendTarget.toMap, mapa, PrepareMessagePlayWave(waveID, X, Y))
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleForceWAVEToMap de Protocol.bas")
End Sub

''
' Handles the "RoyalArmyMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRoyalArmyMessage(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim Message As String
        Message = Reader.ReadString8()
        
        'Solo dioses, admins, semis y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then
            Call SendData(SendTarget.ToRealYRMs, 0, PrepareMessageConsoleMsg("EJÉRCITO REAL> " & Message, FontTypeNames.FONTTYPE_TALK))
        End If
    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleRoyalArmyMessage de Protocol.bas")
End Sub

''
' Handles the "ChaosLegionMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChaosLegionMessage(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim Message As String
        Message = Reader.ReadString8()

        'Solo dioses, admins, semis y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then
            Call SendData(SendTarget.ToCaosYRMs, 0, PrepareMessageConsoleMsg("FUERZAS DEL CAOS> " & Message, FontTypeNames.FONTTYPE_TALK))
        End If
        
    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleChaosLegionMessage de Protocol.bas")
End Sub

''
' Handles the "TalkAsNPC" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTalkAsNPC(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim Message As String
        Message = Reader.ReadString8()

        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            'Asegurarse haya un NPC seleccionado
            If .flags.TargetNpc > 0 Then
                Call SendData(SendTarget.ToNPCArea, .flags.TargetNpc, PrepareMessageChatOverHead(Message, Npclist(.flags.TargetNpc).Char.CharIndex, vbWhite))
            Else
                Call WriteConsoleMsg(UserIndex, "Debes seleccionar el NPC por el que quieres hablar antes de usar este comando.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleTalkAsNPC de Protocol.bas")
End Sub

''
' Handles the "DestroyAllItemsInArea" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDestroyAllItemsInArea(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Dim X As Long
        Dim Y As Long
        Dim bIsExit As Boolean
        
        For Y = .Pos.Y - MinYBorder + 1 To .Pos.Y + MinYBorder - 1
            For X = .Pos.X - MinXBorder + 1 To .Pos.X + MinXBorder - 1
                If X > 0 And Y > 0 And X < 101 And Y < 101 Then
                    If MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex > 0 Then
                        bIsExit = MapData(.Pos.Map, X, Y).TileExit.Map > 0
                        If ItemNoEsDeMapa(MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex, bIsExit) Then
                            Call EraseObj(MAX_INVENTORY_OBJS, .Pos.Map, X, Y)
                        End If
                    End If
                End If
            Next X
        Next Y
        
        Call LogGM(UserList(UserIndex).Name, "/MASSDEST")
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleDestroyAllItemsInArea de Protocol.bas")
End Sub

''
' Handles the "AcceptRoyalCouncilMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAcceptRoyalCouncilMember(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = Reader.ReadString8()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            tUser = NameIndex(UserName)
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline", FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue aceptado en el honorable Consejo Real de Banderbill.", FontTypeNames.FONTTYPE_CONSEJO))
                With UserList(tUser)
                    If .flags.Privilegios And PlayerType.ChaosCouncil Then .flags.Privilegios = .flags.Privilegios - PlayerType.ChaosCouncil
                    If Not .flags.Privilegios And PlayerType.RoyalCouncil Then .flags.Privilegios = .flags.Privilegios + PlayerType.RoyalCouncil
                    
                    Call WarpUserChar(tUser, .Pos.Map, .Pos.X, .Pos.Y, False)
                End With
            End If
        End If
    
    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleAcceptRoyalCouncilMember de Protocol.bas")
End Sub

''
' Handles the "ChaosCouncilMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAcceptChaosCouncilMember(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim UserName As String
        Dim tUser As Integer
        
        UserName = Reader.ReadString8()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            tUser = NameIndex(UserName)
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline", FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue aceptado en el Concilio de las Sombras.", FontTypeNames.FONTTYPE_CONSEJO))
                
                With UserList(tUser)
                    If .flags.Privilegios And PlayerType.RoyalCouncil Then .flags.Privilegios = .flags.Privilegios - PlayerType.RoyalCouncil
                    If Not .flags.Privilegios And PlayerType.ChaosCouncil Then .flags.Privilegios = .flags.Privilegios + PlayerType.ChaosCouncil

                    Call WarpUserChar(tUser, .Pos.Map, .Pos.X, .Pos.Y, False)
                End With
            End If
        End If
        
    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleAcceptChaosCouncilMember de Protocol.bas")
End Sub

''
' Handles the "ItemsInTheFloor" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleItemsInTheFloor(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub

        Dim tObj As Integer
        Dim X As Long
        Dim Y As Long
        
        For X = 5 To 95
            For Y = 5 To 95
                tObj = MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex
                If tObj > 0 Then
                    If ObjData(tObj).ObjType <> eOBJType.otResource Then
                        Call WriteConsoleMsg(UserIndex, "(" & X & "," & Y & ") " & ObjData(tObj).Name, FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
            Next Y
        Next X
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleItemsInTheFloor de Protocol.bas")
End Sub

''
' Handles the "MakeDumb" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMakeDumb(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim UserName As String
        Dim tUser As Integer
        
        UserName = Reader.ReadString8()

        If ((.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Or ((.flags.Privilegios And (PlayerType.SemiDios Or PlayerType.RoleMaster)) = (PlayerType.SemiDios Or PlayerType.RoleMaster))) Then
            tUser = NameIndex(UserName)
            'para deteccion de aoice
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteDumb(tUser)
            End If
        End If
    End With
    
    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleMakeDumb de Protocol.bas")
End Sub

''
' Handles the "MakeDumbNoMore" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMakeDumbNoMore(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim UserName As String
        Dim tUser As Integer
        
        UserName = Reader.ReadString8()

        If ((.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Or ((.flags.Privilegios And (PlayerType.SemiDios Or PlayerType.RoleMaster)) = (PlayerType.SemiDios Or PlayerType.RoleMaster))) Then
            tUser = NameIndex(UserName)
            'para deteccion de aoice
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteDumbNoMore(tUser)
            End If
        End If
    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleMakeDumbNoMore de Protocol.bas")
End Sub

''
' Handles the "DumpIPTables" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDumpIPTables(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Call SecurityIp.DumpTables
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleDumpIPTables de Protocol.bas")
End Sub

''
' Handles the "CouncilKick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCouncilKick(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim UserName As String
        Dim tUser As Integer
        
        UserName = Reader.ReadString8()
      
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            tUser = NameIndex(UserName)
            If tUser <= 0 Then

                Dim UserId As Long
                UserId = GetUserID(UserName)
                If UserId <> 0 Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline, echando de los consejos.", FontTypeNames.FONTTYPE_INFO)
                    
                    Call ExpellFromCouncilDB(UserId)
                Else
                    Call WriteConsoleMsg(UserIndex, "No se encuentra el charfile " & CharPath & UserName & ".chr", FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                With UserList(tUser)
                    If .flags.Privilegios And PlayerType.RoyalCouncil Then
                        Call WriteConsoleMsg(tUser, "Has sido echado del consejo de Banderbill.", FontTypeNames.FONTTYPE_TALK)
                        .flags.Privilegios = .flags.Privilegios - PlayerType.RoyalCouncil
                        
                        Call WarpUserChar(tUser, .Pos.Map, .Pos.X, .Pos.Y, False)
                        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue expulsado del consejo de Banderbill.", FontTypeNames.FONTTYPE_CONSEJO))
                    End If
                    
                    If .flags.Privilegios And PlayerType.ChaosCouncil Then
                        Call WriteConsoleMsg(tUser, "Has sido echado del Concilio de las Sombras.", FontTypeNames.FONTTYPE_TALK)
                        .flags.Privilegios = .flags.Privilegios - PlayerType.ChaosCouncil
                        
                        Call WarpUserChar(tUser, .Pos.Map, .Pos.X, .Pos.Y, False)
                        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue expulsado del Concilio de las Sombras.", FontTypeNames.FONTTYPE_CONSEJO))
                    End If
                End With
            End If
        End If

    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleCouncilKick de Protocol.bas")
End Sub

''
' Handles the "SetTrigger" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSetTrigger(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim tTrigger As Byte
        Dim tLog As String
        Dim ObjIndex As Integer
        
        tTrigger = Reader.ReadInt8()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        If tTrigger >= 0 Then
            If MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger = eTrigger.zonaOscura Then
                If tTrigger <> eTrigger.zonaOscura Then
                    If Not (.flags.AdminInvisible = 1) Then _
                        Call SendData(SendTarget.ToUsersAndRmsAndCounselorsAreaButGMs, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
                    
                    ObjIndex = MapData(.Pos.Map, .Pos.X, .Pos.Y).ObjInfo.ObjIndex
                    
                    If ObjIndex > 0 Then
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageObjectCreate(ObjData(ObjIndex).GrhIndex, .Pos.X, .Pos.Y, ObjData(ObjIndex).ObjType, 0, CanBeTransparent:=ObjData(ObjIndex).CanBeTransparent))
                    End If
                End If
            Else
                If tTrigger = eTrigger.zonaOscura Then
                    If Not (.flags.AdminInvisible = 1) Then _
                        Call SendData(SendTarget.ToUsersAndRmsAndCounselorsAreaButGMs, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, True))
                    
                    ObjIndex = MapData(.Pos.Map, .Pos.X, .Pos.Y).ObjInfo.ObjIndex
                    
                    If ObjIndex > 0 Then
                        Call SendData(SendTarget.ToUsersAndRmsAndCounselorsAreaButGMs, UserIndex, PrepareMessageObjectDelete(.Pos.X, .Pos.Y))
                    End If
                End If
            End If
            
            MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger = tTrigger
            tLog = "Trigger " & tTrigger & " en mapa " & .Pos.Map & " " & .Pos.X & "," & .Pos.Y
            
            Call LogGM(.Name, tLog)
            Call WriteConsoleMsg(UserIndex, tLog, FontTypeNames.FONTTYPE_INFO)
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleSetTrigger de Protocol.bas")
End Sub

''
' Handles the "AskTrigger" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAskTrigger(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 04/13/07
'
'***************************************************
On Error GoTo ErrHandler
  
    Dim tTrigger As Byte
    
    With UserList(UserIndex)

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        tTrigger = MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger
        
        Call LogGM(.Name, "Miro el trigger en " & .Pos.Map & "," & .Pos.X & "," & .Pos.Y & ". Era " & tTrigger)
        
        Call WriteConsoleMsg(UserIndex, _
            "Trigger " & tTrigger & " en mapa " & .Pos.Map & " " & .Pos.X & ", " & .Pos.Y _
            , FontTypeNames.FONTTYPE_INFO)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleAskTrigger de Protocol.bas")
End Sub

''
' Handles the "BannedIPList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBannedIPList(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Dim Lista As String
        Dim LoopC As Long
        
        Call LogGM(.Name, "/BANIPLIST")
        
        For LoopC = 1 To BanIps.Count
            Lista = Lista & BanIps.Item(LoopC) & ", "
        Next LoopC
        
        If LenB(Lista) <> 0 Then Lista = Left$(Lista, Len(Lista) - 2)
        
        Call WriteConsoleMsg(UserIndex, Lista, FontTypeNames.FONTTYPE_INFO)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleBannedIPList de Protocol.bas")
End Sub

''
' Handles the "BannedIPReload" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBannedIPReload(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call BanIpGuardar
        Call BanIpCargar
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleBannedIPReload de Protocol.bas")
End Sub

''
' Handles the "GuildBan" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildBan(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim GuildName As String
        Dim cantMembers As Integer
        Dim LoopC As Long
        Dim member As String
        Dim Count As Byte
        Dim GuildIndex As Integer
        Dim tFile As String
        
        GuildName = Reader.ReadString8()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            'tFile = App.Path & "\guilds\" & GuildName & "-members.mem"
            
            GuildIndex = GetGuildIndex(GuildName)
            
            If GuildIndex = 0 Then
                Call WriteConsoleMsg(UserIndex, "No existe el clan: " & GuildName & ".", FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & " baneó al clan " & UCase$(GuildName), FontTypeNames.FONTTYPE_FIGHT))
                
                'baneamos a los miembros
                Call LogGM(.Name, "BANCLAN a " & UCase$(GuildName))
                
                '' Llamar al banclan
                'Call BanGuild(GuildIndex, .ID)
                
            End If
        End If
        
    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleGuildBan de Protocol.bas")
End Sub

''
' Handles the "BanIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBanIP(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 07/02/09
'Agregado un CopyBuffer porque se producia un bucle
'inifito al intentar banear una ip ya baneada. (NicoNZ)
'07/02/09 Pato - Ahora no es posible saber si un gm está o no online.
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim bannedIP As String
        Dim tUser As Integer
        Dim Reason As String
        Dim I As Long
        
        ' Is it by ip??
        If Reader.ReadBool() Then
            bannedIP = Reader.ReadInt8() & "."
            bannedIP = bannedIP & Reader.ReadInt8() & "."
            bannedIP = bannedIP & Reader.ReadInt8() & "."
            bannedIP = bannedIP & Reader.ReadInt8()
        Else
            tUser = NameIndex(Reader.ReadString8())
            
            If tUser > 0 Then bannedIP = UserList(tUser).IP
        End If
        
        Reason = Reader.ReadString8()
        
        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
            If LenB(bannedIP) > 0 Then
                Call LogGM(.Name, "/BanIP " & bannedIP & " por " & Reason)
                
                If BanIpBuscar(bannedIP) > 0 Then
                    Call WriteConsoleMsg(UserIndex, "La IP " & bannedIP & " ya se encuentra en la lista de bans.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call BanIpAgrega(bannedIP)
                    Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.Name & " baneó la IP " & bannedIP & " por " & Reason, FontTypeNames.FONTTYPE_FIGHT))
                    
                    'Find every player with that ip and ban him!
                    For I = 1 To LastUser
                            If UserList(I).ConnIDValida Then
                                If UserList(I).IP = bannedIP Then
                                    ' TODO: Nightw - Remove hardcoded values
                                    Call BanCharacter(UserIndex, UserList(I).Name, Reason, "", 2)
                                End If
                            End If
                        Next I
                End If
            ElseIf tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "El personaje no está online.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If

    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleBanIP de Protocol.bas")
End Sub

''
' Handles the "UnbanIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUnbanIP(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/30/06
'
'***************************************************
On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim bannedIP As String
        
        bannedIP = Reader.ReadInt8() & "."
        bannedIP = bannedIP & Reader.ReadInt8() & "."
        bannedIP = bannedIP & Reader.ReadInt8() & "."
        bannedIP = bannedIP & Reader.ReadInt8()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        If BanIpQuita(bannedIP) Then
            Call WriteConsoleMsg(UserIndex, "La IP """ & bannedIP & """ se ha quitado de la lista de bans.", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, "La IP """ & bannedIP & """ NO se encuentra en la lista de bans.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleUnbanIP de Protocol.bas")
End Sub

''
' Handles the "CreateItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCreateItem(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************

On Error GoTo ErrHandler
    
    With UserList(UserIndex)

        Dim tObj As Integer
        tObj = Reader.ReadInt16()

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
                
        Dim mapa As Integer
        Dim X As Byte
        Dim Y As Byte
        
        mapa = .Pos.Map
        X = .Pos.X
        Y = .Pos.Y
            
        Call LogGM(.Name, "/CI: " & tObj & " en mapa " & _
            mapa & " (" & X & "," & Y & ")")
        
        If MapData(mapa, X, Y - 1).ObjInfo.ObjIndex > 0 Then _
            Exit Sub
        
        If MapData(mapa, X, Y - 1).TileExit.Map > 0 Then _
            Exit Sub
        
        If tObj < 1 Or tObj > NumObjDatas Then _
            Exit Sub
        
        'Is the object not null?
        If LenB(ObjData(tObj).Name) = 0 Then Exit Sub
        
        ' Resources can't be created manually
        If ObjData(tObj).ObjType = otResource Then
            Call WriteConsoleMsg(UserIndex, "No se pueden crear objetos del tipo Recurso.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Dim Objeto As Obj
        Call WriteConsoleMsg(UserIndex, "¡¡ATENCIÓN: FUERON CREADOS ***100*** ÍTEMS, TIRE Y /DEST LOS QUE NO NECESITE!!", FontTypeNames.FONTTYPE_GUILD)
        
        Objeto.Amount = 100
        Objeto.ObjIndex = tObj
        Call MakeObj(Objeto, mapa, X, Y - 1)
        
        If ObjData(tObj).Log = 1 Then
            Call LogDesarrollo(.Name & " /CI: [" & tObj & "]" & ObjData(tObj).Name & " en mapa " & _
                mapa & " (" & X & "," & Y & ")")
        End If
        
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleCreateItem de Protocol.bas")
End Sub

''
' Handles the "DestroyItems" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDestroyItems(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Dim mapa As Integer
        Dim X As Byte
        Dim Y As Byte
        
        mapa = .Pos.Map
        X = .Pos.X
        Y = .Pos.Y
        
        Dim ObjIndex As Integer
        ObjIndex = MapData(mapa, X, Y).ObjInfo.ObjIndex
        
        If ObjIndex = 0 Then Exit Sub
        
        Call LogGM(.Name, "/DEST " & ObjIndex & " en mapa " & _
            mapa & " (" & X & "," & Y & "). Cantidad: " & MapData(mapa, X, Y).ObjInfo.Amount)
        
        If ObjData(ObjIndex).ObjType = eOBJType.otTeleport And _
            MapData(mapa, X, Y).TileExit.Map > 0 Then
            
            Call WriteConsoleMsg(UserIndex, "No puede destruir teleports así. Utilice /DT.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Call EraseObj(10000, mapa, X, Y)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleDestroyItems de Protocol.bas")
End Sub

Private Sub HandleFactionKick(ByVal UserIndex As Integer)
On Error GoTo ErrHandler
    Dim UserName As String
    Dim Reason As String
    Dim tUser As Integer
        
    UserName = Reader.ReadString8()
    
    With UserList(UserIndex)
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And _
                (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Or _
                .flags.PrivEspecial Then
        
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "\", "")
            End If
            If (InStrB(UserName, "/") <> 0) Then
                UserName = Replace(UserName, "/", "")
            End If
            
            Call ModFacciones.KickFromFactionByName(UserIndex, UserName)
                
        End If
    End With
        
    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleFactionKick de Protocol.bas")
End Sub

''
' Handles the "ForceMIDIAll" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForceMIDIAll(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************

On Error GoTo ErrHandler
    
    With UserList(UserIndex)

        Dim midiID As Byte
        midiID = Reader.ReadInt8()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & " broadcast música: " & midiID, FontTypeNames.FONTTYPE_SERVER))
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayMusic(midiID))
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleForceMIDIAll de Protocol.bas")
End Sub

''
' Handles the "ForceWAVEAll" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForceWAVEAll(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************

On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        Dim waveID As Byte
        waveID = Reader.ReadInt8()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(waveID, NO_3D_SOUND, NO_3D_SOUND))
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleForceWAVEAll de Protocol.bas")
End Sub

''
' Handles the "RemovePunishment" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRemovePunishment(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 1/05/07
'Pablo (ToxicWaste): 1/05/07, You can now edit the punishment.
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim UserName As String
        Dim Punishment As Byte
        Dim NewText As String
        
        UserName = Reader.ReadString8()
        Punishment = Reader.ReadInt8
        NewText = Reader.ReadString8()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            If LenB(UserName) = 0 Then
                Call WriteConsoleMsg(UserIndex, "Utilice /borrarpena Nick@NumeroDePena@NuevaPena", FontTypeNames.FONTTYPE_INFO)
            Else
            
                Dim UserId As Long
                UserId = GetUserID(UserName)
                
                If UserId <> 0 Then
                    
                    Dim punishmentID As Long
                    Dim PunishmentDescrip As String
                    punishmentID = GetUserPunishmentDB(UserId, Punishment, PunishmentDescrip)
                    
                    If punishmentID <> 0 Then
                        Call LogGM(.Name, " borro la pena: " & Punishment & "-" & _
                            PunishmentDescrip & " de " & UserName & " y la cambió por: " & NewText)
                      
                        Call UpdateUserPunishmentDB(UserId, punishmentID, NewText)
                    Else
                        Call WriteConsoleMsg(UserIndex, "El usuario no posee esa pena.", FontTypeNames.FONTTYPE_INFO)
                    End If
                    Call WriteConsoleMsg(UserIndex, "Pena modificada.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleRemovePunishment de Protocol.bas")
End Sub

''
' Handles the "TileBlockedToggle" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTileBlockedToggle(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub

        Call LogGM(.Name, "/BLOQ")
        
        If MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked = 0 Then
            MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked = 1
        Else
            MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked = 0
        End If
        
        Call Bloquear(True, .Pos.Map, .Pos.X, .Pos.Y, MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleTileBlockedToggle de Protocol.bas")
End Sub

''
' Handles the "KillNPCNoRespawn" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKillNPCNoRespawn(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        If .flags.TargetNpc = 0 Then Exit Sub
        
        If Npclist(.flags.TargetNpc).flags.Boss > 0 Then
            Call RestartBossSpawn(Npclist(.flags.TargetNpc).flags.Boss)
        End If
        
        Call QuitarNPC(.flags.TargetNpc)
        Call LogGM(.Name, "/MATA " & Npclist(.flags.TargetNpc).Name)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleKillNPCNoRespawn de Protocol.bas")
End Sub

''
' Handles the "KillAllNearbyNPCs" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKillAllNearbyNPCs(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Dim X As Long
        Dim Y As Long
        
        For Y = .Pos.Y - MinYBorder + 1 To .Pos.Y + MinYBorder - 1
            For X = .Pos.X - MinXBorder + 1 To .Pos.X + MinXBorder - 1
                If X > 0 And Y > 0 And X < 101 And Y < 101 Then
                    If MapData(.Pos.Map, X, Y).NpcIndex > 0 Then Call QuitarNPC(MapData(.Pos.Map, X, Y).NpcIndex)
                End If
            Next X
        Next Y
        Call LogGM(.Name, "/MASSKILL")
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleKillAllNearbyNPCs de Protocol.bas")
End Sub

''
' Handles the "LastIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLastIP(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim UserName As String
        Dim Lista As String
        Dim LoopC As Byte
        Dim Priv As Integer
        Dim validCheck As Boolean
        
        UserName = Reader.ReadString8()

        Priv = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            'Handle special chars
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "\", "")
            End If
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "/", "")
            End If
            If (InStrB(UserName, "+") <> 0) Then
                UserName = Replace(UserName, "+", " ")
            End If
            
            'Only Gods and Admins can see the ips of adminsitrative characters. All others can be seen by every adminsitrative char.
            If NameIndex(UserName) > 0 Then
                validCheck = (UserList(NameIndex(UserName)).flags.Privilegios And Priv) = 0 Or (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
            Else
                validCheck = (UserDarPrivilegioLevel(UserName) And Priv) = 0 Or (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
            End If
            
            If Not validCheck Then
                Call WriteConsoleMsg(UserIndex, UserName & " es de mayor jerarquía que vos.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            Call LogGM(.Name, "/LASTIP " & UserName)

            Dim UserId As Long
            UserId = GetUserID(UserName)
            
            If UserId <= 0 Then
                Call WriteConsoleMsg(UserIndex, "El usuario """ & UserName & """ no existe.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            Call SendUserIps(UserIndex, UserId, UserName)
            
        End If
    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleLastIP de Protocol.bas")
End Sub

''
' Handles the "ChatColor" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleChatColor(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Change the user`s chat color
'***************************************************

On Error GoTo ErrHandler
    
    With UserList(UserIndex)

        Dim color As Long
        
        color = RGB(Reader.ReadInt8(), Reader.ReadInt8(), Reader.ReadInt8())
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.RoleMaster)) Then
            .flags.ChatColor = color
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleChatColor de Protocol.bas")
End Sub

''
' Handles the "Ignored" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleIgnored(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Ignore the user
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then
            .flags.AdminPerseguible = Not .flags.AdminPerseguible
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleIgnored de Protocol.bas")
End Sub

''
' Handles the "CheckSlot" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleCheckSlot(ByVal UserIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 07/06/2010
'Check one Users Slot in Particular from Inventory
'07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        'Reads the UserName and Slot Packets
        Dim UserName As String
        Dim Slot As Byte
        Dim tIndex As Integer
        
        Dim UserIsAdmin As Boolean
        Dim OtherUserIsAdmin As Boolean
                
        UserName = Reader.ReadString8() 'Que UserName?
        Slot = Reader.ReadInt8() 'Que Slot?

        UserIsAdmin = (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0

        If (.flags.Privilegios And PlayerType.SemiDios) <> 0 Or UserIsAdmin Then
            
            Call LogGM(.Name, .Name & " Checkeó el slot " & Slot & " de " & UserName)
            
            tIndex = NameIndex(UserName)  'Que user index?
            OtherUserIsAdmin = EsDios(UserName) Or EsAdmin(UserName)
            
            If tIndex > 0 Then
                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    If Slot > 0 And Slot <= UserList(tIndex).CurrentInventorySlots Then
                        If UserList(tIndex).Invent.Object(Slot).ObjIndex > 0 Then
                            Call WriteConsoleMsg(UserIndex, " Objeto " & Slot & ") " & ObjData(UserList(tIndex).Invent.Object(Slot).ObjIndex).Name & " Cantidad:" & UserList(tIndex).Invent.Object(Slot).Amount, FontTypeNames.FONTTYPE_INFO)
                        Else
                            Call WriteConsoleMsg(UserIndex, "No hay ningún objeto en slot seleccionado.", FontTypeNames.FONTTYPE_INFO)
                        End If
                    Else
                        Call WriteConsoleMsg(UserIndex, "Slot Inválido.", FontTypeNames.FONTTYPE_TALK)
                    End If
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes ver slots de un dios o admin.", FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_TALK)
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes ver slots de un dios o admin.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleCheckSlot de Protocol.bas")
End Sub

''
' Handles the "ResetAutoUpdate" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleResetAutoUpdate(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Reset the AutoUpdate
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        If UCase$(.Name) <> "NIGHTW" Then Exit Sub
        
        Call WriteConsoleMsg(UserIndex, "TID: " & CStr(ReiniciarAutoUpdate()), FontTypeNames.FONTTYPE_INFO)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleResetAutoUpdate de Protocol.bas")
End Sub

''
' Handles the "Restart" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleRestart(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Restart the game
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        If UCase$(.Name) <> "NIGHTW" Then Exit Sub
        
        'time and Time BUG!
        Call LogGM(.Name, .Name & " reinició el mundo.")
        
        Call ReiniciarServidor(True)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleRestart de Protocol.bas")
End Sub

''
' Handles the "ReloadObjects" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleReloadObjects(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Reload the objects
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha recargado los objetos.")
        
        Call LoadOBJData
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleReloadObjects de Protocol.bas")
End Sub

''
' Handles the "ReloadSpells" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleReloadSpells(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Reload the spells
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha recargado los hechizos.")
        
        Call CargarHechizos
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleReloadSpells de Protocol.bas")
End Sub

''
' Handle the "ReloadServerIni" message.
'
' @param userIndex The index of the user sending the message

Public Sub HandleReloadServerIni(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Reload the Server`s INI
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha recargado los INITs.")
        
        Call LoadSini
        
        Call WriteConsoleMsg(UserIndex, "Server.ini actualizado correctamente", FontTypeNames.FONTTYPE_INFO)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleReloadServerIni de Protocol.bas")
End Sub

''
' Handle the "ReloadNPCs" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleReloadNPCs(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Reload the Server`s NPC
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
         
        Call LogGM(.Name, .Name & " ha recargado los NPCs.")
    
        Call CargaNpcsDat
    
        Call WriteConsoleMsg(UserIndex, "Npcs.dat recargado.", FontTypeNames.FONTTYPE_INFO)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleReloadNPCs de Protocol.bas")
End Sub

''
' Handle the "KickAllChars" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleKickAllChars(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Kick all the chars that are online
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha echado a todos los personajes.")
        
        Call EcharPjsNoPrivilegiados
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleKickAllChars de Protocol.bas")
End Sub

''
' Handle the "ShowServerForm" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleShowServerForm(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Show the server form
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha solicitado mostrar el formulario del servidor.")
        Call frmMain.mnuMostrar_Click
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleShowServerForm de Protocol.bas")
End Sub

''
' Handle the "CleanSOS" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleCleanSOS(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Clean the SOS
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha borrado los SOS.")
        
        Call Ayuda.Reset
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleCleanSOS de Protocol.bas")
End Sub

''
' Handle the "SaveChars" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSaveChars(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Save the characters
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha guardado todos los chars.")
        
        Call mdParty.ActualizaExperiencias
        Call GuardarUsuarios
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleSaveChars de Protocol.bas")
End Sub

''
' Handle the "ChangeMapInfoBackup" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoBackup(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/24/06
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Change the backup`s info of the map
'***************************************************

On Error GoTo ErrHandler
      
    With UserList(UserIndex)

        Dim doTheBackUp As Boolean
        
        doTheBackUp = Reader.ReadBool()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) = 0 Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha cambiado la información sobre el BackUp.")
        
        'Change the boolean to byte in a fast way
        If doTheBackUp Then
            MapInfo(.Pos.Map).BackUp = 1
        Else
            MapInfo(.Pos.Map).BackUp = 0
        End If
        
        'Change the boolean to string in a fast way
        Call WriteVar(MapPath & "mapa" & .Pos.Map & ".dat", "Mapa" & .Pos.Map, "backup", MapInfo(.Pos.Map).BackUp)
        
        Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " Backup: " & MapInfo(.Pos.Map).BackUp, FontTypeNames.FONTTYPE_INFO)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleChangeMapInfoBackup de Protocol.bas")
End Sub

''
' Handle the "ChangeMapInfoPK" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoPK(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/24/06
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Change the pk`s info of the  map
'***************************************************

On Error GoTo ErrHandler
      
    With UserList(UserIndex)

        Dim isMapPk As Boolean
        
        isMapPk = Reader.ReadBool()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) = 0 Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha cambiado la información sobre si es PK el mapa.")
        
        MapInfo(.Pos.Map).Pk = isMapPk
        
        'Change the boolean to string in a fast way
        Call WriteVar(MapPath & "mapa" & .Pos.Map & ".dat", "Mapa" & .Pos.Map, "Pk", IIf(isMapPk, "1", "0"))

        Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " PK: " & MapInfo(.Pos.Map).Pk, FontTypeNames.FONTTYPE_INFO)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleChangeMapInfoPK de Protocol.bas")
End Sub

''
' Handle the "ChangeMapInfoRestricted" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoRestricted(ByVal UserIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Restringido -> Options: "NEWBIE", "NO", "ARMADA", "CAOS", "FACCION".
'***************************************************

On Error GoTo ErrHandler
    Dim tStr As String
    
    With UserList(UserIndex)

        tStr = Reader.ReadString8()

        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If tStr = "NEWBIE" Or tStr = "NO" Or tStr = "ARMADA" Or tStr = "CAOS" Or tStr = "FACCION" Then
                Call LogGM(.Name, .Name & " ha cambiado la información sobre si es restringido el mapa.")
                
                MapInfo(UserList(UserIndex).Pos.Map).Restringir = RestrictStringToByte(tStr)
                
                Call WriteVar(MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "Restringir", tStr)
                Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " Restringido: " & RestrictByteToString(MapInfo(.Pos.Map).Restringir), FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Opciones para restringir: 'NEWBIE', 'NO', 'ARMADA', 'CAOS', 'FACCION'", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleChangeMapInfoRestricted de Protocol.bas")
End Sub

''
' Handle the "ChangeMapInfoNoMagic" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoMagic(ByVal UserIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'MagiaSinEfecto -> Options: "1" , "0".
'***************************************************

On Error GoTo ErrHandler
      
    Dim nomagic As Boolean
    
    With UserList(UserIndex)

        nomagic = Reader.ReadBool
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            Call LogGM(.Name, .Name & " ha cambiado la información sobre si está permitido usar la magia el mapa.")
            MapInfo(UserList(UserIndex).Pos.Map).MagiaSinEfecto = nomagic
            Call WriteVar(MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "MagiaSinEfecto", nomagic)
            Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " MagiaSinEfecto: " & MapInfo(.Pos.Map).MagiaSinEfecto, FontTypeNames.FONTTYPE_INFO)
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleChangeMapInfoNoMagic de Protocol.bas")
End Sub

''
' Handle the "ChangeMapInfoNoInvi" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoInvi(ByVal UserIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'InviSinEfecto -> Options: "1", "0"
'***************************************************

On Error GoTo ErrHandler
      
    Dim noinvi As Boolean
    
    With UserList(UserIndex)

        noinvi = Reader.ReadBool()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            Call LogGM(.Name, .Name & " ha cambiado la información sobre si está permitido usar la invisibilidad en el mapa.")
            MapInfo(UserList(UserIndex).Pos.Map).InviSinEfecto = noinvi
            Call WriteVar(MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "InviSinEfecto", noinvi)
            Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " InviSinEfecto: " & MapInfo(.Pos.Map).InviSinEfecto, FontTypeNames.FONTTYPE_INFO)
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleChangeMapInfoNoInvi de Protocol.bas")
End Sub
            
''
' Handle the "ChangeMapInfoNoResu" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoResu(ByVal UserIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'ResuSinEfecto -> Options: "1", "0"
'***************************************************

On Error GoTo ErrHandler
  
    Dim noresu As Boolean
    
    With UserList(UserIndex)

        noresu = Reader.ReadBool()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            Call LogGM(.Name, .Name & " ha cambiado la información sobre si está permitido usar el resucitar en el mapa.")
            MapInfo(UserList(UserIndex).Pos.Map).ResuSinEfecto = noresu
            Call WriteVar(MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "ResuSinEfecto", noresu)
            Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " ResuSinEfecto: " & MapInfo(.Pos.Map).ResuSinEfecto, FontTypeNames.FONTTYPE_INFO)
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleChangeMapInfoNoResu de Protocol.bas")
End Sub

''
' Handle the "ChangeMapInfoLand" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoLand(ByVal UserIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Terreno -> Opciones: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
'***************************************************

On Error GoTo ErrHandler
    Dim tStr As String
    
    With UserList(UserIndex)

        tStr = Reader.ReadString8()

        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If tStr = "BOSQUE" Or tStr = "NIEVE" Or tStr = "DESIERTO" Or tStr = "CIUDAD" Or tStr = "CAMPO" Or tStr = "DUNGEON" Then
                Call LogGM(.Name, .Name & " ha cambiado la información del terreno del mapa.")
                
                MapInfo(UserList(UserIndex).Pos.Map).Terreno = TerrainZoneStringToByte(tStr)
                
                Call WriteVar(MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "Terreno", tStr)
                Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " Terreno: " & tStr, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Opciones para terreno: 'BOSQUE', 'NIEVE', 'DESIERTO', 'CIUDAD', 'CAMPO', 'DUNGEON'", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(UserIndex, "Igualmente, el único útil es 'NIEVE' ya que al ingresarlo, la gente muere de frío en el mapa.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleChangeMapInfoLand de Protocol.bas")
End Sub

''
' Handle the "ChangeMapInfoZone" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoZone(ByVal UserIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Zona -> Opciones: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
'***************************************************

On Error GoTo ErrHandler
    Dim tStr As String
    
    With UserList(UserIndex)

        tStr = Reader.ReadString8()

        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If tStr = "BOSQUE" Or tStr = "NIEVE" Or tStr = "DESIERTO" Or tStr = "CIUDAD" Or tStr = "CAMPO" Or tStr = "DUNGEON" Then
                Call LogGM(.Name, .Name & " ha cambiado la información de la zona del mapa.")
                MapInfo(UserList(UserIndex).Pos.Map).Zona = TerrainZoneStringToByte(tStr)
                Call WriteVar(MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "Zona", tStr)
                Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " Zona: " & tStr, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Opciones para terreno: 'BOSQUE', 'NIEVE', 'DESIERTO', 'CIUDAD', 'CAMPO', 'DUNGEON'", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(UserIndex, "Igualmente, el único útil es 'DUNGEON' ya que al ingresarlo, NO se sentirá el efecto de la lluvia en este mapa.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleChangeMapInfoZone de Protocol.bas")
End Sub
            
''
' Handle the "ChangeMapInfoStealNp" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoStealNpc(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 25/07/2010
'RoboNpcsPermitido -> Options: "1", "0"
'***************************************************

On Error GoTo ErrHandler
  
    Dim RoboNpc As Byte
    
    With UserList(UserIndex)

        RoboNpc = Val(IIf(Reader.ReadBool(), 1, 0))
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            Call LogGM(.Name, .Name & " ha cambiado la información sobre si está permitido robar npcs en el mapa.")
            
            MapInfo(UserList(UserIndex).Pos.Map).RoboNpcsPermitido = RoboNpc
            
            Call WriteVar(MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "RoboNpcsPermitido", RoboNpc)
            Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " RoboNpcsPermitido: " & MapInfo(.Pos.Map).RoboNpcsPermitido, FontTypeNames.FONTTYPE_INFO)
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleChangeMapInfoStealNpc de Protocol.bas")
End Sub
            
''
' Handle the "ChangeMapInfoNoOcultar" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoOcultar(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 18/09/2010
'OcultarSinEfecto -> Options: "1", "0"
'***************************************************

On Error GoTo ErrHandler
  
    Dim NoOcultar As Byte
    Dim mapa As Integer
    
    With UserList(UserIndex)

        NoOcultar = Val(IIf(Reader.ReadBool(), 1, 0))
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            
            mapa = .Pos.Map
            
            Call LogGM(.Name, .Name & " ha cambiado la información sobre si está permitido ocultarse en el mapa " & mapa & ".")
            
            MapInfo(mapa).OcultarSinEfecto = NoOcultar
            
            Call WriteVar(MapPath & "mapa" & mapa & ".dat", "Mapa" & mapa, "OcultarSinEfecto", NoOcultar)
            Call WriteConsoleMsg(UserIndex, "Mapa " & mapa & " OcultarSinEfecto: " & NoOcultar, FontTypeNames.FONTTYPE_INFO)
        End If
        
    End With
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleChangeMapInfoNoOcultar de Protocol.bas")
End Sub
           
''
' Handle the "ChangeMapInfoNoInvocar" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoInvocar(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 18/09/2010
'InvocarSinEfecto -> Options: "1", "0"
'***************************************************

On Error GoTo ErrHandler
    
    Dim NoInvocar As Byte
    Dim mapa As Integer
    
    With UserList(UserIndex)

        NoInvocar = Val(IIf(Reader.ReadBool(), 1, 0))
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            
            mapa = .Pos.Map
            
            Call LogGM(.Name, .Name & " ha cambiado la información sobre si está permitido invocar en el mapa " & mapa & ".")
            
            MapInfo(mapa).InvocarSinEfecto = NoInvocar
            
            Call WriteVar(MapPath & "mapa" & mapa & ".dat", "Mapa" & mapa, "InvocarSinEfecto", NoInvocar)
            Call WriteConsoleMsg(UserIndex, "Mapa " & mapa & " InvocarSinEfecto: " & NoInvocar, FontTypeNames.FONTTYPE_INFO)
        End If
        
    End With
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleChangeMapInfoNoInvocar de Protocol.bas")
End Sub

''
' Handle the "SaveMap" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSaveMap(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/24/06
'Saves the map
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha guardado el mapa " & CStr(.Pos.Map))
        
        Call GrabarMapa(.Pos.Map, ServerConfiguration.ResourcesPaths.WorldBackup & "Mapa" & .Pos.Map)
        
        Call WriteConsoleMsg(UserIndex, "Mapa Guardado.", FontTypeNames.FONTTYPE_INFO)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleSaveMap de Protocol.bas")
End Sub

''
' Handle the "DoBackUp" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleDoBackUp(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/24/06
'Show guilds messages
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha hecho un backup.")
        
        Call ES.DoBackUp 'Sino lo confunde con la id del paquete
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleDoBackUp de Protocol.bas")
End Sub

''
' Handle the "AlterName" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleAlterName(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/26/06
'Change user name
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        'Reads the userName and newUser Packets
        Dim UserName As String
        Dim newName As String
        Dim changeNameUI As Integer
        Dim GuildIndex As Integer
        
        UserName = Reader.ReadString8()
        newName = Reader.ReadString8()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Or .flags.PrivEspecial Then
            If LenB(UserName) = 0 Or LenB(newName) = 0 Then
                Call WriteConsoleMsg(UserIndex, "Usar: /ANAME origen@destino", FontTypeNames.FONTTYPE_INFO)
            Else
                changeNameUI = NameIndex(UserName)
                
                If changeNameUI > 0 Then
                    Call WriteConsoleMsg(UserIndex, "El Pj está online, debe salir para hacer el cambio.", FontTypeNames.FONTTYPE_WARNING)
                Else
                    Dim UserId As Long
                    UserId = GetUserID(UserName)
                    
                    If UserId = 0 Then
                        Call WriteConsoleMsg(UserIndex, "El pj " & UserName & " es inexistente.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        GuildIndex = CInt(Val(GetCharData("USER_INFO", "GUILD_ID", UserId)))
                        
                        If GuildIndex > 0 Then
                            Call WriteConsoleMsg(UserIndex, "El pj " & UserName & " pertenece a un clan, debe salir del mismo con /salirclan para ser transferido.", FontTypeNames.FONTTYPE_INFO)
                        Else
                            Dim TempID As Long
                            TempID = GetUserID(newName)
                            
                            If TempID = 0 Then
                                If BlacklistIsValidNickname(newName) Then
                                    Call UpdateUserNameDB(UserId, newName)
                                    Call BlacklistAppend(UserName)
                                    Call WriteConsoleMsg(UserIndex, "Transferencia exitosa.", FontTypeNames.FONTTYPE_INFO)
                                    Call LogGM(.Name, "Ha cambiado de nombre al usuario " & _
                                               UserName & ". Ahora se llama " & newName)
                                Else
                                    Call WriteConsoleMsg(UserIndex, "El nombre solicitado es inválido.", _
                                                         FontTypeNames.FONTTYPE_INFO)
                                End If
                            Else
                                Call WriteConsoleMsg(UserIndex, "El nick solicitado ya existe.", FontTypeNames.FONTTYPE_INFO)
                            End If
                        End If
                    End If
                End If
            End If
        End If
    
    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleAlterName de Protocol.bas")
End Sub

''
' Handle the "AlterMail" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleAlterMail(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/01/2015 (D'Artagnan)
'Change user password
'05/01/2015: D'Artagnan - Migrated to accounts.
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim UserName As String
        Dim newMail As String
        Dim nAccountID As Long
        
        UserName = Reader.ReadString8()
        newMail = Reader.ReadString8()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            If LenB(UserName) = 0 Or LenB(newMail) = 0 Then
                Call WriteConsoleMsg(UserIndex, "usar /AEMAIL <pj>-<nuevomail>", FontTypeNames.FONTTYPE_INFO)
            Else
                Dim UserId As Long
                UserId = GetUserID(UserName)
                
                If UserId = 0 Then
                    Call WriteConsoleMsg(UserIndex, "No existe el personaje.", FontTypeNames.FONTTYPE_INFO)
                Else
                    nAccountID = GetAccountIDByUserID(UserId)
                    If Not AccountEmailExists(newMail) Then
                        Call UpdateAccountMail(nAccountID, newMail)
                        Call WriteConsoleMsg(UserIndex, "Email de la cuenta perteneciente a " & _
                                             UserName & " cambiado a: " & newMail & ".", _
                                             FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call WriteConsoleMsg(UserIndex, "La dirección ingresada ya está en uso.", _
                                             FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
                
                Call LogGM(.Name, "Le ha cambiado el mail a " & UserName)
            End If
        End If
        

    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleAlterMail de Protocol.bas")
End Sub

''
' Handle the "AlterPassword" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleAlterPassword(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/26/06
'Change user password
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Dim UserName As String
            Dim NewPassword As String
            Dim Password As String
        
            UserName = Replace(Reader.ReadString8(), "+", " ")

                NewPassword = Replace(Reader.ReadString8(), "+", " ")

                If LenB(UserName) = 0 Or LenB(NewPassword) = 0 Then
                Call WriteConsoleMsg(UserIndex, "Alguno de los parámetros está vacío. Use /APASS <user>@<nuevaPass>", FontTypeNames.FONTTYPE_INFO)
            Else
                If LenB(NewPassword) = 0 Then
                    Call WriteConsoleMsg(UserIndex, "No se puede utilizar una contraseña vacía. Por favor, asegúrese de utilizar el comando /APASS <user>@<nuevaPass>", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call UpdateCharPassword(UserName, NewPassword)
                
                    Call WriteConsoleMsg(UserIndex, .Name & " ha cambiado la password de " & UserName & ".", FontTypeNames.FONTTYPE_INFO)
                    Call LogGM(.Name, "Ha alterado la contraseña de " & UserName)
                End If
            End If
        End If
    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleAlterPassword de Protocol.bas")
End Sub

''
' Handle the "AlterName" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleAlterGuildName(ByVal UserIndex As Integer)
'***************************************************
'Author: Lex!
'Last Modification: 14/05/12
'Change guild name
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        'Reads the userName and newUser Packets
        Dim GuildSpecialName As String
        Dim newGuildSpecialName As String
        Dim GuildName As String
        Dim newGuildName As String
        Dim GuildIndex As Integer
        
        GuildSpecialName = Trim$(Reader.ReadString8())
        newGuildSpecialName = Trim$(Reader.ReadString8())

    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleAlterGuildName de Protocol.bas")
End Sub

''
' Handle the "HandleCreateNPC" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleCreateNPC(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 26/09/2010
'26/09/2010: ZaMa - Ya no se pueden crear npcs pretorianos.
'***************************************************

On Error GoTo ErrHandler
      
    With UserList(UserIndex)

        Dim NpcIndex As Integer
        
        NpcIndex = Reader.ReadInt16()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        If NpcIndex >= 900 And NpcIndex <= 923 Then
            Call WriteConsoleMsg(UserIndex, "No puedes sumonear miembros del clan pretoriano de esta forma, utiliza /CrearClanPretoriano.", FontTypeNames.FONTTYPE_WARNING)
            Exit Sub
        End If
        
        NpcIndex = SpawnNpc(NpcIndex, .Pos, True, False)
        
        If NpcIndex <> 0 Then
            Call LogGM(.Name, "Sumoneó a " & Npclist(NpcIndex).Name & " en mapa " & .Pos.Map)
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleCreateNPC de Protocol.bas")
End Sub


''
' Handle the "CreateNPCWithRespawn" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleCreateNPCWithRespawn(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 26/09/2010
'26/09/2010: ZaMa - Ya no se pueden crear npcs pretorianos.
'***************************************************

On Error GoTo ErrHandler
    
    With UserList(UserIndex)

        Dim NpcIndex As Integer
        
        NpcIndex = Reader.ReadInt16()
        
        If NpcIndex >= 900 And NpcIndex <= 923 Then
            Call WriteConsoleMsg(UserIndex, "No puedes sumonear miembros del clan pretoriano de esta forma, utiliza /CrearClanPretoriano.", FontTypeNames.FONTTYPE_WARNING)
            Exit Sub
        End If
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        NpcIndex = SpawnNpc(NpcIndex, .Pos, True, True)
        
        If NpcIndex <> 0 Then
            Call LogGM(.Name, "Sumoneó con respawn " & Npclist(NpcIndex).Name & " en mapa " & .Pos.Map)
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleCreateNPCWithRespawn de Protocol.bas")
End Sub

''
' Handle the "ImperialArmour" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleImperialArmour(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/24/06
'
'***************************************************
On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim Index As Byte
        Dim ObjIndex As Integer
        
        Index = Reader.ReadInt8()
        ObjIndex = Reader.ReadInt16()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Select Case Index
            Case 1
                ArmaduraImperial1 = ObjIndex
            
            Case 2
                ArmaduraImperial2 = ObjIndex
            
            Case 3
                ArmaduraImperial3 = ObjIndex
            
            Case 4
                TunicaMagoImperial = ObjIndex
        End Select
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleImperialArmour de Protocol.bas")
End Sub

''
' Handle the "ChaosArmour" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChaosArmour(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/24/06
'
'***************************************************

On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        Dim Index As Byte
        Dim ObjIndex As Integer
        
        Index = Reader.ReadInt8()
        ObjIndex = Reader.ReadInt16()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Select Case Index
            Case 1
                ArmaduraCaos1 = ObjIndex
            
            Case 2
                ArmaduraCaos2 = ObjIndex
            
            Case 3
                ArmaduraCaos3 = ObjIndex
            
            Case 4
                TunicaMagoCaos = ObjIndex
        End Select
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleChaosArmour de Protocol.bas")
End Sub

''
' Handle the "NavigateToggle" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleNavigateToggle(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/12/07
'
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub
        
        If .flags.Navegando = 1 Then
            .flags.Navegando = 0
            Call WriteNavigateChange(UserIndex, False)
        Else
            .flags.Navegando = 1
            Call WriteNavigateChange(UserIndex, True)
        End If
        
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleNavigateToggle de Protocol.bas")
End Sub

''
' Handle the "ServerOpenToUsersToggle" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleServerOpenToUsersToggle(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/24/06
'
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        If ServerSoloGMs > 0 Then
            Call WriteConsoleMsg(UserIndex, "Servidor habilitado para todos.", FontTypeNames.FONTTYPE_INFO)
            ServerSoloGMs = 0
            frmMain.chkServerHabilitado.Value = vbUnchecked
        Else
            Call WriteConsoleMsg(UserIndex, "Servidor restringido a administradores.", FontTypeNames.FONTTYPE_INFO)
            ServerSoloGMs = 1
            frmMain.chkServerHabilitado.Value = vbChecked
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleServerOpenToUsersToggle de Protocol.bas")
End Sub


''
' Handle the "ResetFactions" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleResetFactions(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/03/2015
'04/03/2015: D'Artagnan - Minor changes and code optimization.
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim UserName As String
        Dim tUser As Integer
        Dim Char As String
        Dim cantPenas As Byte
        
        UserName = Reader.ReadString8()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Or .flags.PrivEspecial Then
            Call LogGM(.Name, "/RAJAR " & UserName)
            
            tUser = NameIndex(UserName)
            Dim UserId As Long
            
            If tUser > 0 Then
                UserId = UserList(tUser).Id
            Else
                UserId = GetUserID(UserName)
            End If
            
            If UserId <> 0 Then
                If tUser > 0 Then
                    Call ResetFacciones(tUser)
                Else
                    Call ResetUserFactionDB(UserId)
                End If
                Call AddPunishmentDB_OLDBORRAR(UserId, .Id, 45, Now, "Personaje reincorporado a la facción.", "")
                Call WriteConsoleMsg(UserIndex, "Estadísticas de facción restablecidas.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "El personaje " & UserName & " no existe.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
    
    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleResetFactions de Protocol.bas")
End Sub

''
' Handle the "RequestCharMail" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleRequestCharMail(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/01/2015 (D'Artagnan)
'Request user mail
'01/01/2015: D'Artagnan - Retrieve email from the accounts table.
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim UserName As String
        Dim mail As String
        
        UserName = Reader.ReadString8()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Or .flags.PrivEspecial Then
            Dim UserId As Long
            UserId = GetUserID(UserName)
            
            If UserId <> 0 Then
                mail = GetAccountData("EMAIL", GetAccountIDByUserID(UserId))
                Call WriteConsoleMsg(UserIndex, "Last email de " & UserName & ": " & mail, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "El usuario no existe.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleRequestCharMail de Protocol.bas")
End Sub

''
' Handle the "SystemMessage" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSystemMessage(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/29/06
'Send a message to all the users
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim Message As String
        Message = Reader.ReadString8()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.Name, "Mensaje de sistema:" & Message)
            
            Call SendData(SendTarget.ToAll, 0, PrepareMessageShowMessageBox(Message))
        End If
    
    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleSystemMessage de Protocol.bas")
End Sub

''
' Handle the "SetMOTD" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSetMOTD(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 03/31/07
'Set the MOTD
'Modified by: Juan Martín Sotuyo Dodero (Maraxus)
'   - Fixed a bug that prevented from properly setting the new number of lines.
'   - Fixed a bug that caused the player to be kicked.
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim newMOTD As String
        Dim auxiliaryString() As String
        Dim LoopC As Long
        
        newMOTD = Reader.ReadString8()

        auxiliaryString = Split(newMOTD, vbCrLf)
        
        If ((Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And _
             (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios))) Or _
              .flags.PrivEspecial Then
              
            Call LogGM(.Name, "Ha fijado un nuevo MOTD")
            
            MaxLines = UBound(auxiliaryString()) + 1
            
            ReDim MOTD(1 To MaxLines)
            
            Call WriteVar(ServerConfiguration.ResourcesPaths.Dats & "Motd.ini", "INIT", "NumLines", CStr(MaxLines))
            
            For LoopC = 1 To MaxLines
                Call WriteVar(ServerConfiguration.ResourcesPaths.Dats & "Motd.ini", "Motd", "Line" & CStr(LoopC), auxiliaryString(LoopC - 1))
                
                MOTD(LoopC).texto = auxiliaryString(LoopC - 1)
            Next LoopC
            
            Call WriteConsoleMsg(UserIndex, "Se ha cambiado el MOTD con éxito.", FontTypeNames.FONTTYPE_INFO)
        End If
    
    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleSetMOTD de Protocol.bas")
End Sub

''
' Handle the "ChangeMOTD" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMOTD(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín sotuyo Dodero (Maraxus)
'Last Modification: 12/29/06
'Change the MOTD
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        If (.flags.Privilegios And (PlayerType.RoleMaster Or PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios)) Then
            Exit Sub
        End If
        
        Dim auxiliaryString As String
        Dim LoopC As Long
        
        For LoopC = LBound(MOTD()) To UBound(MOTD())
            auxiliaryString = auxiliaryString & MOTD(LoopC).texto & vbCrLf
        Next LoopC
        
        If Len(auxiliaryString) >= 2 Then
            If Right$(auxiliaryString, 2) = vbCrLf Then
                auxiliaryString = Left$(auxiliaryString, Len(auxiliaryString) - 2)
            End If
        End If
        
        Call WriteShowMOTDEditionForm(UserIndex, auxiliaryString)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleChangeMOTD de Protocol.bas")
End Sub

''
' Handle the "Ping" message
'
' @param userIndex The index of the user sending the message

Public Sub HandlePing(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/24/06
'Show guilds messages
'***************************************************
On Error GoTo ErrHandler

        Call WritePong(UserIndex, Reader.ReadInt32())
        
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandlePing de Protocol.bas")
End Sub

''
' Handle the "SetIniVar" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSetIniVar(ByVal UserIndex As Integer)
'***************************************************
'Author: Brian Chaia (BrianPr)
'Last Modification: 01/23/10 (Marco)
'Modify server.ini
'***************************************************

On Error GoTo ErrHandler

    With UserList(UserIndex)
        Dim sLlave As String
        Dim sClave As String
        Dim sValor As String

        'Obtengo los parámetros
        sLlave = Reader.ReadString8()
        sClave = Reader.ReadString8()
        sValor = Reader.ReadString8()

        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
            Dim sTmp As String

            'No podemos modificar [INIT]Dioses ni [Dioses]*
            If (UCase$(sLlave) = "INIT" And UCase$(sClave) = "DIOSES") Or UCase$(sLlave) = "DIOSES" Then
                Call WriteConsoleMsg(UserIndex, "¡No puedes modificar esa información desde aquí!", FontTypeNames.FONTTYPE_INFO)
            Else
                'Obtengo el valor según llave y clave
                sTmp = GetVar(IniPath & "Server.ini", sLlave, sClave)

                'Si obtengo un valor escribo en el server.ini
                If LenB(sTmp) Then
                    Call WriteVar(IniPath & "Server.ini", sLlave, sClave, sValor)
                    Call LogGM(.Name, "Modificó en server.ini (" & sLlave & " " & sClave & ") el valor " & sTmp & " por " & sValor)
                    Call WriteConsoleMsg(UserIndex, "Modificó " & sLlave & " " & sClave & " a " & sValor & ". Valor anterior " & sTmp, FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, "No existe la llave y/o clave", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleSetIniVar de Protocol.bas")
End Sub

''
' Handle the "CreatePretorianClan" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleCreatePretorianClan(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 29/10/2010
'***************************************************

On Error GoTo ErrHandler

    Dim Map As Integer
    Dim X As Byte
    Dim Y As Byte
    Dim Index As Long
    
    With UserList(UserIndex)

        Map = Reader.ReadInt16()
        X = Reader.ReadInt8()
        Y = Reader.ReadInt8()
        
        ' User Admin?
         If ((.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) = 0) Or ((.flags.Privilegios And PlayerType.RoleMaster) <> 0) Then Exit Sub
        
        ' Valid pos?
        If Not InMapBounds(Map, X, Y) Then
            Call WriteConsoleMsg(UserIndex, "Posición inválida.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        ' Choose pretorian clan Index
        If Map = MAPA_PRETORIANO Then
            Index = 1 ' Default clan
        Else
            Index = 2 ' Custom Clan
        End If
            
        ' Is already active any clan?
        If Not ClanPretoriano(Index).Active Then
            
            If Not ClanPretoriano(Index).SpawnClan(Map, X, Y, Index) Then
                Call WriteConsoleMsg(UserIndex, "La posición no es apropiada para crear el clan", FontTypeNames.FONTTYPE_INFO)
            End If
        
        Else
            Call WriteConsoleMsg(UserIndex, "El clan pretoriano se encuentra activo en el mapa " & _
                ClanPretoriano(Index).ClanMap & ". Utilice /EliminarPretorianos MAPA y reintente.", FontTypeNames.FONTTYPE_INFO)
        End If
        
        Call LogGM(.Name, "Utilizó el comando /CREARPRETORIANOS " & Map & " " & X & " " & Y)
        
    End With

    Exit Sub

ErrHandler:
    Call LogError("Error en HandleCreatePretorianClan. Error: " & Err.Number & " - " & Err.Description)
End Sub

''
' Handle the "CreatePretorianClan" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleDeletePretorianClan(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 29/10/2010
'***************************************************

On Error GoTo ErrHandler
    
    Dim Map As Integer
    Dim Index As Long
    
    With UserList(UserIndex)

        Map = Reader.ReadInt16()
        
        ' User Admin?
        If ((.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) = 0) Or ((.flags.Privilegios And PlayerType.RoleMaster) <> 0) Then Exit Sub
        
        ' Valid map?
        If Map < 1 Or Map > NumMaps Then
            Call WriteConsoleMsg(UserIndex, "Mapa inválido.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    
        For Index = 1 To UBound(ClanPretoriano)
         
            ' Search for the clan to be deleted
            If ClanPretoriano(Index).ClanMap = Map Then
                ClanPretoriano(Index).DeleteClan
                Exit For
            End If
        
        Next Index
        
        Call LogGM(.Name, "Utilizó el comando /ELIMINARPRETORIANOS " & Map)
        
    End With

    Exit Sub

ErrHandler:
    Call LogError("Error en HandleDeletePretorianClan. Error: " & Err.Number & " - " & Err.Description)
End Sub


Public Sub WriteConnectedMessage(ByVal UserIndex As Integer)

    Call Writer.WriteInt8(ServerPacketID.Connected)
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Writes the "Logged" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLoggedMessage(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Logged" message to the given user's outgoing data buffer
'***************************************************

    With UserList(UserIndex)
        Call Writer.WriteInt8(ServerPacketID.logged)
            
    End With
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Writes the "RemoveDialogs" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemoveAllDialogs(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RemoveDialogs" message to the given user's outgoing data buffer
'***************************************************

    Call Writer.WriteInt8(ServerPacketID.RemoveDialogs)
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Writes the "RemoveCharDialog" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character whose dialog will be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemoveCharDialog(ByVal UserIndex As Integer, ByVal CharIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RemoveCharDialog" message to the given user's outgoing data buffer
'***************************************************

    Call SendData(ToUser, UserIndex, PrepareMessageRemoveCharDialog(CharIndex))
    
End Sub

''
' Writes the "NavigateChange" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNavigateChange(ByVal UserIndex As Integer, ByVal NewValue As Boolean)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "NavigateChange" message to the given user's outgoing data buffer
'***************************************************

    Call Writer.WriteInt8(ServerPacketID.NavigateChange)
    Call Writer.WriteBool(NewValue)
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Writes the "Disconnect" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDisconnect(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Disconnect" message to the given user's outgoing data buffer
'***************************************************

    Call Writer.WriteInt8(ServerPacketID.Disconnect)
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Writes the "UserOfferConfirm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserOfferConfirm(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 14/12/2009
'Writes the "UserOfferConfirm" message to the given user's outgoing data buffer
'***************************************************

    Call Writer.WriteInt8(ServerPacketID.UserOfferConfirm)
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub


''
' Writes the "CommerceEnd" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceEnd(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CommerceEnd" message to the given user's outgoing data buffer
'***************************************************

    Call Writer.WriteInt8(ServerPacketID.CommerceEnd)
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Writes the "BankEnd" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankEnd(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankEnd" message to the given user's outgoing data buffer
'***************************************************

    Call Writer.WriteInt8(ServerPacketID.BankEnd)
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Writes the "CommerceInit" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceInit(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CommerceInit" message to the given user's outgoing data buffer
'***************************************************

    Call Writer.WriteInt8(ServerPacketID.CommerceInit)
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Writes the "BankInit" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankInit(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankInit" message to the given user's outgoing data buffer
'***************************************************

    Call Writer.WriteInt8(ServerPacketID.BankInit)
    Call Writer.WriteInt32(UserList(UserIndex).Stats.Banco)
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Writes the "UserCommerceInit" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceInit(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserCommerceInit" message to the given user's outgoing data buffer
'***************************************************

    Call Writer.WriteInt8(ServerPacketID.UserCommerceInit)
    Call Writer.WriteString8(UserList(UserIndex).ComUsu.DestNick)
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Writes the "UserCommerceEnd" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceEnd(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserCommerceEnd" message to the given user's outgoing data buffer
'***************************************************

    Call Writer.WriteInt8(ServerPacketID.UserCommerceEnd)
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Writes the "ShowBlacksmithForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowCraftForm(ByVal UserIndex As Integer)

    Call Writer.WriteInt8(ServerPacketID.ShowCraftForm)
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Writes the "UpdateSta" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateSta(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateMana" message to the given user's outgoing data buffer
'***************************************************

        Call Writer.WriteInt8(ServerPacketID.UpdateSta)
        Call Writer.WriteInt16(UserList(UserIndex).Stats.MinSta)
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Writes the "UpdateMana" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateMana(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateMana" message to the given user's outgoing data buffer
'***************************************************

        Call Writer.WriteInt8(ServerPacketID.UpdateMana)
        Call Writer.WriteInt16(UserList(UserIndex).Stats.MinMAN)
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Writes the "UpdateHP" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateHP(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateMana" message to the given user's outgoing data buffer
'***************************************************

        Call Writer.WriteInt8(ServerPacketID.UpdateHP)
        Call Writer.WriteInt16(UserList(UserIndex).Stats.MinHp)
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

Public Sub WriteUpdateChallenge(ByVal UserIndex As Integer)
'***************************************************
'Author:
'Last Modification:
'***************************************************

        Call Writer.WriteInt8(ServerPacketID.UpdateChallenge)
        
        Dim Sand As Byte
        Sand = UserList(UserIndex).Challenge.InSand
        
        Call Writer.WriteInt32(SandsChallenge(Sand).Amount_gold)
        Call Writer.WriteInt8(SandsChallenge(Sand).Maxim_dead)
        Call Writer.WriteInt8(SandsChallenge(Sand).Event_time)
        Call Writer.WriteInt8(SandsChallenge(Sand).Time_start)
        Call Writer.WriteInt8(SandsChallenge(Sand).Event_map)
        Call Writer.WriteInt8(SandsChallenge(Sand).Invisibility)
        Call Writer.WriteInt8(SandsChallenge(Sand).Resucitar)
        Call Writer.WriteInt8(SandsChallenge(Sand).Elementary)
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

' TODO: Nightw - The Challenge system was replaced by the Duel system. This should be removed.
Public Sub WriteUpdateChallengeStat(ByVal UserIndex As Integer)


        Call Writer.WriteInt8(ServerPacketID.UpdateChallengeStat)
        
        'Dim Sand As Byte
        'Sand = UserList(UserIndex).Challenge.InSand
        
        'Call Writer.WriteInt8(Sand)
        
        'Call Writer.WriteInt32(SandsChallenge(Sand).Amount_gold)
        'Call Writer.WriteInt8(SandsChallenge(Sand).Maxim_dead)
        'Call Writer.WriteInt8(SandsChallenge(Sand).Event_time)
        'Call Writer.WriteInt8(SandsChallenge(Sand).Time_start)
        'Call Writer.WriteInt8(SandsChallenge(Sand).Event_map)
        'Call Writer.WriteInt8(SandsChallenge(Sand).Invisibility)
        'Call Writer.WriteInt8(SandsChallenge(Sand).Resucitar)
        'Call Writer.WriteInt8(SandsChallenge(Sand).Elementary)
        
        'Call Writer.WriteInt8(UserList(UserIndex).Challenge.TeamSelect)
        'Call Writer.WriteInt8(SandsChallenge(Sand).DeadPoints(1))
        'Call Writer.WriteInt8(SandsChallenge(Sand).DeadPoints(2))
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Writes the "UpdateGold" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateGold(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateGold" message to the given user's outgoing data buffer
'***************************************************

        Call Writer.WriteInt8(ServerPacketID.UpdateGold)
        Call Writer.WriteInt32(UserList(UserIndex).Stats.GLD)
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Writes the "UpdateBankGold" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateBankGold(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 14/12/2009
'Writes the "UpdateBankGold" message to the given user's outgoing data buffer
'***************************************************

        Call Writer.WriteInt8(ServerPacketID.UpdateBankGold)
        Call Writer.WriteInt32(UserList(UserIndex).Stats.Banco)
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub


''
' Writes the "UpdateExp" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateExp(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateExp" message to the given user's outgoing data buffer
'***************************************************

        Call Writer.WriteInt8(ServerPacketID.UpdateExp)
        Call Writer.WriteInt32(UserList(UserIndex).Stats.Exp)
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateStrenghtAndDexterity(ByVal UserIndex As Integer)
'***************************************************
'Author: Budi
'Last Modification: 11/26/09
'Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer
'***************************************************

    With UserList(UserIndex)
        .flags.bStrDextRunningOutNotified = False

            Call Writer.WriteInt8(ServerPacketID.UpdateStrenghtAndDexterity)
            Call Writer.WriteInt8(.Stats.UserAtributos(eAtributos.Fuerza))
            Call Writer.WriteInt8(.Stats.UserAtributos(eAtributos.Agilidad))

    End With
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub


' Writes the "WriteUpdateCharacterInfo" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUpdateCharacterInfo(ByVal UserIndex As Integer)

    With UserList(UserIndex)

        Call Writer.WriteInt8(ServerPacketID.UpdateCharacterInfo)
        Call Writer.WriteInt(.clase)
        Call Writer.WriteInt(.raza)
        Call Writer.WriteInt(.Genero)

    End With
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

' Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateDexterity(ByVal UserIndex As Integer)
'***************************************************
'Author: Budi
'Last Modification: 11/26/09
'Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer
'***************************************************

    With UserList(UserIndex)
        .flags.bStrDextRunningOutNotified = False

            Call Writer.WriteInt8(ServerPacketID.UpdateDexterity)
            Call Writer.WriteInt8(.Stats.UserAtributos(eAtributos.Agilidad))

    End With
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

' Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateStrenght(ByVal UserIndex As Integer)
'***************************************************
'Author: Budi
'Last Modification: 11/26/09
'Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer
'***************************************************

    With UserList(UserIndex)
        .flags.bStrDextRunningOutNotified = False

            Call Writer.WriteInt8(ServerPacketID.UpdateStrenght)
            Call Writer.WriteInt8(.Stats.UserAtributos(eAtributos.Fuerza))

    End With
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Writes the "ChangeMap" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    map The new map to load.
' @param    version The version of the map in the server to check if client is properly updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMap(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal version As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeMap" message to the given user's outgoing data buffer
'***************************************************

    Call Writer.WriteInt8(ServerPacketID.ChangeMap)
    Call Writer.WriteInt16(Map)
    Call Writer.WriteInt16(version)
    Call Writer.WriteInt8(MapInfo(Map).Reverb) ' New reverb property
    Call Writer.WriteBool(MapInfo(Map).CraftingStoreAllowed)
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Writes the "PosUpdate" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePosUpdate(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PosUpdate" message to the given user's outgoing data buffer
'***************************************************

        Call Writer.WriteInt8(ServerPacketID.PosUpdate)
        Call Writer.WriteInt8(UserList(UserIndex).Pos.X)
        Call Writer.WriteInt8(UserList(UserIndex).Pos.Y)
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Writes the "ChatOverHead" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @param    CharIndex The character uppon which the chat will be displayed.
' @param    Color The color to be used when displaying the chat.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChatOverHead(ByVal UserIndex As Integer, ByVal Chat As String, ByVal CharIndex As Integer, ByVal color As Long)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChatOverHead" message to the given user's outgoing data buffer
'***************************************************

    Call SendData(ToUser, UserIndex, PrepareMessageChatOverHead(Chat, CharIndex, color))

End Sub
Public Sub WriteChatPersonalizado(ByVal UserIndex As Integer, ByVal Chat As String, ByVal CharIndex As Integer, ByVal Tipo As Byte)
'***************************************************
'Author: Juan Dalmasso (CHOTS)
'Last Modification: 11/06/2011
'Writes the "ChatPersonalizado" message to the given user's outgoing data buffer
'***************************************************

    Call SendData(ToUser, UserIndex, PrepareMessageChatPersonalizado(Chat, CharIndex, Tipo))
 
End Sub

''
' Writes the "ConsoleMsg" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @param    FontIndex Index of the FONTTYPE structure to use.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteConsoleMsg(ByVal UserIndex As Integer, ByVal Chat As String, ByVal FontIndex As FontTypeNames, Optional ByVal MessageType As eMessageType = info)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ConsoleMsg" message to the given user's outgoing data buffer
'***************************************************

    Call SendData(ToUser, UserIndex, PrepareMessageConsoleMsg(Chat, FontIndex, MessageType))

End Sub

Public Sub WriteCommerceChat(ByVal UserIndex As Integer, ByVal Chat As String, ByVal FontIndex As FontTypeNames)
'***************************************************
'Author: ZaMa
'Last Modification: 05/17/06
'Writes the "ConsoleMsg" message to the given user's outgoing data buffer
'***************************************************

    Call SendData(ToUser, UserIndex, PrepareCommerceConsoleMsg(Chat, FontIndex))
  
End Sub

''
' Writes the "ShowMessageBox" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Message Text to be displayed in the message box.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowMessageBox(ByVal UserIndex As Integer, ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowMessageBox" message to the given user's outgoing data buffer
'***************************************************

        Call Writer.WriteInt8(ServerPacketID.ShowMessageBox)
        Call Writer.WriteString8(Message)
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Writes the "UserIndexInServer" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserIndexInServer(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserIndexInServer" message to the given user's outgoing data buffer
'***************************************************

        Call Writer.WriteInt8(ServerPacketID.UserIndexInServer)
        Call Writer.WriteInt16(UserIndex)
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Writes the "UserCharIndexInServer" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCharIndexInServer(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserIndexInServer" message to the given user's outgoing data buffer
'***************************************************

        Call Writer.WriteInt8(ServerPacketID.UserCharIndexInServer)
        Call Writer.WriteInt16(UserList(UserIndex).Char.CharIndex)
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Writes the "CharacterCreate" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    X X coord of the new character's position.
' @param    Y Y coord of the new character's position.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @param    name Name of the new character.
' @param    criminal Determines if the character is a criminal or not.
' @param    privileges Sets if the character is a normal one or any kind of administrative character.
' @param    bHostile Determines if the NPC is hostile or not.
' @param    bMerchant Determines if the NPC is merchant or not.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterCreate(ByVal UserIndex As Integer, ByVal body As Integer, ByVal head As Integer, ByVal heading As eHeading, _
                                ByVal CharIndex As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal Weapon As Integer, ByVal shield As Integer, _
                                ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer, ByVal Name As String, ByVal NickColor As Byte, ByVal Alignment As eCharacterAlignment, _
                                ByVal Privileges As Byte, ByVal bHostile As Boolean, ByVal bMerchant As Boolean, ByVal isSailing As Boolean, Optional ByVal NpcNumber As Integer = 0, _
                                Optional ByVal OverHeadIcon As Integer = 0)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 17/09/14
'Writes the "CharacterCreate" message to the given user's outgoing data buffer
'17/09/14: D'Artagnan - bHostile and bMerchant parameters.
'***************************************************

    Call SendData(ToUser, UserIndex, _
        PrepareMessageCharacterCreate( _
            body, head, heading, CharIndex, X, Y, Weapon, shield, FX, FXLoops, _
            helmet, Name, NickColor, Alignment, Privileges, bHostile, bMerchant, isSailing, NpcNumber, OverHeadIcon _
        ) _
    )

End Sub

''
' Writes the "CharacterRemove" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character to be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterRemove(ByVal UserIndex As Integer, ByVal CharIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CharacterRemove" message to the given user's outgoing data buffer
'***************************************************

    Call SendData(ToUser, UserIndex, PrepareMessageCharacterRemove(CharIndex))

End Sub

''
' Writes the "CharacterMove" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterMove(ByVal UserIndex As Integer, ByVal CharIndex As Integer, ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CharacterMove" message to the given user's outgoing data buffer
'***************************************************

    Call SendData(ToUser, UserIndex, PrepareMessageCharacterMove(CharIndex, X, Y, False))

End Sub

Public Sub WriteForceCharMove(ByVal UserIndex, ByVal Direccion As eHeading)
'***************************************************
'Author: ZaMa
'Last Modification: 26/03/2009
'Writes the "ForceCharMove" message to the given user's outgoing data buffer
'***************************************************

    Call SendData(ToUser, UserIndex, PrepareMessageForceCharMove(Direccion))

End Sub

''
' Writes the "CharacterChange" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterChange(ByVal UserIndex As Integer, ByVal body As Integer, ByVal head As Integer, ByVal heading As eHeading, _
                                ByVal CharIndex As Integer, ByVal Weapon As Integer, ByVal shield As Integer, _
                                ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer, ByVal isSailing As Boolean, _
                                ByVal IsDead As Boolean, ByVal OverHeadIcon As Integer, ByVal Alignment As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CharacterChange" message to the given user's outgoing data buffer
'***************************************************

    Call SendData(ToUser, UserIndex, PrepareMessageCharacterChange(body, head, heading, CharIndex, Weapon, shield, FX, FXLoops, helmet, isSailing, IsDead, OverHeadIcon, Alignment))

End Sub

''
' Writes the "ObjectCreate" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    GrhIndex Grh of the object.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteObjectCreate(ByVal UserIndex As Integer, ByVal GrhIndex As Integer, ByVal X As Byte, ByVal Y As Byte, Optional ByVal Luminous As Boolean = False, _
                                Optional ByVal OffsetX As Integer, Optional ByVal OffsetY As Integer, Optional ByVal LightSize As Integer, _
                                Optional ByVal CanBeTransparent As Boolean = False, Optional ByVal ObjType As Byte = 0, Optional ByVal ObjMetadata As Long = 0)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ObjectCreate" message to the given user's outgoing data buffer
'***************************************************

    Call SendData(ToUser, UserIndex, PrepareMessageObjectCreate(GrhIndex, X, Y, ObjType, ObjMetadata, Luminous, OffsetX, OffsetY, LightSize, CanBeTransparent))

End Sub

''
' Writes the "ObjectDelete" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteObjectDelete(ByVal UserIndex As Integer, ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ObjectDelete" message to the given user's outgoing data buffer
'***************************************************

    Call SendData(ToUser, UserIndex, PrepareMessageObjectDelete(X, Y))

End Sub

''
' Writes the "BlockPosition" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @param    Blocked True if the position is blocked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlockPosition(ByVal UserIndex As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal Blocked As Boolean)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BlockPosition" message to the given user's outgoing data buffer
'***************************************************

        Call Writer.WriteInt8(ServerPacketID.BlockPosition)
        Call Writer.WriteInt8(X)
        Call Writer.WriteInt8(Y)
        Call Writer.WriteBool(Blocked)
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Writes the "PlayMusic" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    midi The midi to be played.
' @param    loops Number of repets for the midi.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePlayMusic(ByVal UserIndex As Integer, ByVal Map As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PlayMusic" message to the given user's outgoing data buffer
'***************************************************

    Call SendData(ToUser, UserIndex, PrepareMessagePlayMusic(Map))

End Sub

''
' Writes the "PlayWave" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    wave The wave to be played.
' @param    X The X position in map coordinates from where the sound comes.
' @param    Y The Y position in map coordinates from where the sound comes.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePlayWave(ByVal UserIndex As Integer, ByVal wave As Byte, ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 08/08/07
'Last Modified by: Rapsodius
'Added X and Y positions for 3D Sounds
'***************************************************

    Call SendData(ToUser, UserIndex, PrepareMessagePlayWave(wave, X, Y))

End Sub

''
' Writes the "PauseToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePauseToggle(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PauseToggle" message to the given user's outgoing data buffer
'***************************************************

    Call SendData(ToUser, UserIndex, PrepareMessagePauseToggle())

End Sub

''
' Writes the "RainToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRainToggle(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RainToggle" message to the given user's outgoing data buffer
'***************************************************

    Call SendData(ToUser, UserIndex, PrepareMessageRainToggle())

End Sub

''
' Writes the "CreateFX" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character upon which the FX will be created.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateFX(ByVal UserIndex As Integer, ByVal CharIndex As Integer, ByVal FX As Integer, ByVal FXLoops As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CreateFX" message to the given user's outgoing data buffer
'***************************************************

    Call SendData(ToUser, UserIndex, PrepareMessageCreateFX(CharIndex, FX, FXLoops))

End Sub

''
' Writes the "UpdateUserStats" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateUserStats(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateUserStats" message to the given user's outgoing data buffer
'***************************************************

        Call Writer.WriteInt8(ServerPacketID.UpdateUserStats)
        Call Writer.WriteInt16(UserList(UserIndex).Stats.MaxHp)
        Call Writer.WriteInt16(UserList(UserIndex).Stats.MinHp)
        Call Writer.WriteInt16(UserList(UserIndex).Stats.MaxMan)
        Call Writer.WriteInt16(UserList(UserIndex).Stats.MinMAN)
        Call Writer.WriteInt16(UserList(UserIndex).Stats.MaxSta)
        Call Writer.WriteInt16(UserList(UserIndex).Stats.MinSta)
        Call Writer.WriteInt32(UserList(UserIndex).Stats.GLD)
        Call Writer.WriteInt8(UserList(UserIndex).Stats.ELV)
        Call Writer.WriteInt32(UserList(UserIndex).Stats.ELU)
        Call Writer.WriteInt32(UserList(UserIndex).Stats.Exp)
        Call Writer.WriteBool(UserList(UserIndex).Stats.ELV = ConstantesBalance.MaxLvl)
        Call Writer.WriteInt16(UserList(UserIndex).Stats.MasteryPoints)
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Writes the "WorkRequestTarget" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Skill The skill for which we request a target.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWorkRequestTarget(ByVal UserIndex As Integer, ByVal skill As eSkill)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "WorkRequestTarget" message to the given user's outgoing data buffer
'***************************************************

        Call Writer.WriteInt8(ServerPacketID.WorkRequestTarget)
        Call Writer.WriteInt8(skill)
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Writes the "ChangeInventorySlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Inventory slot which needs to be updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeInventorySlot(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal ObjIndex As Integer, _
    ByVal Amount As Integer, ByVal Equipped As Byte, ByVal CanUse As Boolean)

    
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 25/05/2011 (Amraphen)
'Writes the "ChangeInventorySlot" message to the given user's outgoing data buffer
'3/12/09: Budi - Ahora se envia MaxDef y MinDef en lugar de Def
'25/05/2011: Amraphen - Ahora se envía la defensa según se tiene equipado armadura de segunda jerarquía o no.
'***************************************************

        Call Writer.WriteInt8(ServerPacketID.ChangeInventorySlot)
        Call Writer.WriteInt8(Slot)
        Call Writer.WriteInt16(ObjIndex)
        Call Writer.WriteInt16(Amount)
        Call Writer.WriteBool(Equipped)
        Call Writer.WriteBool(CanUse)
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

Public Sub WriteAddSlots(ByVal UserIndex As Integer, ByVal Mochila As eMochilas)
'***************************************************
'Author: Budi
'Last Modification: 01/12/09
'Writes the "AddSlots" message to the given user's outgoing data buffer
'***************************************************

        Call Writer.WriteInt8(ServerPacketID.AddSlots)
        Call Writer.WriteInt8(Mochila)
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub


''
' Writes the "ChangeBankSlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Inventory slot which needs to be updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeBankSlot(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal ObjIndex As Integer, _
                                                            ByVal Amount As Integer, ByVal CanUse As Boolean)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/03/09
'Writes the "ChangeBankSlot" message to the given user's outgoing data buffer
'12/03/09: Budi - Ahora se envia MaxDef y MinDef en lugar de sólo Def
'03/05/2012: ZaMa - Optimizo el tamaño del paquete (los datos los saca del cliente).
'***************************************************

        Call Writer.WriteInt8(ServerPacketID.ChangeBankSlot)
        Call Writer.WriteInt8(Slot)
        Call Writer.WriteInt16(ObjIndex)
        Call Writer.WriteInt16(Amount)
        Call Writer.WriteBool(CanUse)
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Writes the "ChangeSpellSlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Spell slot to update.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeSpellSlot(ByVal UserIndex As Integer, ByVal Slot As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeSpellSlot" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo ErrHandler:

    With UserList(UserIndex)
        Call Writer.WriteInt8(ServerPacketID.ChangeSpellSlot)
        Call Writer.WriteInt(Slot)
        Call Writer.WriteInt(.Stats.UserHechizos(Slot).SpellNumber)
        
        If .Stats.UserHechizos(Slot).SpellNumber > 0 Then
            Call Writer.WriteString8(Hechizos(.Stats.UserHechizos(Slot).SpellNumber).Nombre)
            Call Writer.WriteInt(Hechizos(.Stats.UserHechizos(Slot).SpellNumber).SpellCastInterval)
        Else
            Call Writer.WriteString8("(None)")
            Call Writer.WriteInt(0)
        End If
        
        
    End With
    
    Call SendData(ToUser, UserIndex, vbNullString)
    Exit Sub
ErrHandler:
    Debug.Print Err.Description
End Sub

''
' Writes the "Atributes" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAttributes(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Atributes" message to the given user's outgoing data buffer
'***************************************************

    With UserList(UserIndex)
        Call Writer.WriteInt8(ServerPacketID.Atributes)
        Call Writer.WriteInt8(.Stats.UserAtributos(eAtributos.Fuerza))
        Call Writer.WriteInt8(.Stats.UserAtributos(eAtributos.Agilidad))
        Call Writer.WriteInt8(.Stats.UserAtributos(eAtributos.Inteligencia))
        Call Writer.WriteInt8(.Stats.UserAtributos(eAtributos.Carisma))
        Call Writer.WriteInt8(.Stats.UserAtributos(eAtributos.Constitucion))
    End With
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub


Public Sub WriteCraftableRecipes(ByVal UserIndex As Integer, ByVal Profession As Byte)
    Dim I As Integer, J  As Integer, k As Integer

    ' Stop if there's no profession configured for the item, or the profession has no recipe groups
    If Profession <= 0 Then Exit Sub
    If Professions(Profession).CraftingRecipeGroupsQty <= 0 Then Exit Sub
        
    Dim ValidElementIndexes() As Integer
    Dim ValidElementIndexesQty As Integer
    Dim CurrentIndex As Integer
    
    Dim BlacksmithSkills As Byte
    Dim CarpentrySkills As Byte
    Dim TailoringSkills As Byte
    
    
    With UserList(UserIndex)
        BlacksmithSkills = GetSkills(UserIndex, eSkill.Herreria)
        CarpentrySkills = GetSkills(UserIndex, eSkill.Carpinteria)
        TailoringSkills = GetSkills(UserIndex, eSkill.Sastreria)
    End With
    
    With Professions(Profession)
    
        Call Writer.WriteInt8(ServerPacketID.CraftableRecipes)
        Call Writer.WriteInt(.CraftingRecipeGroupsQty)
    
        For I = 1 To .CraftingRecipeGroupsQty
            ValidElementIndexesQty = 0
            
            If .CraftingRecipeGroups(I).RecipesQty >= 0 Then
                ReDim ValidElementIndexes(1 To .CraftingRecipeGroups(I).RecipesQty)
                
                For J = 1 To .CraftingRecipeGroups(I).RecipesQty
                    If BlacksmithSkills >= .CraftingRecipeGroups(I).Recipes(J).BlacksmithSkillNeeded And _
                        CarpentrySkills >= .CraftingRecipeGroups(I).Recipes(J).CarpenterSkillNeeded And _
                        TailoringSkills >= .CraftingRecipeGroups(I).Recipes(J).TailoringSkillNeeded Then
                        
                        ValidElementIndexesQty = ValidElementIndexesQty + 1
                        ValidElementIndexes(ValidElementIndexesQty) = J
                    End If
                Next J
                
                If ValidElementIndexesQty > 0 Then
                    ReDim Preserve ValidElementIndexes(1 To ValidElementIndexesQty)
                End If
             
                Call Writer.WriteString16(.CraftingRecipeGroups(I).TabTitle)
                Call Writer.WriteString16(.CraftingRecipeGroups(I).TabImage)
                Call Writer.WriteInt(.CraftingRecipeGroups(I).ProfessionType)
                Call Writer.WriteInt(ValidElementIndexesQty)
                
                For J = 1 To ValidElementIndexesQty
                    CurrentIndex = ValidElementIndexes(J)
                    
                    Call Writer.WriteInt(CurrentIndex)
                    Call Writer.WriteInt(.CraftingRecipeGroups(I).Recipes(CurrentIndex).ObjIndex)
                    Call Writer.WriteInt(.CraftingRecipeGroups(I).Recipes(CurrentIndex).MaterialsQty)
                    
                    For k = 1 To .CraftingRecipeGroups(I).Recipes(CurrentIndex).MaterialsQty
                        Call Writer.WriteInt(.CraftingRecipeGroups(I).Recipes(CurrentIndex).Materials(k).ObjIndex)
                        Call Writer.WriteInt(.CraftingRecipeGroups(I).Recipes(CurrentIndex).Materials(k).Amount)
                    Next k
                    
                Next J
            
            End If
        Next I
    End With
    Exit Sub
    
End Sub

''
' Writes the "RestOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRestOK(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RestOK" message to the given user's outgoing data buffer
'***************************************************

    Call Writer.WriteInt8(ServerPacketID.RestOK)
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Writes the "ErrorMsg" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    message The error message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteErrorMsg(ByVal UserIndex As Integer, ByVal Message As String, _
                         Optional ByVal bCloseConnection As Boolean = True)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 07/14/2014 (D'Artagnan)
'Writes the "ErrorMsg" Message to the given user's outgoing data buffer
'07/14/2014: D'Artagnan - New optional parameter: bCloseConnection.
'***************************************************

        Call SendData(ToUser, UserIndex, PrepareMessageErrorMsg(Message))


    
End Sub

''
' Writes the "Blind" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlind(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Blind" message to the given user's outgoing data buffer
'***************************************************

    Call Writer.WriteInt8(ServerPacketID.Blind)
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Writes the "Dumb" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDumb(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Dumb" message to the given user's outgoing data buffer
'***************************************************

    Call Writer.WriteInt8(ServerPacketID.Dumb)
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Writes the "ShowSignal" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    objIndex Index of the signal to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowSignal(ByVal UserIndex As Integer, ByVal ObjIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowSignal" message to the given user's outgoing data buffer
'***************************************************

    
        Call Writer.WriteInt8(ServerPacketID.ShowSignal)
        Call Writer.WriteString8(ObjData(ObjIndex).texto)
        Call Writer.WriteInt16(ObjData(ObjIndex).GrhSecundario)
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Writes the "ChangeNPCInventorySlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex   User to which the message is intended.
' @param    slot        The inventory slot in which this item is to be placed.
' @param    obj         The object to be set in the NPC's inventory window.
' @param    price       The value the NPC asks for the object.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeNPCInventorySlot(ByVal UserIndex As Integer, ByVal Slot As Byte, _
                                                                          ByVal Amount As Integer, ByVal Price As Single, _
                                                                          ByVal ObjIndex As Integer, ByVal CanUse As Boolean)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 03/05/2012
'Last Modified by: Budi
'Writes the "ChangeNPCInventorySlot" message to the given user's outgoing data buffer
'12/03/09: Budi - Ahora se envia MaxDef y MinDef en lugar de sólo Def
'03/05/2012: ZaMa - Aliviano el paquete y envío menos info (lo demas lo saca desde el cliente).
'***************************************************

    
        Call Writer.WriteInt8(ServerPacketID.ChangeNPCInventorySlot)
        Call Writer.WriteInt8(Slot)
        Call Writer.WriteInt16(Amount)
        Call Writer.WriteReal32(Price)
        Call Writer.WriteInt16(ObjIndex)
        Call Writer.WriteBool(CanUse)
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

Public Sub WriteSetUserSalePrice(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Price As Long)

        Call Writer.WriteInt8(ServerPacketID.SetUserSalePrice)
        Call Writer.WriteInt8(Slot)
        Call Writer.WriteInt32(Price)
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Writes the "UpdateHungerAndThirst" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateHungerAndThirst(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateHungerAndThirst" message to the given user's outgoing data buffer
'***************************************************

    
        Call Writer.WriteInt8(ServerPacketID.UpdateHungerAndThirst)
        Call Writer.WriteInt8(UserList(UserIndex).Stats.MaxAGU)
        Call Writer.WriteInt8(UserList(UserIndex).Stats.MinAGU)
        Call Writer.WriteInt8(UserList(UserIndex).Stats.MaxHam)
        Call Writer.WriteInt8(UserList(UserIndex).Stats.MinHam)
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub


''
' Writes the "MiniStats" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMiniStats(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "MiniStats" message to the given user's outgoing data buffer
'***************************************************

    
        Call Writer.WriteInt8(ServerPacketID.MiniStats)
        
        Call Writer.WriteInt32(UserList(UserIndex).Faccion.CiudadanosMatados)
        Call Writer.WriteInt32(UserList(UserIndex).Faccion.CriminalesMatados)
        
'TODO : Este valor es calculable, no debería NI EXISTIR, ya sea en el servidor ni en el cliente!!!
        Call Writer.WriteInt32(UserList(UserIndex).Stats.UsuariosMatados)
        
        Call Writer.WriteInt32(UserList(UserIndex).Stats.NPCsMuertos)
        
        Call Writer.WriteInt8(UserList(UserIndex).clase)
        Call Writer.WriteInt32(UserList(UserIndex).Counters.Pena)
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Writes the "LevelUp" message to the given user's outgoing data buffer.
'
' @param    skillPoints The number of free skill points the player has.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLevelUp(ByVal UserIndex As Integer, ByVal skillPoints As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "LevelUp" message to the given user's outgoing data buffer
'***************************************************

    
        Call Writer.WriteInt8(ServerPacketID.LevelUp)
        Call Writer.WriteInt16(skillPoints)
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Writes the "AddForumMsg" message to the given user's outgoing data buffer.
'
' @param    title The title of the message to display.
' @param    message The message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAddForumMsg(ByVal UserIndex As Integer, ByVal ForumType As eForumType, _
                    ByRef Title As String, ByRef Author As String, ByRef Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 02/01/2010
'Writes the "AddForumMsg" message to the given user's outgoing data buffer
'02/01/2010: ZaMa - Now sends Author and forum type
'***************************************************

        Call Writer.WriteInt8(ServerPacketID.AddForumMsg)
        Call Writer.WriteInt8(ForumType)
        Call Writer.WriteString8(Title)
        Call Writer.WriteString8(Author)
        Call Writer.WriteString8(Message)
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

Public Sub WriteShowGuildCreate(ByVal UserIndex As Integer)
    Dim Faccion As eGuildAlignment
    
    If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
        Faccion = eGuildAlignment.Real
    ElseIf UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
        Faccion = eGuildAlignment.Evil
    Else
        Faccion = eGuildAlignment.Neutral
    End If
    
    Call Writer.WriteInt8(ServerPacketID.ShowGuildCreate)
    Call Writer.WriteInt8(Faccion)
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub


Public Sub WriteShowGuildForm(ByVal UserIndex As Integer)

    Call Writer.WriteInt8(ServerPacketID.ShowGuildForm)
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Writes the "ShowForumForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowForumForm(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowForumForm" message to the given user's outgoing data buffer
'***************************************************

    Dim Visibilidad As Byte
    Dim CanMakeSticky As Byte
    
    With UserList(UserIndex)
        Call Writer.WriteInt8(ServerPacketID.ShowForumForm)
        
        Visibilidad = eForumVisibility.ieGENERAL_MEMBER
        
        If esCaos(UserIndex) Or EsGm(UserIndex) Then
            Visibilidad = Visibilidad Or eForumVisibility.ieCAOS_MEMBER
        End If
        
        If esArmada(UserIndex) Or EsGm(UserIndex) Then
            Visibilidad = Visibilidad Or eForumVisibility.ieREAL_MEMBER
        End If
        
        Call Writer.WriteInt8(Visibilidad)
        
        ' Pueden mandar sticky los gms o los del consejo de armada/caos
        If EsGm(UserIndex) Then
            CanMakeSticky = 2
        ElseIf (.flags.Privilegios And PlayerType.ChaosCouncil) <> 0 Then
            CanMakeSticky = 1
        ElseIf (.flags.Privilegios And PlayerType.RoyalCouncil) <> 0 Then
            CanMakeSticky = 1
        End If
        
        Call Writer.WriteInt8(CanMakeSticky)
    End With
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Writes the "SetInvisible" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex The char turning visible / invisible.
' @param    invisible True if the char is no longer visible, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetInvisible(ByVal UserIndex As Integer, ByVal CharIndex As Integer, ByVal invisible As Boolean, Optional ByVal Transparency As Boolean = False)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SetInvisible" message to the given user's outgoing data buffer
'***************************************************

    Call SendData(ToUser, UserIndex, PrepareMessageSetInvisible(CharIndex, invisible, Transparency))

End Sub

''
' Writes the "MeditateToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMeditateToggle(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "MeditateToggle" message to the given user's outgoing data buffer
'***************************************************

    Call Writer.WriteInt8(ServerPacketID.MeditateToggle)
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Writes the "BlindNoMore" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlindNoMore(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BlindNoMore" message to the given user's outgoing data buffer
'***************************************************

    Call Writer.WriteInt8(ServerPacketID.BlindNoMore)
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Writes the "DumbNoMore" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDumbNoMore(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "DumbNoMore" message to the given user's outgoing data buffer
'***************************************************

    Call Writer.WriteInt8(ServerPacketID.DumbNoMore)
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Writes the "SendSkills" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSendSkills(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 28/04/2015
'Writes the "SendSkills" message to the given user's outgoing data buffer
'11/19/09: Pato - Now send the percentage of progress of the skills.
'28/04/2015: D'Artagnan - Send both natural and assigned skills.
'***************************************************


    Dim lSkill As Long
    Dim Percentage As Integer
    
    With UserList(UserIndex)
        Call Writer.WriteInt8(ServerPacketID.SendSkills)
        
        For lSkill = 1 To NUMSKILLS
            Call Writer.WriteInt8(GetNaturalSkills(UserIndex, lSkill))
            Call Writer.WriteInt8(GetAssignedSkills(UserIndex, lSkill))
            
            If GetSkills(UserIndex, lSkill) < MAX_SKILL_POINTS Then
                Percentage = Int(.Stats.ExpSkills(lSkill) * 100 / .Stats.EluSkills(lSkill))
                If Percentage > 100 Then Percentage = 100

                Call Writer.WriteInt8(Percentage)

            Else
                Call Writer.WriteInt8(0)
            End If
        Next lSkill
    End With
        
    Call SendData(ToUser, UserIndex, vbNullString)

End Sub

''
' Writes the "TrainerCreatureList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    npcIndex The index of the requested trainer.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTrainerCreatureList(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TrainerCreatureList" message to the given user's outgoing data buffer
'***************************************************

    Dim I As Long
    Dim str As String
    
    
        Call Writer.WriteInt8(ServerPacketID.TrainerCreatureList)
        
        For I = 1 To Npclist(NpcIndex).NroCriaturas
            str = str & Npclist(NpcIndex).Criaturas(I).NpcName & SEPARATOR
        Next I
        
        If LenB(str) > 0 Then _
            str = Left$(str, Len(str) - 1)
        
        Call Writer.WriteString8(str)
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Writes the "CharacterInfo" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    charName The requested char's name.
' @param    race The requested char's race.
' @param    class The requested char's class.
' @param    gender The requested char's gender.
' @param    level The requested char's level.
' @param    gold The requested char's gold.
' @param    reputation The requested char's reputation.
' @param    previousPetitions The requested char's previous petitions to enter guilds.
' @param    currentGuild The requested char's current guild.
' @param    previousGuilds The requested char's previous guilds.
' @param    currentPetitions The requested char's current active guild requests.
' @param    RoyalArmy True if tha char belongs to the Royal Army.
' @param    CaosLegion True if tha char belongs to the Caos Legion.
' @param    citicensKilled The number of citicens killed by the requested char.
' @param    criminalsKilled The number of criminals killed by the requested char.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterInfo(ByVal UserIndex As Integer, ByVal charName As String, ByVal race As eRaza, ByVal Class As eClass, _
    ByVal gender As eGenero, ByVal level As Byte, ByVal Gold As Long, ByVal Bank As Long, ByVal reputation As Long, _
    ByVal previousPetitions As String, ByVal CurrentGuild As String, ByVal previousGuilds As String, ByVal currentPetitions As String, ByVal RoyalArmy As Boolean, _
    ByVal CaosLegion As Boolean, ByVal citicensKilled As Long, ByVal criminalsKilled As Long, _
    ByVal RankingPoints As Long)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/08/06
'12/08/2014 thesword: adding current petitions list to the parameters
'Writes the "CharacterInfo" message to the given user's outgoing data buffer
'***************************************************

    
        Call Writer.WriteInt8(ServerPacketID.CharacterInfo)
        
        Call Writer.WriteString8(charName)
        Call Writer.WriteInt8(race)
        Call Writer.WriteInt8(Class)
        Call Writer.WriteInt8(gender)
        
        Call Writer.WriteInt8(level)
        Call Writer.WriteInt32(Gold)
        Call Writer.WriteInt32(Bank)
        Call Writer.WriteInt32(reputation)
        
        Call Writer.WriteString8(previousPetitions)
        Call Writer.WriteString8(CurrentGuild)
        Call Writer.WriteString8(previousGuilds)
        Call Writer.WriteString8(currentPetitions)
        
        Call Writer.WriteBool(RoyalArmy)
        Call Writer.WriteBool(CaosLegion)
        
        Call Writer.WriteInt32(citicensKilled)
        Call Writer.WriteInt32(criminalsKilled)
        
        Call Writer.WriteInt32(RankingPoints)
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Writes the "ParalizeOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteParalizeOK(ByVal UserIndex As Integer, Optional ByVal UpdatePos As Boolean = True)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 08/12/07
'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
'Writes the "ParalizeOK" message to the given user's outgoing data buffer
'And updates user position
'***************************************************

    Call Writer.WriteInt8(ServerPacketID.ParalizeOK)
    Call SendData(ToUser, UserIndex, vbNullString)
    
    If UpdatePos Then Call WritePosUpdate(UserIndex)

End Sub

''
' Writes the "ShowUserRequest" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    details DEtails of the char's request.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowUserRequest(ByVal UserIndex As Integer, ByVal Details As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowUserRequest" message to the given user's outgoing data buffer
'***************************************************

    
        Call Writer.WriteInt8(ServerPacketID.ShowUserRequest)
        
        Call Writer.WriteString8(Details)
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Writes the "TradeOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTradeOK(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TradeOK" message to the given user's outgoing data buffer
'***************************************************

    Call Writer.WriteInt8(ServerPacketID.TradeOK)
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Writes the "BankOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankOK(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankOK" message to the given user's outgoing data buffer
'***************************************************

    Call Writer.WriteInt8(ServerPacketID.BankOK)
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Writes the "ChangeUserTradeSlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    ObjIndex The object's index.
' @param    amount The number of objects offered.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeUserTradeSlot(ByVal UserIndex As Integer, ByVal OfferSlot As Byte, ByVal ObjIndex As Integer, _
                                    ByVal Amount As Long, ByVal CanUse As Boolean)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/03/09
'Writes the "ChangeUserTradeSlot" message to the given user's outgoing data buffer
'25/11/2009: ZaMa - Now sends the specific offer slot to be modified.
'12/03/09: Budi - Ahora se envia MaxDef y MinDef en lugar de sólo Def
'08/05/2012: ZaMa - Reduzco el tamaño del paquete (lo maneja todo el cliente ahora).
'***************************************************

    
        Call Writer.WriteInt8(ServerPacketID.ChangeUserTradeSlot)
        
        Call Writer.WriteInt8(OfferSlot)
        Call Writer.WriteInt16(ObjIndex)
        Call Writer.WriteInt32(Amount)
        Call Writer.WriteBool(CanUse)
        
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Writes the "WriteChangeUserTradeGold" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Amount The number of objects offered.
Public Sub WriteChangeUserTradeGold(ByVal UserIndex As Integer, ByVal Amount As Long)
    
    Call Writer.WriteInt8(ServerPacketID.ChangeUserTradeGold)
    
    Call Writer.WriteInt32(Amount)
        
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Writes the "SpawnList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    npcNames The names of the creatures that can be spawned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpawnList(ByVal UserIndex As Integer, ByRef npcNames() As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SpawnList" message to the given user's outgoing data buffer
'***************************************************

    Dim I As Long
    Dim Tmp As String
    
    
        Call Writer.WriteInt8(ServerPacketID.SpawnList)
        
        For I = LBound(npcNames()) To UBound(npcNames())
            Tmp = Tmp & npcNames(I) & SEPARATOR
        Next I
        
        If Len(Tmp) Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call Writer.WriteString8(Tmp)
        
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Writes the "ShowSOSForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowSOSForm(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowSOSForm" message to the given user's outgoing data buffer
'***************************************************

    Dim I As Long
    Dim Tmp As String
    
    
        Call Writer.WriteInt8(ServerPacketID.ShowSOSForm)
        
        For I = 1 To Ayuda.Longitud
            Tmp = Tmp & Ayuda.VerElemento(I) & SEPARATOR
        Next I
        
        If LenB(Tmp) <> 0 Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call Writer.WriteString8(Tmp)
       
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Writes the "ShowDenounces" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowDenounces(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 14/11/2010
'Writes the "ShowDenounces" message to the given user's outgoing data buffer
'***************************************************

    Dim DenounceIndex As Long
    Dim DenounceList As String
    
    
        Call Writer.WriteInt8(ServerPacketID.ShowDenounces)
        
        For DenounceIndex = 1 To Denuncias.Longitud
            DenounceList = DenounceList & Denuncias.VerElemento(DenounceIndex, False) & SEPARATOR
        Next DenounceIndex
        
        If LenB(DenounceList) <> 0 Then _
            DenounceList = Left$(DenounceList, Len(DenounceList) - 1)
        
        Call Writer.WriteString8(DenounceList)
       
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Writes the "ShowSOSForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowPartyForm(ByVal UserIndex As Integer)
'***************************************************
'Author: Budi
'Last Modification: 11/26/09
'Writes the "ShowPartyForm" message to the given user's outgoing data buffer
'***************************************************

    Dim I As Long
    Dim Tmp As String
    Dim PI As Integer
    Dim Members() As Integer
    ReDim Members(Constantes.MaxPartyMembers) As Integer
    
    
        Call Writer.WriteInt8(ServerPacketID.ShowPartyForm)
        
        PI = UserList(UserIndex).PartyIndex
        Call Writer.WriteInt8(CByte(Parties(PI).EsPartyLeader(UserIndex)))
        
        If PI > 0 Then
            Call Parties(PI).ObtenerMiembrosOnline(Members())
            For I = 1 To Constantes.MaxPartyMembers
                If Members(I) > 0 Then
                    Tmp = Tmp & UserList(Members(I)).Name & " (" & Fix(Parties(PI).MiExperiencia(Members(I))) & ")" & SEPARATOR
                End If
            Next I
        End If
        
        If LenB(Tmp) <> 0 Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
            
        Call Writer.WriteString8(Tmp)
        Call Writer.WriteInt32(Parties(PI).ObtenerExperienciaTotal)
       
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Writes the "ShowMOTDEditionForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    currentMOTD The current Message Of The Day.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowMOTDEditionForm(ByVal UserIndex As Integer, ByVal currentMOTD As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowMOTDEditionForm" message to the given user's outgoing data buffer
'***************************************************

    
        Call Writer.WriteInt8(ServerPacketID.ShowMOTDEditionForm)
        
        Call Writer.WriteString8(currentMOTD)
       
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Writes the "ShowGMPanelForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowGMPanelForm(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowGMPanelForm" message to the given user's outgoing data buffer
'***************************************************

    Call Writer.WriteInt8(ServerPacketID.ShowGMPanelForm)
       
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Writes the "UserNameList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    userNameList List of user names.
' @param    Cant Number of names to send.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserNameList(ByVal UserIndex As Integer, ByRef userNamesList() As String, ByVal cant As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06 NIGO:
'Writes the "UserNameList" message to the given user's outgoing data buffer
'***************************************************

    Dim I As Long
    Dim Tmp As String
    
    
        Call Writer.WriteInt8(ServerPacketID.UserNameList)
        
        ' Prepare user's names list
        For I = 1 To cant
            Tmp = Tmp & userNamesList(I) & SEPARATOR
        Next I
        
        If Len(Tmp) Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call Writer.WriteString8(Tmp)
        
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Writes the "Pong" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePong(ByVal UserIndex As Integer, ByVal Tick As Long)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Pong" message to the given user's outgoing data buffer
'***************************************************

    Call Writer.WriteInt8(ServerPacketID.Pong)
    Call Writer.WriteInt32(Tick)
       
    Call SendData(ToUser, UserIndex, vbNullString, , True)
        
End Sub

''
' Prepares the "SetInvisible" message and returns it.
'
' @param    CharIndex The char turning visible / invisible.
' @param    invisible True if the char is no longer visible, False otherwise.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The message is written to no outgoing buffer, but only prepared in a single string to be easily sent to several clients.

Public Function PrepareMessageSetInvisible(ByVal CharIndex As Integer, ByVal invisible As Boolean, Optional ByVal Transparency As Boolean = False) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "SetInvisible" message and returns it.
'***************************************************
On Error GoTo ErrHandler

        Call Writer.WriteInt8(ServerPacketID.SetInvisible)
        
        Call Writer.WriteInt16(CharIndex)
        Call Writer.WriteBool(invisible)
        Call Writer.WriteBool(Transparency)

  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function PrepareMessageSetInvisible de Protocol.bas")
End Function

Public Function PrepareMessageCharacterChangeNick(ByVal CharIndex As Integer, ByVal newNick As String) As String
'***************************************************
'Author: Budi
'Last Modification: 07/23/09
'Prepares the "Change Nick" message and returns it.
'***************************************************
On Error GoTo ErrHandler

        Call Writer.WriteInt8(ServerPacketID.CharacterChangeNick)
        
        Call Writer.WriteInt16(CharIndex)
        Call Writer.WriteString8(newNick)
        
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function PrepareMessageCharacterChangeNick de Protocol.bas")
End Function

''
' Prepares the "ChatOverHead" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @param    CharIndex The character uppon which the chat will be displayed.
' @param    Color The color to be used when displaying the chat.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The message is written to no outgoing buffer, but only prepared in a single string to be easily sent to several clients.

Public Function PrepareMessageChatOverHead(ByVal Chat As String, ByVal CharIndex As Integer, ByVal color As Long) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "ChatOverHead" message and returns it.
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ServerPacketID.ChatOverHead)
        Call Writer.WriteString8(Chat)
        Call Writer.WriteInt16(CharIndex)
        
        ' Write rgb channels and save one byte from long :D
        Call Writer.WriteInt8(color And &HFF)
        Call Writer.WriteInt8((color And &HFF00&) \ &H100&)
        Call Writer.WriteInt8((color And &HFF0000) \ &H10000)

  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function PrepareMessageChatOverHead de Protocol.bas")
End Function

Public Function PrepareMessageChatPersonalizado(ByVal Chat As String, ByVal CharIndex As Integer, ByVal Tipo As Byte) As String
'***************************************************
'Author: Juan Dalmasso (CHOTS)
'Last Modification: 11/06/2011
'Prepares the "ChatPersonalizado" message and returns it.
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ServerPacketID.ChatPersonalizado)
        Call Writer.WriteString8(Chat)
        Call Writer.WriteInt16(CharIndex)
        
        ' Write the type of message
        '1=normal
        '2=clan
        '3=party
        '4=gritar
        '5=palabras magicas
        '6=susurrar
        Call Writer.WriteInt8(Tipo)

  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function PrepareMessageChatPersonalizado de Protocol.bas")
End Function

''
' Prepares the "ConsoleMsg" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @param    FontIndex Index of the FONTTYPE structure to use.
' @param    MessageType type of console message (General, Guild, Party)
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageConsoleMsg(ByVal Chat As String, ByVal FontIndex As FontTypeNames, Optional ByVal MessageType As eMessageType = info) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/05/11 (D'Artagnan)
'Prepares the "MessageType" message and returns it.
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ServerPacketID.ConsoleMsg)
        Call Writer.WriteString8(Chat)
        Call Writer.WriteInt8(FontIndex)
        Call Writer.WriteInt8(MessageType)

  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function PrepareMessageConsoleMsg de Protocol.bas")
End Function

Public Function PrepareCommerceConsoleMsg(ByRef Chat As String, ByVal FontIndex As FontTypeNames) As String
'***************************************************
'Author: ZaMa
'Last Modification: 03/12/2009
'Prepares the "CommerceConsoleMsg" message and returns it.
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ServerPacketID.CommerceChat)
        Call Writer.WriteString8(Chat)
        Call Writer.WriteInt8(FontIndex)

  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function PrepareCommerceConsoleMsg de Protocol.bas")
End Function

''
' Prepares the "CreateFX" message and returns it.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character upon which the FX will be created.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    IsGlow Determines if the FX sent is a glow effect or a normal one.
'           Glow effects are drawed in a different layer
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCreateFX(ByVal CharIndex As Integer, ByVal FX As Integer, ByVal FXLoops As Integer, Optional ByVal IsGlowEffect As Boolean = False, Optional ByVal Slot As Byte = 255) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "CreateFX" message and returns it
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ServerPacketID.CreateFX)
        Call Writer.WriteInt16(CharIndex)
        Call Writer.WriteInt16(FX)
        Call Writer.WriteInt16(FXLoops)
        Call Writer.WriteInt8(Slot)
        Call Writer.WriteBool(IsGlowEffect)

  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function PrepareMessageCreateFX de Protocol.bas")
End Function

''
' Prepares the "PlayWave" message and returns it.
'
' @param    wave The wave to be played.
' @param    X The X position in map coordinates from where the sound comes.
' @param    Y The Y position in map coordinates from where the sound comes.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePlayWave(ByVal wave As Integer, ByVal X As Byte, ByVal Y As Byte, Optional ByVal Entity As Long = 0) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 08/08/07
'Last Modified by: Rapsodius
'Added X and Y positions for 3D Sounds
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ServerPacketID.PlayEffect)
        Call Writer.WriteInt16(wave)
        Call Writer.WriteInt8(X)
        Call Writer.WriteInt8(Y)
        Call Writer.WriteInt(Entity)
        
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function PrepareMessagePlayWave de Protocol.bas")
End Function

''
' Prepares the "GuildChat" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageGuildChat(ByVal Chat As String, Optional ByVal IsMOTD As Boolean = False) As String

On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ServerPacketID.GuildChat)
        Call Writer.WriteString8(Chat)
        Call Writer.WriteBool(IsMOTD)
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function PrepareMessageGuildChat de Protocol.bas")
End Function

''
' Prepares the "ShowMessageBox" message and returns it.
'
' @param    Message Text to be displayed in the message box.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageShowMessageBox(ByVal Chat As String) As String
'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 01/08/07
'Prepares the "ShowMessageBox" message and returns it
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ServerPacketID.ShowMessageBox)
        Call Writer.WriteString8(Chat)

  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function PrepareMessageShowMessageBox de Protocol.bas")
End Function


''
' Prepares the "PlayMusic" message and returns it.
'
' @param    Map The map where the Music information will be extracted from
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePlayMusic(ByVal Map As Integer) As String
On Error GoTo ErrHandler
    
    If MapInfo(Map).NumMusic = 0 Then Exit Function
        
    Call Writer.WriteInt8(ServerPacketID.PlayMusic)
    Call Writer.WriteSafeArrayInt32(MapInfo(Map).Music)

  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function PrepareMessagePlayMusic de Protocol.bas")
End Function

''
' Prepares the "PauseToggle" message and returns it.
'
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePauseToggle() As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "PauseToggle" message and returns it
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ServerPacketID.PauseToggle)

  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function PrepareMessagePauseToggle de Protocol.bas")
End Function

''
' Prepares the "RainToggle" message and returns it.
'
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageRainToggle() As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "RainToggle" message and returns it
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ServerPacketID.RainToggle)

  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function PrepareMessageRainToggle de Protocol.bas")
End Function

''
' Prepares the "ObjectDelete" message and returns it.
'
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageObjectDelete(ByVal X As Byte, ByVal Y As Byte) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "ObjectDelete" message and returns it
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ServerPacketID.ObjectDelete)
        Call Writer.WriteInt8(X)
        Call Writer.WriteInt8(Y)

  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function PrepareMessageObjectDelete de Protocol.bas")
End Function

''
' Prepares the "ObjectUpdate" message and returns it.

Public Function PrepareMessageObjectUpdate(ByVal X As Byte, ByVal Y As Byte, ByVal GrhIndex As Integer, ByVal ObjType As Long, Optional ByVal ObjMetadata As Long = 0) As String

On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ServerPacketID.ObjectUpdate)
        Call Writer.WriteInt8(X)
        Call Writer.WriteInt8(Y)
        Call Writer.WriteInt(GrhIndex)
        Call Writer.WriteInt(ObjType)
        Call Writer.WriteInt(ObjMetadata)

  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function PrepareMessageObjectUpdate de Protocol.bas")
End Function

''
' Prepares the "BlockPosition" message and returns it.
'
' @param    X X coord of the tile to block/unblock.
' @param    Y Y coord of the tile to block/unblock.
' @param    Blocked Blocked status of the tile
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageBlockPosition(ByVal X As Byte, ByVal Y As Byte, ByVal Blocked As Boolean) As String
'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 01/08/07
'Prepares the "BlockPosition" message and returns it
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ServerPacketID.BlockPosition)
        Call Writer.WriteInt8(X)
        Call Writer.WriteInt8(Y)
        Call Writer.WriteBool(Blocked)

  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function PrepareMessageBlockPosition de Protocol.bas")
End Function

''
' Prepares the "ObjectCreate" message and returns it.
'
' @param    GrhIndex Grh of the object.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageObjectCreate(ByVal GrhIndex As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal ObjType As Byte, _
                                            Optional ByVal ObjMetadata As Long = 0, Optional ByVal Luminous As Boolean = False, _
                                            Optional ByVal OffsetX As Integer, Optional ByVal OffsetY As Integer, _
                                            Optional ByVal LightSize As Integer, Optional ByVal CanBeTransparent As Boolean = False) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'prepares the "ObjectCreate" message and returns it
'***************************************************
On Error GoTo ErrHandler

        Call Writer.WriteInt8(ServerPacketID.ObjectCreate)
        Call Writer.WriteInt8(X)
        Call Writer.WriteInt8(Y)
        Call Writer.WriteInt16(GrhIndex)
        Call Writer.WriteBool(Luminous)
        Call Writer.WriteInt16(OffsetX)
        Call Writer.WriteInt16(OffsetY)
        Call Writer.WriteInt16(LightSize)
        Call Writer.WriteBool(CanBeTransparent)
        Call Writer.WriteInt8(ObjType)
        Call Writer.WriteInt(ObjMetadata)
        

  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function PrepareMessageObjectCreate de Protocol.bas")
End Function

''
' Prepares the "CharacterRemove" message and returns it.
'
' @param    CharIndex Character to be removed.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterRemove(ByVal CharIndex As Integer) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "CharacterRemove" message and returns it
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ServerPacketID.CharacterRemove)
        Call Writer.WriteInt16(CharIndex)

  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function PrepareMessageCharacterRemove de Protocol.bas")
End Function

''
' Prepares the "RemoveCharDialog" message and returns it.
'
' @param    CharIndex Character whose dialog will be removed.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageRemoveCharDialog(ByVal CharIndex As Integer) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RemoveCharDialog" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ServerPacketID.RemoveCharDialog)
        Call Writer.WriteInt16(CharIndex)
        
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function PrepareMessageRemoveCharDialog de Protocol.bas")
End Function

''
' Writes the "CharacterCreate" message to the given user's outgoing data buffer.
'
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    X X coord of the new character's position.
' @param    Y Y coord of the new character's position.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @param    name Name of the new character.
' @param    NickColor Determines if the character is a criminal or not, and if can be atacked by someone
' @param    privileges Sets if the character is a normal one or any kind of administrative character.
' @param    bHostile Determines if the NPC is hostile or not.
' @param    bMerchant Determines if the NPC is merchant or not.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterCreate(ByVal body As Integer, ByVal head As Integer, ByVal heading As eHeading, _
                                ByVal CharIndex As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal Weapon As Integer, ByVal shield As Integer, _
                                ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer, ByVal Name As String, ByVal NickColor As Byte, ByVal Alignment As eCharacterAlignment, _
                                ByVal Privileges As Byte, ByVal bHostile As Boolean, ByVal bMerchant As Boolean, ByVal isSailing As Boolean, Optional ByVal NpcNumber As Integer = 0, _
                                Optional ByVal OverHeadIcon As Integer = 0) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 17/09/14
'Prepares the "CharacterCreate" message and returns it
'17/09/14: D'Artagnan - bHostile and bMerchant parameters.
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ServerPacketID.CharacterCreate)
        
        Call Writer.WriteInt16(CharIndex)
        Call Writer.WriteInt16(body)
        Call Writer.WriteInt16(head)
        Call Writer.WriteInt8(heading)
        Call Writer.WriteInt8(X)
        Call Writer.WriteInt8(Y)
        Call Writer.WriteInt16(Weapon)
        Call Writer.WriteInt16(shield)
        Call Writer.WriteInt16(helmet)
        Call Writer.WriteInt16(FX)
        Call Writer.WriteInt16(FXLoops)
        Call Writer.WriteString8(Name)
        Call Writer.WriteInt8(NickColor)
        Call Writer.WriteInt8(Alignment)
        Call Writer.WriteInt8(Privileges)
        Call Writer.WriteBool(bHostile)
        Call Writer.WriteBool(bMerchant)
        Call Writer.WriteBool(isSailing)
        Call Writer.WriteInt16(NpcNumber)
        Call Writer.WriteInt(OverHeadIcon)
        
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function PrepareMessageCharacterCreate de Protocol.bas")
End Function

''
' Prepares the "CharacterChange" message and returns it.
'
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterChange(ByVal body As Integer, ByVal head As Integer, ByVal heading As eHeading, _
                                ByVal CharIndex As Integer, ByVal Weapon As Integer, ByVal shield As Integer, _
                                ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer, ByVal isSailing As Boolean, _
                                ByVal IsDead As Boolean, ByVal OverHeadIcon As Integer, ByVal Alignment As Byte) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "CharacterChange" message and returns it
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ServerPacketID.CharacterChange)
        
        Call Writer.WriteInt16(CharIndex)
        Call Writer.WriteInt16(body)
        Call Writer.WriteInt16(head)
        Call Writer.WriteInt8(heading)
        Call Writer.WriteInt16(Weapon)
        Call Writer.WriteInt16(shield)
        Call Writer.WriteInt16(helmet)
        Call Writer.WriteInt16(FX)
        Call Writer.WriteInt16(FXLoops)
        Call Writer.WriteBool(isSailing)
        Call Writer.WriteBool(IsDead)
        Call Writer.WriteInt(OverHeadIcon)
        Call Writer.WriteInt8(Alignment)

  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function PrepareMessageCharacterChange de Protocol.bas")
End Function

''
' Prepares the "CharacterMove" message and returns it.
'
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterMove(ByVal CharIndex As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal Warped As Boolean) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "CharacterMove" message and returns it
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ServerPacketID.CharacterMove)
        Call Writer.WriteInt16(CharIndex)
        Call Writer.WriteInt8(X)
        Call Writer.WriteInt8(Y)
        Call Writer.WriteBool(Warped)

  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function PrepareMessageCharacterMove de Protocol.bas")
End Function

Public Function PrepareMessageForceCharMove(ByVal Direccion As eHeading) As String
'***************************************************
'Author: ZaMa
'Last Modification: 26/03/2009
'Prepares the "ForceCharMove" message and returns it
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ServerPacketID.ForceCharMove)
        Call Writer.WriteInt8(Direccion)

  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function PrepareMessageForceCharMove de Protocol.bas")
End Function

''
' Prepares the "UpdateTagAndStatus" message and returns it.
'
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageUpdateTagAndStatus(ByVal UserIndex As Integer, ByVal NickColor As Byte, _
                                                ByRef Tag As String) As String
'***************************************************
'Author: Alejandro Salvo (Salvito)
'Last Modification: 04/07/07
'Last Modified By: Juan Martín Sotuyo Dodero (Maraxus)
'Prepares the "UpdateTagAndStatus" message and returns it
'15/01/2010: ZaMa - Now sends the nick color instead of the status.
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ServerPacketID.UpdateTagAndStatus)
        
        Call Writer.WriteInt16(UserList(UserIndex).Char.CharIndex)
        Call Writer.WriteInt8(NickColor)
        Call Writer.WriteString8(Tag)

  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function PrepareMessageUpdateTagAndStatus de Protocol.bas")
End Function

''
' Prepares the "ErrorMsg" message and returns it.
'
' @param    message The error message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageErrorMsg(ByVal Message As String, Optional ByVal bCloseConnection As Boolean = False) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "ErrorMsg" message and returns it
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ServerPacketID.ErrorMsg)
        Call Writer.WriteString8(Message)
        Call Writer.WriteBool(bCloseConnection)
            
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function PrepareMessageErrorMsg de Protocol.bas")
End Function

''
' Writes the "StopWorking" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.

Public Sub WriteStopWorking(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 21/02/2010
'
'***************************************************

    Call Writer.WriteInt8(ServerPacketID.StopWorking)
       
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Writes the "CancelOfferItem" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Slot      The slot to cancel.

Public Sub WriteCancelOfferItem(ByVal UserIndex As Integer, ByVal Slot As Byte)
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 05/03/2010
'
'***************************************************

        Call Writer.WriteInt8(ServerPacketID.CancelOfferItem)
        Call Writer.WriteInt8(Slot)
       
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Handles the "SetDialog" message.
'
' @param UserIndex The index of the user sending the message

Public Sub HandleSetDialog(ByVal UserIndex As Integer)
'***************************************************
'Author: Amraphen
'Last Modification: 18/11/2010
'20/11/2010: ZaMa - Arreglo privilegios.
'***************************************************

On Error GoTo ErrHandler

    With UserList(UserIndex)
        Dim NewDialog As String
        NewDialog = Reader.ReadString8

        If .flags.TargetNpc > 0 Then
            ' Dsgm/Dsrm/Rm
            If Not ((.flags.Privilegios And PlayerType.Dios) = 0 And (.flags.Privilegios And (PlayerType.SemiDios Or PlayerType.RoleMaster)) <> (PlayerType.SemiDios Or PlayerType.RoleMaster)) Then
                'Replace the NPC's dialog.
                Npclist(.flags.TargetNpc).Desc = NewDialog
            End If
        End If
    End With
    
    Exit Sub
    
ErrHandler:

    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleSetDialog de Protocol.bas")
End Sub

''
' Handles the "Impersonate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleImpersonate(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 20/11/2010
'
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        ' Dsgm/Dsrm/Rm
        If (.flags.Privilegios And PlayerType.Dios) = 0 And _
           (.flags.Privilegios And (PlayerType.SemiDios Or PlayerType.RoleMaster)) <> (PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

        
        Dim NpcIndex As Integer
        NpcIndex = .flags.TargetNpc
        
        If NpcIndex = 0 Then Exit Sub
        
        ' Copy head, body and desc
        Call ImitateNpc(UserIndex, NpcIndex)
        
        ' Teleports user to npc's coords
        Call WarpUserChar(UserIndex, Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, _
            Npclist(NpcIndex).Pos.Y, False, True)
        
        ' Log gm
        Call LogGM(.Name, "/IMPERSONAR con " & Npclist(NpcIndex).Name & " en mapa " & .Pos.Map)
        
        ' Remove npc
        Call QuitarNPC(NpcIndex)
        
    End With
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleImpersonate de Protocol.bas")
End Sub

''
' Handles the "Imitate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleImitate(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 20/11/2010
'
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        ' Dsgm/Dsrm/Rm/ConseRm
        If (.flags.Privilegios And PlayerType.Dios) = 0 And _
           (.flags.Privilegios And (PlayerType.SemiDios Or PlayerType.RoleMaster)) <> (PlayerType.SemiDios Or PlayerType.RoleMaster) And _
           (.flags.Privilegios And (PlayerType.Consejero Or PlayerType.RoleMaster)) <> (PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub
        
        Dim NpcIndex As Integer
        NpcIndex = .flags.TargetNpc
        
        If NpcIndex = 0 Then Exit Sub
        
        ' Copy head, body and desc
        Call ImitateNpc(UserIndex, NpcIndex)
        Call LogGM(.Name, "/MIMETIZAR con " & Npclist(NpcIndex).Name & " en mapa " & .Pos.Map)
        
    End With
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleImitate de Protocol.bas")
End Sub

''
' Handles the "RecordAdd" message.
'
' @param UserIndex The index of the user sending the message
           
Public Sub HandleRecordAdd(ByVal UserIndex As Integer)
'**************************************************************
'Author: Amraphen
'Last Modify Date: 29/11/2010
'
'**************************************************************

On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim UserName As String
        Dim Reason As String
        
        UserName = Reader.ReadString8
        Reason = Reader.ReadString8

        If (.flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster)) = 0 Then
            
            'Verificamos que exista el personaje
            Dim UserId As Long
            UserId = GetUserID(UserName)
            
            If UserId = 0 Then
                Call WriteShowMessageBox(UserIndex, "El personaje no existe")
            Else
                'Agregamos el seguimiento
                Call AddRecord(UserIndex, UserName, Reason)
                
                'Enviamos la nueva lista de personajes
                Call WriteRecordList(UserIndex)
            End If
        End If

    End With
    
    Exit Sub
    
ErrHandler:

    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleRecordAdd de Protocol.bas")
End Sub

''
' Handles the "RecordAddObs" message.
'
' @param UserIndex The index of the user sending the message.

Public Sub HandleRecordAddObs(ByVal UserIndex As Integer)
'**************************************************************
'Author: Amraphen
'Last Modify Date: 29/11/2010
'
'**************************************************************

On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim RecordIndex As Byte
        Dim Obs As String
        
        RecordIndex = Reader.ReadInt8
        Obs = Reader.ReadString8

        If (.flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster)) = 0 Then
            'Agregamos la observación
            Call AddObs(UserIndex, RecordIndex, Obs)
            
            'Actualizamos la información
            Call WriteRecordDetails(UserIndex, RecordIndex)
        End If
    End With
    
    Exit Sub
    
ErrHandler:

    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleRecordAddObs de Protocol.bas")
End Sub

''
' Handles the "RecordRemove" message.
'
' @param UserIndex The index of the user sending the message.

Public Sub HandleRecordRemove(ByVal UserIndex As Integer)
'***************************************************
'Author: Amraphen
'Last Modification: 29/11/2010
'
'***************************************************
On Error GoTo ErrHandler

    Dim RecordIndex As Integer

    With UserList(UserIndex)

        RecordIndex = Reader.ReadInt8
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub
        
        'Sólo dioses pueden remover los seguimientos, los otros reciben una advertencia:
        If (.flags.Privilegios And PlayerType.Dios) Then
            Call RemoveRecord(RecordIndex)
            Call WriteShowMessageBox(UserIndex, "Se ha eliminado el seguimiento.")
            Call WriteRecordList(UserIndex)
        Else
            Call WriteShowMessageBox(UserIndex, "Sólo los dioses pueden eliminar seguimientos.")
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleRecordRemove de Protocol.bas")
End Sub

''
' Handles the "RecordListRequest" message.
'
' @param UserIndex The index of the user sending the message.
            
Public Sub HandleRecordListRequest(ByVal UserIndex As Integer)
'***************************************************
'Author: Amraphen
'Last Modification: 29/11/2010
'
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub

        Call WriteRecordList(UserIndex)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleRecordListRequest de Protocol.bas")
End Sub

''
' Writes the "RecordDetails" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRecordDetails(ByVal UserIndex As Integer, ByVal RecordIndex As Integer)
'***************************************************
'Author: Amraphen
'Last Modification: 29/11/2010
'Writes the "RecordDetails" message to the given user's outgoing data buffer
'***************************************************
Dim I As Long
Dim tIndex As Integer
Dim TmpStr As String
Dim TempDate As Date

    
        Call Writer.WriteInt8(ServerPacketID.RecordDetails)
        
        'Creador y motivo
        Call Writer.WriteString8(Records(RecordIndex).Creador)
        Call Writer.WriteString8(Records(RecordIndex).Motivo)
        
        tIndex = NameIndex(Records(RecordIndex).Usuario)
        
        'Status del pj (online?)
        Call Writer.WriteBool(tIndex > 0)
        
        'Escribo la IP según el estado del personaje
        If tIndex > 0 Then
            'La IP Actual
            TmpStr = UserList(tIndex).IP
        Else 'String nulo
            TmpStr = vbNullString
        End If
        Call Writer.WriteString8(TmpStr)
        
        'Escribo tiempo online según el estado del personaje
        If tIndex > 0 Then
            'Tiempo logueado.
            TempDate = Now - UserList(tIndex).LogOnTime
            TmpStr = Hour(TempDate) & ":" & Minute(TempDate) & ":" & Second(TempDate)
        Else
            'Envío string nulo.
            TmpStr = vbNullString
        End If
        Call Writer.WriteString8(TmpStr)

        'Escribo observaciones:
        TmpStr = vbNullString
        If Records(RecordIndex).NumObs Then
            For I = 1 To Records(RecordIndex).NumObs
                TmpStr = TmpStr & Records(RecordIndex).Obs(I).Creador & "> " & Records(RecordIndex).Obs(I).Detalles & vbCrLf
            Next I
            
            TmpStr = Left$(TmpStr, Len(TmpStr) - 1)
        End If
        Call Writer.WriteString8(TmpStr)
       
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Writes the "RecordList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRecordList(ByVal UserIndex As Integer)
'***************************************************
'Author: Amraphen
'Last Modification: 29/11/2010
'Writes the "RecordList" message to the given user's outgoing data buffer
'***************************************************
Dim I As Long

    
        Call Writer.WriteInt8(ServerPacketID.RecordList)
        
        Call Writer.WriteInt8(NumRecords)
        For I = 1 To NumRecords
            Call Writer.WriteString8(Records(I).Usuario)
        Next I
       
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Writes the "ShowMenu" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    MenuIndex: The menu index.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowMenu(ByVal UserIndex As Integer, ByVal MenuIndex As Byte)
'***************************************************
'Author: ZaMa
'Last Modification: 10/05/2011
'Writes the "ShowMenu" message to the given user's outgoing data buffer
'***************************************************

        Call Writer.WriteInt8(ServerPacketID.ShowMenu)
        
        Call Writer.WriteInt8(MenuIndex)
       
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

''
' Handles the "RecordDetailsRequest" message.
'
' @param UserIndex The index of the user sending the message.
            
Public Sub HandleRecordDetailsRequest(ByVal UserIndex As Integer)
'***************************************************
'Author: Amraphen
'Last Modification: 07/04/2011
'Handles the "RecordListRequest" message
'***************************************************
On Error GoTo ErrHandler
  
Dim RecordIndex As Byte

    With UserList(UserIndex)

        RecordIndex = Reader.ReadInt8
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub
        
        Call WriteRecordDetails(UserIndex, RecordIndex)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleRecordDetailsRequest de Protocol.bas")
End Sub

Public Sub HandleMoveItem(ByVal UserIndex As Integer)
'***************************************************
'Author: Ignacio Mariano Tirabasso (Budi)
'Last Modification: 07/08/2014 (D'Artagnan)
'07/08/2014: D'Artagnan - Bank item support.
'***************************************************

    On Error GoTo ErrHandler
      
    With UserList(UserIndex)
    
        Dim originalSlot As Byte
        Dim newSlot As Byte
        Dim moveType As eMoveType

        originalSlot = Reader.ReadInt8
        newSlot = Reader.ReadInt8
        moveType = Reader.ReadInt8
        
        Select Case moveType
            Case eMoveType.Inventory
                Call InvUsuario.moveItem(UserIndex, originalSlot, newSlot)
            
            Case eMoveType.Bank
                Call MoveBankItem(UserIndex, originalSlot, newSlot)
        End Select
        
    End With
    
      
    Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleMoveItem de Protocol.bas")
End Sub

Public Function PrepareMessageCharacterAttackMovement(ByVal CharIndex As Integer) As String
'***************************************************
'Author: Amraphen
'Last Modification: 24/05/2011
'Prepares the "CharacterAttackMovement" message and returns it.
'***************************************************
On Error GoTo ErrHandler
  
        Call Writer.WriteInt8(ServerPacketID.CharacterAttackMovement)
        Call Writer.WriteInt16(CharIndex)
        
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function PrepareMessageCharacterAttackMovement de Protocol.bas")
End Function

''
' Writes the "StrDextRunningOut" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Seconds Seconds left.

Public Sub WriteStrDextRunningOut(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Dalmasso (CHOTS)
'Last Modification: 08/06/2011
'
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ServerPacketID.StrDextRunningOut)
       
    Call SendData(ToUser, UserIndex, vbNullString)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteStrDextRunningOut de Protocol.bas")
End Sub

''
' Handles the "PMSend" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandlePMSend(ByVal UserIndex As Integer)
'***************************************************
'Author: Amraphen
'Last Modification: 04/08/2011
'Handles the "PMSend" message.
'***************************************************

On Error GoTo ErrHandler

    With UserList(UserIndex)
        Dim UserName As String
        Dim Message As String
        Dim TargetIndex As Integer

        UserName = Reader.ReadString8
        Message = Reader.ReadString8
        
        TargetIndex = NameIndex(UserName)

        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios) Then
            If TargetIndex = 0 Then 'Offline

                Dim UserId As Long
                UserId = GetUserID(UserName)
                
                If UserId <> 0 Then
                    Call AddUserPrivateMsjDB(UserId, Message)

                    Call WriteConsoleMsg(UserIndex, "Mensaje enviado.", FontTypeNames.FONTTYPE_GM)
                Else
                    Call WriteConsoleMsg(UserIndex, "Usuario inexistente.", FontTypeNames.FONTTYPE_INFO)
                End If
            Else 'Online
                Call AgregarMensaje(TargetIndex, .Name, Message)
                Call WriteConsoleMsg(UserIndex, "Mensaje enviado.", FontTypeNames.FONTTYPE_GM)
            End If
            
            Call LogGM(.Name, "/ENVIARMP " & UserName & " Mensaje: " & Message)
        End If
    End With
    
    Exit Sub
    
ErrHandler:

    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandlePMSend de Protocol.bas")
End Sub

''
' Handles the "PMList" message.
'
' @param    UserIndex The index of the user sending the message.

Public Sub HandlePMList(ByVal UserIndex As Integer)
'***************************************************
'Author: Amraphen
'Last Modification: 04/08/2011
'Handles the "PMList" message.
'***************************************************
On Error GoTo ErrHandler
  
Dim LoopC As Long
    
    With UserList(UserIndex)

        If .UltimoMensaje = 0 Then
            Call WriteConsoleMsg(UserIndex, "No tienes mensajes privados.", FontTypeNames.FONTTYPE_INFOBOLD)
        Else
            'Envía la lista de mensajes privados al usuario:
            Call WriteConsoleMsg(UserIndex, "Mensajes privados: ", FontTypeNames.FONTTYPE_INFOBOLD)
            
            For LoopC = 1 To .UltimoMensaje
                With .Mensajes(LoopC)
                    If .Contenido = vbNullString Then
                        Call WriteConsoleMsg(UserIndex, "MENSAJE " & LoopC & "> VACÍO.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        If .Nuevo Then
                            Call WriteConsoleMsg(UserIndex, "MENSAJE " & LoopC & "> (!)" & .Contenido, FontTypeNames.FONTTYPE_FIGHT)
                            .Nuevo = False
                        Else
                            Call WriteConsoleMsg(UserIndex, "MENSAJE " & LoopC & "> " & .Contenido, FontTypeNames.FONTTYPE_INFO)
                        End If
                    End If
                End With
            Next LoopC
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandlePMList de Protocol.bas")
End Sub

Public Sub HandlePMDeleteList(ByVal UserIndex As Integer)
'***************************************************
'Author: Amraphen
'Last Modification: 04/08/2011
'Handles the "PMDeleteList" message.
'***************************************************
On Error GoTo ErrHandler
  

    With UserList(UserIndex)

        Call LimpiarMensajes(UserIndex)
        Call WriteConsoleMsg(UserIndex, "Se han borrado tus mensajes privados.", FontTypeNames.FONTTYPE_INFO)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandlePMDeleteList de Protocol.bas")
End Sub

Public Sub HandlePMDeleteUser(ByVal UserIndex As Integer)
'***************************************************
'Author: Amraphen
'Last Modification: 04/08/2011
'Handles the "PMDeleteUser" message.
'***************************************************

On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim UserName As String
        Dim MpIndex As Byte
        Dim TargetIndex As Integer
        
        UserName = Reader.ReadString8
        MpIndex = Reader.ReadInt8

        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios) Then
            TargetIndex = NameIndex(UserName)
            If TargetIndex = 0 Then 'Offline

                Dim UserId As Long
                UserId = GetUserID(UserName)
                
                If UserId <> 0 Then
                    Call DeleteUserPrivateMsjDB(UserId, MpIndex)
                    Call WriteConsoleMsg(UserIndex, "Mensaje/s borrado/s.", FontTypeNames.FONTTYPE_INFO)

                Else
                    Call WriteConsoleMsg(UserIndex, "Usuario inexistente.", FontTypeNames.FONTTYPE_INFO)
                End If
            Else 'Online
                If MpIndex = 0 Then
                    Call LimpiarMensajes(TargetIndex)
                    Call WriteConsoleMsg(UserIndex, "Mensajes borrados.", FontTypeNames.FONTTYPE_INFO)
                    
                    Call WriteConsoleMsg(UserIndex, "Mensajes borrados.", FontTypeNames.FONTTYPE_INFO)
                ElseIf MpIndex >= 1 And MpIndex <= Constantes.MaxPrivateMessages Then
                    Call BorrarMensaje(TargetIndex, MpIndex)
                    
                    Call WriteConsoleMsg(UserIndex, "Mensaje borrado.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
    End With
    
    Exit Sub
    
ErrHandler:

    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandlePMDeleteUser de Protocol.bas")
End Sub

Public Sub HandlePMListUser(ByVal UserIndex As Integer)
'***************************************************
'Author: Amraphen
'Last Modification: 04/08/2011
'Handles the "PMListUser" message.
'***************************************************

On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim UserName As String
        Dim MpIndex As Byte
        Dim TargetIndex As Integer
        Dim LoopC As Long
        
        UserName = Reader.ReadString8
        MpIndex = Reader.ReadInt8

        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios) Then
            TargetIndex = NameIndex(UserName)
            If TargetIndex = 0 Then 'Offline
            
                Dim UserId As Long
                UserId = GetUserID(UserName)
                
                If UserId <> 0 Then
                    Call SendUserMessagesDB(UserIndex, UserId, UserName, MpIndex)

                Else
                    Call WriteConsoleMsg(UserIndex, "Usuario inexistente.", FontTypeNames.FONTTYPE_INFO)
                End If
            Else 'Online
                With UserList(TargetIndex)
                    If MpIndex <> 0 Then
                        Call WriteConsoleMsg(UserIndex, "Mensajes privados de " & UserName & ":", FontTypeNames.FONTTYPE_INFOBOLD)
                        Call EnviarMensaje(UserIndex, MpIndex, .Mensajes(MpIndex).Contenido, .Mensajes(MpIndex).Nuevo)
                        
                    ElseIf .UltimoMensaje = 0 Then
                        Call WriteConsoleMsg(UserIndex, "El usuario " & UserName & " no tiene mensajes privados.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call WriteConsoleMsg(UserIndex, "Mensajes privados de " & UserName & ":", FontTypeNames.FONTTYPE_INFOBOLD)
                        
                        For LoopC = 1 To .UltimoMensaje
                            Call EnviarMensaje(UserIndex, LoopC, .Mensajes(LoopC).Contenido, .Mensajes(LoopC).Nuevo)
                        Next LoopC
                    End If
                End With
            End If
        End If
    End With
   
    Exit Sub
    
ErrHandler:

    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandlePMListUser de Protocol.bas")
End Sub

''
' Handles the "HigherAdminsMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleHigherAdminsMessage(ByVal UserIndex As Integer)
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 03/30/12
'
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim Message As String
        
        Message = Reader.ReadString8()

        If ((.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0) And ((.flags.Privilegios And PlayerType.RoleMaster) = 0) Then
            Call LogGM(.Name, "Mensaje a Dioses:" & Message)
        
            If LenB(Message) <> 0 Then
                'Analize chat...
                Call Statistics.ParseChat(Message)
            
                Call SendData(SendTarget.ToHigherAdminsButRMs, 0, PrepareMessageConsoleMsg(.Name & "(Sólo Dioses)> " & Message, FontTypeNames.FONTTYPE_GMMSG))
            End If
        End If
    End With
   
    Exit Sub
    
ErrHandler:

    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleHigherAdminsMessage de Protocol.bas")
End Sub

Public Sub HandleMenuAction(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 22/03/2012
'Handles the "MenuAction" message.
'***************************************************
  
On Error GoTo ErrHandler
 
    With UserList(UserIndex)

        ' Perform action
        Dim iAction As Integer
        iAction = Reader.ReadInt16
        
        Dim Slot As Byte
        Slot = Reader.ReadInt8
        
        Call PerformMenuAction(UserIndex, iAction, Slot)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleMenuAction de Protocol.bas")
End Sub

Public Sub HandleTournamentParticipate(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 31/05/2012
'***************************************************
On Error GoTo ErrHandler

    ' Validate registration
    If UserCanRegisterIntoTorunament(UserIndex) Then
        ' Registrate
        Call RegisterUserToTournament(UserIndex)
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleTournamentParticipate de Protocol.bas")
End Sub
 
Public Sub HandleRequestTournamentCompetitors(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 31/05/2012
'***************************************************
On Error GoTo ErrHandler
  
    
    With UserList(UserIndex)

        ' Gms, Dioses, Admins (No Rms)
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 And _
           (.flags.Privilegios And (PlayerType.RoleMaster)) = 0 Then
           
           Call SendCompetitorsList(UserIndex)
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleRequestTournamentCompetitors de Protocol.bas")
End Sub
 
Public Sub HandleTournamentDisqualify(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 31/05/2012
'***************************************************
On Error GoTo ErrHandler

    Dim CompetitorName As String
    Dim CompetitorIndex As Integer
     
    With UserList(UserIndex)

        ' Competitor Name
        CompetitorName = Reader.ReadString8

        ' Gms, Dioses, Admins (No Rms)
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 And _
           (.flags.Privilegios And (PlayerType.RoleMaster)) = 0 Then
            
            ' Tuornament active?
            If Not Tournament.Active Then
                Call WriteConsoleMsg(UserIndex, "No hay ningún torneo activo.", FontTypeNames.FONTTYPE_SERVER)
            
            ' Any competitor remaining?
            ElseIf Tournament.CompetitorsList.Longitud = 0 Then
                Call WriteConsoleMsg(UserIndex, "No hay ningún participante en la lista.", FontTypeNames.FONTTYPE_SERVER)
                
            ' User is in the list?
            ElseIf Tournament.CompetitorsList.Existe(CompetitorName) Then
                Tournament.CompetitorsList.Quitar (CompetitorName)
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(CompetitorName & " ha sido descalificado.", FontTypeNames.FONTTYPE_SERVER))
            Else
                Call WriteConsoleMsg(UserIndex, "El usuario no está participando en el torneo.", FontTypeNames.FONTTYPE_SERVER)
            End If
        End If
    End With

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleTournamentDisqualify de Protocol.bas")
End Sub
 
Public Sub HandleTournamentFight(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 31/05/2012
'**************************************************
On Error GoTo ErrHandler

    With UserList(UserIndex)
        Dim Competitor1 As String
        Competitor1 = Reader.ReadString8
        
        Dim Competitor2 As String
        Competitor2 = Reader.ReadString8
        
        Dim ArenaIndex As Integer
        ArenaIndex = Reader.ReadInt16
        
        ' Gms, Dioses, Admins (No Rms)
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 And _
           (.flags.Privilegios And (PlayerType.RoleMaster)) = 0 Then
            
            Call TournamentFightBegin(UserIndex, Competitor1, Competitor2, ArenaIndex)
        End If
        
    End With

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleTournamentFight de Protocol.bas")
End Sub
 
Public Sub HandleTorunamentCancel(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 31/05/2012
'***************************************************
On Error GoTo ErrHandler
  

    With UserList(UserIndex)

        ' Gms, Dioses, Admins (No Rms)
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 And _
           (.flags.Privilegios And (PlayerType.RoleMaster)) = 0 Then
        
            With Tournament
                If .Active Then
                
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Torneo Finalizado.", FontTypeNames.FONTTYPE_SERVER))
                    
                    Dim Index As Long
                    Dim CompetitorIndex As Integer
                    
                    ' Reset remaining competitors state
                    For Index = 1 To .CompetitorsList.Longitud
                        CompetitorIndex = NameIndex(.CompetitorsList.VerElemento(Index))
                        If CompetitorIndex <> 0 Then
                            UserList(CompetitorIndex).flags.TournamentState = eTournamentState.ieNone
                            Call TournamentUserExpell(CompetitorIndex, eTournamentExpellMotive.ieMassiveExpell)
                        End If
                    Next Index
                    
                    .Active = False
                End If
            End With
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleTorunamentCancel de Protocol.bas")
End Sub

Public Sub HandleTournamentEdit(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 31/05/2012
'Allow admins to edit or save the tournament configuration.
'***************************************************
On Error GoTo ErrHandler
  

    With UserList(UserIndex)

        Dim EditOption As Byte
        EditOption = Reader.ReadInt8
        
        Dim yTemp As Byte, yTemp2 As Byte
        Dim lTemp As Long
        Dim iTemp As Long
        
        ' Gms, Dioses, Admins (No Rms)
        Dim SaveEdition As Boolean
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 And _
           (.flags.Privilegios And (PlayerType.RoleMaster)) = 0 Then
           
            If Not Tournament.Active Or EditOption = eTournamentEdit.ieSaveConfig Then
                SaveEdition = True
            Else
                Call WriteConsoleMsg(UserIndex, "No se puede editar las condiciones de un torneo mientras esta activo.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        Select Case EditOption
            ' Num of competitors
            Case eTournamentEdit.ieMaxCompetitor
                
                yTemp = Reader.ReadInt8
                If SaveEdition Then
                    Tournament.MaxCompetitors = yTemp
                End If
                
            ' Min level
            Case eTournamentEdit.ieMinLevel
                
                yTemp = Reader.ReadInt8
                If SaveEdition Then
                    Tournament.MinLevel = yTemp
                End If
                
            ' Max level
            Case eTournamentEdit.ieMaxLevel
            
                yTemp = Reader.ReadInt8
                If SaveEdition Then
                    Tournament.MinLevel = yTemp
                End If
                
            ' Gold Required
            Case eTournamentEdit.ieRequiredGold
            
                lTemp = Reader.ReadInt32
                If SaveEdition Then
                    Tournament.RequiredGold = lTemp
                End If
                
            ' Forbidden items
            Case eTournamentEdit.ieForbiddenItems
                
                yTemp = Reader.ReadInt8
                If SaveEdition Then
                    Tournament.NumForbiddenItems = yTemp
                    
                    If yTemp > 0 Then _
                        ReDim Tournament.ForbiddenItem(1 To yTemp)
                End If
                
                For lTemp = 1 To yTemp
                    iTemp = Reader.ReadInt16
                    If SaveEdition Then
                        Tournament.ForbiddenItem(lTemp) = iTemp
                    End If
                Next lTemp
                
            ' Permited class
            Case eTournamentEdit.iePermitedClass
                
                yTemp = Reader.ReadInt8
                If SaveEdition Then
                    Tournament.NumPermitedClass = yTemp
                    
                    If yTemp > 0 Then _
                        ReDim Tournament.PermitedClass(1 To yTemp)
                End If
                
                For lTemp = 1 To yTemp
                    yTemp2 = Reader.ReadInt8
                    If SaveEdition Then
                        Tournament.PermitedClass(lTemp) = yTemp2
                    End If
                Next lTemp
                
            ' Num. Rounds
            Case eTournamentEdit.ieNumRoundsToWin
                
                yTemp = Reader.ReadInt8
                If SaveEdition Then
                    Tournament.NumRoundsToWin = yTemp
                End If
                
            ' Kill After Loose
            Case eTournamentEdit.ieKillAfterLoose
            
                yTemp = Reader.ReadInt8
                If SaveEdition Then
                    Tournament.KillAfterLoose = yTemp
                End If
                
            ' Waiting Map
            Case eTournamentEdit.ieWaitingMap
                
                iTemp = Reader.ReadInt16
                yTemp = Reader.ReadInt8
                yTemp2 = Reader.ReadInt8
                If SaveEdition Then
                    Tournament.WaitingMap.Map = iTemp
                    Tournament.WaitingMap.X = yTemp
                    Tournament.WaitingMap.Y = yTemp2
                End If
                
            ' Arena Position
            Case eTournamentEdit.ieArenaPosition
                yTemp = Reader.ReadInt8
                If SaveEdition Then
                    Tournament.Arenas(yTemp).Map = Reader.ReadInt16
                    Tournament.Arenas(yTemp).UserPos1.X = Reader.ReadInt8
                    Tournament.Arenas(yTemp).UserPos1.Y = Reader.ReadInt8
                    Tournament.Arenas(yTemp).UserPos2.X = Reader.ReadInt8
                    Tournament.Arenas(yTemp).UserPos2.Y = Reader.ReadInt8
                Else
                    Reader.ReadInt16
                    Reader.ReadInt8
                    Reader.ReadInt8
                    Reader.ReadInt8
                    Reader.ReadInt8
                End If
                
            ' Final Map
            Case eTournamentEdit.ieFinalMap
                
                iTemp = Reader.ReadInt16
                yTemp = Reader.ReadInt8
                yTemp2 = Reader.ReadInt8
                If SaveEdition Then
                    Tournament.FinalMap.Map = iTemp
                    Tournament.FinalMap.X = yTemp
                    Tournament.FinalMap.Y = yTemp2
                End If
            
            ' Save Config
            Case eTournamentEdit.ieSaveConfig
                If SaveEdition Then
                    'TODO_TORNEO: dump al dat :O
                End If
        End Select
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleTournamentEdit de Protocol.bas")
End Sub

Public Sub HandleTorunamentBegin(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 31/05/2012
'Begins tournament by sending its configuration info to users and opening registration
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        ' Gms, Dioses, Admins (No Rms)
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 And _
           (.flags.Privilegios And (PlayerType.RoleMaster)) = 0 Then
            
            With Tournament
                If Not .Active Then
                    Dim Mensaje As String
                    Mensaje = UCase$(UserList(UserIndex).Name) & " ESTÁ ORGANIZANDO UN TORNEO, CON " & .MaxCompetitors & " CUPOS, "
                    
                    If .MinLevel <> 0 Then
                        Mensaje = Mensaje & "LEVEL MÍNIMO: " & .MinLevel & ", "
                    End If
                    
                    If .MaxLevel <> 0 Then
                        Mensaje = Mensaje & "LEVEL MÁXIMO: " & .MaxLevel & ", "
                    End If
                    'TODO_TORNEO: Toda la config
                    
                    Mensaje = Mensaje & ", PARA PARTICIPAR ENVIA /PARTICIPAR."
                    
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(Mensaje, FontTypeNames.FONTTYPE_CENTINELA))
                    
                    .RegistrationCountdown = 10
                    .CountdownActivated = True
                Else
                    Call WriteConsoleMsg(UserIndex, "Ya hay un torneo.", FontTypeNames.FONTTYPE_INFO)
                End If
            End With
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleTorunamentBegin de Protocol.bas")
End Sub

''
' Writes the "TournamentCompetitorsList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    npcNames The names of the creatures that can be spawned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTournamentCompetitorList(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 07/06/2012
'Writes the "TournamentCompetitorsList" message to the given user's outgoing data buffer
'***************************************************

    Dim Index As Long
    Dim Tmp As String
    
    
        Call Writer.WriteInt8(ServerPacketID.TournamentCompetitorList)
        
        With Tournament.CompetitorsList
            For Index = 1 To .Longitud
                Tmp = Tmp & .VerElemento(Index) & SEPARATOR
            Next Index
        End With
        
        If Len(Tmp) Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call Writer.WriteString8(Tmp)
       
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub
 
Public Sub HandleRequestTournamentConfig(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 06/07/2012
'***************************************************
On Error GoTo ErrHandler
  
    
    With UserList(UserIndex)

    'TODO_TORNEO:
        ' Gms, Dioses, Admins (No Rms)
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 And _
           (.flags.Privilegios And (PlayerType.RoleMaster)) = 0 Then
           
           'Call SendCompetitorsList(UserIndex)
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleRequestTournamentConfig de Protocol.bas")
End Sub

Public Sub WriteTournamentConfig(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 06/07/2012
'***************************************************
On Error GoTo ErrHandler
  
    
    Dim lTemp As Long
    
    
        Call Writer.WriteInt8(ServerPacketID.TournamentConfig)
        
        ' General
        Call Writer.WriteInt8(Tournament.MinLevel)
        Call Writer.WriteInt8(Tournament.MaxLevel)
        Call Writer.WriteInt8(Tournament.MaxCompetitors)
        Call Writer.WriteInt8(Tournament.NumRoundsToWin)
        Call Writer.WriteInt32(Tournament.RequiredGold)
        Call Writer.WriteInt8(Tournament.KillAfterLoose)
        
        ' Classes
        Call Writer.WriteInt8(Tournament.NumPermitedClass)
        For lTemp = 1 To Tournament.NumPermitedClass
            Call Writer.WriteInt8(Tournament.PermitedClass(lTemp))
        Next lTemp
        
        ' Items
        Call Writer.WriteInt8(Tournament.NumForbiddenItems)
        For lTemp = 1 To Tournament.NumForbiddenItems
            Call Writer.WriteInt16(Tournament.ForbiddenItem(lTemp))
        Next lTemp
    
        ' Maps
        Call Writer.WriteInt16(Tournament.WaitingMap.Map)
        Call Writer.WriteInt8(Tournament.WaitingMap.X)
        Call Writer.WriteInt8(Tournament.WaitingMap.Y)
        
        Call Writer.WriteInt16(Tournament.FinalMap.Map)
        Call Writer.WriteInt8(Tournament.FinalMap.X)
        Call Writer.WriteInt8(Tournament.FinalMap.Y)
        
        For lTemp = 1 To MAX_ARENAS
            Call Writer.WriteInt16(Tournament.Arenas(lTemp).Map)
            
            Call Writer.WriteInt8(Tournament.Arenas(lTemp).UserPos1.X)
            Call Writer.WriteInt8(Tournament.Arenas(lTemp).UserPos1.X)
            
            Call Writer.WriteInt8(Tournament.Arenas(lTemp).UserPos2.X)
            Call Writer.WriteInt8(Tournament.Arenas(lTemp).UserPos2.X)
        Next lTemp
       
    Call SendData(ToUser, UserIndex, vbNullString)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteTournamentConfig de Protocol.bas")
End Sub

Private Sub HandleAccountLoginChar(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Creation Date: 23/01/2014
'Last Modification: 20/06/2014 (D'Artagnan)
'Logs account char.
'20/06/2014: D'Artagnan - Read account session token.
'***************************************************

On Error GoTo ErrHandler
    
    With UserList(UserIndex)

        Dim CharSlot As Byte
        Dim sVersion As String
        Dim ClientTempCode As String
        Dim SessionToken As String
        Dim SessionIndex As Integer
        
        CharSlot = Reader.ReadInt8()
        sVersion = CStr(Reader.ReadInt8()) & "." & CStr(Reader.ReadInt8()) & "." & CStr(Reader.ReadInt8())
        SessionToken = Reader.ReadString8()
        ClientTempCode = Reader.ReadString8()

        .ClientTempCode = ClientTempCode
        
        If Not VersionOK(sVersion) Then
            Call WriteCloseForm(UserIndex, "frmAccount")
            Call DisconnectWithMessage(UserIndex, OUTDATED_VERSION)
        End If
        
        ' Valid slot?
        If CharSlot > MAX_ACCOUNT_CHARS Then Exit Sub
        
        ' Check if the session is valid
        If Not modSession.IsTokenValid(SessionToken, UserIndex, SessionIndex) Then
            Call WriteCloseForm(UserIndex, "frmAccount")
            Call DisconnectWithMessage(UserIndex, "Su sesión es inválida. Por favor vuelva a ingresar.")
        End If
        
        ' Check if the session expired
        If modSession.SessionExpired(SessionIndex) Then
            Call modSession.CleanSessionSlot(SessionIndex)
            Call WriteCloseForm(UserIndex, "frmAccount")
            Call DisconnectWithMessage(UserIndex, "Su sesión ha expirado. Por favor vuelva a ingresar.")
        End If
        
        .nSessionId = SessionIndex
        .AccountId = aActiveSessions(SessionIndex).nAccountID
        .AccountName = aActiveSessions(SessionIndex).sAccountName
        .AccountEmail = aActiveSessions(SessionIndex).AccountEmail
        
        If modAccount.ConnectChar(UserIndex, CharSlot) Then
            ' Remove token once used.
            Call modSession.CleanSessionSlot(SessionIndex)
        Else
            Call WriteCloseForm(UserIndex, "frmAccount")
            Call TCP.CloseSocket(UserIndex)
        End If
        
        .ClientTempCode = ClientTempCode
    End With
    
    Exit Sub
    
ErrHandler:

    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleAccountLoginChar de Protocol.bas")
        
End Sub
 
Private Sub HandleAccountCreateChar(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Creation Date: 23/01/2014
'Last Modification: 02/06/2014 (D'Artagnan)
'Crea y conecta un personaje de la cuenta.
'***************************************************

On Error GoTo ErrHandler

'
' @ Crea un personaje.
 
    Dim s_name      As String
    Dim s_Genero    As eGenero
    Dim s_Raza      As eRaza
    Dim s_Clase     As eClass
    Dim s_Head      As Integer
    Dim s_Home      As Byte
    Dim sVersion    As String
    Dim SessionToken      As String
    Dim sClientTempCode As String
    Dim SessionIndex As Integer

    s_name = Trim(Reader.ReadString8())
    s_Raza = Reader.ReadInt8()
    s_Genero = Reader.ReadInt8()
    s_Clase = Reader.ReadInt8()
    s_Head = Reader.ReadInt16()
    s_Home = Reader.ReadInt8()
    sVersion = CStr(Reader.ReadInt8()) & "." & CStr(Reader.ReadInt8()) & "." & CStr(Reader.ReadInt8())
    SessionToken = Reader.ReadString8()
    sClientTempCode = Reader.ReadString8()

    UserList(UserIndex).ClientTempCode = sClientTempCode

    If Not VersionOK(sVersion) Then
        Call DisconnectWithMessage(UserIndex, OUTDATED_VERSION)
        Exit Sub
    End If
    
    ' Check if the session is valid
    If Not modSession.IsTokenValid(SessionToken, UserIndex, SessionIndex) Then
        Call WriteCloseForm(UserIndex, "frmAccount")
        Call DisconnectWithMessage(UserIndex, "Su sesión es inválida. Por favor vuelva a ingresar.")
        Exit Sub
    End If
    
    ' Check if the session expired
    If modSession.SessionExpired(SessionIndex) Then
        Call modSession.CleanSessionSlot(SessionIndex)
        Call WriteCloseForm(UserIndex, "frmAccount")
        Call DisconnectWithMessage(UserIndex, "Su sesión ha expirado. Por favor vuelva a ingresar.")
        Exit Sub
    End If
    
    If aActiveSessions(SessionIndex).nAccountID = 0 Then
        Call modSession.CleanSessionSlot(SessionIndex)
        Call WriteCloseForm(UserIndex, "frmAccount")
        Call DisconnectWithMessage(UserIndex, "Error al comprobar la integridad de la sesión. Por favor vuelva a ingresar.")
        Exit Sub
    End If
    
    If Len(s_name) > MAX_NICKNAME_SIZE Then
        Call DisconnectWithMessage(UserIndex, "El nombre de tu personaje no puede superar los " & MAX_NICKNAME_SIZE & " caracteres")
        Exit Sub
    End If
    
    If Not Classes(s_Clase).Enabled Then
        Call DisconnectWithMessage(UserIndex, "La clase seleccionada no se encuentra habilitada.")
        Exit Sub
    End If
        
    If PuedeCrearPersonajes = 0 Then
        Call DisconnectWithMessage(UserIndex, "La creación de personajes se encuentra deshabilitada.")
        Exit Sub
    End If
        
    ' Extend the session lifetime
    Call modSession.ExtendSessionLifetime(SessionIndex)
        
    ' Create the character and connect
    If modAccount.ConectarNuevoPersonaje(UserIndex, s_name, s_Genero, s_Raza, s_Clase, s_Head, s_Home, SessionIndex) Then
        ' Assign the session data and clean the slot.
        With UserList(UserIndex)
            .nSessionId = SessionIndex
            .AccountId = aActiveSessions(SessionIndex).nAccountID
            .AccountName = aActiveSessions(SessionIndex).sAccountName
        End With
        ' Clean the session slot as it's not used anymore. The caracter will be logged in.
        Call modSession.CleanSessionSlot(SessionIndex)
    End If
      
  Exit Sub
  
ErrHandler:

    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleAccountCreateChar de Protocol.bas")

    Call WriteErrorMsg(UserIndex, "Se produjo un error al crear el personaje. Intente denuevo más tarde", False)
    Call WriteCloseForm(UserIndex, "frmCrearPersonaje")
    Call CloseSocket(UserIndex)
End Sub
 
Private Sub HandleAccountLogin(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Creation Date: 23/01/2014
'Last Modification: 23/01/2014
'Logs in an account.
'***************************************************
Dim isValidData As Boolean
isValidData = True

On Error GoTo ErrHandler
    
    With UserList(UserIndex)

        Dim AccountName As String
        Dim Password As String
        Dim sPreviousToken As String
        Dim clientMD5 As String
        Dim version As String
        Dim sClientTempCode As String

        AccountName = Reader.ReadString8()
        Password = Reader.ReadString8()
        sPreviousToken = Reader.ReadString8()
        sClientTempCode = Reader.ReadString8
        version = CStr(Reader.ReadInt8()) & "." & CStr(Reader.ReadInt8()) & "." & CStr(Reader.ReadInt8())
        
        .ClientTempCode = sClientTempCode

#If EnableSecurity Then
        clientMD5 = Reader.ReadString8()
        
        If MD5ClientesActivado Then
            If Not MD5ok(clientMD5) Then
                Call WriteErrorMsg(UserIndex, "El cliente está dañado, por favor descarguelo nuevamente desde www.alkononline.com.ar")
                isValidData = False
            End If
        End If
#End If

            If Not VersionOK(version) Then
            Call WriteErrorMsg(UserIndex, OUTDATED_VERSION)
            isValidData = False
        End If

        ' Nightw TODO:
        Debug.Print "Login Previous Token: " & sPreviousToken
        
        ' No previous token?
        If Len(sPreviousToken) = 1 Then sPreviousToken = vbNullString
        
        If isValidData Then
            If modAccount.connect(UserIndex, AccountName, Password, sPreviousToken, , sClientTempCode) Then
                    ' If the user connected succesfully, then send the token.
                    Call SecurityIp.IpCleanConnectionInterval(.IPLong)

                Else
                Call CloseSocket(UserIndex)
            End If
        Else
            Call CloseSocket(UserIndex)
        End If
        
        Call CloseSocket(UserIndex)
            
    End With
    
    Exit Sub
    
ErrHandler:

    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleAccountLogin de Protocol.bas")
End Sub

Private Sub HandleAccountCreate(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Creation Date: 23/01/2014
'Last Modification: 17/06/2014 (D'Artagnan)
'Creates Account
'17/06/2014: D'Artagnan - Show the account form if succeed.
'***************************************************

On Error GoTo ErrHandler
    
    With UserList(UserIndex)

        Dim AccountName As String
        Dim Password As String
        Dim Email As String
        Dim SecretQuestion As String
        Dim Answer As String
        Dim sVersion As String

        AccountName = Reader.ReadString8()
    
        Password = Reader.ReadString8()

        Email = Reader.ReadString8()
        SecretQuestion = Reader.ReadString8()
        Answer = Reader.ReadString8()
        sVersion = CStr(Reader.ReadInt8()) & "." & CStr(Reader.ReadInt8()) & "." & CStr(Reader.ReadInt8())
        .ClientTempCode = Reader.ReadString8()

        If Not VersionOK(sVersion) Then
            Call WriteCloseForm(UserIndex, "frmAccountCreate")
            Call DisconnectWithMessage(UserIndex, OUTDATED_VERSION)
            Exit Sub
        End If
        
        If modAccount.CreateAccount(UserIndex, AccountName, Password, SecretQuestion, Answer, Email) Then
        
            ' An external account validation process is a bit more complex, requiring the user to interact with an email
            ' If we're using this complex validation process, the account will remain inactive, and we should send the user back to the
            ' account login screen.
            If ServerConfiguration.UseExternalAccountValidation Then
                Call Protocol.WriteErrorMsg(UserIndex, "Su cuenta ha sido creada con exito y un correo electrónico fue enviado a su casilla de correo. Por favor, complete el proceso de verificación antes de ingresar al juego.", True)
                Call WriteCloseForm(UserIndex, "frmAccountCreate")
            Else
                Call modAccount.connect(UserIndex, AccountName, Password, vbNullString, True, .ClientTempCode)
            End If
        End If
         
    End With
    
    Exit Sub
    
ErrHandler:

    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleAccountCreate de Protocol.bas")
End Sub
 
Private Sub HandleAccountDeleteChar(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Creation Date: 23/01/2014
'Last Modification: 19/07/2014 (D'Artagnan)
'Deletes char slot.
'***************************************************

On Error GoTo ErrHandler
      
    With UserList(UserIndex)
    
        Dim SessionToken As String
        Dim CharSlot As Byte
        Dim sVersion As String
        Dim accToken As String
        Dim SessionIndex As Integer

        CharSlot = Reader.ReadInt8()
        sVersion = CStr(Reader.ReadInt8()) & "." & CStr(Reader.ReadInt8()) & "." & _
                   CStr(Reader.ReadInt8())
        SessionToken = Reader.ReadString8()
        .ClientTempCode = Reader.ReadString8()
        accToken = Reader.ReadString8()

        If Not VersionOK(sVersion) Then
            Call DisconnectWithMessage(UserIndex, OUTDATED_VERSION)
            Exit Sub
        End If
        
        ' Check if the session is valid
        If Not modSession.IsTokenValid(SessionToken, UserIndex, SessionIndex) Then
            Call WriteCloseForm(UserIndex, "frmAccount")
            Call DisconnectWithMessage(UserIndex, "Su sesión es inválida. Por favor vuelva a ingresar.")
            Exit Sub
        End If
        
        ' Check if the session expired
        If modSession.SessionExpired(SessionIndex) Then
            Call WriteCloseForm(UserIndex, "frmAccount")
            Call modSession.CleanSessionSlot(SessionIndex)
            Call DisconnectWithMessage(UserIndex, "Su sesión ha expirado. Por favor vuelva a ingresar.")
            Exit Sub
        End If
        
        ' Valid slot?
        If CharSlot > MAX_ACCOUNT_CHARS Then Exit Sub
        
        Call modAccount.AccountChar_Delete(UserIndex, SessionIndex, CharSlot, accToken)
        Call modSession.ExtendSessionLifetime(SessionIndex)
    End With
 
    Exit Sub
    
ErrHandler:

    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleAccountDeleteChar de Protocol.bas")
End Sub
 
Private Sub HandleAccountRecover(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Creation Date: 23/01/2014
'Last Modification: 23/01/2014
'Recovers account.
'***************************************************

On Error GoTo ErrHandler

    Dim RequestingQuestion As Boolean
    Dim sVersion As String
       
    Dim AccountName As String
    Dim Email As String
    Dim UserToken As String
    Dim TmpString As String
    
    AccountName = Reader.ReadString8()
    
    TmpString = Reader.ReadString8
        
    sVersion = CStr(Reader.ReadInt8()) & "." & CStr(Reader.ReadInt8()) & "." & CStr(Reader.ReadInt8())
    UserList(UserIndex).ClientTempCode = Reader.ReadString8()

    If Not VersionOK(sVersion) Then
        Call DisconnectWithMessage(UserIndex, OUTDATED_VERSION)
        Exit Sub
    End If
    
    ' If we're using the external account validation, then we don't need to continue the process, as this code should
    ' have never been reached in the first place.
    If ServerConfiguration.UseExternalAccountValidation Then
        Exit Sub
    End If
    
    ' Account exists?
    Dim sSecretQuestion As String
    Dim sSecretAnswer As String
    Dim AccountId As Long
    Dim accountEmail As String
    
    If ServerConfiguration.UseExternalAccountValidation Then
        Email = TmpString
    Else
        UserToken = UCase$(TmpString)
    End If

    AccountId = GetAccountID(AccountName, , accountEmail, sSecretQuestion, sSecretAnswer)
    If AccountId = 0 Then
        Call DisconnectWithMessage(UserIndex, "Alguno de los datos es incorrecto.")
        Exit Sub
    End If
        
    ' valid secret code? We are using the "ANSWER" field to store the token.
    If UCase$(UserToken) <> UCase$(sSecretAnswer) Then
        Call DisconnectWithMessage(UserIndex, "Alguno de los datos es incorrecto.")
        Exit Sub
    End If
    
    'TODO: Nightw: Is this needed? This might cause issues with connections generated after this packet is processed
    ' keeping this accountid.
    UserList(UserIndex).AccountId = AccountId
    UserList(UserIndex).AccountName = AccountName
    
    Dim NewPass As Integer
    NewPass = RandomNumber(1555, 6666)
    
    ' Update Password
    Call UpdateAccountPassword(UserList(UserIndex).AccountId, CStr(NewPass), False)
    
    Call WriteCloseForm(UserIndex, "frmAccountRecover")
    Call DisconnectWithMessage(UserIndex, "¡Has recuperado la cuenta! La nueva contraseña de " & UserList(UserIndex).AccountName & " es: " & CStr(NewPass) & ".")
    
    
    'Ya procesamos los datos, cierra el sub.
    Exit Sub
   
    
 
 
    Exit Sub
    
ErrHandler:

    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleAccountRecover de Protocol.bas")
End Sub

Private Sub HandleAccountChangePassword(ByVal UserIndex As Integer)
'***************************************************
'Author: D'Artagnan
'Creation Date: 07/11/2014
'Last Modification: 07/11/2014
'
'***************************************************

On Error GoTo ErrHandler

    Dim sCurrentPassword As String
    Dim sCurrentPasswordInput As String
    Dim sNewPassword As String
    Dim SessionToken As String
    Dim SessionIndex As Integer
        
    With UserList(UserIndex)

        sCurrentPasswordInput = Reader.ReadString8()
        sNewPassword = Reader.ReadString8()

        SessionToken = Reader.ReadString8()
        .ClientTempCode = Reader.ReadString8()
        
        ' TODO: Need to validate the version sent by the client.
               
        ' Check if the session is valid
        If Not modSession.IsTokenValid(SessionToken, UserIndex, SessionIndex) Then
            Call WriteCloseForm(UserIndex, "frmAccount")
            Call DisconnectWithMessage(UserIndex, "Su sesión es inválida. Por favor vuelva a ingresar.")
            Exit Sub
        End If
        
        ' Check if the session expired
        If modSession.SessionExpired(SessionIndex) Then
            Call WriteCloseForm(UserIndex, "frmAccountChangePassword")
            Call WriteCloseForm(UserIndex, "frmAccount")
            Call modSession.CleanSessionSlot(SessionIndex)
            Call DisconnectWithMessage(UserIndex, "Su sesión ha expirado. Por favor vuelva a ingresar.")
            Exit Sub
        End If
        
        ' Get the account data.
        Dim AccountData As tAccountData
        AccountData = GetAccountById(modSession.aActiveSessions(SessionIndex).nAccountID)
        
        ' Current password must match
        If AccountData.Password <> sCurrentPasswordInput Then
            Call DisconnectWithMessage(UserIndex, "Contraseña incorrecta.")
            Exit Sub
        End If
        
        ' Current password must be different than the new password
        If AccountData.Password = sNewPassword Then
            Call DisconnectWithMessage(UserIndex, "La nueva contraseña debe ser diferente a la anterior.")
            Exit Sub
        End If
        
        Call UpdateAccountPassword(AccountData.Id, sNewPassword, True)

            ' TODO: Send password update signal
            If modMessageQueueProxy.IsProxyServerOnline() Then
            Call modMessageQueueProxy.SendAccountPasswordChangedMessage(AccountData.Id, AccountData.Name, AccountData.Email, sNewPassword, .IP)
        End If

        Call WriteCloseForm(UserIndex, "frmAccountChangePassword")
        Call DisconnectWithMessage(UserIndex, "La contraseña ha sido cambiada con éxito.")
        
        Call modSession.ExtendSessionLifetime(SessionIndex)
                
        ' Clean the connection interval to allow the user to connect again immediately.
        Call SecurityIp.IpCleanConnectionInterval(.IPLong)
            
    End With
     
    Exit Sub
    
ErrHandler:

    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleAccountChangePassword de Protocol.bas")
End Sub
 
Public Sub WriteAccountPersonaje(ByVal UserIndex As Integer, ByVal CharSlot As Byte, _
    ByRef CharDetail As Char_Acc_Data)
'***************************************************
'Author: ZaMa
'Creation Date: 23/01/2014
'Last Modification: 23/01/2014
'Sends account char detail.
'***************************************************
On Error GoTo ErrHandler
  
 
    
         Call Writer.WriteInt8(ServerPacketID.AccountPersonaje)
         Call Writer.WriteInt8(CharSlot)
         
         Call Writer.WriteString8(CharDetail.Nick_Name)
         Call Writer.WriteString8(CharDetail.Pos_Map)
         Call Writer.WriteBool(CharDetail.Muerto)
         Call Writer.WriteBool(CharDetail.bSailing)
         Call Writer.WriteInt8(CharDetail.Nivel)
         Call Writer.WriteInt8(CharDetail.Alignment)
         Call Writer.WriteInt(CharDetail.IdGuild)
         Call Writer.WriteString8(CharDetail.GuildName)
         Call Writer.WriteInt(CharDetail.JailRemainingTime)
         Call Writer.WriteBool(CharDetail.Banned)
         
         Call Writer.WriteInt16(CharDetail.Character.body)
         Call Writer.WriteInt16(CharDetail.Character.head)
         Call Writer.WriteInt16(CharDetail.Character.WeaponAnim)
         Call Writer.WriteInt16(CharDetail.Character.ShieldAnim)
         Call Writer.WriteInt16(CharDetail.Character.CascoAnim)
       
    Call SendData(ToUser, UserIndex, vbNullString)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteAccountPersonaje de Protocol.bas")
End Sub


Public Sub WriteAccountRemoveChar(ByVal UserIndex As Integer, ByVal nCharSlot As Byte)
'***************************************************
'Author: D'Artagnan
'Date: 20/06/2014
'Last Modification: 20/06/2014
'Remove the character at the specified slot from the account form.
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ServerPacketID.AccountRemoveChar)
        Call Writer.WriteInt8(nCharSlot)
        
    Call SendData(ToUser, UserIndex, vbNullString)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteAccountRemoveChar de Protocol.bas")
End Sub

Public Sub WriteAccountShow(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Creation Date: 23/01/2014
'Last Modification: 20/06/2014 (D'Artagnan)
'Shows Account form.
'20/06/2014: D'Artagnan - Send account session token.
'***************************************************
On Error GoTo ErrHandler
  
    Writer.WriteInt8 (ServerPacketID.AccountShow)
          
    Call SendData(ToUser, UserIndex, vbNullString)
    
    Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteAccountShow de Protocol.bas")
End Sub

Public Sub WriteLoginScreenShow(ByVal UserIndex As Integer)
'***************************************************
'Author: Nightw
'Creation Date: 13/10/2014
'Shows Login form.
'***************************************************
On Error GoTo ErrHandler
  

    
         Call Writer.WriteInt8(ServerPacketID.LoginScreenShow)
       
    Call SendData(ToUser, UserIndex, vbNullString)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteLoginScreenShow de Protocol.bas")
End Sub

Public Sub WriteCloseForm(ByVal UserIndex As Integer, ByRef sFormName As String)
'***************************************************
'Author: D'Artagnan
'Creation Date: 07/11/2014
'Last Modification: 07/11/2014
'Close the specified form.
'***************************************************
On Error GoTo ErrHandler
  
    
        Call Writer.WriteInt8(ServerPacketID.CloseForm)
        Call Writer.WriteString8(sFormName)
       
    Call SendData(ToUser, UserIndex, vbNullString)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteCloseForm de Protocol.bas")
End Sub
 
Public Sub WriteAccountQuestion(ByVal UserIndex As Integer, ByRef sQuestion As String)
'***************************************************
'Author: ZaMa
'Creation Date: 23/01/2014
'Last Modification: 23/01/2014
'Sends secret question
'***************************************************
On Error GoTo ErrHandler
  
 
    
         Call Writer.WriteInt8(ServerPacketID.AccountQuestion)
         Call Writer.WriteString8(sQuestion)
       
    Call SendData(ToUser, UserIndex, vbNullString)
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteAccountQuestion de Protocol.bas")
End Sub

Private Sub HandleChatDesafio(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 03/12/2009
'
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim Chat As String
        
        Chat = Reader.ReadString8()

        If LenB(Chat) <> 0 Then
            'Analize chat...
            Call Statistics.ParseChat(Chat)
                
            Chat = UserList(UserIndex).Name & "> " & Chat
            
            ' Acá faltaría WriteWarMsg()
            'Call WriteWarMsg(UserIndex, Chat, FontTypeNames.FONTTYPE_PARTY)
            'Call WriteWarMsg(UserList(UserIndex).Challenge.IndexOther, Chat, FontTypeNames.FONTTYPE_PARTY)

        End If

    End With
     
    Exit Sub
    
ErrHandler:

    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleChatDesafio de Protocol.bas")
End Sub

Private Sub HandleAceptDesafio(ByVal UserIndex As Integer)
On Error GoTo ErrHandler
  
    Dim OtherUser As Integer
    
    With UserList(UserIndex)

        OtherUser = .Challenge.IndexOther
        
        'logeado
        If UserList(OtherUser).flags.UserLogged Then

            If UserList(OtherUser).Challenge.Aceptar = True Then
                'comenzar evento
                Call Start_challenge(UserIndex, OtherUser, .Challenge.InSand)
            Else
                .Challenge.Aceptar = True
                'Call WriteGuildWar(UserIndex, 4) 'inhabilito al usuario
                'Call WriteGuildWar(OtherUser, 3) 'habilito al otro
                Call WriteConsoleMsg(UserIndex, "WriteGuildWar() no existe.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleAceptDesafio de Protocol.bas")
End Sub

Private Sub HandleCancelDesafio(ByVal UserIndex As Integer)
On Error GoTo ErrHandler
  
    Dim OtherUser As Integer
    
    With UserList(UserIndex)

        OtherUser = .Challenge.IndexOther
        
        Call Cancel_challenge(UserIndex)
        Call Cancel_challenge(OtherUser)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleCancelDesafio de Protocol.bas")
End Sub

Private Sub HandleDatosDesafio(ByVal UserIndex As Integer)

On Error GoTo ErrHandler

    With UserList(UserIndex)
        Dim Amount_gold As Long
        Dim Maxim_dead As Byte
        Dim Event_time As Byte
        Dim Time_start As Byte
        Dim Event_map As Byte
        Dim Invisibility As Byte
        Dim Resucitar As Byte
        Dim Elementary As Byte

        Amount_gold = Reader.ReadInt32()
        Maxim_dead = Reader.ReadInt8()
        Event_time = Reader.ReadInt8()
        Time_start = Reader.ReadInt8()
        Event_map = Reader.ReadInt8()
        Invisibility = Reader.ReadInt8()
        Resucitar = Reader.ReadInt8()
        Elementary = Reader.ReadInt8()

        'cotrolar valores asignados
        If Not MapaValido(Event_map) Or MapInfo(Event_map).Pk = False Then
            Call WriteConsoleMsg(UserIndex, "El mapa asignado no es válido.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If .Stats.GLD < Amount_gold Then
            Call WriteConsoleMsg(UserIndex, "No tienes suficientes monedas de oro.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'enviar a la arena
        SandsChallenge(.Challenge.InSand).InUse = 1
        SandsChallenge(.Challenge.InSand).Amount_gold = Amount_gold
        
        SandsChallenge(.Challenge.InSand).Elementary = Elementary
        SandsChallenge(.Challenge.InSand).Resucitar = Resucitar
        SandsChallenge(.Challenge.InSand).Invisibility = Invisibility
        
        SandsChallenge(.Challenge.InSand).Time_start = Time_start
        SandsChallenge(.Challenge.InSand).Event_map = Event_map
        SandsChallenge(.Challenge.InSand).Event_time = Event_time
        
        SandsChallenge(.Challenge.InSand).Maxim_dead = Maxim_dead
        
        'enviar info
        Call WriteUpdateChallenge(.Challenge.IndexOther)
        
        'inhabilitar en enviar
        'Call WriteGuildWar(UserIndex, 2) 'inhabilito el enviar
        'Call WriteGuildWar(.Challenge.IndexOther, 3) 'le habilito el aceptar al otro
        Call WriteConsoleMsg(UserIndex, "WriteGuildWar() no existe.", FontTypeNames.FONTTYPE_INFO)
    End With
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleDatosDesafio de Protocol.bas")
End Sub

Public Sub HandleGetPunishmentList(ByVal UserIndex As Integer)
'***************************************************
'Author: Nightw
'Creation Date: 30/09/2014
'Last Modification:
' Handles the petition for sending the punishment types
'***************************************************

On Error GoTo ErrHandler
    Dim UserName As String
    Dim UserId As Long
    Dim PunishmentSubtype As Byte

    PunishmentSubtype = Reader.ReadInt8()
    UserName = Reader.ReadString8()

    If (Not UserList(UserIndex).flags.Privilegios And PlayerType.RoleMaster) <> 0 And (Not UserList(UserIndex).flags.Privilegios And PlayerType.User) <> 0 Then
        'Is the user a semi, dios or admin?
        Call WritePunishmentTypeList(UserIndex, UserName, PunishmentSubtype)
    End If
    
    Exit Sub

ErrHandler:

    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleGetPunishmentList de Protocol.bas")
End Sub

Public Sub WritePunishmentTypeList(ByVal UserIndex As Integer, ByVal UserName As String, ByVal PunishmentSubtype As Integer)

    Dim I As Integer, J As Integer
    Dim list() As tPunishmentType
            
    Select Case PunishmentSubtype
        Case ePunishmentSubType.Ban
            list = listBanTypes
        Case ePunishmentSubType.Jail
            list = listJailTypes
        Case ePunishmentSubType.Warning
            list = listWarningTypes
    End Select
            
            
    Call Writer.WriteInt8(ServerPacketID.PunishmentTypeList)
    
    If UBound(list) >= 0 Then
        Call Writer.WriteInt8(PunishmentSubtype)
        Call Writer.WriteString8(UserName)
               
        ' Write the amount of punishment types
        Call Writer.WriteInt16(UBound(list))
        
        For I = 0 To UBound(list) - 1
            Call Writer.WriteInt16(list(I).Id)
            Call Writer.WriteString8(list(I).Name)
        Next I
    End If
            
    Call SendData(ToUser, UserIndex, vbNullString)
            
End Sub

Public Sub WriteBerserkerEnabled(ByVal UserIndex As Integer, ByVal Enabled As Boolean)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BerserkerEnabled" message to the given user's outgoing data buffer
'***************************************************

    
        Call Writer.WriteInt8(ServerPacketID.EnableBerserker)
        Call Writer.WriteBool(Enabled)
       
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

Private Sub HandleDueloPublico(ByVal UserIndex As Integer)
'---------------------------------------------------------------------------------------
' Module    : Protocol
' Author    : Anagrama
' Date      : 19/08/2016
' Purpose   : Ingresa a la lista de espera de duelos.
'---------------------------------------------------------------------------------------
On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        Call IngresarDueloPublico(UserIndex)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleDueloPublico de Protocol.bas")
End Sub

Public Sub WriteOkDueloPublico(ByVal UserIndex As Integer)
'---------------------------------------------------------------------------------------
' Module    : Protocol
' Author    : Anagrama
' Date      : 19/08/2016
' Purpose   : Envia el OK que dice que pudo entrar a la lista y ahora esta esperando.
'---------------------------------------------------------------------------------------

    
        Call Writer.WriteInt8(ServerPacketID.OkDueloPublico)
       
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

Private Sub HandleCancelarEspera(ByVal UserIndex As Integer)
'---------------------------------------------------------------------------------------
' Module    : Protocol
' Author    : Anagrama
' Date      : 19/08/2016
' Purpose   : Sale de la lista de espera de duelos.
'---------------------------------------------------------------------------------------
On Error GoTo ErrHandler
  
    With UserList(UserIndex)
        Call CancelarEsperaDuelo(UserIndex)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleCancelarEspera de Protocol.bas")
End Sub

Private Sub HandleDuelos(ByVal UserIndex As Integer)
'---------------------------------------------------------------------------------------
' Module    : Protocol
' Author    : Anagrama
' Date      : 19/08/2016
' Purpose   : Recibe la peticion para enviar un duelo.
'---------------------------------------------------------------------------------------
On Error GoTo ErrHandler
  
    If PuedeDuelo(UserIndex, 0) Then
        Call WriteDuelos(UserIndex)
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleDuelos de Protocol.bas")
End Sub

Private Sub HandleAceptarDuelo(ByVal UserIndex As Integer)
'---------------------------------------------------------------------------------------
' Module    : Protocol
' Author    : Anagrama
' Date      : 19/08/2016
' Purpose   : Acepta el duelo.
'---------------------------------------------------------------------------------------
On Error GoTo ErrHandler

    If UserList(UserIndex).flags.DueloIndex > 0 Then
        If Not DuelData.Duelo(UserList(UserIndex).flags.DueloIndex).estado = eDuelState.Esperando_Jugadores Then Exit Sub
        If GetUserTeam(UserList(UserIndex).flags.DueloIndex, UserIndex) > 0 Then Exit Sub
        If PuedeAceptarDuelo(UserIndex, UserList(UserIndex).flags.DueloIndex) Then
            Call AssignTeamMember(UserList(UserIndex).flags.DueloIndex, UserList(UserIndex).flags.DueloTeam, UserIndex)
        End If
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleAceptarDuelo de Protocol.bas")
End Sub

Private Sub HandleRechazarDuelo(ByVal UserIndex As Integer)
'---------------------------------------------------------------------------------------
' Module    : Protocol
' Author    : Anagrama
' Date      : 19/08/2016
' Purpose   : Rechaza el duelo.
'---------------------------------------------------------------------------------------
On Error GoTo ErrHandler

    If UserList(UserIndex).flags.DueloIndex > 0 Then
        Dim DIndex As Byte
        DIndex = UserList(UserIndex).flags.DueloIndex
        
        If Not DuelData.Duelo(DIndex).estado = eDuelState.Esperando_Jugadores Then Exit Sub
        
        If GetUserTeam(DIndex, UserIndex) > 0 Then Exit Sub
        UserList(UserIndex).flags.DueloIndex = 0
        UserList(UserIndex).flags.DueloTeam = 0
        Call CancelarDuelo(DIndex)
        Call WriteConsoleMsg(UserIndex, "Has rechazado la invitación al Duelo.", FontTypeNames.FONTTYPE_INFO)
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleRechazarDuelo de Protocol.bas")
End Sub

Private Sub HandleRetar(ByVal UserIndex As Integer)
'---------------------------------------------------------------------------------------
' Module    : Protocol
' Author    : Anagrama
' Date      : 19/08/2016
' Purpose   : Recibe la peticion de duelo.
'---------------------------------------------------------------------------------------

On Error GoTo ErrHandler

        Dim Amount As Long
        Dim Gold   As Long
        Dim Drop   As Boolean
        Dim Nicks(1 To 7) As String
        Dim Resucitate As Boolean
        
        Amount = Reader.ReadInt8()
        Gold = Reader.ReadInt32()
        Drop = Reader.ReadBool()
        
        Select Case Amount
            Case 1
                Nicks(1) = Reader.ReadString8()
                
                Call PeticionDuelo(UserIndex, Amount, Gold, Drop, Nicks(1))
            Case 2
                Nicks(1) = Reader.ReadString8()
                Nicks(2) = Reader.ReadString8()
                Nicks(3) = Reader.ReadString8()
                Resucitate = Reader.ReadBool()
                
                Call PeticionDuelo(UserIndex, Amount, Gold, Drop, Nicks(1), Nicks(2), Nicks(3), Resucitate)
            Case 3
                Nicks(1) = Reader.ReadString8()
                Nicks(2) = Reader.ReadString8()
                Nicks(3) = Reader.ReadString8()
                Resucitate = Reader.ReadBool()
                Nicks(4) = Reader.ReadString8()
                Nicks(5) = Reader.ReadString8()
                
                Call PeticionDuelo(UserIndex, Amount, Gold, Drop, Nicks(1), Nicks(2), Nicks(3), Resucitate, Nicks(4), Nicks(5))
            Case 4
                Nicks(1) = Reader.ReadString8()
                Nicks(2) = Reader.ReadString8()
                Nicks(3) = Reader.ReadString8()
                Resucitate = Reader.ReadBool()
                Nicks(4) = Reader.ReadString8()
                Nicks(5) = Reader.ReadString8()
                Nicks(6) = Reader.ReadString8()
                Nicks(7) = Reader.ReadString8()
                
                Call PeticionDuelo(UserIndex, Amount, Gold, Drop, Nicks(1), Nicks(2), Nicks(3), Resucitate, Nicks(4), Nicks(5), Nicks(6), Nicks(7))
        End Select

    
    Exit Sub

ErrHandler:

    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleRetar de Protocol.bas")
End Sub

Public Sub WriteDuelos(ByVal UserIndex As Integer)
'---------------------------------------------------------------------------------------
' Module    : Protocol
' Author    : Anagrama
' Date      : 19/08/2016
' Purpose   : Envia la confirmacion para enviar un duelo.
'---------------------------------------------------------------------------------------

    Call Writer.WriteInt8(ServerPacketID.Retar)
       
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

Public Sub WriteMensajeDuelo(ByVal UserIndex As Integer, ByVal Slot As Byte, Optional ByVal TeamMate As Boolean, _
                            Optional ByVal P1 As String, Optional ByVal P2 As String, Optional ByVal P3 As String, _
                            Optional ByVal P4 As String, Optional ByVal P5 As String, Optional ByVal P6 As String, _
                            Optional ByVal P7 As String)
'---------------------------------------------------------------------------------------
' Module    : Protocol
' Author    : Anagrama
' Date      : 19/08/2016
' Purpose   : Envia el mensaje de duelo con toda la informacion.
'---------------------------------------------------------------------------------------

    With UserList(UserIndex)
        Call Writer.WriteInt8(ServerPacketID.MensajeDuelo)
        
        Call Writer.WriteInt8(CByte(GetTipoDuelo(Slot)))
        Select Case GetTipoDuelo(Slot)
            Case eDuelType.vs1
                Call Writer.WriteString8(P1)
                Call Writer.WriteInt32(DuelData.Duelo(Slot).Oro)
                Call Writer.WriteBool(DuelData.Duelo(Slot).Drop)
            Case eDuelType.vs2
                Call Writer.WriteString8(P1)
                Call Writer.WriteString8(P2)
                Call Writer.WriteString8(P3)
                Call Writer.WriteInt32(DuelData.Duelo(Slot).Oro)
                Call Writer.WriteBool(DuelData.Duelo(Slot).Drop)
                Call Writer.WriteBool(DuelData.Duelo(Slot).Resucitar)
                Call Writer.WriteBool(TeamMate)
            Case eDuelType.vs3
                Call Writer.WriteString8(P1)
                Call Writer.WriteString8(P2)
                Call Writer.WriteString8(P3)
                Call Writer.WriteString8(P4)
                Call Writer.WriteString8(P5)
                Call Writer.WriteInt32(DuelData.Duelo(Slot).Oro)
                Call Writer.WriteBool(DuelData.Duelo(Slot).Drop)
                Call Writer.WriteBool(DuelData.Duelo(Slot).Resucitar)
                Call Writer.WriteBool(TeamMate)
            Case eDuelType.vs4
                Call Writer.WriteString8(P1)
                Call Writer.WriteString8(P2)
                Call Writer.WriteString8(P3)
                Call Writer.WriteString8(P4)
                Call Writer.WriteString8(P5)
                Call Writer.WriteString8(P6)
                Call Writer.WriteString8(P7)
                Call Writer.WriteInt32(DuelData.Duelo(Slot).Oro)
                Call Writer.WriteBool(DuelData.Duelo(Slot).Drop)
                Call Writer.WriteBool(DuelData.Duelo(Slot).Resucitar)
                Call Writer.WriteBool(TeamMate)
        End Select
    End With
       
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

Private Sub HandleCancelarElDuelo(ByVal UserIndex As Integer)
'---------------------------------------------------------------------------------------
' Module    : Protocol
' Author    : Anagrama
' Date      : 19/08/2016
' Purpose   : Cancela el duelo actual.
'---------------------------------------------------------------------------------------
On Error GoTo ErrHandler

    If UserList(UserIndex).flags.DueloIndex > 0 Then
        Dim DIndex As Byte
        DIndex = UserList(UserIndex).flags.DueloIndex
        
        If Not DuelData.Duelo(DIndex).estado = eDuelState.Esperando_Jugadores Then Exit Sub

        If Not GetUserTeam(DIndex, UserIndex) = 1 Then Exit Sub
        If Not GetTeamSlot(DIndex, GetUserTeam(DIndex, UserIndex), UserIndex) = 1 Then Exit Sub
        
        UserList(UserIndex).flags.DueloIndex = 0
        UserList(UserIndex).flags.DueloTeam = 0
        Call CancelarDuelo(DIndex)
        Call WriteConsoleMsg(UserIndex, "Has cancelado el duelo.", FontTypeNames.FONTTYPE_INFO)
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleCancelarElDuelo de Protocol.bas")
End Sub


Public Sub WriteChangeAccBankSlot(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal ObjIndex As Integer, _
    ByVal Amount As Integer, ByVal CanUse As Boolean)
'***************************************************
'Author: Anagrama
'Last Modification: 22/08/2016
'Envia el slot de la boveda de cuenta
'***************************************************

    
        Call Writer.WriteInt8(ServerPacketID.AccBankChangeSlot)
        Call Writer.WriteInt8(Slot)
        Call Writer.WriteInt16(ObjIndex)
        Call Writer.WriteInt16(Amount)
        Call Writer.WriteBool(CanUse)
       
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

Public Sub WriteAccBankInit(ByVal UserIndex As Integer)
'***************************************************
'Author: Anagrama
'Last Modification: 22/08/2016
'Inicia la boveda de cuenta
'***************************************************

    Call Writer.WriteInt8(ServerPacketID.AccBankInit)
    Call Writer.WriteInt32(BovedaCuenta(UserList(UserIndex).flags.AccountBank).Oro)
       
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

Private Sub HandleAccBankExtractItem(ByVal UserIndex As Integer)
'***************************************************
'Author: Anagrama
'Last Modification: 22/08/2016
'Usuario retira objeto de la boveda
'***************************************************

On Error GoTo ErrHandler
  

    With UserList(UserIndex)

        Dim Slot As Byte
        Dim Amount As Integer
        
        Slot = Reader.ReadInt8()
        Amount = Reader.ReadInt16()
        
        'Dead people can't commerce
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        '¿El target es un NPC valido?
        If .flags.TargetNpc < 1 Then Exit Sub
        
        '¿Es el banquero?
        If Npclist(.flags.TargetNpc).NPCtype <> eNPCType.Banquero Then
            Exit Sub
        End If
        
        'User retira el item del slot
        Call AccBankRetirarItem(UserIndex, Slot, Amount)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleAccBankExtractItem de Protocol.bas")
End Sub

Private Sub HandleAccBankDepositItem(ByVal UserIndex As Integer)
'***************************************************
'Author: Anagrama
'Last Modification: 22/08/2016
'Usuario deposita objeto en al boveda
'***************************************************

On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim Slot As Byte
        Dim Amount As Integer
        
        Slot = Reader.ReadInt8()
        Amount = Reader.ReadInt16()
        
        'Dead people can't commerce...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        '¿El target es un NPC valido?
        If .flags.TargetNpc < 1 Then Exit Sub
        
        '¿El NPC puede comerciar?
        If Npclist(.flags.TargetNpc).NPCtype <> eNPCType.Banquero Then
            Exit Sub
        End If
        
        'User deposita el item del slot rdata
        Call UserDepositaAccBankItem(UserIndex, Slot, Amount)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleAccBankDepositItem de Protocol.bas")
End Sub

Private Sub HandleAccBankExtractGold(ByVal UserIndex As Integer)
'***************************************************
'Author: Anagrama
'Last Modification: 22/08/2016
'Usuario retira oro
'***************************************************

On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim Amount As Long
        
        Amount = Reader.ReadInt32()
        
        'Dead people can't open the bank.
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNpc = 0 Then
             Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
             Exit Sub
        End If
        
        If Npclist(.flags.TargetNpc).NPCtype <> eNPCType.Banquero Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNpc).Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If .flags.AccountBank = 0 Then Exit Sub
        
        If Amount > 0 And Amount <= BovedaCuenta(.flags.AccountBank).Oro Then
             BovedaCuenta(.flags.AccountBank).Oro = BovedaCuenta(.flags.AccountBank).Oro - Amount
             .Stats.GLD = .Stats.GLD + Amount
             Call WriteChatOverHead(UserIndex, "Tenés " & BovedaCuenta(.flags.AccountBank).Oro & " monedas de oro en tu cuenta.", Npclist(.flags.TargetNpc).Char.CharIndex, vbWhite)
        Else
            Call WriteChatOverHead(UserIndex, "No tienes esa cantidad.", Npclist(.flags.TargetNpc).Char.CharIndex, vbWhite)
        End If
        
        Call WriteUpdateGold(UserIndex)
        Call WriteUpdateAccBankGold(UserIndex)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleAccBankExtractGold de Protocol.bas")
End Sub

Private Sub HandleAccBankDepositGold(ByVal UserIndex As Integer)
'***************************************************
'Author: Anagrama
'Last Modification: 22/08/2016
'Usuario deposita oro
'***************************************************

On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim Amount As Long
        
        Amount = Reader.ReadInt32()
        
        'Dead people can't leave a faction.. they can't talk...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNpc = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Distancia(Npclist(.flags.TargetNpc).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Npclist(.flags.TargetNpc).NPCtype <> eNPCType.Banquero Then Exit Sub
        
        If .flags.AccountBank = 0 Then Exit Sub
        
        If Amount > 0 And Amount <= .Stats.GLD Then
            BovedaCuenta(.flags.AccountBank).Oro = BovedaCuenta(.flags.AccountBank).Oro + Amount
            .Stats.GLD = .Stats.GLD - Amount
            Call WriteChatOverHead(UserIndex, "Tenés " & BovedaCuenta(.flags.AccountBank).Oro & " monedas de oro en tu cuenta.", Npclist(.flags.TargetNpc).Char.CharIndex, vbWhite)
            
            Call WriteUpdateGold(UserIndex)
            Call WriteUpdateAccBankGold(UserIndex)
        Else
            Call WriteChatOverHead(UserIndex, "No tenés esa cantidad.", Npclist(.flags.TargetNpc).Char.CharIndex, vbWhite)
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleAccBankDepositGold de Protocol.bas")
End Sub

Public Sub WriteUpdateAccBankGold(ByVal UserIndex As Integer)
'***************************************************
'Author: Anagrama
'Last Modification: 22/08/2016
'Envia el oro de la boveda de cuenta
'***************************************************

    
        Call Writer.WriteInt8(ServerPacketID.AccBankUpdateGold)
        Call Writer.WriteInt32(BovedaCuenta(UserList(UserIndex).flags.AccountBank).Oro)
       
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

Private Sub HandleAccBankStart(ByVal UserIndex As Integer)
'***************************************************
'Author: Anagrama
'Last Modification: 06/04/2017 - G Toyz
'Recibe la solicitud para iniciar la boveda de cuenta
'***************************************************

On Error GoTo ErrHandler
    
    Dim Pass As String

        Pass = Reader.ReadString8()

        If Not CommerceAllowed(UserIndex) Then
            Exit Sub
        End If
        
        'Validate target NPC
        If UserList(UserIndex).flags.TargetNpc > 0 Then
            If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype = eNPCType.Banquero Then
                If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).Pos, UserList(UserIndex).Pos) > 3 Then
                    Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos del banco.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                Call IniciarDepositoAcc(UserIndex, Pass)
            Else
                Call WriteConsoleMsg(UserIndex, "Este " & IIf(Npclist(UserList(UserIndex).flags.TargetNpc).Comercia, "comerciante", "NPC") & " no es un banco.", FontTypeNames.FONTTYPE_INFO)
            End If
        Else
            Call WriteConsoleMsg(UserIndex, "Primero haz click izquierdo sobre el personaje.", FontTypeNames.FONTTYPE_INFO)
        End If

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleAccBankStart de Protocol.bas")
End Sub

Private Sub HandleAccBankEnd(ByVal UserIndex As Integer)
'***************************************************
'Author: Anagrama
'Last Modification: 22/08/2016
'Cierra al boveda en el servidor
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)

        'User exits banking mode
        .flags.Comerciando = 0
        Call WriteAccBankEnd(UserIndex)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleAccBankEnd de Protocol.bas")
End Sub

Public Sub WriteAccBankEnd(ByVal UserIndex As Integer)
'***************************************************
'Author: Anagrama
'Last Modification: 22/08/2016
'Cierra la boveda en el cliente
'***************************************************

    Call Writer.WriteInt8(ServerPacketID.AccBankEnd)
       
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

Private Sub HandleAccBankChangePass(ByVal UserIndex As Integer)
'***************************************************
'Author: Anagrama
'Last Modification: 22/08/2016
'Cambia la contraseña de la boveda de cuenta
'***************************************************

On Error GoTo ErrHandler
  
    Dim Pass As String
    Dim Token As String

        Token = Reader.ReadString8()
        Pass = Reader.ReadString8()
   
        Call ChangeAccBankPass(UserIndex, Token, Pass)

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleAccBankChangePass de Protocol.bas")
End Sub

Public Sub WriteAccBankRequestPass(ByVal UserIndex As Integer)
'***************************************************
'Author: Anagrama
'Last Modification: 22/08/2016
'Solicita la contraseña para ingresar a la boveda
'***************************************************

    Call Writer.WriteInt8(ServerPacketID.AccBankRequestPass)
       
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

Public Sub HandleChangeMapInfoNoInmo(ByVal UserIndex As Integer)
'***************************************************
'Author: Anagrama
'Last Modification: 09/09/2016
'Cambia la posibilidad de usar o no inmo en el mapa
'***************************************************

On Error GoTo ErrHandler
      
    Dim NoInmo As Boolean
    
    With UserList(UserIndex)
        NoInmo = Reader.ReadBool()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            Call LogGM(.Name, .Name & " ha cambiado la información sobre si está permitido usar inmovilizar en el mapa.")
            MapInfo(UserList(UserIndex).Pos.Map).InmovilizarSinEfecto = NoInmo
            Call WriteVar(MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "InmovilizarSinEfecto", NoInmo)
            Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " InmovilizarSinEfecto: " & MapInfo(.Pos.Map).InmovilizarSinEfecto, FontTypeNames.FONTTYPE_INFO)
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleChangeMapInfoNoInmo de Protocol.bas")
End Sub

Public Sub HandleChangeMapInfoMismoBando(ByVal UserIndex As Integer)
'***************************************************
'Author: Anagrama
'Last Modification: 09/09/2016
'Cambia la posibilidad de que ciudas no puedan atacar armadas y crimis no puedan atacar caos y al verre.
'***************************************************

    
On Error GoTo ErrHandler
  
    Dim MismoBando As Boolean
    
    With UserList(UserIndex)
        MismoBando = Reader.ReadBool()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            Call LogGM(.Name, .Name & " ha cambiado la información sobre si está permitido que ciudas ataquen armadas y criminales ataquen caos en el mapa.")
            MapInfo(UserList(UserIndex).Pos.Map).MismoBando = MismoBando
            Call WriteVar(MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "MismoBando", MismoBando)
            Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " MismoBando: " & MapInfo(.Pos.Map).MismoBando, FontTypeNames.FONTTYPE_INFO)
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleChangeMapInfoMismoBando de Protocol.bas")
End Sub

''
' Prepares the "StartEffect" message and returns it.
'
' @param    the effect id to start.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareStartEffect(ByVal effect As Byte, Optional ByVal Arg1 As String) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "GuildChat" message and returns it
'14/07/2016: Anagrama - Ahora envia una lista de los posibles midi/mp3 del mapa.
'***************************************************
On Error GoTo ErrHandler

        Call Writer.WriteInt8(ServerPacketID.StartPresentEffect)
        Call Writer.WriteInt8(effect)
        
        Select Case effect
            Case ePresentEffect.SpawnBoss
                Call Writer.WriteString8(Arg1)
        End Select

  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function PrepareStartEffect de Protocol.bas")
End Function

Private Sub HandleSpawnBoss(ByVal UserIndex As Integer)
On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim BossId As Byte
        Dim BossNpcIndex As Integer
        Dim J As Byte
        Dim BossText As String
        
        BossId = Reader.ReadInt8()

        If BossId < 1 Or BossId > UBound(BossData) Then
            Call WriteConsoleMsg(UserIndex, "El boss que intentas invocar es inválido.", FontTypeNames.FONTTYPE_INFO)
            Call LogGM(.Name, "Intento invocar al boss inválido número:" & BossId)
            Exit Sub
        End If
        
        ' Only Admins
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) = 0 Then
            Exit Sub
                    End If
    
        If BossData(BossId).Alive Then
                Call WriteConsoleMsg(UserIndex, "El boss " & NpcData(BossData(BossId).NpcIndex).Name & " se encuentra vivo.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
        Call modBosses.SpawnBoss(BossId)
            
        Call LogGM(.Name, "Invocó al boss número:" & BossId)

    End With
    
    Exit Sub
    
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function PrepareStartEffect de Protocol.bas")
End Sub

Private Sub HandleCraftItem(ByVal UserIndex As Integer)
'***************************************************
'Author: Anagrama
'Last Modification: 06/04/2017
'Recibe la petición para fabricar un item y su estación de trabajo correspondiente.
'***************************************************

On Error GoTo ErrHandler
  
    With Reader

        Dim CraftingGroup As Byte
        Dim RecipeIndex As Integer
        Dim FromMacro As Boolean
        
        CraftingGroup = .ReadInt8()
        RecipeIndex = .ReadInt()
        FromMacro = .ReadBool()

        If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
        
        If Not PuedeConstruir(UserIndex, CraftingGroup, RecipeIndex) And FromMacro Then
            Call WriteCloseForm(UserIndex, "frmCraft")
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleCraftItem de Protocol.bas")
End Sub

Private Sub HandleSelectPet(ByVal UserIndex As Integer)
'***************************************************
'Author: Anagrama
'Last Modification: 30/07/2017
'Setea la mascota en estado de seleccionada.
'***************************************************

On Error GoTo ErrHandler
  
    With Reader
        Dim MascotaIndex As Byte
        
        MascotaIndex = .ReadInt8()

        If MascotaIndex < 1 Or MascotaIndex > 3 Then
            Call WriteConsoleMsg(UserIndex, "Debes seleccionar una mascota", FontTypeNames.FONTTYPE_INFO, info)
            Exit Sub
        End If
        
        If UserList(UserIndex).TammedPetsCount = 0 Then
            Call WriteConsoleMsg(UserIndex, "No tienes ninguna mascota.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If UserList(UserIndex).TammedPets(MascotaIndex).NpcNumber = 0 Then
            Call WriteConsoleMsg(UserIndex, "No tienes ninguna mascota en ese slot.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        UserList(UserIndex).SelectedPet = MascotaIndex
        
        Call WriteConsoleMsg(UserIndex, "Has seleccionado a tu " & NpcData(UserList(UserIndex).TammedPets(MascotaIndex).NpcNumber).Name & " como mascota predeterminada. Ahora podrás resucitarla e invocarla.", FONTTYPE_INFO, info)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleSelectPet de Protocol.bas")
End Sub

Private Sub HandleRequestPetSelection(ByVal UserIndex As Integer)
'***************************************************
'Author: Anagrama
'Last Modification: 13/08/2017
'Recibe la solicitud para abrir el formulario de selección de mascota.
'***************************************************
On Error GoTo ErrHandler
  
    With Reader
        Call WriteSendPetSelection(UserIndex)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleRequestPetSelection de Protocol.bas")
End Sub

Public Sub WriteSendPetSelection(ByVal UserIndex As Integer)
'***************************************************
'Author: Anagrama
'Last Modification: 13/08/2017
'Envía la mascota seleccionada actualmente para abrir el formulario.
'***************************************************
On Error GoTo ErrHandler
    Dim I As Byte
    
    With UserList(UserIndex)
        Call Writer.WriteInt8(ServerPacketID.SendPetSelection)
        Call Writer.WriteInt8(.SelectedPet)
        
        For I = 1 To Classes(.clase).ClassMods.MaxTammedPets
            If .TammedPets(I).NpcIndex > 0 Then
                Call Writer.WriteString8(NpcData(.TammedPets(I).NpcIndex).Name)
                Call Writer.WriteInt16(NpcData(.TammedPets(I).NpcIndex).Char.body)
                Call Writer.WriteInt16(NpcData(.TammedPets(I).NpcIndex).Char.head)
            Else
                Call Writer.WriteString8("")
                Call Writer.WriteInt16(0)
                Call Writer.WriteInt16(0)
            End If
        Next I
        
    End With
           
    Call SendData(ToUser, UserIndex, vbNullString)
    
Exit Sub

ErrHandler:
        Call LogError("Error " & Err.Number & " (" & Err.Description & ") en Sub WriteSendPetSelection del Módulo Protocol")
End Sub

Public Sub WriteSendPetList(ByVal UserIndex As Integer)
On Error GoTo ErrHandler
    Dim I As Byte
    
    With UserList(UserIndex)
        Call Writer.WriteInt8(ServerPacketID.SendPetList)
        Call Writer.WriteInt8(Classes(.clase).ClassMods.MaxTammedPets)
        
        For I = 1 To Classes(.clase).ClassMods.MaxTammedPets
            If .TammedPets(I).NpcNumber > 0 Then
                Call Writer.WriteInt16(NpcData(.TammedPets(I).NpcNumber).Char.body)
            Else
                Call Writer.WriteInt16(0)
            End If
        Next I
        
    End With
           
    Call SendData(ToUser, UserIndex, vbNullString)
    
Exit Sub

ErrHandler:
        Call LogError("Error " & Err.Number & " (" & Err.Description & ") en Sub WriteSendPetSelection del Módulo Protocol")

End Sub

Public Sub WriteSendSessionToken(ByVal UserIndex As Integer, ByVal Token As Integer)
On Error GoTo ErrHandler

    Call Writer.WriteInt8(ServerPacketID.SendSessionToken)
    Call Writer.WriteString8(aActiveSessions(Token).Token)

    Call SendData(ToUser, UserIndex, vbNullString)
    
    Exit Sub
ErrHandler:
        Call LogError("Error " & Err.Number & " (" & Err.Description & ") en Sub WriteSendPetSelection del Módulo Protocol")
End Sub

Public Sub HandleMasteryAssign(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    Dim MasteryGroup As Integer
    Dim MasteryToAssign As Integer
    
    MasteryGroup = Reader.ReadInt16
    MasteryToAssign = Reader.ReadInt16
    
    ' Different validations
    
    ' Minimum level required
    If UserList(UserIndex).Stats.ELV < ConstantesBalance.MaxLvl Then
        Call WriteConsoleMsg(UserIndex, "Necesitas ser nivel " & ConstantesBalance.MaxLvl & " para aprender una maestría.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    ' Mastery already exists
    If modMasteries.HasMasteryAssigned(UserIndex, MasteryGroup, MasteryToAssign) Then
        Call WriteConsoleMsg(UserIndex, "Ya posees la maestria " & Masteries(MasteryToAssign).Name, FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    ' TODO: Check if there's a requirement of a previously aquired mastery
    If Masteries(MasteryToAssign).Enabled = False Then
        Call WriteConsoleMsg(UserIndex, "La maestria " & Masteries(MasteryToAssign).Name & " se encuentra desactivada", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If UserList(UserIndex).Stats.MasteryPoints < Masteries(MasteryToAssign).PointsRequired Then
        Call WriteConsoleMsg(UserIndex, "No tienes la suficiente cantidad de puntos para obtener esta maestria", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If UserList(UserIndex).Stats.GLD < Masteries(MasteryToAssign).GoldRequired Then
        Call WriteConsoleMsg(UserIndex, "No tienes el suficiente oro para obtener esta maestria", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
        
    ' Aquire the mastery
    Call modMasteries.AquireMastery(UserIndex, MasteryGroup, MasteryToAssign, True)
    
    ' Send the new list of masteries to the user
    Call WriteSendMasteries(UserIndex, CharacterMasteries)
        
Exit Sub

ErrHandler:
        Call LogError("Error " & Err.Number & " (" & Err.Description & ") en Sub HandleMasteryAssign de Protocol.bas")
End Sub

Public Sub WriteSendMasteries(ByVal UserIndex As Integer, ByVal TypeMasteries As eSendMasteryType)
On Error GoTo ErrHandler:
    
    Dim I As Byte
    Dim J As Integer
    Dim FoundMastery As Boolean
    With UserList(UserIndex)

        Call Writer.WriteInt8(ServerPacketID.SendMasteries)
        
        ' Class Masteries or Character Masteries?
        Call Writer.WriteInt8(TypeMasteries)
        
        If TypeMasteries = eSendMasteryType.ClassMasteries Then
                        
            'Send groups qty
            Call Writer.WriteInt16(Classes(.clase).MasteryGroupsQty)
            
            For I = 1 To Classes(.clase).MasteryGroupsQty
                'Send masteries qty
                Call Writer.WriteInt16(Classes(.clase).MasteryGroups(I).MasteriesQty)
                
                For J = 1 To Classes(.clase).MasteryGroups(I).MasteriesQty
                    Call Writer.WriteInt16(Classes(.clase).MasteryGroups(I).Masteries(J))
                Next J
                
            Next I
        Else
            Call Writer.WriteInt16(Classes(.clase).MasteryGroupsQty)
            
            For I = 1 To Classes(.clase).MasteryGroupsQty
                Call Writer.WriteInt16(I) ' Send the mastery group
                
                If I <= .Masteries.GroupsQty Then
                    For J = 1 To Classes(.clase).MasteryGroups(I).MasteriesQty
                        If Not modMasteries.HasMasteryAssigned(UserIndex, I, Classes(.clase).MasteryGroups(I).Masteries(J)) Then
                            Call Writer.WriteInt16(1)
                            Call Writer.WriteInt16(Classes(.clase).MasteryGroups(I).Masteries(J))
                            FoundMastery = True
                            Exit For
                        End If
                    Next J
                    If Not FoundMastery Then
                        ' There's no mastery available on this group, so we send one element with a mastery ID of 0.
                        Call Writer.WriteInt16(1)
                        Call Writer.WriteInt16(0)
                    End If

                Else
                    ' Send only one mastery, the number
                    Call Writer.WriteInt16(1)
                    Call Writer.WriteInt16(Classes(.clase).MasteryGroups(I).Masteries(1))
                End If
                
                FoundMastery = False
            Next I
            
        End If
        
    End With
           
    Call SendData(ToUser, UserIndex, vbNullString)

    Exit Sub
ErrHandler:
    Call LogError("Error " & Err.Number & " (" & Err.Description & ") en Sub WriteSendMasteries de Protocol.bas")
End Sub


Public Sub InitProtocol()
    LAST_CLIENT_PACKET_ID = ClientPacketID.LastClientPacketId - 1
End Sub

Private Sub HandleGuildCreate(ByVal UserIndexLeader As Integer)
On Error GoTo ErrHandler
    
    Dim GuildName As String
    Dim CreatedGuildIndex As Integer
    Dim QtyGoldOfMember As Long
    Dim Alignment As eGuildAlignment
    
    GuildName = Reader.ReadString16()
    
    With UserList(UserIndexLeader)
        
        If .Guild.IdGuild <> 0 Then
            Call WriteErrorMsg(UserIndexLeader, "Ya se encuentra en un Clan, para crear uno nuevo debe abandonar el anterior.", False)
            Exit Sub
        End If
        
        If (.Stats.ELV < GuildConfiguration.CreationLeaderRequiredLevel) Then
            Call WriteErrorMsg(UserIndexLeader, "Necesitas ser nivel " & GuildConfiguration.CreationLeaderRequiredLevel & " para crear un Clan.", False)
            Exit Sub
        End If
        
        If Len(Trim(GuildName)) > MAX_GUILD_NAME_LEN Then
            Call WriteErrorMsg(UserIndexLeader, "El nombre del clan no puede tener más de " & MAX_GUILD_NAME_LEN & " caracteres.", False)
            Exit Sub
        End If
        
        If Not NombrePermitido(GuildName) Or Not AsciiValidos(GuildName, False) Then
            Call WriteErrorMsg(UserIndexLeader, "El nombre seleccionado posee palabras o caracteres no permitidos.", False)
            Exit Sub
        End If
        
        If GuildNameIsUsed(GuildName) Then
            Call WriteErrorMsg(UserIndexLeader, "El nombre del clan ya se encuentra en uso.", False)
            Exit Sub
        End If
        
        If .Stats.GLD < GuildConfiguration.CreationRequiredGold Then
            Call WriteErrorMsg(UserIndexLeader, "No posees suficiente oro, para crear un se necesitan " & GuildConfiguration.CreationRequiredGold & " monedas", False)
            Exit Sub
        End If
                              
        If .Pos.Map <> Npclist(.flags.TargetNpc).Pos.Map Or Not InRangoVision(UserIndexLeader, Npclist(.flags.TargetNpc).Pos.X, Npclist(.flags.TargetNpc).Pos.Y) Then
            Call WriteErrorMsg(UserIndexLeader, "No se encuentra en el rango de vision del maestro de Clanes.", False)
            Exit Sub
        End If
        
        If Not modGuild_Functions.CanUseGuildNameByReservation(.AccountEmail, GuildName) Then
            Call WriteErrorMsg(UserIndexLeader, "No puedes crear un clan con ese nombre ya que es un nombre reservado", False)
            Exit Sub
        End If
                        
        If UserList(UserIndexLeader).Faccion.ArmadaReal = 1 Then
            Alignment = eGuildAlignment.Real
        ElseIf UserList(UserIndexLeader).Faccion.FuerzasCaos = 1 Then
            Alignment = eGuildAlignment.Evil
        Else
            Alignment = eGuildAlignment.Neutral
        End If
        
        CreatedGuildIndex = CreateGuildFromDB(GuildName, .Id, Alignment, ConstantesBalance.GuildRankingStartingPoints)
        
        If CreatedGuildIndex <> -1 Then
            .Guild.IdGuild = GuildList(CreatedGuildIndex).IdGuild
            .Guild.GuildIndex = CreatedGuildIndex
            .Guild.RoleIndex = 1
            .Guild.GuildMemberIndex = 1
            .Guild.RoleId = GetGuildRoleId(.Guild.GuildIndex, UserIndexLeader)
            .Stats.GLD = .Stats.GLD - GuildConfiguration.CreationRequiredGold
            .Guild.GuildMemberIndex = GetMemberIndexOf(UserIndexLeader)
            
            Call WriteUpdateGold(UserIndexLeader)
            Call WriteGuildInfo(UserIndexLeader)
            Call AddOnlineMember(UserIndexLeader)
            Call WriteGuildRolesList(UserIndexLeader)
            Call WriteGuildMembersList(UserIndexLeader)
            Call WriteGuildBankList(UserIndexLeader)
            Call WriteGuildUpgradesList(UserIndexLeader)
            Call WriteGuildUpgradesAcquired(UserIndexLeader)
            Call WriteGuildQuestsCompletedList(UserIndexLeader)
            
            Call WriteGuildCreated(UserIndexLeader)
            Call RefreshCharStatus(UserIndexLeader, False)
            
        End If
        
    End With
    
    Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleGuildCreate de Protocol.bas")
End Sub

Private Sub HandleGuildUserInvitationResponse(ByVal UserIndex As Integer)
On Error GoTo ErrHandler
    Dim InvitationIndex As Integer
    Dim IsAccepted As Integer
    Dim GuildIndex As Integer

    GuildIndex = Reader.ReadInt()
    InvitationIndex = Reader.ReadInt()
    IsAccepted = Reader.ReadBool()
    
    Call modGuild_Functions.UserInvitationResponse(UserIndex, GuildIndex, InvitationIndex, IsAccepted)
    
    Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleGuildUserInvitationResponse de Protocol.bas")
End Sub

Private Sub HandleGuildMember(ByVal UserIndex As Integer)
On Error GoTo ErrHandler
    
    Dim MemberAction As Byte
    Dim UserIndexMember As Long
    Dim UserIndexInvited As Integer
    Dim UserName As String
    Dim GuildIndex As Integer
    Dim GuildId As Long
    Dim I As Integer
    Dim UserIdTarget As Integer
    Dim Accepted As Boolean
    Dim MemberAlignment As eGuildAlignment
    Dim Message As String
    Dim RoleName As String
    Dim KickedName As String
    Dim KickedUserId As Long
    Dim UserKickedIndex As Integer
    
    MemberAction = Reader.ReadInt8()
    KickedUserId = Reader.ReadInt()
    GuildId = Reader.ReadInt16()
    If MemberAction = eMemberAction.SendInvitation Then
        UserName = Reader.ReadString16()
    End If
    Accepted = Reader.ReadBool()
    
    With UserList(UserIndex)
        Select Case MemberAction
                
            Case eMemberAction.KickMember
                
                If Not HasPermission(UserIndex, EGuildPermission.MEMBER_KICK) Then
                    Call WriteConsoleMsg(UserIndex, NOPERMISSIONOFGUILD, FONTTYPE_INFO, info)
                    Exit Sub
                End If
                
                UserKickedIndex = GetUserIndexFromUserId(KickedUserId)
                GuildIndex = GuildIndexOf(GuildId)

                For I = 1 To GuildList(GuildIndex).MemberCount
                   If GuildList(GuildIndex).Members(I).IdUser = KickedUserId Then
                       If GuildList(GuildIndex).Members(I).IdRole = ID_ROLE_LEADER Then
                           'leader cannot be kicked
                           Exit Sub
                       End If
                   End If
                Next I
                
                Call GuildRemoveMember(GuildIndex, KickedUserId, KickedName, True)
                
                If UserKickedIndex > 0 Then
                     Call WriteConsoleMsg(UserKickedIndex, "Has sido expulsado del clan ", FontTypeNames.FONTTYPE_GUILD, eMessageType.Guild)
                End If
                
                If GuildLastMemberOnline(GuildIndex) > 0 Then
                    For I = 1 To GuildLastMemberOnline(GuildIndex)
                        With GuildList(GuildIndex).OnlineMembers(I)
                             Call WriteConsoleMsg(.MemberUserIndex, KickedName & " ha sido expulsado", FontTypeNames.FONTTYPE_GUILD, eMessageType.Guild)
                        End With
                    Next I
                End If
                        
            Case eMemberAction.SendInvitation
                Call modGuild_Functions.InviteUser(UserIndex, UserName)
            Case eMemberAction.LeaveGuild
                
                GuildIndex = .Guild.GuildIndex
                
                If UserList(UserIndex).Guild.RoleId = ID_ROLE_LEADER And GuildList(.Guild.GuildIndex).IdRightHand <= 0 Then
                    Call WriteConsoleMsg(UserIndex, "Para salir del clan primero tienes que asignar una mano derecha", FONTTYPE_INFO, info)
                    Exit Sub
                End If
               
                ' Expell the old member
                Call GuildRemoveMember(GuildIndex, .Id, KickedName)
                
                Call RemoveOnlineMember(UserIndex)
                Call ResetGuildInfo(UserIndex)
                Call WriteConsoleMsg(UserIndex, "Has salido del clan ", FontTypeNames.FONTTYPE_GUILD, eMessageType.Guild)
                Call WriteGuildMemberKicked(UserIndex)
                Call RefreshCharStatus(UserIndex, False)
                
                If GuildLastMemberOnline(GuildIndex) > 0 Then
                    For I = 1 To GuildLastMemberOnline(GuildIndex)
                        With GuildList(GuildIndex).OnlineMembers(I)
                            If UserList(.MemberUserIndex).Id = .IdUser Then
                                Call WriteConsoleMsg(.MemberUserIndex, KickedName & " ha salido del clan", FontTypeNames.FONTTYPE_GUILD, eMessageType.Guild)
                            End If
                        End With
                    Next I
                End If
                
        End Select
            
    End With
    
    Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleGuildMember de Protocol.bas")
End Sub

Private Sub HandleGuildExchange(ByVal UserIndex As Integer)
On Error GoTo ErrHandler

    Dim ExchangeType As Byte
    Dim ExchangeAction As Byte
    Dim Quantity As Long
    Dim GuildIndex As Integer
    Dim Slot As Integer
    Dim Box As Integer
    Dim Success As Boolean
    Dim I As Integer
    
    ExchangeType = Reader.ReadInt8()
    ExchangeAction = Reader.ReadInt8()
    Quantity = Reader.ReadInt32()
    Slot = Reader.ReadInt16()
    Box = Reader.ReadInt16()
    
    Success = False
    
    With UserList(UserIndex)
            GuildIndex = .Guild.GuildIndex
            If (.flags.TargetNpc) = 0 Then Exit Sub
            
             If Npclist(.flags.TargetNpc).NPCtype <> eNPCType.GuildMaster Then
                Exit Sub
            End If
            If Not InRangoVision(UserIndex, Npclist(.flags.TargetNpc).Pos.X, Npclist(.flags.TargetNpc).Pos.Y) Then
                Exit Sub
            End If
            
            Select Case ExchangeType
                
                Case eExchangeType.IsGold
                    
                    Select Case ExchangeAction
                        
                        Case eExchangeAction.Withdraw
                            
                            If Not HasPermission(UserIndex, BANK_WITHDRAW_GOLD) Then
                                Call WriteConsoleMsg(UserIndex, NOPERMISSIONOFGUILD, FONTTYPE_INFO, info)
                                Exit Sub
                            End If
                            
                            If (GuildList(GuildIndex).BankGold < Quantity) Then
                                Quantity = GuildList(GuildIndex).BankGold
                            End If
                            
                            GuildList(GuildIndex).BankGold = GuildList(GuildIndex).BankGold - Quantity
                            .Stats.GLD = .Stats.GLD + Quantity
               
                        Case eExchangeAction.Deposit
                            
                            If Not HasPermission(UserIndex, BANK_DEPOSIT_GOLD) Then
                                Call WriteConsoleMsg(UserIndex, NOPERMISSIONOFGUILD, FONTTYPE_INFO, info)
                                Exit Sub
                            End If
                            
                            If (.Stats.GLD < Quantity) Then
                                Quantity = .Stats.GLD
                            End If
                            
                            GuildList(GuildIndex).BankGold = GuildList(GuildIndex).BankGold + Quantity
                            .Stats.GLD = .Stats.GLD - Quantity
                        
                    End Select
                    
                    Call WriteUpdateGold(UserIndex)
                    
                    If GuildLastMemberOnline(GuildIndex) > 0 Then
                        For I = 1 To GuildLastMemberOnline(GuildIndex)
                            With GuildList(GuildIndex).OnlineMembers(I)
                                If UserList(.MemberUserIndex).Id = .IdUser Then
                                    Call WriteGuildMemberStatusChange(.MemberUserIndex, .IdUser, eChangeMember.GoldGBChange, , GuildList(GuildIndex).BankGold)
                                End If
                            End With
                        Next I
                    End If
                    
                Case eExchangeType.IsObject
                    
                    Select Case ExchangeAction
                    
                        Case eExchangeAction.Withdraw
                            
                            If Not HasPermission(UserIndex, BANK_WITHDRAW_ITEM) Then
                                Call WriteErrorMsg(UserIndex, NOPERMISSIONOFGUILD, False)
                                Exit Sub
                            End If
                                
                            Call UserWithdrawItemFromGuildBank(UserIndex, Slot, Quantity, Box, GuildIndex)

                        
                        Case eExchangeAction.Deposit
                            
                            If Not HasPermission(UserIndex, BANK_DEPOSIT_ITEM) Then
                                Call WriteConsoleMsg(UserIndex, NOPERMISSIONOFGUILD, FONTTYPE_INFO, info)
                                Exit Sub
                            End If
                            
                            Call UserDepositItemInGuildBank(UserIndex, Slot, Quantity, Box, GuildIndex)

                    End Select
                    
            End Select
        
    End With
    

    
    Exit Sub

ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleGuildExchange de Protocol.bas")
End Sub

Private Sub HandleGuildRole(ByVal UserIndex As Integer)
On Error GoTo ErrHandler
    
    Dim ActionType As Byte
    Dim J As Integer
    Dim I As Integer
    Dim GuildIndex As Integer
    Dim TargetUserId As Long
    Dim QtyPermission As Integer
    Dim RoleId As Long
    Dim Permissions() As String
    Dim PermissionStr As String
    Dim RoleName As String
    Dim RoleIndex As Integer
    Dim Message As String
    Dim UserName As String
    
    ActionType = Reader.ReadInt8()
    RoleId = Reader.ReadInt32()
    
    With UserList(UserIndex)
        Select Case ActionType
                    
            Case eRoleAction.Assign
                
                TargetUserId = Reader.ReadInt()
                GuildIndex = .Guild.GuildIndex
                
                If RoleId = ID_ROLE_LEADER Then
                    Exit Sub
                End If
                
                If RoleId = ID_ROLE_RIGHTHAND And Not HasPermission(UserIndex, RIGHT_HAND_ASSIGN) Then
                    Call WriteConsoleMsg(UserIndex, NOPERMISSIONOFGUILD, FONTTYPE_INFO, info)
                    Exit Sub
                End If
                
                If Not HasPermission(UserIndex, ROLE_ASSIGN) Then
                    Call WriteConsoleMsg(UserIndex, NOPERMISSIONOFGUILD, FONTTYPE_INFO, info)
                    Exit Sub
                End If
                    
                'check if target id is leader
                If GuildList(GuildIndex).IdLeader = TargetUserId Then
                    Exit Sub
                End If
                
                ' If the target user is the righthand and we're assigning a role different to the ID_ROLE_RIGHTHAND
                ' Then we should remove it from the guild metadata
                If GuildList(GuildIndex).IdRightHand = TargetUserId And RoleId <> ID_ROLE_RIGHTHAND Then
                    GuildList(GuildIndex).IdRightHand = 0
                End If
                
                Call modGuild_DB.AssignRoleFromDB(.Id, .Guild.IdGuild, TargetUserId, RoleId)
                
                With GuildList(GuildIndex)
                    If RoleId = ID_ROLE_RIGHTHAND Then
                        .IdRightHand = TargetUserId
                    End If
                    
                    RoleIndex = modGuild_Functions.GetRoleIndexFromRoleId(GuildIndex, RoleId)
                    RoleName = .Roles(RoleIndex).RoleName
                    
                    For I = 1 To UBound(.Members)
                        If .Members(I).IdUser = TargetUserId Then
                            .Members(I).IdRole = RoleId
                            .Members(I).RoleIndex = RoleIndex
                            UserName = .Members(I).NameUser
                            Exit For
                        End If
                    Next I
                    
                    Call UpdateGuildLeadership(GuildList(GuildIndex).IdGuild, GuildList(GuildIndex).IdLeader, GuildList(GuildIndex).IdRightHand)
                    
                    If GuildLastMemberOnline(GuildIndex) > 0 Then
                        For I = 1 To GuildLastMemberOnline(GuildIndex)
                            
                            If .OnlineMembers(I).IdUser = TargetUserId Then
                                'update user info
                                UserList(.OnlineMembers(I).MemberUserIndex).Guild.RoleId = RoleId
                                UserList(.OnlineMembers(I).MemberUserIndex).Guild.RoleIndex = RoleIndex
                                
                                Call WriteGuildMemberStatusChange(.OnlineMembers(I).MemberUserIndex, TargetUserId, eChangeMember.RoleChange, RoleId, 0)
                                Message = "Has sido cambiado al rol " & RoleName & "."
                            Else
                                Message = UserName & " ha sido cambiado al rol " & RoleName & "."
                            End If
                            
                            Call WriteGuildMembersList(.OnlineMembers(I).MemberUserIndex)
                            Call WriteGuildInfo(.OnlineMembers(I).MemberUserIndex)
                            
                                Call WriteConsoleMsg(.OnlineMembers(I).MemberUserIndex, Message, FontTypeNames.FONTTYPE_GUILD, eMessageType.Guild)
                        Next I
                    End If
                End With
            Case eRoleAction.Create
                RoleName = Reader.ReadString8
                QtyPermission = Reader.ReadInt8
                
                If QtyPermission > 0 Then
                
                    ReDim Permissions(1 To QtyPermission) As String
                    For I = 1 To QtyPermission
                        Permissions(I) = Reader.ReadString8
                        PermissionStr = PermissionStr & Permissions(I)
                        If I <> QtyPermission Then
                            PermissionStr = PermissionStr & ","
                        End If
                    Next I
                End If
                
                Call modGuild_Functions.RoleUpsert(GuildIndex, RoleId, RoleName, UserIndex, PermissionStr)
                
            Case eRoleAction.Delete
            
                Call modGuild_Functions.RoleDelete(.Guild.GuildIndex, RoleId, UserIndex)
        End Select
    End With
    
    Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleGuildRole de Protocol.bas")
End Sub
Private Sub HandleGuildUpgrade(ByVal UserIndex As Integer)
On Error GoTo ErrHandler
    
    Dim UpgradeId As Byte
    Dim RoleId As Integer
    
    RoleId = UserList(UserIndex).Guild.RoleId
    
    UpgradeId = Reader.ReadInt8()

    Call modGuild_Functions.BuyGuildUpgrade(UserIndex, UpgradeId)
    
    Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleGuildUpgrade de Protocol.bas")
End Sub


Public Sub WriteGuildCreated(ByVal UserIndex As Integer)
On Error GoTo ErrHandler

    Call Writer.WriteInt8(ServerPacketID.GuildCreated)

    Call SendData(ToUser, UserIndex, vbNullString)
    
    Exit Sub
ErrHandler:
        Call LogError("Error " & Err.Number & " (" & Err.Description & ") en Sub WriteGuildCreated del Módulo Protocol")
End Sub

Public Sub WriteGuildInfo(ByVal UserIndex As Integer)
On Error GoTo ErrHandler
    Dim GuildIndex As Integer, MaxMembers As Integer, MaxRoles As Integer, MaxSlotBank As Integer, MaxBoxes As Integer
    Dim MaxContribution As Long
    Dim BankAvailable As Boolean
    
    GuildIndex = UserList(UserIndex).Guild.GuildIndex
    MaxMembers = GetLimitOfGuildMember(GuildIndex)
    MaxRoles = GetLimitOfGuildRoles(GuildIndex)
    MaxSlotBank = GetLimitOfGuildBankSlots(GuildIndex)
    MaxBoxes = GetLimitOfGuildBankBoxes(GuildIndex)
    MaxContribution = GetLimitOfGuildContribution(GuildIndex)
    BankAvailable = GuildList(GuildIndex).UpgradeEffect.IsGuildBank
    
    Call Writer.WriteInt8(ServerPacketID.GuildInfo)
    
    With GuildList(GuildIndex)
        Call Writer.WriteInt32(.IdGuild)
        Call Writer.WriteString16(.Name)
        Call Writer.WriteString16(.Description)
        Call Writer.WriteInt8(.Alignment)
        Call Writer.WriteString16(.CreationTime)
        Call Writer.WriteInt8(.Status)
        Call Writer.WriteInt32(.IdLeader)
        Call Writer.WriteInt32(.IdRightHand)
        Call Writer.WriteInt8(.MemberCount)
        Call Writer.WriteInt16(.CurrentQuest.IdQuest)
        Call Writer.WriteString16(.CurrentQuest.StartedDate)
        Call Writer.WriteInt32(.ContributionEarned)
        Call Writer.WriteInt32(.ContributionAvailable)
        Call Writer.WriteInt32(.BankGold)
        Call Writer.WriteInt32(UserList(UserIndex).Guild.RoleId)
        Call Writer.WriteInt8(MaxMembers)
        Call Writer.WriteInt8(MaxRoles)
        Call Writer.WriteInt8(MaxSlotBank)
        Call Writer.WriteInt8(MaxBoxes)
        Call Writer.WriteInt32(MaxContribution)
        Call Writer.WriteBool(BankAvailable)
    End With
           
    Call SendData(ToUser, UserIndex, vbNullString)
           
    Exit Sub
ErrHandler:
        Call LogError("Error " & Err.Number & " (" & Err.Description & ") en Sub WriteGuildInfo del Módulo Protocol")
End Sub


Public Sub WriteGuildMembersList(ByVal UserIndex As Integer)
On Error GoTo ErrHandler
    Dim GuildIndex As Integer, I As Integer
    Dim MemberName As String
    
    GuildIndex = UserList(UserIndex).Guild.GuildIndex
    
    Call Writer.WriteInt8(ServerPacketID.GuildMembersList)
    
    With GuildList(GuildIndex)
            Call Writer.WriteInt8(.MemberCount)
            For I = 1 To .MemberCount
                Call Writer.WriteInt32(.Members(I).IdUser)
                Call Writer.WriteInt32(.Members(I).IdRole)
                Call Writer.WriteString16(.Members(I).NameUser)
                Call Writer.WriteBool(IsMemberOnline(GuildIndex, .Members(I).IdUser))
            Next I
    End With
    
    Call SendData(ToUser, UserIndex, vbNullString)

    Exit Sub
ErrHandler:
        Call LogError("Error " & Err.Number & " (" & Err.Description & ") en Sub WriteGuildMembersList del Módulo Protocol")
End Sub


Public Sub WriteGuildRolesList(ByVal UserIndex As Integer)
On Error GoTo ErrHandler
    Dim GuildIndex As Integer
    Dim I As Integer
    Dim J As Integer
    Dim QtyRoles As Integer
    Dim QtyPermissions As Integer
    Dim MemberName As String
    
    GuildIndex = UserList(UserIndex).Guild.GuildIndex
    
    Call Writer.WriteInt8(ServerPacketID.GuildRolesList)
    
    With GuildList(GuildIndex)
        QtyRoles = UBound(.Roles)
        Call Writer.WriteInt16(QtyRoles)
        For I = 1 To QtyRoles
            Call Writer.WriteInt32(.Roles(I).IdRole)
            Call Writer.WriteString16(.Roles(I).RoleName)
            
            Call Writer.WriteBool(.Roles(I).IsDeleteable)
            Call Writer.WriteBool(.Roles(I).IdRole > 2 And .Roles(I).IdRole <> .IdDefaultRole) ' Can update permissions if it's not the leader, righthand or default/recruit role
            Call Writer.WriteBool(.Roles(I).IdRole > 2) ' Can rename if Is Default/Recruit Role or a custom role
            
            Call Writer.WriteInt16(.Roles(I).PermissionCount)
            
            QtyPermissions = UBound(.Roles(I).RolePermission)
            
            For J = 1 To QtyPermissions
                If .Roles(I).RolePermission(J).IsEnabled Then
                    Call Writer.WriteInt32(.Roles(I).RolePermission(J).IdPermission)
                    Call Writer.WriteString16(.Roles(I).RolePermission(J).Key)
                End If
            Next J
        Next I
        
    End With
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
    Exit Sub
ErrHandler:
        Call LogError("Error " & Err.Number & " (" & Err.Description & ") en Sub WriteGuildRolesList del Módulo Protocol")
End Sub


Public Sub WriteGuildUpgradesAcquired(ByVal UserIndex As Integer, Optional ByVal UpgradeId As Integer = 0)
On Error GoTo ErrHandler
    Dim GuildIndex As Integer
    Dim I As Integer
    Dim QtyUpgrade As Integer
    Dim MemberName As String
    
    GuildIndex = UserList(UserIndex).Guild.GuildIndex
    
    Call Writer.WriteInt8(ServerPacketID.GuildUpgradesAcquired)
    
    If Not Not GuildList(GuildIndex).Upgrades Then
        QtyUpgrade = UBound(GuildList(GuildIndex).Upgrades)
    End If
    
    If UpgradeId = 0 Then
    
        Call Writer.WriteInt8(QtyUpgrade)
        
        If QtyUpgrade = 0 Then Exit Sub
        
        With GuildList(GuildIndex)
            
                For I = 1 To QtyUpgrade
                    Call Writer.WriteInt32(.Upgrades(I).IdUpgrade)
                    Call Writer.WriteBool(.Upgrades(I).IsEnabled)
                    Call Writer.WriteString16(.Upgrades(I).UpgradeBy)
                    Call Writer.WriteString16(.Upgrades(I).UpgradeDate)
                    Call Writer.WriteInt8(.Upgrades(I).UpgradeLevel)
                Next I
        End With
        
    Else
        Call Writer.WriteInt8(1)
        
        With GuildList(GuildIndex)
            Call Writer.WriteInt32(.Upgrades(QtyUpgrade).IdUpgrade)
            Call Writer.WriteBool(.Upgrades(QtyUpgrade).IsEnabled)
            Call Writer.WriteString16(.Upgrades(QtyUpgrade).UpgradeBy)
            Call Writer.WriteString16(.Upgrades(QtyUpgrade).UpgradeDate)
            Call Writer.WriteInt8(.Upgrades(QtyUpgrade).UpgradeLevel)

        End With
        
    End If
        
    Call SendData(ToUser, UserIndex, vbNullString)
    
    Exit Sub
ErrHandler:
        Call LogError("Error " & Err.Number & " (" & Err.Description & ") en Sub WriteGuildUpgradesAcquired del Módulo Protocol")
End Sub

Public Sub WriteGuildMemberStatusChange(ByVal UserIndex As Integer, ByVal IdUser As Long, ByVal TypeChanged As Byte, Optional ByVal ValueChangedByte As Byte, Optional ByVal ValueChangedLong As Long)
On Error GoTo ErrHandler

    Call Writer.WriteInt8(ServerPacketID.GuildMemberStatusChange)

    Call Writer.WriteInt32(IdUser)
    Call Writer.WriteInt8(TypeChanged) '1= conexion, 2= rolchange, 3=gold
    Call Writer.WriteInt8(ValueChangedByte)
    Call Writer.WriteInt32(ValueChangedLong)

    Call SendData(ToUser, UserIndex, vbNullString)
    
    Exit Sub
ErrHandler:
        Call LogError("Error " & Err.Number & " (" & Err.Description & ") en Sub WriteGuildMemberStatusChange del Módulo Protocol")
End Sub

Public Sub WriteGuildMemberKicked(ByVal UserIndex As Integer)
On Error GoTo ErrHandler
        
    Call Writer.WriteInt8(ServerPacketID.GuildMemberKicked)
    Call SendData(ToUser, UserIndex, vbNullString)
    
    Exit Sub
ErrHandler:
        Call LogError("Error " & Err.Number & " (" & Err.Description & ") en Sub WriteGuildMemberKicked del Módulo Protocol")
End Sub

Public Sub WriteGuildBankList(ByVal UserIndex As Integer)
On Error GoTo ErrHandler
    Dim GuildIndex As Integer
    Dim I As Integer
    Dim QtyBankSlots As Integer
    Dim CanUse As Boolean
    
    GuildIndex = UserList(UserIndex).Guild.GuildIndex
    
    Call Writer.WriteInt8(ServerPacketID.GuildBankList)
    
    With GuildList(GuildIndex)
        QtyBankSlots = UBound(.Bank)
        Call Writer.WriteInt16(QtyBankSlots)
            For I = 1 To QtyBankSlots
                If .Bank(I).IdObject > 0 Then
                    CanUse = General.checkCanUseItem(UserIndex, .Bank(I).IdObject)
                End If
                Call Writer.WriteInt32(.Bank(I).IdObject)
                Call Writer.WriteInt16(.Bank(I).Box)
                Call Writer.WriteInt16(.Bank(I).Slot)
                Call Writer.WriteInt16(.Bank(I).Amount)
                Call Writer.WriteBool(CanUse)
                
            Next I
    End With
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
    Exit Sub
ErrHandler:
        Call LogError("Error " & Err.Number & " (" & Err.Description & ") en Sub WriteGuildBankList del Módulo Protocol")
End Sub

Public Sub WriteGuildBankChangeSlot(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal ObjIndex As Integer, _
                                                            ByVal Amount As Integer, ByVal Box As Integer, ByVal CanUse As Boolean)

        Call Writer.WriteInt8(ServerPacketID.GuildBankChangeSlot)
        Call Writer.WriteInt8(Slot)
        Call Writer.WriteInt16(Box)
        Call Writer.WriteInt16(ObjIndex)
        Call Writer.WriteInt16(Amount)
        Call Writer.WriteBool(CanUse)
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

Public Sub HandleGuildBankEnd(ByVal UserIndex As Integer)
On Error GoTo ErrHandler
      
    Call SaveGuildBankDB(UserList(UserIndex).Guild.GuildIndex)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleGuildBankEnd de Protocol.bas")
End Sub

Public Sub HandleGuildQuest(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
      
        Dim QuestId As Integer
        Dim GuildIndex As Integer
        Dim I As Integer
      
        QuestId = Reader.ReadInt
        GuildIndex = UserList(UserIndex).Guild.GuildIndex
        
        With GuildList(GuildIndex).CurrentQuest
            If UserList(UserIndex).Guild.RoleId = ID_ROLE_LEADER Then
                If GuildList(GuildIndex).CurrentQuest.IdQuest = 0 Then
                    Call modQuestSystem.StartQuest(UserIndex, QuestId)
                Else
                    Call modQuestSystem.CancelCurrentGuildQuest(GuildIndex, True)
                End If
            End If
        End With
        
    Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleGuildQuest de Protocol.bas")
End Sub

Public Sub HandleGuildQuestAddObject(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
      
        Dim InventorySlot As Integer
        Dim Quantity As Integer
        InventorySlot = Reader.ReadInt
        Quantity = Reader.ReadInt

        Call modQuestSystem.GuildQuestAddObject(UserIndex, InventorySlot, Quantity)
        
    Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleGuildQuestAddObject de Protocol.bas")
End Sub

Public Sub WriteGuildInvitation(ByVal GuildIndex As Long, ByVal InvitedByUserIndex As Long, ByVal TargetUserIndex As Long, ByVal InvitationIndex As Integer, ByVal InvitationLifeTimeInMinutes As Long)
On Error GoTo ErrHandler
    
    Call Writer.WriteInt8(ServerPacketID.GuildSendInvitation)
    Call Writer.WriteString16(UserList(InvitedByUserIndex).Name)
    
    Call Writer.WriteInt(GuildIndex)
    Call Writer.WriteString16(GuildList(GuildIndex).Name)
    
    Call Writer.WriteInt(InvitationIndex)
    Call Writer.WriteInt(InvitationLifeTimeInMinutes)

    Call SendData(ToUser, TargetUserIndex, vbNullString)
    
    Exit Sub
ErrHandler:
        Call LogError("Error " & Err.Number & " (" & Err.Description & ") en Sub WriteGuildInvitation del Módulo Protocol")
End Sub


Public Sub WriteGuildUpgradesList(ByVal UserIndex As Integer)
On Error GoTo ErrHandler
    Dim GuildIndex As Integer
    Dim I As Integer, J As Integer, k As Integer
    Dim QtyUpgrade As Integer, QtyGroup As Integer
    Dim MemberName As String
    Dim UpgReqQty As Integer, QstReqQty As Integer
    Dim Obtained As Boolean
    
    GuildIndex = UserList(UserIndex).Guild.GuildIndex
    
    Call Writer.WriteInt8(ServerPacketID.GuildUpgradesList)
        
    QtyUpgrade = UBound(GuildConfiguration.GuildUpgradesList)
    QtyGroup = GuildConfiguration.UpgradesGroupsQty
    
    With GuildConfiguration
        
        Call Writer.WriteInt16(QtyUpgrade)
        Call Writer.WriteInt16(GuildConfiguration.UpgradesGroupsQty)
        
        For I = 1 To QtyGroup
            Call Writer.WriteInt16(GuildConfiguration.GuildUpgradeGroup(I).UpgradeQty)
            For J = 1 To GuildConfiguration.GuildUpgradeGroup(I).UpgradeQty
                Call Writer.WriteInt16(GuildConfiguration.GuildUpgradeGroup(I).Upgrades(J))
            Next J
        Next I
        
        For I = 1 To QtyUpgrade
            Call Writer.WriteString16(.GuildUpgradesList(I).Name)
            Call Writer.WriteString16(.GuildUpgradesList(I).Description)
            Call Writer.WriteInt32(.GuildUpgradesList(I).IconGraph)
            Call Writer.WriteInt32(.GuildUpgradesList(I).GoldCost)
            Call Writer.WriteInt32(.GuildUpgradesList(I).ContributionCost)
            UpgReqQty = UpgradeRequireSize(.GuildUpgradesList(I).UpgradeRequired)
            Call Writer.WriteInt16(UpgReqQty)
                For J = 1 To UpgReqQty
                Call Writer.WriteInt16(.GuildUpgradesList(I).UpgradeRequired(J))
                Next J
            QstReqQty = GetQuestQtyReq(I)
            Call Writer.WriteInt16(QstReqQty)
                For J = 1 To QstReqQty
                    Call Writer.WriteInt16(.GuildUpgradesList(I).QuestRequired(J).Id)
                    Obtained = False
                    For k = 1 To GuildList(GuildIndex).QuestCompletedCount
                        If GuildList(GuildIndex).QuestCompleted(k).IdQuest = .GuildUpgradesList(I).QuestRequired(J).Id Then
                            Obtained = True
                        End If
                    Next k
                    Call Writer.WriteString16(.GuildUpgradesList(I).QuestRequired(J).Title)
                    Call Writer.WriteBool(Obtained)
                Next J
           
        Next I
    End With
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
    Exit Sub
ErrHandler:
        Call LogError("Error " & Err.Number & " (" & Err.Description & ") en Sub WriteGuildUpgradesList del Módulo Protocol")
End Sub

Public Sub WriteGuildInfoChange(ByVal UserIndex As Integer, ByVal TypeChanged As Byte, Optional ByVal ValueChangedByte As Byte, Optional ByVal ValueChangedLong As Long)
On Error GoTo ErrHandler

    Call Writer.WriteInt8(ServerPacketID.GuildInfoChange)

    Call Writer.WriteInt8(TypeChanged)
    Call Writer.WriteInt8(ValueChangedByte)
    Call Writer.WriteInt32(ValueChangedLong)

    Call SendData(ToUser, UserIndex, vbNullString)
    
    Exit Sub
ErrHandler:
        Call LogError("Error " & Err.Number & " (" & Err.Description & ") en Sub WriteGuildInfoChange del Módulo Protocol")
End Sub

Public Sub HandleWorkerStore(ByVal UserIndex As Integer)
On Error GoTo ErrHandler

    Dim MessageType As Byte
    
    MessageType = Reader.ReadInt8()
    
    Select Case MessageType
        Case eWorkerStoreAction.WorkerStoreGetRecipes
            Call HandleWorkerStore_WorkerStoreGetRecipes(UserIndex)
        
        Case eWorkerStoreAction.WorkerStoreCreate
            Call HandleWorkerStore_Create(UserIndex)
        
        Case eWorkerStoreAction.WorkerStoreClose
            Call HandleWorkerStore_Close(UserIndex)
            
        Case eWorkerStoreAction.WorkerStoreCraftItem
            Call HandleWorkerStore_CraftItem(UserIndex)
            
    End Select
    
        
    Exit Sub
ErrHandler:
    Call LogError("Error " & Err.Number & " (" & Err.Description & ") en Sub WriteGuildMemberStatusChange del Módulo Protocol")
End Sub

Public Sub HandleWorkerStore_Create(ByVal UserIndex As Integer)
    
    Dim CraftingItemsQty As Integer
    Dim CraftingStoreItems() As tCraftingStoreItem

    CraftingItemsQty = Reader.ReadInt16
    If CraftingItemsQty <= 0 Then Exit Sub
    
    ReDim CraftingStoreItems(1 To CraftingItemsQty)
    Dim I As Integer

    With UserList(UserIndex)
    
        For I = 1 To CraftingItemsQty
            CraftingStoreItems(I).Recipe = Reader.ReadInt
            CraftingStoreItems(I).RecipeIndex = Reader.ReadInt
            CraftingStoreItems(I).ConstructionPrice = Reader.ReadInt
            CraftingStoreItems(I).MaterialsPrice = Reader.ReadInt
            CraftingStoreItems(I).RecipeGroup = Reader.ReadInt8
        Next I
    
        If .Pos.Map <> Constantes.CraftingStoreMap Or MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger = eTrigger.ANTIPIQUETE Or MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger = POSINVALIDA Then
            Call WriteConsoleMsg(UserIndex, "No puedes abrir tu tienda en este lugar.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If .CraftingStore.IsOpen Then
            Call WriteConsoleMsg(UserIndex, "Ya tienes una tienda abierta.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        If modCrafting.IsStoreOpenNearby(UserIndex) Then
            Call WriteConsoleMsg(UserIndex, "No puedes abrir tu tienda tan cerca de otro usuario.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
               
    End With
        
    Call modCrafting.CreateWorkerStore(UserIndex, CraftingStoreItems)
    Call WriteWorkerStore_Open(UserIndex)

End Sub

Public Sub HandleWorkerStore_Close(ByVal UserIndex As Integer)
    
    ' Read the list of items
    Call modCrafting.CloseWorkerStore(UserIndex)

End Sub

Public Sub HandleWorkerStore_CraftItem(ByVal UserIndex As Integer)
    
    Call modCrafting.CraftItemOnDemand(UserIndex, UserList(UserIndex).flags.TargetUser, Reader.ReadInt16, Reader.ReadString16)
    
    
End Sub

Public Sub WriteWorkerStore_Open(ByVal UserIndex As Integer)

    Call Writer.WriteInt8(ServerPacketID.WorkerStore)
    Call Writer.WriteInt8(eWorkerStoreServerSubAction.OpenStore)

    Call SendData(ToUser, UserIndex, vbNullString)
        
End Sub

Public Sub WriteWorkerStore_ItemCraftedNotification(ByVal UserIndex As Integer, ByVal BuyerName As String, ByVal ItemNumber As Integer, _
                                                    ByVal ItemQuantity As Double, ByVal ConstructionPrice As Double)
        
    
    Call Writer.WriteInt8(ServerPacketID.WorkerStore)
    Call Writer.WriteInt8(eWorkerStoreServerSubAction.ItemCrafted)
    
    Call Writer.WriteString16(BuyerName)
    Call Writer.WriteInt16(ItemNumber)
    Call Writer.WriteInt16(ItemQuantity)
    Call Writer.WriteInt32(ConstructionPrice)

    Call SendData(ToUser, UserIndex, vbNullString)
        
End Sub


Public Sub WriteWorkerStore_Show(ByVal UserIndex As Integer, ByVal WorkerUserIndex As Integer)
    
    With UserList(WorkerUserIndex)
    
        If UserIndex = WorkerUserIndex Then Exit Sub
    
        Call Writer.WriteInt8(ServerPacketID.WorkerStore)
        Call Writer.WriteInt8(eWorkerStoreServerSubAction.ShowStore)
        
        Call Writer.WriteString16(.Name)
        Call Writer.WriteString16(.CraftingStore.InstanceId)
        Call Writer.WriteInt16(.CraftingStore.ItemsQty)
        
        Dim I As Integer
        For I = 1 To .CraftingStore.ItemsQty
            Call Writer.WriteInt16(.CraftingStore.Items(I).RecipeItem)
            Call Writer.WriteInt16(.CraftingStore.Items(I).RecipeIndex)
            Call Writer.WriteInt16(.CraftingStore.Items(I).ConstructionPrice)
            Call Writer.WriteInt16(.CraftingStore.Items(I).MaterialsPrice)
            Call Writer.WriteInt8(.CraftingStore.Items(I).RecipeGroup)

            ' Write the materials required
            Dim J As Integer
            Dim RequiredMaterialsQty As Integer
            Dim RecipeIndex As Integer
            Dim CraftingGroup As Byte
            Dim ProfessionType As Byte
            
            RecipeIndex = .CraftingStore.Items(I).RecipeIndex
            CraftingGroup = .CraftingStore.Items(I).RecipeGroup
            ProfessionType = .CraftingStore.ProfessionType
            
            RequiredMaterialsQty = Professions(ProfessionType).CraftingRecipeGroups(CraftingGroup).Recipes(RecipeIndex).MaterialsQty
            
            Call Writer.WriteInt16(RequiredMaterialsQty)
            For J = 1 To RequiredMaterialsQty
                Call Writer.WriteInt16(Professions(ProfessionType).CraftingRecipeGroups(CraftingGroup).Recipes(RecipeIndex).Materials(J).ObjIndex)
                Call Writer.WriteInt16(Professions(ProfessionType).CraftingRecipeGroups(CraftingGroup).Recipes(RecipeIndex).Materials(J).Amount)
            Next J

        Next I
    End With
    
    Call SendData(SendTarget.ToUser, UserIndex, vbNullString)
    
End Sub

Public Sub HandleWorkerStore_WorkerStoreGetRecipes(ByVal UserIndex As Integer)
    
    Dim ProfessionType As Byte
    With UserList(UserIndex)

        If .Pos.Map <> Constantes.CraftingStoreMap Or MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger = eTrigger.ANTIPIQUETE Or MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger = POSINVALIDA Then
            Call WriteConsoleMsg(UserIndex, "No puedes abrir tu tienda en este lugar.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    
        If .CraftingStore.IsOpen Then Exit Sub
        
        If MapInfo(.Pos.Map).Pk = True Then
            Call WriteConsoleMsg(UserIndex, "No puedes abrir una tienda en zona insegura.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If .Invent.WeaponEqpObjIndex <= 0 Then
            Call WriteConsoleMsg(UserIndex, "Necesitas equipar una herramienta para poder abrir una tienda.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        ProfessionType = ObjData(.Invent.WeaponEqpObjIndex).ProfessionType
        
        If ProfessionType <= 0 Then
            Call WriteConsoleMsg(UserIndex, "Necesitas equipar una herramienta para poder abrir una tienda.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If modCrafting.IsStoreOpenNearby(UserIndex) Then
            Call WriteConsoleMsg(UserIndex, "No puedes abrir tu tienda tan cerca de otro usuario.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

    End With
        
    If ProfessionType <= 0 Or ProfessionType > UBound(Professions) Then Exit Sub
    
    If Not Professions(ProfessionType).Enabled Then
        Call WriteConsoleMsg(UserIndex, "La profesión asignada a tu herramienta equipada se encuentra desactivada. Intenta denuevo más tarde.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If

    Call WriteCraftableRecipes(UserIndex, ProfessionType)
    Call WriteWorkerStore_OpenFormForCreation(UserIndex)
    
End Sub


Public Sub WriteWorkerStore_OpenFormForCreation(ByVal UserIndex As Integer)

On Error GoTo ErrHandler:

    Call Writer.WriteInt8(ServerPacketID.WorkerStore)
    Call Writer.WriteInt8(eWorkerStoreServerSubAction.OpenFormForCreation)

    Call SendData(SendTarget.ToUser, UserIndex, vbNullString)
    
    Exit Sub
    
ErrHandler:
    Call LogError("Error " & Err.Number & " (" & Err.Description & ") en Sub WriteGuildMemberStatusChange del Módulo Protocol")
End Sub


Public Sub WriteConsoleFormattedMessage(ByVal UserIndex As Integer, ByVal MessageId As Integer, ByRef Parameters() As Variant, ByVal ParametersCount As Integer)
On Error GoTo ErrHandler:

    Dim I As Integer
    
    Call Writer.WriteInt8(ServerPacketID.ConsoleFormattedMessage)
    Call Writer.WriteInt(MessageId)
    Call Writer.WriteInt(ParametersCount)
    
    For I = 0 To ParametersCount - 1
        If VarType(Parameters(I)) = vbString Then
            Call Writer.WriteInt(eMessageParameterType.Text)
            Call Writer.WriteString8(Parameters(I))
        Else
            Call Writer.WriteInt(eMessageParameterType.Number)
            Call Writer.WriteInt(Parameters(I))
        End If
    Next I
    
    Call SendData(SendTarget.ToUser, UserIndex, vbNullString)
    
    Exit Sub

ErrHandler:
    Call LogError("Error " & Err.Number & " (" & Err.Description & ") en Sub WriteConsoleFormattedMessage de Protocol.bas")
End Sub


Public Sub WriteShowFormattedMessageBox(ByVal UserIndex As Integer, ByVal MessageId As Integer, ByRef Parameters() As Variant, ByVal ParametersCount As Integer)
On Error GoTo ErrHandler:

    Dim I As Integer
    
    Call Writer.WriteInt8(ServerPacketID.ShowFormattedMessageBox)
    Call Writer.WriteInt(MessageId)
    Call Writer.WriteInt(ParametersCount)
    
    For I = 0 To ParametersCount - 1
        If VarType(Parameters(I)) = vbString Then
            Call Writer.WriteInt(eMessageParameterType.Text)
            Call Writer.WriteString8(Parameters(I))
        Else
            Call Writer.WriteInt(eMessageParameterType.Number)
            Call Writer.WriteInt(Parameters(I))
        End If
    Next I
    
    Call SendData(SendTarget.ToUser, UserIndex, vbNullString)
    
    Exit Sub

ErrHandler:
    Call LogError("Error " & Err.Number & " (" & Err.Description & ") en Sub WriteShowFormattedMessageBox de Protocol.bas")
End Sub

Public Sub WriteGuildQuestsCompletedList(ByVal UserIndex As Integer)
On Error GoTo ErrHandler
    
    Dim I As Integer
    Dim GuildIndex As Integer
    
    GuildIndex = UserList(UserIndex).Guild.GuildIndex
    
    With GuildList(GuildIndex)
        Call Writer.WriteInt8(ServerPacketID.GuildQuestsCompletedList)
        Call Writer.WriteInt(.QuestCompletedCount)
        
        For I = 1 To .QuestCompletedCount
            Call Writer.WriteInt(.QuestCompleted(I).IdQuest)
        Next I
        
    End With
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
    Exit Sub
ErrHandler:
        Call LogError("Error " & Err.Number & " (" & Err.Description & ") en Sub WriteGuildQuestsCompletedList del Módulo Protocol")
End Sub

Public Sub WriteGuildQuestUpdate_Finished(ByVal UserIndex As Integer, ByVal Failed As Boolean, ByVal QuestId As Integer, ByVal StageNumber As Integer)

    Call Writer.WriteInt8(ServerPacketID.GuildQuestUpdateStatus)
    Call Writer.WriteInt8(eQuestUpdateEvent.EventQuestFinished)
    Call Writer.WriteInt(QuestId)
    Call Writer.WriteInt(StageNumber)
    Call Writer.WriteBool(Failed)
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub


Public Sub WriteGuildCurrentQuestInfo(ByVal UserIndex As Integer)
On Error GoTo ErrHandler
    Dim QuestId As Integer
    Dim CurrentStage As Integer
    Dim I As Long
    Dim TotalSeconds As Long
    Dim GuildIndex As Integer
    
    GuildIndex = UserList(UserIndex).Guild.GuildIndex
    
    If GuildIndex <= 0 Then Exit Sub
    
    QuestId = GuildList(GuildIndex).CurrentQuest.IdQuest
    
    ' If there's no started quest, we don't send anything
    If QuestId = 0 Then Exit Sub
        
    Call Writer.WriteInt8(ServerPacketID.GuildCurrentQuestInfo)

    With GuildList(GuildIndex).CurrentQuest
        Call Writer.WriteInt(.IdQuest)
        
        TotalSeconds = modQuestSystem.GetQuestStageRemainingTime(GuildIndex)
        
        Call Writer.WriteInt(TotalSeconds)
        
        Call Writer.WriteInt(.CurrentStage)
        
        ' The NPC we need to talk with to advance/complete the quest.
        Call Writer.WriteInt(modQuestSystem.GetGuildQuestTalkNpc(GuildIndex))
        
        ' NPC Kills requirements.
        Call Writer.WriteInt(.CurrentNpcKillsQuantity)
        For I = 1 To .CurrentNpcKillsQuantity
            Call Writer.WriteInt(.CurrentNpcKills(I).Quantity)
        Next I

        ' Obj requirements.
        'Excepcion aca, la lista de objetos solo deberia iterarse en el modulo modRequiredObjecList
        Call Writer.WriteInt(.CurrentObjectList.ItemsCount)
        
        For I = 0 To .CurrentObjectList.ItemsCount - 1
            Call Writer.WriteInt(.CurrentObjectList.Items(I).ObjIndex)
            Call Writer.WriteInt(.CurrentObjectList.Items(I).Quantity)
        Next I
            
        ' Army frags requirements.
        Call Writer.WriteInt(.CurrentFrags.Army.Qty)
        
        ' Legion frags requirements.
        Call Writer.WriteInt(.CurrentFrags.Legion.Qty)
        
        ' Neutral frags requirements.
        Call Writer.WriteInt(.CurrentFrags.Neutral.Qty)
        
        Call Writer.WriteBool(.StageIsCompleted)
        
    End With
    
    Call SendData(ToUser, UserIndex, vbNullString)
    
    
Exit Sub
ErrHandler:
End Sub

Public Sub WriteGuildQuestUpdateReqStatus_ObjCollect(ByVal UserIndex As Integer, ByVal QuestId As Integer, ByVal StageNumber As Integer, ByVal ObjectIndex As Integer, ByVal Quantity As Integer)
On Error GoTo ErrHandler

    Call Writer.WriteInt8(ServerPacketID.GuildQuestUpdateStatus)
    Call Writer.WriteInt8(eQuestRequirement.ObjCollect) ' Subpacket: Frags
    Call Writer.WriteInt(QuestId)
    Call Writer.WriteInt(StageNumber)
    Call Writer.WriteInt(ObjectIndex)
    Call Writer.WriteInt(Quantity)
   
    Call SendData(ToUser, UserIndex, vbNullString)
    Exit Sub
    
ErrHandler:
    Call LogError("Error " & Err.Number & " (" & Err.Description & ") en Sub WriteGuildQuestUpdateReqStatus_NPCs del Módulo Protocol")
End Sub

Public Sub WriteGuildQuestUpdateReqStatus_NpcKill(ByVal UserIndex As Integer, ByVal QuestId As Integer, ByVal StageNumber As Integer, ByVal NpcNumber As Integer, ByVal RequirementIndex As Integer, ByVal Amount As Integer)
On Error GoTo ErrHandler

    Call Writer.WriteInt8(ServerPacketID.GuildQuestUpdateStatus)
    Call Writer.WriteInt8(eQuestRequirement.NpcKill) ' Subpacket: Frags
    Call Writer.WriteInt(QuestId)
    Call Writer.WriteInt(StageNumber)
    Call Writer.WriteInt(NpcNumber)
    Call Writer.WriteInt(RequirementIndex)
    Call Writer.WriteInt(Amount)
   
    Call SendData(ToUser, UserIndex, vbNullString)
    Exit Sub
    
ErrHandler:
    Call LogError("Error " & Err.Number & " (" & Err.Description & ") en Sub WriteGuildQuestUpdateReqStatus_NPCs del Módulo Protocol")
End Sub
    

Public Sub WriteGuildQuestUpdateReqStatus_UserKill(ByVal UserIndex As Integer, ByVal QuestId As Integer, ByVal StageNumber As Integer, ByVal NeutralFrags As Integer, ByVal ArmyFrags As Integer, ByVal LegionFrags As Integer)
On Error GoTo ErrHandler

    Call Writer.WriteInt8(ServerPacketID.GuildQuestUpdateStatus)
    Call Writer.WriteInt8(eQuestRequirement.UserKill)
    Call Writer.WriteInt(QuestId)
    Call Writer.WriteInt(StageNumber)
    Call Writer.WriteInt(NeutralFrags)
    Call Writer.WriteInt(ArmyFrags)
    Call Writer.WriteInt(LegionFrags)
   
    Call SendData(ToUser, UserIndex, vbNullString)
    Exit Sub
ErrHandler:
    Call LogError("Error " & Err.Number & " (" & Err.Description & ") en Sub WriteGuildQuestUpdateReqStatus_Frags del Módulo Protocol")
End Sub

Public Sub WriteGuildQuestUpdateStatus(ByVal UserIndex As Integer, ByVal QuestId As Integer, ByVal StageNumber As Integer, ByVal Requirement As Integer, ByVal ExtraInfo As Integer, ByVal Quantity As Integer)
On Error GoTo ErrHandler

    Call Writer.WriteInt8(ServerPacketID.GuildQuestUpdateStatus)
    Call Writer.WriteInt(QuestId)
    Call Writer.WriteInt(StageNumber)
    Call Writer.WriteInt(Requirement)
    Call Writer.WriteInt(ExtraInfo)
    Call Writer.WriteInt(Quantity)
   
    Call SendData(ToUser, UserIndex, vbNullString)
    Exit Sub
ErrHandler:
    Call LogError("Error " & Err.Number & " (" & Err.Description & ") en Sub WriteGuildQuestUpdateStatus del Módulo Protocol")
End Sub

Private Sub HandlePartyInviteMember(ByVal UserIndex As Integer)

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim UserName As String
        Dim tUser As Integer
        Dim rank As Integer
        Dim razon As String
        Dim PI As Integer
        Dim CommandCaster As String
        
    
        UserName = Reader.ReadString8()
        tUser = NameIndex(UserName)
        CommandCaster = UserList(UserIndex).Name
        
        If UserList(UserIndex).Stats.ELV < ConstantesBalance.MinCrearPartyLevel Then
            Call WriteConsoleMsg(UserIndex, "Tu nivel es muy bajo para crear o ingresar a una party.", FontTypeNames.FONTTYPE_PARTY, eMessageType.Party)
            Exit Sub
        End If
        
        PI = UserList(UserIndex).PartyIndex
        rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
        
        If UserList(UserIndex).flags.Muerto Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_PARTY, eMessageType.Party)
            Exit Sub
        End If
        
        If UserName = "" And UserList(UserIndex).flags.TargetUser = 0 Then
            Call WriteConsoleMsg(UserIndex, "Debe seleccionar a alguien para invitar al grupo.", FontTypeNames.FONTTYPE_PARTY, eMessageType.Party)
            Exit Sub
        End If
        
        If tUser = 0 And UserList(UserIndex).flags.TargetUser > 0 Then
            tUser = UserList(UserIndex).flags.TargetUser
        End If
        
        If tUser <= 0 Then
            Call WriteConsoleMsg(UserIndex, UserName & " no se encuentra conectado.", FontTypeNames.FONTTYPE_PARTY, eMessageType.Party)
            Exit Sub
        End If
        
        If (UserList(tUser).flags.Privilegios And rank) > (.flags.Privilegios And rank) Then
            Call WriteConsoleMsg(UserIndex, "No puedes incorporar a tu grupo a personajes de mayor jerarquía.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If UserList(tUser).PartyIndex > 0 Then
            Call WriteConsoleMsg(UserIndex, UserName & " ya posee grupo.", FontTypeNames.FONTTYPE_PARTY, eMessageType.Party)
            Exit Sub
        End If
        
        If PI > 0 Then
            If Not Parties(PI).EsPartyLeader(UserIndex) Then
                Call WriteConsoleMsg(UserIndex, "¡No eres el líder de tu grupo!", FontTypeNames.FONTTYPE_PARTY, eMessageType.Party)
                Exit Sub
            End If
            If Parties(PI).CantMiembros >= Constantes.MaxPartyMembers Then
                Call WriteConsoleMsg(UserIndex, "El grupo ya ha llegado al limite de miembros!", FontTypeNames.FONTTYPE_PARTY, eMessageType.Party)
                Exit Sub
            End If
        End If
        
        If tUser = UserIndex Then
            Call WriteConsoleMsg(UserIndex, "No puedes invitarte a ti mismo a un grupo.", FontTypeNames.FONTTYPE_PARTY, eMessageType.Party)
            Exit Sub
        End If
           
        'send invitation
        Call WritePartyInvitation(tUser, UserIndex, UserList(UserIndex).Name)
        Call WriteConsoleMsg(tUser, "El usuario " & UserList(UserIndex).Name & " te invitó a un grupo. Presiona el boton Grupos para aceptar esta invitación.", FontTypeNames.FONTTYPE_PARTY, eMessageType.Party)
        
    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") por " & CommandCaster & " y target " & UserName & " en Sub HandlePartyInviteMember de Protocol.bas")
End Sub

Public Sub WritePartyInvitation(ByVal UserIndex As Integer, ByVal UserIndexRequest As Integer, ByVal UserNameRequest As String)
    
        Call Writer.WriteInt8(ServerPacketID.PartyInvitation)
        
        Call Writer.WriteString16(UserNameRequest)
        Call Writer.WriteInt8(UserIndexRequest)
       
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub


Public Sub HandlePartyAcceptInvitation(ByVal UserIndex As Integer)

    Dim UserName As String
    Dim tUser As Integer
    Dim Answer As Boolean
    
    UserName = Reader.ReadString16()
    Answer = Reader.ReadBool()
   
    tUser = NameIndex(UserName)
    
    If tUser = 0 Then
        Call WriteConsoleMsg(UserIndex, UserName & " no se encuentra conectado.", FontTypeNames.FONTTYPE_PARTY, eMessageType.Party)
        Exit Sub
    End If
    
    If Not Answer = True Then
        Exit Sub
    End If

    If UserList(tUser).PartyIndex > 0 Then
        'unirse
        Call mdParty.JoinMemberToParty(tUser, UserIndex)
    Else
        'crear
        Call CrearParty(tUser)
        Call mdParty.JoinMemberToParty(tUser, UserIndex)
    End If
    
    Exit Sub
    
End Sub

Public Sub HandleShowGuildMessages(ByVal UserIndex As Integer)

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim Guild As String

        Guild = Reader.ReadString8()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call modGuild_Functions.AdminListenToGuild(UserIndex, Guild)
        End If

    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleShowGuildMessages de Protocol.bas")
End Sub

Public Sub HandleForgive(ByVal UserIndex As Integer)

On Error GoTo ErrHandler
    With UserList(UserIndex)

        Dim UserName As String

        UserName = Reader.ReadString8()

        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call ModFacciones.ForgiveCharacter(UserIndex, Trim(UserName))
        End If
        
    End With

    Exit Sub

ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleForgive de Protocol.bas")
End Sub

Public Sub WriteSetIntervals(ByVal UserIndex As Integer)
    
    Call Writer.WriteInt8(ServerPacketID.SetIntervals)
    
    Call Writer.WriteInt(ServerConfiguration.Intervals.IntervalSpellMacro)          'INT_MACRO_HECHIS
    Call Writer.WriteInt(ServerConfiguration.Intervals.IntervalWorkMacro)                      'INT_MACRO_TRABAJO
    
    Call Writer.WriteInt(ServerConfiguration.Intervals.IntervalAction)               'INT_ACTION
    Call Writer.WriteInt(ServerConfiguration.Intervals.IntervaloUserPuedeAtacar)    'INT_ATTACK
    Call Writer.WriteInt(ServerConfiguration.Intervals.IntervaloFlechasCazadores)   'INT_ARROWS
    
    Call Writer.WriteInt(ServerConfiguration.Intervals.IntervaloUserPuedeCastear)   'INT_CAST_SPELL
    Call Writer.WriteInt(ServerConfiguration.Intervals.IntervaloMagiaGolpe)         'INT_CAST_ATTACK
    Call Writer.WriteInt(ServerConfiguration.Intervals.IntervaloGolpeMagia)         'INT_CAST_ATTACK
    Call Writer.WriteInt(ServerConfiguration.Intervals.IntervaloUserPuedeTrabajar)  'INT_WORK
    Call Writer.WriteInt(ServerConfiguration.Intervals.IntervaloUserPuedeUsarU)     'INT_USEITEMU
    Call Writer.WriteInt(ServerConfiguration.Intervals.IntervaloUserPuedeUsar)      'INT_USEITEMDCK
    Call Writer.WriteInt(ServerConfiguration.Intervals.IntervalRequestPosition)     'INT_SENTRPU
    Call Writer.WriteInt(ServerConfiguration.Intervals.IntervalMeditate)            'INT_MEDITATE
    
       
    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

Public Sub WriteSpellAttackResult(ByVal UserIndex As Integer, ByVal Success As Boolean, ByVal SpellIndex As Integer)
    
    Call Writer.WriteInt8(ServerPacketID.SpellAttackResult)
    
    Call Writer.WriteBool(Success)
    Call Writer.WriteInt(SpellIndex)

    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

Public Sub WriteAttackResult(ByVal UserIndex As Integer, ByVal Success As Boolean)
    
    Call Writer.WriteInt8(ServerPacketID.AttackResult)
    
    Call Writer.WriteBool(Success)

    Call SendData(ToUser, UserIndex, vbNullString)
    
End Sub

