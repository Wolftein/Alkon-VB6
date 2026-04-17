Attribute VB_Name = "Mod_Declaraciones"
'Argentum Online 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'
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
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez


Option Explicit

'Engine (Aurora)
Public Aurora_Audio    As Aurora_Engine.Audio_Service
Public Aurora_Content  As Aurora_Engine.Content_Service
Public Aurora_Graphics As Aurora_Engine.Graphic_Service
Public Aurora_Network  As Aurora_Engine.Network_Service
Public Aurora_Renderer As Aurora_Engine.Graphic_Renderer
Public Aurora_Scene    As Aurora_Engine.Partitioner

'Objetos públicos
Public DialogosClanes As clsGuildDlg
Public Dialogos As clsDialogs
Public Inventario As clsGraphicalInventory
Public InvBanco(1) As clsGraphicalInventory
Public AccBank(1) As clsGraphicalInventory

'Inventarios de comercio con usuario
Public InvComUsu As clsGraphicalInventory  ' Inventario del usuario visible en el comercio
Public InvComNpc As clsGraphicalInventory ' Inventario con los items que ofrece el npc
Public InvOroComUsu(2) As clsGraphicalInventory  ' Inventarios de oro (ambos usuarios)
Public InvOfferComUsu(1) As clsGraphicalInventory  ' Inventarios de ofertas (ambos usuarios)

'Inventarios de herreria
Public Const MAX_LIST_ITEMS As Byte = 7
Public Const MAX_CRAFT_MATERIAL As Byte = 8
Public InvCraftItem(1 To MAX_LIST_ITEMS) As clsGraphicalInventory
Public InvCraftSelItem As clsGraphicalInventory
Public InvCraftMaterial(1 To MAX_CRAFT_MATERIAL) As clsGraphicalInventory

Public CustomKeys As clsCustomKeys
Public CustomMessages As clsCustomMessages

''
'The main timer of the game.
Public MainTimer As clsTimer

'Random client token. This will be generated when the client starts
Public RandomClientToken As String

Public Type tCraftingStoreItemMaterial
    ItemNumber As Integer
    Quantity As Integer
End Type

Public Type tCraftingStoreItem
    SelectedCraftingGroup As Byte
    ConstructionPrice As Long
    MaterialsPrice As Long
    RecipeNumber As Integer
    RecipeIndex As Integer
    ItemNumber As Integer
    MaterialsQty As Integer
    Materials() As tCraftingStoreItemMaterial
End Type

Public Type tCurrentOpenStore
    OwnerUserIndex As Integer
    OwnerName As String
    Type As Byte
    ItemsQty As Integer
    Items() As tCraftingStoreItem
End Type


Public WorkerStoreItemsToSell As tCurrentOpenStore
Public CurrentOpenStore As tCurrentOpenStore

Public CraftingStoreWeapons As tCurrentOpenStore
Public CraftingStoreArmors As tCurrentOpenStore
Public CraftingStoreCarpentry As tCurrentOpenStore

Public Const MAX_NICKNAME_SIZE As Integer = 15

'Error code
Public Const TOO_FAST As Long = 24036
Public Const REFUSED As Long = 24061
Public Const TIME_OUT As Long = 24060

'Sonidos
Public Const SND_CLICK As String = "click.wav"
Public Const SND_PASOS1 As String = "23.wav"
Public Const SND_PASOS2 As String = "24.wav"
Public Const SND_NAVEGANDO As String = "50.wav"

' Head index of the casper. Used to know if a char is killed

' Interval for the default connection actions (account connect, char connect, char creation, etc)
Public Const INT_ACTION As Integer = 1000

Public Const CASPER_HEAD As Integer = 500
Public Const FRAGATA_FANTASMAL As Integer = 87

Public Const NUMATRIBUTES As Byte = 5

Public Const HUMANO_H_PRIMER_CABEZA As Integer = 1
Public Const HUMANO_H_ULTIMA_CABEZA As Integer = 40 'En verdad es hasta la 51, pero como son muchas estas las dejamos no seleccionables
Public Const HUMANO_H_CUERPO_DESNUDO As Integer = 21

Public Const ELFO_H_PRIMER_CABEZA As Integer = 101
Public Const ELFO_H_ULTIMA_CABEZA As Integer = 122
Public Const ELFO_H_CUERPO_DESNUDO As Integer = 210

Public Const DROW_H_PRIMER_CABEZA As Integer = 201
Public Const DROW_H_ULTIMA_CABEZA As Integer = 221
Public Const DROW_H_CUERPO_DESNUDO As Integer = 32

Public Const ENANO_H_PRIMER_CABEZA As Integer = 301
Public Const ENANO_H_ULTIMA_CABEZA As Integer = 319
Public Const ENANO_H_CUERPO_DESNUDO As Integer = 53

Public Const GNOMO_H_PRIMER_CABEZA As Integer = 401
Public Const GNOMO_H_ULTIMA_CABEZA As Integer = 416
Public Const GNOMO_H_CUERPO_DESNUDO As Integer = 222
'**************************************************
Public Const HUMANO_M_PRIMER_CABEZA As Integer = 70
Public Const HUMANO_M_ULTIMA_CABEZA As Integer = 89
Public Const HUMANO_M_CUERPO_DESNUDO As Integer = 39

Public Const ELFO_M_PRIMER_CABEZA As Integer = 170
Public Const ELFO_M_ULTIMA_CABEZA As Integer = 188
Public Const ELFO_M_CUERPO_DESNUDO As Integer = 259

Public Const DROW_M_PRIMER_CABEZA As Integer = 270
Public Const DROW_M_ULTIMA_CABEZA As Integer = 288
Public Const DROW_M_CUERPO_DESNUDO As Integer = 40

Public Const ENANO_M_PRIMER_CABEZA As Integer = 370
Public Const ENANO_M_ULTIMA_CABEZA As Integer = 384
Public Const ENANO_M_CUERPO_DESNUDO As Integer = 60

Public Const GNOMO_M_PRIMER_CABEZA As Integer = 470
Public Const GNOMO_M_ULTIMA_CABEZA As Integer = 484
Public Const GNOMO_M_CUERPO_DESNUDO As Integer = 260

'Musica
Public Const MP3_Inicio As Byte = 1

Public RawServersList As String

Public Type tColor
    r As Byte
    g As Byte
    b As Byte
End Type

Public ColoresPJ(0 To 50) As tColor

'CHOTS | Colores de diálogos customizables
Public Const MAXCOLORESDIALOGOS As Byte = 6

Public ColoresDialogos(1 To MAXCOLORESDIALOGOS) As tColor
'Referencias:
'1=Normal
'2=Clan
'3=Party
'4=Gritar
'5=Palabras Mágicas
'6=Susurrar
'CHOTS

'Usado para definir cuál es el nombre de la UI por defecto que el cliente va a usar.
Public Const SELECTED_UI As String = "Default\"

Public Type tServerInfo
    Ip As String
    Puerto As Integer
    Desc As String
    PassRecPort As Integer
    PanelPassRecoveryUrl As String
End Type
    
Public ServersLst() As tServerInfo
Public ServersRecibidos As Boolean

Public CurServer As Integer

Public CreandoClan As Boolean
Public ClanName As String
Public secClanName As String
Public Site As String

Public UserCiego As Boolean
Public UserEstupido As Boolean
Public UserCompletedQuests() As Integer

Public NoRes As Boolean 'no cambiar la resolucion

Public RainBufferIndex As Long

Public Const bCabeza = 1
Public Const bPiernaIzquierda = 2
Public Const bPiernaDerecha = 3
Public Const bBrazoDerecho = 4
Public Const bBrazoIzquierdo = 5
Public Const bTorso = 6

Public NumEscudosAnims As Integer


' Crafting

Public Type tCraftingRecipeMaterial
    ObjNumber As Integer
    Amount As Integer
End Type

Public Type tCrafgintRecipe
    RecipeIndex As Integer
    ObjNumber As Integer
    
    CraftingProbability As Byte
    BlacksmithSkillNeeded As Byte
    CarpenterSkillNeeded As Byte
    
    MaterialsPrice As Long
    ConstructionPrice As Long
    
    SelectedCraftingGroup As Byte
    
    MaterialsQty As Integer
    Materials() As tCraftingRecipeMaterial
End Type

Public Type tCraftingRecipeGroup
    TabTitle As String
    TabImage As String
    OwnerName As String
    ProfessionType As Byte
   
    RecipesQty As Integer
    Recipes() As tCrafgintRecipe
End Type

Public ArmasHerrero() As tItemsConstruibles
Public ArmadurasHerrero() As tItemsConstruibles
Public ObjCarpintero() As tItemsConstruibles
Public UpgradeHerrero() As tItemsConstruibles
Public UpgradeCarpintero() As tItemsConstruibles



Public UsaMacro As Boolean
Public CnTd As Byte


Public Const MAX_BANCOINVENTORY_SLOTS As Byte = 40
Public UserBancoInventory() As Inventory

Public AccBankInventory() As Inventory

Public Const MAX_QUESTINVENTORY_SLOTS As Byte = 20
Public UserQuestInventory(1 To MAX_QUESTINVENTORY_SLOTS) As Inventory

Public TradingUserName As String

Public Enum eCharacterAlignment
    Newbie = 0
    Neutral = 1
    FactionRoyal = 2
    FactionLegion = 3
End Enum

'Direcciones
Public Enum E_Heading
    NORTH = 1
    EAST = 2
    SOUTH = 3
    WEST = 4
End Enum

'Objetos
Public Const MAX_INVENTORY_OBJS As Integer = 10000
Public Const MAX_INVENTORY_SLOTS As Byte = 30
Public Const MAX_NPC_INVENTORY_SLOTS As Byte = 20
Public Const MAXHECHI As Byte = 35

Public Const INV_OFFER_SLOTS As Byte = 25
Public Const INV_GOLD_SLOTS As Byte = 1

Public Const MAXSKILLPOINTS As Byte = 100

Public Const MAXATRIBUTOS As Byte = 40

Public Const FLAGORO As Integer = MAX_INVENTORY_SLOTS + 1
Public Const GOLD_OFFER_SLOT As Integer = INV_OFFER_SLOTS + 1

Public Enum eClass
    Mage = 1    'Mago
    Cleric      'Clérigo
    Warrior     'Guerrero
    Assasin     'Asesino
    Thief       'Ladrón
    Bard        'Bardo
    Druid       'Druida
    Bandit      'Bandido
    Paladin     'Paladín
    Worker      'Trabajador
    Hunter      'Cazador
    
End Enum

Public Enum eCiudad
    cUllathorpe = 1
    cNix
    cBanderbill
    cLindos
    cArghal
End Enum

Enum eRaza
    Humano = 1
    Elfo
    ElfoOscuro
    Gnomo
    Enano
End Enum

Public Enum eSkill
    Magia = 1
    Robar = 2
    Tacticas = 3
    Armas = 4
    Meditar = 5
    Apuñalar = 6
    Ocultarse = 7
    Supervivencia = 8
    Talar = 9
    Defensa = 10
    Pesca = 11
    Mineria = 12
    Carpinteria = 13
    Herreria = 14
    Domar = 15
    Proyectiles = 16
    Wrestling = 17
    Sastreria = 18
End Enum

Public Enum eAtributos
    Fuerza = 1
    Agilidad = 2
    Inteligencia = 3
    Carisma = 4
    Constitucion = 5
End Enum

Enum eGenero
    Hombre = 1
    Mujer
End Enum

Public Enum PlayerType
    User = &H1
    Consejero = &H2
    SemiDios = &H4
    Dios = &H8
    Admin = &H10
    RoleMaster = &H20
    ChaosCouncil = &H40
    RoyalCouncil = &H80
End Enum

Public Enum eObjType
    otUseOnce = 1
    otWeapon = 2
    otArmadura = 3
    otArboles = 4
    otGuita = 5
    otPuertas = 6
    otContenedores = 7
    otCarteles = 8
    otLlaves = 9
    otForos = 10
    otPociones = 11
    otBebidas = 13
    otLeña = 14
    otFogata = 15
    otEscudo = 16
    otCasco = 17
    otAnillo = 18
    otTeleport = 19
    otYacimiento = 22
    otMinerales = 23
    otPergaminos = 24
    otInstrumentos = 26
    otYunque = 27
    otFragua = 28
    otBarcos = 31
    otFlechas = 32
    otBotellaVacia = 33
    otBotellaLlena = 34
    otManchas = 35          'No se usa
    otArbolElfico = 36
    otMochilas = 37
    otTool = 46
    otCualquiera = 1000
End Enum

Public MaxInventorySlots As Byte

Public Const FundirMetal As Integer = 88

' Determina el color del nick
Public Enum eNickColor
    ieNeutral = &H1
    ieCriminal = &H2
    ieCiudadano = &H3
    ieAtacable = &H4
End Enum

Public Enum eGMCommands
    GMMessage = 1           '/GMSG
    showName                '/SHOWNAME
    OnlineRoyalArmy         '/ONLINEREAL
    OnlineChaosLegion       '/ONLINECAOS
    GoNearby                '/IRCERCA
    Comment                 '/REM
    serverTime              '/HORA
    Where                   '/DONDE
    CreaturesInMap          '/NENE
    WarpMeToTarget          '/TELEPLOC
    WarpChar                '/TELEP
    Silence                 '/SILENCIAR
    SOSShowList             '/SHOW SOS
    SOSRemove               'SOSDONE
    GoToChar                '/IRA
    invisible               '/INVISIBLE
    GMPanel                 '/PANELGM
    RequestUserList         'LISTUSU
    Jail                    '/CARCEL
    KillNPC                 '/RMATA
    WarnUser                '/ADVERTENCIA
    EditChar                '/MOD
    RequestStatsBosses
    RequestCharInfo         '/INFO
    RequestCharStats        '/STAT
    RequestCharGold         '/BAL
    RequestCharInventory    '/INV
    RequestCharBank         '/BOV
    RequestCharSkills       '/SKILLS
    Forgive                 '/PERDON
    ReviveChar              '/REVIVIR
    OnlineGM                '/ONLINEGM
    OnlineMap               '/ONLINEMAP
    Kick                    '/ECHAR
    Execute                 '/EJECUTAR
    banChar                 '/BAN
    UnbanChar               '/UNBAN
    NPCFollow               '/SEGUIR
    SummonChar              '/SUM
    SpawnListRequest        '/CC
    SpawnCreature           'SPA
    ResetNPCInventory       '/RESETINV
    CleanWorld              '/LIMPIAR
    ServerMessage           '/RMSG
    nickToIP                '/NICK2IP
    IPToNick                '/IP2NICK
    GuildOnlineMembers      '/ONCLAN
    TeleportCreate          '/CT
    TeleportDestroy         '/DT
    RainToggle              '/LLUVIA
    SetCharDescription      '/SETDESC
    ForceMIDIToMap          '/FORCEMIDIMAP
    ForceWAVEToMap          '/FORCEWAVMAP
    RoyalArmyMessage        '/REALMSG
    ChaosLegionMessage      '/CAOSMSG
    TalkAsNPC               '/TALKAS
    DestroyAllItemsInArea   '/MASSDEST
    AcceptRoyalCouncilMember '/ACEPTCONSE
    AcceptChaosCouncilMember '/ACEPTCONSECAOS
    ItemsInTheFloor         '/PISO
    MakeDumb                '/ESTUPIDO
    MakeDumbNoMore          '/NOESTUPIDO
    dumpIPTables            '/DUMPSECURITY
    CouncilKick             '/KICKCONSE
    SetTrigger              '/TRIGGER
    AskTrigger              '/TRIGGER with no args
    BannedIPList            '/BANIPLIST
    BannedIPReload          '/BANIPRELOAD
    GuildMemberList         '/MIEMBROSCLAN
    GuildBan                '/BANCLAN
    BanIP                   '/BANIP
    UnbanIP                 '/UNBANIP
    CreateItem              '/CI
    DestroyItems            '/DEST
    FactionKick             '/ECHARFACCION
    ForceMIDIAll            '/FORCEMIDI
    ForceWAVEAll            '/FORCEWAV
    RemovePunishment        '/BORRARPENA
    TileBlockedToggle       '/BLOQ
    KillNPCNoRespawn        '/MATA
    KillAllNearbyNPCs       '/MASSKILL
    LastIP                  '/LASTIP
    ChangeMOTD              '/MOTDCAMBIA
    SetMOTD                 'ZMOTD
    SystemMessage           '/SMSG
    CreateNPC               '/ACC
    CreateNPCWithRespawn    '/RACC
    ImperialArmour          '/AI1 - 4
    ChaosArmour             '/AC1 - 4
    NavigateToggle          '/NAVE
    ServerOpenToUsersToggle '/HABILITAR
    ResetFactions           '/RAJAR
    RemoveCharFromGuild     '/ECHARCLAN
    ModGuildContribution    '/MODCLANCONTRI
    RequestCharMail         '/LASTEMAIL
    AlterPassword           '/APASS
    AlterMail               '/AEMAIL
    AlterName               '/ANAME
    DoBackUp                '/DOBACKUP
    ShowGuildMessages       '/SHOWCMSG
    SaveMap                 '/GUARDAMAPA
    ChangeMapInfoPK         '/MODMAPINFO PK
    ChangeMapInfoBackup     '/MODMAPINFO BACKUP
    ChangeMapInfoRestricted '/MODMAPINFO RESTRINGIR
    ChangeMapInfoNoMagic    '/MODMAPINFO MAGIASINEFECTO
    ChangeMapInfoNoInvi     '/MODMAPINFO INVISINEFECTO
    ChangeMapInfoNoResu     '/MODMAPINFO RESUSINEFECTO
    ChangeMapInfoLand       '/MODMAPINFO TERRENO
    ChangeMapInfoZone       '/MODMAPINFO ZONA
    ChangeMapInfoStealNpc   '/MODMAPINFO ROBONPCm
    ChangeMapInfoNoOcultar  '/MODMAPINFO OCULTARSINEFECTO
    ChangeMapInfoNoInvocar  '/MODMAPINFO INVOCARSINEFECTO
    SaveChars               '/GRABAR
    CleanSOS                '/BORRAR SOS
    ShowServerForm          '/SHOW INT
    KickAllChars            '/ECHARTODOSPJS
    ReloadNPCs              '/RELOADNPCS
    ReloadServerIni         '/RELOADSINI
    ReloadSpells            '/RELOADHECHIZOS
    ReloadObjects           '/RELOADOBJ
    Restart                 '/REINICIAR
    ResetAutoUpdate         '/AUTOUPDATE
    ChatColor               '/CHATCOLOR
    Ignored                 '/IGNORADO
    CheckSlot               '/SLOT
    SetIniVar               '/SETINIVAR LLAVE CLAVE VALOR
    CreatePretorianClan     '/CREARPRETORIANOS
    RemovePretorianClan     '/ELIMINARPRETORIANOS
    EnableDenounces         '/DENUNCIAS
    ShowDenouncesList       '/SHOW DENUNCIAS
    MapMessage              '/MAPMSG
    SetDialog               '/SETDIALOG
    Impersonate             '/IMPERSONAR
    Imitate                 '/MIMETIZAR
    RecordAdd
    RecordRemove
    RecordAddObs
    RecordListRequest
    RecordDetailsRequest
    PMSend
    PMDeleteUser
    PMListUser
    RequestTournamentCompetitors
    Descalificar
    Pelea
    CerrarTorneo
    IniciarTorneo
    TorunamentEdit
    RequestTournamentConfig
    AlterGuildName
    HigherAdminsMessage
    GetPunishmentTypeList
    AdminChangeGuildAlign
    ChangeMapInfoNoInmo
    ChangeMapInfoMismoBando
    SpawnBoss
End Enum

'
' Mensajes
'
' MENSAJE_*  --> Mensajes de texto que se muestran en el cuadro de texto
'

Public Const MENSAJE_CRIATURA_FALLA_GOLPE As String = "¡¡¡La criatura falló el golpe!!!"
Public Const MENSAJE_CRIATURA_MATADO As String = "¡¡¡La criatura te ha matado!!!"
Public Const MENSAJE_RECHAZO_ATAQUE_ESCUDO As String = "¡¡¡Has rechazado el ataque con el escudo!!!"
Public Const MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO  As String = "¡¡¡El usuario rechazó el ataque con su escudo!!!"
Public Const MENSAJE_FALLADO_GOLPE As String = "¡¡¡Has fallado el golpe!!!"
Public Const MENSAJE_SEGURO_ACTIVADO As String = ">>SEGURO ACTIVADO<<"
Public Const MENSAJE_SEGURO_DESACTIVADO As String = ">>SEGURO DESACTIVADO<<"
Public Const MENSAJE_SEGURO_ADVIERTE As String = "Recuerda que si matas una criatura o ayudas a otros ciudadanos, puedes volverte uno de ellos."
Public Const MENSAJE_PIERDE_NOBLEZA As String = "¡¡Has perdido puntaje de nobleza y ganado puntaje de criminalidad!! Si sigues ayudando a criminales te convertirás en uno de ellos y serás perseguido por las tropas de las ciudades."
Public Const MENSAJE_USAR_MEDITANDO As String = "¡Estás meditando! Debes dejar de meditar para usar objetos."

Public Const MENSAJE_SEGURO_RESU_ON As String = "SEGURO DE RESURRECCION ACTIVADO"
Public Const MENSAJE_SEGURO_RESU_OFF As String = "SEGURO DE RESURRECCION DESACTIVADO"

Public Const MENSAJE_GOLPE_CABEZA As String = "¡¡La criatura te ha pegado en la cabeza por "
Public Const MENSAJE_GOLPE_BRAZO_IZQ As String = "¡¡La criatura te ha pegado el brazo izquierdo por "
Public Const MENSAJE_GOLPE_BRAZO_DER As String = "¡¡La criatura te ha pegado el brazo derecho por "
Public Const MENSAJE_GOLPE_PIERNA_IZQ As String = "¡¡La criatura te ha pegado la pierna izquierda por "
Public Const MENSAJE_GOLPE_PIERNA_DER As String = "¡¡La criatura te ha pegado la pierna derecha por "
Public Const MENSAJE_GOLPE_TORSO  As String = "¡¡La criatura te ha pegado en el torso por "

' MENSAJE_[12]: Aparecen antes y despues del valor de los mensajes anteriores (MENSAJE_GOLPE_*)
Public Const MENSAJE_1 As String = "¡¡"
Public Const MENSAJE_2 As String = "!!"
Public Const MENSAJE_22 As String = "!"

Public Const MENSAJE_GOLPE_CRIATURA_1 As String = "¡¡Le has pegado a la criatura por "

Public Const MENSAJE_ATAQUE_FALLO As String = " te atacó y falló!!"

Public Const MENSAJE_RECIVE_IMPACTO_CABEZA As String = " te ha pegado en la cabeza por "
Public Const MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ As String = " te ha pegado el brazo izquierdo por "
Public Const MENSAJE_RECIVE_IMPACTO_BRAZO_DER As String = " te ha pegado el brazo derecho por "
Public Const MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ As String = " te ha pegado la pierna izquierda por "
Public Const MENSAJE_RECIVE_IMPACTO_PIERNA_DER As String = " te ha pegado la pierna derecha por "
Public Const MENSAJE_RECIVE_IMPACTO_TORSO As String = " te ha pegado en el torso por "

Public Const MENSAJE_PRODUCE_IMPACTO_1 As String = "¡¡Le has pegado a "
Public Const MENSAJE_PRODUCE_IMPACTO_CABEZA As String = " en la cabeza por "
Public Const MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ As String = " en el brazo izquierdo por "
Public Const MENSAJE_PRODUCE_IMPACTO_BRAZO_DER As String = " en el brazo derecho por "
Public Const MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ As String = " en la pierna izquierda por "
Public Const MENSAJE_PRODUCE_IMPACTO_PIERNA_DER As String = " en la pierna derecha por "
Public Const MENSAJE_PRODUCE_IMPACTO_TORSO As String = " en el torso por "

Public Const MENSAJE_TRABAJO_MAGIA As String = "Haz click sobre el objetivo..."
Public Const MENSAJE_TRABAJO_PESCA As String = "Haz click sobre el sitio donde quieres pescar..."
Public Const MENSAJE_TRABAJO_ROBAR As String = "Haz click sobre la víctima..."
Public Const MENSAJE_TRABAJO_TALAR As String = "Haz click sobre el árbol..."
Public Const MENSAJE_TRABAJO_MINERIA As String = "Haz click sobre el yacimiento..."
Public Const MENSAJE_TRABAJO_FUNDIRMETAL As String = "Haz click sobre la fragua..."
Public Const MENSAJE_TRABAJO_PROYECTILES As String = "Haz click sobre la victima..."

Public Const MENSAJE_NO_VES_NADA_INTERESANTE As String = "No ves nada interesante."
Public Const MENSAJE_HAS_MATADO_A As String = "Has matado a "
Public Const MENSAJE_TE_HA_MATADO As String = " te ha matado!"
Public Const MENSAJE_ERA_NIVEL As String = " era nivel "

Public Const MENSAJE_HOGAR As String = "Has llegado a tu hogar. El viaje ha finalizado."
Public Const MENSAJE_HOGAR_CANCEL As String = "Tu viaje ha sido cancelado."

Public Enum eMessages
    DontSeeAnything
    NPCSwing
    NPCKillUser
    BlockedWithShieldUser
    BlockedWithShieldOther
    UserSwing
    SafeModeOn
    SafeModeOff
    ResuscitationSafeOff
    ResuscitationSafeOn
    NobilityLost
    CantUseWhileMeditating
    NPCHitUser
    UserHitNPC
    UserAttackedSwing
    UserHittedByUser
    UserHittedUser
    WorkRequestTarget
    SpellCastRequestTarget
    HaveKilledUser
    UserKill
    EarnExp
    GoHome
    CancelGoHome
    FinishHome
End Enum

'Inventario
Type Inventory
    ObjIndex As Integer
    Name As String
    GrhIndex As Integer
    '[Alejo]: tipo de datos ahora es Long
    Amount As Long
    '[/Alejo]
    Equipped As Byte
    Valor As Single
    OBJType As Integer
    MaxDef As Integer
    MinDef As Integer 'Budi
    MaxHit As Integer
    MinHit As Integer
    MaxAmount As Long
    CanUse As Boolean
End Type

Type NpCinV
    ObjIndex As Integer
    Name As String
    GrhIndex As Integer
    Amount As Integer
    Valor As Single
    OBJType As Integer
    MaxDef As Integer
    MinDef As Integer
    MaxHit As Integer
    MinHit As Integer
    CanUse As Boolean
    C1 As String
    C2 As String
    C3 As String
    C4 As String
    C5 As String
    C6 As String
    C7 As String
End Type

Type tEstadisticasUsu
    CiudadanosMatados As Long
    CriminalesMatados As Long
    UsuariosMatados As Long
    NpcsMatados As Long
    Clase As String
    PenaCarcel As Long
End Type

Type tCraftItem
    ObjIndex As Integer
    Amount As Integer
End Type

Type tItemsConstruibles
    Name As String
    ObjIndex As Integer
    StationRecipeIndex As Integer
    GrhIndex As Integer
    CraftItem() As tCraftItem
End Type

Public Enum eNombresView
    off = 0
    Rollover = 1
    Fixed = 2
End Enum

Public Nombres As eNombresView

'User status vars

Public UserHechizos(1 To MAXHECHI) As Integer

Public NPCInventory() As NpCinV

Public UserMeditar As Boolean
Public UserName As String
Public uName As String 'nick con caracteres especiales
Public UserPassword As String
Public UserMaxHP As Integer
Public UserMinHP As Integer
Public UserMaxMAN As Integer
Public UserMinMAN As Integer
Public UserMaxSTA As Integer
Public UserMinSTA As Integer
Public UserMaxAGU As Byte
Public UserMinAGU As Byte
Public UserMaxHAM As Byte
Public UserMinHAM As Byte
Public UserGLD As Double
Public UserLvl As Integer
Public UserPort As Integer
Public UserServerIP As String
Public UserEstado As Byte '0 = Vivo & 1 = Muerto
Public UserPasarNivel As Long
Public UserExp As Long
Public IsMaxLevel As Boolean

Public UserMasteryPoints As Integer
Public UserEstadisticas As tEstadisticasUsu
Public UserDescansar As Boolean

Public PrimeraVez As Boolean
Public bShowTutorial As Boolean

Public pausa As Boolean
Public UserParalizado As Boolean
Public UserNavegando As Boolean
Public UserHogar As eCiudad

Public UserFuerza As Byte
Public UserAgilidad As Byte

Public UserWeaponEqpSlot As Byte
Public UserArmourEqpSlot As Byte
Public UserSecArmourEqpSlot As Byte  ' Secondary armour.
Public UserHelmEqpSlot As Byte
Public UserShieldEqpSlot As Byte

'<-------------------------NUEVO-------------------------->
Public Comerciando As Boolean
Public ViewingFormCantMove As Boolean
Public MirandoForo As Boolean
Public MirandoAsignarSkills As Boolean
Public MirandoEstadisticas As Boolean
Public MirandoParty As Boolean
Public MirandoCarpinteria As Boolean
Public MirandoHerreria As Boolean
Public MirandoPanelQuest As Boolean
'<-------------------------NUEVO-------------------------->

'<-------------------------PET SYSTEM-------------------------->
Public HasPets As Boolean
Public PetSelectedIndex As Byte
Public PetList() As Integer
Public PetListQty As Integer
'<-------------------------PET SYSTEM-------------------------->



Public UserEmail As String

Public Const NUMCIUDADES As Byte = 5
Public Const NUMSKILLS As Byte = 18
Public Const NUMATRIBUTOS As Byte = 5
Public Const NumClases As Byte = 11
Public Const EnabledClassesQty As Byte = 9
Public Const NUMRAZAS As Byte = 5

Public UserSkills(1 To NUMSKILLS) As clsSkill
Public PorcentajeSkills(1 To NUMSKILLS) As Byte
Public SkillsNames(1 To NUMSKILLS) As String

Public UserAtributos(1 To NUMATRIBUTOS) As Byte
Public AtributosNames(1 To NUMATRIBUTOS) As String

Public Ciudades(1 To NUMCIUDADES) As String

Public ListaRazas(1 To NUMRAZAS) As String
Public ListaClases(1 To NumClases) As String
Public ListEnabledClasses(1 To EnabledClassesQty) As Byte

Public SkillPoints As Integer
Public Alocados As Integer
Public Flags() As Integer
Public Logged As Boolean

Public UsingSkill As Integer
Public CastedSpellNumber As Integer
Public CastedSpellIndex As Integer

Public MD5HushYo As String * 16

Public currentPingTime As Long
Public HeartBeatTime As Integer
Public ShowPerformanceData As Boolean

Public EsPartyLeader As Boolean

Public Type tPartyTempInvitacion
    UserNameRequest     As String
    UserIndexRequest    As Integer
End Type

Public PartyTempInvitation As tPartyTempInvitacion

Public Enum FxMeditar
    CHICO = 4
    MEDIANO = 5
    GRANDE = 6
    XGRANDE = 16
    XXGRANDE = 34
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

''
' TRIGGERS
'
' @param NADA nada
' @param BAJOTECHO bajo techo
' @param trigger_2 ???
' @param POSINVALIDA los npcs no pueden pisar tiles con este trigger
' @param ZONASEGURA no se puede robar o pelear desde este trigger
' @param ANTIPIQUETE
' @param ZONAPELEA al pelear en este trigger no se caen las cosas y no cambia el estado de ciuda o crimi
' @param ZONAOSCURA lo que haya en este trigger no será visible
' @param CASA todo lo que tenga este trigger forma parte de una casa
'
Public Enum eTrigger
    NADA = 0
    BAJOTECHO = 1
    trigger_2 = 2
    POSINVALIDA = 3
    ZONASEGURA = 4
    ANTIPIQUETE = 5
    ZONAPELEA = 6
    ZONAOSCURA = 7
    CASA = 8
End Enum

'Server stuff
Public stxtbuffer As String 'Holds temp raw data from server
Public stxtbuffercmsg As String 'Holds temp raw data from server
Public Connected As Boolean 'True when connected to server
Public UserMap As Integer

'Control
Public prgRun As Boolean 'When true the program ends

Public IPdelServidor As String
Public PuertoDelServidor As String

'
'********** FUNCIONES API ***********
'
Public Declare Function mciSendString Lib "winmm" Alias "mciSendStringA" (ByVal lpstrCommand As String, _
                                                                            ByVal lpstrReturnString As String, _
                                                                            ByVal uReturnLength As Long, _
                                                                            ByVal hwndCallback As Long) As Long
                                                                    

Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function timeGetTime Lib "winmm.dll" () As Long

'para escribir y leer variables
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpFileName As String) As Long

'Teclado
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Para ejecutar el browser y programas externos
Public Const SW_SHOWNORMAL As Long = 1
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Cambio de resolución

Public Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long


' New cursors
Public Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long


'Lista de cabezas
Public Type tIndiceCabeza
    Head(1 To 4) As Integer
End Type

Public Type tIndiceCuerpo
    Body(1 To 4) As Integer
    HeadOffsetX As Integer
    HeadOffsetY As Integer
End Type

Public Type tIndiceFx
    Animacion As Integer
    OffsetX As Integer
    OffsetY As Integer
End Type

Public EsperandoLevel As Boolean

' Tipos de mensajes
Public Enum eForumMsgType
    ieGeneral
    ieGENERAL_STICKY
    ieREAL
    ieREAL_STICKY
    ieCAOS
    ieCAOS_STICKY
End Enum

' Indica los privilegios para visualizar los diferentes foros
Public Enum eForumVisibility
    ieGENERAL_MEMBER = &H1
    ieREAL_MEMBER = &H2
    ieCAOS_MEMBER = &H4
End Enum

' Indica el tipo de foro
Public Enum eForumType
    ieGeneral
    ieREAL
    ieCAOS
End Enum

' Limite de posts
Public Const MAX_STICKY_POST As Byte = 5
Public Const MAX_GENERAL_POST As Byte = 30
Public Const STICKY_FORUM_OFFSET As Byte = 50

' Estructura contenedora de mensajes
Public Type tForo
    StickyTitle(1 To MAX_STICKY_POST) As String
    StickyPost(1 To MAX_STICKY_POST) As String
    StickyAuthor(1 To MAX_STICKY_POST) As String
    GeneralTitle(1 To MAX_GENERAL_POST) As String
    GeneralPost(1 To MAX_GENERAL_POST) As String
    GeneralAuthor(1 To MAX_GENERAL_POST) As String
End Type

' 1 foro general y 2 faccionarios
Public Foros(0 To 2) As tForo

' Forum info handler
Public clsForos As clsForum

Public Traveling As Boolean

Public Const OFFSET_HEAD As Integer = -34

Public Enum eSMType
    sResucitation
    sSafemode
    mSpells
    mWork
    mPets
End Enum

Public Const SM_CANT As Byte = 4
Public SMStatus(SM_CANT) As Boolean

'Hardcoded grhs and items

Public Const ORO_INDEX As Integer = 12
Public Const ORO_GRH As Integer = 511

Public picMouseIcon As Picture

Public Enum eMoveType
    Inventory = 1
    Target
    InventoryToTarget
    TargetToInventory
    None
End Enum

'Caracteres
Public Const CAR_ESPECIALES = "áàäâÁÀÄÂéèëêÉÈËÊíìïîÍÌÏÎóòöôÓÒÖÔúùüûÚÙÜÛñÑ'"
Public Const CAR_COMUNES = "aaaaAAAAeeeeEEEEiiiiIIIIooooOOOOuuuuUUUUnN "
Public Const CAR_ESPECIALES_CLANES = ".;,'"
Public Const CAR_COMUNES_CLANES = "    "

Public Const DAT_PATH As String = "\DAT\"

'Modificador de defensa para armaduras de segunda jerarquía.
Public Const MOD_DEF_SEG_JERARQUIA As Single = 1.25

Public Type tCabecera 'Cabecera de los con
    Desc As String * 255
    CRC As Long
    MagicWord As Long
End Type
Public MiCabecera As tCabecera

Public Enum eOrigenSkills
    ieAsignacion = 1
    ieEstadisticas = 2
End Enum

Public OrigenSkills As eOrigenSkills

Public Enum E_MODO
    Normal = 1
    Dados = 2
    AccountCreate = 3
    AccountLogin = 4
    AccountDeleteChar = 5
    AccountLoginChar = 6
    AccountCreateChar = 7
    AccountRecover = 8
    AccountChangePassword = 9
End Enum

Public Type tMasteryGroup
    GroupId As Integer
    MasteriesQty As Integer
    Masteries() As Integer
End Type

Public Type tCurrentPlayerMap
    Number As Integer
    version As Integer
    CraftingStoreAllowed As Boolean
End Type

Public Type tPlayerIntervals
    SpellCastMacro As Integer                'INT_MACRO_HECHIS As Integer = 2788 = ???? Crear nuevo
    WorkMacro As Integer                'INT_MACRO_TRABAJO As Integer = 1010 = ???? Crear nuevo
    Work As Integer                     'INT_WORK As Integer = 700 = IntervaloTrabajo
    Actions As Integer                  'INT_ACTION As Integer = 1000 = ???? Crear nuevo.
    PlayerAttack As Integer             'INT_ATTACK As Integer = 1500 = IntervaloUserPuedeAtacar
    PlayerAttackArrow As Integer        'INT_ARROWS As Integer = 1400 = IntervaloFlechasCazadores
    PlayerCastSpell As Integer          'INT_CAST_SPELL As Integer = 1400 = IntervaloLanzaHechizo
    PlayerAttackAfterSpell As Integer   'INT_CAST_ATTACK As Integer = 1000 = IntervaloMagiaGolpe
    PlayerCastSpellAfterAttack As Integer   'IntervaloGolpeMagia
    UseItemWithKey As Integer           'INT_USEITEMU As Integer = 450 = ???? Crear nuevo
    UseItemDoubleClick As Integer       'INT_USEITEMDCK As Integer = 125 = IntervaloUserPuedeUsar
    RequestPositionUpdate As Integer    'INT_SENTRPU As Integer = 2000 = ???? Crear nuevo
    Meditate As Integer                 'INT_MEDITATE As Integer = 750 = ???? Crear nuevo
End Type

' Type to hold all the player's current data.
Public Type tPlayerData
    OnDemandCraftingStoreOpen As Boolean
    ClassMasteryGroupsQty As Integer
    ClassMasteryGroups() As tMasteryGroup
    
    MasteryGroupsQty As Integer
    MasteryGroups() As tMasteryGroup
    
    Guild As tGuildInfo
    
    CurrentMap As tCurrentPlayerMap
    
    Class As eClass
    Gender As eGenero
    Race As eRaza
    Intervals As tPlayerIntervals
    
    CraftingRecipeGroups() As tCraftingRecipeGroup
    CraftingRecipeGroupsQty As Integer
    
End Type


' Game Metadata
Public Type tMetadataMastery
    Id As Integer
    Name As String
    Description As String
    Enabled As Boolean
    RequiredMastery As Integer
    RequiredPoints As Integer
    RequiredGold As Long
    IconGrh As Integer
    
End Type

Public Type tMetadataClassMasteries
    ClassMasteriesIds() As Integer
End Type

Public Type tNPC
    Name As String
    MiniatureFileName As String
End Type

Public Type tObjData
    Name As String
    GrhIndex As Integer
    OBJType As eObjType
    
    Real As Integer
    Caos As Integer

    MaxHit As Integer
    MinHit As Integer
    MaxDef As Integer
    MinDef As Integer
    MinimumLevel As Byte
    Valor As Long
End Type

Public Type tGameMetadata
    MasteriesQty As Integer
    Masteries() As tMetadataMastery
    
    ClassMasteries() As tMetadataClassMasteries
    
    GuildQuestsQty As Integer
    GuildQuests() As tQuest
    
    Npcs() As tNPC
    NpcsQty As Integer
    
    Objs() As tObjData
    ObjsQty As Integer
End Type

Public EstadoLogin As E_MODO

Public MD5 As clsMD5

Public Declare Sub OutputDebugString Lib "kernel32" Alias "OutputDebugStringA" (ByVal lpOutputString As String)

Public GameMetadata As tGameMetadata
Public PlayerData As tPlayerData

Public Const MAP_URL As String = "https://manual.alkononline.com.ar/index.php?title=El_mundo"

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


Public Type tFont
    red As Byte
    green As Byte
    blue As Byte
    bold As Boolean
    italic As Boolean
End Type


Public FontTypes(23) As tFont

