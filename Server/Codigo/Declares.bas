Attribute VB_Name = "Declaraciones"
'Argentum Online 0.14.0
'Copyright (C) 2002 Márquez Pablo Ignacio
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

''
' Modulo de declaraciones. Aca hay de todo.
'

Public Aurora_Network As Aurora_Engine.Network_Service

' Debug
Public Declare Sub OutputDebugString Lib "kernel32" Alias "OutputDebugStringA" (ByVal lpOutputString As String)

' Tipos de present effect para enviar al cliente.
Public Enum ePresentEffect
    SpawnBoss
End Enum

'Si el tiempo de inactividad esta desactivado.
Public IdleOff As Boolean

#If EnableSecurity Then
Public aDos As clsAntiDoS
#End If

Public aClon As clsAntiMassClon
Public TrashCollector As Collection

Public Const INFINITE_LOOPS As Integer = -1

''
' The color of chats over head of dead characters.
Public Const CHAT_COLOR_DEAD_CHAR As Long = &HC0C0C0

''
' The color of yells made by any kind of game administrator.
Public Const CHAT_COLOR_GM_YELL As Long = &HF82FF


' Punishment type ids for reserved types.
Public Enum PunishmentStaticIds
    UnBanChar = 1
    PermanentBan = 2
End Enum

''
' Coordinates for normal sounds (not 3D, like rain)
Public Const NO_3D_SOUND As Byte = 0

Public Enum eMenues
    ieComerciante = 1
    ieSacerdote
    ieGobernador
    ieMascota
    ieMascotaQuieta
    ieEntrenador
    ieFogata
    ieFogataDescansando
    ieBanquero
    ieEnlistadorFaccion
    ieApostador
    ieYunque
    ieFragua
    ieOtroUser
    ieOtroUserCompartiendoNpc
    ieNpcDomable
    ieLenia
    ieRamas
End Enum

Public Enum eMenuAction
    ieCommerce = 1
    iePriestHeal
    ieHogar
    iePetStand
    iePetFollow
    ieReleasePet
    ieTrain
    ieSummonLastNpc
    ieRestToggle
    ieBank
    ieFactionEnlist
    ieFactionReward
    ieFactionWithdraw
    ieFactionInfo
    ieGamble
    ieBlacksmith
    ieMakeLingot
    ieMeltDown
    ieShareNpc
    ieStopSharingNpc
    ieTameNpc
    ieMakeFireWood
    ieLightFire
End Enum

Public Enum eMimeType
    ieNone
    ieTerrain
    ieAquatic
    ieBoth
End Enum

Public Enum iMinerales
    HierroCrudo = 192
    PlataCruda = 193
    OroCrudo = 194
    LingoteDeHierro = 386
    LingoteDePlata = 387
    LingoteDeOro = 388
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

Public Enum ePrivileges
    Admin = 1
    Dios
    Especial
    SemiDios
    Consejero
    RoleMaster
End Enum

Public Enum eClass
    Mage = 1       'Mago
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
    cArkhein
    cLastCity
End Enum

Public Enum eRaza
    Humano = 1
    Elfo
    Drow
    Gnomo
    Enano
End Enum

Enum eGenero
    Hombre = 1
    Mujer
End Enum

Public Enum eMoveType
    Inventory = 1
    Bank
End Enum

Public Type tCabecera 'Cabecera de los con
    Desc As String * 255
    Crc As Long
    MagicWord As Long
End Type

Public MiCabecera As tCabecera

Public Const MaxMascotasEntrenador As Byte = 7

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
    zonaOscura = 7
    CASA = 8
End Enum

''
' constantes para el trigger 6
'
' @see eTrigger
' @param TRIGGER6_PERMITE TRIGGER6_PERMITE
' @param TRIGGER6_PROHIBE TRIGGER6_PROHIBE
' @param TRIGGER6_AUSENTE El trigger no aparece
'
Public Enum eTrigger6
    TRIGGER6_PERMITE = 1
    TRIGGER6_PROHIBE = 2
    TRIGGER6_AUSENTE = 3
End Enum

Public Enum eTerrainZone
    terrain_bosque = 0
    terrain_nieve = 1
    terrain_desierto = 2
    zone_ciudad = 3
    zone_campo = 4
    zone_dungeon = 5
End Enum

Public Enum eRestrict
    restrict_no = 0
    restrict_newbie = 1
    restrict_armada = 2
    restrict_caos = 3
    restrict_faccion = 4
End Enum
' <<<<<< Targets >>>>>>
Public Enum TargetType
    uUsuarios = 1
    uNPC = 2
    uUsuariosYnpc = 3
    uTerreno = 4
End Enum

Public Enum eReviveTarget
    User = 1
    Pet
End Enum

' <<<<<< Acciona sobre >>>>>>
Public Enum TipoHechizo
    uPropiedades = 1
    uEstado = 2
    uMaterializa = 3    'Nose usa
    uInvocacion = 4
End Enum

Public Const MAXUSERHECHIZOS As Byte = 35
Public Const MAXUSERPASSIVES As Byte = 6

' La utilidad de esto es casi nula, sólo se revisa si fue a la cabeza...
Public Enum PartesCuerpo
    bCabeza = 1
    bPiernaIzquierda = 2
    bPiernaDerecha = 3
    bBrazoDerecho = 4
    bBrazoIzquierdo = 5
    bTorso = 6
End Enum

Public Const Guardias As Integer = 6

Public Const MAX_ORO_EDIT As Long = 5000000
Public Const MAX_VIDA_EDIT As Long = 30000

Public Const MAX_PP_EDIT As Long = 100000
Public Const MIN_PP_EDIT As Long = -100000

Public Const STANDARD_BOUNTY_HUNTER_MESSAGE As String = "Se te ha otorgado un premio por ayudar al proyecto reportando bugs, el mismo está disponible en tu bóveda."
Public Const TAG_CONSULT_MODE As String = "[CONSULTA]"

Public Const MaxOro As Long = 90000000

Public Const MAXNPCS As Integer = 10000
Public Const MAXCHARS As Integer = 10000

Public Const RED_PESCA As Integer = 543

Public Enum eNPCType
    Comun = 0
    Revividor = 1
    GuardiaReal = 2
    Entrenador = 3
    Banquero = 4
    Noble = 5
    DRAGON = 6
    Timbero = 7
    GuardiasCaos = 8
    ResucitadorNewbie = 9
    Pretoriano = 10
    Gobernador = 11
    GuildMaster = 12
End Enum

'********** CONSTANTANTES ***********

''
' Cantidad de skills
Public Const NUMSKILLS As Byte = 18

''
' Cantidad de Atributos
Public Const NUMATRIBUTOS As Byte = 5

''
' Cantidad de Clases
Public Const NUMCLASES As Byte = 11

''
' Cantidad de Razas
Public Const NUMRAZAS As Byte = 5


''
' Valor maximo de cada skill
Public Const MAX_SKILL_POINTS As Byte = 100

''
' Valor maximo de skills libres
Public Const MAX_SKILLS_LIBRES As Integer = 1000

''
' Cantidad de Ciudades
Public Const NUMCIUDADES As Byte = 6

'Número de pjs por cuenta.
Public Const MAX_ACCOUNT_CHARS As Byte = 8

''
'Direccion
'
' @param NORTH Norte
' @param EAST Este
' @param SOUTH Sur
' @param WEST Oeste
'
Public Enum eHeading
    NORTH = 1
    EAST = 2
    SOUTH = 3
    WEST = 4
End Enum

Public Const Pescado As Byte = 139

Public ListaPeces() As Integer


'%%%%%%%%%% CONSTANTES DE INDICES %%%%%%%%%%%%%%%
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

Public Enum eMochilas
    Mediana = 1
    Grande = 2
End Enum

Public Const FundirMetal = 88

Public Enum eAtributos
    Fuerza = 1
    Agilidad = 2
    Inteligencia = 3
    Carisma = 4
    Constitucion = 5
End Enum

Public Enum eProfessions
    Woodcutting
    Fishing
    Minning
    Carpentry
    Blacksmithing
    Tailoring
End Enum

'Tamaño del mapa
Public Const XMaxMapSize As Byte = 100
Public Const XMinMapSize As Byte = 1
Public Const YMaxMapSize As Byte = 100
Public Const YMinMapSize As Byte = 1

'Tamaño en Tiles de la pantalla de visualizacion
Public Const XWindow As Byte = 17
Public Const YWindow As Byte = 13

' Cantidad maxima de objetos por slot de inventario
Public Const MAX_INVENTORY_OBJS As Integer = 10000

''
' Cantidad de "slots" en el inventario con mochila
Public Const MAX_INVENTORY_SLOTS As Byte = 30

''
' Cantidad de "slots" en el inventario sin mochila
Public Const MAX_NORMAL_INVENTORY_SLOTS As Byte = 25

''
' Cantidad de "slots" en el inventario de quest
Public Const MAX_QUEST_INVENTORY_SLOTS As Byte = 20

''
' Constante para indicar que se esta usando ORO
Public Const FLAGORO As Integer = MAX_INVENTORY_SLOTS + 1


' CATEGORIAS PRINCIPALES
Public Enum eOBJType
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
    otESCUDO = 16
    otCASCO = 17
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
    otYacimientoPez = 38
    otTrigger = 39
    otSurpriseBox = 40
    otGuildBook = 42
    otCraftingMaterial = 43
    otTrampa = 44
    otResource = 45         ' Extractable Resouces (Trees, Mines, Fish pools, etc)
    otTool = 46             ' Tools used to extract resources. TODO: Make the crafting tools to use this type
    otQuest = 47
    otCualquiera = 1000
End Enum

Public Enum eDamageType
    Weapon
    Spell
    BareHand
    NpcDamage
    NpcSpell
End Enum

'Texto
Public Const FONTTYPE_TALK As String = "~255~255~255~0~0"
Public Const FONTTYPE_FIGHT As String = "~255~0~0~1~0"
Public Const FONTTYPE_WARNING As String = "~32~51~223~1~1"
Public Const FONTTYPE_INFO As String = "~65~190~156~0~0"
Public Const FONTTYPE_INFOBOLD As String = "~65~190~156~1~0"
Public Const FONTTYPE_EJECUCION As String = "~130~130~130~1~0"
Public Const FONTTYPE_PARTY As String = "~255~180~255~0~0"
Public Const FONTTYPE_VENENO As String = "~0~255~0~0~0"
Public Const FONTTYPE_GUILD As String = "~255~255~255~1~0"
Public Const FONTTYPE_SERVER As String = "~0~185~0~0~0"
Public Const FONTTYPE_GUILDMSG As String = "~228~199~27~0~0"
Public Const FONTTYPE_CONSEJO As String = "~130~130~255~1~0"
Public Const FONTTYPE_CONSEJOCAOS As String = "~255~60~00~1~0"
Public Const FONTTYPE_CONSEJOVesA As String = "~0~200~255~1~0"
Public Const FONTTYPE_CONSEJOCAOSVesA As String = "~255~50~0~1~0"
Public Const FONTTYPE_CENTINELA As String = "~0~255~0~1~0"

' **************************************************************
' **************************************************************
' ************************ TIPOS *******************************
' **************************************************************
' **************************************************************


Public Type tEloCombatResult
    AttackerPreviousPoints As Long
    AttackerNewPoints As Long
    AttackerPointsDifference As Long
    
    VictimPreviousPoints As Long
    VictimNewPoints As Long
    VictimPointsDifference As Long
    
    SkewDistanceUsed As Integer
    
End Type

Public Type tObservacion
    Creador As String
    Fecha As Date
    
    Detalles As String
End Type

Public Type tRecord
    Usuario As String
    Motivo As String
    Creador As String
    Fecha As Date
    
    NumObs As Byte
    Obs() As tObservacion
End Type

Private Type tDamagerOverTimeData
    IsDot As Boolean
    WaitForFirstTick As Boolean
    TickCount As Integer
    TickInterval As Long
    MaxStackEffect As Integer
End Type

Public Type tHechizo
    Nombre As String
    Desc As String
    PalabrasMagicas As String
    
    HechizeroMsg As String
    TargetMsg As String
    PropioMsg As String
    
'    Resis As Byte
    
    Putrefaccion As Byte
    Teletransportacion As Byte
    Salta As Byte
    DistanciaSalto As Byte
    Petrificar As Byte
    
    Area As Byte
    AreaEfficacy() As Byte
    CasterAffected As Boolean
    
    Atraer As Byte
    ByPassPassive As Byte
    
    Tipo As TipoHechizo
    
    WAV As Integer
    FXgrh As Integer
    Loops As Byte
    
    SubeHP As Byte
    MinHp As Integer
    MaxHp As Integer
    
    SubeMana As Byte
    MinMana As Integer
    MaxMana As Integer
    
    SubeSta As Byte
    MinSta As Integer
    MaxSta As Integer
    
    SubeHam As Byte
    MinHam As Integer
    MaxHam As Integer
    
    SubeSed As Byte
    MinSed As Integer
    MaxSed As Integer
    
    SubeAgilidad As Byte
    MinAgilidad As Integer
    MaxAgilidad As Integer
    
    SubeFuerza As Byte
    MinFuerza As Integer
    MaxFuerza As Integer
    
    SubeCarisma As Byte
    MinCarisma As Integer
    MaxCarisma As Integer
    
    Invisibilidad As Byte
    Paraliza As Byte
    Inmoviliza As Byte
    RemoverParalisis As Byte
    RemoverEstupidez As Byte
    CuraVeneno As Byte
    Envenena As Byte
    Maldicion As Byte
    RemoverMaldicion As Byte
    Bendicion As Byte
    Estupidez As Byte
    Ceguera As Byte
    Revivir As eReviveTarget
    Morph As Byte
    Mimetiza As Byte
    RemueveInvisibilidadParcial As Byte
    LifeLeechPerc As Byte
    
    Warp As Byte
    Invoca As Byte
    NumNpc As Integer
    cant As Integer

'    Materializa As Byte
'    ItemIndex As Byte
    
    MinSkill As Integer
    ManaRequerido As Integer
    ManaRequeridoPerc As Integer
    
    RequireFullMana As Byte

    'Barrin 29/9/03
    StaRequerido As Integer

    TargetUser As Boolean
    TargetNpc As Boolean
    TargetObj As Boolean
    TargetTerrain As Boolean

    MagicCastPowerRequired As Integer
    
    MinLevel As Byte
    
    DamageOverTime As tDamagerOverTimeData
    IgnoreMagicDefensePerc As Byte
    
    SpellCastInterval As Double
End Type

Public Type LevelSkill
    LevelValue As Integer
End Type

Public Type UserOBJ
    ObjIndex As Integer
    Amount As Integer
    Equipped As Byte
End Type

Public Type Inventario
    Object(1 To MAX_INVENTORY_SLOTS) As UserOBJ
    WeaponEqpObjIndex As Integer
    WeaponEqpSlot As Byte
    ArmourEqpObjIndex As Integer
    ArmourEqpSlot As Byte
    EscudoEqpObjIndex As Integer
    EscudoEqpSlot As Byte
    CascoEqpObjIndex As Integer
    CascoEqpSlot As Byte
    MunicionEqpObjIndex As Integer
    MunicionEqpSlot As Byte
    AnilloEqpObjIndex As Integer
    AnilloEqpSlot As Byte
    BarcoObjIndex As Integer
    BarcoSlot As Byte
    MochilaEqpObjIndex As Integer
    MochilaEqpSlot As Byte
    FactionArmourEqpObjIndex As Integer
    FactionArmourEqpSlot As Byte
    NroItems As Integer
End Type

Public Type Position
    X As Integer
    Y As Integer
End Type

Public Type WorldPos
    Map As Integer
    X As Integer
    Y As Integer
End Type

Public Type FXdata
    Nombre As String
    GrhIndex As Integer
    Delay As Integer
End Type

'Datos de user o npc
Public Type Char
    CharIndex As Integer
    head As Integer
    body As Integer
    
    WeaponAnim As Integer
    ShieldAnim As Integer
    CascoAnim As Integer
    
    FX As Integer
    Loops As Integer
    
    heading As eHeading
End Type

Private Type tTriggerSpell
    Index As Integer    ' Spell Index
    
    WAV As Integer      ' Overrides spell-wav
    FXgrh As Integer    ' Overrides spell-fx
    Loops As Byte       ' Overrides spell-Loops
    Interval As Long    ' Overrides spell-interval
    MinHit As Integer   ' Overrides spell-MinHit
    MaxHit As Integer   ' Overrides spell-MaxHit
    
    InvokeNpcIndex As Integer ' Invokes npc
    
    DamageOverTime As tDamagerOverTimeData
    
End Type

' Trigger Data
Private Type tTriggerData
    Visible As Byte ' Users can see it or not
    Dissapears As Byte ' Dissapears when stepping on it?
    Animation As Integer ' Has own animation (-1 no animation)
    CanDetect As Byte ' Can be detected on click?
    CanDisarm As Byte ' Can be Disarm?
    CanTake As Byte ' Can be taken in order to use it again.
    
    AffectNpc As Boolean
    AffectUser As Boolean
    
    ActivationMessage As String
    
    NumSpells As Byte
    Spells() As tTriggerSpell
    
    ' Damage over time triggers
End Type

Public Type tSpellPosition
    Pos As WorldPos
    DistanceFromTarget As Byte
End Type

Public Type tAttackPosition
    ReducedDamageFromSplash As Boolean
    Pos As WorldPos
End Type

Public Const PROB_MULTIPLIER As Long = 10000000
Public Const MAX_DIGIT_SB_RND As Long = 100 * PROB_MULTIPLIER

Public Type tSurpriseObj 'No uso el del drop complejo de npc porque aca tmb llevan probabilidad
    ObjIndex As Integer
    Amount As Long
    prob As Long
End Type
    
Public Type tSurpriseDrops
    NroItems As Long
    Drop() As tSurpriseObj
End Type

Public Type tCraftingItem
    ObjIndex As Integer
    Amount As Integer
End Type

Public Enum eCraftingStoreMaterialProvider
    StoreOwnerMaterials
    CustomerMaterials
End Enum

Public Type tCraftingStoreItem
    Recipe As Integer
    RecipeIndex As Integer
    RecipeItem As Integer
    ConstructionPrice As Long
    MaterialsPrice As Long
    RecipeGroup As Byte
    AmountCrafted As Long
    MoneyEarned As Double
    ProfessionType As Byte
End Type

Public Type tCraftingStore
    Items() As tCraftingStoreItem
    ItemsQty As Integer
    StoreType As eCraftingStoreMaterialProvider
    IsOpen As Boolean
    
    CraftedObjectsQty As Long
    MoneyEarned As Double
    ProfessionType As Byte
    LastCraftedObjectAt As Long
    
    InstanceId As String
End Type

Public Type tUserCraftingRecipes
    Calculated As Boolean

    CarpentryItemsQty As Integer
    CarpentryItems() As Integer
    
    BlacksmithWeaponsQty As Integer
    BlacksmithWeapons() As Integer
    
    BlacksmithArmorsQty As Integer
    BlacksmithArmors() As Integer
End Type

Public Type tProfessionCrafgintRecipe

    ObjIndex As Integer
    Materials() As tCraftingItem
    MaterialsQty As Integer
    
    CraftingAmount As Integer
    
    RecipeIndex As Integer
    
    CraftingProbability As Byte
    BlacksmithSkillNeeded As Byte
    CarpenterSkillNeeded As Byte
    TailoringSkillNeeded As Byte
    ProduceAmount As Integer
    
End Type

Public Type tProfessionCraftingRecipeGroup
    TabTitle As String
    TabImage As String
    DatFileName As String
    
    ProfessionType As Byte
    
    RecipesQty As Integer
    Recipes() As tProfessionCrafgintRecipe
End Type


'Profesiones
Public Type tProfession
    Name As String
    Enabled As Boolean
    SkillNumber As Byte
    SkillExpSuccess As Byte
    SkillExpFailure As Byte
    RequiredStaminaWorker As Byte
    RequiredStaminaOther As Byte
    MinRemovableResourcesPercent As Byte
    MaxRemovableResourcesPercent As Byte
    EnabledInSafeZone As Boolean
    SuccessFx As Integer
    
    CraftingRecipeGroupsQty As Byte
    CraftingRecipeGroups() As tProfessionCraftingRecipeGroup
End Type

'Recursos
Public Type tResource
    ResourceNumber As Integer
    ObjIndex As Integer
    ExtractionProbability As Single
    MinPerTickWorker As Integer
    MaxPerTickWorker As Integer
    MinPerTickOther As Integer
    MaxPerTickOther As Integer
    MaxAvailableQuantity As Integer
    UnlimitedResource As Boolean
    MinToolPower As Byte
End Type

Public Type tMinMaxIsPercent
    Min As Integer
    Max As Integer
    IsPercent As Boolean
End Type


'Tipos de objetos
Public Type ObjData
    Name As String 'Nombre del obj
    
    ObjType As eOBJType 'Tipo enum que determina cuales son las caract del obj
    
    GrhIndex As Integer ' Indice del grafico que representa el obj
    GrhSecundario As Integer
    
    'Solo contenedores
    MaxItems As Integer
    Conte As Inventario
    Apuñala As Byte
    Acuchilla As Byte
    
    Cupos As Byte 'Cupos que le da al clan al usarlo
    
    TrapActivatedObject As Integer
    TrapActivableLevelActivate As Integer
    TrapActivableLevelDeactivate As Integer
    TrapActivable As Byte
    
    StabDamageReduction As Byte
    
    HechizoIndex As Integer
    
    ForoID As String
    
    MinHp As Integer ' Minimo puntos de vida
    MaxHp As Integer ' Maximo puntos de vida
    
    
    MineralIndex As Integer
    LingoteInex As Integer
    
    
    proyectil As Integer
    Municion As Integer
    
    Crucial As Byte
    Newbie As Integer
    ItemGM As Byte
    
    'Puntos de Stamina que da
    MinSta As Integer ' Minimo puntos de stamina
    
    'Pociones
    TipoPocion As Byte
    
    AffectsMana As tMinMaxIsPercent
    AffectsHealth As tMinMaxIsPercent
    AffectsAgility As tMinMaxIsPercent
    AffectsStrength As tMinMaxIsPercent
    
    MaxModificador As Integer
    MinModificador As Integer
    DuracionEfecto As Long
    MinSkill As Integer
    LingoteIndex As Integer
    
    MinHit As Integer 'Minimo golpe
    MaxHit As Integer 'Maximo golpe
    
    MinHam As Integer
    MinSed As Integer
    
    Def As Integer
    MinDef As Integer ' Armaduras
    MaxDef As Integer ' Armaduras
    
    Ropaje As Integer 'Indice del grafico del ropaje
    
    NumRopajeGenerico As Integer
    NumRopajeMujerAlto As Integer
    NumRopajeHombreAlto As Integer
    NumRopajeMujerBajo As Integer
    NumRopajeHombreBajo As Integer
    NumRopajeMujerDrow As Integer
    NumRopajeHombreDrow As Integer
    
    NumBodyNeutral As Integer
    NumBodyRoyal As Integer
    NumBodyLegion As Integer
    
    WeaponAnim As Integer ' Apunta a una anim de armas
    WeaponRazaEnanaAnim As Integer
    ShieldAnim As Integer ' Apunta a una anim de escudo
    CascoAnim As Integer
    
    Valor As Long     ' Precio
    SalePrice As Long ' Precio de venta
    
    Cerrada As Integer
    Llave As Byte
    clave As Long 'si clave=llave la puerta se abre o cierra
    
    Radio As Integer ' Para teleps: El radio para calcular el random de la pos destino
    
    MochilaType As Byte 'Tipo de Mochila (1 la chica, 2 la grande)
    
    Guante As Byte ' Indica si es un guante o no.
    Critical As Byte ' Critical damage?
    
    IndexAbierta As Integer
    IndexCerrada As Integer
    IndexCerradaLlave As Integer
    
    Mujer As Byte
    Hombre As Byte
    
    Envenena As Byte
    Paraliza As Byte
    
    Agarrable As Byte
    
    Luminous As Boolean
    LightOffsetX As Integer
    LightOffsetY As Integer
    LightSize As Integer
    
    CanBeTransparent As Boolean
    
    LingH As Integer
    LingO As Integer
    LingP As Integer
    Madera As Integer
    MaderaElfica As Integer
    
    SkHerreria As Integer
    SkCarpinteria As Integer
    
    texto As String
    
    'Clases que no tienen permitido usar este obj
    ClaseProhibida(1 To NUMCLASES) As eClass
    
    RazaProhibida(1 To NUMRAZAS) As eRaza
    
    Snd1 As Integer
    Snd2 As Integer
    Snd3 As Integer
    
    Real As Integer
    Caos As Integer
    
    NoSeCae As Integer
    NoSeTira As Integer
    NoRobable As Integer
    NoComerciable As Integer
    Intransferible As Integer
    
    MagicCastPower As Integer
    MagicDamageBonus As Integer
    
    DefensaMagicaMax As Integer
    DefensaMagicaMin As Integer
    Perforation As Byte
    TwoHanded As Byte
    
    ImpideParalizar As Byte
    ImpideInmobilizar As Byte
    ImpideAturdir As Byte
    ImpideCegar As Byte

    Log As Byte 'es un objeto que queremos loguear? Pablo (ToxicWaste) 07/09/07
    NoLog As Byte 'es un objeto que esta prohibido loguear?
    
    Upgrade As Integer
    
    MenuIndex As Byte
    ItemQuest As Byte
    
    Trigger As tTriggerData ' Trigger info
    
    SurpriseDrops As tSurpriseDrops
    
    ' Resource Extraction System
    DepletedGrhIndex As Integer 'ObjIndex to use when the resource is emptied
    SoundNumber As Integer 'WAV to play after the resource is emptied
    MaxExtractedQuantity As Long 'Max resource qty to gather
    RespawnCooldown As Long 'Time to respawn after emptied. Used by State Server
    NumResources As Byte ' Amount of resources that will be loaded
    Resources() As tResource 'Resources that can be extracted from this object
    ProfessionType As Byte 'Profession index
    ToolPower As Byte 'For tools, determines the power to extract certain resources.
    ' /Resource Extraction System
    
    MaxStaRecoveryPerc As Byte ' Maximum stamina recovery percentage. Used by campfires
    DisappearTimeInSec As Integer ' Time that is considered by the state server to remove the object after it was created. Used by campfires
    
    CampfireObj As Integer ' Object to be used when spawning a new campfire.
    
    MaxDistanceFromTarget As Byte
    RequiredStamina As Integer
    
    SplashDamage As Boolean
    SplashDamageType As Byte
    SplashDamageReduction As Double
    
    SizeWidth As Byte   ' Entity size
    SizeHeight As Byte  ' Entity size
    
    AllowResting As Boolean
    MinimumLevel As Byte
End Type

Public Type Obj
    ObjIndex As Integer
    Amount As Integer
    
    CurrentGrhIndex As Integer
    PendingQty As Long 'Pending qty, decreased when workers use the resource
    Resources() As tResource 'Resources that can be extracted from this object
    ActivatedByUser As Integer 'Hunter who active
End Type

'''''''
' QUEST
'''''''

Public Type tQuestObj
    ObjIndex As Integer
    ObjQty As Long
End Type

Public Type tQuestRewards
    Gold As Long
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
    Army As tQuestFragAlign
    Legion As tQuestFragAlign
    MinLevel As Integer
End Type

Public Type tQuestStage

    StarterNpc As tQuestNpc
    EndNpc As tQuestNpc

    ObjsCollectQuantity As Integer
    ObjsCollect() As modRequiredObjectList.RequiredObjectListItem
    
    NpcsKillsQuantity As Integer
    NpcKill() As tQuestNpc

    Frags As tQuestFrags
    Rewards As tQuestRewards

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

Public GuildQuestList() As tQuest

Public Enum eQuestUserAlign
    Army = 1
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

Public Enum eQuestRequirementDb
    NpcKill = 1
    ObjCollect
    FragNeutral
    FragArmada
    FragLegion
End Enum


Public Type tModClassRace
    StartingHealth As Integer
    HealthPerLevelMin As Integer
    HealthPerLevelMax As Integer
    
    ExtraHealthAtLevel() As Integer
End Type

'[Pablo ToxicWaste]
Public Type ModClase
    Evasion As Double
    AtaqueArmas As Double
    AtaqueProyectiles As Double
    AtaqueWrestling As Double
    DamageWeapons As Double
    DamageProjectiles As Double
    DamageWrestling As Double
    PhysicalDamage As Double
    Escudo As Double
    Taming As Double
    Work As Double
    
    BaseDamage As Double
    DistanceDmgReduction As Double
    DistanceDamageReductionStart As Byte
    
    StabChance As Byte
    StabDamageMultiplier As Double
    
    ManaPerLevelMultiplier As Single
    ManaStarterMultiplier As Single
    
    HealthPerLevelMin As Byte
    HealthPerLevelMax As Byte
    
    StaminaPerLevel As Integer
    StaminaStarter As Integer
    
    SkillsPerLevel As Integer
    SkillsStarter As Integer
    
    MagicDamageBonus As Integer
    MagicCastPower As Integer
    
    MaxInvokedPets As Integer
    MaxTammedPets As Integer
    MaxActivePets As Integer
    
    DamageWrestlingMin As Integer
    DamageWrestlingMax As Integer
    
    HidingChance As Integer
    HidingDuration As Double
    
    StealingChance As Integer
    StealingAmount As Double
    
    StartingHealth As Integer
    
End Type

Public Type ModRaza
    Fuerza As Single
    Agilidad As Single
    Inteligencia As Single
    Carisma As Single
    Constitucion As Single
End Type
'[/Pablo ToxicWaste]

'[KEVIN]
'Banco Objs
Public Const MAX_BANCOINVENTORY_SLOTS As Byte = 40
Public Const MAX_BANCOINVENTORY_SLOTS_FIX As Byte = 35
'[/KEVIN]

'[KEVIN]
Public Type BancoInventario
    Object(1 To MAX_BANCOINVENTORY_SLOTS) As UserOBJ
    NroItems As Integer
End Type
'[/KEVIN]

' Determina el color del nick
Public Enum eNickColor
    ieNeutral = &H1
    ieCriminal = &H2
    ieCiudadano = &H3
    ieAtacable = &H4
End Enum

'*******
'FOROS *
'*******

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

' Estructura contenedora de mensajes
'Public Type tForo
    'StickyTitle(1 To Constantes.MaxStickyPost) As String
    'StickyPost(1 To Constantes.MaxStickyPost) As String
    'GeneralTitle(1 To Constantes.MaxGeneralPost) As String
    'GeneralPost(1 To Constantes.MaxGeneralPost) As String
'End Type

#If NewQuest = 0 Then

'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************
'******* T I P O S   D E    U S U A R I O S **************
'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************

Public Type tQuestRewards
    RewardGLD As Long
    RewardEXP As Long
    
    RewardSkills As Byte
    RewardSkill() As tQuestSkills
    
    RewardOBJs As Byte
    RewardOBJ() As Obj
    
    RewardSpells As Byte
    RewardSpell() As Byte
End Type

Public Type tStage

    Desc As String
    ' Final npc (End of stage; if = 0 => automatic reward)
    NpcDest As Integer
    TimeLimit As Long
    
    ' Requirements / objectives
    RequiredObjs As Byte
    RequiredObj() As Obj
    
    RequiredNpcs As Byte
    RequiredNpc() As tQuestNpc
    
    RequiredPk As Byte
    RequiredCitizens As Byte
    RequiredArmys As Byte
    RequiredChaos As Byte

    ' Rewards
    Rewards As tQuestRewards
End Type

Public Type tQuest

    Nombre As String
    Desc As String
    
    ' Requisitos
    RequiredLevelMin As Byte
    RequiredLevelMax As Byte
    RequiredStatus As Byte
    RequiredQuest As Integer
    RequiredRace As Byte
    RequiredClass As Byte
    RequiredGender As Byte
    RequiredFaction As Byte

    ' Limitaciones
    ForbiddenQuest As Byte
    ForbiddenRace As Byte
    ForbiddenClass As Byte

    Time As Long

    Repeatable As Byte
    
    ' Etapas
    NumStages As Integer
    Stages() As tStage
        
    ' Recompensas
    Rewards As tQuestRewards
End Type
#End If

' Passive skill type
Public Type tPassiveSkill
    Id As Integer
    Name As String
    Enabled As Boolean
    AllowedByClass As Boolean
    Active As Boolean
End Type

Public Type tUserSpell
    SpellNumber As Integer
    LastUsedAt As Long
    LastUsedSuccessfully As Boolean
End Type

'Estadisticas de los usuarios
Public Type UserStats
    GLD As Long 'Dinero
    Banco As Long
    
    MaxHp As Integer
    MinHp As Integer
    
    MaxSta As Integer
    MinSta As Integer
    MaxMan As Integer
    MinMAN As Integer

    MaxHam As Integer
    MinHam As Integer
    
    MaxAGU As Integer
    MinAGU As Integer
        
    Def As Integer
    Exp As Double
    ELV As Byte
    ELU As Long
    MasteryPoints As Integer
    
    DuelosGanados As Long
    DuelosPerdidos As Long
    OroDuelos As Long
    
    UserAtributos(1 To NUMATRIBUTOS) As Byte
    UserAtributosBackUP(1 To NUMATRIBUTOS) As Byte
    
    UserHechizos(1 To MAXUSERHECHIZOS) As tUserSpell
    
    UsuariosMatados As Long
    CriminalesMatados As Long
    NPCsMuertos As Long
    
    SkillPts As Integer
    
    ExpSkills(1 To NUMSKILLS) As Long
    EluSkills(1 To NUMSKILLS) As Long
    
    ' Don't use these attributes directly.
    ' Use the modSkills' functions.
    AssignedSkills(1 To NUMSKILLS) As Byte
    NaturalSkills(1 To NUMSKILLS) As Byte
    
    RankingPoints As Long
    
    UserPassives() As tPassiveSkill
    
End Type

'Flags
Public Type UserFlags
    Muerto As Byte '¿Esta muerto?
    Escondido As Byte '¿Esta escondido?
    Comerciando As Integer   ' 4 Ops. 1) MAXNPCS +1 = Bank, 1) > 0 trading with npc, 3) < 0 trading with user, 4) = 0 no trading.
    UserLogged As Boolean '¿Esta online?
    Meditando As Boolean
    Descuento As String
    Hambre As Byte
    Sed As Byte
    PuedeMoverse As Byte
    TimerLanzarSpell As Long
    PuedeTrabajar As Byte
    Envenenado As Byte
    Paralizado As Byte
    Inmovilizado As Byte
    Estupidez As Byte
    Ceguera As Byte
    Inmunidad As Byte
    invisible As Byte
    Maldicion As Byte
    Bendicion As Byte
    Oculto As Byte
    Desnudo As Byte
    Descansar As Boolean
    CastedSpellNumber As Integer
    CastedSpellIndex As Integer
    TomoPocion As Boolean
    TipoPocion As Byte
    Putrefaccion As Integer
    Petrificado As Byte
    
    AccountBank As Integer
    
    NoPuedeSerAtacado As Boolean
    AtacablePor As Integer
    ShareNpcWith As Integer
    
    DueloPublico As Integer
    DueloIndex As Byte
    DueloTeam As Byte
    
    Vuela As Byte
    Navegando As Byte
    Seguro As Boolean
    SeguroResu As Boolean
    
    DuracionEfecto As Long
    TargetNpc As Integer ' Npc señalado por el usuario
    TargetNpcTipo As eNPCType ' Tipo del npc señalado
    OwnedNpc As Integer ' Npc que le pertenece (no puede ser atacado)
    NpcInv As Integer

    Ban As Byte
    AdministrativeBan As Byte
    
    TargetUser As Integer ' Usuario señalado
    
    TargetObj As Integer ' Obj señalado
    TargetObjMap As Integer
    TargetObjX As Integer
    TargetObjY As Integer
    
    TargetMap As Integer
    TargetX As Integer
    TargetY As Integer
    
    TargetObjInvIndex As Integer
    TargetObjInvSlot As Integer
    
    AtacadoPorNpc As Integer
    AtacadoPorUser As Integer
    NPCAtacado As Integer
    Ignorado As Boolean
    
    HelpMode As Boolean
    HelpingUser As Integer
    HelpingUserName As String
    HelpedBy As Integer
    HelpedByUserName As String
    'EnConsulta As Boolean
    
    SendDenounces As Boolean
    
    StatsChanged As Byte
    Privilegios As PlayerType
    PrivEspecial As Boolean
    
    ValCoDe As Integer
        
    AdminInvisible As Byte
    AdminPerseguible As Boolean
    
    ChatColor As Long
    
    '[el oso]
    MD5Reportado As String
    '[/el oso]
    
    '[Barrin 30-11-03]
    TimesWalk As Long
    StartWalk As Long
    CountSH As Long
    '[/Barrin 30-11-03]
    
    '[CDT 17-02-04]
    UltimoMensaje As Byte
    '[/CDT]
    
    Silenciado As Byte
    
    Mimetizado As Byte
    MimetizadoType As Byte
    
    lastMap As Integer
    Traveling As Byte 'Travelin Band ¿?
    
    CountQuestTime As Boolean
    ParalizedBy As String
    ParalizedByIndex As Integer
    ParalizedByNpcIndex As Integer
    LastNpcInvoked As Integer
    
    TournamentState As Byte
    bStrDextRunningOutNotified As Boolean
    nCommerceSourceUser As Integer
    LastTamedPet As Integer
    
#If EnableSecurity Then
    SecurityFlags As tSecurityFlags
#End If
    ActiveTraps() As WorldPos
End Type

Public Type UserCounters
    IdleCount As Long
    AttackCounter As Integer
    HPCounter As Long
    RestingHPCounter As Long
    RegenerationCounter As Long
    STACounter As Long
    RestingSTACounter As Long
    Frio As Long
    Lava As Long
    COMCounter As Long
    AGUACounter As Long
    Veneno As Long
    Paralisis As Long
    Ceguera As Long
    Estupidez As Long
    Putrefaccion As Long
    PutrefaccionDmg As Long
    Petrificado As Long
    
    Inmunidad As Long
    Invisibilidad As Long
    TiempoOculto As Long
    
    Mimetismo As Long
    PiqueteC As Long
    Pena As Long
    SendMapCounter As WorldPos
    '[Gonzalo]
    Saliendo As Boolean
    Salir As Integer
    '[/Gonzalo]
    
    'Barrin 3/10/03
    tInicioMeditar As Long
    bPuedeMeditar As Boolean
    'Barrin
    
    TimerLanzarSpell As Long
    TimerPuedeAtacar As Long
    TimerPuedeUsarArco As Long
    TimerPuedeTrabajar As Long
    TimerUsar As Long
    TimerMagiaGolpe As Long
    TimerGolpeMagia As Long
    TimerGolpeUsar As Long
    TimerPuedeSerAtacado As Long
    TimerPerteneceNpc As Long
    TimerEstadoAtacable As Long
    TimerHide As Long
    
    failedUsageAttempts As Long
    
    goHome As Long
    
#If EnableSecurity Then
    SecurityCounters As tSecurityCounters
#End If
End Type

Public Enum eCharacterAlignment
    Newbie = 0
    Neutral = 1
    FactionRoyal = 2
    FactionLegion = 3
End Enum

'Cosas faccionarias.
Public Type tFacciones
    Alignment As eCharacterAlignment
    ArmadaReal As Byte
    FuerzasCaos As Byte
    NeutralsKilled As Long
    CriminalesMatados As Long
    CiudadanosMatados As Long
    RecompensasReal As Long
    RecompensasCaos As Long
    RecibioExpInicialReal As Byte
    RecibioExpInicialCaos As Byte
    RecibioArmaduraReal As Byte
    RecibioArmaduraCaos As Byte
    Reenlistadas As Byte
    NivelIngreso As Integer
    FechaIngreso As Date
    MatadosIngreso As Integer 'Para Armadas nada mas
    NextRecompensa As Integer
End Type

Public Type tCrafting
    Cantidad As Long
    PorCiclo As Integer
End Type

Public Type tUserMensaje
    Contenido As String
    Nuevo As Boolean
End Type

Private Type tTrainningData
    startTick As Long
    trainningTime As Long
End Type

Public Type tPet
    NpcIndex As Integer     ' The Index of the NPC in the server
    NpcNumber As Integer    ' The NPC Number. Previously known as PetType
    RemainingLife As Integer      ' The remaining life of the NPC.
    IsInvoked As Boolean
End Type

Public Type tAccountData
    Id As Long
    Name As String
    Email As String
    Password As String
    Token As String
    Status As Byte
    BanDetail As String
    CreationDate As String
    BankGold As Long
    BankPassword As String
End Type

Public Type tUserPunishment
    Id As Long
    Punisher As String
    Reason As String
    EndDate As Date
End Type

Public Type tUserGuild
    IdGuild As Integer
    GuildIndex As Integer
    
    GuildRange As Byte

    RoleId As Integer
    RoleIndex As Integer
    
    GuildMemberIndex As Integer
End Type


    
'Tipo de los Usuarios
Public Type User
    Name As String
    secName As String
    Id As Long
    
    CraftingStore As tCraftingStore
    
    OverHeadIcon As Integer

    LastCompletedPacket As Integer

    RestObjectCoords As WorldPos

    AccountName As String
    AccountId As Long
    AccountEmail As String
    
    AccountCharNames(1 To MAX_ACCOUNT_CHARS) As String
    nSessionId As Integer
    ClientTempCode As String
    bIsPremium As Boolean
    
    ShowName As Boolean 'Permite que los GMs oculten su nick con el comando /SHOWNAME
    
    Char As Char 'Define la apariencia
    OrigChar As Char
    
    Desc As String ' Descripcion
    DescRM As String
    
    clase As eClass
    raza As eRaza
    Genero As eGenero
    
    Hogar As eCiudad
        
    Invent As Inventario
    
    Pos As WorldPos
    VolverDueloPos As WorldPos
    
    'Outgoing and incoming messages
    Connection   As Network_Client
    ConnIDValida As Boolean
    IP           As String
    IPLong       As Long
    

    '[KEVIN]
    BancoInvent As BancoInventario
    '[/KEVIN]
    
    Counters As UserCounters
    bShowAccountForm As Boolean
    bForceCloseAccount As Boolean
    
    Construir As tCrafting
    
    TammedPets() As tPet
    TammedPetsCount As Integer
    
    InvokedPets() As tPet
    InvokedPetsCount As Integer
    
    
    'MascotasIndex() As Integer
    'MascotasType() As Integer
    'MascotasLife() As Integer
    
    'NroMascotas As Integer
    PetAliveCount As Integer
    
    SelectedPet As Byte
    
    Stats As UserStats
    flags As UserFlags
        
    Faccion As tFacciones
    
    CraftableElements As tUserCraftingRecipes
    
#If EnableSecurity Then
    Security As SecurityData
#End If

#If ConUpTime Then
    LogOnTime As Date
    UpTime As Long
#End If

    ComUsu As tCOmercioUsuario
    Challenge As t_challenge
    
    Guild As tUserGuild
   
    'FundandoGuildAlineacion As ALINEACION_GUILD     'esto esta aca hasta que se parchee el cliente y se pongan cadenas de datos distintas para cada alineacion
    ListeningGuild As Integer
    AspiranteA As Integer
    
    PartyIndex As Integer   'index a la party q es miembro
    PartySolicitud As Integer   'index a la party q solicito
    
    KeyCrypt As Integer

    
    CurrentInventorySlots As Byte
    
    trainningData As tTrainningData
    
    Mensajes() As tUserMensaje
    UltimoMensaje As Byte
    
    InstanceId As Long 'unique identifier for user
    
    Punishment As tUserPunishment
    
    DbConnectionEventId As Long
    
    Masteries As tUserMasteryBoost
    
    InvitationGuildIndex As Integer
End Type


'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************
'**  T I P O S   D E    N P C S **************************
'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************

Public Type tIntervalos
    Walk As Integer
    Hit As Integer
    MoveAttack As Integer
End Type

Public Type NPCStats
    Alineacion As Integer
    MaxHp As Long
    MinHp As Long
    MaxHit As Integer
    MinHit As Integer
    Def As Integer
    DefM As Integer
End Type

Public Type NpcCounters
    Paralisis As Long
    TiempoExistencia As Long
End Type

Public Type NPCFlags
    AfectaParalisis As Byte
    Domable As Integer
    ItemToTame As Integer
    Respawn As Byte
    NPCActive As Boolean '¿Esta vivo?
    Follow As Boolean
    Faccion As Byte
    AtacaDoble As Byte
    LanzaSpells As Byte
    
    ShowName As Byte
    
    Boss As Byte
    
    KeepHeading As Byte
    
    Invocador As Integer
    Invocacion() As Integer
    MaxInvocaciones As Byte
    
    ExpCount As Long
    
    OldMovement As TipoAI
    OldHostil As Byte
    
    DistanciaMaxima As Byte
    VolviendoOrig As Byte
    VolviendoInt As Byte
    
    AguaValida As Byte
    TierraInvalida As Byte
    
    Sound As Integer
    AttackedBy As String
    AttackedFirstBy As String
    BackUp As Byte
    RespawnOrigPos As Byte
    
    Envenenado As Byte
    Paralizado As Byte
    Inmovilizado As Byte
    invisible As Byte
    Maldicion As Byte
    Bendicion As Byte
    
    Snd1 As Integer
    Snd2 As Integer
    Snd3 As Integer
    
    isAffectedByDOT As Boolean
End Type

Public Type tCriaturasEntrenador
    NpcIndex As Integer
    NpcName As String
    tmpIndex As Integer
End Type

' New type for holding the pathfinding info
Public Type NpcPathFindingInfo
    Path() As tVertice      ' This array holds the path
    Target As Position      ' The location where the NPC has to go
    PathLenght As Integer   ' Number of steps *
    CurPos As Integer       ' Current location of the npc
    TargetUser As Integer   ' UserIndex chased
    NoPath As Boolean       ' If it is true there is no path to the target location
    
    '* By setting PathLenght to 0 we force the recalculation
    '  of the path, this is very useful. For example,
    '  if a NPC or a User moves over the npc's path, blocking
    '  its way, the function NpcLegalPos set PathLenght to 0
    '  forcing the seek of a new path.
    
End Type
' New type for holding the pathfinding info

Public Type tDrops
    DropIndex As Integer
    Probabilidad As Single
    NoExcluyente As Byte
End Type


Public Type Char_Acc_Data
    Id                  As Long
    Nick_Name           As String
    Alignment           As Byte
    IdGuild             As Long
    GuildName           As String
    Character           As Char
    Nivel               As Byte
    Pos_Map             As String
    Muerto              As Boolean
    bSailing            As Boolean
    JailRemainingTime   As Long
    Banned              As Boolean
End Type

Public Type npc
    Name As String
    Char As Char 'Define como se vera
    Desc As String
    
    ExtraBodies As Byte
    ExtraBody() As Integer
    ActualBody As Byte
    
    Tag As String
    
    NumInvocaciones As Byte
    NpcsInvocables() As Integer
    
    NPCtype As eNPCType
    Numero As Integer

    InvReSpawn As Byte

    Comercia As Integer
    Target As Long
    TargetNpc As Long
    TipoItems As Integer

    Veneno As Byte

    Pos As WorldPos 'Posicion
    Orig As WorldPos
    SkillDomar As Integer

    PathFinding As Byte 'Tipo de PathFinding que usa
    
    Intervalos As tIntervalos
    Timers As clsTimers
    
    Movement As TipoAI
    Attackable As Byte
    Hostile As Byte
    PoderAtaque As Long
    PoderEvasion As Long

    ' Quest
    NumQuests As Integer
    Quest() As Integer
    
    Owner As Integer

    GiveEXP As Long
    GiveEXPTierra As Long
    
    GiveGLD As Long
    Drop() As tDrops
    NroDrops As Integer
    
    'Flecha en NPC
    TengoFlechas(1 To 6) As Integer
    
    Stats As NPCStats
    flags As NPCFlags
    Contadores As NpcCounters
    
    Invent As Inventario
    CanAttack As Byte
    
    NroExpresiones As Byte
    Expresiones() As String ' le da vida ;)
    
    NroSpells As Byte
    Spells() As Integer  ' le da vida ;)
    
    '<<<<Entrenadores>>>>>
    NroCriaturas As Integer
    Criaturas() As tCriaturasEntrenador
    MaestroUser As Integer
    MaestroNpc As Integer
    Mascotas As Integer
    
    ' New!! Needed for pathfindig
    PFINFO As NpcPathFindingInfo

    'Hogar
    Ciudad As Byte
    
    MenuIndex As Byte
    
    'Para diferenciar entre clanes
    ClanIndex As Integer
    
    Exists As Boolean
    
    level As Byte
    OffsetReducedExp As Byte
    OffsetModificator As Single
    
    MasteryStarter As Boolean
    OverHeadIcon As Integer
    
    SizeWidth As Byte ' Graphical entity's width
    SizeHeight As Byte ' Graphical entity's height
    
    InstanceId As Long 'unique identifier for npc
End Type

'**********************************************************
'**********************************************************
'******************** Tipos del mapa **********************
'**********************************************************
'**********************************************************
'Spawn Pos
Public Type tNpcSpawnPos
    Pos() As Position
End Type

'Tile
Public Type MapBlock
    Blocked As Byte
    Graphic(1 To 4) As Integer
    UserIndex As Integer
    NpcIndex As Integer
    ObjInfo As Obj
    TileExit As WorldPos
    Trigger As eTrigger
End Type

' Extractable Resource System
Type tExtractableResourceGroup
    ObjNumber As Integer ' The ObjNumber associated to the resource
    
    ResourceList() As WorldPos ' A list of all the resource positions in the map
    ResourceQty As Integer ' The amount of resources of the same type
    
    EmptyResourcePositions() As WorldPos ' A list of positions where the resources are empty
    EmptyResourceQty As Integer ' The amount of empty resources
End Type

Type tMapExtractableResourceData
    ResourceGroup() As tExtractableResourceGroup ' A list of grouped resources
    ResourceGroupQty As Integer ' Amount of resource groups (grouped by obj number)
End Type
' / Extractable Resource System

'Info del mapa
Type MapInfo
    NumUsers As Integer
    
    NpcSpawnPos(0 To 1) As tNpcSpawnPos
    
    NumMusic As Byte
    Music() As Long
    
    Name As String
    StartPos As WorldPos
    OnDeathGoTo As WorldPos
    
    MapVersion As Integer
    Pk As Boolean
    MagiaSinEfecto As Byte
    NoEncriptarMP As Byte
    MismoBando As Byte
    Reverb As Byte
    
    ' Anti Magias/Habilidades
    InviSinEfecto As Byte
    ResuSinEfecto As Byte
    OcultarSinEfecto As Byte
    InvocarSinEfecto As Byte
    InmovilizarSinEfecto As Byte
    CraftingStoreAllowed As Boolean
    
    RoboNpcsPermitido As Byte
    
    MapaTierra As Byte
    
    NakedLosesHealth As Boolean
    NakedLosesEnergy As Boolean
    
    Terreno As Byte
    Zona As Byte
    Restringir As Byte
    BackUp As Byte
    
    MapResources As tMapExtractableResourceData    ' Holds all the information about resources in the map.
End Type

'********** V A R I A B L E S     P U B L I C A S ***********

Public Const MAX_NICKNAME_SIZE As Integer = 15

Public ULTIMAVERSION As String
Public OUTDATED_VERSION As String

Public ListaRazas(1 To NUMRAZAS) As String
Public SkillsNames(1 To NUMSKILLS) As String
Public ListaClases(1 To NUMCLASES) As String
Public ListaAtributos(1 To NUMATRIBUTOS) As String


Public RECORDusuarios As Long

'
'Directorios
'

''
'Ruta base del server, en donde esta el "server.ini"
Public IniPath As String

''
'Ruta base para guardar los chars
Public CharPath As String

''
'Ruta base para los archivos de mapas
Public MapPath As String

''
'Ruta base para los DATs
Public DatPath As String

''
'Bordes del mapa
Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte

''
'Numero de usuarios actual
Public NumUsers As Integer
Public LastUser As Integer
Public LastChar As Integer
Public NumChars As Integer
Public LastNPC As Integer
Public NumNpcs As Integer
Public NumMaps As Integer
Public NumObjDatas As Integer
Public NumNpcsDat As Integer
Public NumeroHechizos As Integer
Public AllowMultiLogins As Byte
Public IdleLimit As Integer
Public MaxUsers As Integer
Public HideMe As Byte
Public LastBackup As String
Public Minutos As String
Public haciendoBK As Boolean
Public PuedeCrearPersonajes As Integer
Public ServerSoloGMs As Integer
Public NumRecords As Integer
Public HappyHour As Single
Public HappyHourActivated As Boolean
Public HappyHourDays(1 To 7) As Single
Public lNumHappyDays As Long

''
'Esta activada la verificacion MD5 ?
Public MD5ClientesActivado As Byte


Public EnPausa As Boolean
Public EnTesting As Boolean


'*****************ARRAYS PUBLICOS*************************
Public UserList() As User 'USUARIOS
Public Npclist(1 To MAXNPCS) As npc 'NPCS
Public MapData() As MapBlock
Public MapInfo() As MapInfo
Public Hechizos() As tHechizo
Public charList(1 To MAXCHARS) As Integer
Public ObjData() As ObjData
Public NpcData() As npc
Public FX() As FXdata
Public SpawnList() As tCriaturasEntrenador
Public LevelSkill(1 To 50) As LevelSkill
Public ForbidenNames() As String
Public MD5s() As String
Public BanIps As Collection
Public Parties(1 To MAX_PARTIES) As clsParty
Public ModRaza(1 To NUMRAZAS) As ModRaza
Public ModTrabajo(1 To NUMCLASES) As Double
Public TablaExperiencia(1 To 50) As Long
Public DistribucionEnteraVida(1 To 5) As Integer
Public DistribucionSemienteraVida(1 To 4) As Integer
Public Ciudades(1 To NUMCIUDADES) As WorldPos
Public ListaCiudades(1 To NUMCIUDADES) As String
Public Records() As tRecord
Public Professions() As tProfession
Public Resources() As tResource
'*********************************************************

Type HomeDistance
    distanceToCity(1 To NUMCIUDADES) As Integer
End Type

Public Nix As WorldPos
Public Ullathorpe As WorldPos
Public Banderbill As WorldPos
Public Lindos As WorldPos
Public Arghal As WorldPos
Public Arkhein As WorldPos
Public Nemahuak As WorldPos

Public Prision As WorldPos
Public Libertad As WorldPos

Public Ayuda As cCola
Public Denuncias As cCola
Public ConsultaPopular As ConsultasPopulares
Public SonidosMapas As SoundMapInfo

Public Declare Function GetRealTickCount Lib "kernel32" Alias "GetTickCount" () As Long

Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef Destination As Any, ByVal Length As Long)

Public Enum e_ObjetosCriticos
    Manzana = 1
    Manzana2 = 2
    ManzanaNewbie = 467
End Enum

Public Enum eMessages
    DontSeeAnything
    NPCSwing
    NPCKillUser
    BlockedWithShieldUser
    BlockedWithShieldother
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
    Home
    CancelHome
    FinishHome
End Enum

Public Enum eGMCommands
    GMMessage = 1           '/GMSG
    ShowName                '/SHOWNAME
    OnlineRoyalArmy         '/ONLINEREAL
    OnlineChaosLegion       '/ONLINECAOS
    GoNearby                '/IRCERCA
    comment                 '/REM
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
    BanChar                 '/BAN
    UnBanChar               '/UNBAN
    NPCFollow               '/SEGUIR
    SummonChar              '/SUM
    SpawnListRequest        '/CC
    SpawnCreature           'SPA
    ResetNPCInventory       '/RESETINV
    CleanWorld              '/LIMPIAR
    ServerMessage           '/RMSG
    NickToIP                '/NICK2IP
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
    DumpIPTables            '/DUMPSECURITY
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
    ChangeMapInfoStealNpc   '/MODMAPINFO ROBONPC
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
    Descalificar            '/DESCALIFICAR
    Pelea                   '/PELEA Nick@Nick
    CerrarTorneo            '/CERRARTORNEO
    IniciarTorneo           '/HTORNEO
    TorunamentEdit
    RequestTournamentConfig
    AlterGuildName
    HigherAdminsMessage
    GetPunishmenttypelist   ' Get the list of punishment type enabled in the server.
    AdminChangeGuildAlign
    ChangeMapInfoNoInmo
    ChangeMapInfoMismoBando
    SpawnBoss
End Enum

'''''''
'' Pretorianos
'''''''
Public ClanPretoriano() As clsClanPretoriano

'Mensajes de los NPCs enlistadores (Nobles):
Public Const MENSAJE_REY_CAOS As String = "¿Esperabas pasar desapercibido, intruso? Los servidores del Demonio no son bienvenidos, ¡Guardias, a él!"
Public Const MENSAJE_REY_CRIMINAL_NOENLISTABLE As String = "Tus pecados son grandes, pero aún así puedes redimirte. El pasado deja huellas, pero aún puedes limpiar tu alma."
Public Const MENSAJE_REY_CRIMINAL_ENLISTABLE As String = "Limpia tu reputación y paga por los delitos cometidos. Un miembro de la Armada Real debe tener un comportamiento ejemplar."

Public Const MENSAJE_DEMONIO_REAL As String = "Lacayo de Tancredo, ve y dile a tu gente que nadie pisará estas tierras si no se arrodilla ante mi."
Public Const MENSAJE_DEMONIO_CIUDADANO_NOENLISTABLE As String = "Tu indecisión te ha condenado a una vida sin sentido, aún tienes elección... Pero ten mucho cuidado, mis hordas nunca descansan."
Public Const MENSAJE_DEMONIO_CIUDADANO_ENLISTABLE As String = "Siento el miedo por tus venas. Deja de ser escoria y únete a mis filas, sabrás que es el mejor camino."

Public Administradores As clsIniManager

Public Const MIN_AMOUNT_LOG As Integer = 1000
Public Const MIN_VALUE_LOG As Long = 50000
Public Const MIN_GOLD_AMOUNT_LOG As Long = 10000

Public AnimHogar(1 To 4) As Integer
Public AnimHogarNavegando(1 To 4) As Integer

'Caracteres
Public Const CAR_ESPECIALES = "áàäâÁÀÄÂéèëêÉÈËÊíìïîÍÌÏÎóòöôÓÒÖÔúùüûÚÙÜÛñÑ.,:;()¡!¿?-_"

''
' Constante para indicar que estamos bovedeando
Public Const TRADING_BANK As Integer = MAXNPCS + 1


' PUNISHMENT SUBTYTPES
Public listBanTypes() As tPunishmentType
Public listJailTypes() As tPunishmentType
Public listWarningTypes() As tPunishmentType

Public Enum ePunishmentSubType
    Jail = 1
    Ban = 2
    Warning = 3
End Enum

Public Type tPunishmentRule
    Count As Integer
    severity As Integer
End Type


Public Type tPunishmentType
    Id As Integer
    Name As String
    BaseType As Byte
    Rules() As tPunishmentRule
    EndDate As Date
    AddJail As Boolean
    AddBan As Boolean
    NextPunishment As Integer
End Type

Public Type tPunishmentDbResponse
    PunishmentTypeId As Integer
    PunishmentBaseType As Byte
    PunishmentSeverity As Long
    
    ForcedPunishmentTypeId As Integer
    ForcedPunismentBaseType As Byte
    ForcedPunishmentSeverity As Long
    
    LastInsertedPunishmentId As Long
End Type



' Let's try to avoid the MySQL Timeout
Public mysql_requery_time As Byte

'Passive skills Enum

Public Enum ePassiveSpells
    ParalysisImmunity = 1
    IndomitableWill
    Regeneration
    VitalRestoration
    Berserk
    LastElement 'Do not use as a passive
End Enum

'Constantes del servidor.

Public Type tConstantes
    MaxPrivateMessages As Byte
    TiempoCarcelPiquete As Byte
    MaxNpcDrops As Byte
    GoHomePenalty As Byte
    MaxDenuncias As Byte
    MaxPartyMembers As Byte
    MaxPartyDifLevel As Byte
    MaxMensajesForo As Byte
    MaxAnunciosForo As Byte
    GuildLevelMax As Byte
    MinGuildMembers As Integer
    MaxGuildMembers As Integer
    ElfCampfireDuration As Long
    NormalCampfireDuration As Long
    MinStaRecoveryPerc As Long
    MaxStaRecoveryPerc As Long
    CraftingStoreMap As Integer
    CraftingStoreOverheadIcon As Integer
End Type

Public Constantes As tConstantes

Public Type tConstantesBalance
    LimiteNewbie As Byte
    MaxRep As Long
    MaxOro As Double
    MaxExp As Double
    MaxUsersMatados As Long
    MaxAtributos As Byte
    MinAtributos As Byte
    MaxLvl As Byte
    MaxHp As Integer
    MaxSta As Integer
    MaxMan As Integer
    SelfWorkerMaps() As Integer
    SelfWorkerMapsQty As Integer
    
    EluSkillInicial As Integer
    ExpAciertoSkill As Integer
    SkillExpCampfireSuccess As Integer
    SkillExpNpcKilled As Integer
    ExpFalloSkill As Byte
    ModDefSegJerarquia As Single
    MinCrearPartyLevel As Byte
    IntMoveAttack As Integer
    ModGoldMultiplier As Double
    ModExpMultiplier As Double
    ModTrainingExpMultiplier As Double
    MaxActiveTrapQty As Integer
    RankingMinLevel As Integer
    PlayerRankingStartingPoints As Integer
    GuildRankingStartingPoints As Integer
    RankingSkewDistance As Integer
    
    DuelProhibitedSpellsQty As Integer
    DuelProhibitedSpells() As Integer
    
    HomeWaitingTime As Long
    
    FactionMinLevel As Byte
    FactionMaxRejoins As Byte
    
    AlignmentAttackActionMatrix(3, 3) As Boolean
    AlignmentHelpActionMatrix(3, 3) As Boolean
End Type

Public ConstantesBalance As tConstantesBalance

Public Type tConstantesFX
    FxSangre As Byte
    FxWarp As Byte
    FxMeditarGrande As Byte
End Type

Public ConstantesMeditations() As Long
Public ConstantesFX As tConstantesFX

Public Type tConstantesGRH
    FragataFantasmal As Integer
    FragataReal As Integer
    FragataCaos As Integer
    Barca As Integer
    Galera As Integer
    Galeon As Integer
    BarcaCiuda As Integer
    BarcaCiudaAtacable As Integer
    GaleraCiuda As Integer
    GaleraCiudaAtacable As Integer
    GaleonCiuda As Integer
    GaleonCiudaAtacable As Integer
    BarcaReal As Integer
    BarcaRealAtacable As Integer
    GaleraReal As Integer
    GaleraRealAtacable As Integer
    GaleonReal As Integer
    GaleonRealAtacable As Integer
    BarcaPk As Integer
    GaleraPk As Integer
    GaleonPk As Integer
    BarcaCaos As Integer
    GaleraCaos As Integer
    GaleonCaos As Integer
    NingunEscudo As Byte
    NingunCasco As Byte
    NingunArma As Byte
    CuerpoMuerto As Integer
    CabezaMuerto As Integer
    HumanoHPrimerCabeza As Integer
    HumanoHUltimaCabeza As Integer
    ElfoHPrimerCabeza As Integer
    ElfoHUltimaCabeza As Integer
    DrowHPrimerCabeza As Integer
    DrowHUltimaCabeza As Integer
    EnanoHPrimerCabeza As Integer
    EnanoHUltimaCabeza As Integer
    GnomoHPrimerCabeza As Integer
    GnomoHUltimaCabeza As Integer
    HumanoMPrimerCabeza As Integer
    HumanoMUltimaCabeza As Integer
    ElfoMPrimerCabeza As Integer
    ElfoMUltimaCabeza As Integer
    DrowMPrimerCabeza As Integer
    DrowMUltimaCabeza As Integer
    EnanoMPrimerCabeza As Integer
    EnanoMUltimaCabeza As Integer
    GnomoMPrimerCabeza As Integer
    GnomoMUltimaCabeza As Integer
End Type

Public ConstantesGRH As tConstantesGRH

Public Type tConstantesItems
    EspadaMataDragones As Integer
    LingoteHierro As Integer
    LingotePlata As Integer
    LingoteOro As Integer
    Leña As Integer
    LeñaElfica As Integer
    HachaLeñador As Integer
    HachaLeñaElfica As Integer
    PiqueteMinero As Integer
    HachaLeñadorNW As Integer
    PiqueteMineroNW As Integer
    CañaPescaNW As Integer
    SerruchoCarpinteroNW As Integer
    MartilloHerreroNW As Integer
    Daga As Integer
    FogataApagada As Integer
    Fogata As Integer
    FogataElfica As Integer
    RamitaElfica As Integer
    MartilloHerrero As Integer
    SerruchoCarpintero As Integer
    RedPesca As Integer
    CañaPesca As Integer
    Flecha As Integer
    Flecha1 As Integer
    Flecha2 As Integer
    Flecha3 As Integer
    FlechaNewbie As Integer
    Cuchillas As Integer
    Oro As Integer
    NumPescados As Integer
    Pescado1 As Integer
    Pescado2 As Integer
    Pescado3 As Integer
    Pescado4 As Integer
    Telep As Integer
    GuanteHurto As Integer
    RequiredGuildItem As Integer
    RequiredGuildItemCant As Integer
End Type

Public ConstantesItems As tConstantesItems

Public Type tConstantesHechizos
    Apocalipsis As Byte
    Descarga As Byte
    EleFuego As Byte
    EleAgua As Byte
    EleTierra As Byte
End Type

Public ConstantesHechizos As tConstantesHechizos

Public Type tConstantesCombate
    ProbAcuchillar As Byte
    DañoAcuchillar As Single
    AssassinNpcStabChance As Byte
End Type

Public ConstantesCombate As tConstantesCombate

Public Type tConstantesTrabajo
    EsfuerzoTalarGeneral As Byte
    EsfuerzoTalarLeñador As Byte
    EsfuerzoPescarGeneral As Byte
    EsfuerzoPescarPescador As Byte
    EsfuerzoExcavarGeneral As Byte
    EsfuerzoExcavarMinero As Byte
    PorcentajeMaterialesUpgrade As Single
End Type

Public ConstantesTrabajo As tConstantesTrabajo

Public Type tConstantesNPCs
    EleFuego As Integer
    EleTierra As Integer
    EleAgua As Integer
End Type

Public ConstantesNPCs As tConstantesNPCs

Public Type tConstantesReputacion
    Asalto As Integer
    Asesino As Integer
    AsesinoGuardiaBueno As Integer
    AsesinoNPCMalo As Integer
    AsesinoCiuda As Integer
    AtacoNPCMalo As Integer
    Cazador As Integer
    Noble As Integer
    Ladron As Integer
    RestarLadron As Integer
    Proleta As Integer
End Type

Public ConstantesReputacion As tConstantesReputacion

Public Type tConstantesSonidos
    Swing As Integer
    Talar As Integer
    Pescar As Integer
    Minero As Integer
    Warp As Integer
    Puerta As Integer
    Nivel As Integer
    UserMuere As Integer
    Impacto As Integer
    Impacto2 As Integer
    Leñador As Integer
    Fogata As Integer
    Ave As Integer
    Ave2 As Integer
    Ave3 As Integer
    Grillo As Integer
    Grillo2 As Integer
    SacarArma As Integer
    Escudo As Integer
    MartilloHerrero As Integer
    TrabajoCarpintero As Integer
    Tomar As Integer
End Type

Public ConstantesSonidos As tConstantesSonidos

Public Type tConstantesBosses
    BossDMCastPutrefaccion As Byte
    BossDMCastAparicion As Byte
    BossDMSpellPutrefaccion As Integer
    BossDMSpellAparicion As Integer
    BossDVDistance As Byte
    BossDVChangeTarget As Byte
    BossDVCastDescarga As Byte
    BossDVCastTormenta As Byte
    BossDVCastPetrificar As Byte
    BossDVSpellDescarga As Integer
    BossDVSpellTormenta As Integer
    BossDVSpellPetrificar As Integer
    BossDINumDebuff As Byte
    BossDISpellDebuff() As Integer
    BossDISpellBola As Integer
    BossDACastTorrente As Byte
    BossDACastTentaculo As Byte
    BossDACastAtraer As Byte
    BossDACastAplastar As Byte
    BossDAAplastarArea As Byte
    BossDASpellTorrente As Integer
    BossDASpellTentaculo As Integer
    BossDASpellAtraer As Integer
End Type

Public ConstantesBosses As tConstantesBosses

Public Enum ClientPresentEffects
    BossAppears = 0
    ENUMSIZE
End Enum


' Configuration

Public Type ExternalToolConfigurationPreperties
    Enabled As Boolean
    ListenPort As Long
    ExePath As String
End Type

Public Type ExternalToolsConfiguration
    StateServer As ExternalToolConfigurationPreperties
    ProxyServer As ExternalToolConfigurationPreperties
End Type

Public Type AccountSessionConfigType
    Lifetime As Long
    MaxQuantity As Integer
    TokenSize As Byte
End Type

Public Type ResourcesPathsConfigType
    Dats As String
    Maps As String
    WorldBackup As String
End Type

Public Type LogsPathsConfigType
    GeneralPath As String
    GameMastersPath As String
    DevelopmentPath As String
    GuildsPath As String
End Type

Public Type tServerIntervals
    SanaIntervaloSinDescansar As Long
    StaminaIntervaloSinDescansar As Long
    SanaIntervaloDescansar As Long
    StaminaIntervaloDescansar As Long
    IntervaloSed As Long
    IntervaloHambre As Long
    IntervaloVeneno As Long
    IntervaloPutrefaccionDmg As Long
    IntervaloParalizado As Long
    IntervaloParalizadoReducido As Long
    IntervaloNPCParalizado As Long
    IntervaloInvisible As Long
    IntervaloMimetismo As Long
    IntervaloFrio As Long
    IntervaloLava As Long
    IntervaloInvocacion As Long
    IntervaloOculto As Long
    IntervaloUserPuedeAtacar As Long
    IntervaloGolpeUsar As Long
    IntervaloMagiaGolpe As Long
    IntervaloGolpeMagia As Long
    IntervaloUserPuedeCastear As Long
    IntervaloUserPuedeTrabajar As Long
    IntervaloIdleKick As Long
    IntervaloCerrarConexion As Long
    IntervaloUserPuedeUsar As Long
    IntervaloUserPuedeUsarU As Long
    IntervaloFlechasCazadores As Long
    IntervaloPuedeSerAtacado As Long
    IntervaloAtacable As Long
    IntervaloOwnedNpc As Long
    IntervaloOcultar As Long
    IntervaloInmunidad As Long
    
    IntervalRequestPosition As Long
    IntervalMeditate As Long
    IntervalAction As Long
    IntervalWorkMacro As Long
    IntervalSpellMacro As Long
End Type

Public Type tPassiveSkillsConfiguration
    Name As String
    Enabled As Boolean
    UnlockLevel As Integer
    AllowedClassesQty As Integer
    AllowedClasses() As eClass
End Type

Public Type tServerConfigurationType
    ExternalTools As ExternalToolsConfiguration
    Session As AccountSessionConfigType
    ResourcesPaths As ResourcesPathsConfigType
    LogsPaths As LogsPathsConfigType
    LogToDebuggerWindow As Boolean
    UseExternalAccountValidation As Boolean
    
    IpTablesSecurityLogFailedEnabled As Boolean
    IpTablesSecurityEnabled As Boolean
        
    Intervals As tServerIntervals
    
    PassiveSkills() As tPassiveSkillsConfiguration
    PassiveSkillsQty As Long
    
    StartPositionsQty As Integer
    StartPositions() As WorldPos
    
End Type

Public ServerConfiguration As tServerConfigurationType

Public Enum eTypeTarget
    isUser = 1
    IsNPC
End Enum

Public Type eTypeClassStartingItem
    ItemNumber As Integer
    Quantity As Integer
    Equipped As Byte
End Type

Public Type eTypeClassConfiguration
    Name As String
    Enabled As Boolean
    
    StartingItemsQty As Integer
    StartingItems() As eTypeClassStartingItem
    
    StartingSpellsQty As Integer
    StartingSpells() As Integer
    
    ClassMods As ModClase 'We should move the old class mods here
    
    RaceMods() As tModClassRace
        
    MasteryGroupsQty As Integer
    MasteryGroups() As tMasteryBoostGroupConfig
End Type

Public Enum eSplashDamageType
    None = 0
    Swing
    Lance
    Pike
End Enum

Public Classes() As eTypeClassConfiguration

Public MasteriesQty As Integer
Public Masteries() As tMasteryBoost
Public Type PermissionType
    IdPermission As Long
    Key As String
    PermissionName As String
    IsEnabled As Boolean
End Type

Public PermissionConfig() As PermissionType

Public Type GuildRoleType
    IdRole As Long
    RoleName As String
    PermissionCount As Integer
    RolePermission() As PermissionType
    IsDeleteable As Boolean
    IsDirty As Boolean
End Type

Public Type GuildUpgradeEffectType
    IsChatOverHead As Boolean
    IsFriendlyFireProtection As Boolean
    AddMemberLimit As Integer
    AddRolesGuild As Byte
    IsGuildBank As Boolean
    AddBankSlot As Byte
    AddBankBox As Byte
    AddMaxContribution As Long
    IsSeeInvisibleGuildMember As Boolean
End Type

Public Type GuildUpgradeGroupConfig
    UpgradeQty As Integer
    Upgrades() As Integer
End Type

Public Type GuildQuestReq
    Id As Integer
    Title As String
    Obtained As Boolean
End Type

Public Type GuildConfUpgradeType
    IsEnabled As Boolean
    Name As String
    Description As String
    IconGraph As Integer
    GoldCost As Long
    ContributionCost As Long
    QuestRequired() As GuildQuestReq
    UpgradeRequired() As Integer
    UpgradeEffect As GuildUpgradeEffectType
End Type

Public Type tGuildReservedEmail
    AccountEmail As String
    GuildName As String
End Type

Public Type GuildConfigurationType
    MemberQty As Integer
    BankSlotQty As Integer
    BankBoxesQty As Integer
    MaxGold As Long
    MaxContribution As Long
    
    MaxGuilds As Integer
    RolsQty As Integer
    
    CreationEnabled As Boolean
    CreationLeaderRequiredLevel As Byte
    CreationRigthHandRequiredLevel As Byte
    CreationRequiredGold As Long
    
    UpgradesQty As Integer
    UpgradesGroupsQty As Integer
    
    GuildUpgradeGroup() As GuildUpgradeGroupConfig
    GuildUpgradesList() As GuildConfUpgradeType
    InvitationLifeTimeInMinutes As Integer
    
    InvalidNamesQty As Long
    InvalidNames() As String
    
    ReservedNamesQty As Integer
    ReservedNames() As tGuildReservedEmail
End Type

Public GuildConfiguration As GuildConfigurationType

Public Type GuildBankType
    Box As Integer
    Slot As Integer
    IdObject As Long
    Amount As Integer
End Type

Public Type GuildMemberType
    IdUser As Long
    NameUser As String
    JoinDate As Date
    IdRole As Long
    RoleIndex As Integer
    RoleAssignedBy As Long
    ContributionEarner As Long
    IsDirty As Boolean
End Type

Public Type GuildQuestCompletedType
    IdContribution As Long
    IdQuest As Long
    CompletedDate As Date
    MembersContributed As Long
    ContributionGained As Long
End Type

Public Type GuildUpgradeType
    IdUpgrade As Integer
    UpgradeLevel As Integer
    UpgradeDate As Date
    UpgradeBy As Long
    IsEnabled As Boolean
End Type

Public Type GuildMembersOnline
    IdUser As Long
    MemberUserIndex As Integer
    IdRole As Long
End Type

Public Enum eGuildAlignment
    Neutral = 1
    Real
    Evil
    GameMaster
    
    LastElement
End Enum

Public Type tCurrentQuest
    IdQuest As Integer
    CurrentStage As Integer
    StageIsCompleted As Boolean
    IsFirstTime As Boolean
    CanStartNextStage As Boolean
    StartedDate As Date
    SecondsLeft As Long
    
    ServerStartedDate As Date
    
    CurrentFrags As tQuestFrags
    CurrentNpcKillsQuantity As Integer
    
    CurrentNpcKills() As tQuestNpc
    CurrentObjectList As RequiredObjectList
    
End Type

Public Type GuildInvitation
    InvitedByUserIndex As Long
    InvitedByUserId As Long

    TargetUserIndex As Long
    TargetUserId As Long
    InvitationDate As Date
End Type

Public Type GuildType
    'guild info
    IdGuild As Long
    Name As String
    Description As String
    Alignment As Integer
    CreationTime As Date
    Status As Integer
    IdLeader As Long
    IdRightHand As Long
    MemberCount As Integer
    ContributionEarned As Long
    ContributionAvailable As Long
    
    RankingPoints As Long
    
    ' guild quest
    QuestCompletedCount As Integer
    QuestCompleted() As GuildQuestCompletedType
    CurrentQuest As tCurrentQuest
    
    BankGold As Long
    IdDefaultRole As Long
    
    IsDirty As Boolean ' this field is used for save

    Bank() As GuildBankType
    Members() As GuildMemberType
    Roles() As GuildRoleType
    Upgrades() As GuildUpgradeType
    UpgradeEffect As GuildUpgradeEffectType
    
    OnlineMembers() As GuildMembersOnline
    OnlineMemberCount As Integer
    
    ListeningAdmins() As Integer
    ListeningAdminsCount As Integer
    
    Invitations() As GuildInvitation
End Type

Public MaxGuildQty As Integer
Public GuildList() As GuildType

Public Const ID_ROLE_LEADER As Integer = 1
Public Const ID_ROLE_RIGHTHAND As Integer = 2

Public Enum EGuildPermission
    EDIT_GUILD_DESC = 1
    RIGHT_HAND_ASSIGN
    ROLE_ASSIGN
    ROLE_CREATE_DELETE
    ROLE_MODIFY
    BANK_DEPOSIT_ITEM
    BANK_WITHDRAW_ITEM
    BANK_DEPOSIT_GOLD
    BANK_WITHDRAW_GOLD
    MEMBER_ACCEPT
    MEMBER_KICK
End Enum

Public Enum eChangeMember
    OnlineChange = 1
    RoleChange
    GoldGBChange
End Enum

Public Enum eWorkerStoreAction
    WorkerStoreGetRecipes
    WorkerStoreCreate
    WorkerStoreClose
    WorkerStoreCraftItem
End Enum

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

Public Const MAX_LENGTH_NAME As Integer = 15

Public Const PUNISHMENT_TYPE_RECORD As Byte = 45

'''' QPT'No se usa

Public Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
            
Public Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long


Public Function CalculateExecutionTime(ByVal started As Currency, ByVal ended As Currency, ByVal frequency As Currency) As Long
    CalculateExecutionTime = (ended - started) * 1000 / frequency
End Function


'''' END QPT

