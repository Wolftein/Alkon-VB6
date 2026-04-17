Attribute VB_Name = "General"
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

Public Running As Boolean

Public Enum ShutDownFrom
    Nobody
    System
    User
End Enum

Public ShutdownWithBackup As Boolean
Public ShutdownBy As ShutDownFrom

Sub DarCuerpoDesnudo(ByVal UserIndex As Integer, Optional ByVal Mimetizado As Boolean = False)
'***************************************************
'Autor: Nacho (Integer)
'Last Modification: 03/14/07
'Da cuerpo desnudo a un usuario
'23/11/2009: ZaMa - Optimizacion de codigo.
'***************************************************
On Error GoTo ErrHandler
  

    Dim CuerpoDesnudo As Integer
    
    With UserList(UserIndex)
        Select Case .Genero
            Case eGenero.Hombre
                Select Case .raza
                    Case eRaza.Humano
                        CuerpoDesnudo = 21
                    Case eRaza.Drow
                        CuerpoDesnudo = 32
                    Case eRaza.Elfo
                        CuerpoDesnudo = 210
                    Case eRaza.Gnomo
                        CuerpoDesnudo = 222
                    Case eRaza.Enano
                        CuerpoDesnudo = 53
                End Select
            Case eGenero.Mujer
                Select Case .raza
                    Case eRaza.Humano
                        CuerpoDesnudo = 39
                    Case eRaza.Drow
                        CuerpoDesnudo = 40
                    Case eRaza.Elfo
                        CuerpoDesnudo = 259
                    Case eRaza.Gnomo
                        CuerpoDesnudo = 260
                    Case eRaza.Enano
                        CuerpoDesnudo = 60
                End Select
        End Select
        
        If Mimetizado Then
            .OrigChar.body = CuerpoDesnudo
        Else
            .Char.body = CuerpoDesnudo
        End If
        
        .flags.Desnudo = 1
    End With

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub DarCuerpoDesnudo de General.bas")
End Sub

Sub Bloquear(ByVal toMap As Boolean, ByVal sndIndex As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal b As Boolean)
'***************************************************
'Author: Unknown
'Last Modification: -
'b ahora es boolean,
'b=true bloquea el tile en (x,y)
'b=false desbloquea el tile en (x,y)
'toMap = true -> Envia los datos a todo el mapa
'toMap = false -> Envia los datos al user
'Unifique los tres parametros (sndIndex,sndMap y map) en sndIndex... pero de todas formas, el mapa jamas se indica.. eso esta bien asi?
'Puede llegar a ser, que se quiera mandar el mapa, habria que agregar un nuevo parametro y modificar.. lo quite porque no se usaba ni aca ni en el cliente :s
'***************************************************
On Error GoTo ErrHandler
  

    If toMap Then
        Call SendData(SendTarget.toMap, sndIndex, PrepareMessageBlockPosition(X, Y, b))
    Else
        Call WriteBlockPosition(sndIndex, X, Y, b)
    End If

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Bloquear de General.bas")
End Sub

Function HayAgua(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    If Map > 0 And Map < NumMaps + 1 And X > 0 And X < 101 And Y > 0 And Y < 101 Then
        With MapData(Map, X, Y)
            If ((.Graphic(1) >= 1505 And .Graphic(1) <= 1520) Or _
            (.Graphic(1) >= 5665 And .Graphic(1) <= 5680) Or _
            (.Graphic(1) >= 18974 And .Graphic(1) <= 18989) Or _
            (.Graphic(1) >= 13547 And .Graphic(1) <= 13562)) And _
               .Graphic(2) = 0 Then
                    HayAgua = True
            Else
                    HayAgua = False
            End If
        End With
    Else
      HayAgua = False
    End If

  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function HayAgua de General.bas")
End Function

Private Function HayLava(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
'***************************************************
'Autor: Nacho (Integer)
'Last Modification: 03/12/07
'***************************************************
On Error GoTo ErrHandler
  
    If Map > 0 And Map < NumMaps + 1 And X > 0 And X < 101 And Y > 0 And Y < 101 Then
        If MapData(Map, X, Y).Graphic(1) >= 5837 And MapData(Map, X, Y).Graphic(1) <= 5852 Then
            HayLava = True
        Else
            HayLava = False
        End If
    Else
      HayLava = False
    End If

  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function HayLava de General.bas")
End Function

Public Function IsPortal(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
    Dim ObjIndex As Integer
    
    ObjIndex = MapData(Map, X, Y).ObjInfo.ObjIndex
    
    If ObjIndex <> 0 Then
        IsPortal = ObjData(ObjIndex).ObjType = eOBJType.otTeleport
    End If
    
End Function


Sub LimpiarMundo()
'***************************************************
'Author: Unknow
'Last Modification: 04/15/2008
'01/14/2008: Marcos Martinez (ByVal) - La funcion FOR estaba mal. En ves de i habia un 1.
'04/15/2008: (NicoNZ) - La funcion FOR estaba mal, de la forma que se hacia tiraba error.
'***************************************************
On Error GoTo ErrHandler

    Dim trashId As Integer
    Dim d As cGarbage
    Set d = New cGarbage
    
    For trashId = TrashCollector.Count To 1 Step -1
        Set d = TrashCollector(trashId)
        Call EraseObj(1, d.Map, d.X, d.Y)
        Call TrashCollector.Remove(trashId)
        Set d = Nothing
    Next trashId
    
    Call SecurityIp.IpTableSecurityCleanIpTime
    
    Exit Sub

ErrHandler:
    Call LogError("Error producido en el sub LimpiarMundo: " & Err.Description)
End Sub

Sub EnviarSpawnList(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    Dim k As Long
    Dim npcNames() As String
    
    ReDim npcNames(1 To UBound(Declaraciones.SpawnList)) As String
    
    For k = 1 To UBound(Declaraciones.SpawnList)
        npcNames(k) = Declaraciones.SpawnList(k).NpcName
    Next k
    
    Call WriteSpawnList(UserIndex, npcNames())

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub EnviarSpawnList de General.bas")
End Sub

Sub Main()
'***************************************************
'Author: Unknown
'Last Modification: 15/03/2011
'15/03/2011: ZaMa - Modularice todo, para que quede mas claro.
'***************************************************
On Error GoTo ErrHandler

On Error Resume Next

    ChDir App.Path
    ChDrive App.Path
        
    ' Initialize Aurora
    Dim Aurora_Configuration As Aurora_Engine.Kernel_Properties
    Aurora_Configuration.LogFilename = App.Path & "/Aurora.log"
    Call Kernel.Initialize(eKernelModeServer, Aurora_Configuration)
    
    Set Aurora_Network = Kernel.Network
        
    ' Initialize the Rnd seed
    Call Randomize
    
     ' Server.ini & Apuestas.dat
    frmCargando.Label1(2).Caption = "Cargando Server.ini"
    Call LoadSini
    
    If Not AreResourcesPathsSet Then
        Call InitResourcesPaths
        Exit Sub
    End If
    
    Call LoadMotd
    Call BanIpCargar
    
    frmMain.Caption = frmMain.Caption & " V." & App.Major & "." & App.Minor & "." & App.Revision
    
    ' Start loading..
    frmCargando.Show
    
    ' Constants & vars
    frmCargando.Label1(2).Caption = "Cargando constantes..."
    Call LoadConstants
    DoEvents
    
    ' Balance.dat
    frmCargando.Label1(2).Caption = "Cargando Balance.Dat"
    Call LoadBalance
    DoEvents
    
    ' DB
    frmCargando.Label1(2).Caption = "Conectando a db..."
    If Not ConnectDB Then End
    mysql_requery_time = 20
    
    DoEvents
    
    frmCargando.Label1(2).Caption = "Cargando lista de penas..."
    Call GetPunishmentTypeCount
    
    ' Arrays
    frmCargando.Label1(2).Caption = "Iniciando Arrays..."
    Call LoadArrays
    
    ' Gameserver intervals
    Call LoadIntervals
    
    Call CargaApuestas
    
        ' Masteries
    frmCargando.Label1(2).Caption = "Cargando Masteries.dat"
    Call LoadMasteries
    
    ' Classes.dat
    frmCargando.Label1(2).Caption = "Cargando Classes.dat"
    Call LoadClasses
    
    'Professions.dat
    frmCargando.Label1(2).Caption = "Cargando Professions.Dat"
    Call LoadProfessions
    
    'Resources.dat
    frmCargando.Label1(2).Caption = "Cargando Resources.Dat"
    Call LoadResources
    
    ' Hechizos.dat
    frmCargando.Label1(2).Caption = "Cargando Hechizos.Dat"
    Call CargarHechizos

    ' Obj.dat
    frmCargando.Label1(2).Caption = "Cargando Obj.Dat"
    Call LoadOBJData
    
    frmCargando.Label1(2).Caption = "Cargando Passive Skills.Dat"
    Call LoadPassiveSkillsConfig
    
    ' Npcs.dat
    frmCargando.Label1(2).Caption = "Cargando NPCs.Dat"
    Call CargaNpcsDat
    
    ' Quests
    frmCargando.Label1(2).Caption = "Iniciando lista de quests..."
    Call LoadGuildQuests
        
    ' Armaduras faccionarias
    frmCargando.Label1(2).Caption = "Cargando ArmadurasFaccionarias.dat"
    Call LoadArmadurasFaccion
    
    ' Animaciones
    frmCargando.Label1(2).Caption = "Cargando Animaciones"
    Call LoadAnimations
    
    ' Pretorianos
    frmCargando.Label1(2).Caption = "Cargando Pretorianos.dat"
    Call LoadPretorianData

    ' Duelos
    frmCargando.Label1(2).Caption = "Cargando Duelos.dat"
    Call LoadDuelData
    
    ' Drops
    frmCargando.Label1(2).Caption = "Cargando Drops.dat"
    Call LoadDropData
    
    frmCargando.Label1(2).Caption = "Cargando Guild.dat"
    Call LoadGuildConfiguration
    
    frmCargando.Label1(2).Caption = "Cargando Guild Permissions desde la DB"
    Call LoadGuildPermissionDB
    
    frmCargando.Label1(2).Caption = "Cargando Guilds desde la DB"
    Call LoadGuilds
    
    ' Mapas
    If BootDelBackUp Then
        frmCargando.Label1(2).Caption = "Cargando BackUp"
        Call CargarBackUp
    Else
        frmCargando.Label1(2).Caption = "Cargando Mapas"
        Call LoadMapData
    End If
    
    ' Bosses.dat
    frmCargando.Label1(2).Caption = "Cargando Bosses.Dat"
    Call LoadBossData
    
    OUTDATED_VERSION = "Esta versión del juego es obsoleta, la versión correcta es la " & _
                       ULTIMAVERSION & ". La misma se encuentra disponible en www.alkononline.com.ar"
    
    ' Map Sounds
    Set SonidosMapas = New SoundMapInfo
    Call SonidosMapas.LoadSoundMapInfo
    
    'Init the protocol
#If EnableSecurity Then
    Call ProtocolPackets.InitProtocol
#Else
    Call Protocol.InitProtocol
#End If

    ' Connections
    Call ResetUsersConnections
    
    ' Timers
    Call InitMainTimers
    
    ' Redim Arrays
    Call InitUserArrays
    
    ' Boveda de cuenta
    Call InitAccBank
    
    ' Sockets
    Call SocketConfig

    'Init the protocol
#If EnableSecurity Then
    Call Security.StartSecurity
#End If


    ' Start the listening sockets for external tools
    Call frmMain.ListenMQ
    Call frmMain.ListenProxySender
    
    Call frmMain.ListenForRemoteTools
    
    ' End loading..
    Unload frmCargando
    
    'Log start time
    LogServerStartTime
    
    'Ocultar
    If HideMe = 1 Then
        Call frmMain.InitMain(1)
    Else
        Call frmMain.InitMain(0)
    End If
    
    tInicioServer = GetTickCount()
    Call InicializaEstadisticas
    
    'Spawn default pretorian clan
    If Not ClanPretoriano(1).Active Then
        Call ClanPretoriano(1).SpawnClan(163, 35, 25, 1)
    End If
    
    ' Account Session System
    Call InitSessionSystem
            
    Running = True
    
    While (Running)
        DoEvents
        
        Call Kernel.Tick
    Wend

    Call CloseServer
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Main de General.bas")
End Sub

Private Sub LoadConstants()
'*****************************************************************
'Author: ZaMa
'Last Modify Date: 15/03/2011
'Loads all constants and general parameters.
'*****************************************************************
On Error GoTo ErrHandler
  
On Error Resume Next

    Dim Leer As clsIniManager
    Set Leer = New clsIniManager
   
    LastBackup = Format(Now, "Short Time")
    Minutos = Format(Now, "Short Time")
    
    ' Paths
    CharPath = App.Path & "\Charfile\"
    
    ' Skills by level
    LevelSkill(1).LevelValue = 3
    LevelSkill(2).LevelValue = 5
    LevelSkill(3).LevelValue = 7
    LevelSkill(4).LevelValue = 10
    LevelSkill(5).LevelValue = 13
    LevelSkill(6).LevelValue = 15
    LevelSkill(7).LevelValue = 17
    LevelSkill(8).LevelValue = 20
    LevelSkill(9).LevelValue = 23
    LevelSkill(10).LevelValue = 25
    LevelSkill(11).LevelValue = 27
    LevelSkill(12).LevelValue = 30
    LevelSkill(13).LevelValue = 33
    LevelSkill(14).LevelValue = 35
    LevelSkill(15).LevelValue = 37
    LevelSkill(16).LevelValue = 40
    LevelSkill(17).LevelValue = 43
    LevelSkill(18).LevelValue = 45
    LevelSkill(19).LevelValue = 47
    LevelSkill(20).LevelValue = 50
    LevelSkill(21).LevelValue = 53
    LevelSkill(22).LevelValue = 55
    LevelSkill(23).LevelValue = 57
    LevelSkill(24).LevelValue = 60
    LevelSkill(25).LevelValue = 63
    LevelSkill(26).LevelValue = 65
    LevelSkill(27).LevelValue = 67
    LevelSkill(28).LevelValue = 70
    LevelSkill(29).LevelValue = 73
    LevelSkill(30).LevelValue = 75
    LevelSkill(31).LevelValue = 77
    LevelSkill(32).LevelValue = 80
    LevelSkill(33).LevelValue = 83
    LevelSkill(34).LevelValue = 85
    LevelSkill(35).LevelValue = 87
    LevelSkill(36).LevelValue = 90
    LevelSkill(37).LevelValue = 93
    LevelSkill(38).LevelValue = 95
    LevelSkill(39).LevelValue = 97
    LevelSkill(40).LevelValue = 100
    LevelSkill(41).LevelValue = 100
    LevelSkill(42).LevelValue = 100
    LevelSkill(43).LevelValue = 100
    LevelSkill(44).LevelValue = 100
    LevelSkill(45).LevelValue = 100
    LevelSkill(46).LevelValue = 100
    LevelSkill(47).LevelValue = 100
    LevelSkill(48).LevelValue = 100
    LevelSkill(49).LevelValue = 100
    LevelSkill(50).LevelValue = 100
    
    ' Races
    ListaRazas(eRaza.Humano) = "Humano"
    ListaRazas(eRaza.Elfo) = "Elfo"
    ListaRazas(eRaza.Drow) = "Drow"
    ListaRazas(eRaza.Gnomo) = "Gnomo"
    ListaRazas(eRaza.Enano) = "Enano"
    
    ' Classes
    ListaClases(eClass.Mage) = "Mago"
    ListaClases(eClass.Cleric) = "Clerigo"
    ListaClases(eClass.Warrior) = "Guerrero"
    ListaClases(eClass.Assasin) = "Asesino"
    ListaClases(eClass.Thief) = "Ladron"
    ListaClases(eClass.Bard) = "Bardo"
    ListaClases(eClass.Druid) = "Druida"
    ListaClases(eClass.Bandit) = "Bandido"
    ListaClases(eClass.Paladin) = "Paladin"
    ListaClases(eClass.Worker) = "Trabajador"
    ListaClases(eClass.Hunter) = "Cazador"
    
    ' Skills
    SkillsNames(eSkill.Magia) = "Magia"
    SkillsNames(eSkill.Robar) = "Robar"
    SkillsNames(eSkill.Tacticas) = "Evasión en combate"
    SkillsNames(eSkill.Armas) = "Combate con armas"
    SkillsNames(eSkill.Meditar) = "Meditar"
    SkillsNames(eSkill.Apuñalar) = "Apuñalar"
    SkillsNames(eSkill.Ocultarse) = "Ocultarse"
    SkillsNames(eSkill.Supervivencia) = "Supervivencia"
    SkillsNames(eSkill.Talar) = "Talar"
    SkillsNames(eSkill.Defensa) = "Defensa con escudos"
    SkillsNames(eSkill.Pesca) = "Pesca"
    SkillsNames(eSkill.Mineria) = "Mineria"
    SkillsNames(eSkill.Carpinteria) = "Carpinteria"
    SkillsNames(eSkill.Herreria) = "Herreria"
    SkillsNames(eSkill.Domar) = "Domar animales"
    SkillsNames(eSkill.Proyectiles) = "Combate a distancia"
    SkillsNames(eSkill.Wrestling) = "Combate sin armas"
    SkillsNames(eSkill.Sastreria) = "Sastreria"
    
    ' Attributes
    ListaAtributos(eAtributos.Fuerza) = "Fuerza"
    ListaAtributos(eAtributos.Agilidad) = "Agilidad"
    ListaAtributos(eAtributos.Inteligencia) = "Inteligencia"
    ListaAtributos(eAtributos.Carisma) = "Carisma"
    ListaAtributos(eAtributos.Constitucion) = "Constitucion"
    
    Call Leer.Initialize(DatPath & "Constantes.dat")
    
    With Constantes
        .GoHomePenalty = Val(Leer.GetValue("Misc", "GoHomePenalty"))
        .MaxPrivateMessages = Val(Leer.GetValue("Misc", "MaxPrivateMessages"))
        .TiempoCarcelPiquete = Val(Leer.GetValue("Misc", "TiempoCarcelPiquete"))
        .MaxNpcDrops = Val(Leer.GetValue("Misc", "MaxNpcDrops"))
        .MaxDenuncias = Val(Leer.GetValue("Misc", "MaxDenuncias"))
        .MaxPartyMembers = Val(Leer.GetValue("Misc", "MaxPartyMembers"))
        .MaxPartyDifLevel = Val(Leer.GetValue("Misc", "MaxPartyDifLevel"))
        .MaxMensajesForo = Val(Leer.GetValue("Misc", "MaxMensajesForo"))
        .MaxAnunciosForo = Val(Leer.GetValue("Misc", "MaxAnunciosForo"))
        .GuildLevelMax = Val(Leer.GetValue("Misc", "GuildLevelMax"))
        .MinGuildMembers = Val(Leer.GetValue("Misc", "MinGuildMembers"))
        .MaxGuildMembers = Val(Leer.GetValue("Misc", "MaxGuildMembers"))
        .ElfCampfireDuration = Val(Leer.GetValue("Misc", "ElfCampfireDuration"))
        .NormalCampfireDuration = Val(Leer.GetValue("Misc", "NormalCampfireDuration"))
        .MinStaRecoveryPerc = Val(Leer.GetValue("Misc", "MinStaRecoveryPerc"))
        .MaxStaRecoveryPerc = Val(Leer.GetValue("Misc", "MaxStaRecoveryPerc"))
        .CraftingStoreMap = Val(Leer.GetValue("Misc", "CraftingStoreMap"))
        .CraftingStoreOverheadIcon = Val(Leer.GetValue("Misc", "CraftingStoreOverheadIcon"))
    End With
        
    With ConstantesFX
        .FxSangre = Val(Leer.GetValue("FX", "Sangre"))
        .FxWarp = Val(Leer.GetValue("FX", "Warp"))
        .FxMeditarGrande = Val(Leer.GetValue("FX", "MeditarGrande"))
    End With
    
    With ConstantesGRH
        .FragataFantasmal = Val(Leer.GetValue("GRH", "FragataFantasmal"))
        .FragataReal = Val(Leer.GetValue("GRH", "FragataReal"))
        .FragataCaos = Val(Leer.GetValue("GRH", "FragataCaos"))
        .Barca = Val(Leer.GetValue("GRH", "Barca"))
        .Galera = Val(Leer.GetValue("GRH", "Galera"))
        .Galeon = Val(Leer.GetValue("GRH", "Galeon"))
        .BarcaCiuda = Val(Leer.GetValue("GRH", "BarcaCiuda"))
        .BarcaCiudaAtacable = Val(Leer.GetValue("GRH", "BarcaCiudaAtacable"))
        .GaleraCiuda = Val(Leer.GetValue("GRH", "GaleraCiuda"))
        .GaleraCiudaAtacable = Val(Leer.GetValue("GRH", "GaleraCiudaAtacable"))
        .GaleonCiuda = Val(Leer.GetValue("GRH", "GaleonCiuda"))
        .GaleonCiudaAtacable = Val(Leer.GetValue("GRH", "GaleonCiudaAtacable"))
        .BarcaReal = Val(Leer.GetValue("GRH", "BarcaReal"))
        .BarcaRealAtacable = Val(Leer.GetValue("GRH", "BarcaRealAtacable"))
        .GaleraReal = Val(Leer.GetValue("GRH", "GaleraReal"))
        .GaleraRealAtacable = Val(Leer.GetValue("GRH", "GaleraRealAtacable"))
        .GaleonReal = Val(Leer.GetValue("GRH", "GaleonReal"))
        .GaleonRealAtacable = Val(Leer.GetValue("GRH", "GaleonRealAtacable"))
        .BarcaPk = Val(Leer.GetValue("GRH", "BarcaPk"))
        .GaleraPk = Val(Leer.GetValue("GRH", "GaleraPk"))
        .GaleonPk = Val(Leer.GetValue("GRH", "GaleonPk"))
        .BarcaCaos = Val(Leer.GetValue("GRH", "BarcaCaos"))
        .GaleraCaos = Val(Leer.GetValue("GRH", "GaleraCaos"))
        .GaleonCaos = Val(Leer.GetValue("GRH", "GaleonCaos"))
        .NingunEscudo = Val(Leer.GetValue("GRH", "NingunEscudo"))
        .NingunCasco = Val(Leer.GetValue("GRH", "NingunCasco"))
        .NingunArma = Val(Leer.GetValue("GRH", "NingunArma"))
        .CuerpoMuerto = Val(Leer.GetValue("GRH", "CuerpoMuerto"))
        .CabezaMuerto = Val(Leer.GetValue("GRH", "CabezaMuerto"))
        .HumanoHPrimerCabeza = Val(Leer.GetValue("GRH", "HumanoHPrimerCabeza"))
        .HumanoHUltimaCabeza = Val(Leer.GetValue("GRH", "HumanoHUltimaCabeza"))
        .ElfoHPrimerCabeza = Val(Leer.GetValue("GRH", "ElfoHPrimerCabeza"))
        .ElfoHUltimaCabeza = Val(Leer.GetValue("GRH", "ElfoHUltimaCabeza"))
        .DrowHPrimerCabeza = Val(Leer.GetValue("GRH", "DrowHPrimerCabeza"))
        .DrowHUltimaCabeza = Val(Leer.GetValue("GRH", "DrowHUltimaCabeza"))
        .EnanoHPrimerCabeza = Val(Leer.GetValue("GRH", "EnanoHPrimerCabeza"))
        .EnanoHUltimaCabeza = Val(Leer.GetValue("GRH", "EnanoHUltimaCabeza"))
        .GnomoHPrimerCabeza = Val(Leer.GetValue("GRH", "GnomoHPrimerCabeza"))
        .GnomoHUltimaCabeza = Val(Leer.GetValue("GRH", "GnomoHUltimaCabeza"))
        .HumanoMPrimerCabeza = Val(Leer.GetValue("GRH", "HumanoMPrimerCabeza"))
        .HumanoMUltimaCabeza = Val(Leer.GetValue("GRH", "HumanoMUltimaCabeza"))
        .ElfoMPrimerCabeza = Val(Leer.GetValue("GRH", "ElfoMPrimerCabeza"))
        .ElfoMUltimaCabeza = Val(Leer.GetValue("GRH", "ElfoMUltimaCabeza"))
        .DrowMPrimerCabeza = Val(Leer.GetValue("GRH", "DrowMPrimerCabeza"))
        .DrowMUltimaCabeza = Val(Leer.GetValue("GRH", "DrowMUltimaCabeza"))
        .EnanoMPrimerCabeza = Val(Leer.GetValue("GRH", "EnanoMPrimerCabeza"))
        .EnanoMUltimaCabeza = Val(Leer.GetValue("GRH", "EnanoMUltimaCabeza"))
        .GnomoMPrimerCabeza = Val(Leer.GetValue("GRH", "GnomoMPrimerCabeza"))
        .GnomoMUltimaCabeza = Val(Leer.GetValue("GRH", "GnomoMUltimaCabeza"))
    End With
    
    With ConstantesItems
        .EspadaMataDragones = Val(Leer.GetValue("Items", "EspadaMataDragones"))
        .LingoteHierro = Val(Leer.GetValue("Items", "LingoteHierro"))
        .LingotePlata = Val(Leer.GetValue("Items", "LingotePlata"))
        .LingoteOro = Val(Leer.GetValue("Items", "LingoteOro"))
        .Leña = Val(Leer.GetValue("Items", "Leña"))
        .LeñaElfica = Val(Leer.GetValue("Items", "LeñaElfica"))
        .HachaLeñador = Val(Leer.GetValue("Items", "HachaLeñador"))
        .HachaLeñaElfica = Val(Leer.GetValue("Items", "HachaLeñaElfica"))
        .PiqueteMinero = Val(Leer.GetValue("Items", "PiqueteMinero"))
        .HachaLeñadorNW = Val(Leer.GetValue("Items", "HachaLeñadorNW"))
        .PiqueteMineroNW = Val(Leer.GetValue("Items", "PiqueteMineroNW"))
        .CañaPescaNW = Val(Leer.GetValue("Items", "CañaPescaNW"))
        .SerruchoCarpinteroNW = Val(Leer.GetValue("Items", "SerruchoCarpinteroNW"))
        .MartilloHerreroNW = Val(Leer.GetValue("Items", "MartilloHerreroNW"))
        .Daga = Val(Leer.GetValue("Items", "Daga"))
        .FogataApagada = Val(Leer.GetValue("Items", "FogataApagada"))
        .Fogata = Val(Leer.GetValue("Items", "Fogata"))
        .FogataElfica = Val(Leer.GetValue("Items", "FogataElfica"))
        .RamitaElfica = Val(Leer.GetValue("Items", "RamitaElfica"))
        .MartilloHerrero = Val(Leer.GetValue("Items", "MartilloHerrero"))
        .SerruchoCarpintero = Val(Leer.GetValue("Items", "SerruchoCarpintero"))
        .RedPesca = Val(Leer.GetValue("Items", "RedPesca"))
        .CañaPesca = Val(Leer.GetValue("Items", "CañaPesca"))
        .Flecha = Val(Leer.GetValue("Items", "Flecha"))
        .Flecha1 = Val(Leer.GetValue("Items", "Flecha1"))
        .Flecha2 = Val(Leer.GetValue("Items", "Flecha2"))
        .Flecha3 = Val(Leer.GetValue("Items", "Flecha3"))
        .FlechaNewbie = Val(Leer.GetValue("Items", "FlechaNewbie"))
        .Cuchillas = Val(Leer.GetValue("Items", "Cuchillas"))
        .Oro = Val(Leer.GetValue("Items", "Oro"))
        .NumPescados = Val(Leer.GetValue("Items", "NumPescados"))
        .Pescado1 = Val(Leer.GetValue("Items", "Pescado1"))
        .Pescado2 = Val(Leer.GetValue("Items", "Pescado2"))
        .Pescado3 = Val(Leer.GetValue("Items", "Pescado3"))
        .Pescado4 = Val(Leer.GetValue("Items", "Pescado4"))
        .Telep = Val(Leer.GetValue("Items", "Telep"))
        .GuanteHurto = Val(Leer.GetValue("Items", "GuanteHurto"))
    End With
    
    With ConstantesHechizos
        .Apocalipsis = Val(Leer.GetValue("Hechizos", "Apocalipsis"))
        .Descarga = Val(Leer.GetValue("Hechizos", "Descarga"))
        .EleFuego = Val(Leer.GetValue("Hechizos", "EleFuego"))
        .EleTierra = Val(Leer.GetValue("Hechizos", "EleTierra"))
        .EleAgua = Val(Leer.GetValue("Hechizos", "EleAgua"))
    End With
    
    With ConstantesCombate
        .ProbAcuchillar = Val(Leer.GetValue("Combate", "ProbAcuchillar"))
        .DañoAcuchillar = Val(Leer.GetValue("Combate", "DañoAcuchillar"))
        .AssassinNpcStabChance = Val(Leer.GetValue("Combate", "AssassinNpcStabChance"))
    End With
    
    With ConstantesTrabajo
        .EsfuerzoTalarGeneral = Val(Leer.GetValue("Trabajo", "EsfuerzoTalarGeneral"))
        .EsfuerzoTalarLeñador = Val(Leer.GetValue("Trabajo", "EsfuerzoTalarLeñador"))
        .EsfuerzoPescarGeneral = Val(Leer.GetValue("Trabajo", "EsfuerzoPescarGeneral"))
        .EsfuerzoPescarPescador = Val(Leer.GetValue("Trabajo", "EsfuerzoPescarPescador"))
        .EsfuerzoExcavarGeneral = Val(Leer.GetValue("Trabajo", "EsfuerzoExcavarGeneral"))
        .EsfuerzoExcavarMinero = Val(Leer.GetValue("Trabajo", "EsfuerzoExcavarMinero"))
        .PorcentajeMaterialesUpgrade = Val(Leer.GetValue("Trabajo", "PorcentajeMaterialesUpgrade"))
    End With
    
    With ConstantesNPCs
        .EleFuego = Val(Leer.GetValue("NPCs", "EleFuego"))
        .EleTierra = Val(Leer.GetValue("NPCs", "EleTierra"))
        .EleAgua = Val(Leer.GetValue("NPCs", "EleAgua"))
    End With
    
    With ConstantesReputacion
        .Asalto = Val(Leer.GetValue("Reputacion", "Asalto"))
        .Asesino = Val(Leer.GetValue("Reputacion", "Asesino"))
        .AsesinoGuardiaBueno = Val(Leer.GetValue("Reputacion", "AsesinoGuardiaBueno"))
        .AsesinoNPCMalo = Val(Leer.GetValue("Reputacion", "AsesinoNPCMalo"))
        .AsesinoCiuda = Val(Leer.GetValue("Reputacion", "AsesinoCiuda"))
        .AtacoNPCMalo = Val(Leer.GetValue("Reputacion", "AtacoNPCMalo"))
        .Cazador = Val(Leer.GetValue("Reputacion", "Cazador"))
        .Noble = Val(Leer.GetValue("Reputacion", "Noble"))
        .Ladron = Val(Leer.GetValue("Reputacion", "Ladron"))
        .RestarLadron = Val(Leer.GetValue("Reputacion", "RestarLadron"))
        .Proleta = Val(Leer.GetValue("Reputacion", "Proleta"))
    End With
    
    With ConstantesSonidos
        .Swing = Val(Leer.GetValue("Sonidos", "Swing"))
        .Talar = Val(Leer.GetValue("Sonidos", "Talar"))
        .Pescar = Val(Leer.GetValue("Sonidos", "Pescar"))
        .Minero = Val(Leer.GetValue("Sonidos", "Minero"))
        .Warp = Val(Leer.GetValue("Sonidos", "Warp"))
        .Puerta = Val(Leer.GetValue("Sonidos", "Puerta"))
        .Nivel = Val(Leer.GetValue("Sonidos", "Nivel"))
        .UserMuere = Val(Leer.GetValue("Sonidos", "UserMuere"))
        .Impacto = Val(Leer.GetValue("Sonidos", "Impacto"))
        .Impacto2 = Val(Leer.GetValue("Sonidos", "Impacto2"))
        .Leñador = Val(Leer.GetValue("Sonidos", "Leñador"))
        .Fogata = Val(Leer.GetValue("Sonidos", "Fogata"))
        .Ave = Val(Leer.GetValue("Sonidos", "Ave"))
        .Ave2 = Val(Leer.GetValue("Sonidos", "Ave2"))
        .Ave3 = Val(Leer.GetValue("Sonidos", "Ave3"))
        .Grillo = Val(Leer.GetValue("Sonidos", "Grillo"))
        .Grillo2 = Val(Leer.GetValue("Sonidos", "Grillo2"))
        .SacarArma = Val(Leer.GetValue("Sonidos", "SacarArma"))
        .Escudo = Val(Leer.GetValue("Sonidos", "Escudo"))
        .MartilloHerrero = Val(Leer.GetValue("Sonidos", "MartilloHerrero"))
        .TrabajoCarpintero = Val(Leer.GetValue("Sonidos", "TrabajoCarpintero"))
        .Tomar = Val(Leer.GetValue("Sonidos", "Tomar"))
    End With
    
    With ConstantesBosses
        .BossDMCastPutrefaccion = Val(Leer.GetValue("Bosses", "BossDMCastPutrefaccion"))
        .BossDMCastAparicion = Val(Leer.GetValue("Bosses", "BossDMCastAparicion"))
        .BossDMSpellAparicion = Val(Leer.GetValue("Bosses", "BossDMSpellAparicion"))
        .BossDMSpellPutrefaccion = Val(Leer.GetValue("Bosses", "BossDMSpellPutrefaccion"))
        .BossDVDistance = Val(Leer.GetValue("Bosses", "BossDVDistance"))
        .BossDVChangeTarget = Val(Leer.GetValue("Bosses", "BossDVChangeTarget"))
        .BossDVCastPetrificar = Val(Leer.GetValue("Bosses", "BossDVCastPetrificar"))
        .BossDVCastTormenta = Val(Leer.GetValue("Bosses", "BossDVCastTormenta"))
        .BossDVCastDescarga = Val(Leer.GetValue("Bosses", "BossDVCastDescarga"))
        .BossDVSpellPetrificar = Val(Leer.GetValue("Bosses", "BossDVSpellPetrificar"))
        .BossDVSpellTormenta = Val(Leer.GetValue("Bosses", "BossDVSpellTormenta"))
        .BossDVSpellDescarga = Val(Leer.GetValue("Bosses", "BossDVSpellDescarga"))
        .BossDINumDebuff = Val(Leer.GetValue("Bosses", "BossDINumDebuff"))
        ReDim .BossDISpellDebuff(1 To .BossDINumDebuff) As Integer
        
        Dim BossId As Byte
        For BossId = 1 To .BossDINumDebuff
            .BossDISpellDebuff(BossId) = Val(Leer.GetValue("Bosses", "BossDISpellDebuff" & BossId))
        Next BossId
        .BossDISpellBola = Val(Leer.GetValue("Bosses", "BossDISpellBola"))
        .BossDACastTorrente = Val(Leer.GetValue("Bosses", "BossDACastTorrente"))
        .BossDACastTentaculo = Val(Leer.GetValue("Bosses", "BossDACastTentaculo"))
        .BossDACastAtraer = Val(Leer.GetValue("Bosses", "BossDACastAtraer"))
        .BossDACastAplastar = Val(Leer.GetValue("Bosses", "BossDACastAplastar"))
        .BossDAAplastarArea = Val(Leer.GetValue("Bosses", "BossDAAplastarArea"))
        .BossDASpellTorrente = Val(Leer.GetValue("Bosses", "BossDASpellTorrente"))
        .BossDASpellTentaculo = Val(Leer.GetValue("Bosses", "BossDASpellTentaculo"))
        .BossDASpellAtraer = Val(Leer.GetValue("Bosses", "BossDASpellAtraer"))
    End With
    
    Set Leer = Nothing
    
    ReDim ListaPeces(1 To ConstantesItems.NumPescados) As Integer
    
    ' Fishes
    ListaPeces(1) = ConstantesItems.Pescado1
    ListaPeces(2) = ConstantesItems.Pescado2
    ListaPeces(3) = ConstantesItems.Pescado3
    ListaPeces(4) = ConstantesItems.Pescado4

    'Bordes del mapa
    MinXBorder = XMinMapSize + (XWindow \ 2)
    MaxXBorder = XMaxMapSize - (XWindow \ 2)
    MinYBorder = YMinMapSize + (YWindow \ 2)
    MaxYBorder = YMaxMapSize - (YWindow \ 2)
    
    Set Ayuda = New cCola
    Set Denuncias = New cCola
    Denuncias.MaxLenght = Constantes.MaxDenuncias

    With Prision
        .Map = 66
        .X = 75
        .Y = 47
    End With
    
    With Libertad
        .Map = 66
        .X = 75
        .Y = 65
    End With

    ' Initialize classes
    Protocol.InitAuxiliarBuffer
#If EnableSecurity Then
    Set aDos = New clsAntiDoS
#End If

    Set aClon = New clsAntiMassClon
    Set TrashCollector = New Collection

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadConstants de General.bas")
End Sub

Private Sub LoadArrays()
'*****************************************************************
'Author: ZaMa
'Last Modify Date: 15/03/2011
'Loads all arrays
'*****************************************************************
On Error GoTo ErrHandler
  
On Error Resume Next
    ' Load Records
    Call LoadRecords
    ' Load guilds info
    ' Load spawn list
    Call CargarSpawnList
    ' Load forbidden words
    Call CargarForbidenWords
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadArrays de General.bas")
End Sub

Private Sub ResetUsersConnections()
'*****************************************************************
'Author: ZaMa
'Last Modify Date: 15/03/2011
'Resets Users Connections.
'*****************************************************************
On Error GoTo ErrHandler
  
On Error Resume Next

    Dim LoopC As Long
    For LoopC = 1 To MaxUsers
        UserList(LoopC).ConnIDValida = False
    Next LoopC
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ResetUsersConnections de General.bas")
End Sub


Public Sub InitUserArrays()
On Error GoTo ErrHandler:

    Dim LoopC
    For LoopC = 1 To MaxUsers
        With UserList(LoopC)
            'Private Messages
            ReDim .Mensajes(1 To Constantes.MaxPrivateMessages)
            
            ReDim .flags.ActiveTraps(1 To ConstantesBalance.MaxActiveTrapQty)
        End With
    Next LoopC
    Exit Sub
    
ErrHandler:
    Call LogError("Error en InitUserArrays: " & Err.Description)
    
End Sub

Private Sub InitMainTimers()
'*****************************************************************
'Author: ZaMa
'Last Modify Date: 15/03/2011
'Initializes Main Timers.
'*****************************************************************
On Error GoTo ErrHandler
  
On Error Resume Next

    With frmMain
        .AutoSave.Enabled = True
        .tLluvia.Enabled = True
        .tPiqueteC.Enabled = True
        .GameTimer.Enabled = True
        .tLluviaEvent.Enabled = False
        .FX.Enabled = True
        .Auditoria.Enabled = True
        .KillLog.Enabled = True
        .TIMER_AI.Enabled = True
        '.npcataca.Enabled = True
        
#If EnableSecurity Then
        .securityTimer.Enabled = True
#End If
    End With
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub InitMainTimers de General.bas")
End Sub

Private Sub SocketConfig()
'*****************************************************************
'Author: ZaMa
'Last Modify Date: 15/03/2011
'Sets socket config.
'*****************************************************************
On Error GoTo ErrHandler
  
On Error Resume Next

    Call SecurityIp.InitIpTables(1000)
    
    TCP.Listen "0.0.0.0", CStr(Puerto)
    
    If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SocketConfig de General.bas")
End Sub

Private Sub LogServerStartTime()
'*****************************************************************
'Author: ZaMa
'Last Modify Date: 15/03/2011
'Logs Server Start Time.
'*****************************************************************
On Error GoTo ErrHandler
  
    Dim N As Integer
    N = FreeFile
    Open ServerConfiguration.LogsPaths.GeneralPath & "Main.log" For Append Shared As #N
    Print #N, Date & " " & Time & " server iniciado " & App.Major & "."; App.Minor & "." & App.Revision
    Close #N

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LogServerStartTime de General.bas")
End Sub

Function FileExist(ByVal File As String, Optional FileType As VbFileAttribute = vbNormal) As Boolean
'*****************************************************************
'Se fija si existe el archivo
'*****************************************************************
On Error GoTo ErrHandler
  

    FileExist = LenB(dir$(File, FileType)) <> 0
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function FileExist de General.bas")
End Function

Function ReadField(ByVal Pos As Integer, ByRef Text As String, ByVal SepASCII As Byte) As String
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/15/2004
'Gets a field from a delimited string
'*****************************************************************
On Error GoTo ErrHandler
  

    Dim positionIndex As Long
    Dim lastPos As Long
    Dim CurrentPos As Long
    Dim delimiter As String * 1
    
    delimiter = Chr$(SepASCII)
    
    For positionIndex = 1 To Pos
        lastPos = CurrentPos
        CurrentPos = InStr(lastPos + 1, Text, delimiter, vbBinaryCompare)
    Next positionIndex
    
    If CurrentPos = 0 Then
        ReadField = mid$(Text, lastPos + 1, Len(Text) - lastPos)
    Else
        ReadField = mid$(Text, lastPos + 1, CurrentPos - lastPos - 1)
    End If
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function ReadField de General.bas")
End Function

Function MapaValido(ByVal Map As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    MapaValido = Map >= 1 And Map <= NumMaps
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function MapaValido de General.bas")
End Function

Sub MostrarNumUsers()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    frmMain.txtNumUsers.Text = NumUsers

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub MostrarNumUsers de General.bas")
End Sub

Public Sub LogCriticEvent(Desc As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo ErrHandler

    Call SaveCriticEventDB(Desc)

    Exit Sub

ErrHandler:

End Sub

Public Sub LogEjercitoReal(Desc As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo ErrHandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open ServerConfiguration.LogsPaths.GeneralPath & "EjercitoReal.log" For Append Shared As #nfile
    Print #nfile, Desc
    Close #nfile
    
    Exit Sub

ErrHandler:

End Sub

Public Sub LogEjercitoCaos(Desc As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo ErrHandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open ServerConfiguration.LogsPaths.GeneralPath & "EjercitoCaos.log" For Append Shared As #nfile
    Print #nfile, Desc
    Close #nfile

Exit Sub

ErrHandler:

End Sub

Public Sub LogError(Desc As String)

    Desc = "AO: S - " & Desc
    Call OutputDebugString(Desc)
    
End Sub

Public Sub LogErrorDB(Desc As String, Sql As String)
'***************************************************
'Author: ZaMa
'Last Modification: -
'
'***************************************************

On Error GoTo ErrHandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open ServerConfiguration.LogsPaths.GeneralPath & "erroresDB.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & Desc & "; Sql: " & Sql
    Close #nfile

    Exit Sub

ErrHandler:

End Sub

Public Sub LogDesarrollo(ByVal str As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open ServerConfiguration.LogsPaths.DevelopmentPath & "desarrollo" & Month(Date) & Year(Date) & ".log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & str
    Close #nfile

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LogDesarrollo de General.bas")
End Sub

Public Sub LogGM(Nombre As String, texto As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo ErrHandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    'Guardamos todo en el mismo lugar. Pablo (ToxicWaste) 18/05/07
    Open ServerConfiguration.LogsPaths.GameMastersPath & Nombre & ".log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & texto
    Close #nfile
    
    Exit Sub

ErrHandler:

End Sub

Public Sub LogAsesinato(texto As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo ErrHandler
    Dim nfile As Integer
    
    nfile = FreeFile ' obtenemos un canal
    
    Open ServerConfiguration.LogsPaths.GeneralPath & "asesinatos.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & texto
    Close #nfile
    
    Exit Sub

ErrHandler:

End Sub

Public Sub logVentaCasa(ByVal texto As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo ErrHandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    
    Open ServerConfiguration.LogsPaths.GeneralPath & "propiedades.log" For Append Shared As #nfile
    Print #nfile, "----------------------------------------------------------"
    Print #nfile, Date & " " & Time & " " & texto
    Print #nfile, "----------------------------------------------------------"
    Close #nfile
    
    Exit Sub

ErrHandler:

End Sub
Public Sub LogHackAttemp(texto As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo ErrHandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open ServerConfiguration.LogsPaths.GeneralPath & "HackAttemps.log" For Append Shared As #nfile
    Print #nfile, "----------------------------------------------------------"
    Print #nfile, Date & " " & Time & " " & texto
    Print #nfile, "----------------------------------------------------------"
    Close #nfile
    
    Exit Sub

ErrHandler:

End Sub

Public Sub LogCheating(texto As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo ErrHandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open ServerConfiguration.LogsPaths.GeneralPath & "CH.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & texto
    Close #nfile
    
    Exit Sub

ErrHandler:

End Sub

Sub Restart()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

'Se asegura de que los sockets estan cerrados e ignora cualquier err
On Error Resume Next

    If frmMain.Visible Then frmMain.txStatus.Caption = "Reiniciando."
    
    Dim LoopC As Long
    
    'Cierra el socket de escucha
    TCP.Disconnect
    
    'Inicia el socket de escucha
    TCP.Listen "0.0.0.0", CStr(Puerto)
    
    For LoopC = 1 To MaxUsers
        Call CloseSocket(LoopC)
    Next

    ReDim UserList(1 To MaxUsers) As User
    
    For LoopC = 1 To MaxUsers
        UserList(LoopC).ConnIDValida = False
    Next LoopC
    
    LastUser = 0
    NumUsers = 0
    
    Call InitUserArrays
    
    Call FreeNPCs
    Call FreeCharIndexes
    
    Call LoadSini
    
    Call ResetForums
    Call LoadOBJData
    
    Call LoadMapData
    
    Call CargarHechizos
    
    If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."
    
    'Log it
    Dim N As Integer
    N = FreeFile
    Open ServerConfiguration.LogsPaths.GeneralPath & "Main.log" For Append Shared As #N
    Print #N, Date & " " & Time & " servidor reiniciado."
    Close #N
    
    'Ocultar
    
    If HideMe = 1 Then
        Call frmMain.InitMain(1)
    Else
        Call frmMain.InitMain(0)
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Restart de General.bas")
End Sub


Public Function Intemperie(ByVal UserIndex As Integer) As Boolean
'**************************************************************
'Author: Unknown
'Last Modify Date: 15/11/2009
'15/11/2009: ZaMa - La lluvia no quita stamina en las arenas.
'23/11/2009: ZaMa - Optimizacion de codigo.
'**************************************************************
On Error GoTo ErrHandler
  

    With UserList(UserIndex)
        If MapInfo(.Pos.Map).Zona <> eTerrainZone.zone_dungeon Then
            If MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger <> eTrigger.BAJOTECHO And _
               MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger <> eTrigger.trigger_2 And _
               MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger <> eTrigger.ZONASEGURA Then Intemperie = True
        Else
            Intemperie = False
        End If
    End With
    
    'En las arenas no te afecta la lluvia
    If IsArena(UserIndex) Then Intemperie = False
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function Intemperie de General.bas")
End Function

Public Sub EfectoLluvia(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo ErrHandler

    If UserList(UserIndex).flags.UserLogged Then
        If Intemperie(UserIndex) Then
            Dim modifi As Long
            modifi = Porcentaje(UserList(UserIndex).Stats.MaxSta, 3)
            Call QuitarSta(UserIndex, modifi)
        End If
    End If
    
    Exit Sub
ErrHandler:
    LogError ("Error en EfectoLluvia")
End Sub



Public Sub TiempoInvocacion(ByVal UserIndex As Integer)
On Error GoTo ErrHandler
    Dim I As Integer
    
    With UserList(UserIndex)
        For I = 1 To Classes(.clase).ClassMods.MaxInvokedPets
            If .InvokedPets(I).NpcIndex > 0 Then
                If IsIntervalReached(Npclist(.InvokedPets(I).NpcIndex).Contadores.TiempoExistencia) Then
                    Call QuitarInvocacion(UserIndex, I)
                End If
                
            End If
        Next I
    End With

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub TiempoInvocacion de General.bas")
End Sub

Public Sub EfectoFrio(ByVal UserIndex As Integer)
On Error GoTo ErrHandler
  
    With UserList(UserIndex)
        If IsIntervalReached(.Counters.Frio) Then
            
            If MapInfo(.Pos.Map).NakedLosesEnergy Then
                Call QuitarSta(UserIndex, Porcentaje(.Stats.MaxSta, 5))
                Call WriteUpdateSta(UserIndex)
            End If
            
            If MapInfo(.Pos.Map).NakedLosesHealth Then
                Call WriteConsoleMsg(UserIndex, "¡¡Estas expuesto a la intemperie, vístete o morirás!!", FontTypeNames.FONTTYPE_INFO)
       
                .Stats.MinHp = .Stats.MinHp - MinimoInt(Porcentaje(.Stats.MaxHp, 5), 15)
                
                If .Stats.MinHp < 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡¡Has muerto de frío!!", FontTypeNames.FONTTYPE_INFO)
                    .Stats.MinHp = 0
                    Call UserDie(UserIndex)
                End If
                
                Call WriteUpdateHP(UserIndex)
            End If
            
            .Counters.Frio = SetIntervalEnd(ServerConfiguration.Intervals.IntervaloFrio)

        End If
    End With
  
    Exit Sub
  
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub EfectoFrio de General.bas")
End Sub

Public Sub EfectoLava(ByVal UserIndex As Integer)
On Error GoTo ErrHandler
  
      With UserList(UserIndex)
        If IsIntervalReached(.Counters.Lava) Then
            If HayLava(.Pos.Map, .Pos.X, .Pos.Y) Then
                Call WriteConsoleMsg(UserIndex, "¡¡Quitate de la lava, te estás quemando!!", FontTypeNames.FONTTYPE_INFO)
                
                .Stats.MinHp = .Stats.MinHp - Porcentaje(.Stats.MaxHp, 5)
                
                If .Stats.MinHp < 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡¡Has muerto quemado!!", FontTypeNames.FONTTYPE_INFO)
                    .Stats.MinHp = 0
                    Call UserDie(UserIndex)
                End If

                Call WriteUpdateHP(UserIndex)

            End If
            
            .Counters.Lava = SetIntervalEnd(ServerConfiguration.Intervals.IntervaloLava)

        End If
    End With
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub EfectoLava de General.bas")
End Sub

''
' Maneja  el efecto del estado atacable
'
' @param UserIndex  El index del usuario a ser afectado por el estado atacable
'

Public Sub EfectoEstadoAtacable(ByVal UserIndex As Integer)
'******************************************************
'Author: ZaMa
'Last Update: 18/09/2010 (ZaMa)
'18/09/2010: ZaMa - Ahora se activa el seguro cuando dejas de ser atacable.
'******************************************************
On Error GoTo ErrHandler
  

    ' Si ya paso el tiempo de penalizacion
    If Not IntervaloEstadoAtacable(UserIndex) Then
        ' Deja de poder ser atacado
        UserList(UserIndex).flags.AtacablePor = 0
        
        ' Activo el seguro si deja de estar atacable
        If Not UserList(UserIndex).flags.Seguro Then
            Call WriteMultiMessage(UserIndex, eMessages.SafeModeOn)
        End If
        
        ' Send nick normal
        Call RefreshCharStatus(UserIndex, False)
    End If
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub EfectoEstadoAtacable de General.bas")
End Sub

''
' Maneja el tiempo de arrivo al hogar
'
' @param UserIndex  El index del usuario a ser afectado por el /hogar
'

Public Sub TravelingEffect(ByVal UserIndex As Integer)
'******************************************************
'Author: ZaMa
'Last Update: 01/06/2010 (ZaMa)
'******************************************************
On Error GoTo ErrHandler
  

    ' Si ya paso el tiempo de penalizacion
    If IntervaloGoHome(UserIndex) Then
        Call HomeArrival(UserIndex)
    End If

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub TravelingEffect de General.bas")
End Sub

''
' Maneja el tiempo y el efecto del mimetismo
'
' @param UserIndex  El index del usuario a ser afectado por el mimetismo
'

Public Sub EfectoMimetismo(ByVal UserIndex As Integer)
'******************************************************
'Author: Unknown
'Last Update: 16/09/2010 (ZaMa)
'12/01/2010: ZaMa - Los druidas pierden la inmunidad de ser atacados cuando pierden el efecto del mimetismo.
'16/09/2010: ZaMa - Se recupera la apariencia de la barca correspondiente despues de terminado el mimetismo.
'******************************************************
On Error GoTo ErrHandler
  
    
    With UserList(UserIndex)
        If .Counters.Mimetismo = -1 Then Exit Sub  '/MIMETIZAR, /IMPERSONAR
        
        If IsIntervalReached(.Counters.Mimetismo) Then
            Call EndMimic(UserIndex, True, True)
            
            With .Char
                Call ChangeUserChar(UserIndex, .body, .head, .heading, .WeaponAnim, .ShieldAnim, .CascoAnim)
            End With
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub EfectoMimetismo de General.bas")
End Sub

Public Sub EfectoInmunidad(ByVal UserIndex As Integer)
On Error GoTo ErrHandler
    With UserList(UserIndex)
        If IsIntervalReached(.Counters.Inmunidad) Then
            .Counters.Inmunidad = 0
            .flags.Inmunidad = 0
            
            ' Step on trigger?
            Call CheckTriggerActivation(UserIndex, 0, .Pos.Map, .Pos.X, .Pos.Y, False)
            
            'Call WriteConsoleMsg(UserIndex, "Has Perdido la proteccion a las trampas.", FontTypeNames.FONTTYPE_INFO)
        End If
        Exit Sub
    End With
    
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub EfectoInmunidad de General.bas")
End Sub

Public Sub EfectoInvisibilidad(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 16/09/2010 (ZaMa)
'16/09/2010: ZaMa - Al perder el invi cuando navegas, no se manda el mensaje de sacar invi (ya estas visible).
'***************************************************
On Error GoTo ErrHandler


    With UserList(UserIndex)
        If IsIntervalReached(.Counters.Invisibilidad) Then
            .Counters.Invisibilidad = 0
            .flags.invisible = 0
            If .flags.Oculto = 0 Then
                Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
    
                ' Si navega ya esta visible..
                If Not ThiefRestoreBoatAppearance(UserIndex) Then
                    'Si está en un oscuro no lo hacemos visible
                    If MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger <> eTrigger.zonaOscura Then
                        Call UsUaRiOs.SetInvisible(UserIndex, .Char.CharIndex, False)
                    End If
                End If
            End If
                    
            If BerzerkConditionMet(UserIndex) Then
                Call ActivatePassive(UserIndex, ePassiveSpells.Berserk, True)
                Call SendBerserkEffect(UserIndex, ePassiveSpells.Berserk, True)
            End If
        End If
        
    End With


    Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub EfectoInvisibilidad de General.bas")
End Sub

Public Sub EfectoParalisisNpc(ByVal NpcIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    With Npclist(NpcIndex)
        If IsIntervalReached(.Contadores.Paralisis) Then
            .flags.Paralizado = 0
            .flags.Inmovilizado = 0
        End If
    End With

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub EfectoParalisisNpc de General.bas")
End Sub

Public Sub EfectoCegueEstu(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    With UserList(UserIndex)
        If IsIntervalReached(.Counters.Ceguera) Then
            If .flags.Ceguera = 1 Then
                .flags.Ceguera = 0
                Call WriteBlindNoMore(UserIndex)
            End If
            If .flags.Estupidez = 1 Then
                .flags.Estupidez = 0
                Call WriteDumbNoMore(UserIndex)
            End If
        End If
    End With

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub EfectoCegueEstu de General.bas")
End Sub

Public Sub EfectoParalisisUser(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 02/12/2010
'02/12/2010: ZaMa - Now non-magic clases lose paralisis effect under certain circunstances.
'***************************************************
On Error GoTo ErrHandler
  

    With UserList(UserIndex)
        If .flags.Putrefaccion Then Exit Sub
        If .flags.Petrificado Then Exit Sub
        
        If IsIntervalReached(.Counters.Paralisis) Then
            Call RemoveParalisis(UserIndex)
        Else
            Dim CasterIndex As Integer
            CasterIndex = .flags.ParalizedByIndex
        
            ' Only aplies to non-magic clases
            If .Stats.MaxMan = 0 Then
                ' Paralized by user?
                If CasterIndex <> 0 Then
                
                    ' Close? => Remove Paralisis
                    If UserList(CasterIndex).Name <> .flags.ParalizedBy Then
                        Call RemoveParalisis(UserIndex)
                        Exit Sub
                        
                    ' Caster dead? => Remove Paralisis
                    ElseIf UserList(CasterIndex).flags.Muerto = 1 Then
                        Call RemoveParalisis(UserIndex)
                        Exit Sub
                    
                    ElseIf GetIntervalRemainingTime(.Counters.Paralisis) > ServerConfiguration.Intervals.IntervaloParalizadoReducido Then
                        ' Out of vision range? => Reduce paralisis counter
                        If Not InVisionRangeAndMap(UserIndex, UserList(CasterIndex).Pos) Then
                            ' 1500 ms
                            .Counters.Paralisis = SetIntervalEnd(ServerConfiguration.Intervals.IntervaloParalizadoReducido)
                            Exit Sub
                        End If
                    End If
                
                ' Npc?
                Else
                    CasterIndex = .flags.ParalizedByNpcIndex
                    
                    ' Paralized by npc?
                    If CasterIndex <> 0 Then
                    
                        If GetIntervalRemainingTime(.Counters.Paralisis) > ServerConfiguration.Intervals.IntervaloParalizadoReducido Then
                            ' Out of vision range? => Reduce paralisis counter
                            If Not InVisionRangeAndMap(UserIndex, Npclist(CasterIndex).Pos) Then
                                ' 1500 ms
                                .Counters.Paralisis = SetIntervalEnd(ServerConfiguration.Intervals.IntervaloParalizadoReducido)
                                Exit Sub
                            End If
                        End If
                    End If
                    
                End If
            End If
        End If
    End With

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub EfectoParalisisUser de General.bas")
End Sub

Public Sub RemoveParalisis(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 20/11/2010
'Removes paralisis effect from user.
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)
        .flags.Paralizado = 0
        .flags.Inmovilizado = 0
        .flags.ParalizedBy = vbNullString
        .flags.ParalizedByIndex = 0
        .flags.ParalizedByNpcIndex = 0
        .flags.Putrefaccion = 0
        .flags.Petrificado = 0
        .Counters.Paralisis = 0
        .Counters.Putrefaccion = 0
        Call WriteParalizeOK(UserIndex)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub RemoveParalisis de General.bas")
End Sub

Public Sub EfectoPetrificado(ByVal UserIndex As Integer)
On Error GoTo ErrHandler
    
    With UserList(UserIndex)
        If Not IsIntervalReached(.Counters.Petrificado) Then Exit Sub

        .flags.Petrificado = 0
        .flags.Paralizado = 0
        Call WriteParalizeOK(UserIndex)
    End With

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub EfectoPetrificado de General.bas")
End Sub

Public Sub EfectoPutrefaccion(ByVal UserIndex As Integer)
On Error GoTo ErrHandler
    Dim Damage As Integer
    
    With UserList(UserIndex)
        If IsIntervalReached(.Counters.Putrefaccion) Then
            .flags.Putrefaccion = 0
            .flags.Paralizado = 0
            Call WriteParalizeOK(UserIndex)
        ElseIf IsIntervalReached(.Counters.PutrefaccionDmg) Then
            .Counters.PutrefaccionDmg = SetIntervalEnd(ServerConfiguration.Intervals.IntervaloPutrefaccionDmg)

            Damage = RandomNumber(Hechizos(.flags.Putrefaccion).MinHp, Hechizos(.flags.Putrefaccion).MaxHp)
            
            If .flags.Privilegios And PlayerType.User Then
            .Stats.MinHp = .Stats.MinHp - Damage
            End If
            
            Call WriteConsoleMsg(UserIndex, "La putrefacción te ha quitado " & Damage & " de vida.", FontTypeNames.FONTTYPE_FIGHT, _
                                  eMessageType.Combate)
                                 
            If .Stats.MinHp < 1 Then Call UserDie(UserIndex)
            Call WriteUpdateHP(UserIndex)
        End If

    End With

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub EfectoPutrefaccion de General.bas")
End Sub

Public Sub EfectoVeneno(ByVal UserIndex As Integer)
On Error GoTo ErrHandler
    
    With UserList(UserIndex)
        If Not IsIntervalReached(.Counters.Veneno) Then Exit Sub
        
        Call WriteConsoleMsg(UserIndex, "Estás envenenado, si no te curas morirás.", FontTypeNames.FONTTYPE_VENENO, eMessageType.Combate)
        .Counters.Veneno = SetIntervalEnd(ServerConfiguration.Intervals.IntervaloVeneno)
        .Stats.MinHp = .Stats.MinHp - RandomNumber(1, 5)

        If .Stats.MinHp < 1 Then Call UserDie(UserIndex)
        Call WriteUpdateHP(UserIndex)

    End With
  
    Exit Sub
  
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub EfectoVeneno de General.bas")
End Sub

Public Sub DuracionPociones(ByVal UserIndex As Integer)
'***************************************************
'Author: ??????
'Last Modification: 08/06/11 (CHOTS)
'Le agregué que avise antes cuando se te está por ir
'
'Cuando se pierde el efecto de la poción updatea fz y agi (No me gusta que ambos atributos aunque se haya modificado solo uno, pero bueno :p)
'***************************************************
On Error GoTo ErrHandler
  

Const SEGUNDOS_AVISO As Byte = 5
'CHOTS | Los segundos antes que se te acabe que te avisa

    With UserList(UserIndex)
        'Controla la duracion de las pociones
        If .flags.DuracionEfecto > 0 Then
            .flags.DuracionEfecto = .flags.DuracionEfecto - 1
            If ((.flags.DuracionEfecto / 25) <= SEGUNDOS_AVISO) And (Not .flags.bStrDextRunningOutNotified) Then
            
                If .Stats.UserAtributos(eAtributos.Agilidad) > .Stats.UserAtributosBackUP(eAtributos.Agilidad) Or .Stats.UserAtributos(eAtributos.Fuerza) > .Stats.UserAtributosBackUP(eAtributos.Fuerza) Then
                    Call WriteStrDextRunningOut(UserIndex)
                    '.flags.UltimoMensaje = 221
                    .flags.bStrDextRunningOutNotified = True
                End If
            End If
            If .flags.DuracionEfecto = 0 Then
                .flags.UltimoMensaje = 222
                .flags.TomoPocion = False
                .flags.TipoPocion = 0
                'volvemos los atributos al estado normal
                Dim loopX As Integer
                
                For loopX = 1 To NUMATRIBUTOS
                    .Stats.UserAtributos(loopX) = .Stats.UserAtributosBackUP(loopX)
                Next loopX
                
                Call WriteUpdateStrenghtAndDexterity(UserIndex)
                .flags.bStrDextRunningOutNotified = False
                
                ' Disable the Berzerk
                If HasPassiveAssigned(UserIndex, ePassiveSpells.Berserk) Then
                    If HasPassiveActivated(UserIndex, ePassiveSpells.Berserk) Then
                        Call ActivatePassive(UserIndex, ePassiveSpells.Berserk, False)
                        Call SendBerserkEffect(UserIndex, ePassiveSpells.Berserk, False)
                    End If
                End If
                
           End If
        End If
    End With

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub DuracionPociones de General.bas")
End Sub

Public Sub DoSed(ByVal UserIndex As Integer, fenviarAyS As Boolean)
On Error GoTo ErrHandler
  
    With UserList(UserIndex)
        If ServerConfiguration.Intervals.IntervaloSed = 0 Then Exit Sub
        If Not .Stats.MinAGU > 0 Then Exit Sub
        If .CraftingStore.IsOpen Then Exit Sub
        If Not IsIntervalReached(.Counters.AGUACounter) Then Exit Sub
        
        .Counters.AGUACounter = SetIntervalEnd(ServerConfiguration.Intervals.IntervaloSed)
        .Stats.MinAGU = .Stats.MinAGU - 10
        
        If .Stats.MinAGU <= 0 Then
            .Stats.MinAGU = 0
            .flags.Sed = 1
        End If
        
        fenviarAyS = True
    End With

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub DoHambre de General.bas")
End Sub

Public Sub DoHambre(ByVal UserIndex As Integer, ByRef fenviarAyS As Boolean)
On Error GoTo ErrHandler

    With UserList(UserIndex)
        If ServerConfiguration.Intervals.IntervaloHambre = 0 Then Exit Sub
        If Not .Stats.MinHam > 0 Then Exit Sub
        If .CraftingStore.IsOpen Then Exit Sub
        If Not IsIntervalReached(.Counters.COMCounter) Then Exit Sub

        .Counters.COMCounter = SetIntervalEnd(ServerConfiguration.Intervals.IntervaloHambre)
        .Stats.MinHam = .Stats.MinHam - 10
        
        If .Stats.MinHam <= 0 Then
               .Stats.MinHam = 0
               .flags.Hambre = 1
        End If
        
        fenviarAyS = True
    End With

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub DoSed de General.bas")
End Sub

Private Sub DoStaminaRecovery(ByVal UserIndex As Integer, ByRef EnviarStats As Boolean, Optional ByVal MaxStaRecoveryPerc As Byte)
On Error GoTo ErrHandler

    Dim AddStamina As Integer
    With UserList(UserIndex)
        
        EnviarStats = True
        
        MaxStaRecoveryPerc = MaxStaRecoveryPerc Or Constantes.MaxStaRecoveryPerc
        AddStamina = RandomNumber(Porcentaje(.Stats.MaxSta, Constantes.MinStaRecoveryPerc), Porcentaje(.Stats.MaxSta, MaxStaRecoveryPerc))
        
        .Stats.MinSta = .Stats.MinSta + AddStamina
        If .Stats.MinSta > .Stats.MaxSta Then
            .Stats.MinSta = .Stats.MaxSta
        End If
        
    End With

Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub DoStaminaRecovery de General.bas")
End Sub

Private Function IsValidStaminaRecovery(ByVal UserIndex As Integer) As Boolean
On Error GoTo ErrHandler

    With UserList(UserIndex)
        IsValidStaminaRecovery = .Stats.MinSta < .Stats.MaxSta And Not .flags.Desnudo = 1
    End With
    
Exit Function
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub IsValidStaminaRecovery de General.bas")
End Function

Public Sub DoNormalStaminaRecovery(ByVal UserIndex As Integer, ByRef EnviarStats As Boolean)
On Error GoTo ErrHandler

     With UserList(UserIndex)
        If Not IsValidStaminaRecovery(UserIndex) Then Exit Sub
        If Not IsIntervalReached(.Counters.STACounter) Then Exit Sub
        
        .Counters.STACounter = SetIntervalEnd(ServerConfiguration.Intervals.StaminaIntervaloSinDescansar)
        Call DoStaminaRecovery(UserIndex, EnviarStats)
    End With

Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub DoNormalStaminaRecovery de General.bas")
End Sub

Public Sub DoRestingStaminaRecovery(ByVal UserIndex As Integer, ByRef EnviarStats As Boolean)
On Error GoTo ErrHandler

     With UserList(UserIndex)
        If Not IsValidStaminaRecovery(UserIndex) Then Exit Sub
        If Not IsIntervalReached(.Counters.RestingSTACounter) Then Exit Sub
        
        Dim CampfireObj As Integer
        CampfireObj = MapData(.RestObjectCoords.Map, .RestObjectCoords.X, .RestObjectCoords.Y).ObjInfo.ObjIndex
        
        If ObjData(CampfireObj).ObjType = otFogata Then
            .Counters.RestingSTACounter = SetIntervalEnd(ServerConfiguration.Intervals.StaminaIntervaloDescansar)

            Call DoStaminaRecovery(UserIndex, EnviarStats, ObjData(CampfireObj).MaxStaRecoveryPerc)
        End If
        
        
    End With

Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub DoRestingStaminaRecovery de General.bas")
End Sub

Private Sub DoHeal(ByVal UserIndex As Integer, ByRef EnviarStats As Boolean)
On Error GoTo ErrHandler

    Dim Rnd As Integer
    With UserList(UserIndex)
        Rnd = RandomNumber(2, Porcentaje(.Stats.MaxSta, 5))
        
        .Stats.MinHp = .Stats.MinHp + Rnd
        If .Stats.MinHp > .Stats.MaxHp Then .Stats.MinHp = .Stats.MaxHp
        Call WriteConsoleMsg(UserIndex, "Has sanado.", FontTypeNames.FONTTYPE_INFO)
        EnviarStats = True
        
    End With

Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub DoHeal de General.bas")
End Sub

Private Function IsValidHeal(ByVal UserIndex As Integer) As Boolean
On Error GoTo ErrHandler

    With UserList(UserIndex)
        IsValidHeal = .Stats.MinHp < .Stats.MaxHp
    End With
    
Exit Function
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub IsValidHeal de General.bas")
End Function

Public Sub DoNormalHeal(ByVal UserIndex As Integer, ByRef EnviarStats As Boolean)
On Error GoTo ErrHandler

     With UserList(UserIndex)
        If Not IsValidHeal(UserIndex) Then Exit Sub
        If Not IsIntervalReached(.Counters.HPCounter) Then Exit Sub
        
        .Counters.HPCounter = SetIntervalEnd(ServerConfiguration.Intervals.SanaIntervaloSinDescansar)
        Call DoHeal(UserIndex, EnviarStats)
    End With

Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub DoNormalHeal de General.bas")
End Sub

Public Sub DoRestingHeal(ByVal UserIndex As Integer, ByRef EnviarStats As Boolean)
On Error GoTo ErrHandler

     With UserList(UserIndex)
        If Not IsValidHeal(UserIndex) Then Exit Sub
        If Not IsIntervalReached(.Counters.RestingHPCounter) Then Exit Sub
        
        .Counters.RestingHPCounter = SetIntervalEnd(ServerConfiguration.Intervals.SanaIntervaloDescansar)
        Call DoHeal(UserIndex, EnviarStats)
    End With

Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub DoRestingHeal de General.bas")
End Sub


Public Sub DoRegeneration(ByVal UserIndex As Integer, ByRef EnviarStats As Boolean)
'***************************************************
'Author: Nightw
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler

    With UserList(UserIndex)
        If .Stats.UserAtributos(eAtributos.Agilidad) > MinimoInt(ConstantesBalance.MaxAtributos, .Stats.UserAtributosBackUP(Agilidad) * 2) Then Exit Sub
        If .Stats.UserAtributos(eAtributos.Fuerza) > MinimoInt(ConstantesBalance.MaxAtributos, .Stats.UserAtributosBackUP(Fuerza) * 2) Then Exit Sub
        
        If Not IsIntervalReached(.Counters.RegenerationCounter) Then Exit Sub
        .Counters.RegenerationCounter = SetIntervalEnd(ServerConfiguration.Intervals.SanaIntervaloDescansar)
        
        Dim Rnd As Integer

        If IsValidHeal(UserIndex) Then
            Rnd = RandomNumber(1, Porcentaje(.Stats.MaxSta, 5))
            .Stats.MinHp = .Stats.MinHp + Rnd
            
            If .Stats.MinHp > .Stats.MaxHp Then .Stats.MinHp = .Stats.MaxHp
        End If

        If .Stats.UserAtributos(eAtributos.Agilidad) <= MinimoInt(ConstantesBalance.MaxAtributos, .Stats.UserAtributosBackUP(Agilidad) * 2) And _
           .Stats.UserAtributos(eAtributos.Fuerza) <= MinimoInt(ConstantesBalance.MaxAtributos, .Stats.UserAtributosBackUP(Fuerza) * 2) Then
           
            Rnd = RandomNumber(2, 4)
            
            .flags.TomoPocion = True
            .flags.DuracionEfecto = 1200
        
            .Stats.UserAtributos(eAtributos.Agilidad) = .Stats.UserAtributos(eAtributos.Agilidad) + Rnd
            If .Stats.UserAtributos(eAtributos.Agilidad) > ConstantesBalance.MaxAtributos Then _
                .Stats.UserAtributos(eAtributos.Agilidad) = ConstantesBalance.MaxAtributos
            If .Stats.UserAtributos(eAtributos.Agilidad) > 2 * .Stats.UserAtributosBackUP(Agilidad) Then .Stats.UserAtributos(eAtributos.Agilidad) = 2 * .Stats.UserAtributosBackUP(Agilidad)
            
            .Stats.UserAtributos(eAtributos.Fuerza) = .Stats.UserAtributos(eAtributos.Fuerza) + Rnd
            If .Stats.UserAtributos(eAtributos.Fuerza) > ConstantesBalance.MaxAtributos Then _
                .Stats.UserAtributos(eAtributos.Fuerza) = ConstantesBalance.MaxAtributos
            If .Stats.UserAtributos(eAtributos.Fuerza) > 2 * .Stats.UserAtributosBackUP(Fuerza) Then .Stats.UserAtributos(eAtributos.Fuerza) = 2 * .Stats.UserAtributosBackUP(Fuerza)
            
            Call WriteUpdateStrenghtAndDexterity(UserIndex)
            
        End If
        
        Call WriteConsoleMsg(UserIndex, "Te has regenerado.", FontTypeNames.FONTTYPE_INFO)
        
        ' Enable the Berzerk if needed
        If BerzerkConditionMet(UserIndex) Then
            Call ActivatePassive(UserIndex, ePassiveSpells.Berserk, True)
            Call SendBerserkEffect(UserIndex, ePassiveSpells.Berserk, True)
        End If
        
        EnviarStats = True
 
    End With

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub DoRegeneration de General.bas")
End Sub

Public Sub CargaNpcsDat()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  
    
    Dim LoopC As Long
    Dim ln As String
    Dim tmpValue As Integer
    Dim Leer As clsIniManager
    Set Leer = New clsIniManager
    Call Leer.Initialize(DatPath & "NPCs.dat")
    
    'Dim NumNpcs As Long
    NumNpcsDat = CInt(Val(Leer.GetValue("INIT", "NumNPCs")))
    
    If NumNpcsDat <> 0 Then _
        ReDim NpcData(1 To NumNpcsDat)
    
    Dim NpcNumber As Long
    For NpcNumber = 1 To NumNpcsDat
       
        If Leer.KeyExists("NPC" & NpcNumber) Then
            With NpcData(NpcNumber)
                .Name = Leer.GetValue("NPC" & NpcNumber, "Name")
                .desc = Leer.GetValue("NPC" & NpcNumber, "Desc")
                .Tag = Leer.GetValue("NPC" & NpcNumber, "Tag")
                
                .Movement = Val(Leer.GetValue("NPC" & NpcNumber, "Movement"))
                .PathFinding = Val(Leer.GetValue("NPC" & NpcNumber, "PathFinding"))
                
                .NumInvocaciones = Val(Leer.GetValue("NPC" & NpcNumber, "NumInvocaciones"))
                If .NumInvocaciones > 0 Then
                    ReDim .NpcsInvocables(1 To .NumInvocaciones) As Integer
                    For LoopC = 1 To .NumInvocaciones
                        .NpcsInvocables(LoopC) = Val(Leer.GetValue("NPC" & NpcNumber, "Invocacion" & LoopC))
                    Next LoopC
                End If
                .flags.MaxInvocaciones = Val(Leer.GetValue("NPC" & NpcNumber, "MaxInvocaciones"))
                
                .ExtraBodies = Val(Leer.GetValue("NPC" & NpcNumber, "ExtraBodies"))
                If .ExtraBodies > 0 Then
                    ReDim .ExtraBody(1 To .ExtraBodies) As Integer
                    For LoopC = 1 To .ExtraBodies
                        .ExtraBody(LoopC) = Val(Leer.GetValue("NPC" & NpcNumber, "ExtraBody" & LoopC))
                    Next LoopC
                End If
         
                .flags.AguaValida = Val(Leer.GetValue("NPC" & NpcNumber, "AguaValida"))
                .flags.TierraInvalida = Val(Leer.GetValue("NPC" & NpcNumber, "TierraInValida"))
                .flags.Faccion = Val(Leer.GetValue("NPC" & NpcNumber, "Faccion"))
                .flags.AtacaDoble = Val(Leer.GetValue("NPC" & NpcNumber, "AtacaDoble"))
                
                .NPCtype = Val(Leer.GetValue("NPC" & NpcNumber, "NpcType"))
                
                .Char.body = Val(Leer.GetValue("NPC" & NpcNumber, "Body"))
                .Char.head = Val(Leer.GetValue("NPC" & NpcNumber, "Head"))
                .Char.heading = Val(Leer.GetValue("NPC" & NpcNumber, "Heading"))
                
                tmpValue = Val(Leer.GetValue("NPC" & NpcNumber, "EquippedShield"))
                If tmpValue > 0 Then
                    .Char.ShieldAnim = ObjData(tmpValue).ShieldAnim
                End If
                
                tmpValue = Val(Leer.GetValue("NPC" & NpcNumber, "EquippedHelmet"))
                If tmpValue > 0 Then
                    .Char.CascoAnim = ObjData(tmpValue).CascoAnim
                End If
                
                tmpValue = Val(Leer.GetValue("NPC" & NpcNumber, "EquippedWeapon"))
                If tmpValue > 0 Then
                    .Char.WeaponAnim = ObjData(tmpValue).WeaponAnim
                End If

                .Attackable = Val(Leer.GetValue("NPC" & NpcNumber, "Attackable"))
                .Comercia = Val(Leer.GetValue("NPC" & NpcNumber, "Comercia"))
                .Hostile = Val(Leer.GetValue("NPC" & NpcNumber, "Hostile"))
                
                .GiveEXP = Val(Leer.GetValue("NPC" & NpcNumber, "GiveEXP")) ' * 20
                .GiveEXPTierra = Val(Leer.GetValue("NPC" & NpcNumber, "GiveEXPTierra"))
                
                If .GiveEXPTierra = 0 Then .GiveEXPTierra = .GiveEXP ' / 20
                
                .Veneno = Val(Leer.GetValue("NPC" & NpcNumber, "Veneno"))
                
                .flags.Domable = Val(Leer.GetValue("NPC" & NpcNumber, "Domable"))
                .flags.ItemToTame = Val(Leer.GetValue("NPC" & NpcNumber, "ItemToTame"))

                .GiveGLD = Val(Leer.GetValue("NPC" & NpcNumber, "GiveGLD"))
                
                ' Load quests
                Dim NumQuests As Integer
                NumQuests = Val(Leer.GetValue("NPC" & NpcNumber, "NumQuests"))
                .NumQuests = NumQuests
                
                If NumQuests > 0 Then
                    ReDim .Quest(1 To NumQuests) As Integer
                    
                    For LoopC = 1 To NumQuests
                        .Quest(LoopC) = Val(Leer.GetValue("NPC" & NpcNumber, "Quest" & LoopC))
                    Next LoopC
                End If
        
                
                .PoderAtaque = Val(Leer.GetValue("NPC" & NpcNumber, "PoderAtaque"))
                .PoderEvasion = Val(Leer.GetValue("NPC" & NpcNumber, "PoderEvasion"))
                
                .InvReSpawn = Val(Leer.GetValue("NPC" & NpcNumber, "InvReSpawn"))
                
                With .Stats
                    .MaxHp = Val(Leer.GetValue("NPC" & NpcNumber, "MaxHP"))
                    .MinHp = Val(Leer.GetValue("NPC" & NpcNumber, "MinHP"))
                    .MaxHit = Val(Leer.GetValue("NPC" & NpcNumber, "MaxHIT"))
                    .MinHit = Val(Leer.GetValue("NPC" & NpcNumber, "MinHIT"))
                    .Def = Val(Leer.GetValue("NPC" & NpcNumber, "DEF"))
                    .DefM = Val(Leer.GetValue("NPC" & NpcNumber, "DEFm"))
                    .Alineacion = Val(Leer.GetValue("NPC" & NpcNumber, "Alineacion"))
                End With
                
                .Invent.NroItems = Val(Leer.GetValue("NPC" & NpcNumber, "NROITEMS"))
                For LoopC = 1 To .Invent.NroItems
                    ln = Leer.GetValue("NPC" & NpcNumber, "Obj" & LoopC)
                    .Invent.Object(LoopC).ObjIndex = Val(ReadField(1, ln, 45))
                    .Invent.Object(LoopC).Amount = Val(ReadField(2, ln, 45))
                Next LoopC
                
                .NroDrops = Val(Leer.GetValue("NPC" & NpcNumber, "NroDrops"))
                If .NroDrops > 0 Then
                    ReDim .Drop(1 To .NroDrops) As tDrops
                    For LoopC = 1 To .NroDrops
                        ln = Leer.GetValue("NPC" & NpcNumber, "Drop" & LoopC)
                        .Drop(LoopC).DropIndex = Val(ReadField(1, ln, 45))
                        .Drop(LoopC).Probabilidad = Val(ReadField(2, ln, 45))
                        .Drop(LoopC).NoExcluyente = Val(ReadField(3, ln, 45))
                    Next LoopC
                End If
                
                With .Intervalos
                    .Walk = Val(Leer.GetValue("NPC" & NpcNumber, "IntWalk"))
                    .Hit = Val(Leer.GetValue("NPC" & NpcNumber, "IntHit"))
                End With
                
                .flags.LanzaSpells = Val(Leer.GetValue("NPC" & NpcNumber, "LanzaSpells"))
                If .flags.LanzaSpells > 0 Then ReDim .Spells(1 To .flags.LanzaSpells)
                For LoopC = 1 To .flags.LanzaSpells
                    .Spells(LoopC) = Val(Leer.GetValue("NPC" & NpcNumber, "Sp" & LoopC))
                Next LoopC
                
                .NroCriaturas = Val(Leer.GetValue("NPC" & NpcNumber, "NroCriaturas"))
                If .NPCtype = eNPCType.Entrenador And .NroCriaturas > 0 Then
                    
                    ReDim .Criaturas(1 To .NroCriaturas) As tCriaturasEntrenador
                    For LoopC = 1 To .NroCriaturas
                        .Criaturas(LoopC).NpcIndex = Leer.GetValue("NPC" & NpcNumber, "CI" & LoopC)
                        .Criaturas(LoopC).NpcName = Leer.GetValue("NPC" & NpcNumber, "CN" & LoopC)
                    Next LoopC
                End If
                
                With .flags
                    .Respawn = Val(Leer.GetValue("NPC" & NpcNumber, "ReSpawn"))
                    
                    .ShowName = Val(Leer.GetValue("NPC" & NpcNumber, "ShowName"))
                    
                    .BackUp = Val(Leer.GetValue("NPC" & NpcNumber, "BackUp"))
                    .RespawnOrigPos = Val(Leer.GetValue("NPC" & NpcNumber, "OrigPos"))
                    .AfectaParalisis = Val(Leer.GetValue("NPC" & NpcNumber, "AfectaParalisis"))
                    
                    .DistanciaMaxima = Val(Leer.GetValue("NPC" & NpcNumber, "DistanciaMaxima"))
                    
                    .Snd1 = Val(Leer.GetValue("NPC" & NpcNumber, "Snd1"))
                    .Snd2 = Val(Leer.GetValue("NPC" & NpcNumber, "Snd2"))
                    .Snd3 = Val(Leer.GetValue("NPC" & NpcNumber, "Snd3"))
                End With
                
                '<<<<<<<<<<<<<< Expresiones >>>>>>>>>>>>>>>>
                .NroExpresiones = Val(Leer.GetValue("NPC" & NpcNumber, "NROEXP"))
                If .NroExpresiones > 0 Then ReDim .Expresiones(1 To .NroExpresiones) As String
                For LoopC = 1 To .NroExpresiones
                    .Expresiones(LoopC) = Leer.GetValue("NPC" & NpcNumber, "Exp" & LoopC)
                Next LoopC
                '<<<<<<<<<<<<<< Expresiones >>>>>>>>>>>>>>>>
                
                ' Menu desplegable p/npc
                Select Case .NPCtype
                    Case eNPCType.Banquero
                        .MenuIndex = eMenues.ieBanquero
                        
                    Case eNPCType.Entrenador
                        .MenuIndex = eMenues.ieEntrenador
                        
                    Case eNPCType.Gobernador
                        .MenuIndex = eMenues.ieGobernador
                        
                    Case eNPCType.Noble
                        .MenuIndex = eMenues.ieEnlistadorFaccion
                        
                    Case eNPCType.ResucitadorNewbie, eNPCType.Revividor
                        .MenuIndex = eMenues.ieSacerdote
                        
                    Case eNPCType.Timbero
                        .MenuIndex = eMenues.ieApostador
                        
                    Case Else
                        If .flags.Domable <> 0 Then
                            If .flags.Follow Then
                                .MenuIndex = eMenues.ieMascota
                            Else
                                .MenuIndex = eMenues.ieNpcDomable
                            End If
                        End If
                End Select
                
                If .Comercia = 1 Then .MenuIndex = eMenues.ieComerciante
                
                'Tipo de items con los que comercia
                .TipoItems = Val(Leer.GetValue("NPC" & NpcNumber, "TipoItems"))
                
                .Ciudad = Val(Leer.GetValue("NPC" & NpcNumber, "Ciudad"))
                
                .Exists = True
                .level = Val(Leer.GetValue("NPC" & NpcNumber, "Level"))
                .OffsetReducedExp = Val(Leer.GetValue("NPC" & NpcNumber, "OffsetReducedExp"))
                .OffsetModificator = Val(Leer.GetValue("NPC" & NpcNumber, "OffsetModificator"))
                
                .MasteryStarter = CBool(Val(Leer.GetValue("NPC" & NpcNumber, "MasteryStarter")))
                
                .OverHeadIcon = Val(Leer.GetValue("NPC" & NpcNumber, "OverHeadIcon"))
                
                
                ' Default values for width and height
                .SizeWidth = ModAreas.DEFAULT_ENTITY_WIDTH
                .SizeHeight = ModAreas.DEFAULT_ENTITY_HEIGHT
                
                If Leer.GetValue("NPC" & NpcNumber, "SizeWidth") <> "" Then
                    .SizeWidth = CByte(Val(Leer.GetValue("NPC" & NpcNumber, "SizeWidth")))
                End If
                
                If Leer.GetValue("NPC" & NpcNumber, "SizeHeight") <> "" Then
                    .SizeHeight = CByte(Val(Leer.GetValue("NPC" & NpcNumber, "SizeHeight")))
                End If

            End With
        End If
    Next NpcNumber
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") cargando NPC " & NpcNumber & " en Sub CargaNpcsDat de General.bas")
End Sub

Sub PasarSegundo()
'***************************************************
'Author: Unknown
'Last Modification: 18/06/2014
'18/06/2014: D'Artagnan - Show the account form when logged out.
'***************************************************

On Error GoTo ErrHandler
    Dim I As Long
    
    For I = 1 To LastUser
        With UserList(I)
            If .flags.UserLogged Then
                'Cerrar usuario
                If .Counters.Saliendo Then
                    .Counters.Salir = .Counters.Salir - 1
                    If .Counters.Salir <= 0 Then
                        Call WriteConsoleMsg(I, "Gracias por jugar Argentum Online", FontTypeNames.FONTTYPE_INFO)
                        Call WriteDisconnect(I)
                        
                        If .bShowAccountForm Then
                            ' Save user first to update data.
                            'If .PartyIndex > 0 Then Call mdParty.SalirDeParty(I)
                            'Call SaveUserDB(I, True, False, vbNullString)
                            ' Show the account form when logged out.
                            Call modAccount.connect(I, .AccountName, vbNullString, vbNullString, False, .ClientTempCode)
                        ElseIf .bForceCloseAccount Then
                            Call WriteLoginScreenShow(I)
                        End If
                                         
                        Call CloseSocket(I, True)
                    End If
                End If
            End If
        End With
    Next I
    
Exit Sub

ErrHandler:
    Call LogError("Error en PasarSegundo. Err: " & Err.Description & " - " & Err.Number & " - UserIndex: " & I)
    Resume Next
End Sub

Public Function ReiniciarAutoUpdate() As Double
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  
    ReiniciarAutoUpdate = Shell(App.Path & "\autoupdater\aoau.exe", vbMinimizedNoFocus)
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function ReiniciarAutoUpdate de General.bas")
End Function
 
Public Sub ReiniciarServidor(Optional ByVal EjecutarLauncher As Boolean = True)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    'WorldSave
    Call ES.DoBackUp

    'commit experiencias
    Call mdParty.ActualizaExperiencias

    'Guardar Pjs
    Call GuardarUsuarios
    
    If EjecutarLauncher Then Shell (App.Path & "\launcher.exe")

    'Chauuu
    Unload frmMain

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ReiniciarServidor de General.bas")
End Sub

 
Sub GuardarUsuarios()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    haciendoBK = True
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Grabando Personajes", FontTypeNames.FONTTYPE_SERVER, eMessageType.Admin), IsUrgent:=True)
    
    Dim I As Long
    For I = 1 To LastUser
        If UserList(I).flags.UserLogged Then
            Call SaveUserDB(I, False, False, vbNullString)
        End If
    Next I
    
    'se guardan los seguimientos
    Call SaveRecords
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Personajes Grabados", FontTypeNames.FONTTYPE_SERVER, eMessageType.Admin))
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle(), IsUrgent:=True)

    haciendoBK = False
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub GuardarUsuarios de General.bas")
End Sub


Sub InicializaEstadisticas()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    Dim Ta As Long
    Ta = GetTickCount()
    
    Set EstadisticasWeb = New clsEstadisticasIPC
    Call EstadisticasWeb.Inicializa(frmMain.hWnd)
    Call EstadisticasWeb.Informar(CANTIDAD_MAPAS, NumMaps)
    Call EstadisticasWeb.Informar(CANTIDAD_ONLINE, NumUsers)
    Call EstadisticasWeb.Informar(UPTIME_SERVER, (Ta - tInicioServer) / 1000)
    Call EstadisticasWeb.Informar(RECORD_USUARIOS, RECORDusuarios)

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub InicializaEstadisticas de General.bas")
End Sub

Public Sub FreeNPCs()
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Releases all NPC Indexes
'***************************************************
On Error GoTo ErrHandler
  
    Dim LoopC As Long
    
    ' Free all NPC indexes
    For LoopC = 1 To MAXNPCS
        Npclist(LoopC).flags.NPCActive = False
    Next LoopC
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub FreeNPCs de General.bas")
End Sub

Public Sub FreeCharIndexes()
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Releases all char indexes
'***************************************************
    ' Free all char indexes (set them all to 0)
On Error GoTo ErrHandler
  
    Call ZeroMemory(charList(1), MAXCHARS * Len(charList(1)))
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub FreeCharIndexes de General.bas")
End Sub

Public Function isTrading(ByVal UserIndex As Integer) As Boolean
'***************************************************
'Autor: Torres Patricio (Pato)
'Last Modification: 05/04/12
'
'***************************************************
On Error GoTo ErrHandler
  
    
    isTrading = (UserList(UserIndex).flags.Comerciando <> 0)
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function isTrading de General.bas")
End Function


Public Function isTradingWithUser(ByVal UserIndex As Integer) As Boolean
'***************************************************
'Autor: Torres Patricio (Pato)
'Last Modification: 05/04/12
'
'***************************************************
On Error GoTo ErrHandler
  
    
    isTradingWithUser = (UserList(UserIndex).flags.Comerciando < 0)
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function isTradingWithUser de General.bas")
End Function

Public Function isTradingWithNPC(ByVal UserIndex As Integer) As Boolean
'***************************************************
'Autor: Torres Patricio (Pato)
'Last Modification: 05/04/12
'
'***************************************************
On Error GoTo ErrHandler
  
    
    isTradingWithNPC = (UserList(UserIndex).flags.Comerciando > 0)
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function isTradingWithNPC de General.bas")
End Function

Public Function isTradingWithBank(ByVal UserIndex As Integer) As Boolean
'***************************************************
'Autor: Torres Patricio (Pato)
'Last Modification: 05/04/12
'
'***************************************************
On Error GoTo ErrHandler
  
    
    isTradingWithBank = (UserList(UserIndex).flags.Comerciando = (MAXNPCS + 1))
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function isTradingWithBank de General.bas")
End Function

Public Function getTradingUser(ByVal UserIndex As Integer) As Integer
'***************************************************
'Autor: Torres Patricio (Pato)
'Last Modification: 05/04/12
'
'***************************************************
On Error GoTo ErrHandler
  
    
    getTradingUser = (Not UserList(UserIndex).flags.Comerciando)
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function getTradingUser de General.bas")
End Function

Public Function getTradingNPC(ByVal UserIndex As Integer) As Integer
'***************************************************
'Autor: Torres Patricio (Pato)
'Last Modification: 05/04/12
'
'***************************************************
On Error GoTo ErrHandler
  
    
    getTradingNPC = UserList(UserIndex).flags.Comerciando
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function getTradingNPC de General.bas")
End Function

Public Function checkCanUseItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean
    
    If ObjIndex <= 0 Then
        checkCanUseItem = True
        Exit Function
    End If
    
    With ObjData(ObjIndex)
        If Not .ObjType = otAnillo And .ObjType = otArmadura And .ObjType = otBarcos And .ObjType = otCASCO _
                                        And .ObjType = otESCUDO And .ObjType = otFlechas And .ObjType = otWeapon Then
            checkCanUseItem = True
            Exit Function
        End If
        
        If .MinimumLevel > UserList(UserIndex).Stats.ELV Then Exit Function
        
        If Not ClasePuedeUsarItem(UserIndex, ObjIndex) Then Exit Function
        If Not SexoPuedeUsarItem(UserIndex, ObjIndex) Then Exit Function
        If .ObjType = otArmadura Then: If Not CheckRazaUsaRopa(UserIndex, ObjIndex) Then Exit Function
        If Not FaccionPuedeUsarItem(UserIndex, ObjIndex) Then Exit Function
    End With
    
    checkCanUseItem = True
End Function

Public Sub InitSessionSystem()
    If ServerConfiguration.Session.Lifetime <= 0 Or ServerConfiguration.Session.MaxQuantity <= 0 Or ServerConfiguration.Session.TokenSize <= 1 Then
        Call MsgBox("La configuración del sistema de sesiones de cuenta posee valores invalidos")
        End
    End If

    Call modSession.InitializeSessionSystem
End Sub

Public Function GetTickCount() As Long
    GetTickCount = GetRealTickCount() And &H7FFFFFFF
End Function

Public Function GetUserIndexFromUserId(ByVal UserId As Long) As Integer
On Error GoTo ErrHandler
    Dim currentUserIndex As Integer
    
    For currentUserIndex = 1 To LastUser
        If UserList(currentUserIndex).ID = UserId Then
            GetUserIndexFromUserId = currentUserIndex
            Exit Function
        End If
    Next currentUserIndex
    
    Exit Function
    
ErrHandler:
    Call LogError("Error " & Err.Number & " (" & Err.Description & ") en Function GetUserIndexFromUserId del Módulo General")
End Function

Public Sub StartShutDown(from As ShutDownFrom, withBackup As Boolean)
    Call LogMain("Starting shutdown from " & from & " with backup:" & withBackup)
    
    ShutdownWithBackup = withBackup
    ShutdownBy = from

    Running = False
End Sub

Private Sub MakeServerBackup()
On Error GoTo ErrHandler
    FrmStat.Show
   
    'WorldSave
    Call ES.DoBackUp

    'commit experiencia
    Call mdParty.ActualizaExperiencias

    'Guardar Pjs
    Call GuardarUsuarios
    
    ' Save Guild Headers
    Call modGuild_Functions.SaveAllGuilds
    
 Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub MakeServerBackup de General.frm")

End Sub

Private Sub CloseServer()
On Error Resume Next
On Error GoTo ErrHandler
    Dim LogMessage As String
    
    If ShutdownWithBackup Then
        Call MakeServerBackup
    End If

    'Save stats!!!
    Call Statistics.DumpStatistics

    Call frmMain.QuitarIconoSystray

    Dim LoopC As Integer

    For LoopC = 1 To MaxUsers
        If UserList(LoopC).ConnIDValida Then Call CloseSocket(LoopC)
    Next

    Dim N As Integer
    'Dejamos en 0 el contador de numusers.log
    N = FreeFile
    Open ServerConfiguration.LogsPaths.GeneralPath & "numusers.log" For Output As N
        Print #N, 0
    Close #N
        
    'Log
    LogMessage = "Server cerrado."
    
    Call LogMain(LogMessage)
    
    Set SonidosMapas = Nothing

    End
    
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CloseServer de General.bas")
End Sub

Public Sub OutputToLog(FileName As String, Message As String)
    Dim N As Integer
    N = FreeFile
    Open ServerConfiguration.LogsPaths.GeneralPath & FileName For Append Shared As #N
        Print #N, Date & " " & Time & " " & Message
    Close #N
End Sub

Public Sub LogMain(Message As String)
    Call OutputToLog("Main.log", Message)
End Sub

Public Function GetClassTypeFromName(ByRef ClassName As String) As Byte
    Dim I As Byte
    
    For I = 1 To NUMCLASES
        If UCase$(ListaClases(I)) = UCase$(ClassName) Then
            GetClassTypeFromName = I
            Exit Function
        End If
    Next I
    
    GetClassTypeFromName = 0
    
End Function

Public Function RandomString(ByVal cb As Integer) As String
On Error GoTo ErrHandler
  
    Randomize
    Dim rgch As String
    rgch = "abcdefghijklmnopqrstuvwxyz"
    rgch = rgch & UCase(rgch) & "0123456789"

    Dim I As Long
    For I = 1 To cb
        RandomString = RandomString & mid$(rgch, Int(Rnd() * Len(rgch) + 1), 1)
    Next

  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function RandomString de General.bas")
End Function


Public Function GetRandomCharacterStartPosition() As WorldPos
    Dim RandNumb As Integer
    If ServerConfiguration.StartPositionsQty <= 0 Then
        GetRandomCharacterStartPosition = Nemahuak
        Exit Function
    End If
    
    
    RandNumb = RandomNumber(1, ServerConfiguration.StartPositionsQty)
    GetRandomCharacterStartPosition = ServerConfiguration.StartPositions(RandNumb)

End Function
