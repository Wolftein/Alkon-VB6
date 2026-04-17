Attribute VB_Name = "Admin"
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

Public Type tMotd
    texto As String
    Formato As String
End Type

Public MaxLines As Integer
Public MOTD() As tMotd

Public Type tAPuestas
    Ganancias As Long
    Perdidas As Long
    Jugadas As Long
End Type
Public Apuestas As tAPuestas

Public tInicioServer As Long
Public EstadisticasWeb As clsEstadisticasIPC



'BALANCE

Public PorcentajeRecuperoMana As Integer

Public MinutosWs As Long
Public MinutosGuardarUsuarios As Long
Public MinutosMotd As Long
Public Puerto As Integer

Public BootDelBackUp As Byte
Public Lloviendo As Boolean

Function VersionOK(ByVal Ver As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  
    VersionOK = (Ver = ULTIMAVERSION)
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function VersionOK de Admin.bas")
End Function

Sub ReSpawnOrigPosNpcs()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

On Error Resume Next

    Dim I As Integer
    Dim MiNPC As npc
       
    For I = 1 To LastNPC
       'OJO
       If Npclist(I).flags.NPCActive Then
            
            If InMapBounds(Npclist(I).Orig.Map, Npclist(I).Orig.X, Npclist(I).Orig.Y) And Npclist(I).Numero = Guardias Then
                    MiNPC = Npclist(I)
                    Call QuitarNPC(I)
                    Call ReSpawnNpc(MiNPC)
            End If
            
            'tildada por sugerencia de yind
            'If Npclist(I).Contadores.TiempoExistencia > 0 Then
            '        Call MuereNpc(I, 0)
            'End If
       End If
       
    Next I
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ReSpawnOrigPosNpcs de Admin.bas")
End Sub

Sub WorldSave()
'***************************************************
'Author: Unknown
'Last Modification: 12/10/2014
'12/10/2014: D'Artagnan - Create backup folder.
'***************************************************
On Error GoTo ErrHandler
  

On Error Resume Next

    Dim loopX As Integer
    Dim hFile As Integer
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Iniciando WorldSave", FontTypeNames.FONTTYPE_SERVER, eMessageType.Admin), IsUrgent:=True)
    
    If Not FileExist(ServerConfiguration.ResourcesPaths.WorldBackup, vbDirectory) Then
        Call MkDir(ServerConfiguration.ResourcesPaths.WorldBackup)
    End If
    
    #If EnableSecurity Then
        Encriptacion.StringValidacion = Encriptacion.ArmarStringValidacion
    #End If
    
    Call ReSpawnOrigPosNpcs 'respawn de los guardias en las pos originales
    
    Dim J As Integer, k As Integer
    
    For J = 1 To NumMaps
        If MapInfo(J).BackUp = 1 Then k = k + 1
    Next J
    
    FrmStat.ProgressBar1.Min = 0
    FrmStat.ProgressBar1.Max = k
    FrmStat.ProgressBar1.value = 0
    
    For loopX = 1 To NumMaps
        'DoEvents
        
        If MapInfo(loopX).BackUp = 1 Then
            Call GrabarMapa(loopX, ServerConfiguration.ResourcesPaths.WorldBackup & "Mapa" & loopX)
            FrmStat.ProgressBar1.value = FrmStat.ProgressBar1.value + 1
        End If
    
    Next loopX
    
    FrmStat.Visible = False
    
    If FileExist(DatPath & "\bkNpcs.dat") Then Kill (DatPath & "bkNpcs.dat")
    
    hFile = FreeFile()
    
    Open DatPath & "\bkNpcs.dat" For Output As hFile
    
        For loopX = 1 To LastNPC
            If Npclist(loopX).flags.BackUp = 1 Then
                Call BackUPnPc(loopX, hFile)
            End If
        Next loopX
        
    Close hFile
    
    Call SaveForums
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> WorldSave ha concluído.", FontTypeNames.FONTTYPE_SERVER, eMessageType.Admin))
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WorldSave de Admin.bas")
End Sub

Public Sub PurgarPenas()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    Dim I As Long
    
    For I = 1 To LastUser
        If UserList(I).flags.UserLogged Then
            If UserList(I).Counters.Pena > 0 Then
                UserList(I).Counters.Pena = UserList(I).Counters.Pena - 1
                
                If UserList(I).Counters.Pena < 1 Then
                    UserList(I).Counters.Pena = 0
                    Call WarpUserChar(I, Libertad.Map, Libertad.X, Libertad.Y, True)
                    Call WriteConsoleMsg(I, "¡Has sido liberado!", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
    Next I
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub PurgarPenas de Admin.bas")
End Sub


Public Sub Encarcelar(ByVal UserIndex As Integer, ByVal Minutos As Long, Optional ByVal GmName As String = vbNullString)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    UserList(UserIndex).Counters.Pena = Minutos
    
    
    Call WarpUserChar(UserIndex, Prision.Map, Prision.X, Prision.Y, True)
    
    If LenB(GmName) = 0 Then
        Call WriteConsoleMsg(UserIndex, "Has sido encarcelado, deberás permanecer en la cárcel " & Minutos & " minutos.", FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(UserIndex, GmName & " te ha encarcelado, deberás permanecer en la cárcel " & Minutos & " minutos.", FontTypeNames.FONTTYPE_INFO)
    End If
    If UserList(UserIndex).flags.Traveling = 1 Then
        Call EndTravel(UserIndex, True)
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Encarcelar de Admin.bas")
End Sub

Public Function PersonajeExiste(ByVal Name As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

  PersonajeExiste = GetUserID(Name) <> 0


  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function PersonajeExiste de Admin.bas")
End Function

Public Function MD5ok(ByVal md5formateado As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler

    Dim I As Integer
    
    If MD5ClientesActivado = 1 Then
        For I = 0 To UBound(MD5s)
            If (md5formateado = MD5s(I)) Then
                MD5ok = True
                Exit Function
            End If
        Next I
        MD5ok = False
    Else
        MD5ok = True
    End If

  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function MD5ok de Admin.bas")
End Function

Public Sub MD5sCarga()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler

    Dim LoopC As Integer
    
    MD5ClientesActivado = Val(GetVar(IniPath & "Server.ini", "MD5Hush", "Activado"))
    
    If MD5ClientesActivado = 1 Then
        ReDim MD5s(Val(GetVar(IniPath & "Server.ini", "MD5Hush", "MD5Aceptados")))
        For LoopC = 0 To UBound(MD5s)
            MD5s(LoopC) = GetVar(IniPath & "Server.ini", "MD5Hush", "MD5Aceptado" & (LoopC + 1))
            MD5s(LoopC) = txtOffset(hexMd52Asc(MD5s(LoopC)), 55)
        Next LoopC
    End If

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub MD5sCarga de Admin.bas")
End Sub

Public Sub BanIpAgrega(ByVal IP As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler

    BanIps.Add IP
    
    Call BanIpGuardar
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub BanIpAgrega de Admin.bas")
End Sub

Public Function BanIpBuscar(ByVal IP As String) As Long
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler

    Dim Dale As Boolean
    Dim LoopC As Long
    
    Dale = True
    LoopC = 1
    Do While LoopC <= BanIps.Count And Dale
        Dale = (BanIps.Item(LoopC) <> IP)
        LoopC = LoopC + 1
    Loop
    
    If Dale Then
        BanIpBuscar = 0
    Else
        BanIpBuscar = LoopC - 1
    End If
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function BanIpBuscar de Admin.bas")
End Function

Public Function BanIpQuita(ByVal IP As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler

On Error Resume Next

    Dim N As Long
    
    N = BanIpBuscar(IP)
    If N > 0 Then
        BanIps.Remove N
        BanIpGuardar
        BanIpQuita = True
    Else
        BanIpQuita = False
    End If

  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function BanIpQuita de Admin.bas")
End Function

Public Sub BanIpGuardar()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    Dim ArchivoBanIp As String
    Dim ArchN As Long
    Dim LoopC As Long
    
    ArchivoBanIp = ServerConfiguration.ResourcesPaths.Dats & "BanIps.dat"
    
    ArchN = FreeFile()
    Open ArchivoBanIp For Output As #ArchN
    
    For LoopC = 1 To BanIps.Count
        Print #ArchN, BanIps.Item(LoopC)
    Next LoopC
    
    Close #ArchN

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub BanIpGuardar de Admin.bas")
End Sub

Public Sub BanIpCargar()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    Dim ArchN As Long
    Dim Tmp As String
    Dim ArchivoBanIp As String
    
    ArchivoBanIp = ServerConfiguration.ResourcesPaths.Dats & "BanIps.dat"
    
    Set BanIps = New Collection
    
    ArchN = FreeFile()
    Open ArchivoBanIp For Input As #ArchN
    
    Do While Not EOF(ArchN)
        Line Input #ArchN, Tmp
        BanIps.Add Tmp
    Loop
    
    Close #ArchN

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub BanIpCargar de Admin.bas")
End Sub

Public Sub ActualizaEstadisticasWeb()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    Static Andando As Boolean
    Static Contador As Long
    Dim Tmp As Boolean
    
    Contador = Contador + 1
    
    If Contador >= 10 Then
        Contador = 0
        Tmp = EstadisticasWeb.EstadisticasAndando()
        
        If Andando = False And Tmp = True Then
            Call InicializaEstadisticas
        End If
        
        Andando = Tmp
    End If

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ActualizaEstadisticasWeb de Admin.bas")
End Sub

Public Function UserDarPrivilegioLevel(ByVal Name As String) As PlayerType
'***************************************************
'Author: Unknown
'Last Modification: 03/02/07
'Last Modified By: Juan Martín Sotuyo Dodero (Maraxus)
'***************************************************
On Error GoTo ErrHandler
  

    If EsAdmin(Name) Then
        UserDarPrivilegioLevel = PlayerType.Admin
    ElseIf EsDios(Name) Then
        UserDarPrivilegioLevel = PlayerType.Dios
    ElseIf EsSemiDios(Name) Then
        UserDarPrivilegioLevel = PlayerType.SemiDios
    ElseIf EsConsejero(Name) Then
        UserDarPrivilegioLevel = PlayerType.Consejero
    Else
        UserDarPrivilegioLevel = PlayerType.User
    End If
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function UserDarPrivilegioLevel de Admin.bas")
End Function


Public Sub BanCharacter(ByVal bannerUserIndex As Integer, ByVal UserName As String, ByVal Reason As String, ByRef AdminNotes As String, ByVal punishmentType As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 03/01/15
'22/05/2010: Ya no se peude banear admins de mayor rango si estan online.
'03/01/2015: Nightw - Hice algo de limpieza en el código y arreglé el bug que no guardaba las penas si el user estaba online
'***************************************************
On Error GoTo ErrHandler
  

    Dim bannedUserIndex As Integer
    Dim UserPriv As Byte
    Dim cantPenas As Byte
    Dim rank As Integer
    Dim UserId As Long
    Dim GuildId As Long
    Dim bUserBanned As Boolean
    Dim puedeBanear As Boolean
    Dim GuildIndex As Integer
    
    puedeBanear = False
    
    If InStrB(UserName, "+") Then
        UserName = Replace(UserName, "+", " ")
    End If
    
    bannedUserIndex = NameIndex(UserName)
    
    rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero

    If bannedUserIndex > 0 Then
        UserId = UserList(bannedUserIndex).ID
        GuildId = UserList(bannedUserIndex).Guild.IdGuild
        UserName = UserList(bannedUserIndex).Name
        GuildIndex = UserList(bannedUserIndex).Guild.GuildIndex
        
        bUserBanned = False
    Else
        Call GetCharInfoWithGuild(UserName, UserId, GuildId, bUserBanned)
        GuildIndex = GetGuildIndexByGuildId(GuildId)
    End If
    
    
    If UserId = 0 Then
        Call WriteConsoleMsg(bannerUserIndex, "El pj " & UserName & " no existe.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
   
    With UserList(bannerUserIndex)
        If bannedUserIndex <= 0 Then
            Call WriteConsoleMsg(bannerUserIndex, "El usuario no está online.", FontTypeNames.FONTTYPE_TALK)
            
            UserPriv = UserDarPrivilegioLevel(UserName)
            
            If (UserPriv And rank) > (.flags.Privilegios And rank) Then
                Call WriteConsoleMsg(bannerUserIndex, "No puedes banear a alguien de mayor jerarquía.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If

            If bUserBanned Then
                Call WriteConsoleMsg(bannerUserIndex, "El personaje ya se encuentra baneado.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            Call SendData(SendTarget.ToAdminsButCounselorsAndRms, 0, PrepareMessageConsoleMsg("Servidor> " & .Name & " ha baneado a " & UserName & ".", FontTypeNames.FONTTYPE_SERVER, eMessageType.Admin))
            puedeBanear = True
            
            If GuildIndex > 0 Then
                Call modGuild_Functions.GuildBanChar(GuildIndex, UserId)
            End If
            
            If (UserPriv And rank) = (.flags.Privilegios And rank) Then
                .flags.Ban = 1
                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.Name & " banned by the server por bannear un Administrador.", FontTypeNames.FONTTYPE_FIGHT))
                Call CloseSocket(bannerUserIndex)
            End If
 
        Else
            If (UserList(bannedUserIndex).flags.Privilegios And rank) > (.flags.Privilegios And rank) Then
                Call WriteConsoleMsg(bannerUserIndex, "No puedes banear a alguien de mayor jerarquía.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            puedeBanear = True
            
            Call LogBan(UserList(bannedUserIndex).Name, UserList(bannerUserIndex).Name, Reason)
            Call SendData(SendTarget.ToAdminsButCounselorsAndRms, 0, PrepareMessageConsoleMsg("Servidor> " & .Name & " ha baneado a " & UserList(bannedUserIndex).Name & ".", FontTypeNames.FONTTYPE_SERVER, eMessageType.Admin))

            'Ponemos el flag de ban a 1
            UserList(bannedUserIndex).flags.Ban = 1
            
            If GuildIndex > 0 Then
            Call modGuild_Functions.GuildBanChar(GuildIndex, UserId)
            End If
            
            If (UserList(bannedUserIndex).flags.Privilegios And rank) = (.flags.Privilegios And rank) Then
                .flags.Ban = 1
                
                Call SendData(SendTarget.ToAdminsButCounselorsAndRms, 0, PrepareMessageConsoleMsg(.Name & " banned by the server por bannear un Administrador.", FontTypeNames.FONTTYPE_FIGHT))
                Call CloseSocket(bannerUserIndex)
            End If
 
            Call CloseSocket(bannedUserIndex)
            
        End If
        
        Call LogGM(.Name, "BAN a " & UserName & ". Razón: " & Reason)
        
        If puedeBanear Then
        
            Call AddPunishmentDB(UserId, .ID, punishmentType, Reason, AdminNotes)
            
        End If

    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub BanCharacter de Admin.bas")
End Sub

