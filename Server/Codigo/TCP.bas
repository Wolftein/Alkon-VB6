Attribute VB_Name = "TCP"
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
'Cñdigo Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Public Server As Network_Server
Private ServerProtocol As Network_Protocol


Public Sub OnServerConnect(ByVal Connection As Network_Client)
    On Error GoTo ErrHandler

    Dim IP As String
    IP = Connection.GetStatistics().Address
    
    Dim IPLong As Long
    IPLong = GetLongIp(IP)
    
    'Busca si esta banneada la ip
    Dim I As Long
    For I = 1 To BanIps.Count
        If BanIps.Item(I) = IP Then
            Call SendAndClose(Connection, PrepareMessageErrorMsg("Su IP se encuentra bloqueada en este servidor."))
            Exit Sub
        End If
    Next I
    
    If ServerConfiguration.IpTablesSecurityEnabled Then
        If Not SecurityIp.IpSecurityAceptarNuevaConexion(IPLong) Then
            If ServerConfiguration.IpTablesSecurityLogFailedEnabled Then Call LogError("Failed connection: IpSecurityAceptarNuevaConexion - " & IP)
            
            Call Connection.Close(True)
            Exit Sub
        End If
        
        If SecurityIp.IPSecuritySuperaLimiteConexiones(IPLong) Then
            If ServerConfiguration.IpTablesSecurityLogFailedEnabled Then Call LogError("Failed connection: IPSecuritySuperaLimiteConexiones - " & IP)
            
            Call SendAndClose(Connection, PrepareMessageErrorMsg("Limite de conexiones para su IP alcanzado."))
            Exit Sub
        End If
    End If
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   BIENVENIDO AL SERVIDOR!!!!!!!!
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim UserIndex As Integer
    UserIndex = NextOpenUser
    
    If (UserIndex = 0) Then
        Call SendAndClose(Connection, PrepareMessageErrorMsg("El servidor se encuentra lleno en este momento. Disculpe las molestias ocasionadas."))
        Exit Sub
    End If

    If UserIndex > LastUser Then LastUser = UserIndex
    
    Set UserList(UserIndex).Connection = Connection
    UserList(UserIndex).ConnIDValida = True
    UserList(UserIndex).IP = IP
    UserList(UserIndex).IPLong = IPLong
    UserList(UserIndex).Counters.IdleCount = SetIntervalEnd(ServerConfiguration.Intervals.IntervaloIdleKick)
    Call Connection.SetAttachment(UserIndex)
    
    Call WriteConnectedMessage(UserIndex)

    Exit Sub
  
ErrHandler:

  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub OnServerConnect de TCP.bas")

End Sub

Public Sub OnServerClose(ByVal Connection As Network_Client)
On Error GoTo ErrHandler

    If ServerConfiguration.IpTablesSecurityEnabled Then
        Call SecurityIp.IpRestarConexion(GetLongIp(Connection.GetStatistics().Address))
    End If
    
    Dim UserIndex As Integer
    UserIndex = Connection.GetAttachment()
    
    'Rejected on accept
    If UserIndex <> 0 Then
        If UserList(UserIndex).flags.UserLogged Then
            Call Cerrar_Usuario(UserIndex)
        Else
            Call CloseSocket(UserIndex)
        End If
        
        UserList(UserIndex).ConnIDValida = False
        Set UserList(UserIndex).Connection = Nothing
    End If

    Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub OnServerClose de cSession.cls")
End Sub

Public Sub OnServerSend(ByVal Connection As Network_Client, ByVal Message As BinaryReader)
    ' #SECURITY
End Sub

Public Sub OnServerRecv(ByVal Connection As Network_Client, ByVal Message As BinaryReader)
    Set Reader = Message
    
    Dim UserIndex As Integer
    UserIndex = Connection.GetAttachment()
        
    If UserIndex <> 0 Then
        While (Message.GetAvailable() > 0)
            Call HandleIncomingData(UserIndex)
        Wend
    End If
  
    Set Reader = Nothing
End Sub

Public Sub OnServerError(ByVal Connection As Network_Client, ByVal Error As Long, ByVal Description As String)
    ' #SECURITY
End Sub

Public Sub Listen(ByVal Address As String, ByVal Service As String)
    Set ServerProtocol = New Network_Protocol
    Call ServerProtocol.Attach(AddressOf OnServerConnect, AddressOf OnServerClose, AddressOf OnServerRecv, AddressOf OnServerSend, AddressOf OnServerError)
    
    Set Server = Aurora_Network.Listen(Address, Service)
    Call Server.SetProtocol(ServerProtocol)
    
End Sub

Public Sub Disconnect()
    Set Server = Nothing
End Sub

Public Sub Send(ByVal Connection As Network_Client, Optional ByVal Urgent As Boolean = False)

    If Connection Is Nothing Then
        Exit Sub
    End If
    
    Call Connection.Write(Protocol.Writer)
    
    If (Urgent) Then
        Call Connection.Flush
    End If
End Sub

Public Sub SendAndClose(ByVal Connection As Network_Client, ByVal Blank As String)
    Call Connection.Write(Protocol.Writer)
    Call Connection.Close(False)
    
    Protocol.Writer.Clear
End Sub


Sub DarCuerpo(ByVal UserIndex As Integer)
'*************************************************
'Author: Nacho (Integer)
'Last modified: 14/03/2007
'Elije una cabeza para el usuario y le da un body
'*************************************************
On Error GoTo ErrHandler
  
Dim NewBody As Integer
Dim UserRaza As Byte
Dim UserGenero As Byte

UserGenero = UserList(UserIndex).Genero
UserRaza = UserList(UserIndex).raza

Select Case UserGenero
   Case eGenero.Hombre
        Select Case UserRaza
            Case eRaza.Humano
                NewBody = 1
            Case eRaza.Elfo
                NewBody = 2
            Case eRaza.Drow
                NewBody = 3
            Case eRaza.Enano
                NewBody = 300
            Case eRaza.Gnomo
                NewBody = 300
        End Select
   Case eGenero.Mujer
        Select Case UserRaza
            Case eRaza.Humano
                NewBody = 1
            Case eRaza.Elfo
                NewBody = 2
            Case eRaza.Drow
                NewBody = 3
            Case eRaza.Gnomo
                NewBody = 300
            Case eRaza.Enano
                NewBody = 300
        End Select
End Select

UserList(UserIndex).Char.body = NewBody
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub DarCuerpo de TCP.bas")
End Sub

Public Function ValidarCabeza(ByVal UserRaza As Byte, ByVal UserGenero As Byte, ByVal head As Integer) As Boolean
On Error GoTo ErrHandler
  

Select Case UserGenero
    Case eGenero.Hombre
        Select Case UserRaza
            Case eRaza.Humano
                ValidarCabeza = (head >= ConstantesGRH.HumanoHPrimerCabeza And _
                                head <= ConstantesGRH.HumanoHUltimaCabeza)
            Case eRaza.Elfo
                ValidarCabeza = (head >= ConstantesGRH.ElfoHPrimerCabeza And _
                                head <= ConstantesGRH.ElfoHUltimaCabeza)
            Case eRaza.Drow
                ValidarCabeza = (head >= ConstantesGRH.DrowHPrimerCabeza And _
                                head <= ConstantesGRH.DrowHUltimaCabeza)
            Case eRaza.Enano
                ValidarCabeza = (head >= ConstantesGRH.EnanoHPrimerCabeza And _
                                head <= ConstantesGRH.EnanoHUltimaCabeza)
            Case eRaza.Gnomo
                ValidarCabeza = (head >= ConstantesGRH.GnomoHPrimerCabeza And _
                                head <= ConstantesGRH.GnomoHUltimaCabeza)
        End Select
    
    Case eGenero.Mujer
        Select Case UserRaza
            Case eRaza.Humano
                ValidarCabeza = (head >= ConstantesGRH.HumanoMPrimerCabeza And _
                                head <= ConstantesGRH.HumanoMUltimaCabeza)
            Case eRaza.Elfo
                ValidarCabeza = (head >= ConstantesGRH.ElfoMPrimerCabeza And _
                                head <= ConstantesGRH.ElfoMUltimaCabeza)
            Case eRaza.Drow
                ValidarCabeza = (head >= ConstantesGRH.DrowMPrimerCabeza And _
                                head <= ConstantesGRH.DrowMUltimaCabeza)
            Case eRaza.Enano
                ValidarCabeza = (head >= ConstantesGRH.EnanoMPrimerCabeza And _
                                head <= ConstantesGRH.EnanoMUltimaCabeza)
            Case eRaza.Gnomo
                ValidarCabeza = (head >= ConstantesGRH.GnomoMPrimerCabeza And _
                                head <= ConstantesGRH.GnomoMUltimaCabeza)
        End Select
End Select
        
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function ValidarCabeza de TCP.bas")
End Function

Function AsciiValidos(ByVal cad As String, ByVal cEspec As Boolean, _
                      Optional ByVal bIncludeNumbers As Boolean = False, _
                      Optional ByVal bIncludeSpaces As Boolean = True) As Boolean
On Error GoTo ErrHandler
  
    Dim car As Byte
    Dim I As Long
    Dim J As Long
    
    cad = LCase$(cad)
    
    For I = 1 To Len(cad)
        AsciiValidos = False
        car = Asc(Mid$(cad, I, 1))
        If ((car >= 97 And car <= 122)) Or (bIncludeSpaces And car = 32) Or _
           (bIncludeNumbers And car >= 48 And car <= 57) Then
            AsciiValidos = True
        Else
            If cEspec Then 'Asciivalidos especiales
                For J = 1 To Len(CAR_ESPECIALES)
                    If car = Asc(Mid$(CAR_ESPECIALES, J, 1)) Then
                        AsciiValidos = True
                        Exit For
                    End If
                Next J
                If AsciiValidos = False Then Exit Function
            Else
                Exit Function
            End If
        End If
    Next I
    
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function AsciiValidos de TCP.bas")
End Function

Function NombrePermitido(ByVal Nombre As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

Dim I As Integer

For I = 1 To UBound(ForbidenNames)
    If InStr(Nombre, ForbidenNames(I)) Then
        NombrePermitido = False
        Exit Function
    End If
Next I

NombrePermitido = True

  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function NombrePermitido de TCP.bas")
End Function

Function ValidateSkills(ByVal UserIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

Dim LoopC As Integer

For LoopC = 1 To NUMSKILLS
    If GetSkills(UserIndex, LoopC) < 0 Then
        Exit Function
        If GetSkills(UserIndex, LoopC) > 100 Then
            Call ZeroSkills(UserIndex, LoopC)
            Call AddNaturalSkills(UserIndex, LoopC, 100)
        End If
    End If
Next LoopC

ValidateSkills = True
    
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function ValidateSkills de TCP.bas")
End Function

Public Function ValidCharacterData(ByVal UserIndex As Integer, ByRef Name As String, ByVal UserRaza As eRaza, _
                                   ByVal UserClass As eClass, ByVal UserSexo As eGenero, ByVal head As Integer) As Boolean
'*************************************************
'Author: D'Artagnan (Taken from ConnecNewUser)
'Date: 02/06/2014
'Return True if all provided values are valid. False otherwise.
'*************************************************
On Error GoTo ErrHandler
  

ValidCharacterData = False

With UserList(UserIndex)
    
    If Len(Name) > MAX_LENGTH_NAME Then
        Call WriteErrorMsg(UserIndex, "El nombre es demasiado largo, debe tener " & MAX_LENGTH_NAME & " caracteres o menos.")
        Exit Function
    End If

    If Not AsciiValidos(Name, False) Or LenB(Name) = 0 Or Not BlacklistIsValidNickname(Name) Then
        Call WriteErrorMsg(UserIndex, "Nombre inválido.")
        Exit Function
    End If
    
    If UserList(UserIndex).flags.UserLogged Then
        Call LogCheating("El usuario " & UserList(UserIndex).Name & " ha intentado crear a " & Name & " desde la IP " & UserList(UserIndex).IP)
        
        'Kick player ( and leave character inside :D )!
        Call CloseSocketSL(UserIndex)
        Call Cerrar_Usuario(UserIndex)
        
        Exit Function
    End If
    
    If Not ValidarCabeza(UserRaza, UserSexo, head) Then
        Call LogCheating("El usuario " & Name & " ha seleccionado la cabeza " & head & " desde la IP " & .IP)
        
        Call WriteErrorMsg(UserIndex, "Cabeza inválida, elija una cabeza seleccionable.")
        Exit Function
    End If
    
    ' Check if the class is enabled in the class.dat file
    If Not Classes(UserClass).Enabled Then
        Call WriteErrorMsg(UserIndex, "La clase seleccionada no se encuentra disponible en este momento.")
        Exit Function
    End If
    
    '¿Existe el personaje?
    Dim UserId As Long
    If GetCharInfo(Name, UserId, False) Then
        Call WriteErrorMsg(UserIndex, "Ya existe el personaje.")
        Exit Function
    End If
    
    
End With

ValidCharacterData = True

  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function ValidCharacterData de TCP.bas")
End Function

Public Sub CreateCharacter(ByVal UserIndex As Integer, ByRef Name As String, ByRef Password As String, ByVal UserRaza As eRaza, _
                           ByVal UserSexo As eGenero, ByVal UserClase As eClass, ByVal head As Integer, ByVal nAccountID As Long)
'*************************************************
'Author: D'Artagnan (Taken from ConnecNewUser)
'Date: 02/06/2014
'Last modification: 14/04/2015
'Set initial values and store them in the database.
'The account ID is optional, since the accounts system
'may be disabled.
'14/04/2015: D'Artagnan - Free skills for each class.
'*************************************************
On Error GoTo ErrHandler
  
Dim I As Long
Dim Hogar As eCiudad

With UserList(UserIndex)
    .flags.Muerto = 0
    .flags.Escondido = 0
    
    .Name = Name
    .clase = UserClase
    .raza = UserRaza
    .Genero = UserSexo
    .Faccion.Alignment = eCharacterAlignment.Newbie
    
    .AccountId = -1
    .AccountName = vbNullString
    .AccountEmail = vbNullString
    
    .Hogar = Hogar
    .AccountId = nAccountID
    
    .Guild.IdGuild = 0
    .Guild.RoleId = 0
    
    .Pos = GetRandomCharacterStartPosition()
    
    ' Set attributes from
    Call RecalculateUserAttributes(UserIndex)

    
    For I = 1 To NUMSKILLS
        Call ZeroSkills(UserIndex, I)
        Call CheckEluSkill(UserIndex, I, True)
    Next I
    
    ' Initial skills for each class.
    .Stats.SkillPts = Classes(.clase).ClassMods.SkillsStarter
    
    .Char.heading = eHeading.SOUTH
    
    Call DarCuerpo(UserIndex)
    .Char.head = head
    
    .OrigChar = .Char
    
    .Stats.MaxHp = GetStartingHealth(.clase, .raza)
    .Stats.MinHp = .Stats.MaxHp
    
    .Stats.MaxSta = Classes(.clase).ClassMods.StaminaStarter
    .Stats.MinSta = .Stats.MaxSta
    
    .Stats.MaxAGU = 100
    .Stats.MinAGU = 100
    
    .Stats.MaxHam = 100
    .Stats.MinHam = 100
        
    .Stats.MaxMan = GetStartingMana(UserIndex)
    .Stats.MinMAN = .Stats.MaxMan
    
    .Stats.GLD = 0
    
    .Stats.Exp = 0
    .Stats.ELU = TablaExperiencia(1)
    .Stats.ELV = 1
    
    .Stats.RankingPoints = ConstantesBalance.PlayerRankingStartingPoints
            
    '============ HECHIZOS ============
    For I = 1 To Classes(.clase).StartingSpellsQty
        .Stats.UserHechizos(I).SpellNumber = Classes(.clase).StartingSpells(I)
        .Stats.UserHechizos(I).LastUsedAt = 0
        .Stats.UserHechizos(I).LastUsedSuccessfully = False
    Next I
    
    '============ INVENTARIO ============
    For I = 1 To Classes(.clase).StartingItemsQty
        .Invent.Object(I).ObjIndex = Classes(.clase).StartingItems(I).ItemNumber
        .Invent.Object(I).Amount = Classes(.clase).StartingItems(I).Quantity
        .Invent.Object(I).Equipped = Classes(.clase).StartingItems(I).Equipped
        
        ' Equip the first weapon from the list.
        If ObjData(.Invent.Object(I).ObjIndex).ObjType = otWeapon And .Invent.Object(I).Equipped And .Invent.WeaponEqpObjIndex = 0 Then
            .Invent.WeaponEqpObjIndex = .Invent.Object(I).ObjIndex
            .Invent.WeaponEqpSlot = I
            .Char.WeaponAnim = GetWeaponAnim(UserIndex, .Invent.WeaponEqpObjIndex)
        End If
        
        If ObjData(.Invent.Object(I).ObjIndex).ObjType = otArmadura And .Invent.Object(I).Equipped And .Invent.ArmourEqpObjIndex = 0 Then
            .Invent.ArmourEqpObjIndex = .Invent.Object(I).ObjIndex
            .Invent.ArmourEqpSlot = I
            .Char.body = GetBodyForUser(UserIndex, .Invent.ArmourEqpObjIndex)
        End If
        
        If ObjData(.Invent.Object(I).ObjIndex).ObjType = otFlechas And .Invent.Object(I).Equipped And .Invent.MunicionEqpObjIndex = 0 Then
            .Invent.MunicionEqpObjIndex = .Invent.Object(I).ObjIndex
            .Invent.MunicionEqpSlot = I
        End If
        
    Next I
    
    .Invent.NroItems = Classes(.clase).StartingItemsQty

#If ConUpTime Then
    .LogOnTime = Now
    .UpTime = 0
#End If

    ' Crea la boveda de cuenta si no existe
    .flags.AccountBank = GetAccBankIndex(.AccountId)
    If .flags.AccountBank = 0 Then
        Call LoadAccountBankDB(UserIndex)
    End If
    
End With
    
    'Valores Default de facciones al Activar nuevo usuario
    Call ResetFacciones(UserIndex)

    Call SaveUserDB(UserIndex, True, True, Password)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CreateCharacter de TCP.bas")
  Err.Raise Err.Number
End Sub


Public Sub DisconnectWithMessage(ByVal nUserIndex As Integer, ByRef sMessage As String)
'***************************************************
'Author: D'Artagnan
'Last Modification: 31/05/2015
'
'***************************************************
On Error GoTo ErrHandler
  
    Call WriteErrorMsg(nUserIndex, sMessage, True)
    Call CloseSocket(nUserIndex)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub DisconnectWithMessage de TCP.bas")
End Sub

Sub CloseSocket(ByVal UserIndex As Integer, Optional ByVal bSaveUser As Boolean = True)
'***************************************************
'Author: Unknown
'Last Modification: 28/07/2014 (D'Artagnan)
'28/07/2014: D'Artagnan - Closing an user is now optional.
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)
        'Call SecurityIp.IpRestarConexion(.IPLong)
         
        If .ConnIDValida Then
            Call CloseSocketSL(UserIndex)
        End If
                
        'mato los comercios seguros
        If isTradingWithUser(UserIndex) Then
            Dim tempUsu As Integer
            
            tempUsu = getTradingUser(UserIndex)
            
            If UserList(tempUsu).flags.UserLogged Then
                If getTradingUser(tempUsu) = UserIndex Then
                    Call WriteConsoleMsg(tempUsu, "Comercio cancelado por el otro usuario", FontTypeNames.FONTTYPE_TALK)
                    Call FinComerciarUsu(tempUsu)
                End If
            End If
        End If
        
        ' if in tournamet, expells it
        If .flags.TournamentState <> eTournamentState.ieNone Then
            Call TournamentUserExpell(UserIndex, eTournamentExpellMotive.ieAbandon)
        End If

        If .flags.UserLogged Then
            Call CloseUser(UserIndex, bSaveUser)
        Else
            Call ResetUserSlot(UserIndex)
        End If
        
        Call FreeSlot(UserIndex)
    End With
Exit Sub

ErrHandler:
    Call ResetUserSlot(UserIndex)

    Call FreeSlot(UserIndex)

    Call LogError("CloseSocket - Error = " & Err.Number & " - Descripción = " & Err.Description & " - UserIndex = " & UserIndex)
End Sub

'[Alejo-21-5]: Cierra un socket sin limpiar el slot
Sub CloseSocketSL(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  
  
    With UserList(UserIndex)
        If (.ConnIDValida) Then
            Call .Connection.Close(False)
            
            Set .Connection = Nothing
            .ConnIDValida = False
        End If
    End With

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CloseSocketSL de TCP.bas")
End Sub


Function EstaPCarea(Index As Integer, Index2 As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  
Dim X As Integer, Y As Integer

For Y = UserList(Index).Pos.Y - MinYBorder + 1 To UserList(Index).Pos.Y + MinYBorder - 1
        For X = UserList(Index).Pos.X - MinXBorder + 1 To UserList(Index).Pos.X + MinXBorder - 1

            If MapData(UserList(Index).Pos.Map, X, Y).UserIndex = Index2 Then
                EstaPCarea = True
                Exit Function
            End If
        
        Next X
Next Y
EstaPCarea = False
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function EstaPCarea de TCP.bas")
End Function

Function HayPCarea(Pos As WorldPos) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

Dim X As Integer, Y As Integer
For Y = Pos.Y - MinYBorder + 1 To Pos.Y + MinYBorder - 1
        For X = Pos.X - MinXBorder + 1 To Pos.X + MinXBorder - 1
            If X > 0 And Y > 0 And X < 101 And Y < 101 Then
                If MapData(Pos.Map, X, Y).UserIndex > 0 Then
                    HayPCarea = True
                    Exit Function
                End If
            End If
        Next X
Next Y
HayPCarea = False
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function HayPCarea de TCP.bas")
End Function

Function Validatechr(ByVal UserIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

Validatechr = UserList(UserIndex).Char.head <> 0 _
                And UserList(UserIndex).Char.body <> 0 _
                And ValidateSkills(UserIndex)

  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function Validatechr de TCP.bas")
End Function

Public Function ConnectUser(ByVal UserIndex As Integer, ByRef Name As String, ByVal UserId As Long, _
                            ByRef uName As String, ByVal IsNewChar As Boolean) As Boolean
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 08/01/2015 (D'Artagnan)
'26/03/2009: ZaMa - Agrego por default que el color de dialogo de los dioses, sea como el de su nick.
'12/06/2009: ZaMa - Agrego chequeo de nivel al loguear
'14/09/2009: ZaMa - Ahora el usuario esta protegido del ataque de npcs al loguear
'11/27/2009: Budi - Se envian los InvStats del personaje y su Fuerza y Agilidad
'03/12/2009: Budi - Optimización del código
'24/07/2010: ZaMa - La posicion de comienzo es namehuak, como se habia definido inicialmente.
'16/02/2014: D'Artagnan - Bug fix: head visible on factional users when sailing.
'04/06/2014: D'Artagnan - Removed password checking when using accounts.
'08/01/2015: D'Artagnan - Safe mode always on.
'***************************************************
On Error GoTo ErrHandler
  
Dim N As Integer
Dim tStr As String
Dim I As Integer

With UserList(UserIndex)
    If .flags.UserLogged Then
        Call LogCheating("El usuario " & .Name & " ha intentado loguear a " & Name & " desde la IP " & .IP)
        'Kick player ( and leave character inside :D )!
        Call CloseSocketSL(UserIndex)
        Call Cerrar_Usuario(UserIndex)
        Exit Function
    End If
    
    'Reseteamos los FLAGS
    .flags.Escondido = 0
    .flags.TargetNPC = 0
    .flags.TargetNpcTipo = eNPCType.Comun
    .flags.TargetObj = 0
    .flags.TargetUser = 0
    .Char.FX = 0
       
    .OverHeadIcon = 0 ' 0 Means NO ICON
        
    'Controlamos no pasar el maximo de usuarios
    If NumUsers >= MaxUsers Then
        Call DisconnectWithMessage(UserIndex, "El servidor ha alcanzado el mñximo de usuarios soportado, por favor vuelva a intertarlo mñs tarde.")
        Exit Function
    End If
    
    'Este IP ya esta conectado?
    If AllowMultiLogins = 0 Then
        If CheckForSameIP(UserIndex, .IP) = True Then
            Call DisconnectWithMessage(UserIndex, "No es posible usar más de un personaje al mismo tiempo.")
            Exit Function
        End If
    End If
    
    ReDim .Mensajes(1 To Constantes.MaxPrivateMessages) As tUserMensaje
    
    '¿Existe el personaje?
    If Not IsNewChar Then
        '¿Ya esta conectado el personaje?
        If CheckForSameName(Name) Then
            If UserList(NameIndex(Name)).Counters.Saliendo Then
                Call DisconnectWithMessage(UserIndex, "El usuario está saliendo.")
            Else
                Call DisconnectWithMessage(UserIndex, "El personaje está conectado.")
            End If
            Exit Function
        End If
    End If
    
    'Reseteamos los privilegios
    .flags.Privilegios = 0
        
    'Vemos que clase de user es (se lo usa para setear los privilegios al loguear el PJ)
    If EsAdmin(Name) Then
        .flags.Privilegios = .flags.Privilegios Or PlayerType.Admin
        Call LogGM(Name, "Se conecto con ip:" & .IP)
    ElseIf EsDios(Name) Then
        .flags.Privilegios = .flags.Privilegios Or PlayerType.Dios
        Call LogGM(Name, "Se conecto con ip:" & .IP)
    ElseIf EsSemiDios(Name) Then
        .flags.Privilegios = .flags.Privilegios Or PlayerType.SemiDios
        
        .flags.PrivEspecial = EsGmEspecial(Name)
        
        Call LogGM(Name, "Se conecto con ip:" & .IP)
    ElseIf EsConsejero(Name) Then
        .flags.Privilegios = .flags.Privilegios Or PlayerType.Consejero
        Call LogGM(Name, "Se conecto con ip:" & .IP)
    Else
        .flags.Privilegios = .flags.Privilegios Or PlayerType.User
        .flags.AdminPerseguible = True
    End If
    
    'Add RM flag if needed
    If EsRolesMaster(Name) Then
        .flags.Privilegios = .flags.Privilegios Or PlayerType.RoleMaster
    End If
    
    If ServerSoloGMs > 0 Then
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero)) = 0 Then
            Call DisconnectWithMessage(UserIndex, "Servidor restringido a administradores. Por favor reintente en unos momentos.")
            Exit Function
        End If
    End If
    
    'Now we have userId
    Dim CharacterLoaded As Boolean, PunishmentDesc As String, PunisherName As String, PunishmentEndDate As Date
    CharacterLoaded = LoadCharFromDB(UserIndex, UserId, PunishmentDesc, PunisherName, PunishmentEndDate)
        
    ' Is the user banned?
    If PunishmentDesc <> vbNullString Then
        Call DisconnectWithMessage(UserIndex, _
            "Se te ha prohibido la entrada a Argentum Online hasta el " & PunishmentEndDate & " debido a tu mal comportamiento. " & _
            "Puedes consultar el reglamento y el sistema de soporte desde www.alkononline.com.ar" & _
            vbNewLine & vbNewLine & "Motivo: " & PunishmentDesc & "." & vbNewLine & "Sanción aplicada por: " & PunisherName & ".")
        Exit Function
    End If
    
    ' Check if the class is enabled in the class.dat file
    If Not Classes(.clase).Enabled Then
        Call DisconnectWithMessage(UserIndex, "La clase seleccionada no se encuentra disponible en este momento.")
        Exit Function
    End If
    
    ' Recalculate user attributes
    Call RecalculateUserAttributes(UserIndex)

    ' HP and Mana needs to be recalculated after logging in
    .Stats.MaxHp = RecalculateCharacterMaxHealth(UserIndex)
    If .Stats.MinHp > .Stats.MaxHp Then .Stats.MinHp = .Stats.MaxHp
    
    .Stats.MaxMan = RecalculateCharacterMaxMana(UserIndex)
    If .Stats.MinMAN > .Stats.MaxMan Then .Stats.MinMAN = .Stats.MaxMan
    
    ' Recalculate Passive skills
    Call SetUserPassiveDefaults(UserIndex)
    Call RecalculateUserPassives(UserIndex, False)
    
    ' We send the intervals the client need to use for this session.
    Call WriteSetIntervals(UserIndex)
      
    'Carga la boveda de cuenta
    .flags.AccountBank = GetAccBankIndex(.AccountId)
    If .flags.AccountBank = 0 Then
        Call LoadAccountBankDB(UserIndex)
    End If
    
    If Not Validatechr(UserIndex) Then
        Call DisconnectWithMessage(UserIndex, "El personaje no pudo ser cargado. Contacte a los administradores para solucionar este inconveniente.")
        Exit Function
    End If
    
    If .Invent.EscudoEqpSlot = 0 Then .Char.ShieldAnim = ConstantesGRH.NingunEscudo
    If .Invent.CascoEqpSlot = 0 Then .Char.CascoAnim = ConstantesGRH.NingunCasco
    If .Invent.WeaponEqpSlot = 0 Then .Char.WeaponAnim = ConstantesGRH.NingunArma
    
    If .Invent.MochilaEqpSlot > 0 Then
        .CurrentInventorySlots = MAX_NORMAL_INVENTORY_SLOTS + ObjData(.Invent.Object(.Invent.MochilaEqpSlot).ObjIndex).MochilaType * 5
    Else
        .CurrentInventorySlots = MAX_NORMAL_INVENTORY_SLOTS
    End If
    If (.flags.Muerto = 0) Then
        .flags.SeguroResu = False
        Call WriteMultiMessage(UserIndex, eMessages.ResuscitationSafeOff)
    Else
        .flags.SeguroResu = True
        Call WriteMultiMessage(UserIndex, eMessages.ResuscitationSafeOn)
    End If
    
    Call UpdateUserInv(True, UserIndex, 0)
    Call UpdateUserHechizos(True, UserIndex, 1)
    
    If .flags.Paralizado Then
        Call WriteParalizeOK(UserIndex)
    End If
    
    Dim mapa As Integer
    mapa = .Pos.Map

    If Not MapaValido(mapa) Then
        Call WriteErrorMsg(UserIndex, "El PJ se encuenta en un mapa inválido.")
        Call CloseSocket(UserIndex)
        Exit Function
    End If
    
    ' If map has different initial coords, update it
    Dim StartMap As Integer
    StartMap = MapInfo(mapa).StartPos.Map
    If StartMap <> 0 Then
        If MapaValido(StartMap) Then
            .Pos = MapInfo(mapa).StartPos
            mapa = StartMap
        End If
    End If

    'Tratamos de evitar en lo posible el "Telefrag". Solo 1 intento de loguear en pos adjacentes.
    'Codigo por Pablo (ToxicWaste) y revisado por Nacho (Integer), corregido para que realmente ande y no tire el server por Juan Martín Sotuyo Dodero (Maraxus)
    If MapData(mapa, .Pos.X, .Pos.Y).UserIndex <> 0 Or MapData(mapa, .Pos.X, .Pos.Y).NpcIndex <> 0 Then
        Dim FoundPlace As Boolean
        Dim esAgua As Boolean
        Dim tX As Long
        Dim tY As Long
        
        FoundPlace = False
        esAgua = HayAgua(mapa, .Pos.X, .Pos.Y)
        
        For tY = .Pos.Y - 1 To .Pos.Y + 1
            For tX = .Pos.X - 1 To .Pos.X + 1
                If esAgua Then
                    'reviso que sea pos legal en agua, que no haya User ni NPC para poder loguear.
                    If LegalPos(mapa, tX, tY, True, False) Then
                        FoundPlace = True
                        Exit For
                    End If
                Else
                    'reviso que sea pos legal en tierra, que no haya User ni NPC para poder loguear.
                    If LegalPos(mapa, tX, tY, False, True) Then
                        FoundPlace = True
                        Exit For
                    End If
                End If
            Next tX
            
            If FoundPlace Then _
                Exit For
        Next tY
        
        If FoundPlace Then 'Si encontramos un lugar, listo, nos quedamos ahi
            .Pos.X = tX
            .Pos.Y = tY
        Else
            Dim tempUsu As Integer
            
            tempUsu = MapData(mapa, .Pos.X, .Pos.Y).UserIndex
            
            'Si no encontramos un lugar, sacamos al usuario que tenemos abajo, y si es un NPC, lo pisamos.
            If tempUsu <> 0 Then
               'Si no encontramos lugar, y abajo teniamos a un usuario, lo pisamos y cerramos su comercio seguro
                If isTradingWithUser(tempUsu) Then
                    tempUsu = getTradingUser(tempUsu)
                    
                    'Le avisamos al que estaba comerciando que se tuvo que ir.
                    If UserList(tempUsu).flags.UserLogged Then
                        Call FinComerciarUsu(tempUsu)
                        Call WriteConsoleMsg(tempUsu, "Comercio cancelado. El otro usuario se ha desconectado.", FontTypeNames.FONTTYPE_TALK)
                    End If
                    
                    tempUsu = MapData(mapa, .Pos.X, .Pos.Y).UserIndex
                    
                    'Lo sacamos.
                    If UserList(tempUsu).flags.UserLogged Then
                        Call FinComerciarUsu(tempUsu)
                        Call WriteErrorMsg(tempUsu, "Alguien se ha conectado donde te encontrabas, por favor reconéctate...")
                    End If
                End If
                
                Call CloseSocket(tempUsu)
            End If
        End If
    End If
    
    'Nombre de sistema
    .Name = Name
    .secName = uName
    
    .ShowName = True 'Por default los nombres son visibles
    
    If UserList(UserIndex).Guild.IdGuild > 0 Then
    
        .Guild.GuildIndex = GuildIndexOf(.Guild.IdGuild)
        .Guild.GuildMemberIndex = GetMemberIndexOf(UserIndex)
        .Guild.RoleId = GetGuildRoleId(.Guild.GuildIndex, UserIndex)
        
        Call WriteGuildInfo(UserIndex)
        Call WriteGuildRolesList(UserIndex)
        Call AddOnlineMember(UserIndex)
        Call WriteGuildMembersList(UserIndex)
        Call WriteGuildBankList(UserIndex)
        Call WriteGuildUpgradesList(UserIndex)
        Call WriteGuildUpgradesAcquired(UserIndex)
        Call WriteGuildQuestsCompletedList(UserIndex)
        Call WriteGuildCurrentQuestInfo(UserIndex)
        
        Call NotifyMemberConnection(UserIndex, True)
    End If
    
    'If in the water, and has a boat, equip it!
    If .Invent.BarcoObjIndex > 0 And _
           (HayAgua(mapa, .Pos.X, .Pos.Y) Or UserAreaHasWater(UserIndex) Or _
        BodyIsBoat(.Char.body)) Then
        .Char.head = 0
        If .flags.Muerto = 0 Then
            Call ToggleBoatBody(UserIndex)
        Else
            .Char.body = ConstantesGRH.FragataFantasmal
            .Char.ShieldAnim = ConstantesGRH.NingunEscudo
            .Char.WeaponAnim = ConstantesGRH.NingunArma
            .Char.CascoAnim = ConstantesGRH.NingunCasco
        End If
    
        .flags.Navegando = 1
    End If
    
    Call WriteUpdateCharacterInfo(UserIndex)
    'Info
    Call WriteUserIndexInServer(UserIndex) 'Enviamos el User index

    Call WriteChangeMap(UserIndex, .Pos.Map, MapInfo(.Pos.Map).MapVersion) 'Carga el mapa
    Call WritePlayMusic(UserIndex, .Pos.Map)
    
    If .flags.Privilegios = PlayerType.Dios Or .flags.Privilegios = (PlayerType.Dios Or PlayerType.ChaosCouncil) Or .flags.Privilegios = (PlayerType.Dios Or PlayerType.ChaosCouncil) Or .flags.Privilegios = (PlayerType.Dios Or PlayerType.RoyalCouncil) Then
        'Gods, with or without being in the council.
        .flags.ChatColor = RGB(250, 250, 150)
    ElseIf .flags.Privilegios = PlayerType.SemiDios Or .flags.Privilegios = (PlayerType.SemiDios Or PlayerType.ChaosCouncil) Or .flags.Privilegios = (PlayerType.SemiDios Or PlayerType.ChaosCouncil) Or .flags.Privilegios = (PlayerType.SemiDios Or PlayerType.RoyalCouncil) Then
        .flags.ChatColor = RGB(0, 255, 0)
    ElseIf .flags.Privilegios = PlayerType.Consejero Or .flags.Privilegios = (PlayerType.Consejero Or PlayerType.ChaosCouncil) Or .flags.Privilegios = (PlayerType.Consejero Or PlayerType.ChaosCouncil) Or .flags.Privilegios = (PlayerType.Consejero Or PlayerType.RoyalCouncil) Then
        .flags.ChatColor = RGB(0, 255, 0)
    ElseIf .flags.Privilegios = (PlayerType.User Or PlayerType.RoyalCouncil) Then
        .flags.ChatColor = RGB(0, 255, 255)
    ElseIf .flags.Privilegios = (PlayerType.User Or PlayerType.ChaosCouncil) Then
        .flags.ChatColor = RGB(255, 128, 64)
    Else
        .flags.ChatColor = vbWhite
    End If
    
    
    ''[EL OSO]: TRAIGO ESTO ACA ARRIBA PARA DARLE EL IP!
    #If ConUpTime Then
        .LogOnTime = Now
    #End If
    
    ' Sometimes the head is visible even when sailing
    If .Char.head > 0 And .flags.Navegando Then _
        .Char.head = 0
        
    Dim IsAdminPlayer As Boolean
    
    'Crea  el personaje del usuario
    If Not NewUserChar(UserIndex) Then
        Exit Function
    End If
    
    IsAdminPlayer = (.flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster)) = 0
    
    'El Admin tiene que loguear invisible.
    If IsAdminPlayer Then

        .flags.SendDenounces = True
        Call DoAdminInvisible(UserIndex)
    End If
    

    
    If Not IsAdminPlayer Then
        If MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger = eTrigger.zonaOscura Then
            Call SendData(SendTarget.ToUsersAndRmsAndCounselorsAreaButGMs, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, True))
        End If
    End If
    
    Call WriteUserCharIndexInServer(UserIndex)
    ''[/el oso]
    
    Call DoTileEvents(UserIndex, .Pos.Map, .Pos.X, .Pos.Y)
    
    Call WriteUpdateUserStats(UserIndex)
    Call CheckUserLevel(UserIndex)
    
    Call WriteUpdateHungerAndThirst(UserIndex)
    Call WriteUpdateStrenghtAndDexterity(UserIndex)
    
    ' Step on trigger?
    Call CheckTriggerActivation(UserIndex, 0, .Pos.Map, .Pos.X, .Pos.Y, False)
    
    If haciendoBK Then
        Call WritePauseToggle(UserIndex)
        Call WriteConsoleMsg(UserIndex, "Servidor> Por favor espera algunos segundos, el WorldSave está ejecutándose.", FontTypeNames.FONTTYPE_SERVER, eMessageType.Admin)
    End If
    
    If EnPausa Then
        Call WritePauseToggle(UserIndex)
        Call WriteConsoleMsg(UserIndex, "Servidor> Lo sentimos mucho pero el servidor se encuentra actualmente detenido. Intenta ingresar más tarde.", FontTypeNames.FONTTYPE_SERVER, eMessageType.Admin)
    End If
    
    If EnTesting And .Stats.ELV >= 18 Then
        Call DisconnectWithMessage(UserIndex, "Servidor en Testing por unos minutos, conectese con PJs de nivel menor a 18. No se conecte con Pjs que puedan resultar importantes por ahora pues pueden arruinarse.")
        Exit Function
    End If
    
    If TieneMensajesNuevos(UserIndex) Then
        Call WriteConsoleMsg(UserIndex, "¡Tienes mensajes privados sin leer! Puedes leerlos escribiendo /LISTAMPS", FontTypeNames.FONTTYPE_FIGHT)
    End If
    
    'Actualiza el Num de usuarios
    'DE ACA EN ADELANTE GRABA EL CHARFILE, OJO!
    NumUsers = NumUsers + 1
    .flags.UserLogged = True
    
    .DbConnectionEventId = SetUserLoggedStateDB(.ID, True, .IP, 0)

    'Call EstadisticasWeb.Informar(CANTIDAD_ONLINE, NumUsers)
    
    MapInfo(.Pos.Map).NumUsers = MapInfo(.Pos.Map).NumUsers + 1

    If .Stats.SkillPts > 0 Then
        Call WriteLevelUp(UserIndex, .Stats.SkillPts)
    End If
    
    If NumUsers > RECORDusuarios Then
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("RECORD de usuarios conectados simultaneamente." & "Hay " & NumUsers & " usuarios.", FontTypeNames.FONTTYPE_INFO))
        RECORDusuarios = NumUsers
        Call WriteVar(IniPath & "Server.ini", "INIT", "RECORD", Str(RECORDusuarios))
        
        'Call EstadisticasWeb.Informar(RECORD_USUARIOS, RECORDusuarios)
    End If
    
    If .TammedPetsCount > 0 And MapInfo(.Pos.Map).Pk Then
        
        Call WarpMascotas(UserIndex)
        Dim MascotasIndex As Integer

    ElseIf .TammedPetsCount > 0 And MapInfo(.Pos.Map).Pk = False Then
        Call WriteConsoleMsg(UserIndex, "No se permiten mascotas en zona segura. Éstas te esperarán afuera.", FontTypeNames.FONTTYPE_INFO)
    End If
    
    ' Send the pet list to the user.
    Call WriteSendPetList(UserIndex)
    
    If .flags.Navegando = 1 Then
        Call WriteNavigateChange(UserIndex, True)
    End If
    
    ' Safe mode always on.
    .flags.Seguro = True
    Call WriteMultiMessage(UserIndex, eMessages.SafeModeOn) 'Call WriteSafeModeOn(UserIndex)
    
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, ConstantesFX.FxWarp, 0))
    
    Call WriteLoggedMessage(UserIndex)
    
    Call CheckIsBeingFollowed(UserIndex)

    ' Esta protegido del ataque de npcs por 5 segundos, si no realiza ninguna accion
    Call IntervaloPermiteSerAtacado(UserIndex, True)
    
    If Lloviendo Then
        Call WriteRainToggle(UserIndex)
    End If
        
    '.bIsPremium = AccountIsPremium(.accountId)
    
    Call MostrarNumUsers
    
    #If EnableSecurity Then
        Call Security.UserConnected(UserIndex)
    #End If

    N = FreeFile
    Open ServerConfiguration.LogsPaths.GeneralPath & "numusers.log" For Output As N
    Print #N, NumUsers
    Close #N
    
    .InstanceId = GetTickCount()
    
    Call SendDamageOverTimeLoad(.ID)
    
    ' Set default counters
    .Counters.COMCounter = SetIntervalEnd(ServerConfiguration.Intervals.IntervaloHambre)
    .Counters.AGUACounter = SetIntervalEnd(ServerConfiguration.Intervals.IntervaloSed)
    
    ConnectUser = True
    
End With
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function ConnectUser de TCP.bas")
End Function

Sub ResetFacciones(ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 23/01/2007
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'23/01/2007 Pablo (ToxicWaste) - Agrego NivelIngreso, FechaIngreso, MatadosIngreso y NextRecompensa.
'*************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex).Faccion
        .ArmadaReal = 0
        .NeutralsKilled = 0
        .CiudadanosMatados = 0
        .CriminalesMatados = 0
        .FuerzasCaos = 0
        .FechaIngreso = 0
        .RecibioArmaduraCaos = 0
        .RecibioArmaduraReal = 0
        .RecibioExpInicialCaos = 0
        .RecibioExpInicialReal = 0
        .RecompensasCaos = 0
        .RecompensasReal = 0
        .Reenlistadas = 0
        .NivelIngreso = 0
        .MatadosIngreso = 0
        .NextRecompensa = 0
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ResetFacciones de TCP.bas")
End Sub

Sub ResetContadores(ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 10/07/2010
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'05/20/2007 Integer - Agregue todas las variables que faltaban.
'10/07/2010: ZaMa - Agrego los counters que faltaban.
'*************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex).Counters
        .Petrificado = 0
        .Putrefaccion = 0
        .AGUACounter = 0
        .AttackCounter = 0
        .bPuedeMeditar = True
        .Ceguera = 0
        .COMCounter = 0
        .Estupidez = 0
        .failedUsageAttempts = 0
        .Frio = 0
        .goHome = 0
        .HPCounter = 0
        .RestingHPCounter = 0
        .IdleCount = 0
        .Invisibilidad = 0
        .Lava = 0
        .Mimetismo = 0
        .Paralisis = 0
        .Pena = 0
        .PiqueteC = 0
        .Saliendo = False
        .Salir = 0
        .STACounter = 0
        .RestingSTACounter = 0
        .TiempoOculto = 0
        .TimerEstadoAtacable = 0
        .TimerGolpeMagia = 0
        .TimerGolpeUsar = 0
        .TimerLanzarSpell = 0
        .TimerMagiaGolpe = 0
        .TimerPerteneceNpc = 0
        .TimerPuedeAtacar = 0
        .TimerPuedeSerAtacado = 0
        .TimerPuedeTrabajar = 0
        .TimerPuedeUsarArco = 0
        .TimerUsar = 0
        .TimerHide = 0
        .Veneno = 0
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ResetContadores de TCP.bas")
End Sub

Sub ResetCharInfo(ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'*************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex).Char
        .body = 0
        .CascoAnim = 0
        .CharIndex = 0
        .FX = 0
        .head = 0
        .Loops = 0
        .heading = 0
        .Loops = 0
        .ShieldAnim = 0
        .WeaponAnim = 0
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ResetCharInfo de TCP.bas")
End Sub

Sub ResetBasicUserInfo(ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'*************************************************
On Error GoTo ErrHandler
  
    Dim LoopC As Byte

    With UserList(UserIndex)
        .ID = 0
        .Name = vbNullString
        .secName = vbNullString
        .Desc = vbNullString
        .DescRM = vbNullString
        .Pos.Map = 0
        .Pos.X = 0
        .Pos.Y = 0
        .clase = 0
        .Genero = 0
        .Hogar = 0
        .raza = 0
        
        .AccountId = -1
        .AccountName = vbNullString
        .AccountEmail = vbNullString
            
        For LoopC = 1 To 8
            .AccountCharNames(LoopC) = vbNullString
        Next LoopC
        
        .nSessionId = -1
        .ClientTempCode = vbNullString
        
        .PartyIndex = 0
        .PartySolicitud = 0
        
        .OverHeadIcon = 0

        With .Stats
            .Banco = 0
            .ELV = 0
            .ELU = 0
            .Exp = 0
            .Def = 0
            '.CriminalesMatados = 0
            .NPCsMuertos = 0
            .UsuariosMatados = 0
            .SkillPts = 0
            .GLD = 0
            .UserAtributos(1) = 0
            .UserAtributos(2) = 0
            .UserAtributos(3) = 0
            .UserAtributos(4) = 0
            .UserAtributos(5) = 0
            .UserAtributosBackUP(1) = 0
            .UserAtributosBackUP(2) = 0
            .UserAtributosBackUP(3) = 0
            .UserAtributosBackUP(4) = 0
            .UserAtributosBackUP(5) = 0
            .MasteryPoints = 0
            .RankingPoints = 0
            
            If ServerConfiguration.PassiveSkillsQty > 0 Then
                Erase .UserPassives
                'For LoopC = 1 To ServerConfiguration.PassiveSkillsQty
                '    .UserPassives(LoopC).Id = 0
                '    .UserPassives(LoopC).Enabled = False
                '    .UserPassives(LoopC).Name = ""
                'Next LoopC
            End If
            
            
        End With
        
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ResetBasicUserInfo de TCP.bas")
End Sub

Sub ResetGuildInfo(ByVal UserIndex As Integer)
On Error GoTo ErrHandler
  
    Dim LoopC As Byte

    With UserList(UserIndex).Guild
        .IdGuild = 0
        .GuildIndex = 0
        .GuildMemberIndex = 0
        .RoleId = 0
        .RoleIndex = 0
        .GuildRange = 0
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ResetBasicUserInfo de TCP.bas")
End Sub

Sub ResetUserFlags(ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 06/28/2008
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'06/28/2008 NicoNZ - Agrego el flag Inmovilizado
'*************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex).flags
        .Petrificado = 0
        .Putrefaccion = 0
        .DueloIndex = 0
        .DueloPublico = 0
        .DueloTeam = 0
        .AccountBank = 0
        .Comerciando = False
        .Ban = 0
        .Escondido = 0
        .DuracionEfecto = 0
        .NpcInv = 0
        .StatsChanged = 0
        .TargetNPC = 0
        .TargetNpcTipo = eNPCType.Comun
        .TargetObj = 0
        .TargetObjMap = 0
        .TargetObjX = 0
        .TargetObjY = 0
        .TargetUser = 0
        .TipoPocion = 0
        .TomoPocion = False
        .Descuento = vbNullString
        .Hambre = 0
        .Sed = 0
        .Descansar = False
        .Vuela = 0
        .Navegando = 0
        .Oculto = 0
        .Envenenado = 0
        .invisible = 0
        .Paralizado = 0
        .Inmovilizado = 0
        .Maldicion = 0
        .Bendicion = 0
        .Meditando = 0
        .Privilegios = 0
        .PrivEspecial = False
        .PuedeMoverse = 0
        .AdminInvisible = 0
        .ValCoDe = 0
        .CastedSpellNumber = 0
        .CastedSpellIndex = 0
        .TimesWalk = 0
        .StartWalk = 0
        .CountSH = 0
        .Silenciado = 0
        .AdminPerseguible = False
        .lastMap = 0
        .Traveling = 0
        .AtacablePor = 0
        .AtacadoPorNpc = 0
        .AtacadoPorUser = 0
        .NoPuedeSerAtacado = False
        .ShareNpcWith = 0
        .HelpMode = False
        .HelpingUser = 0
        .HelpingUserName = vbNullString
        .HelpedBy = 0
        .HelpedByUserName = vbNullString
        .Ignorado = False
        .SendDenounces = False
        .ParalizedBy = vbNullString
        .ParalizedByIndex = 0
        .ParalizedByNpcIndex = 0
        .CountQuestTime = False
        .Mimetizado = 0
        .MimetizadoType = 0
        .LastNpcInvoked = 0
        .TournamentState = 0
        .nCommerceSourceUser = 0
        
        .LastTamedPet = 0
        
        If .OwnedNpc <> 0 Then
            Call PerdioNpc(UserIndex)
        End If
        
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ResetUserFlags de TCP.bas")
End Sub

Sub ResetUserSpells(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    Dim LoopC As Long
    For LoopC = 1 To MAXUSERHECHIZOS
        UserList(UserIndex).Stats.UserHechizos(LoopC).SpellNumber = 0
        UserList(UserIndex).Stats.UserHechizos(LoopC).LastUsedAt = 0
        UserList(UserIndex).Stats.UserHechizos(LoopC).LastUsedSuccessfully = False
        
    Next LoopC
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ResetUserSpells de TCP.bas")
End Sub

Sub ResetUserPets(ByVal UserIndex As Integer)
On Error GoTo ErrHandler

    Dim LoopC As Long
    
    With UserList(UserIndex)
            
        For LoopC = 1 To .TammedPetsCount
            .TammedPets(LoopC).NpcIndex = 0
            .TammedPets(LoopC).NpcNumber = 0
            .TammedPets(LoopC).RemainingLife = 0
        Next LoopC
        
        For LoopC = 1 To .InvokedPetsCount
            .InvokedPets(LoopC).NpcIndex = 0
            .InvokedPets(LoopC).NpcNumber = 0
            .InvokedPets(LoopC).RemainingLife = 0
        Next LoopC
        
        .TammedPetsCount = 0
        .InvokedPetsCount = 0
        
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ResetUserPets de TCP.bas")
End Sub

Sub ResetUserBanco(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    Dim LoopC As Long
    
    For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
          UserList(UserIndex).BancoInvent.Object(LoopC).Amount = 0
          UserList(UserIndex).BancoInvent.Object(LoopC).Equipped = 0
          UserList(UserIndex).BancoInvent.Object(LoopC).ObjIndex = 0
    Next LoopC
    
    UserList(UserIndex).BancoInvent.NroItems = 0
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ResetUserBanco de TCP.bas")
End Sub


Public Sub ResetUserWorkerStore(ByVal UserIndex)
On Error GoTo ErrHandler
    
    With UserList(UserIndex).CraftingStore
        .IsOpen = False
        .ItemsQty = 0
        Erase .Items
        
        .LastCraftedObjectAt = DateSerial(1900, 1, 1)
        
        .MoneyEarned = 0
        .CraftedObjectsQty = 0
    
    End With
    
    Exit Sub
    
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ResetUserBanco de TCP.bas")
End Sub



Public Sub LimpiarComercioSeguro(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    If isTradingWithUser(UserIndex) Then
        Call FinComerciarUsu(getTradingUser(UserIndex))
        Call FinComerciarUsu(UserIndex)
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LimpiarComercioSeguro de TCP.bas")
End Sub

Sub ResetUserSlot(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

Dim I As Long


With UserList(UserIndex)

    .AccountName = vbNullString
    .AccountEmail = vbNullString
    .AccountId = 0
    .ClientTempCode = vbNullString
    .nSessionId = -1
    
    .Punishment.ID = 0
    .Punishment.Punisher = vbNullString
    .Punishment.Reason = vbNullString
    .Punishment.EndDate = DateSerial(1900, 1, 1)

    .DbConnectionEventId = 0
End With

Call LimpiarComercioSeguro(UserIndex)
Call ResetFacciones(UserIndex)
Call ResetContadores(UserIndex)
Call ResetCharInfo(UserIndex)
Call ResetBasicUserInfo(UserIndex)
Call ResetGuildInfo(UserIndex)
Call ResetUserFlags(UserIndex)
Call LimpiarInventario(UserIndex)
Call ResetUserSpells(UserIndex)
Call ResetUserPets(UserIndex)
Call ResetUserBanco(UserIndex)
Call ResetUserWorkerStore(UserIndex)
Call ResetMasteries(UserIndex)
Call LimpiarMensajes(UserIndex)

With UserList(UserIndex).ComUsu
    .Acepto = False
    
    For I = 1 To MAX_OFFER_SLOTS
        .cant(I) = 0
        .Objeto(I) = 0
    Next I
    
    .GoldAmount = 0
    .DestNick = vbNullString
End With
 
#If EnableSecurity Then
    Call resetSecurity(UserIndex)
#End If

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ResetUserSlot de TCP.bas")
End Sub

Sub CloseUser(ByVal UserIndex As Integer, Optional ByVal bSaveUser As Boolean = True)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo ErrHandler

Dim N As Integer
Dim Map As Integer
Dim Name As String
Dim I As Integer
Dim TrapNumber As Integer

Dim aN As Integer

If NumUsers > 0 Then NumUsers = NumUsers - 1

With UserList(UserIndex)

    If UserList(UserIndex).flags.DueloPublico > 0 Then
        PublicDuel(UserList(UserIndex).flags.DueloPublico) = 0
    End If
    
    If UserList(UserIndex).flags.DueloIndex > 0 Then
        If DuelData.Duelo(UserList(UserIndex).flags.DueloIndex).estado = eDuelState.Esperando_Inicio Or DuelData.Duelo(UserList(UserIndex).flags.DueloIndex).estado = eDuelState.Iniciado Then
            Call AbandonarDuelo(UserList(UserIndex).flags.DueloIndex, UserIndex)
        ElseIf DuelData.Duelo(UserList(UserIndex).flags.DueloIndex).estado = eDuelState.Esperando_Jugadores Then
            Call SendData(SendTarget.ToDuelo, UserList(UserIndex).flags.DueloIndex, PrepareMessageConsoleMsg("El Duelo ha sido cancelado, " & UserList(UserIndex).Name & " se ha desconectado.", FontTypeNames.FONTTYPE_INFO))
            Call CancelarDuelo(UserList(UserIndex).flags.DueloIndex)
        ElseIf DuelData.Duelo(UserList(UserIndex).flags.DueloIndex).estado = eDuelState.Esperando_Final Then
            Call ApresurarFinalDuelo(UserList(UserIndex).flags.DueloIndex)
        End If
    End If
    
    ' Send a message to the state server to persist all the DoT effects
    Call modStateServer.SendDamageOverTimePersist(UserList(UserIndex).ID)
    
    Call DisableAllTrapsForUser(UserIndex)
    aN = .flags.AtacadoPorNpc
    If aN > 0 Then
          Npclist(aN).Movement = Npclist(aN).flags.OldMovement
          Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
          Npclist(aN).flags.AttackedBy = vbNullString
    End If
    
    aN = .flags.NPCAtacado
    If aN > 0 Then
        If Npclist(aN).flags.AttackedFirstBy = .Name Then
            Npclist(aN).flags.AttackedFirstBy = vbNullString
        End If
    End If
    .flags.AtacadoPorNpc = 0
    .flags.NPCAtacado = 0
    
    Map = .Pos.Map
    Name = UCase$(.Name)
    
    .Char.FX = 0
    .Char.Loops = 0
    Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0))
    
    .flags.UserLogged = False
    .Counters.Saliendo = False
    
    'Le devolvemos el body y head originales
    .flags.AdminInvisible = 0

    'Save statistics
    Call Statistics.UserDisconnected(UserIndex)
    
    ' Grabamos el personaje del usuario
    If bSaveUser Then
        'si esta en party le devolvemos la experiencia
        If .PartyIndex > 0 Then Call mdParty.SalirDeParty(UserIndex)
        Call SaveUserDB(UserIndex, True, False, "")
    End If
    
    Call SetUserLoggedStateDB(.ID, False, .IP, .DbConnectionEventId)
    
    Call CloseAccBank(UserList(UserIndex).flags.AccountBank)

    'Quitar el dialogo
    'If MapInfo(Map).NumUsers > 0 Then
    '    Call SendToUserArea(UserIndex, "QDL" & .Char.charindex)
    'End If
    
    If MapaValido(Map) Then
        If MapInfo(Map).NumUsers > 0 Then
            Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageRemoveCharDialog(.Char.CharIndex))
        End If
    
        'Update Map Users
        MapInfo(Map).NumUsers = MapInfo(Map).NumUsers - 1

        If MapInfo(Map).NumUsers < 0 Then
            MapInfo(Map).NumUsers = 0
        End If
    End If

    'Borrar el personaje
    If .Char.CharIndex > 0 Then
        Call EraseUserChar(UserIndex)
    End If
    
    'Borrar mascotas
    For I = 1 To Classes(.clase).ClassMods.MaxTammedPets
        If .TammedPets(I).NpcIndex > 0 Then
            If Npclist(.TammedPets(I).NpcIndex).flags.NPCActive Then _
                Call QuitarNPC(.TammedPets(I).NpcIndex)
        End If
    Next I
    
    For I = 1 To Classes(.clase).ClassMods.MaxInvokedPets
        If .InvokedPets(I).NpcIndex > 0 Then
            If Npclist(.InvokedPets(I).NpcIndex).flags.NPCActive Then _
                Call QuitarNPC(.InvokedPets(I).NpcIndex)
        End If
    Next I
    
    ' Si el usuario habia dejado un msg en la gm's queue lo borramos
    If Ayuda.Existe(.Name) Then Call Ayuda.Quitar(.Name)
    
    Call RemoveOnlineMember(UserIndex)
    If .InvitationGuildIndex > 0 Then
        Call modGuild_Functions.ClearInvitationByUserId(.InvitationGuildIndex, .Id)
    End If
    Call NotifyMemberConnection(UserIndex, False)
    
    .Guild.GuildMemberIndex = 0
    
    Call ResetUserSlot(UserIndex)
    
    Call MostrarNumUsers
    
    Call EstadisticasWeb.Informar(CANTIDAD_ONLINE, NumUsers)
    
End With

Exit Sub

ErrHandler:

    Dim UserName As String
    If UserIndex > 0 Then UserName = UserList(UserIndex).Name

    Call LogError("Error en CloseUser. Número " & Err.Number & " Descripción: " & Err.Description & _
        ".User: " & UserName & "(" & UserIndex & "). Map: " & Map)
End Sub

Public Sub EcharPjsNoPrivilegiados()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    Dim LoopC As Long
    
    For LoopC = 1 To LastUser
        If UserList(LoopC).flags.UserLogged And UserList(LoopC).ConnIDValida Then
            If UserList(LoopC).flags.Privilegios And PlayerType.User Then
                Call CloseSocket(LoopC)
            End If
        End If
    Next LoopC

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub EcharPjsNoPrivilegiados de TCP.bas")
End Sub

Public Function GetLongIp(IPAddress As String) As Long
    Dim arrTemp As Variant
    Dim I       As Integer
    Dim lngTemp As Double
    
    arrTemp = Split(IPAddress, ".")
    
    For I = 0 To UBound(arrTemp)
        lngTemp = lngTemp + CLng(arrTemp(I)) * (256 ^ (3 - I))
    Next

    GetLongIp = UnsignedLongToSigned(lngTemp)
End Function

Public Function UnsignedLongToSigned(ByVal Value As Double) As Long
    If Value <= 2147483647 Then
        UnsignedLongToSigned = Value
    Else
        UnsignedLongToSigned = Value - 4294967296#
    End If
End Function
