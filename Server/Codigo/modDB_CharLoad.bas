Attribute VB_Name = "modDB_CharLoad"
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

Option Explicit


Public Function LoadCharFromDB(ByVal UserIndex As Integer, ByVal UserId As Long, Optional ByRef PunishmentDescription As String = vbNullString, _
                            Optional ByRef PunisherName As String = vbNullString, Optional ByRef PunishmentEndDate As Date) As Boolean
'***************************************************
'Loads general char info from DB
'***************************************************
On Error GoTo ErrHandler
  
       
    Dim Rs As Recordset
    Dim Cmd As ADODB.Command
    Set Cmd = New ADODB.Command
    
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "sp_LoadChar"
    
    Cmd.Parameters.Append Cmd.CreateParameter("userID", adInteger, adParamInput, 1, UserId)
    
    Set Rs = ExecuteSqlCommand(Cmd)
    
    ' This stored procedure will return multiple recordsets, in this order:
    ' - Character punishment status
    ' - Character tables related to USER_INFO with a 1to1 relationship:  USER_INFO, USER_STATS, USER_FLAGS, USER_ATTRIBUTES, USER_FACTION
    ' - Character inventory
    ' - Character bank
    ' - Character spells
    ' - Character skills
    ' - Character pets
    ' - Character passive skills
    ' - Character masteries
    
   
    With UserList(UserIndex)
        
        If Not Rs.EOF Then
        
            ' If the user is banned, skip loading the rest of the tables,
            ' as they're not going to be returned by the stored procedure
            If CLng(Rs.Fields("PUNISHMENT_ID")) > 0 Then
                PunisherName = CStr(Rs.Fields("PUNISHER_NAME"))
                PunishmentEndDate = CDate(Rs.Fields("PUNISHMENT_END_DATE"))
                PunishmentDescription = CStr(Rs.Fields("PUNISHMENT_REASON"))
                
                LoadCharFromDB = False
                Exit Function
            End If
        
            
            ' Load the main user tables
            Set Rs = Rs.NextRecordset()
            
            'UserID = CLng(Rs.Fields("ID_USER"))
            .ID = UserId
            .Guild.IdGuild = CLng(Rs.Fields("GUILD_ID"))
            
            ' USER_FACTION
           
            With .Faccion
                .Alignment = CByte(Rs.Fields("ALIGNMENT"))
                .ArmadaReal = CByte(Rs.Fields("ARMY"))
                .FuerzasCaos = CByte(Rs.Fields("CHAOS"))
                .NeutralsKilled = CLng(Rs.Fields("NEUTRAL_KILLED"))
                .CiudadanosMatados = CLng(Rs.Fields("CITY_KILLED"))
                .CriminalesMatados = CLng(Rs.Fields("CRI_KILLED"))
                .RecibioArmaduraCaos = CByte(Rs.Fields("CHAOS_ARMOUR_GIVEN"))
                .RecibioArmaduraReal = CByte(Rs.Fields("ARMY_ARMOUR_GIVEN"))
                .RecibioExpInicialCaos = CByte(Rs.Fields("CHAOS_EXP_GIVEN"))
                .RecibioExpInicialReal = CByte(Rs.Fields("ARMY_EXP_GIVEN"))
                .RecompensasCaos = CLng(Rs.Fields("CHAOS_REWARD_GIVEN"))
                .RecompensasReal = CLng(Rs.Fields("ARMY_REWARD_GIVEN"))
                .Reenlistadas = CByte(Rs.Fields("NUM_SIGNS"))
                .NivelIngreso = CInt(Rs.Fields("SIGNING_LEVEL"))
                
                If Not IsNull(Rs.Fields("SIGNING_DATE")) Then
                    .FechaIngreso = CDate(Rs.Fields("SIGNING_DATE"))
                Else
                    .FechaIngreso = 0
                End If
                
                .MatadosIngreso = CInt(Rs.Fields("SIGNING_KILLED"))
                .NextRecompensa = CInt(Rs.Fields("NEXT_REWARD"))
            End With
            
            ' USER_FLAGS
            With .flags
                .Muerto = CByte(Rs.Fields("MUERTO"))
                .Escondido = CByte(Rs.Fields("ESCONDIDO"))
                
                .Hambre = CByte(Rs.Fields("HAMBRE"))
                .Sed = CByte(Rs.Fields("SED"))
                .Desnudo = CByte(Rs.Fields("DESNUDO"))
                .Navegando = CByte(Rs.Fields("NAVEGANDO"))
                .Envenenado = CByte(Rs.Fields("ENVENENADO"))
                .Paralizado = CByte(Rs.Fields("PARALIZADO"))
                
                'Matrix
                .lastMap = CInt(Rs.Fields("LAST_MAP"))
                
                If CByte(Rs.Fields("ROYAL_COUNCIL")) <> 0 Then _
                    .Privilegios = .Privilegios Or PlayerType.RoyalCouncil
                
                If CByte(Rs.Fields("CHAOS_COUNCIL")) <> 0 Then _
                    .Privilegios = .Privilegios Or PlayerType.ChaosCouncil
                    
                .LastTamedPet = CInt(Rs.Fields("LAST_TAMED_PET"))
            End With
            
            If .flags.Paralizado = 1 Then
                .Counters.Paralisis = SetIntervalEnd(ServerConfiguration.Intervals.IntervaloParalizado)
            End If
            
            ' Remaining time in jail
            .Counters.Pena = CLng(Rs.Fields("PUNISHMENT"))
            
            .Genero = CByte(Rs.Fields("GENDER"))
            .clase = CByte(Rs.Fields("CLASS"))
            .raza = CByte(Rs.Fields("RACE"))
            .Hogar = CByte(Rs.Fields("HOME"))
            .Char.heading = CInt(Rs.Fields("HEADING"))
            
            If .Char.heading > 4 Then
                .Char.heading = SOUTH
            End If
            
            ' Trainning Data
            .trainningData.trainningTime = CLng(Rs.Fields("TRAINNING_TIME"))
            .trainningData.startTick = GetTickCount()
        
            With .OrigChar
                .head = CInt(Rs.Fields("HEAD"))
                .body = CInt(Rs.Fields("BODY"))
                .WeaponAnim = CInt(Rs.Fields("WEAPON_ANIM"))
                .ShieldAnim = CInt(Rs.Fields("SHIELD_ANIM"))
                .CascoAnim = CInt(Rs.Fields("HELMET_ANIM"))
                
                .heading = CInt(Rs.Fields("HEADING"))
                
                If .heading > 4 Then
                    .heading = SOUTH
                End If
            End With
            
#If ConUpTime Then
            .UpTime = CLng(Rs.Fields("UP_TIME"))
#End If
            
            If .flags.Muerto = 0 Then
                .Char = .OrigChar
            Else
                .Char.body = ConstantesGRH.CuerpoMuerto
                .Char.head = ConstantesGRH.CabezaMuerto
                .Char.WeaponAnim = ConstantesGRH.NingunArma
                .Char.ShieldAnim = ConstantesGRH.NingunEscudo
                .Char.CascoAnim = ConstantesGRH.NingunCasco
            End If
            
            .desc = CStr(Rs.Fields("DESCRIP"))
            
            Dim sPos As String
            sPos = CStr(Rs.Fields("LAST_POS"))
            .Pos.Map = CInt(ReadField(1, sPos, 45))
            .Pos.X = CInt(ReadField(2, sPos, 45))
            .Pos.Y = CInt(ReadField(3, sPos, 45))
            
            .Guild.IdGuild = CInt(Rs.Fields("GUILD_ID"))
            .AspiranteA = CInt(Rs.Fields("REQUESTING_GUILD"))
            
            Dim sGuildRejectDetail As String
            
            'TODO GUILD remove that
            ' Sends message and update
            Dim Reason As String
            Reason = CStr(Rs.Fields("GUILD_REJECT_DETAIL"))
            If Len(Reason) > 0 Then
                Call WriteShowMessageBox(UserIndex, "Tu solicitud de ingreso al clan ha sido rechazada. El clan te explica que: " & Reason)
                Call UpdateCharInfo("GUILD_REJECT_DETAIL", CStr(Rs.Fields("NAME")), "")
            End If
            
            'Obtiene el indice-objeto del arma
            .Invent.WeaponEqpSlot = CByte(Rs.Fields("WEAPON_SLOT"))
            
            'Obtiene el indice-objeto del armadura
            .Invent.ArmourEqpSlot = CByte(Rs.Fields("ARMOUR_SLOT"))
            
            'Obtiene el indice-objeto del escudo
            .Invent.EscudoEqpSlot = CByte(Rs.Fields("SHIELD_SLOT"))
            
            'Obtiene el indice-objeto del casco
            .Invent.CascoEqpSlot = CByte(Rs.Fields("HELMET_SLOT"))
            
            'Obtiene el indice-objeto barco
            .Invent.BarcoSlot = CByte(Rs.Fields("BOAT_SLOT"))
            
            'Obtiene el indice-objeto municion
            .Invent.MunicionEqpSlot = CByte(Rs.Fields("MUNITION_SLOT"))
            
            'Obtiene el indice-objeto anilo
            .Invent.AnilloEqpSlot = CByte(Rs.Fields("RING_SLOT"))
            
            .Invent.MochilaEqpSlot = CByte(Rs.Fields("SACKPACK_SLOT"))
            
    
            ' STATS
            With .Stats
                .GLD = CLng(Rs.Fields("ORO"))
                .Banco = CLng(Rs.Fields("ORO_BANCO"))
                
                ' HP should be re-calculated on every load (done outside this function) as the character can have a mastery assigned
                .MaxHp = CInt(Rs.Fields("HP_MAX"))
                .MinHp = CInt(Rs.Fields("HP_MIN"))
                
                .MaxSta = CInt(Rs.Fields("STAMINA_MAX"))
                .MinSta = CInt(Rs.Fields("STAMINA_MIN"))
                
                .MaxMan = CInt(Rs.Fields("MANA_MAX"))
                .MinMAN = CInt(Rs.Fields("MANA_MIN"))
                
                .MaxAGU = CByte(Rs.Fields("AGUA_MAX"))
                .MinAGU = CByte(Rs.Fields("AGUA_MIN"))
                
                .MaxHam = CByte(Rs.Fields("HAMBRE_MAX"))
                .MinHam = CByte(Rs.Fields("HAMBRE_MIN"))
                
                .SkillPts = CInt(Rs.Fields("SKILLS"))
                
                .Exp = CDbl(Rs.Fields("EXP"))
                .ELU = CLng(Rs.Fields("EXP_NEXT"))
                .ELV = CByte(Rs.Fields("NIVEL"))
                
                .UsuariosMatados = CLng(Rs.Fields("USERS_KILLED"))
                .NPCsMuertos = CLng(Rs.Fields("NPCS_KILLED"))
                .MasteryPoints = CInt(Rs.Fields("MASTERY_POINTS"))
                .RankingPoints = CLng(Rs.Fields("RANKING_POINTS"))
                
                .DuelosGanados = CLng(Rs.Fields("DUELOS_GANADOS"))
                .DuelosPerdidos = CLng(Rs.Fields("DUELOS_PERDIDOS"))
                .OroDuelos = CLng(Rs.Fields("ORO_DUELOS"))
                
                ' ATTRIBUTES
                .UserAtributos(eAtributos.Fuerza) = CByte(Rs.Fields("STRENGHT"))
                .UserAtributosBackUP(eAtributos.Fuerza) = .UserAtributos(eAtributos.Fuerza)
                
                .UserAtributos(eAtributos.Agilidad) = CByte(Rs.Fields("DEXERITY"))
                .UserAtributosBackUP(eAtributos.Agilidad) = .UserAtributos(eAtributos.Agilidad)
                
                .UserAtributos(eAtributos.Inteligencia) = CByte(Rs.Fields("INTELLIGENCE"))
                .UserAtributosBackUP(eAtributos.Inteligencia) = .UserAtributos(eAtributos.Inteligencia)
                
                .UserAtributos(eAtributos.Carisma) = CByte(Rs.Fields("CHARISM"))
                .UserAtributosBackUP(eAtributos.Carisma) = .UserAtributos(eAtributos.Carisma)
                
                .UserAtributos(eAtributos.Constitucion) = CByte(Rs.Fields("HEALTH"))
                .UserAtributosBackUP(eAtributos.Constitucion) = .UserAtributos(eAtributos.Constitucion)
            End With
        End If
           
    End With
    
    ' Load the inventory and assign the right properties in the user's Invent.
    Set Rs = Rs.NextRecordset()
    Call LoadUserInventoryDB(UserIndex, Rs)
    
    With UserList(UserIndex)
        If .Invent.WeaponEqpSlot > 0 Then
            .Invent.WeaponEqpObjIndex = .Invent.Object(.Invent.WeaponEqpSlot).ObjIndex
        End If

        If .Invent.ArmourEqpSlot > 0 Then
            .Invent.ArmourEqpObjIndex = .Invent.Object(.Invent.ArmourEqpSlot).ObjIndex
            .flags.Desnudo = 0
        Else
            .flags.Desnudo = 1
        End If
        
        If .Invent.EscudoEqpSlot > 0 Then
            .Invent.EscudoEqpObjIndex = .Invent.Object(.Invent.EscudoEqpSlot).ObjIndex
        End If
    
        If .Invent.CascoEqpSlot > 0 Then
            .Invent.CascoEqpObjIndex = .Invent.Object(.Invent.CascoEqpSlot).ObjIndex
        End If
    
        If .Invent.BarcoSlot > 0 Then
            .Invent.BarcoObjIndex = .Invent.Object(.Invent.BarcoSlot).ObjIndex
        End If
    
        If .Invent.MunicionEqpSlot > 0 Then
            .Invent.MunicionEqpObjIndex = .Invent.Object(.Invent.MunicionEqpSlot).ObjIndex
        End If
    
        If .Invent.AnilloEqpSlot > 0 Then
            .Invent.AnilloEqpObjIndex = .Invent.Object(.Invent.AnilloEqpSlot).ObjIndex
        End If
    
        If .Invent.MochilaEqpSlot > 0 Then
            .Invent.MochilaEqpObjIndex = .Invent.Object(.Invent.MochilaEqpSlot).ObjIndex
        End If
    End With
    
    Set Rs = Rs.NextRecordset()
    
    Dim Slot As Byte
    
    With UserList(UserIndex).BancoInvent
        While Not Rs.EOF
            
            .NroItems = .NroItems + 1
            
            Slot = CByte(Rs.Fields("SLOT"))
            .Object(Slot).ObjIndex = CInt(Rs.Fields("OBJ_INDEX"))
            .Object(Slot).Amount = CInt(Rs.Fields("AMOUNT"))
            
            Rs.MoveNext
        Wend
    End With
    
    
    Set Rs = Rs.NextRecordset()
    
    With UserList(UserIndex).Stats
        While Not Rs.EOF
            
            Slot = CByte(Rs.Fields("SLOT"))
            .UserHechizos(Slot).SpellNumber = CInt(Rs.Fields("SPELL_INDEX"))
            
            Rs.MoveNext
        Wend
    End With
    
    
    Set Rs = Rs.NextRecordset()
    
    Dim skill As Byte
    
    With UserList(UserIndex).Stats
        While Not Rs.EOF
            
            skill = CByte(Rs.Fields("SKILL"))
            
            Call ZeroSkills(UserIndex, skill)
            Call AddNaturalSkills(UserIndex, skill, CByte(Rs.Fields("NATURAL_AMOUNT")))
            Call AddAssignedSkills(UserIndex, skill, CByte(Rs.Fields("ASSIGNED_AMOUNT")))
            
            .EluSkills(skill) = CLng(Rs.Fields("SKILL_EXP_NEXT_LEVEL"))
            .ExpSkills(skill) = CLng(Rs.Fields("SKILL_EXP"))
            
            Rs.MoveNext
        Wend
    End With
    
    
    Set Rs = Rs.NextRecordset()
    
    
    With UserList(UserIndex)
        If (Classes(.clase).ClassMods.MaxTammedPets > 0) Then
            ReDim .TammedPets(1 To Classes(.clase).ClassMods.MaxTammedPets) As tPet
        End If
        
        If (Classes(.clase).ClassMods.MaxInvokedPets > 0) Then
            ReDim .InvokedPets(1 To Classes(.clase).ClassMods.MaxInvokedPets) As tPet
        End If
    
        .TammedPetsCount = 0
        .InvokedPetsCount = 0
    
        While Not Rs.EOF
            
            .TammedPetsCount = .TammedPetsCount + 1
            
            Slot = CByte(Rs.Fields("NUM_PET"))
            .TammedPets(Slot).NpcIndex = CInt(Rs.Fields("NPC_INDEX"))
            .TammedPets(Slot).NpcNumber = CInt(Rs.Fields("NPC_TYPE"))
            .TammedPets(Slot).RemainingLife = CInt(Rs.Fields("NPC_LIFE"))
            
            Rs.MoveNext
        Wend
    End With
    
    Set Rs = Rs.NextRecordset()
    Call LoadUserMasteriesDB(UserIndex, Rs)
    
    ' Finish, close the connection and release the objects
    Rs.Close
    Set Rs = Nothing
    Set Cmd = Nothing
    
    LoadCharFromDB = True
    
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") para personaje con ID " & UserId & " en Sub LoadCharFromDB de modDB_CharLoad.bas")
End Function

Private Sub LoadUserInventoryDB(ByVal UserIndex As Integer, ByRef Rs As Recordset)
'***************************************************
'Author: ZaMa
'Last Modification: 25/08/2012
'Loads user inventory info from DB
'***************************************************
On Error GoTo ErrHandler
    Dim Slot As Byte
    
    With UserList(UserIndex).Invent
        While Not Rs.EOF
            
            .NroItems = .NroItems + 1
            
            Slot = CByte(Rs.Fields("SLOT"))
            .Object(Slot).ObjIndex = CInt(Rs.Fields("OBJ_INDEX"))
            .Object(Slot).Amount = CInt(Rs.Fields("AMOUNT"))
            .Object(Slot).Equipped = CByte(Rs.Fields("EQUIPPED"))
            
            Rs.MoveNext
        Wend
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadUserInventoryDB de modDB_CharLoad.bas")
End Sub

Public Sub LoadUserBankDB(ByVal UserIndex As Integer, ByVal UserID As Long)
'***************************************************
'Author: ZaMa
'Last Modification: 25/08/2012
'Loads user bank inventory info from DB
'***************************************************
On Error GoTo ErrHandler
  
    
    Dim Sql As String
    Sql = _
        "SELECT " & _
            "SLOT, " & _
            "OBJ_INDEX, " & _
            "AMOUNT " & _
        "FROM " & _
            "USER_BANK " & _
        "WHERE ID_USER = '" & CStr(UserID) & "' "
    
    Dim Rs As Recordset
    Set Rs = ExecuteSql(Sql)
    
    Dim Slot As Byte
    
    With UserList(UserIndex).BancoInvent
        While Not Rs.EOF
            
            .NroItems = .NroItems + 1
            
            Slot = CByte(Rs.Fields("SLOT"))
            .Object(Slot).ObjIndex = CInt(Rs.Fields("OBJ_INDEX"))
            .Object(Slot).Amount = CInt(Rs.Fields("AMOUNT"))
            
            Rs.MoveNext
        Wend
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadUserBankDB de modDB_CharLoad.bas")
End Sub

Public Sub LoadUserSpellsDB(ByVal UserIndex As Integer, ByVal UserID As Long)
'***************************************************
'Author: ZaMa
'Last Modification: 25/08/2012
'Loads user spells info from DB
'***************************************************
On Error GoTo ErrHandler
  
    
    Dim Sql As String
    Sql = _
        "SELECT " & _
            "SLOT, " & _
            "SPELL_INDEX " & _
        "FROM " & _
            "USER_SPELLS " & _
        "WHERE ID_USER = '" & CStr(UserID) & "' "
    
    Dim Rs As Recordset
    Set Rs = ExecuteSql(Sql)
    
    Dim Slot As Byte
    
    With UserList(UserIndex).Stats
        While Not Rs.EOF
            
            Slot = CByte(Rs.Fields("SLOT"))
            .UserHechizos(Slot).SpellNumber = CInt(Rs.Fields("SPELL_INDEX"))
            
            Rs.MoveNext
        Wend
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadUserSpellsDB de modDB_CharLoad.bas")
End Sub

Public Sub LoadUserSkillsDB(ByVal UserIndex As Integer, ByVal UserID As Long)
'***************************************************
'Author: ZaMa
'Last Modification: 25/08/2012
'Loads user skills info from DB
'***************************************************
On Error GoTo ErrHandler
  
    
    Dim Sql As String
    Sql = _
        "SELECT " & _
            "SKILL, " & _
            "NATURAL_AMOUNT, " & _
            "ASSIGNED_AMOUNT, " & _
            "SKILL_EXP_NEXT_LEVEL, " & _
            "SKILL_EXP " & _
        "FROM " & _
            "USER_SKILLS " & _
        "WHERE ID_USER = '" & CStr(UserID) & "' "
    
    Dim Rs As Recordset
    Set Rs = ExecuteSql(Sql)
    
    Dim skill As Byte
    
    With UserList(UserIndex).Stats
        While Not Rs.EOF
            
            skill = CByte(Rs.Fields("SKILL"))
            
            Call ZeroSkills(UserIndex, skill)
            Call AddNaturalSkills(UserIndex, skill, CByte(Rs.Fields("NATURAL_AMOUNT")))
            Call AddAssignedSkills(UserIndex, skill, CByte(Rs.Fields("ASSIGNED_AMOUNT")))
            
            .EluSkills(skill) = CLng(Rs.Fields("SKILL_EXP_NEXT_LEVEL"))
            .ExpSkills(skill) = CLng(Rs.Fields("SKILL_EXP"))
            
            Rs.MoveNext
        Wend
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadUserSkillsDB de modDB_CharLoad.bas")
End Sub

Public Sub LoadUserPetsDB(ByVal UserIndex As Integer, ByVal UserID As Long)
'***************************************************
'Author: ZaMa
'Last Modification: 25/08/2012
'Loads user pets info from DB
'***************************************************
On Error GoTo ErrHandler
  
    
    Dim Sql As String
    Sql = _
        "SELECT " & _
            "NUM_PET, " & _
            "NPC_INDEX, " & _
            "NPC_TYPE, " & _
            "NPC_LIFE " & _
        "FROM " & _
            "USER_PETS " & _
        "WHERE ID_USER = '" & CStr(UserID) & "' "
    
    Dim Rs As Recordset
    Set Rs = ExecuteSql(Sql)
    
    Dim Slot As Byte
    
    With UserList(UserIndex)
        While Not Rs.EOF
            
            .TammedPetsCount = .TammedPetsCount + 1
            
            Slot = CByte(Rs.Fields("NUM_PET"))
            .TammedPets(Slot).NpcIndex = CInt(Rs.Fields("NPC_INDEX"))
            .TammedPets(Slot).NpcNumber = CInt(Rs.Fields("NPC_TYPE"))
            .TammedPets(Slot).RemainingLife = CInt(Rs.Fields("NPC_LIFE"))
            
            Rs.MoveNext
        Wend
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadUserPetsDB de modDB_CharLoad.bas")
End Sub

Public Sub LoadUserMessagesDB(ByVal UserIndex As Integer, ByVal UserID As Long)
'***************************************************
'Author: ZaMa
'Last Modification: 12/09/2012
'Loads user private messages from DB
'***************************************************
On Error GoTo ErrHandler
  
    
    Dim Sql As String
    Sql = _
        "SELECT " & _
            "MSG_INDEX, " & _
            "MESSAGE, " & _
            "UNREAD " & _
        "FROM " & _
            "USER_MESSAGES " & _
        "WHERE ID_USER = '" & CStr(UserID) & "' " & _
        "ORDER BY MSG_INDEX "
    
    Dim Rs As Recordset
    Set Rs = ExecuteSql(Sql)
    
    Dim Slot As Byte
    
    With UserList(UserIndex)
        While Not Rs.EOF
            
            .UltimoMensaje = CByte(Rs.Fields("MSG_INDEX"))
            With .Mensajes(.UltimoMensaje)
                .Contenido = CStr(Rs.Fields("MESSAGE"))
                .Contenido = (CByte(Rs.Fields("UNREAD")) = 1)
            End With
            
            Rs.MoveNext
        Wend
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadUserMessagesDB de modDB_CharLoad.bas")
End Sub

Public Sub LoadUserPassivesDB(ByVal UserIndex As Integer, ByVal UserId As Long)
'***************************************************
'Author: Lucas Figelj(Luke)
'Last Modification: 21/04/2015
'Loads user passive skills info from DB
'***************************************************
On Error GoTo ErrHandler
  
    
    Dim Sql As String
    Sql = _
        "SELECT " & _
            "ID_PASSIVE " & _
            ",SLOT " & _
        "FROM " & _
            "USER_PASSIVE_SPELLS " & _
        "WHERE ID_USER = '" & CStr(UserID) & "' "
    
    Dim Rs As Recordset
    Set Rs = ExecuteSql(Sql)
    
    Dim Slot As Byte
    
    With UserList(UserIndex).Stats
        While Not Rs.EOF
            
            Slot = CByte(Rs.Fields("SLOT"))
            .UserPassives(Slot).ID = CInt(Rs.Fields("ID_PASSIVE"))
            
            Rs.MoveNext
        Wend
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadUserPassivesDB de modDB_CharLoad.bas")
End Sub

Public Sub LoadUserMasteriesDB(ByVal UserIndex As Integer, ByRef Rs As Recordset)
On Error GoTo ErrHandler
    
    With UserList(UserIndex)
    
        If Classes(.clase).MasteryGroupsQty <= 0 Then Exit Sub
  
        .Masteries.GroupsQty = Classes(.clase).MasteryGroupsQty
        
        Erase .Masteries.Groups
        
        ReDim .Masteries.Groups(1 To .Masteries.GroupsQty)
        .Masteries.GroupsQty = .Masteries.GroupsQty
        
        If Rs.RecordCount = 0 Then
            Exit Sub
        End If
        
        Dim TmpGroups() As String
        ReDim TmpGroups(1 To .Masteries.GroupsQty)
        
        Dim I As Integer
        Dim J As Integer
        Dim CurrentGroup As Integer
        Dim SplittedMasteriesTmp() As String
        Dim StrUnsplittedMasteries As String
        I = 1
        While Not Rs.EOF
            CurrentGroup = CInt(Rs.Fields("ID_MASTERY_GROUP"))
            If CurrentGroup <= UBound(TmpGroups) Then
                StrUnsplittedMasteries = Rs.Fields("MASTERIES")
                                        
                SplittedMasteriesTmp = Split(StrUnsplittedMasteries, ",")
                
                .Masteries.Groups(CurrentGroup).GroupId = CurrentGroup
                .Masteries.Groups(CurrentGroup).MasteriesQty = UBound(SplittedMasteriesTmp) + 1 ' Split uses base 0 for arrays
                
                ReDim .Masteries.Groups(CurrentGroup).Masteries(1 To .Masteries.Groups(CurrentGroup).MasteriesQty)
                 
                For J = 1 To .Masteries.Groups(CurrentGroup).MasteriesQty
                    .Masteries.Groups(CurrentGroup).Masteries(J).ID = SplittedMasteriesTmp(J - 1)
                    
                    Call AssignMasteryPropertiesToUser(UserIndex, .Masteries.Groups(CurrentGroup).Masteries(J).ID)
                Next J
                
            End If
            
            Rs.MoveNext
        Wend
        
        Exit Sub
        
        For I = 1 To UBound(.Masteries.Groups)
            If TmpGroups(I) <> vbNullString Then
                SplittedMasteriesTmp = Split(TmpGroups(I), ",")
                
                .Masteries.Groups(I).MasteriesQty = UBound(SplittedMasteriesTmp) + 1 ' Split uses base 0 for arrays
                
                ReDim .Masteries.Groups(I).Masteries(1 To .Masteries.Groups(I).MasteriesQty)
                
                For J = 1 To .Masteries.Groups(I).MasteriesQty
                    .Masteries.Groups(I).Masteries(J).ID = SplittedMasteriesTmp(J - 1)

                    Call AssignMasteryPropertiesToUser(UserIndex, .Masteries.Groups(I).Masteries(J).ID)
                Next J
            Else
                .Masteries.Groups(I).MasteriesQty = 0
            End If
        Next I
    
    End With

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadUserMasteriesDB de modDB_CharLoad.bas")
End Sub
