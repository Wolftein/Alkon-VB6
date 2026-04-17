Attribute VB_Name = "modDB_Functions"
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


Public Sub SaveUserDB(ByVal UserIndex As Integer, ByVal SaveTimeOnline As Boolean, _
    ByVal NewChar As Boolean, ByRef Password As String)
'***************************************************
'Author: ZaMa
'Creation Date: 09/06/2012
'Last Modification: -
'Saves user to DB
'***************************************************

On Error GoTo ErrHandler

    With UserList(UserIndex)
        'ESTO TIENE QUE EVITAR ESE BUGAZO QUE NO SE POR QUE GRABA USUARIOS NULOS
        'clase=0 es el error, porq el enum empieza de 1!!
        If .clase = 0 Or .Stats.ELV = 0 Then
            Call LogCriticEvent("Estoy intentando guardar un usuario nulo de nombre: " & .Name)
            Exit Sub
        End If
        
        ' Remove mimic
        If .flags.Mimetizado <> 0 Then Call EndMimic(UserIndex, False, False)
    
        ' Header
        Call SaveUserHeaderDB(UserIndex, NewChar, SaveTimeOnline)
        
        ' Skills
        Call SaveUserSkillsDB(UserIndex, .ID, NewChar)
        ' Quest info
        'Call SaveUserQuestInfoDB(UserIndex, .ID, NewChar)
        ' Bank
        Call SaveUserBankDB(UserIndex, .ID, NewChar)
        ' Inventory
        Call SaveUserInventoryDB(UserIndex, .ID, NewChar)
        ' Spells
        Call SaveUserSpellsDB(UserIndex, .ID, NewChar)
        ' Pets
        Call SaveUserPetsDB(UserIndex, .ID, NewChar)
        
        ' Attributes
        Call SaveUserAttributesDB(UserIndex, NewChar)
        ' Account Bank
        Call SaveAccountBankDB(UserIndex, NewChar)
    
        If Not NewChar Then
            ' Private Messages
            Call SaveUserPrivateMsjDB(UserIndex, .ID)
        End If
        
        'Devuelve el head de muerto
        If .flags.Muerto = 1 Then .Char.head = ConstantesGRH.CabezaMuerto
    End With
    
    Exit Sub

ErrHandler:
    Call LogError("Error en SaveUserDB. Error" & Err.Description)
End Sub


Private Sub SaveUserInventoryDB(ByVal UserIndex As Integer, ByVal ID As Long, ByVal NewChar As Boolean)
'***************************************************
'Author: ZaMa
'Creation Date: 09/06/2012
'Last Modification: -
'Saves user inventory into DB
'***************************************************
On Error GoTo ErrHandler:

    With UserList(UserIndex).Invent
        
        Dim Sql As String
        
        ' Delete previous
        If Not NewChar Then
            Sql = _
                "DELETE FROM USER_INVENTORY " & _
                "WHERE ID_USER = '" & CStr(ID) & "' "
                
            Call ExecuteSql(Sql)
        End If
        
        ' Store actual
        Dim Slot As Long
        Dim ProcessedCount As Byte
        ProcessedCount = 0
        
        Sql = "INSERT INTO USER_INVENTORY VALUES "
        
        For Slot = 1 To MAX_INVENTORY_SLOTS
            
            With .Object(Slot)
                If .ObjIndex <> 0 Then
                
                    If ProcessedCount > 0 Then
                        Sql = Sql & ","
                    End If
                    
                    Sql = Sql & _
                        "('" & _
                            CStr(ID) & "','" & _
                            CStr(Slot) & "','" & _
                            CStr(.ObjIndex) & "','" & _
                            CStr(.Amount) & "','" & _
                            CStr(.Equipped) & "' " & _
                        ")"
                    
                    ProcessedCount = ProcessedCount + 1

                End If
            End With
        Next Slot
    End With
    
    If ProcessedCount > 0 Then Call ExecuteSql(Sql)
    
    Exit Sub
ErrHandler:
    LogError ("Error en SaveUserInventoryDB: " & Err.Description)
End Sub

Private Sub SaveUserBankDB(ByVal UserIndex As Integer, ByVal ID As Long, ByVal NewChar As Boolean)
'***************************************************
'Author: ZaMa
'Creation Date: 09/06/2012
'Last Modification: -
'Saves user bank inventory into DB
'***************************************************
On Error GoTo ErrHandler:

    With UserList(UserIndex).BancoInvent
        
        Dim Sql As String
        
        ' Delete previous
        If Not NewChar Then
            Sql = _
                "DELETE FROM USER_BANK " & _
                "WHERE ID_USER = '" & CStr(ID) & "' "
                
            Call ExecuteSql(Sql)
        End If
        
        ' Store actual
        
        Sql = "INSERT INTO USER_BANK VALUES "
        
        Dim Slot As Long
        Dim ProcessedCount As Byte
        ProcessedCount = 0
        For Slot = 1 To MAX_BANCOINVENTORY_SLOTS
            With .Object(Slot)
                If .ObjIndex <> 0 Then
                
                    If ProcessedCount > 0 Then
                        Sql = Sql & ","
                    End If
                
                    Sql = Sql & _
                        "('" & _
                            CStr(ID) & "','" & _
                            CStr(Slot) & "','" & _
                            CStr(.ObjIndex) & "','" & _
                            CStr(.Amount) & "' " & _
                        ")"
                    
                    ProcessedCount = ProcessedCount + 1
                End If
            End With
        Next Slot
    End With
    
    If ProcessedCount > 0 Then Call ExecuteSql(Sql)
    
    Exit Sub
ErrHandler:
    LogError ("Error en SaveUserBankDB: " & Err.Description)
End Sub

Private Sub SaveUserSpellsDB(ByVal UserIndex As Integer, ByVal ID As Long, ByVal NewChar As Boolean)
'***************************************************
'Author: ZaMa
'Creation Date: 09/06/2012
'Last Modification: -
'Saves user spells into DB
'***************************************************
On Error GoTo ErrHandler:

    With UserList(UserIndex).Stats
        
        Dim Sql As String
        
        ' Delete previous
        If Not NewChar Then
            Sql = _
                "DELETE FROM USER_SPELLS " & _
                "WHERE ID_USER = '" & CStr(ID) & "' "
                
            Call ExecuteSql(Sql)
        End If
        
        
        
        
        ' Store actual
        Dim Slot As Long
        Dim ProcessedCount As Byte
        ProcessedCount = 0
        
        Sql = "INSERT INTO USER_SPELLS VALUES"
        
        For Slot = 1 To MAXUSERHECHIZOS
            If .UserHechizos(Slot).SpellNumber <> 0 Then
            
                If ProcessedCount > 0 Then
                    Sql = Sql & ","
                End If
            
                Sql = Sql & _
                    "('" & _
                        CStr(ID) & "','" & _
                        CStr(Slot) & "','" & _
                        CStr(.UserHechizos(Slot).SpellNumber) & "' " & _
                    ")"
                
                ProcessedCount = ProcessedCount + 1
            End If
        Next Slot
    End With
    
    If ProcessedCount > 0 Then Call ExecuteSql(Sql)
    
    Exit Sub
ErrHandler:
    LogError ("Error en SaveUserSpellsDB: " & Err.Description)
End Sub

Private Sub SaveUserSkillsDB(ByVal UserIndex As Integer, ByVal ID As Long, ByVal NewChar As Boolean)
'***************************************************
'Author: ZaMa
'Creation Date: 09/06/2012
'Last Modification: -
'Saves user skills into DB
'***************************************************
On Error GoTo ErrHandler:

    With UserList(UserIndex).Stats
        
        Dim Sql As String
        
        ' Delete previous
        If Not NewChar Then
            Sql = _
                "DELETE FROM USER_SKILLS " & _
                "WHERE ID_USER = '" & CStr(ID) & "' "
                
            Call ExecuteSql(Sql)
        End If
        

        ' Store actual
        Dim skill As Long
        Dim ProcessedCount As Byte
        ProcessedCount = 0
        
        Sql = "INSERT INTO USER_SKILLS VALUES "
        For skill = 1 To NUMSKILLS
            If (GetSkills(UserIndex, skill) <> 0) Or (.ExpSkills(skill) <> 0) Or (.EluSkills(skill) <> 0) Then
            
                If ProcessedCount > 0 Then
                    Sql = Sql & ","
                End If
                
                Sql = Sql & _
                    "('" & _
                        CStr(ID) & "','" & _
                        CStr(skill) & "','" & _
                        CStr(GetNaturalSkills(UserIndex, skill)) & "','" & _
                        CStr(GetAssignedSkills(UserIndex, skill)) & "','" & _
                        CStr(.EluSkills(skill)) & "','" & _
                        CStr(.ExpSkills(skill)) & "' " & _
                    ")"
                          
                ProcessedCount = ProcessedCount + 1
                
            End If
        Next skill
    End With
    
    If ProcessedCount > 0 Then Call ExecuteSql(Sql)
    
    Exit Sub
    
ErrHandler:
    LogError ("Error en SaveUserSkillsDB: " & Err.Description)
End Sub

Private Sub SaveUserPetsDB(ByVal UserIndex As Integer, ByVal ID As Long, ByVal NewChar As Boolean)
'***************************************************
'Author: ZaMa
'Creation Date: 09/06/2012
'Last Modification: -
'Saves user pets into DB
'***************************************************
On Error GoTo ErrHandler:

    With UserList(UserIndex)
        
        Dim Sql As String
        
        ' Delete previous
        If Not NewChar Then
            Sql = _
                "DELETE FROM USER_PETS " & _
                "WHERE ID_USER = '" & CStr(ID) & "' "
                
            Call ExecuteSql(Sql)
        
        
            ' Store actual
            Dim Index As Long
            Dim StorePet As Boolean
            Dim ProcessedCount As Byte
            ProcessedCount = 0
            
            Sql = "INSERT INTO USER_PETS VALUES "

            For Index = 1 To Classes(.clase).ClassMods.MaxTammedPets
                StorePet = False
                If (.TammedPets(Index).NpcIndex <> 0) Then
                    ' Don't save summons
                    StorePet = (Npclist(.TammedPets(Index).NpcIndex).Contadores.TiempoExistencia = 0)
                ElseIf (.TammedPets(Index).NpcNumber <> 0) Then
                    StorePet = True
                End If
                
                ' TODO: Eliminar NpcIndex de la db, no tiene sentido la columna
                If StorePet Then
                
                    If ProcessedCount > 0 Then
                        Sql = Sql & ","
                    End If
                
                    Sql = Sql & _
                        "('" & _
                            CStr(ID) & "','" & _
                            CStr(Index) & "','" & _
                            "0','" & _
                            CStr(.TammedPets(Index).NpcNumber) & "','" & _
                            CStr(.TammedPets(Index).RemainingLife) & "' " & _
                        ")"
                        
                    ProcessedCount = ProcessedCount + 1
                End If
            Next Index
            
        End If
    End With
    
    If ProcessedCount > 0 Then Call ExecuteSql(Sql)
    
    Exit Sub
ErrHandler:
    LogError ("Error en SaveUserPetsDB: " & Err.Description)
End Sub

Private Function SaveUserHeaderDB(ByVal UserIndex As Integer, ByVal NewChar As Boolean, _
                                  ByVal SaveTimeOnline As Boolean) As Long

    Dim Rs As Recordset
    Dim Cmd As ADODB.Command
    Set Cmd = New ADODB.Command
        
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "sp_SaveCharacterHeader"
    
    Dim BodyToSave As Integer
    
    
    With UserList(UserIndex)
       
        #If ConUpTime Then
            If SaveTimeOnline Then
                Dim TempDate As Date
                TempDate = Now - .LogOnTime
                .LogOnTime = Now
                .UpTime = .UpTime + (Abs(Day(TempDate) - 30) * 24 * 3600) + Hour(TempDate) * 3600 + Minute(TempDate) * 60 + Second(TempDate)
            End If
        #End If
     
        ' USER_INFO
        Cmd.Parameters.Append Cmd.CreateParameter("UserID", DataTypeEnum.adInteger, ParameterDirectionEnum.adParamInput, 1, .ID)
        Cmd.Parameters.Append Cmd.CreateParameter("UserName", DataTypeEnum.adBSTR, ParameterDirectionEnum.adParamInput, 1, .Name)
        Cmd.Parameters.Append Cmd.CreateParameter("Gender", DataTypeEnum.adTinyInt, ParameterDirectionEnum.adParamInput, 1, .Genero)
        Cmd.Parameters.Append Cmd.CreateParameter("Race", DataTypeEnum.adTinyInt, ParameterDirectionEnum.adParamInput, 1, .raza)
        Cmd.Parameters.Append Cmd.CreateParameter("Class", DataTypeEnum.adTinyInt, ParameterDirectionEnum.adParamInput, 1, .clase)
        
        Cmd.Parameters.Append Cmd.CreateParameter("Home", DataTypeEnum.adTinyInt, ParameterDirectionEnum.adParamInput, 1, .Hogar)
        Cmd.Parameters.Append Cmd.CreateParameter("CharDescription", DataTypeEnum.adBSTR, ParameterDirectionEnum.adParamInput, 1, .desc)
        Cmd.Parameters.Append Cmd.CreateParameter("Punishment", DataTypeEnum.adTinyInt, ParameterDirectionEnum.adParamInput, 1, .Counters.Pena)
        Cmd.Parameters.Append Cmd.CreateParameter("Heading", DataTypeEnum.adTinyInt, ParameterDirectionEnum.adParamInput, 1, .Char.heading)
        
        Cmd.Parameters.Append Cmd.CreateParameter("Head", DataTypeEnum.adSmallInt, ParameterDirectionEnum.adParamInput, 1, .OrigChar.head)
        Cmd.Parameters.Append Cmd.CreateParameter("Body", DataTypeEnum.adSmallInt, ParameterDirectionEnum.adParamInput, 1, _
                                                        IIf(.flags.Muerto = 0 And .Char.body <> 0, .Char.body, .OrigChar.body))
            
        Cmd.Parameters.Append Cmd.CreateParameter("WeaponAnim", DataTypeEnum.adSmallInt, ParameterDirectionEnum.adParamInput, 1, .Char.WeaponAnim)
        Cmd.Parameters.Append Cmd.CreateParameter("ShieldAnim", DataTypeEnum.adSmallInt, ParameterDirectionEnum.adParamInput, 1, .Char.ShieldAnim)
        
        Cmd.Parameters.Append Cmd.CreateParameter("HelmetAnim", DataTypeEnum.adSmallInt, ParameterDirectionEnum.adParamInput, 1, .Char.CascoAnim)
        Cmd.Parameters.Append Cmd.CreateParameter("Uptime", DataTypeEnum.adInteger, ParameterDirectionEnum.adParamInput, 1, .UpTime)
        Cmd.Parameters.Append Cmd.CreateParameter("LastIP", DataTypeEnum.adBSTR, ParameterDirectionEnum.adParamInput, 1, .IP)
        Cmd.Parameters.Append Cmd.CreateParameter("LastPoss", DataTypeEnum.adBSTR, ParameterDirectionEnum.adParamInput, 1, .Pos.Map & "-" & .Pos.X & "-" & .Pos.Y)
        
        Cmd.Parameters.Append Cmd.CreateParameter("WeaponSlot", DataTypeEnum.adTinyInt, ParameterDirectionEnum.adParamInput, 1, .Invent.WeaponEqpSlot)
        Cmd.Parameters.Append Cmd.CreateParameter("ArmorSlot", DataTypeEnum.adTinyInt, ParameterDirectionEnum.adParamInput, 1, .Invent.ArmourEqpSlot)
        Cmd.Parameters.Append Cmd.CreateParameter("HelmetSlot", DataTypeEnum.adTinyInt, ParameterDirectionEnum.adParamInput, 1, .Invent.CascoEqpSlot)
        Cmd.Parameters.Append Cmd.CreateParameter("ShieldSlot", DataTypeEnum.adTinyInt, ParameterDirectionEnum.adParamInput, 1, .Invent.EscudoEqpSlot)
        
        Cmd.Parameters.Append Cmd.CreateParameter("BoatSlot", DataTypeEnum.adTinyInt, ParameterDirectionEnum.adParamInput, 1, .Invent.BarcoSlot)
        Cmd.Parameters.Append Cmd.CreateParameter("AmmoSlot", DataTypeEnum.adTinyInt, ParameterDirectionEnum.adParamInput, 1, .Invent.MunicionEqpSlot)
        Cmd.Parameters.Append Cmd.CreateParameter("BackpackSlot", DataTypeEnum.adTinyInt, ParameterDirectionEnum.adParamInput, 1, .Invent.MochilaEqpSlot)
        Cmd.Parameters.Append Cmd.CreateParameter("RingSlot", DataTypeEnum.adTinyInt, ParameterDirectionEnum.adParamInput, 1, .Invent.AnilloEqpSlot)
        
        Cmd.Parameters.Append Cmd.CreateParameter("GuildId", DataTypeEnum.adSmallInt, ParameterDirectionEnum.adParamInput, 1, .Guild.IdGuild)
        Cmd.Parameters.Append Cmd.CreateParameter("RequestingGuild", DataTypeEnum.adSmallInt, ParameterDirectionEnum.adParamInput, 1, .AspiranteA)
        Cmd.Parameters.Append Cmd.CreateParameter("IsBanned", DataTypeEnum.adTinyInt, adParamInput, 1, .flags.Ban)
        
        Cmd.Parameters.Append Cmd.CreateParameter("LastPunishmentId", DataTypeEnum.adInteger, ParameterDirectionEnum.adParamInput, 1, .Punishment.ID)
        Cmd.Parameters.Append Cmd.CreateParameter("TrainingTime", DataTypeEnum.adInteger, ParameterDirectionEnum.adParamInput, 1, .trainningData.trainningTime)
        Cmd.Parameters.Append Cmd.CreateParameter("AccountId", DataTypeEnum.adInteger, adParamInput, 1, .AccountId)
    
        ' USER_FLAGS
        Cmd.Parameters.Append Cmd.CreateParameter("IsDead", DataTypeEnum.adTinyInt, adParamInput, 1, .flags.Muerto)
        Cmd.Parameters.Append Cmd.CreateParameter("IsHidding", DataTypeEnum.adTinyInt, ParameterDirectionEnum.adParamInput, 1, .flags.Escondido)
        Cmd.Parameters.Append Cmd.CreateParameter("Hunger", DataTypeEnum.adUnsignedTinyInt, ParameterDirectionEnum.adParamInput, 1, .flags.Hambre)
        Cmd.Parameters.Append Cmd.CreateParameter("Thirst", DataTypeEnum.adUnsignedTinyInt, adParamInput, 1, .flags.Sed)
        
        Cmd.Parameters.Append Cmd.CreateParameter("Naked", DataTypeEnum.adTinyInt, adParamInput, 1, .flags.Desnudo)
        Cmd.Parameters.Append Cmd.CreateParameter("Sailing", DataTypeEnum.adTinyInt, ParameterDirectionEnum.adParamInput, 1, .flags.Navegando)
        Cmd.Parameters.Append Cmd.CreateParameter("Poisoned", DataTypeEnum.adTinyInt, ParameterDirectionEnum.adParamInput, 1, .flags.Envenenado)
        Cmd.Parameters.Append Cmd.CreateParameter("Paralized", DataTypeEnum.adTinyInt, adParamInput, 1, .flags.Paralizado)
        
        Cmd.Parameters.Append Cmd.CreateParameter("LastMap", DataTypeEnum.adSmallInt, adParamInput, 1, .flags.lastMap)
        Cmd.Parameters.Append Cmd.CreateParameter("LastTammedPet", DataTypeEnum.adBigInt, ParameterDirectionEnum.adParamInput, 1, .flags.LastTamedPet)
        
        
        ' USER_STATS
        Cmd.Parameters.Append Cmd.CreateParameter("Gold", DataTypeEnum.adInteger, adParamInput, 1, .Stats.GLD)
        Cmd.Parameters.Append Cmd.CreateParameter("GoldBank", DataTypeEnum.adInteger, adParamInput, 1, .Stats.Banco)
        Cmd.Parameters.Append Cmd.CreateParameter("HpMax", DataTypeEnum.adSmallInt, adParamInput, 1, .Stats.MaxHp)
        Cmd.Parameters.Append Cmd.CreateParameter("HpMin", DataTypeEnum.adSmallInt, adParamInput, 1, .Stats.MinHp)
        
        Cmd.Parameters.Append Cmd.CreateParameter("StaminaMax", DataTypeEnum.adSmallInt, adParamInput, 1, .Stats.MaxSta)
        Cmd.Parameters.Append Cmd.CreateParameter("StaminaMin", DataTypeEnum.adSmallInt, adParamInput, 1, .Stats.MinSta)
        Cmd.Parameters.Append Cmd.CreateParameter("ManaMax", DataTypeEnum.adSmallInt, adParamInput, 1, .Stats.MaxMan)
        Cmd.Parameters.Append Cmd.CreateParameter("ManaMin", DataTypeEnum.adSmallInt, adParamInput, 1, .Stats.MinMAN)

        Cmd.Parameters.Append Cmd.CreateParameter("ThirstMax", DataTypeEnum.adSmallInt, adParamInput, 1, .Stats.MaxAGU)
        Cmd.Parameters.Append Cmd.CreateParameter("ThirstMin", DataTypeEnum.adSmallInt, adParamInput, 1, .Stats.MinAGU)
        
        Cmd.Parameters.Append Cmd.CreateParameter("HungerMax", DataTypeEnum.adSmallInt, adParamInput, 1, .Stats.MaxHam)
        Cmd.Parameters.Append Cmd.CreateParameter("HungerMin", DataTypeEnum.adSmallInt, adParamInput, 1, .Stats.MinHam)
        Cmd.Parameters.Append Cmd.CreateParameter("Skills", DataTypeEnum.adSmallInt, adParamInput, 1, .Stats.SkillPts)
        Cmd.Parameters.Append Cmd.CreateParameter("Exp", DataTypeEnum.adBigInt, adParamInput, 1, .Stats.Exp)
        
        Cmd.Parameters.Append Cmd.CreateParameter("Level", DataTypeEnum.adTinyInt, adParamInput, 1, .Stats.ELV)
        Cmd.Parameters.Append Cmd.CreateParameter("ExpNextLevel", DataTypeEnum.adInteger, adParamInput, 1, .Stats.ELU)
        Cmd.Parameters.Append Cmd.CreateParameter("UsersKilled", DataTypeEnum.adInteger, adParamInput, 1, .Stats.UsuariosMatados)
        Cmd.Parameters.Append Cmd.CreateParameter("NpcsKilled", DataTypeEnum.adInteger, adParamInput, 1, .Stats.NPCsMuertos)
        
        Cmd.Parameters.Append Cmd.CreateParameter("RankingPoints", DataTypeEnum.adInteger, adParamInput, 1, .Stats.RankingPoints)
        Cmd.Parameters.Append Cmd.CreateParameter("MasteryPoints", DataTypeEnum.adInteger, adParamInput, 1, .Stats.MasteryPoints)
        Cmd.Parameters.Append Cmd.CreateParameter("DuelsWon", DataTypeEnum.adInteger, adParamInput, 1, .Stats.DuelosGanados)
        Cmd.Parameters.Append Cmd.CreateParameter("DuelsLost", DataTypeEnum.adInteger, adParamInput, 1, .Stats.DuelosPerdidos)
        Cmd.Parameters.Append Cmd.CreateParameter("DuelsGoldWon", DataTypeEnum.adInteger, adParamInput, 1, .Stats.OroDuelos)
        
        ' USER_FACTION
        Cmd.Parameters.Append Cmd.CreateParameter("Alignment", DataTypeEnum.adTinyInt, adParamInput, 1, .Faccion.Alignment)
        Cmd.Parameters.Append Cmd.CreateParameter("IsArmy", DataTypeEnum.adTinyInt, adParamInput, 1, .Faccion.ArmadaReal)
        Cmd.Parameters.Append Cmd.CreateParameter("IsChaos", DataTypeEnum.adTinyInt, adParamInput, 1, .Faccion.FuerzasCaos)
        Cmd.Parameters.Append Cmd.CreateParameter("NeutralsKilled", DataTypeEnum.adInteger, adParamInput, 1, .Faccion.NeutralsKilled)
        Cmd.Parameters.Append Cmd.CreateParameter("CitizensKilled", DataTypeEnum.adInteger, adParamInput, 1, .Faccion.CiudadanosMatados)
        Cmd.Parameters.Append Cmd.CreateParameter("CriminalsKilled", DataTypeEnum.adInteger, adParamInput, 1, .Faccion.CriminalesMatados)
        Cmd.Parameters.Append Cmd.CreateParameter("ChaosArmorGiven", DataTypeEnum.adTinyInt, adParamInput, 1, .Faccion.RecibioArmaduraCaos)
        
        Cmd.Parameters.Append Cmd.CreateParameter("ArmyArmorGiven", DataTypeEnum.adTinyInt, adParamInput, 1, .Faccion.RecibioArmaduraReal)
        Cmd.Parameters.Append Cmd.CreateParameter("ChaosExpGiven", DataTypeEnum.adTinyInt, adParamInput, 1, .Faccion.RecibioExpInicialCaos)
        Cmd.Parameters.Append Cmd.CreateParameter("ArmyExpGiven", DataTypeEnum.adTinyInt, adParamInput, 1, .Faccion.RecibioExpInicialReal)
        Cmd.Parameters.Append Cmd.CreateParameter("ChaosRewardGiven", DataTypeEnum.adTinyInt, adParamInput, 1, .Faccion.RecompensasCaos)
        Cmd.Parameters.Append Cmd.CreateParameter("ArmyRewardGiven", DataTypeEnum.adTinyInt, adParamInput, 1, .Faccion.RecompensasReal)
        Cmd.Parameters.Append Cmd.CreateParameter("NumSigns", DataTypeEnum.adInteger, adParamInput, 1, .Faccion.Reenlistadas)
        Cmd.Parameters.Append Cmd.CreateParameter("SigningLevel", DataTypeEnum.adTinyInt, adParamInput, 1, .Faccion.NivelIngreso)
        
        Cmd.Parameters.Append Cmd.CreateParameter("SigningDate", DataTypeEnum.adDate, adParamInput, 1, IIf(.Faccion.FechaIngreso = 0, vbNull, Format$(.Faccion.FechaIngreso, "yyyy-mm-dd h:mm:ss")))
        Cmd.Parameters.Append Cmd.CreateParameter("SigningKilled", DataTypeEnum.adInteger, adParamInput, 1, .Faccion.MatadosIngreso)
        Cmd.Parameters.Append Cmd.CreateParameter("NextReward", DataTypeEnum.adInteger, adParamInput, 1, .Faccion.NextRecompensa)
        Cmd.Parameters.Append Cmd.CreateParameter("IsRoyalCouncil", DataTypeEnum.adTinyInt, adParamInput, 1, IIf(UserList(UserIndex).flags.Privilegios And PlayerType.RoyalCouncil, "1", "0"))
        Cmd.Parameters.Append Cmd.CreateParameter("IsChaosCouncil", DataTypeEnum.adTinyInt, adParamInput, 1, IIf(UserList(UserIndex).flags.Privilegios And PlayerType.ChaosCouncil, "1", "0"))

    End With

    Set Rs = ExecuteSqlCommand(Cmd)
    
    If NewChar Then
        UserList(UserIndex).ID = CLng(Rs.Fields("ID_USER"))
    End If
    
    
    SaveUserHeaderDB = UserList(UserIndex).ID

    
End Function


Private Sub SaveUserAttributesDB(ByVal UserIndex As Integer, ByVal NewChar As Boolean)
'***************************************************
'Author: ZaMa
'Creation Date: 09/06/2012
'Last Modification: -
'Saves user Attributes into DB
'***************************************************

On Error GoTo ErrHandler:

    With UserList(UserIndex)
                
        Dim Sql As String
        
        If NewChar Then
            Sql = _
                "INSERT INTO USER_ATTRIBUTES VALUES ('" & _
                    CStr(.ID) & "','" & _
                    CStr(.Stats.UserAtributos(eAtributos.Fuerza)) & "','" & _
                    CStr(.Stats.UserAtributos(eAtributos.Agilidad)) & "','" & _
                    CStr(.Stats.UserAtributos(eAtributos.Inteligencia)) & "','" & _
                    CStr(.Stats.UserAtributos(eAtributos.Carisma)) & "','" & _
                    CStr(.Stats.UserAtributos(eAtributos.Constitucion)) & "' " & _
                ")"
        Else
            
            If Not .flags.TomoPocion Then
                Sql = _
                    "UPDATE USER_ATTRIBUTES SET " & _
                        "STRENGHT ='" & CStr(.Stats.UserAtributos(eAtributos.Fuerza)) & "'," & _
                        "DEXERITY ='" & CStr(.Stats.UserAtributos(eAtributos.Agilidad)) & "'," & _
                        "INTELLIGENCE ='" & CStr(.Stats.UserAtributos(eAtributos.Inteligencia)) & "'," & _
                        "CHARISM ='" & CStr(.Stats.UserAtributos(eAtributos.Carisma)) & "'," & _
                        "HEALTH ='" & CStr(.Stats.UserAtributos(eAtributos.Constitucion)) & "' " & _
                    "WHERE ID_USER = '" & CStr(.ID) & "' "
            Else
                Sql = _
                    "UPDATE USER_ATTRIBUTES SET " & _
                        "STRENGHT ='" & CStr(.Stats.UserAtributosBackUP(eAtributos.Fuerza)) & "'," & _
                        "DEXERITY ='" & CStr(.Stats.UserAtributosBackUP(eAtributos.Agilidad)) & "'," & _
                        "INTELLIGENCE ='" & CStr(.Stats.UserAtributosBackUP(eAtributos.Inteligencia)) & "'," & _
                        "CHARISM ='" & CStr(.Stats.UserAtributosBackUP(eAtributos.Carisma)) & "'," & _
                        "HEALTH ='" & CStr(.Stats.UserAtributosBackUP(eAtributos.Constitucion)) & "' " & _
                    "WHERE ID_USER = '" & CStr(.ID) & "' "
            End If
        End If
        
        Call ExecuteSql(Sql)
    End With
    
    Exit Sub
ErrHandler:
    LogError ("Error en SaveUserAttributesDB: " & Err.Description)
    
End Sub

Private Sub SaveUserPrivateMsjDB(ByVal UserIndex As Integer, ByVal UserId As Long)
'***************************************************
'Author: ZaMa
'Creation Date: 14/09/2012
'Last Modification: -
'Saves user private messaje into DB
'***************************************************
On Error GoTo ErrHandler:

    With UserList(UserIndex)
        
        Dim Sql As String
        
        ' Delete previous
        Call DeleteUserPrivateMsjDB(UserId, 0)
        
        ' Store actual
        Dim Slot As Long
        Dim CurrentSlot As Byte
        For Slot = 1 To Constantes.MaxPrivateMessages
            
            With .Mensajes(Slot)
                .Contenido = Left(.Contenido, 100)
            
                If LenB(.Contenido) <> 0 Then
                    CurrentSlot = CurrentSlot + 1
                    
                    Sql = _
                        "INSERT INTO USER_MESSAGES VALUES ('" & _
                            CStr(UserId) & "','" & _
                            CStr(CurrentSlot) & "','" & _
                            EscapeString(Left$(CStr(.Contenido), 100)) & "','" & _
                            CStr(IIf(.Nuevo, 1, 0)) & "' " & _
                        ")"
                    
                    Call ExecuteSql(Sql)
                
                End If
            End With
        Next Slot
    End With
    Exit Sub
    
ErrHandler:
    LogError ("Error en SaveUserPrivateMsjDB: " & Err.Description)
End Sub

Public Sub DeleteUserPrivateMsjDB(ByVal UserId As Long, ByVal MpIndex As Byte)
'***************************************************
'Author: ZaMa
'Creation Date: 14/09/2012
'Last Modification: -
'Deletes user private messaje from DB and reorder index.
'***************************************************
On Error GoTo ErrHandler
  
        
    Dim Sql As String
    
    ' Delete previous
    Sql = _
        "DELETE FROM USER_MESSAGES " & _
        "WHERE ID_USER = '" & CStr(UserId) & "' "
    
    If MpIndex <> 0 Then
        Sql = Sql & _
            "AND MSG_INDEX = '" & CStr(MpIndex) & "' "
    End If
    
    Call ExecuteSql(Sql)
    
    ' Reorder (Not necesary if it's last one)
    If MpIndex <> 0 And MpIndex <> Constantes.MaxPrivateMessages Then
        Sql = _
            "UPDATE " & _
                "USER_MESSAGES " & _
            "SET " & _
                "MSG_INDEX = MSG_INDEX - 1 " & _
            "WHERE ID_USER = '" & CStr(UserId) & "' " & _
                "AND MSG_INDEX > '" & CStr(MpIndex) & "' "
                
        Call ExecuteSql(Sql)
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub DeleteUserPrivateMsjDB de modDB_Functions.bas")
End Sub

Public Sub SaveCriticEventDB(ByRef sEvent As String)
'***************************************************
'Author: ZaMa
'Creation Date: 29/08/2012
'Last Modification: -
'Saves critic event in DB.
'***************************************************
On Error GoTo ErrHandler

    sEvent = Left(sEvent, 100)

    Dim Sql As String
    Sql = _
        "INSERT INTO CRITIC_EVENTS VALUES ('" & _
            "0" & "','" & _
            Format$(Now, "yyyy-mm-dd h:mm:ss") & "','" & _
            sEvent & "' " & _
        ")"
    
    Call ExecuteSql(Sql)
    
    Exit Sub
    
ErrHandler:
    Call LogErrorDB(Err.Description, Sql)
End Sub



Public Function GetCharInfoByCharId(ByRef UserId As Long, _
    ByRef Banned As Boolean) As Boolean
'***************************************************
'Author: ZaMa
'Creation Date: 29/08/2012
'Last Modification: -
'Get main user info from DB. Returns false if no user name matches.
'***************************************************
On Error GoTo ErrHandler
  
    Dim Sql As String
    Sql = _
        "SELECT " & _
            "ID_USER, " & _
            "BANNED " & _
        "FROM " & _
            "USER_INFO " & _
        "WHERE ID_USER = '" & UserId & "' "
        
    Dim Rs As Recordset
    Set Rs = ExecuteSql(Sql)
    
    With Rs
        If Not .EOF Then
            Banned = (CByte(.Fields("BANNED")) = 1)

            GetCharInfoByCharId = True
        End If
    End With
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetCharInfoByCharId de modDB_Functions.bas")
End Function

Public Function GetCharInfo(ByRef UserName As String, ByRef UserId As Long, _
    ByRef Banned As Boolean) As Boolean
'***************************************************
'Author: ZaMa
'Creation Date: 29/08/2012
'Last Modification: -
'Get main user info from DB. Returns false if no user name matches.
'***************************************************
On Error GoTo ErrHandler
  
    Dim Sql As String
    Sql = _
        "SELECT " & _
            "ID_USER, " & _
            "BANNED " & _
        "FROM " & _
            "USER_INFO " & _
        "WHERE NAME = '" & UserName & "' "
        
    Dim Rs As Recordset
    Set Rs = ExecuteSql(Sql)
    
    With Rs
        If Not .EOF Then
            Banned = (CByte(.Fields("BANNED")) = 1)
            UserId = CLng(.Fields("ID_USER"))
            GetCharInfo = True
        End If
    End With
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetCharInfo de modDB_Functions.bas")
End Function


Public Function GetCharInfoWithGuild(ByRef UserName As String, ByRef UserId As Long, _
    ByRef GuildId As Long, ByRef Banned As Boolean) As Boolean
'***************************************************
'Author: ZaMa
'Creation Date: 29/08/2012
'Last Modification: -
'Get main user info from DB. Returns false if no user name matches.
'***************************************************
On Error GoTo ErrHandler
  
    Dim Sql As String
    Sql = _
        "SELECT " & _
            "ID_USER, " & _
            "GUILD_ID, " & _
            "BANNED " & _
        "FROM " & _
            "USER_INFO " & _
        "WHERE NAME = '" & UserName & "' "
        
    Dim Rs As Recordset
    Set Rs = ExecuteSql(Sql)
    
    With Rs
        If Not .EOF Then
            Banned = (CByte(.Fields("BANNED")) = 1)
            UserId = CLng(.Fields("ID_USER"))
            GuildId = CLng(.Fields("GUILD_ID"))
        End If
    End With
    
    GetCharInfoWithGuild = True
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetCharInfo de modDB_Functions.bas")
End Function

Public Function SetUserLoggedStateDB(ByVal UserId As Long, ByVal Connection As Boolean, ByVal IP As String, ByVal PreviousConnectionEventId As Long) As Long
'***************************************************
'Sets the user logged state by calling the sp_setCharacterLoggedIn stored procedure
'***************************************************
On Error GoTo ErrHandler

    Dim Rs As Recordset
    Dim Cmd As ADODB.Command
    Set Cmd = New ADODB.Command
    
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "sp_setCharacterLoggedIn"
    
    Cmd.Parameters.Append Cmd.CreateParameter("userID", DataTypeEnum.adInteger, ParameterDirectionEnum.adParamInput, 1, UserId)
    Cmd.Parameters.Append Cmd.CreateParameter("userIP", DataTypeEnum.adBSTR, ParameterDirectionEnum.adParamInput, 1, IP)
    Cmd.Parameters.Append Cmd.CreateParameter("isConnectionEvent", DataTypeEnum.adBoolean, ParameterDirectionEnum.adParamInput, 1, Connection)
    Cmd.Parameters.Append Cmd.CreateParameter("currentConnectionEventId", DataTypeEnum.adInteger, adParamInput, 1, PreviousConnectionEventId)
    
    Set Rs = ExecuteSqlCommand(Cmd)
   
    SetUserLoggedStateDB = CLng(Rs.Fields("ID_EVENT"))

  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SetUserLoggedStateDB de modDB_Functions.bas")
End Function

Public Function GetUserID(ByRef sUserName As String) As Long
'***************************************************
'Author: ZaMa
'Creation Date: 07/09/2012
'Last Modification: -
'Returns user DB ID
'***************************************************
On Error GoTo ErrHandler
  
    
    Dim Sql As String
    Sql = _
        "SELECT " & _
            "ID_USER " & _
        "FROM " & _
            "USER_INFO " & _
        "WHERE NAME ='" & Left$(sUserName, 30) & "' "
        
    Dim Rs As Recordset
    Set Rs = ExecuteSql(Sql)
    
    If Not Rs.EOF Then
        GetUserID = CLng(Rs.Fields("ID_USER"))
    End If
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetUserID de modDB_Functions.bas")
End Function
Public Function GetCharData(ByRef sTable As String, ByRef sField As String, _
    ByVal UserId As Long) As String
'***************************************************
'Author: ZaMa
'Creation Date: 08/09/2012
'Last Modification: -
'Returns a char field
'***************************************************
On Error GoTo ErrHandler
  
    
    Dim Sql As String
    Sql = _
        "SELECT " & _
            sField & " " & _
        "FROM " & _
            sTable & " " & _
        "WHERE ID_USER ='" & CStr(UserId) & "' "
        
    Dim Rs As Recordset
    Set Rs = ExecuteSql(Sql)
    
    If Not Rs.EOF Then
        GetCharData = CStr(Rs.Fields(sField))
    End If
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetCharData de modDB_Functions.bas")
End Function

Public Function GetUserNameDataByID(ByRef sTable As String, ByRef sField As String, _
    ByRef UserId As Long) As String
'***************************************************
'Author: Nightw
'Creation Date: 08/09/2012
'Last Modification: -
'Returns a char field by userID
'***************************************************
On Error GoTo ErrHandler
  
    
    Dim Sql As String
    Sql = _
        "SELECT " & _
            sField & " " & _
        "FROM " & _
            sTable & " " & _
        "WHERE ID_USER = " & CStr(UserId)
        
    Dim Rs As Recordset
    Set Rs = ExecuteSql(Sql)
    
    If Not Rs.EOF Then
        GetUserNameDataByID = CStr(Rs.Fields(sField))
    End If
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetUserNameDataByID de modDB_Functions.bas")
End Function

Public Function GetUserNameData(ByRef sTable As String, ByRef sField As String, _
    ByRef UserName As String) As String
'***************************************************
'Author: ZaMa
'Creation Date: 08/09/2012
'Last Modification: -
'Returns a char field with user name.
'***************************************************
On Error GoTo ErrHandler
  
    
    Dim Sql As String
    Sql = _
        "SELECT " & _
            sField & " " & _
        "FROM " & _
            sTable & " " & _
        "WHERE NAME = '" & UserName & "' "
        
    Dim Rs As Recordset
    Set Rs = ExecuteSql(Sql)
    
    If Not Rs.EOF Then
        GetUserNameData = CStr(Rs.Fields(sField))
    End If
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetUserNameData de modDB_Functions.bas")
End Function



Public Function GetCharDataWithName(ByRef sTable As String, ByRef sField As String, _
    ByRef UserName As String) As String
'***************************************************
'Author: ZaMa
'Creation Date: 08/09/2012
'Last Modification: -
'Returns a char field
'***************************************************
On Error GoTo ErrHandler
  
    
    Dim Sql As String
    Sql = _
        "SELECT " & _
            sTable & "." & sField & " " & _
        "FROM " & sTable & _
        " INNER JOIN USER_INFO UI " & _
            " ON UI.ID_USER = " & sTable & ".ID_USER" & _
        " WHERE UI.NAME = '" & UserName & "' "
        
    Dim Rs As Recordset
    Set Rs = ExecuteSql(Sql)
    
    If Not Rs.EOF Then
        GetCharDataWithName = CStr(Rs.Fields(sField))
    End If
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetCharDataWithName de modDB_Functions.bas")
End Function

Public Sub UpdateCharData(ByRef sTable As String, ByRef sField As String, _
    ByVal UserId As Long, ByRef sNewValue As String, Optional ByVal UseQuotes As Boolean = True)
'***************************************************
'Author: ZaMa
'Creation Date: 07/09/2012
'Last Modification: -
'Updates a char field
'***************************************************
On Error GoTo ErrHandler
  
    
    Dim sValue As String
    If UseQuotes Then
        sValue = "'" & sNewValue & "'"
    Else
        sValue = sNewValue
    End If
    
    Dim Sql As String
    Sql = _
        "UPDATE " & sTable & " SET " & _
            sField & " = " & sValue & " " & _
        "WHERE ID_USER ='" & CStr(UserId) & "' "
        
    Call ExecuteSql(Sql)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub UpdateCharData de modDB_Functions.bas")
End Sub

Public Sub UpdateCharInfo(ByRef sField As String, _
    ByRef sUserName As String, ByRef sNewValue As String)
'***************************************************
'Author: ZaMa
'Creation Date: 08/09/2012
'Last Modification: -
'Updates a char field of user_info table
'***************************************************
On Error GoTo ErrHandler
  
    
    Dim Sql As String
    Sql = _
        "UPDATE USER_INFO SET " & _
            sField & " = '" & sNewValue & "' " & _
        "WHERE NAME = '" & sUserName & "' "
        
    Call ExecuteSql(Sql)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub UpdateCharInfo de modDB_Functions.bas")
End Sub

Public Sub UpdateCharSkills(ByVal UserId As Long, ByVal skill As Byte, _
    ByVal NaturalAmount As Byte, ByVal AssignedAmount As Byte, ByVal EluSkills As Long, ByVal ExpSkills As Long)
'***************************************************
'Author: ZaMa
'Creation Date: 08/09/2012
'Last Modification: -
'Updates a char given skill
'***************************************************
On Error GoTo ErrHandler
  
    
    Dim Sql As String
    Sql = _
        "UPDATE USER_SKILLS SET " & _
            "NATURAL_AMOUNT = '" & CStr(NaturalAmount) & "', " & _
            "ASSIGNED_AMOUNT = '" & CStr(AssignedAmount) & "', " & _
            "SKILL_EXP_NEXT_LEVEL = '" & CStr(EluSkills) & "', " & _
            "SKILL_EXP = '" & CStr(ExpSkills) & "' " & _
        "WHERE ID_USER ='" & CStr(UserId) & "' " & _
            "AND SKILL ='" & CStr(skill) & "' "
        
    Call ExecuteSql(Sql)
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub UpdateCharSkills de modDB_Functions.bas")
End Sub

Public Sub GetCharSkillDB(ByVal UserId As Long, ByVal Skill As Byte, ByRef NaturalSkills As Byte, ByRef AssignedSkills As Byte)
On Error GoTo ErrHandler
  
    Dim Sql As String
    Dim Rs As Recordset
    
    Sql = _
        " SELECT  " & _
            "NATURAL_AMOUNT, " & _
            "ASSIGNED_AMOUNT " & _
        " FROM USER_SKILLS " & _
        " WHERE ID_USER ='" & CStr(UserId) & "' " & _
            "AND SKILL ='" & CStr(Skill) & "' "
    
    Set Rs = ExecuteSql(Sql)
    
    NaturalSkills = CByte(Rs.Fields("NATURAL_AMOUNT"))
    AssignedSkills = CByte(Rs.Fields("ASSIGNED_AMOUNT"))

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub GetCharSkillDB de modDB_Functions.bas")
End Sub

Public Function GetPunishmentEndDate(ByVal UserId As Long, ByVal punishmentType As Long) As tPunishmentType
'***************************************************
'Author: Nightw
'Creation Date: 10/09/2014
'Last Modification: 04/04/2015
'Get the amount of time that applies to a specific punishment
'04/04/2015: D'Artagnan - Allow unlimited jail time.
'***************************************************
On Error GoTo ErrHandler
  
    Dim EndDate As Date
    Dim pAmount As Integer
    Dim Sql As String
    Dim Rs As Recordset
    
    ' Get the amount of punishment of the same previously applied to the user
    Sql = "SELECT COUNT(1) as COUNT FROM USER_PUNISHMENT WHERE ID_PUNISHMENT_TYPE = " & CStr(punishmentType) & " AND ID_USER = " & CStr(UserId)
    
    Set Rs = ExecuteSql(Sql)
    If Not Rs.EOF Then
        pAmount = CInt(Rs.Fields("COUNT").value)
    End If
     
     
    ' Get the punishment rules that applies to the next round
    Sql = "SELECT * FROM PUNISHMENT_TYPE_RULES TR " & _
            " INNER JOIN PUNISHMENT_TYPE T ON TR.ID_PUNISHMENT_TYPE = T.ID " & _
            " Where ID_PUNISHMENT_TYPE = " & CStr(punishmentType) & _
            " AND PUNISHMENT_COUNT > " & CStr(pAmount) & _
            " AND ENABLED = 1" & _
            " ORDER BY PUNISHMENT_COUNT LIMIT 1"
    Set Rs = ExecuteSql(Sql)
    If Not Rs.EOF Then
        
        'yyyy Year
        'q Quarter
        'M Month
        'y   Day of year
        'd Day
        'W Weekday
        'ww Week
        'H Hour
        'N Minute
        'S Second
        
        Dim retPunishment As tPunishmentType
        
        With retPunishment
            .ID = CInt(Rs.Fields("ID").value)
            .BaseType = Rs.Fields("BASE_TYPE").value
            .Name = Rs.Fields("DESCRIPTION").value
            
            .AddBan = CBool(Rs.Fields("ADD_BAN").value)
            .AddJail = CBool(Rs.Fields("ADD_JAIL").value)
            .NextPunishment = CInt(Rs.Fields("NEXT_PUNISHMENT_ID").value)
            
            ReDim .Rules(0)
            .Rules(0).Count = CInt(Rs.Fields("PUNISHMENT_COUNT").value)
            .Rules(0).severity = CInt(Rs.Fields("PUNISHMENT_SEVERITY").value)
            
            Select Case CStr(Rs.Fields("BASE_TYPE").value)
                Case ePunishmentSubType.Jail ' Jail
                    If Rs.Fields("PUNISHMENT_SEVERITY").value >= 999 Then
                        .EndDate = DateValue("01/01/2050")
                    ' Jail might be unlimited.
                    ElseIf Rs.Fields("PUNISHMENT_SEVERITY").value > 0 Then
                        .EndDate = DateAdd("N", Rs.Fields("PUNISHMENT_SEVERITY").value, Now)
                    End If
                Case ePunishmentSubType.Ban ' Ban
                    If Rs.Fields("PUNISHMENT_SEVERITY").value >= 999 Then
                        .EndDate = DateValue("01/01/2050")
                    Else
                        .EndDate = DateAdd("d", Rs.Fields("PUNISHMENT_SEVERITY").value, Now)
                    End If
            End Select
            
        End With

        Debug.Print retPunishment.EndDate

    End If
    
    GetPunishmentEndDate = retPunishment 'Return

  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetPunishmentEndDate de modDB_Functions.bas")
End Function

Public Sub AddPunishmentDB2(ByVal UserId As Long, ByRef PunisherName As String, ByRef Description As String)
'***************************************************
'Author: ZaMa
'Creation Date: 11/08/2012
'Last Modification: -
'Saves user punishment into DB
'***************************************************
On Error GoTo ErrHandler
  
    Exit Sub
    Dim Sql As String
    Sql = _
        "INSERT INTO USER_PUNISHMENT VALUES ('" & _
            "0','" & _
            CStr(UserId) & "','" & _
            Format(Date, "yyyy-mm-dd") & "','" & _
            CStr(Time) & "','" & _
            PunisherName & "','" & _
            Left$(Description, 100) & "' " & _
        ")"
    
    Call ExecuteSql(Sql)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub AddPunishmentDB2 de modDB_Functions.bas")
End Sub

Public Function GetPunishmentTypes(ByVal BaseType As Byte)
On Error GoTo ErrHandler
  
    
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetPunishmentTypes de modDB_Functions.bas")
End Function

Public Sub AddPunishmentDB_OLDBORRAR(ByVal UserId As Long, ByVal PunisherId As Long, ByVal PunishmentType As Integer, ByRef EndDate As Date, ByRef Reason As String, ByRef AdminNotes As String)
On Error GoTo ErrHandler
  
    Dim Sql As String
        
    Sql = _
    "INSERT INTO USER_PUNISHMENT VALUES (0," & _
            CStr(UserId) & "," & _
            CStr(PunisherId) & "," & _
            CStr(punishmentType) & ",'" & _
            FormatDateDB(Now) & "','" & _
            FormatDateDB(EndDate) & "','" & _
            EscapeString(Left$(Reason, 255)) & "','" & _
            EscapeString(Left$(AdminNotes, 255)) & "'" & _
            ")"
            
    
    Call ExecuteSql(Sql)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub AddPunishmentDB de modDB_Functions.bas")
End Sub


Public Function AddPunishmentDB(ByVal UserId As Long, ByVal PunisherId As Long, ByVal PunishmentType As Integer, _
                                Optional ByRef Reason As String, Optional ByRef AdminNotes As String) As tPunishmentDbResponse
'***************************************************
'Author: Nightw
'Creation Date: 18/09/2014
'Last Modification: Added some missing values because schema was changed
'Saves user punishment into DB and returns the ID of the
' Returns: The punishment that is assigned
'***************************************************
On Error GoTo ErrHandler
  
                
    Dim Cmd As ADODB.Command
    Set Cmd = New ADODB.Command
        
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "sp_AddCharacterPunishment"
    
    Cmd.Parameters.Append Cmd.CreateParameter("UserID", DataTypeEnum.adInteger, ParameterDirectionEnum.adParamInput, 1, UserId)
    Cmd.Parameters.Append Cmd.CreateParameter("PunisherID", DataTypeEnum.adInteger, ParameterDirectionEnum.adParamInput, 1, PunisherId)
    Cmd.Parameters.Append Cmd.CreateParameter("PunishmentTypeID", DataTypeEnum.adInteger, ParameterDirectionEnum.adParamInput, 1, PunishmentType)
    Cmd.Parameters.Append Cmd.CreateParameter("Reason", DataTypeEnum.adBSTR, ParameterDirectionEnum.adParamInput, 1, Reason)
    Cmd.Parameters.Append Cmd.CreateParameter("AdminNotes", DataTypeEnum.adBSTR, ParameterDirectionEnum.adParamInput, 1, AdminNotes)
    
    Dim Rs As Recordset
        
    Set Rs = ExecuteSqlCommand(Cmd)
    
    
    Dim Response As tPunishmentDbResponse
    Response.PunishmentTypeId = CInt(Rs.Fields("PunishmentTypeID"))
    Response.PunishmentBaseType = CByte(Rs.Fields("PunishmentBaseType"))
    Response.PunishmentSeverity = CLng(Rs.Fields("PunishmentSeverity"))
    
    Response.ForcedPunishmentTypeId = IIf(Not IsNull(Rs.Fields("ForcedPunishmentTypeID")), Rs.Fields("ForcedPunishmentTypeID"), 0)
    Response.ForcedPunismentBaseType = IIf(Not IsNull(Rs.Fields("ForcedPunishmentBaseType")), Rs.Fields("ForcedPunishmentBaseType"), 0)
    Response.ForcedPunishmentSeverity = IIf(Not IsNull(Rs.Fields("ForcedPunishmentSeverity")), Rs.Fields("ForcedPunishmentSeverity"), 0)
    Response.LastInsertedPunishmentId = IIf(Not IsNull(Rs.Fields("LastInsertedPunishment")), Rs.Fields("LastInsertedPunishment"), 0)
    
    
    AddPunishmentDB = Response
       
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub AddPunishmentDB de modDB_Functions.bas")
End Function

Public Sub AdjustPunishmentEndDate(ByVal punishmentID As Long, ByRef EndDate As Date)
'***************************************************
'Author: Nightw
'Creation Date: 04/01/2015
'Adjust the end date for a given punishmentID
'***************************************************
On Error GoTo ErrHandler
  
                
    Dim Sql As String
        
    Sql = _
    "UPDATE USER_PUNISHMENT " & _
    "SET END_DATE = '" & Format(EndDate, "yyyy-mm-dd ttttt") & "' " & _
    "WHERE ID_PUNISHMENT = " & punishmentID
    
    Call ExecuteSql(Sql)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub AdjustPunishmentEndDate de modDB_Functions.bas")
End Sub



Public Function GetLastPunishmentApplied(ByVal UserId As Long)
On Error GoTo ErrHandler
  
    
    Dim Sql As String
    Dim rsQuery As ADODB.Recordset
        
    Sql = _
        "SELECT ID_PUNISHMENT " & _
        "FROM USER_PUNISHMENT " & _
        "Where ID_USER = " & CStr(UserId) & _
        " ORDER BY ID_PUNISHMENT DESC " & _
        "LIMIT 0, 1 "
        
        
    Set rsQuery = ExecuteSql(Sql)
    
    If Not rsQuery.EOF Then
        GetLastPunishmentApplied = CInt(rsQuery.Fields("ID_PUNISHMENT").value)
    End If

  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetLastPunishmentApplied de modDB_Functions.bas")
End Function


Public Sub ExpellFromCouncilDB(ByVal UserId As Long)
'***************************************************
'Author: ZaMa
'Creation Date: 12/09/2012
'Last Modification: -
'Expells user from royal and chaos council
'***************************************************
On Error GoTo ErrHandler
  
    
    Dim Sql As String
    Sql = _
        "UPDATE USER_FACTION SET " & _
            "ROYAL_COUNCIL = '0', " & _
            "CHAOS_COUNCIL = '0' " & _
        "WHERE ID_USER ='" & CStr(UserId) & "' "
        
    Call ExecuteSql(Sql)
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ExpellFromCouncilDB de modDB_Functions.bas")
End Sub

Public Sub SendUserPunishments(ByVal UserIndex As Integer, ByVal UserId As Long)
'***************************************************
'Author: ZaMa
'Creation Date: 12/09/2012
'Last Modification: 8/12/2014
'Sends user punishments list
' 8/12/2014 - Nightw: Some changes to the way we retreive the punishment list.
'***************************************************
On Error GoTo ErrHandler
  

    Dim Sql As String
    
    Sql = "SELECT  PT.DESCRIPTION, " & _
                "UP.REASON, " & _
                "CASE WHEN ISNULL(UI.`NAME`) THEN '@SYSTEM@' ELSE UI.`NAME` END AS PUNISHER_NAME, " & _
                "UP.EVENT_DATE, " & _
                "UP.END_DATE " & _
                "FROM USER_PUNISHMENT UP " & _
                "LEFT JOIN USER_INFO UI " & _
                "    ON UP.ID_PUNISHER = UI.ID_USER " & _
                "INNER JOIN PUNISHMENT_TYPE PT " & _
                "    ON PT.ID = UP.ID_PUNISHMENT_TYPE " & _
                "WHERE UP.ID_USER = " & CStr(UserId) & _
                " ORDER BY ID_PUNISHMENT"
        
    Dim Rs As Recordset
    Set Rs = ExecuteSql(Sql)
    
    Dim NumPunishment As Integer
    With Rs
        While Not .EOF
            
            NumPunishment = NumPunishment + 1
            Call WriteConsoleMsg(UserIndex, _
                NumPunishment & " - " & CStr(.Fields("PUNISHER_NAME")) & ": " & _
                CStr(.Fields("DESCRIPTION")) & " (" & _
                CStr(.Fields("REASON")) & ") - " & _
                CStr(.Fields("EVENT_DATE")) & " - " & _
                CStr(.Fields("END_DATE")), FontTypeNames.FONTTYPE_INFO)
        
            .MoveNext
        Wend
    End With
    
    If NumPunishment = 0 Then
        Call WriteConsoleMsg(UserIndex, "Sin prontuario..", FontTypeNames.FONTTYPE_INFO)
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SendUserPunishments de modDB_Functions.bas")
End Sub

Public Function GetUserPunishmentDB(ByVal UserId As Long, ByVal PunishmentIndex As Byte, ByRef PunishmentDescrip As String) As Long
'***************************************************
'Author: ZaMa
'Creation Date: 12/09/2012
'Last Modification: -
'Returns user punishment ID and descrip
'***************************************************
On Error GoTo ErrHandler
  
    
    ' Invalid Index
    If PunishmentIndex <= 0 Then Exit Function
    
    Dim Sql As String
    Sql = _
        "SELECT " & _
            "ID_PUNISHMENT, " & _
            "DESCRIP " & _
        "FROM " & _
            "USER_PUNISHMENT " & _
        "WHERE " & _
            "ID_USER ='" & CStr(UserId) & "' " & _
        "ORDER BY " & _
            "ID_PUNISHMENT " & _
        "LIMIT 1 " & _
        "OFFSET " & CStr(PunishmentIndex - 1)
        
    Dim Rs As Recordset
    Set Rs = ExecuteSql(Sql)
    
    If Not Rs.EOF Then
        PunishmentDescrip = CStr(Rs.Fields("DESCRIP"))
        GetUserPunishmentDB = CLng(Rs.Fields("ID_PUNISHMENT"))
    End If
    
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetUserPunishmentDB de modDB_Functions.bas")
End Function

Public Sub UpdateUserPunishmentDB(ByVal UserId As Long, ByVal punishmentID As Long, ByRef NewPunishmentDescrip As String)
'***************************************************
'Author: ZaMa
'Creation Date: 12/09/2012
'Last Modification: -
'Updates user punishment index descrip.
'***************************************************
On Error GoTo ErrHandler
  
    Exit Sub
    Dim Sql As String
    Sql = _
        "UPDATE " & _
            "USER_PUNISHMENT " & _
        "SET " & _
            "DESCRIP ='" & Left$(NewPunishmentDescrip, 100) & "' " & _
        "WHERE " & _
            "ID_USER ='" & CStr(UserId) & "' " & _
            "AND ID_PUNISHMENT ='" & CStr(punishmentID) & "' "

    Call ExecuteSql(Sql)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub UpdateUserPunishmentDB de modDB_Functions.bas")
End Sub

Public Sub UpdateUserNameDB(ByVal UserId As Long, ByRef newName As String)
'***************************************************
'Author: ZaMa
'Creation Date: 12/09/2012
'Last Modification: -
'Updates user's name.
'***************************************************
On Error GoTo ErrHandler
  

    Dim Sql As String
    Sql = _
        "UPDATE " & _
            "USER_INFO " & _
        "SET " & _
            "NAME ='" & Left$(newName, 30) & "' " & _
        "WHERE " & _
            "ID_USER ='" & CStr(UserId) & "' "

    Call ExecuteSql(Sql)

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub UpdateUserNameDB de modDB_Functions.bas")
End Sub

Public Sub ExpellUserFromFactionDB(ByVal UserId As Long, ByVal NewAlignment As Byte, ByVal NumSigns As Integer, ByVal ExpellerUserId As Long)
On Error GoTo ErrHandler
  

    Dim Sql As String
    Sql = _
        "UPDATE user_faction " & _
        "SET CHAOS = 0, " & _
        "ARMY = 0, " & _
        "NEXT_REWARD = 0, " & _
        "ROYAL_COUNCIL = 0, " & _
        "CHAOS_COUNCIL = 0, " & _
        "EXPELLER = " & ExpellerUserId & "," & _
        "NUM_SIGNS = " & NumSigns & "," & _
        "ALIGNMENT = " & NewAlignment & " " & _
        "WHERE ID_USER = " & UserId

    Call ExecuteSql(Sql)

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ExpellUserFromFactionDB de modDB_Functions.bas")
End Sub

Public Sub SendUserMessagesDB(ByVal UserIndex As Integer, ByVal UserId As Long, _
    ByRef UserName As String, ByVal MpIndex As Byte)
'***************************************************
'Author: ZaMa
'Creation Date: 13/09/2012
'Last Modification: -
'Sends given user private messages
'***************************************************
On Error GoTo ErrHandler
  
    
    ' Invalid Index
    If MpIndex < 0 Or MpIndex > Constantes.MaxPrivateMessages Then Exit Sub
    
    Dim Sql As String
    Sql = _
        "SELECT " & _
            "MESSAGE, " & _
            "MSG_INDEX, " & _
            "UNREAD " & _
        "FROM " & _
            "USER_MESSAGES " & _
        "WHERE " & _
            "ID_USER ='" & CStr(UserId) & "' "
    
    If MpIndex <> 0 Then
        Sql = Sql & _
            "AND INDEX = '" & CStr(MpIndex) & "' "
    Else
        Sql = Sql & _
            "ORDER BY " & _
                "INDEX "
    End If
        
    Dim Rs As Recordset
    Set Rs = ExecuteSql(Sql)
    
    Call WriteConsoleMsg(UserIndex, "Mensajes privados de " & UserName & ":", FontTypeNames.FONTTYPE_INFOBOLD)
    
    With Rs
        
        If .EOF Then
            If MpIndex = 0 Then
                Call WriteConsoleMsg(UserIndex, "El usuario " & UserName & " no tiene mensajes privados.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call EnviarMensaje(UserIndex, MpIndex, "VACIO", False)
            End If
        Else
            While Not Rs.EOF
                Call EnviarMensaje(UserIndex, CByte(Rs.Fields("INDEX")), _
                    CStr(Rs.Fields("MESSAGE")), CByte(Rs.Fields("UNREAD")) = 1)
                    
                .MoveNext
            Wend
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SendUserMessagesDB de modDB_Functions.bas")
End Sub

Public Sub AddUserPrivateMsjDB(ByVal UserId As Long, ByRef Message As String)
'***************************************************
'Author: ZaMa
'Creation Date: 13/09/2012
'Last Modification: -
'Adds private messages to given user, If exceeds limit, deletes first one and reorders.
'***************************************************
On Error GoTo ErrHandler
  
    
    ' Last index
    Dim Sql As String
    Sql = _
        "SELECT " & _
            "MAX(INDEX) AS LAST_INDEX " & _
        "FROM " & _
            "USER_MESSAGES " & _
        "WHERE ID_USER ='" & CStr(UserId) & "' "
            
            
    Dim Rs As Recordset
    Set Rs = ExecuteSql(Sql)
    
    Dim MpIndex As Byte
    If Not Rs.EOF Then
         MpIndex = CByte(Rs.Fields("LAST_INDEX"))
    End If
    
    ' Delete the first
    If MpIndex = Constantes.MaxPrivateMessages Then
        Call DeleteUserPrivateMsjDB(UserId, 1)
    Else
        MpIndex = MpIndex + 1
    End If
    
    Sql = _
        "INSERT INTO USER_MESSAGES VALUES ('" & _
            CStr(UserId) & "','" & _
            CStr(MpIndex) & "','" & _
            EscapeString(Left$(Message, 100)) & "','" & _
            "1' " & _
        ")"
        
    Call ExecuteSql(Sql)
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub AddUserPrivateMsjDB de modDB_Functions.bas")
End Sub

Public Sub SaveErrorDB(ByRef sError As String)
'***************************************************
'Author: ZaMa
'Creation Date: 29/10/2012
'Last Modification: -
'Saves error in DB.
'***************************************************
On Error GoTo ErrHandler

    sError = Left(sError, 10000)
    
    Dim Sql As String
    Sql = _
        "INSERT INTO ERRORES VALUES ('" & _
            "0" & "','" & _
            Format(Date, "yyyy-mm-dd") & "','" & _
            CStr(Time) & "', '" & _
            sError & "' " & _
        ")"
    
    Call ExecuteSql(Sql)
    
    Exit Sub
    
ErrHandler:
    Call LogErrorDB(Err.Description, Sql)
End Sub

Public Sub SaveStatictisDB(ByVal UserId As Long, ByRef sDescrip As String)
'***************************************************
'Author: ZaMa
'Creation Date: 11/10/2012
'Last Modification: -
'Saves statictics in DB.
'***************************************************
On Error GoTo ErrHandler

    Dim Sql As String
    Sql = _
        "INSERT INTO STATICTICS VALUES ('" & _
            "0" & "','" & _
            Format$(Now, "yyyy-mm-dd h:mm:ss") & "','" & _
            CStr(UserId) & "','" & _
            Left$(sDescrip, 100) & "' " & _
        ")"
    
    Call ExecuteSql(Sql)

    Exit Sub
    
ErrHandler:
    Call LogErrorDB(Err.Description, Sql)
End Sub

Public Sub UpdateCharGuildRequest(ByRef sUserName As String, ByRef sRejectDetail As String)
'***************************************************
'Author: ZaMa
'Creation Date: 15/10/2012
'Last Modification: -
'Updates a char guild request info.
'***************************************************
On Error GoTo ErrHandler
  
    
    Dim Sql As String
    Sql = _
        "UPDATE USER_INFO SET " & _
            "GUILD_REJECT_DETAIL = '" & sRejectDetail & "', " & _
            "REQUESTING_GUILD = '0' " & _
        "WHERE UPPER(NAME) ='" & UCase$(sUserName) & "' "
        
    Call ExecuteSql(Sql)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub UpdateCharGuildRequest de modDB_Functions.bas")
End Sub

Public Sub SaveUserBanDetailDB(ByVal UserId As Long, ByRef sBanedBy As String, ByRef sDetail As String)
'***************************************************
'Author: ZaMa
'Creation Date: 17/10/2012
'Last Modification: -
'Saves user ban detail in DB.
'***************************************************
On Error GoTo ErrHandler

    Dim Sql As String
    Sql = _
        "INSERT INTO USER_BAN_DETAIL VALUES ('" & _
            "0" & "','" & _
            Format$(Now, "yyyy-mm-dd h:mm:ss") & "','" & _
            CStr(UserId) & "','" & _
            Left$(sBanedBy, 30) & "','" & _
            Left$(sDetail, 50) & "' " & _
        ")"
    
    Call ExecuteSql(Sql)

    Exit Sub
    
ErrHandler:
    Call LogErrorDB(Err.Description, Sql)
End Sub

Public Sub SendUserIps(ByVal UserIndex As Integer, ByVal UserId As Long, ByRef UserName As String)
'***************************************************
'Author: ZaMa
'Creation Date: 14/01/2014
'Last Modification: -
'Sends user top 5 ips
'***************************************************
On Error GoTo ErrHandler
  

    Dim Sql As String
    Sql = _
        "SELECT " & _
            "CONNECTION_DATE, " & _
            "DISCONNECTION_DATE, " & _
            "IP " & _
        "FROM " & _
            "USER_CONNECTIONS " & _
        "WHERE " & _
            "ID_USER = '" & CStr(UserId) & "' " & _
            "AND IP <> '' " & _
        "ORDER BY " & _
            "ID_EVENT DESC " & _
        "LIMIT 5 "
        
    Dim Rs As Recordset
    Set Rs = ExecuteSql(Sql)
    
    Dim iCounter As Integer
    
    Dim sLista As String
    sLista = "Las ultimas IPs con las que " & UserName & " se conectó son:"
    With Rs
        While Not .EOF
            
            iCounter = iCounter + 1

            sLista = sLista & vbCrLf & _
                iCounter & " - " & CStr(Rs.Fields("IP")) & " - C: " & _
                    CStr(Rs.Fields("CONNECTION_DATE")) & " | D: " & _
                    modDB.GetStringOr(Rs.Fields("DISCONNECTION_DATE"), "")
            
            .MoveNext
        Wend
    End With
    
    If iCounter = 0 Then
        sLista = sLista & vbCrLf & _
            "Sin Lista.."
    End If
    
    Call WriteConsoleMsg(UserIndex, sLista, FontTypeNames.FONTTYPE_INFO)
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SendUserIps de modDB_Functions.bas")
End Sub

Public Sub SendUserStatsDB(ByVal UserIndex As Integer, ByRef UserName As String)
'***************************************************
'Author: Zama
'Last Modification: 15/10/2014
'
'***************************************************
On Error GoTo ErrHandler
  
    
    Dim UserId As Long
    UserId = GetUserID(UserName)
    
    If UserId = 0 Then
        Call WriteConsoleMsg(UserIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    Dim Sql As String
    Sql = _
        "SELECT " & _
            "USER_STATS.NIVEL," & _
            "USER_STATS.EXP," & _
            "USER_STATS.EXP_NEXT," & _
            "USER_STATS.HP_MAX," & _
            "USER_STATS.HP_MIN," & _
            "USER_STATS.STAMINA_MAX," & _
            "USER_STATS.STAMINA_MIN," & _
            "USER_STATS.MANA_MAX," & _
            "USER_STATS.MANA_MIN," & _
            "USER_STATS.HIT_MAX," & _
            "USER_STATS.HIT_MIN,"
                
   Sql = Sql & _
            "USER_STATS.ORO," & _
            "USER_ATTRIBUTES.STRENGHT, " & _
            "USER_ATTRIBUTES.DEXERITY, " & _
            "USER_ATTRIBUTES.INTELLIGENCE, " & _
            "USER_ATTRIBUTES.CHARISM, " & _
            "USER_ATTRIBUTES.HEALTH, " & _
            "USER_INFO.UP_TIME " & _
        "FROM " & _
            "USER_STATS, " & _
            "USER_INFO, " & _
            "USER_ATTRIBUTES " & _
        "WHERE " & _
            "USER_INFO.ID_USER = '" & CStr(UserId) & "' " & _
            "AND USER_STATS.ID_USER = USER_INFO.ID_USER " & _
            "AND USER_ATTRIBUTES.ID_USER = USER_INFO.ID_USER "
        
    Dim Rs As Recordset
    Set Rs = ExecuteSql(Sql)
    
    Call WriteConsoleMsg(UserIndex, "Estadísticas de: " & UserName, FontTypeNames.FONTTYPE_INFO)
    
    If Not Rs.EOF Then
        Call WriteConsoleMsg(UserIndex, "Nivel: " & CStr(Rs.Fields("NIVEL")) & _
            "  EXP: " & CStr(Rs.Fields("EXP")) & _
            "/" & CStr(Rs.Fields("EXP_NEXT")), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(UserIndex, "Energía: " & CStr(Rs.Fields("STAMINA_MIN")) & _
            "/" & CStr(Rs.Fields("STAMINA_MAX")), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(UserIndex, "Salud: " & CStr(Rs.Fields("HP_MIN")) & _
            "/" & CStr(Rs.Fields("HP_MAX")) & _
            "  Maná: " & CStr(Rs.Fields("MANA_MIN")) & _
            "/" & CStr(Rs.Fields("MANA_MAX")), FontTypeNames.FONTTYPE_INFO)
        
        Call WriteConsoleMsg(UserIndex, "Menor Golpe/Mayor Golpe: " & _
            CStr(Rs.Fields("HIT_MIN")) & _
            "/" & CStr(Rs.Fields("HIT_MAX")), FontTypeNames.FONTTYPE_INFO)
        
        Call WriteConsoleMsg(UserIndex, "Oro: " & CStr(Rs.Fields("ORO")), FontTypeNames.FONTTYPE_INFO)
        
#If ConUpTime Then
        Dim TempSecs As Long
        Dim TempStr As String
        TempSecs = Val(Rs.Fields("UP_TIME"))
        TempStr = (TempSecs \ 86400) & " Días, " & ((TempSecs Mod 86400) \ 3600) & " Horas, " & ((TempSecs Mod 86400) Mod 3600) \ 60 & " Minutos, " & (((TempSecs Mod 86400) Mod 3600) Mod 60) & " Segundos."
        Call WriteConsoleMsg(UserIndex, "Tiempo Logeado: " & TempStr, FontTypeNames.FONTTYPE_INFO)
#End If
        
        Call WriteConsoleMsg(UserIndex, "Dados: " & _
            CStr(Rs.Fields("STRENGHT")) & ", " & _
            CStr(Rs.Fields("DEXERITY")) & ", " & _
            CStr(Rs.Fields("INTELLIGENCE")) & ", " & _
            CStr(Rs.Fields("CHARISM")) & ", " & _
            CStr(Rs.Fields("HEALTH")), FontTypeNames.FONTTYPE_INFO)
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SendUserStatsDB de modDB_Functions.bas")
End Sub

Public Sub SendUserSkillsDB(ByVal UserIndex As Integer, ByRef UserName As String)
'***************************************************
'Author: Zama
'Last Modification: 15/10/2014
'
'***************************************************
On Error GoTo ErrHandler
  
    
    Dim UserId As Long
    UserId = GetUserID(UserName)
    
    If UserId = 0 Then
        Call WriteConsoleMsg(UserIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    Dim Sql As String
    Sql = _
        "SELECT " & _
            "SKILL, " & _
            "NATURAL_AMOUNT, " & _
            "ASSIGNED_AMOUNT " & _
        "FROM " & _
            "USER_SKILLS " & _
        "WHERE " & _
            "ID_USER = '" & CStr(UserId) & "' " & _
        "ORDER BY " & _
            "SKILL ASC "
        
    Dim Rs As Recordset
    Set Rs = ExecuteSql(Sql)
    
    Dim sMensaje As String
    While Not Rs.EOF
        
        sMensaje = sMensaje & _
            "CHAR>" & SkillsNames(CInt(Rs.Fields("SKILL"))) & " = " & _
                CStr(CByte(Rs.Fields("NATURAL_AMOUNT")) + CByte(Rs.Fields("ASSIGNED_AMOUNT"))) & vbCrLf
        Rs.MoveNext
    Wend
    
    ' Libres
    Sql = _
        "SELECT " & _
            "SKILLS " & _
        "FROM " & _
            "USER_STATS " & _
        "WHERE " & _
            "ID_USER = '" & CStr(UserId) & "' "
        
    Set Rs = ExecuteSql(Sql)
    If Not Rs.EOF Then
        sMensaje = sMensaje & _
            "CHAR> Libres: " & CStr(Rs.Fields("SKILLS"))
    End If
    
    Call WriteConsoleMsg(UserIndex, sMensaje, FontTypeNames.FONTTYPE_INFO)
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SendUserSkillsDB de modDB_Functions.bas")
End Sub

Public Sub SendUserMiniStatsDB(ByVal UserIndex As Integer, ByRef UserName As String)
'***************************************************
'Author: Zama
'Last Modification: 15/10/2014
'Shows the users Stats when the user is offline.
'***************************************************
On Error GoTo ErrHandler
  
    
    Dim UserId As Long
    UserId = GetUserID(UserName)
    
    If UserId = 0 Then
        Call WriteConsoleMsg(UserIndex, "El usuario no existe: " & UserName, FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    Dim Sql As String

    Sql = "SELECT " & _
                        "US.USERS_KILLED, " & _
                        "US.NPCS_KILLED, " & _
                        "UF.NEUTRAL_KILLED, " & _
                        "UF.CITY_KILLED, " & _
                        "UF.CRI_KILLED, " & _
                        "UF.ARMY, " & _
                        "UF.CHAOS, " & _
                        "UF.CHAOS_EXP_GIVEN, " & _
                        "UF.ARMY_EXP_GIVEN, " & _
                        "UF.NUM_SIGNS, " & _
                        "UF.SIGNING_LEVEL, " & _
                        "UF.SIGNING_DATE, " & _
                        "UF.SIGNING_KILLED,"

    Sql = Sql & _
                        "UI.GUILD_ID, " & _
                        "UI.PUNISHMENT, " & _
                        "UI.CLASS, " & _
                        "UI.RACE, " & _
                        "UI.BANNED, " & _
                        "UPUNISHER.`NAME` AS BANNED_BY, " & _
                        "PTYPE.DESCRIPTION AS PUNISHMENT_DESC, " & _
                        "UPUNISH.Reason AS PUNISHMENT_REASON " & _
            "FROM USER_INFO  UI " & _
            "INNER JOIN USER_STATS US " & _
                "ON UI.ID_USER = US.ID_USER " & _
            "INNER JOIN USER_FACTION UF " & _
                "ON UI.ID_USER = UF.ID_USER " & _
            "INNER JOIN USER_FLAGS UFLAGS " & _
                "ON UI.ID_USER = UFLAGS.ID_USER " & _
            "LEFT JOIN USER_PUNISHMENT UPUNISH " & _
                "ON UI.ID_BAN_PUNISHMENT = UPUNISH.ID_PUNISHMENT " & _
            "LEFT JOIN PUNISHMENT_TYPE PTYPE " & _
                "ON UPUNISH.ID_PUNISHMENT_TYPE = PTYPE.ID " & _
            "LEFT JOIN USER_INFO UPUNISHER " & _
                "ON UPUNISH.ID_PUNISHER = UPUNISHER.ID_USER " & _
            "WHERE UI.ID_USER = " & UserId

    Dim Rs As Recordset
    Set Rs = ExecuteSql(Sql)
        
    With Rs
        If Not .EOF Then
            
            Call WriteConsoleMsg(UserIndex, "Usuario: " & UserName, FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(UserIndex, "Neutrales Matados: " & Val(.Fields("NEUTRAL_KILLED")) & _
                " Armada Matados: " & Val(.Fields("CITY_KILLED")) & _
                " Caos Matados: " & Val(.Fields("CRI_KILLED")) & _
                " Total Matados: " & Val(.Fields("USERS_KILLED")), FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(UserIndex, "NPCs muertos: " & Val(.Fields("NPCS_KILLED")), FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(UserIndex, "Clase: " & ListaClases(Val(.Fields("CLASS"))), FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(UserIndex, "Raza: " & ListaRazas(Val(.Fields("RACE"))), FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(UserIndex, "Pena: " & Val(.Fields("PUNISHMENT")), FontTypeNames.FONTTYPE_INFO)
            
            Dim sFechaIngreso As String
            
            If Not IsNull(.Fields("SIGNING_DATE")) Then
                sFechaIngreso = Format$(CDate(.Fields("SIGNING_DATE")), "yyyy-mm-dd h:mm:ss")
            Else
                sFechaIngreso = "No ingresó a ninguna Facción"
            End If
            
            Dim sNivelIngreso As String
            sNivelIngreso = CStr(Val(.Fields("SIGNING_LEVEL")))
            
            Dim sReenlistadas As String
            sReenlistadas = CStr(Val(.Fields("NUM_SIGNS")))
            
            If Val(.Fields("ARMY")) = 1 Then
                Call WriteConsoleMsg(UserIndex, "Ejército real desde: " & sFechaIngreso, FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(UserIndex, "Ingresó en nivel: " & sNivelIngreso & " con " & _
                    CStr(Val(.Fields("SIGNING_KILLED"))) & " ciudadanos matados.", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(UserIndex, "Veces que ingresó: " & sReenlistadas, FontTypeNames.FONTTYPE_INFO)
            
            ElseIf Val(.Fields("CHAOS")) = 1 Then
                Call WriteConsoleMsg(UserIndex, "Legión oscura desde: " & sFechaIngreso, FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(UserIndex, "Ingresó en nivel: " & sNivelIngreso, FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(UserIndex, "Veces que ingresó: " & sReenlistadas, FontTypeNames.FONTTYPE_INFO)
            
            ElseIf Val(.Fields("ARMY_EXP_GIVEN")) = 1 Then
                Call WriteConsoleMsg(UserIndex, "Fue ejército real", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(UserIndex, "Veces que ingresó: " & sReenlistadas, FontTypeNames.FONTTYPE_INFO)
            
            ElseIf Val(.Fields("CHAOS_EXP_GIVEN")) = 1 Then
                Call WriteConsoleMsg(UserIndex, "Fue legión oscura", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(UserIndex, "Veces que ingresó: " & sReenlistadas, FontTypeNames.FONTTYPE_INFO)
            End If
        
            Dim Ban As Byte
            Ban = Val(.Fields("BANNED"))
            Call WriteConsoleMsg(UserIndex, "Ban: " & Ban, FontTypeNames.FONTTYPE_INFO)
            
            If Ban = 1 Then
                Dim bannedBy As String
                Dim Desc As String
                Dim Reason As String
                ' if Rs.fields.item("rsField").value & "" = "" then
                If Not (.Fields("BANNED_BY").value & "" = "") Then
                    bannedBy = CStr(.Fields("BANNED_BY"))
                End If
                If Not (.Fields("PUNISHMENT_DESC").value & "" = "") Then
                    Desc = CStr(.Fields("PUNISHMENT_DESC"))
                End If
                If Not (.Fields("PUNISHMENT_REASON").value & "" = "") Then
                    Reason = CStr(.Fields("PUNISHMENT_REASON"))
                End If
                
                Call WriteConsoleMsg(UserIndex, "Ban por: " & bannedBy & _
                    " - Motivo: " & Desc & " (" & Reason & ")", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SendUserMiniStatsDB de modDB_Functions.bas")
End Sub

Public Sub SendUserInvTxtFromDB(ByVal UserIndex As Integer, ByVal UserId As Long, ByRef UserName As String)
'***************************************************
'Author: Zama
'Last Modification: 17/10/2014
'
'***************************************************
On Error GoTo ErrHandler
  
    
    Dim Sql As String
    Sql = _
        "SELECT " & _
            "SLOT, " & _
            "OBJ_INDEX, " & _
            "AMOUNT, " & _
            "EQUIPPED " & _
        "FROM " & _
            "USER_INVENTORY " & _
        "WHERE ID_USER = '" & CStr(UserId) & "' " & _
            "ORDER BY SLOT ASC "
        
    Dim Rs As Recordset
    Set Rs = ExecuteSql(Sql)
    
    Call WriteConsoleMsg(UserIndex, "Items en inventario de " & UserName, FontTypeNames.FONTTYPE_INFO)
    
    Dim sMensaje As String
    Dim lNumItems As Long
    Dim ObjIndex As Integer
    With Rs
        While Not .EOF

            ObjIndex = CInt(.Fields("OBJ_INDEX"))
            If ObjIndex > 0 Then
                sMensaje = _
                    "Objeto " & CStr(.Fields("SLOT")) & _
                    "> " & ObjData(ObjIndex).Name & _
                    " Cantidad:" & Val(.Fields("AMOUNT"))
                
                If CInt(.Fields("EQUIPPED")) = 1 Then
                    sMensaje = sMensaje & " (E)"
                End If
                
                Call WriteConsoleMsg(UserIndex, sMensaje, FontTypeNames.FONTTYPE_INFO)
                lNumItems = lNumItems + 1
            End If
            
            .MoveNext
        Wend
    End With
    
    Call WriteConsoleMsg(UserIndex, "Tiene " & CStr(lNumItems) & " objetos.", FontTypeNames.FONTTYPE_INFO)
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SendUserInvTxtFromDB de modDB_Functions.bas")
End Sub

Public Sub SendUserBovedaTxtFromDB(ByVal UserIndex As Integer, ByVal UserId As Long, ByRef UserName As String)
'***************************************************
'Author: Zama
'Last Modification: 17/10/2014
'
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
        "WHERE ID_USER = '" & CStr(UserId) & "' " & _
            "ORDER BY SLOT ASC "
        
    Dim Rs As Recordset
    Set Rs = ExecuteSql(Sql)
    
    Call WriteConsoleMsg(UserIndex, "Items en bóveda de " & UserName, FontTypeNames.FONTTYPE_INFO)
    
    Dim sMensaje As String
    Dim lNumItems As Long
    Dim ObjIndex As Integer
    With Rs
        While Not .EOF

            ObjIndex = CInt(.Fields("OBJ_INDEX"))
            If ObjIndex > 0 Then
                sMensaje = _
                    "Objeto " & CStr(.Fields("SLOT")) & _
                    "> " & ObjData(ObjIndex).Name & _
                    " Cantidad:" & Val(.Fields("AMOUNT"))
                
                Call WriteConsoleMsg(UserIndex, sMensaje, FontTypeNames.FONTTYPE_INFO)
                lNumItems = lNumItems + 1
            End If
            
            .MoveNext
        Wend
    End With
    
    Call WriteConsoleMsg(UserIndex, "Tiene " & CStr(lNumItems) & " objetos.", FontTypeNames.FONTTYPE_INFO)
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SendUserBovedaTxtFromDB de modDB_Functions.bas")
End Sub


Public Sub ResetUserFactionDB(ByVal UserId As Long)
'***************************************************
'Author: Zama
'Last Modification: 18/10/2014
'
'***************************************************
On Error GoTo ErrHandler
  

    Dim Sql As String
    Sql = _
        "UPDATE USER_FACTION SET " & _
            "ARMY = '0', " & _
            "CHAOS = '0', " & _
            "CITY_KILLED = '0', " & _
            "CRI_KILLED = '0', " & _
            "CHAOS_ARMOUR_GIVEN = '0', " & _
            "ARMY_ARMOUR_GIVEN = '0', " & _
            "CHAOS_EXP_GIVEN = '0', " & _
            "ARMY_EXP_GIVEN = '0', " & _
            "CHAOS_REWARD_GIVEN = '0', " & _
            "ARMY_REWARD_GIVEN = '0', " & _
            "NUM_SIGNS = '0', " & _
            "SIGNING_LEVEL = '0', " & _
            "SIGNING_DATE = '', " & _
            "SIGNING_KILLED = '0', " & _
            "NEXT_REWARD = '0', " & _
            "ROYAL_COUNCIL = '0', " & _
            "CHAOS_COUNCIL = '0' " & _
        "WHERE ID_USER = '" & CStr(UserId) & "' "
        
    Call ExecuteSql(Sql)
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ResetUserFactionDB de modDB_Functions.bas")
End Sub

Public Function SaveAccountInfoDB(ByRef AccountId As Long, ByRef sName As String, ByRef sEmail As String, _
    ByRef sPassword As String, ByRef sSecretQuestion As String, ByRef sAnswer As String, _
    ByVal activationStatus As eAccountStatus) As Boolean
'***************************************************
'Author: ZaMa
'Creation Date: 24/01/2014
'Last Modification: -
'Saves account info in DB.
'***************************************************
On Error GoTo ErrHandler
    
    Dim Sql As String
    
    sName = Left(sName, 30)
    sEmail = Left(sEmail, 50)
    sPassword = Left(sPassword, 32)
    sSecretQuestion = Left(sSecretQuestion, 50)
    sAnswer = Left(sAnswer, 50)
    
    Sql = _
        "INSERT INTO ACCOUNT_INFO " & _
        "(NAME," & _
        "EMAIL," & _
        "PASSWORD," & _
        "SECRET_QUESTION," & _
        "ANSWER," & _
        "STATUS," & _
        "BAN_DETAIL," & _
        "CREATION_DATE," & _
        "BANK_GOLD," & _
        "BANK_PASSWORD) "

Sql = Sql & _
        "VALUES ('" & _
            sName & "','" & _
            sEmail & "','" & _
            sPassword & "','" & _
            sSecretQuestion & "','" & _
            EscapeString(sAnswer) & "','" & _
            CStr(activationStatus) & "','" & _
            vbNullString & "','" & _
            Format$(Now, "yyyy-mm-dd h:mm:ss") & "'," & _
            "0,'" & _
            vbNullString & "'" & _
        ")"
    
    Call ExecuteSql(Sql)
    
    ' Retrieve ID_ACCOUNT
    Sql = "SELECT LAST_INSERT_ID()"
        
    Dim Rs As Recordset
    Set Rs = ExecuteSql(Sql)
    
    AccountId = CLng(Rs.Fields(0))
    
    SaveAccountInfoDB = True
    
    Exit Function
    
ErrHandler:
    Call LogErrorDB(Err.Description, Sql)
End Function

Public Sub SaveAccountCharShareDB(ByVal OwnerAccountID As Long, ByVal SharedAccountID As Long, ByVal UserId As Long)
'***************************************************
'Author: ZaMa
'Creation Date: 24/01/2014
'Last Modification: -
'Saves account char sharing info in DB.
'***************************************************
On Error GoTo ErrHandler

    Dim Sql As String
    Sql = _
        "INSERT INTO ACCOUNT_CHAR_SHARE VALUES ('" & _
            CStr(OwnerAccountID) & "','" & _
            CStr(SharedAccountID) & "', " & _
            CStr(UserId) & "', " & _
            Format(Date, "yyyy-mm-dd") & "','" & _
            CStr(Time) & "' " & _
        ")"
    
    Call ExecuteSql(Sql)
    
    Exit Sub
    
ErrHandler:
    Call LogErrorDB(Err.Description, Sql)
End Sub

Public Function GetAccountID(ByRef sAccountName As String, Optional ByRef sPassword As String, _
    Optional ByRef sEmail As String, Optional ByRef sSecretQuestion As String, Optional ByRef sAnswer As String, _
    Optional ByRef iStatus As Integer, Optional ByRef sBanDetail As String) As Long
'***************************************************
'Author: ZaMa
'Creation Date: 24/01/2014
'Last Modification: -
'Returns account DB ID. Optionally, returns byref other info.
'***************************************************
On Error GoTo ErrHandler
  
    
    Dim Sql As String
    Sql = _
        "SELECT " & _
            "ID_ACCOUNT, " & _
            "PASSWORD, " & _
            "SECRET_QUESTION, " & _
            "ANSWER, " & _
            "STATUS, " & _
            "BAN_DETAIL, " & _
            "EMAIL " & _
        "FROM " & _
            "ACCOUNT_INFO " & _
        "WHERE NAME ='" & Left$(sAccountName, 30) & "' "
        
    Dim Rs As Recordset
    Set Rs = ExecuteSql(Sql)
    
    With Rs
        If Not .EOF Then
            sPassword = CStr(.Fields("PASSWORD"))
            sSecretQuestion = CStr(.Fields("SECRET_QUESTION"))
            sAnswer = CStr(.Fields("ANSWER"))
            iStatus = CInt(.Fields("STATUS"))
            sBanDetail = CStr(.Fields("BAN_DETAIL"))
            sEmail = CStr(.Fields("EMAIL"))
            
            GetAccountID = CLng(.Fields("ID_ACCOUNT"))
        End If
    End With
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetAccountID de modDB_Functions.bas")
End Function

Public Function GetAccountData(ByRef sField As String, _
    ByVal AccountId As Long) As String
'***************************************************
'Author: ZaMa
'Creation Date: 24/01/2014
'Last Modification: -
'Returns an account field
'***************************************************
On Error GoTo ErrHandler
  

    Dim Sql As String
    Sql = _
        "SELECT " & _
            sField & " " & _
        "FROM " & _
            "ACCOUNT_INFO " & _
        "WHERE ID_ACCOUNT ='" & CStr(AccountId) & "' "

    Dim Rs As Recordset
    Set Rs = ExecuteSql(Sql)

    If Not Rs.EOF Then
        GetAccountData = CStr(Rs.Fields(sField))
    End If
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetAccountData de modDB_Functions.bas")
End Function

Public Function AccountEmailExists(ByRef sEmail As String) As Boolean
'***************************************************
'Author: ZaMa
'Creation Date: 24/01/2014
'Last Modification: -
'Returns True if account email exists
'***************************************************
On Error GoTo ErrHandler
  
    
    Dim Sql As String
    Sql = _
        "SELECT " & _
            "ID_ACCOUNT " & _
        "FROM " & _
            "ACCOUNT_INFO " & _
        "WHERE EMAIL ='" & sEmail & "' "
        
    Dim Rs As Recordset
    Set Rs = ExecuteSql(Sql)
    
    If Rs.EOF Then Exit Function
       
    AccountEmailExists = True
    
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function AccountEmailExists de modDB_Functions.bas")
End Function

Public Sub UpdateAccountMail(ByVal nAccountID As Long, ByRef sMail As String)
'***************************************************
'Author: D'Artagnan
'Date: 05/01/2015
'Update the account email.
'***************************************************
On Error GoTo ErrHandler
  

    sMail = Left(sMail, 50)

    Call ExecuteSql("UPDATE ACCOUNT_INFO SET EMAIL = '" & sMail & _
                    "' WHERE ID_ACCOUNT = '" & CStr(nAccountID) & "'")
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub UpdateAccountMail de modDB_Functions.bas")
End Sub

Public Sub UpdateCharPassword(ByRef UserName As String, ByRef sNewPassword As String)
'***************************************************
'Author: ZaMa
'Creation Date: 25/01/2014
'Last Modification: -
'Updates account password.
'***************************************************
On Error GoTo ErrHandler
  

    sNewPassword = Left(sNewPassword, 32)
    
    Dim Sql As String
    Sql = _
        "UPDATE " & _
    "ACCOUNT_INFO " & _
    "SET PASSWORD = '" & sNewPassword & "' " & _
    "WHERE ID_ACCOUNT = " & GetAccountIDByUserID(GetUserID(UserName))

    Call ExecuteSql(Sql)
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub UpdateCharPassword de modDB_Functions.bas")
End Sub

Public Sub UpdateAccountPassword(ByVal AccountId As Long, ByRef sNewPassword As String, ByVal fromMd5Password As Boolean)
'***************************************************
'Author: ZaMa
'Creation Date: 25/01/2014
'Last Modification: -
'Updates account password.
'***************************************************
On Error GoTo ErrHandler
  
    
    Dim Sql As String

    ' If the password we are going to save is not in MD5 format
    ' we need to hash the string. Let's do it in the database directly.
    If fromMd5Password = False Then
        Sql = _
            "UPDATE " & _
                "ACCOUNT_INFO " & _
            "SET PASSWORD = UPPER(MD5('" & sNewPassword & "')) " & _
            "WHERE ID_ACCOUNT ='" & CStr(AccountId) & "' "
    Else
        sNewPassword = Left(sNewPassword, 32)
        
        Sql = _
        "UPDATE " & _
            "ACCOUNT_INFO " & _
        "SET PASSWORD = '" & sNewPassword & "' " & _
        "WHERE ID_ACCOUNT ='" & CStr(AccountId) & "' "
    End If
    
        
    Call ExecuteSql(Sql)
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub UpdateAccountPassword de modDB_Functions.bas")
End Sub

Public Sub GetBanDetails(ByVal nUserID As Long, ByRef sAdminName As String, ByRef sDesc As String, ByRef dEndDate As Date)
'***************************************************
'Author: D'Artagnan
'Last Modification: 07/12/2014
'26/03/2012: Nightw -  Changed the way we retrieve the punishment information
'Retrieve punisher name and description.
'07/12/2014: Changed the rsQuery.RecordCount to use the rsQuery.EOF and removed some garbage code.
'***************************************************
On Error GoTo ErrHandler
  
    Dim rsQuery As Recordset

    Dim str As String
    str = "SELECT IFNULL(UIP.NAME, '') as NAME, IFNULL(PT.Description, '') as DESCRIPTION, IFNULL(UP.END_DATE, '2025-12-30 00:00:00') AS END_DATE " & _
                                "FROM USER_INFO UI " & _
                                "LEFT JOIN USER_PUNISHMENT UP " & _
                                "   ON UI.ID_BAN_PUNISHMENT = UP.ID_PUNISHMENT " & _
                                "LEFT JOIN PUNISHMENT_TYPE PT " & _
                                "   ON UP.ID_PUNISHMENT_TYPE = PT.ID " & _
                                "LEFT JOIN USER_INFO UIP " & _
                                "   ON UIP.ID_USER = UP.ID_PUNISHER " & _
                                "WHERE UI.ID_USER = " & nUserID
                                
    Set rsQuery = ExecuteSql(str)
    
   If Not rsQuery.EOF > 0 Then
                
        sAdminName = CStr(rsQuery.Fields("NAME"))
        sDesc = CStr(rsQuery.Fields("DESCRIPTION"))
        dEndDate = CDate(rsQuery.Fields("END_DATE"))
        
        If sDesc = vbNullString Then
            sDesc = "desconocido."
        End If
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub GetBanDetails de modDB_Functions.bas")
End Sub

Public Sub GetPunishmentTypeCount()
'***************************************************
'Author: Nightw
'Last Modification: 01/09/2014
'Retrieve punishment types
'***************************************************
On Error GoTo ErrHandler
  
    Dim JailCount As Integer
    Dim BanCount As Integer
    Dim WarningCount As Integer
    
    Dim I As Integer, J As Integer
    I = 0
    J = 0
    
    Dim rsQuery As Recordset

    Set rsQuery = ExecuteSql("SELECT (SELECT COUNT(1) " & _
                            "FROM PUNISHMENT_TYPE " & _
                            "WHERE BASE_TYPE = 1) as JAIL,(SELECT COUNT(1) " & _
                            "FROM PUNISHMENT_TYPE " & _
                            "WHERE BASE_TYPE = 2) as BAN, (SELECT COUNT(1) " & _
                            "FROM PUNISHMENT_TYPE WHERE BASE_TYPE = 3) AS WARNING")
    
    
    If Not rsQuery.EOF Then
        JailCount = rsQuery.Fields("JAIL")
        BanCount = rsQuery.Fields("BAN")
        WarningCount = rsQuery.Fields("WARNING")
    End If
    
    
    
    ReDim listBanTypes(BanCount)
    ReDim listJailTypes(JailCount)
    ReDim listWarningTypes(WarningCount)
    
    
    Call GetPunishmentListByType(ePunishmentSubType.Ban, listBanTypes)
    Call GetPunishmentListByType(ePunishmentSubType.Jail, listJailTypes)
    Call GetPunishmentListByType(ePunishmentSubType.Warning, listWarningTypes)
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub GetPunishmentTypeCount de modDB_Functions.bas")
End Sub

'***************************************************
'Author: Nightw
'Last Modification: 01/09/2014
'Get a a list of punishments by type
'***************************************************
Public Sub GetPunishmentListByType(ByVal punishmentType As Integer, ByRef PunishmentList() As tPunishmentType)
    Dim I As Integer, J As Integer
    I = 0
    J = 0
    Dim rsQuery As Recordset
    Set rsQuery = ExecuteSql("SELECT * FROM PUNISHMENT_TYPE WHERE ENABLED = 1 AND BASE_TYPE = " & CStr(punishmentType))
                            
    While Not rsQuery.EOF
        With PunishmentList(I)
            .ID = rsQuery.Fields("ID")
            
            .Name = rsQuery.Fields("DESCRIPTION")
            .BaseType = ePunishmentSubType.Ban
            ReDim .Rules(0)
        End With
        I = I + 1
        rsQuery.MoveNext
    Wend
    
    I = 0
    J = 0

    Set rsQuery = ExecuteSql("SELECT * FROM PUNISHMENT_TYPE_RULES WHERE PUNISHMENT_TYPE_RULES.ID_PUNISHMENT_TYPE IN (SELECT ID FROM PUNISHMENT_TYPE WHERE BASE_TYPE = " & CStr(punishmentType) & ")")
                    
    While Not rsQuery.EOF
        For I = 0 To UBound(PunishmentList) - 1

            If PunishmentList(I).ID = rsQuery.Fields("ID_PUNISHMENT_TYPE") Then
                On Error Resume Next
                J = UBound(PunishmentList(I).Rules) + 1
                On Error GoTo Err:

                ReDim Preserve PunishmentList(I).Rules(J)
                PunishmentList(I).Rules(J).Count = rsQuery.Fields("PUNISHMENT_COUNT")
                PunishmentList(I).Rules(J).severity = rsQuery.Fields("PUNISHMENT_SEVERITY")

            End If
        Next I
        rsQuery.MoveNext
    Wend

    Exit Sub
Err:
End Sub

Public Function FormatDateDB(ByVal Value As Date) As String
'***************************************************
'Author: D'Artagnan
'Last Modification: 05/02/2015
'Just for dates consistency.
'***************************************************
On Error GoTo ErrHandler
  
    FormatDateDB = Format(value, "yyyy-mm-dd hh:mm:ss")
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function FormatDateDB de modDB_Functions.bas")
End Function


Public Sub AddMasteryDB(ByVal UserId As Long, ByVal MasteryGroup As Integer, ByVal MasteryID As Integer, ByVal PointsSpent As Integer)
On Error GoTo ErrHandler
  
    Dim Cmd As ADODB.Command
    Set Cmd = New ADODB.Command
        
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "sp_AddUserMastery"
    
    Dim BodyToSave As Integer
    
 
    Cmd.Parameters.Append Cmd.CreateParameter("userID", DataTypeEnum.adInteger, ParameterDirectionEnum.adParamInput, 1, UserId)
    Cmd.Parameters.Append Cmd.CreateParameter("masteryID", DataTypeEnum.adInteger, ParameterDirectionEnum.adParamInput, 1, MasteryID)
    Cmd.Parameters.Append Cmd.CreateParameter("groupID", DataTypeEnum.adInteger, ParameterDirectionEnum.adParamInput, 1, MasteryGroup)
    Cmd.Parameters.Append Cmd.CreateParameter("pointsSpent", DataTypeEnum.adInteger, ParameterDirectionEnum.adParamInput, 1, PointsSpent)
    
    Call ExecuteSqlCommand(Cmd)
    
    Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub AddMasteryDB de modDB_Functions.bas")
End Sub

Public Sub UnbanCharacterDB(ByVal UserId As Long, ByVal PunisherId As Long, ByVal PunishmentTypeId As Integer, ByRef UnbanReason As String)
On Error GoTo ErrHandler
  
    Dim Cmd As ADODB.Command
    Set Cmd = New ADODB.Command
        
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "sp_UnbanChar"
    
     
    Cmd.Parameters.Append Cmd.CreateParameter("userID", DataTypeEnum.adInteger, ParameterDirectionEnum.adParamInput, 1, UserId)
    Cmd.Parameters.Append Cmd.CreateParameter("punisherID", DataTypeEnum.adInteger, ParameterDirectionEnum.adParamInput, 1, PunisherId)
    Cmd.Parameters.Append Cmd.CreateParameter("punishmentTypeID", DataTypeEnum.adInteger, ParameterDirectionEnum.adParamInput, 1, PunishmentTypeId)
    Cmd.Parameters.Append Cmd.CreateParameter("unbanDate", DataTypeEnum.adDate, ParameterDirectionEnum.adParamInput, 1, FormatDateDB(Now()))
    Cmd.Parameters.Append Cmd.CreateParameter("reason", DataTypeEnum.adBSTR, ParameterDirectionEnum.adParamInput, 1, UnbanReason)
    
    Call ExecuteSqlCommand(Cmd)
    
    Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub UnbanCharacter de modDB_Functions.bas")
End Sub


Public Sub GetCharInformationForFactionKick(ByVal UserID As Long, ByRef CharacterName As String, ByRef CurrentAlignment As Byte, ByRef GuildId As Long)
On Error GoTo ErrHandler
  
    Dim RsQuery As Recordset
    Dim StrQuery As String
    
    StrQuery = "SELECT UI.ID_USER, UI.NAME, UF.ALIGNMENT, UI.GUILD_ID FROM user_info UI " & _
                "INNER JOIN USER_FACTION UF ON UI.ID_USER = UF.ID_USER " & _
                "LEFT JOIN GUILD_INFO GI ON UI.GUILD_ID = GI.ID_GUILD " & _
                "WHERE UI.ID_USER = " & UserID

    Set RsQuery = ExecuteSql(StrQuery)
    
                            
    If Not RsQuery.EOF <= 0 Then
        Exit Sub
    End If
                
    CurrentAlignment = CByte(RsQuery.Fields("ALIGNMENT"))
    GuildId = CLng(RsQuery.Fields("GUILD_ID"))
    CharacterName = CStr(RsQuery.Fields("NAME"))
    
    
    Set RsQuery = Nothing
    
    Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub GetCharInformationForFactionKick de modDB_Functions.bas")
End Sub
