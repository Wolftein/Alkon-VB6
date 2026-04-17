Attribute VB_Name = "NPCs"
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


'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                        Modulo NPC
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Contiene todas las rutinas necesarias para cotrolar los
'NPCs meno la rutina de AI que se encuentra en el modulo
'AI_NPCs para su mejor comprension.
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

Option Explicit

Sub MatarMascota(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
' Kills the user's pet
'***************************************************
On Error GoTo ErrHandler
  

    Dim I As Integer
    
    With UserList(UserIndex)
        For I = 1 To Classes(.clase).ClassMods.MaxTammedPets
            If .TammedPets(I).NpcIndex = NpcIndex Then
                .TammedPets(I).NpcIndex = 0
                .TammedPets(I).RemainingLife = 0
                Exit For
            End If
        Next I
    End With

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub QuitarMascota de MODULO_NPCs.bas")
End Sub

Public Function AlivePetCount(ByVal UserIndex As Integer) As Byte
    Dim Count As Byte
    Dim I As Integer
    Count = 0
    
    With UserList(UserIndex)
        For I = 1 To Classes(.clase).ClassMods.MaxTammedPets
            If .TammedPets(I).RemainingLife > 0 Then
                Count = Count + 1
            End If
        Next I
    End With
    
    AlivePetCount = Count
    
End Function


Sub QuitarMascotaNpc(ByVal Maestro As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    Npclist(Maestro).Mascotas = Npclist(Maestro).Mascotas - 1
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub QuitarMascotaNpc de MODULO_NPCs.bas")
End Sub

Public Sub MuereNpc(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
On Error GoTo ErrHandler

    Dim MiNPC As npc
    MiNPC = Npclist(NpcIndex)
    Dim nExperience As Long
    Dim I As Byte
   
   ' Es pretoriano?
    If MiNPC.NPCtype = eNPCType.Pretoriano Then
        Call ClanPretoriano(MiNPC.ClanIndex).MuerePretoriano(NpcIndex)
    End If
      
    'Invocacion e invocador
    With Npclist(NpcIndex)
    
        If .flags.isAffectedByDOT Then
            Call SendDamageOverTimeRemoveMessage(eTypeTarget.IsNPC, NpcIndex, Npclist(NpcIndex).InstanceId)
        End If
        
        If .flags.Invocador > 0 Then
            For I = 1 To Npclist(.flags.Invocador).flags.MaxInvocaciones
                If Npclist(.flags.Invocador).flags.Invocacion(I) = NpcIndex Then
                    Npclist(.flags.Invocador).flags.Invocacion(I) = 0
                    Exit For
                End If
            Next I
        End If
        
        If .flags.MaxInvocaciones > 0 Then
            For I = 1 To .flags.MaxInvocaciones
                If .flags.Invocacion(I) > 0 Then
                    Call QuitarNPC(.flags.Invocacion(I))
                End If
            Next I
        End If
        
        ' If there's any remaining time, that's because the NPC was killed by a player or NPC
        ' so we must reset the invoked pet element and decrease the counter
        ' Nightw
        If .Contadores.TiempoExistencia > 0 Then
            Dim PetIndex As Integer
            
            PetIndex = GetInvokedPetIndexByNpcIndex(.MaestroUser, NpcIndex)
            
            If PetIndex <> 0 Then
                UserList(.MaestroUser).InvokedPets(PetIndex).IsInvoked = False
                UserList(.MaestroUser).InvokedPets(PetIndex).NpcIndex = 0
                UserList(.MaestroUser).InvokedPets(PetIndex).NpcNumber = 0
                UserList(.MaestroUser).InvokedPets(PetIndex).RemainingLife = 0
                UserList(.MaestroUser).InvokedPetsCount = UserList(.MaestroUser).InvokedPetsCount - 1
            End If
        End If
        
    End With
    
    'Boss spawn y respawn
    If Npclist(NpcIndex).flags.Boss > 0 Then
        Call RestartBossSpawn(Npclist(NpcIndex).flags.Boss)
    End If
    
    Call CheckBossSpawn(NpcIndex)

    If UserIndex > 0 Then ' Lo mato un usuario?
        With UserList(UserIndex)
        
            If MiNPC.flags.Snd3 > 0 Then
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(MiNPC.flags.Snd3, MiNPC.Pos.X, MiNPC.Pos.Y))
            End If
            .flags.TargetNPC = 0
            .flags.TargetNpcTipo = eNPCType.Comun
            
            Dim T As Integer
            
            'El user que lo mato tiene mascotas?
            If .TammedPetsCount > 0 Then
                For T = 1 To Classes(.clase).ClassMods.MaxTammedPets
                      ' La mascota domada está invocada? y viva?
                      If .TammedPets(T).NpcIndex > 0 And .TammedPets(T).RemainingLife Then
                          If Npclist(.TammedPets(T).NpcIndex).TargetNPC = NpcIndex Then
                                  Call FollowAmo(.TammedPets(T).NpcIndex)
                          End If
                      End If
                Next T
            End If
            
            'El user que lo mato tiene mascotas?
            If .InvokedPetsCount > 0 Then
                For T = 1 To Classes(.clase).ClassMods.MaxInvokedPets
                      ' La mascota invocada está invocada? y viva?
                      If .InvokedPets(T).NpcIndex > 0 And .InvokedPets(T).RemainingLife Then
                          If Npclist(.InvokedPets(T).NpcIndex).TargetNPC = NpcIndex Then
                                  Call FollowAmo(.InvokedPets(T).NpcIndex)
                          End If
                      End If
                Next T
            End If
            
            '[KEVIN]
            If MiNPC.MaestroUser = 0 Then
                If MiNPC.flags.ExpCount > 0 Then
                    If .PartyIndex > 0 Then
                        nExperience = ApplyExperienceModifier(UserIndex, NpcIndex, MiNPC.flags.ExpCount)
                        Call mdParty.ObtenerExito(UserIndex, nExperience, MiNPC.Pos.Map, MiNPC.Pos.X, MiNPC.Pos.Y)
                    Else
                        nExperience = ApplyExperienceModifier(UserIndex, NpcIndex, MiNPC.flags.ExpCount)
                        .Stats.Exp = .Stats.Exp + nExperience
                        Call WriteConsoleMsg(UserIndex, "Has ganado " & nExperience & " puntos de experiencia.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                    End If
                    MiNPC.flags.ExpCount = 0
                End If
            End If
            
            '[/KEVIN]
            Call WriteConsoleMsg(UserIndex, "¡Has matado a la criatura!", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
            If .Stats.NPCsMuertos < 9999999 Then _
                .Stats.NPCsMuertos = .Stats.NPCsMuertos + 1
                                  
            Call CheckUserLevel(UserIndex)
            
            If NpcIndex = .flags.ParalizedByNpcIndex Then
                Call RemoveParalisis(UserIndex)
            End If
            
             If GuildHasQuest(.Guild.GuildIndex) Then
                Call modQuestSystem.GuildQuestUpdateStatus(.Guild.GuildIndex, UserIndex, 1, eQuestRequirement.NpcKill, MiNPC.Numero, 1)
            End If
            
        End With
        
        Call SubirSkill(UserIndex, eSkill.Supervivencia, True, ConstantesBalance.SkillExpNpcKilled)
    End If
    
    'Quitamos el npc
    Call QuitarNPC(NpcIndex)
    
    If MiNPC.MaestroUser = 0 Then

        Call NPCDropArrows(MiNPC, UserIndex)
        
        'ReSpawn o no
        Call ReSpawnNpc(MiNPC)
        
        'Tiramos el inventario
        If UserIndex > 0 Then
            Call NPC_TIRAR_ITEMS(MiNPC, MiNPC.NPCtype = eNPCType.Pretoriano, UserIndex, NpcIndex)
        End If
    End If
    
Exit Sub

ErrHandler:
    Call LogError("Error en MuereNpc - Error: " & Err.Number & " - Desc: " & Err.Description)
End Sub

Private Sub ResetNpcFlags(ByVal NpcIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    'Clear the npc's flags
    Dim I As Byte
    
    With Npclist(NpcIndex).flags
        .Boss = 0
        If .MaxInvocaciones > 0 Then
            For I = 1 To .MaxInvocaciones
                .Invocacion(I) = 0
            Next I
        End If
        .Invocador = 0
        .MaxInvocaciones = 0
        .KeepHeading = 0
        .AfectaParalisis = 0
        .AguaValida = 0
        .AttackedBy = vbNullString
        .AttackedFirstBy = vbNullString
        .BackUp = 0
        .Bendicion = 0
        .Domable = 0
        .Envenenado = 0
        .Faccion = 0
        .Follow = False
        .AtacaDoble = 0
        .LanzaSpells = 0
        .invisible = 0
        .Maldicion = 0
        .OldHostil = 0
        .OldMovement = 0
        .Paralizado = 0
        .Inmovilizado = 0
        .Respawn = 0
        .RespawnOrigPos = 0
        .DistanciaMaxima = 0
        .VolviendoOrig = 0
        .VolviendoInt = 0
        .Snd1 = 0
        .Snd2 = 0
        .Snd3 = 0
        .TierraInvalida = 0
        .isAffectedByDOT = False
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ResetNpcFlags de MODULO_NPCs.bas")
End Sub

Private Sub ResetNpcIntervalos(ByVal NpcIndex As Integer)
On Error GoTo ErrHandler
  
    With Npclist(NpcIndex).Intervalos
        .Walk = 0
        .Hit = 0
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ResetNpcIntervalos de MODULO_NPCs.bas")
End Sub

Private Sub ResetNpcCounters(ByVal NpcIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    With Npclist(NpcIndex).Contadores
        .Paralisis = 0
        .TiempoExistencia = 0
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ResetNpcCounters de MODULO_NPCs.bas")
End Sub

Private Sub ResetNpcCharInfo(ByVal NpcIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    With Npclist(NpcIndex).Char
        .body = 0
        .CascoAnim = 0
        .CharIndex = 0
        .FX = 0
        .head = 0
        .heading = 0
        .Loops = 0
        .ShieldAnim = 0
        .WeaponAnim = 0
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ResetNpcCharInfo de MODULO_NPCs.bas")
End Sub

Private Sub ResetNpcCriatures(ByVal NpcIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    Dim J As Long
    
    With Npclist(NpcIndex)
        For J = 1 To .NroCriaturas
            .Criaturas(J).NpcIndex = 0
            .Criaturas(J).NpcName = vbNullString
        Next J
        
        .NroCriaturas = 0
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ResetNpcCriatures de MODULO_NPCs.bas")
End Sub

Sub ResetExpresiones(ByVal NpcIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    Dim J As Long
    
    With Npclist(NpcIndex)
        For J = 1 To .NroExpresiones
            .Expresiones(J) = vbNullString
        Next J
        
        .NroExpresiones = 0
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ResetExpresiones de MODULO_NPCs.bas")
End Sub

Private Sub ResetNpcMainInfo(ByVal NpcIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'22/05/2010: ZaMa - Ahora se resetea el dueño del npc también.
'***************************************************
On Error GoTo ErrHandler
  

    With Npclist(NpcIndex)
        .Attackable = 0
        .CanAttack = 0
        .Comercia = 0
        .GiveEXP = 0
        .GiveGLD = 0
        .Hostile = 0
        .InvReSpawn = 0
        
        Dim I As Long
       
        For I = 1 To 6
            .TengoFlechas(I) = 0
        Next I
        
        If .ExtraBodies > 0 Then
            For I = 1 To .ExtraBodies
                .ExtraBody(I) = 0
            Next I
        End If
        .ExtraBodies = 0
        
        ' Kill the pet and reset the alive pet information, but don't remove the pet
        ' from the user available pet list.
        If .MaestroUser > 0 Then Call MatarMascota(.MaestroUser, NpcIndex)
        If .MaestroNpc > 0 Then Call QuitarMascotaNpc(.MaestroNpc)
        If .Owner > 0 Then Call PerdioNpc(.Owner)
        
        If .NumQuests <> 0 Then Erase .Quest
        .NumQuests = 0
        
        .MaestroUser = 0
        .MaestroNpc = 0
        .Owner = 0
        
        .PathFinding = 0
        .Mascotas = 0
        .Movement = 0
        .Name = vbNullString
        .NPCtype = 0
        .Numero = 0
        .Orig.Map = 0
        .Orig.X = 0
        .Orig.Y = 0
        .PoderAtaque = 0
        .PoderEvasion = 0
        .Pos.Map = 0
        .Pos.X = 0
        .Pos.Y = 0
        .SkillDomar = 0
        .Target = 0
        .TargetNPC = 0
        .TipoItems = 0
        .Veneno = 0
        .Desc = vbNullString
        
        .MenuIndex = 0
        
        .ClanIndex = 0
        
        Dim J As Long
        For J = 1 To .NroSpells
            .Spells(J) = 0
        Next J
    End With
    
    Call ResetNpcCharInfo(NpcIndex)
    Call ResetNpcCriatures(NpcIndex)
    Call ResetExpresiones(NpcIndex)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ResetNpcMainInfo de MODULO_NPCs.bas")
End Sub

Public Sub QuitarNPC(ByVal NpcIndex As Integer)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 16/11/2009
'16/11/2009: ZaMa - Now npcs lose their owner
'***************************************************
On Error GoTo ErrHandler

    With Npclist(NpcIndex)
        .flags.NPCActive = False
        
        If InMapBounds(.Pos.Map, .Pos.X, .Pos.Y) Then
            Call EraseNPCChar(NpcIndex)
        End If
    End With
        
    'Nos aseguramos de que el inventario sea removido...
    'asi los lobos no volveran a tirar armaduras ;))
    Call ResetNpcInv(NpcIndex)
    Call ResetNpcFlags(NpcIndex)
    Call ResetNpcCounters(NpcIndex)
    
    Call ResetNpcMainInfo(NpcIndex)
    
    If NpcIndex = LastNPC Then
        Do Until Npclist(LastNPC).flags.NPCActive
            LastNPC = LastNPC - 1
            If LastNPC < 1 Then Exit Do
        Loop
    End If
        
      
    If NumNpcs <> 0 Then
        NumNpcs = NumNpcs - 1
    End If
Exit Sub

ErrHandler:
    Call LogError("Error en QuitarNPC")
End Sub

Public Sub QuitarPet(ByVal UserIndex As Integer, ByVal PetIndex As Byte)
'***************************************************
'Autor: ZaMa
'Last Modification: 18/11/2009
'Kills a pet
'***************************************************
On Error GoTo ErrHandler

    Dim NpcIndex As Integer
    Dim NpcNumber As Integer

    With UserList(UserIndex)
        
        ' Poco probable que pase, pero por las dudas..
        If PetIndex = 0 Then Exit Sub
        
        NpcNumber = .TammedPets(PetIndex).NpcNumber
        
        ' Validate if the petIndex selected contains a tamed pet.
        If NpcNumber = 0 Then
            Call WriteConsoleMsg(UserIndex, "El slot seleccionado no contiene una mascota válida.", FontTypeNames.FONTTYPE_INFO, eMessageType.info)
            Exit Sub
        End If
        
        NpcIndex = .TammedPets(PetIndex).NpcIndex
        
        ' Limpio el slot de la mascota
        .TammedPetsCount = .TammedPetsCount - 1
        .TammedPets(PetIndex).NpcIndex = 0
        .TammedPets(PetIndex).NpcNumber = 0
        .TammedPets(PetIndex).RemainingLife = 0
        
        ' If the pet is spawned, then we erase it.
        If NpcIndex > 0 Then
            Call QuitarNPC(NpcIndex)
        End If
        
        Call WriteConsoleMsg(UserIndex, "Has liberado a tu mascota.", FontTypeNames.FONTTYPE_INFOBOLD, eMessageType.info)
        
        ' Send the pet list to the user
        Call WriteSendPetList(UserIndex)
    End With
    
    Exit Sub

ErrHandler:
    Call LogError("Error en QuitarPet. Error: " & Err.Number & " Desc: " & Err.Description & " NpcIndex: " & NpcIndex & " UserIndex: " & UserIndex & " PetIndex: " & PetIndex)
End Sub

Public Sub QuitarInvocacion(ByVal UserIndex As Integer, ByVal PetIndex As Byte)
'***************************************************
'Autor: Nightw
'Last Modification: 18/11/2009
'Kills an invoked npc.
'***************************************************
On Error GoTo ErrHandler

    Dim NpcIndex As Integer
    Dim NpcNumber As Integer

    With UserList(UserIndex)
        
        ' Poco probable que pase, pero por las dudas..
        If PetIndex = 0 Then Exit Sub
        
        NpcNumber = .InvokedPets(PetIndex).NpcNumber
        
        ' Validate if the petIndex selected contains an invoked pet.
        If NpcNumber = 0 Then
            Exit Sub
        End If
        
        NpcIndex = .InvokedPets(PetIndex).NpcIndex
        If NpcIndex = 0 Then
            Exit Sub
        End If
        
        ' Limpio el slot de la mascota
        .InvokedPetsCount = .InvokedPetsCount - 1
        .InvokedPets(PetIndex).NpcIndex = 0
        .InvokedPets(PetIndex).NpcNumber = 0
        .InvokedPets(PetIndex).RemainingLife = 0
        
        ' Elimino la mascota
        Call QuitarNPC(NpcIndex)
            
    End With
    
    Exit Sub

ErrHandler:
    Call LogError("Error en QuitarPet. Error: " & Err.Number & " Desc: " & Err.Description & " NpcIndex: " & NpcIndex & " UserIndex: " & UserIndex & " PetIndex: " & PetIndex)
End Sub

Private Function TestSpawnTrigger(Pos As WorldPos, Optional PuedeAgua As Boolean = False) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  
    
    If LegalPos(Pos.Map, Pos.X, Pos.Y, PuedeAgua) Then
        TestSpawnTrigger = _
        MapData(Pos.Map, Pos.X, Pos.Y).Trigger <> 3 And _
        MapData(Pos.Map, Pos.X, Pos.Y).Trigger <> 2 And _
        MapData(Pos.Map, Pos.X, Pos.Y).Trigger <> 1
    End If
    
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function TestSpawnTrigger de MODULO_NPCs.bas")
End Function

Public Function CrearNPC(NroNPC As Integer, mapa As Integer, OrigPos As WorldPos, _
                         Optional ByVal CustomHead As Integer, Optional ByVal ForcePos As Boolean = False) As Integer
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

'Crea un NPC del tipo NRONPC

    Dim Pos As WorldPos
    Dim newpos As WorldPos
    Dim altpos As WorldPos
    Dim nIndex As Integer
    Dim PosicionValida As Boolean
    Dim Iteraciones As Long
    Dim PuedeAgua As Boolean
    Dim PuedeTierra As Boolean
    
    Dim tmpPos As Long
    Dim nextPos As Long
    Dim prevPos As Long
    Dim TipoPos As Byte
    
    Dim FirstValidPos As Long
    
    Dim Map As Integer
    Dim X As Integer
    Dim Y As Integer

    nIndex = OpenNPC(NroNPC, True, mapa) 'Conseguimos un indice
    
    If nIndex > MAXNPCS Then Exit Function
    
    ' Cabeza customizada
    If CustomHead <> 0 Then Npclist(nIndex).Char.head = CustomHead
    
    PuedeAgua = Npclist(nIndex).flags.AguaValida
    PuedeTierra = IIf(Npclist(nIndex).flags.TierraInvalida = 1, False, True)
    
    'Necesita ser respawned en un lugar especifico
    If (Npclist(nIndex).flags.RespawnOrigPos Or ForcePos = True) And InMapBounds(OrigPos.Map, OrigPos.X, OrigPos.Y) Then
        
        Map = OrigPos.Map
        X = OrigPos.X
        Y = OrigPos.Y
        Npclist(nIndex).Orig = OrigPos
        Npclist(nIndex).Pos = OrigPos
       
    Else
        
        Pos.Map = mapa 'mapa
        altpos.Map = mapa
        
        If PuedeAgua = True Then
            If PuedeTierra = True Then
                TipoPos = RandomNumber(0, 1)
            Else
                TipoPos = 1
            End If
        Else
            TipoPos = 0
        End If
        
        If UBound(MapInfo(Pos.Map).NpcSpawnPos(TipoPos).Pos) = 0 Then
            If TipoPos = 1 Then
                TipoPos = 0
            Else
                TipoPos = 1
            End If
        End If
        
        tmpPos = RandomNumber(1, UBound(MapInfo(Pos.Map).NpcSpawnPos(TipoPos).Pos))
        
        nextPos = tmpPos
        prevPos = tmpPos
        
        Do While Not PosicionValida
            Pos.X = MapInfo(Pos.Map).NpcSpawnPos(TipoPos).Pos(tmpPos).X
            Pos.Y = MapInfo(Pos.Map).NpcSpawnPos(TipoPos).Pos(tmpPos).Y
            
            If LegalPosNPC(Pos.Map, Pos.X, Pos.Y, PuedeAgua, ForcePos) And TestSpawnTrigger(Pos, PuedeAgua) Then
                If FirstValidPos = 0 Then FirstValidPos = tmpPos
                
                If Not HayPCarea(Pos) Then
                    With Npclist(nIndex)
                        .Pos.Map = Pos.Map
                        .Pos.X = Pos.X
                        .Pos.Y = Pos.Y
                        .Orig = .Pos
                    End With
                    
                    PosicionValida = True
                End If
            End If
            
            If PosicionValida = False Then
                If tmpPos < nextPos Then
                    If nextPos < UBound(MapInfo(Pos.Map).NpcSpawnPos(TipoPos).Pos) Then
                        nextPos = nextPos + 1
                        tmpPos = nextPos
                    Else
                        If prevPos > 1 Then
                            prevPos = prevPos - 1
                            tmpPos = prevPos
                        Else
                            If FirstValidPos > 0 Then
                                With Npclist(nIndex)
                                    .Pos.Map = Pos.Map
                                    .Pos.X = MapInfo(Pos.Map).NpcSpawnPos(TipoPos).Pos(FirstValidPos).X
                                    .Pos.Y = MapInfo(Pos.Map).NpcSpawnPos(TipoPos).Pos(FirstValidPos).Y
                                    .Orig = .Pos
                                End With
                                
                                PosicionValida = True
                            Else
                                Exit Function
                            End If
                        End If
                    End If
                Else
                    If prevPos > 1 Then
                        prevPos = prevPos - 1
                        tmpPos = prevPos
                    Else
                        If nextPos < UBound(MapInfo(Pos.Map).NpcSpawnPos(TipoPos).Pos) Then
                            nextPos = nextPos + 1
                            tmpPos = nextPos
                        Else
                            If FirstValidPos > 0 Then
                                With Npclist(nIndex)
                                    .Pos.Map = Pos.Map
                                    .Pos.X = MapInfo(Pos.Map).NpcSpawnPos(TipoPos).Pos(FirstValidPos).X
                                    .Pos.Y = MapInfo(Pos.Map).NpcSpawnPos(TipoPos).Pos(FirstValidPos).Y
                                    .Orig = .Pos
                                End With
                                
                                PosicionValida = True
                            Else
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End If
        Loop
            
        'asignamos las nuevas coordenas
        Map = Pos.Map
        X = Npclist(nIndex).Pos.X
        Y = Npclist(nIndex).Pos.Y
    End If
            
    'Crea el NPC
    Call MakeNPCChar(True, Map, nIndex, Map, X, Y)
    
    CrearNPC = nIndex
    
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function CrearNPC de MODULO_NPCs.bas")
End Function

Public Sub MakeNPCChar(ByVal toMap As Boolean, ByVal sndIndex As Integer, ByVal NpcIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 17/09/2014
'
'17/09/2014: D'Artagnan - Send Hostile and Merchant attributes.
'13/07/2016: Anagrama - Ahora se crea con su escudo, casco y arma.
'27/07/2016: Anagrama - Envia nombre si se debe ver el nombre en pantalla.
'***************************************************
On Error GoTo ErrHandler
  
    
    Dim CharIndex As Integer
    Dim Name As String
    Dim Tag As String
    
    If Npclist(NpcIndex).Char.CharIndex = 0 Then
        CharIndex = NextOpenCharIndex
        Npclist(NpcIndex).Char.CharIndex = CharIndex
        charList(CharIndex) = NpcIndex
    End If
    
    MapData(Map, X, Y).NpcIndex = NpcIndex

    If Not toMap Then
        With Npclist(NpcIndex)
            If .flags.ShowName Then
                Name = .Name
                If LenB(.Tag) > 0 Then
                    Name = Name & " <" & .Tag & ">"
                End If
            End If
                        
            Call WriteCharacterCreate(sndIndex, .Char.body, .Char.head, .Char.heading, .Char.CharIndex, X, Y, _
                                      .Char.WeaponAnim, .Char.ShieldAnim, 0, 0, .Char.CascoAnim, Name, 0, eCharacterAlignment.Neutral, 0, CBool(.Hostile), CBool(.Comercia), NpcData(Npclist(NpcIndex).Numero).flags.TierraInvalida = 1, Npclist(NpcIndex).Numero, .OverHeadIcon)
        End With
    Else
        Call ModAreas.CreateEntity(NpcIndex, ENTITY_TYPE_NPC, Npclist(NpcIndex).Pos, Npclist(NpcIndex).SizeWidth, Npclist(NpcIndex).SizeHeight)
    End If
  
    Npclist(NpcIndex).InstanceId = GetTickCount()
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub MakeNPCChar de MODULO_NPCs.bas")
End Sub

Public Sub ChangeNPCChar(ByVal NpcIndex As Integer, ByVal body As Integer, ByVal head As Integer, ByVal heading As eHeading)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler

    If NpcIndex <= 0 Then
        Exit Sub
    End If
    
    With Npclist(NpcIndex).Char
        .body = body
        .head = head
        .heading = heading
        Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCharacterChange(body, head, heading, .CharIndex, 0, 0, 0, 0, 0, CBool(Npclist(NpcIndex).flags.TierraInvalida = 1), False, Npclist(NpcIndex).OverHeadIcon, eCharacterAlignment.Neutral))
    End With
      
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ChangeNPCChar de MODULO_NPCs.bas")
End Sub

Private Sub EraseNPCChar(ByVal NpcIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

If Npclist(NpcIndex).Char.CharIndex <> 0 Then charList(Npclist(NpcIndex).Char.CharIndex) = 0

If Npclist(NpcIndex).Char.CharIndex = LastChar Then
    Do Until charList(LastChar) > 0
        LastChar = LastChar - 1
        If LastChar <= 1 Then Exit Do
    Loop
End If

'Actualizamos el area
Call ModAreas.DeleteEntity(NpcIndex, ENTITY_TYPE_NPC)
      
'Quitamos del mapa
MapData(Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y).NpcIndex = 0

'Update la lista npc
Npclist(NpcIndex).Char.CharIndex = 0


'update NumChars
NumChars = NumChars - 1


  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub EraseNPCChar de MODULO_NPCs.bas")
End Sub

Public Function MoveNPCChar(ByVal NpcIndex As Integer, ByVal nHeading As Byte) As Boolean
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 06/04/2009
'06/04/2009: ZaMa - Now npcs can force to change position with dead character
'01/08/2009: ZaMa - Now npcs can't force to chance position with a dead character if that means to change the terrain the character is in
'26/09/2010: ZaMa - Turn sub into function to know if npc has moved or not.
'***************************************************

On Error GoTo errh

    Dim nPos As WorldPos
    Dim UserIndex As Integer
    Dim isZonaOscura As Boolean
    Dim isZonaOscuraNewPos As Boolean
    
    With Npclist(NpcIndex)
        nPos = .Pos
        Call HeadtoPos(nHeading, nPos)
        
        ' es una posicion legal
        If LegalPosNPC(nPos.Map, nPos.X, nPos.Y, .flags.AguaValida = 1, .MaestroUser <> 0, .flags.TierraInvalida) Then
            
            If .flags.AguaValida = 0 And HayAgua(.Pos.Map, nPos.X, nPos.Y) Then Exit Function
            If .flags.TierraInvalida = 1 And Not HayAgua(.Pos.Map, nPos.X, nPos.Y) Then Exit Function
            
            isZonaOscura = (MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger = eTrigger.zonaOscura)
            isZonaOscuraNewPos = (MapData(nPos.Map, nPos.X, nPos.Y).Trigger = eTrigger.zonaOscura)
            
            UserIndex = MapData(.Pos.Map, nPos.X, nPos.Y).UserIndex
            ' Si hay un usuario a donde se mueve el npc, entonces esta muerto
            If UserIndex > 0 Then
                
                ' No se traslada caspers de agua a tierra
                If HayAgua(.Pos.Map, nPos.X, nPos.Y) And Not HayAgua(.Pos.Map, .Pos.X, .Pos.Y) Then Exit Function
                ' No se traslada caspers de tierra a agua
                If Not HayAgua(.Pos.Map, nPos.X, nPos.Y) And HayAgua(.Pos.Map, .Pos.X, .Pos.Y) Then Exit Function
                
                With UserList(UserIndex)
                    ' Actualizamos posicion y mapa
                    MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex = 0
                    .Pos.X = Npclist(NpcIndex).Pos.X
                    .Pos.Y = Npclist(NpcIndex).Pos.Y
                    MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex = UserIndex
                        
                    ' Si es un admin invisible, no se avisa a los demas clientes
                    If Not (.flags.AdminInvisible = 1) Then
                        'Los valores de visible o invisible están invertidos porque estos flags son del NpcIndex, por lo tanto si el npc entra, el casper sale y viceversa :P
                        If isZonaOscura Then
                            If Not isZonaOscuraNewPos Then
                                Call SendData(SendTarget.ToUsersAndRmsAndCounselorsAreaButGMs, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, True))
                            End If
                        Else
                            If isZonaOscuraNewPos Then
                                Call SendData(SendTarget.ToUsersAndRmsAndCounselorsAreaButGMs, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
                            End If
                        End If
                    End If
                    
                    nHeading = InvertHeading(nHeading)

                    'Forzamos al usuario a moverse
                    Call WriteForceCharMove(UserIndex, nHeading)
                    
                    'Actualizamos las áreas de ser necesario
                    Call ModAreas.UpdateEntity(UserIndex, ENTITY_TYPE_PLAYER, .Pos, False)
                End With
            End If

            'Update map and user pos
            MapData(.Pos.Map, .Pos.X, .Pos.Y).NpcIndex = 0
            .Pos = nPos
            .Char.heading = nHeading
            MapData(.Pos.Map, nPos.X, nPos.Y).NpcIndex = NpcIndex
     
            Call ModAreas.UpdateEntity(NpcIndex, ENTITY_TYPE_NPC, .Pos, False)
            
            If isZonaOscura Then
                If Not isZonaOscuraNewPos Then
                    If (.flags.invisible = 0) Then
                        Call SendData(SendTarget.ToNPCAreaButCounselors, NpcIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
                    End If
                End If
            Else
                If isZonaOscuraNewPos Then
                    If (.flags.invisible = 0) Then
                        Call SendData(SendTarget.ToNPCAreaButCounselors, NpcIndex, PrepareMessageSetInvisible(.Char.CharIndex, True))
                    End If
                End If
            End If

            ' Step into trap?
            Call CheckTriggerActivation(0, NpcIndex, .Pos.Map, .Pos.X, .Pos.Y, False)
            
            If .flags.DistanciaMaxima Then
                If .flags.VolviendoOrig = 0 Then
                    If Distance(.Pos.X, .Pos.Y, .Orig.X, .Orig.Y) > .flags.DistanciaMaxima Then
                        .flags.VolviendoOrig = 1
                    End If
                Else
                    If Distance(.Pos.X, .Pos.Y, .Orig.X, .Orig.Y) = 0 Then
                        .flags.VolviendoOrig = 0
                        .flags.VolviendoInt = 0
                    End If
                End If
            End If
            
            ' Npc has moved
            MoveNPCChar = True
        
        ElseIf .MaestroUser = 0 Then
            If .flags.VolviendoOrig Then
                .flags.VolviendoInt = .flags.VolviendoInt + 1
                If .flags.VolviendoInt > 10 Then
                    .flags.VolviendoOrig = 0
                    .flags.VolviendoInt = 0
                End If
            End If
            If .Movement = TipoAI.NpcPathfinding Then
                'Someone has blocked the npc's way, we must to seek a new path!
                .PFINFO.PathLenght = 0
            End If
        End If
    End With
    
    Exit Function

errh:
    LogError ("Error en move npc " & NpcIndex & ". Error: " & Err.Number & " - " & Err.Description)
End Function

Function NextOpenNPC() As Integer
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo ErrHandler
    Dim LoopC As Long
      
    For LoopC = 1 To MAXNPCS + 1
        If LoopC > MAXNPCS Then Exit For
        If Not Npclist(LoopC).flags.NPCActive Then Exit For
    Next LoopC
      
    NextOpenNPC = LoopC
Exit Function

ErrHandler:
    Call LogError("Error en NextOpenNPC")
End Function

Sub NpcEnvenenarUser(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 10/07/2010
'10/07/2010: ZaMa - Now npcs can't poison dead users.
'***************************************************
On Error GoTo ErrHandler
  

    Dim N As Integer
    
    With UserList(UserIndex)
        If .flags.Muerto = 1 Then Exit Sub
        
        N = RandomNumber(1, 100)
        If N < 30 Then
            .flags.Envenenado = 1
            Call WriteConsoleMsg(UserIndex, "¡¡La criatura te ha envenenado!!", FontTypeNames.FONTTYPE_FIGHT, _
                                 eMessageType.Combate)
        End If
    End With
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub NpcEnvenenarUser de MODULO_NPCs.bas")
End Sub

Function SpawnNpc(ByVal NpcIndex As Integer, Pos As WorldPos, ByVal FX As Boolean, ByVal Respawn As Boolean) As Integer
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 06/15/2008
'23/01/2007 -> Pablo (ToxicWaste): Creates an NPC of the type Npcindex
'06/15/2008 -> Optimizé el codigo. (NicoNZ)
'***************************************************
On Error GoTo ErrHandler
  
Dim newpos As WorldPos
Dim altpos As WorldPos
Dim nIndex As Integer
Dim PosicionValida As Boolean
Dim PuedeAgua As Boolean
Dim PuedeTierra As Boolean


Dim Map As Integer
Dim X As Integer
Dim Y As Integer

nIndex = OpenNPC(NpcIndex, Respawn, Pos.Map)      'Conseguimos un indice

If nIndex > MAXNPCS Then
    SpawnNpc = 0
    Exit Function
End If

PuedeAgua = Npclist(nIndex).flags.AguaValida
PuedeTierra = Not Npclist(nIndex).flags.TierraInvalida = 1
        
Call ClosestLegalPos(Pos, newpos, PuedeAgua, PuedeTierra, , True) 'Nos devuelve la posicion valida mas cercana
Call ClosestLegalPos(Pos, altpos, PuedeAgua)

'fixme: here we should check if spawn tile has teleport or not (only for pets or any other npc?)

'Si X e Y son iguales a 0 significa que no se encontro posicion valida
If newpos.X <> 0 And newpos.Y <> 0 Then
    'Asignamos las nuevas coordenas solo si son validas
    Npclist(nIndex).Pos.Map = newpos.Map
    Npclist(nIndex).Pos.X = newpos.X
    Npclist(nIndex).Pos.Y = newpos.Y
    PosicionValida = True
Else
    If altpos.X <> 0 And altpos.Y <> 0 Then
        Npclist(nIndex).Pos.Map = altpos.Map
        Npclist(nIndex).Pos.X = altpos.X
        Npclist(nIndex).Pos.Y = altpos.Y
        PosicionValida = True
    Else
        PosicionValida = False
    End If
End If

If Not PosicionValida Then
    Call QuitarNPC(nIndex)
    SpawnNpc = 0
    Exit Function
End If

'If Npclist(nIndex).flags.RespawnOrigPos Then
    Npclist(nIndex).Orig.Map = Npclist(nIndex).Pos.Map
    Npclist(nIndex).Orig.X = Npclist(nIndex).Pos.X
    Npclist(nIndex).Orig.Y = Npclist(nIndex).Pos.Y
'End If

'asignamos las nuevas coordenas
Map = newpos.Map
X = Npclist(nIndex).Pos.X
Y = Npclist(nIndex).Pos.Y

'Crea el NPC
Call MakeNPCChar(True, Map, nIndex, Map, X, Y)

If FX Then
    Call SendData(SendTarget.ToNPCArea, nIndex, PrepareMessagePlayWave(ConstantesSonidos.Warp, X, Y, Npclist(nIndex).Char.CharIndex))
    Call SendData(SendTarget.ToNPCArea, nIndex, PrepareMessageCreateFX(Npclist(nIndex).Char.CharIndex, ConstantesFX.FxWarp, 0))
End If

SpawnNpc = nIndex

  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function SpawnNpc de MODULO_NPCs.bas")
End Function

Sub ReSpawnNpc(MiNPC As npc)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    If (MiNPC.flags.Respawn = 0) Then Call CrearNPC(MiNPC.Numero, MiNPC.Pos.Map, MiNPC.Orig)

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ReSpawnNpc de MODULO_NPCs.bas")
End Sub

Public Function OpenNPC(ByVal NpcNumber As Integer, ByVal Respawn As Boolean, ByVal Map As Integer) As Integer
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  
    Dim NpcIndex As Integer
    Dim LoopC As Long
    
    'If requested index is invalid, abort
    If NpcNumber > NumNpcsDat Or NpcNumber <= 0 Then
        OpenNPC = MAXNPCS + 1
        Exit Function
    End If
    
    If Not NpcData(NpcNumber).Exists Then
        OpenNPC = MAXNPCS + 1
        Exit Function
    End If
    
    NpcIndex = NextOpenNPC
    
    If NpcIndex > MAXNPCS Then 'Limite de npcs
        OpenNPC = NpcIndex
        Exit Function
    End If
    
    With Npclist(NpcIndex)
        .NumInvocaciones = NpcData(NpcNumber).NumInvocaciones
        If .NumInvocaciones > 0 Then
            ReDim .NpcsInvocables(1 To .NumInvocaciones) As Integer
            For LoopC = 1 To .NumInvocaciones
                .NpcsInvocables(LoopC) = NpcData(NpcNumber).NpcsInvocables(LoopC)
            Next LoopC
        End If
        
        .flags.MaxInvocaciones = NpcData(NpcNumber).flags.MaxInvocaciones
        If .flags.MaxInvocaciones > 0 Then
            ReDim .flags.Invocacion(1 To .flags.MaxInvocaciones) As Integer
        End If
        
        .ExtraBodies = NpcData(NpcNumber).ExtraBodies
        If .ExtraBodies > 0 Then
            ReDim .ExtraBody(1 To .ExtraBodies) As Integer
            For LoopC = 1 To .ExtraBodies
                .ExtraBody(LoopC) = NpcData(NpcNumber).ExtraBody(LoopC)
            Next LoopC
        End If
        .ActualBody = 0
        
        .Numero = NpcNumber
        .Name = NpcData(NpcNumber).Name
        .Desc = NpcData(NpcNumber).Desc
        
        .Tag = NpcData(NpcNumber).Tag
        
        .PathFinding = NpcData(NpcNumber).PathFinding
        .Movement = NpcData(NpcNumber).Movement
        .flags.OldMovement = .Movement
        
        .NPCtype = NpcData(NpcNumber).NPCtype
        
        .Char.body = NpcData(NpcNumber).Char.body
        .Char.head = NpcData(NpcNumber).Char.head
        .Char.heading = NpcData(NpcNumber).Char.heading
        
        .Char.WeaponAnim = NpcData(NpcNumber).Char.WeaponAnim
        .Char.ShieldAnim = NpcData(NpcNumber).Char.ShieldAnim
        .Char.CascoAnim = NpcData(NpcNumber).Char.CascoAnim
        
        .Attackable = NpcData(NpcNumber).Attackable
        .Comercia = NpcData(NpcNumber).Comercia
        .Hostile = NpcData(NpcNumber).Hostile
        .flags.OldHostil = .Hostile
        
        If MapInfo(Map).MapaTierra = 0 Then
            .GiveEXP = NpcData(NpcNumber).GiveEXP
        Else
            .GiveEXP = NpcData(NpcNumber).GiveEXPTierra
        End If
        
        .flags.ExpCount = .GiveEXP
        
        .Veneno = NpcData(NpcNumber).Veneno
        
        .GiveGLD = NpcData(NpcNumber).GiveGLD
        
        ' Load quests
        Dim NumQuests As Integer
        NumQuests = NpcData(NpcNumber).NumQuests
        .NumQuests = NumQuests
        
        If NumQuests <> 0 Then
            ReDim .Quest(1 To NumQuests) As Integer
            
            'For LoopC = 1 To NumQuests
            '    .Quest(LoopC) = NpcData(NpcNumber).Quest(LoopC)
            'Next LoopC
        End If

        
        .PoderAtaque = NpcData(NpcNumber).PoderAtaque
        .PoderEvasion = NpcData(NpcNumber).PoderEvasion
        
        .InvReSpawn = NpcData(NpcNumber).InvReSpawn
        
        With .Stats
            .MaxHp = NpcData(NpcNumber).Stats.MaxHp
            .MinHp = NpcData(NpcNumber).Stats.MinHp
            .MaxHit = NpcData(NpcNumber).Stats.MaxHit
            .MinHit = NpcData(NpcNumber).Stats.MinHit
            .Def = NpcData(NpcNumber).Stats.Def
            .DefM = NpcData(NpcNumber).Stats.DefM
            .Alineacion = NpcData(NpcNumber).Stats.Alineacion
        End With
        
        With .Invent
            .NroItems = NpcData(NpcNumber).Invent.NroItems
            For LoopC = 1 To .NroItems
                .Object(LoopC).ObjIndex = NpcData(NpcNumber).Invent.Object(LoopC).ObjIndex
                .Object(LoopC).Amount = NpcData(NpcNumber).Invent.Object(LoopC).Amount
            Next LoopC
        End With
        
        .NroDrops = NpcData(NpcNumber).NroDrops
        If .NroDrops > 0 Then
            ReDim .Drop(1 To .NroDrops) As tDrops
            For LoopC = 1 To .NroDrops
                .Drop(LoopC).DropIndex = NpcData(NpcNumber).Drop(LoopC).DropIndex
                .Drop(LoopC).Probabilidad = NpcData(NpcNumber).Drop(LoopC).Probabilidad
                .Drop(LoopC).NoExcluyente = NpcData(NpcNumber).Drop(LoopC).NoExcluyente
            Next LoopC
        End If
        
        .flags.LanzaSpells = NpcData(NpcNumber).flags.LanzaSpells
        If .flags.LanzaSpells > 0 Then ReDim .Spells(1 To .flags.LanzaSpells)
        For LoopC = 1 To .flags.LanzaSpells
            .Spells(LoopC) = NpcData(NpcNumber).Spells(LoopC)
        Next LoopC
        
        .NroCriaturas = NpcData(NpcNumber).NroCriaturas
        If .NPCtype = eNPCType.Entrenador And NpcData(NpcNumber).NroCriaturas > 0 Then

            ReDim .Criaturas(1 To .NroCriaturas) As tCriaturasEntrenador
            For LoopC = 1 To .NroCriaturas
                .Criaturas(LoopC).NpcIndex = NpcData(NpcNumber).Criaturas(LoopC).NpcIndex
                .Criaturas(LoopC).NpcName = NpcData(NpcNumber).Criaturas(LoopC).NpcName
            Next LoopC
        End If
        
        With .flags
            .AguaValida = NpcData(NpcNumber).flags.AguaValida
            .TierraInvalida = NpcData(NpcNumber).flags.TierraInvalida
            .Faccion = NpcData(NpcNumber).flags.Faccion
            .AtacaDoble = NpcData(NpcNumber).flags.AtacaDoble
            .Domable = NpcData(NpcNumber).flags.Domable
            
            .ShowName = NpcData(NpcNumber).flags.ShowName
            
            .NPCActive = True
            
            If Respawn Then
                .Respawn = NpcData(NpcNumber).flags.Respawn
            Else
                .Respawn = 1
            End If
            
            .BackUp = NpcData(NpcNumber).flags.BackUp
            .RespawnOrigPos = NpcData(NpcNumber).flags.RespawnOrigPos
            
            .AfectaParalisis = NpcData(NpcNumber).flags.AfectaParalisis
            
            .DistanciaMaxima = NpcData(NpcNumber).flags.DistanciaMaxima
                        
            .Snd1 = NpcData(NpcNumber).flags.Snd1
            .Snd2 = NpcData(NpcNumber).flags.Snd2
            .Snd3 = NpcData(NpcNumber).flags.Snd3
        End With
        
        With .Intervalos
            .Walk = NpcData(NpcNumber).Intervalos.Walk
            .Hit = NpcData(NpcNumber).Intervalos.Hit
            .MoveAttack = ConstantesBalance.IntMoveAttack
        End With

        Set .Timers = New clsTimers
        Call .Timers.Initialize(NpcIndex)
        
        If .flags.RespawnOrigPos Then
            .Orig.Map = NpcData(NpcNumber).Orig.Map
            .Orig.X = NpcData(NpcNumber).Orig.X
            .Orig.Y = NpcData(NpcNumber).Orig.Y
        End If
    
        '<<<<<<<<<<<<<< Expresiones >>>>>>>>>>>>>>>>
        .NroExpresiones = NpcData(NpcNumber).NroExpresiones
        If .NroExpresiones > 0 Then ReDim .Expresiones(1 To .NroExpresiones) As String
        For LoopC = 1 To .NroExpresiones
            .Expresiones(LoopC) = NpcData(NpcNumber).Expresiones(LoopC)
        Next LoopC
        '<<<<<<<<<<<<<< Expresiones >>>>>>>>>>>>>>>>
        
        .MenuIndex = NpcData(NpcNumber).MenuIndex
        
        'Tipo de items con los que comercia
        .TipoItems = NpcData(NpcNumber).TipoItems
        
        .Ciudad = NpcData(NpcNumber).Ciudad
        .level = NpcData(NpcNumber).level
        .OffsetReducedExp = NpcData(NpcNumber).OffsetReducedExp
        .OffsetModificator = NpcData(NpcNumber).OffsetModificator
        .MasteryStarter = NpcData(NpcNumber).MasteryStarter
        .OverHeadIcon = NpcData(NpcNumber).OverHeadIcon
        .SizeWidth = NpcData(NpcNumber).SizeWidth
        .SizeHeight = NpcData(NpcNumber).SizeHeight
    End With
    
    'Update contadores de NPCs
    If NpcIndex > LastNPC Then LastNPC = NpcIndex
    NumNpcs = NumNpcs + 1
    
    'Devuelve el nuevo Indice
    OpenNPC = NpcIndex
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") para Npc " & NpcNumber & " en mapa " & Map & " en Function OpenNPC de MODULO_NPCs.bas")
End Function

Public Sub DoFollow(ByVal NpcIndex As Integer, ByVal UserName As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    With Npclist(NpcIndex)
        If .flags.Follow Then
            .flags.AttackedBy = vbNullString
            .flags.Follow = False
            .Movement = .flags.OldMovement
            .Hostile = .flags.OldHostil
        Else
            .flags.AttackedBy = UserName
            .flags.Follow = True
            .Movement = TipoAI.NPCDEFENSA
            .Hostile = 0
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub DoFollow de MODULO_NPCs.bas")
End Sub

Public Sub FollowAmo(ByVal NpcIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    With Npclist(NpcIndex)
        .flags.Follow = True
        
        .flags.OldMovement = .Movement
        .flags.OldHostil = .Hostile
        
        .Movement = TipoAI.SigueAmo
        .Hostile = 0
        .Target = 0
        .TargetNPC = 0
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub FollowAmo de MODULO_NPCs.bas")
End Sub

Public Sub ValidarPermanenciaNpc(ByVal NpcIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'Chequea si el npc continua perteneciendo a algún usuario
'***************************************************
On Error GoTo ErrHandler
  

    With Npclist(NpcIndex)
        If IntervaloPerdioNpc(.Owner) Then Call PerdioNpc(.Owner)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ValidarPermanenciaNpc de MODULO_NPCs.bas")
End Sub

Public Function ShouldApplyExpMod(ByVal nUserIndex As Integer, ByVal nNPCIndex As Integer) As Boolean
'***************************************************
'Author: D'Artagnan
'Last Modification: 29/03/2015
'
'***************************************************
On Error GoTo ErrHandler
  
    If nUserIndex > 0 Then
        With Npclist(nNPCIndex)
            ShouldApplyExpMod = .level > 0 And (CInt(UserList(nUserIndex).Stats.ELV) - CInt(.level) >= .OffsetReducedExp)
        End With
    Else
        ShouldApplyExpMod = False
    End If
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function ShouldApplyExpMod de MODULO_NPCs.bas")
End Function

Public Function ApplyExperienceModifier(ByVal nUserIndex As Integer, ByVal nNPCIndex As Integer, _
                                        ByVal nExperience As Long) As Long
'***************************************************
'Author: D'Artagnan
'Last Modification: 29/03/2015
'
'***************************************************
On Error GoTo ErrHandler
  
    ApplyExperienceModifier = nExperience
    With Npclist(nNPCIndex)
        If ShouldApplyExpMod(nUserIndex, nNPCIndex) Then
            ApplyExperienceModifier = ApplyExperienceModifier * .OffsetModificator
        End If
    End With
    
    ApplyExperienceModifier = ApplyExperienceModifier * ConstantesBalance.ModExpMultiplier
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function ApplyExperienceModifier de MODULO_NPCs.bas")
End Function

Public Sub DoNpcInvocacion(ByVal NpcIndex As Integer, ByRef NpcPos As WorldPos)
'---------------------------------------------------------------------------------------
' Module    : AI
' Author    : Anagrama
' Date      : 16/07/2016
' Purpose   : Invoca un npc al azar de su lista si no invocó el máximo y si lo tiene permitido.
'---------------------------------------------------------------------------------------
On Error GoTo ErrHandler
  
    Dim I As Byte

    With Npclist(NpcIndex)
        If .flags.MaxInvocaciones = 0 Then Exit Sub
        
        I = 1
        Do While I <= .flags.MaxInvocaciones
            If .flags.Invocacion(I) = 0 Then
                .flags.Invocacion(I) = SpawnNpc(Int(.NpcsInvocables(RandomNumber(1, UBound(.NpcsInvocables)))), NpcPos, True, False)
                Npclist(.flags.Invocacion(I)).flags.Invocador = NpcIndex
                Exit Do
            End If
            I = I + 1
        Loop
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub DoNpcInvocacion de MODULO_NPCs.bas")
End Sub

Public Sub CheckNpcInvocaciones(ByVal NpcIndex As Integer)
'---------------------------------------------------------------------------------------
' Procedure : AI
' Author    : Anagrama
' Date      : 16/07/2016
' Purpose   : Revisa si hay que matar invocaciones.
'---------------------------------------------------------------------------------------
On Error GoTo ErrHandler
  
    Dim I As Integer

        With Npclist(NpcIndex)
            If .Contadores.TiempoExistencia >= 100 Then
                For I = 1 To Npclist(.flags.Invocador).flags.MaxInvocaciones
                    If Npclist(.flags.Invocador).flags.Invocacion(I) = NpcIndex Then
                        Npclist(.flags.Invocador).flags.Invocacion(I) = 0
                        Exit For
                    End If
                Next I
                
                Call QuitarNPC(NpcIndex)
            Else
                .Contadores.TiempoExistencia = .Contadores.TiempoExistencia + 1
            End If
        End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CheckNpcInvocaciones de MODULO_NPCs.bas")
End Sub
Public Function GetOlderInvokedPetIndex(ByVal UserIndex As Integer) As Integer
    Dim I As Integer
    Dim Counter As Long
    
On Error GoTo ErrHandler

    With UserList(UserIndex)
        For I = 1 To .InvokedPetsCount
            If Counter = 0 Then
                Counter = Npclist(.InvokedPets(I).NpcIndex).Contadores.TiempoExistencia
                GetOlderInvokedPetIndex = I
            Else
                If Npclist(.InvokedPets(I).NpcIndex).Contadores.TiempoExistencia < Counter Then
                    Counter = Npclist(.InvokedPets(I).NpcIndex).Contadores.TiempoExistencia
                    GetOlderInvokedPetIndex = I
                End If
            End If
        Next I
    End With
    
    Exit Function
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetOlderInvokedPetIndex de MODULO_NPCs.bas")
End Function

Public Function GetInvokedPetIndexByNpcIndex(ByVal UserIndex As Integer, ByVal NpcIndex As Integer) As Integer
    Dim PetIndex As Integer
    Dim I As Integer
    
On Error GoTo ErrHandler

    With UserList(UserIndex)
        For I = 1 To Classes(.clase).ClassMods.MaxInvokedPets
            If .InvokedPets(I).NpcIndex = NpcIndex Then
                GetInvokedPetIndexByNpcIndex = I
                Exit Function
            End If
        Next I
    End With

    Exit Function
    
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetInvokedPetIndexByNpcIndex de MODULO_NPCs.bas")
End Function

Public Function GetTammedPetIndexByNpcIndex(ByVal UserIndex As Integer, ByVal NpcIndex As Integer) As Integer
    Dim PetIndex As Integer
    Dim I As Integer
    
On Error GoTo ErrHandler

    With UserList(UserIndex)
        For I = 1 To Classes(.clase).ClassMods.MaxTammedPets
            If .TammedPets(I).NpcIndex = NpcIndex Then
                GetTammedPetIndexByNpcIndex = I
                Exit Function
            End If
        Next I
    End With

    Exit Function
    
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetTammedPetIndexByNpcIndex de MODULO_NPCs.bas")
End Function
