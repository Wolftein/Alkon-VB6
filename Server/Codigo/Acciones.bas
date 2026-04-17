Attribute VB_Name = "Acciones"

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
' Modulo para manejar las acciones (doble click) de los carteles, foro, puerta, ramitas
'

''
' Ejecuta la accion del doble click
'
' @param UserIndex UserIndex
' @param Map Numero de mapa
' @param X X
' @param Y Y

Sub Accion(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    Dim tempIndex As Integer
    Dim GuildIndex As Integer
    Dim NewStageNumber As Integer
    Dim QuestId As Integer
    Dim CurrentStage As Integer
    Dim ObjIndex As Integer
    Dim ObjSlot As Integer
    
On Error Resume Next
    '¿Rango Visión? (ToxicWaste)
    If (Abs(UserList(UserIndex).Pos.Y - Y) > RANGO_VISION_Y) Or (Abs(UserList(UserIndex).Pos.X - X) > RANGO_VISION_X) Then
        Exit Sub
    End If
    
    '¿Posicion valida?
    If InMapBounds(Map, X, Y) Then
        With UserList(UserIndex)
            If MapData(Map, X, Y).UserIndex > 0 Then     'Acciones Usuarios
                tempIndex = MapData(Map, X, Y).UserIndex
                
                If tempIndex = 0 Then Exit Sub
                
                ' Can 't open a store of another player if yours is open.
                If UserList(tempIndex).CraftingStore.IsOpen And Not .CraftingStore.IsOpen Then
                    Call WriteWorkerStore_Show(UserIndex, tempIndex)
                    Exit Sub
                End If
            
            End If
        
            If MapData(Map, X, Y).NpcIndex > 0 Then     'Acciones NPCs
                tempIndex = MapData(Map, X, Y).NpcIndex
                
                'Set the target NPC
                .flags.TargetNPC = tempIndex
                
                GuildIndex = UserList(UserIndex).guild.GuildIndex
                
                'quest actions
                 If GuildHasQuest(GuildIndex) Then
                    If GetQuestNpcEndIndex(GuildIndex) = Npclist(tempIndex).Numero Then
                        
                        If CheckIfCurrentStageIsCompleted(GuildIndex) Then
                            Call WriteChatOverHead(UserIndex, GuildQuestList(GuildList(GuildIndex).CurrentQuest.IdQuest).Stages(GuildList(GuildIndex).CurrentQuest.CurrentStage).EndNpc.desc, Npclist(tempIndex).Char.CharIndex, vbWhite)
                            Call FinishGuildQuestStage(GuildIndex, UserIndex)
                        End If
                        
                        Exit Sub
                    End If
                End If
                
                'check if this is al starter quest NPC (stage = 0)
                If UserList(UserIndex).Guild.IdGuild > 0 Then
                    GuildIndex = UserList(UserIndex).Guild.GuildIndex
                    If GuildList(GuildIndex).CurrentQuest.IdQuest > 0 Then
                        If GuildList(GuildIndex).CurrentQuest.CurrentStage = 0 Then
                            If Npclist(tempIndex).Numero = GuildQuestList(GuildList(GuildIndex).CurrentQuest.IdQuest).Stages(1).StarterNpc.NpcIndex Then
                                Call WriteChatOverHead(UserIndex, GuildQuestList(GuildList(GuildIndex).CurrentQuest.IdQuest).Stages(1).StarterNpc.desc, Npclist(tempIndex).Char.CharIndex, vbWhite)
                                Call UpdateCurrentQuestInfoToOnlineMembers(GuildIndex)
                                Call ChangeGuildQuestStage(GuildIndex, GuildList(GuildIndex).CurrentQuest.IdQuest, 1)
                                Call UpdateCurrentQuestInfoToOnlineMembers(GuildIndex)
                                Exit Sub
                            End If
                        End If
                    End If
                End If
                
                If Npclist(tempIndex).Comercia = 1 Then
                    If Not CommerceAllowed(UserIndex) Then Exit Sub
                    
                    If Distancia(Npclist(tempIndex).Pos, .Pos) > 3 Then
                        Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    'Iniciamos la rutina pa' comerciar.
                    Call IniciarComercioNPC(UserIndex)
                
                ElseIf Npclist(tempIndex).NPCtype = eNPCType.Banquero Then
                    If Not CommerceAllowed(UserIndex) Then Exit Sub
                    
                    If Distancia(Npclist(tempIndex).Pos, .Pos) > 3 Then
                        Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    'A depositar de una
                    Call IniciarDeposito(UserIndex)
                
                ElseIf Npclist(tempIndex).NPCtype = eNPCType.Revividor Or Npclist(tempIndex).NPCtype = eNPCType.ResucitadorNewbie Then
                    If Distancia(.Pos, Npclist(tempIndex).Pos) > 10 Then
                        Call WriteConsoleMsg(UserIndex, "El sacerdote no puede curarte debido a que estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                                     
                    If Npclist(tempIndex).NPCtype = eNPCType.Revividor Then
                        If .flags.Muerto = 1 Then
                            Call RevivirUsuario(UserIndex, False)
                        End If
                        
                        'curamos totalmente
                        .Stats.MinHp = .Stats.MaxHp
                        .Stats.MinMAN = .Stats.MaxMan
                        .Stats.MinSta = .Stats.MaxSta
                        Call WriteUpdateUserStats(UserIndex)
                    End If
                    
                ElseIf Npclist(tempIndex).NPCtype = eNPCType.Entrenador Then
                    '¿Esta el user muerto? Si es asi no puede sacar npcs
                    If .flags.Muerto = 1 Then
                        Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    Call WriteTrainerCreatureList(UserIndex, .flags.TargetNPC)
                ElseIf Npclist(tempIndex).MasteryStarter Then
                    If .flags.Muerto = 1 Then
                        Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    If .Stats.ELV < ConstantesBalance.MaxLvl Then
                        Call WriteChatOverHead(UserIndex, "No eres lo suficientemente poderoso como para aprender una maestría. Regresa cuando seas nivel " & ConstantesBalance.MaxLvl, Npclist(tempIndex).Char.CharIndex, vbWhite)
                        Exit Sub
                    End If
                    
                    Call WriteSendMasteries(UserIndex, eSendMasteryType.ClassMasteries)
                    Call WriteSendMasteries(UserIndex, eSendMasteryType.CharacterMasteries)

                ElseIf Npclist(tempIndex).NPCtype = eNPCType.GuildMaster Then
                    If UserList(UserIndex).Guild.IdGuild = 0 Then
                        Call WriteShowGuildCreate(UserIndex)
                    Else
                        GuildIndex = UserList(UserIndex).Guild.GuildIndex

                        Call WriteShowGuildForm(UserIndex)
                    End If
                End If
            End If
            
            'Hay un obj?
            If MapData(Map, X, Y).ObjInfo.ObjIndex > 0 Then
                tempIndex = MapData(Map, X, Y).ObjInfo.ObjIndex
                
                .flags.TargetObj = tempIndex
                
                Select Case ObjData(tempIndex).ObjType
                    Case eOBJType.otPuertas 'Es una puerta
                        Call AccionParaPuerta(Map, X, Y, UserIndex)
                    Case eOBJType.otCarteles 'Es un cartel
                        Call AccionParaCartel(Map, X, Y, UserIndex)
                    Case eOBJType.otForos 'Foro
                        Call AccionParaForo(Map, X, Y, UserIndex)
                    Case eOBJType.otLeña    'Leña
                        If (tempIndex = ConstantesItems.FogataApagada Or tempIndex = ConstantesItems.RamitaElfica) And .flags.Muerto = 0 Then
                            If tempIndex = ConstantesItems.RamitaElfica And Not HasPassiveAssigned(UserIndex, ePassiveSpells.VitalRestoration) Then
                                If .Stats.UserPassives(ePassiveSpells.VitalRestoration).AllowedByClass Then
                                    Call WriteConsoleMsg(UserIndex, "Debes tener la habilidad pasiva restauración vital.", FontTypeNames.FONTTYPE_INFO)
                                Else
                                    Call WriteConsoleMsg(UserIndex, "Tu clase no te permite realizar fogatas élficas", FontTypeNames.FONTTYPE_INFO)
                                End If
                                Exit Sub
                            End If
                            Call AccionParaRamita(Map, X, Y, UserIndex)
                        End If
                    Case eOBJType.otFogata
                        Call UserRest(UserIndex, True)
                    Case eOBJType.otTrampa
                        Call AccionParaTrampa(Map, X, Y, UserIndex)
                End Select
            '>>>>>>>>>>>OBJETOS QUE OCUPAM MAS DE UN TILE<<<<<<<<<<<<<
            ElseIf MapData(Map, X + 1, Y).ObjInfo.ObjIndex > 0 Then
                tempIndex = MapData(Map, X + 1, Y).ObjInfo.ObjIndex
                .flags.TargetObj = tempIndex
                
                Select Case ObjData(tempIndex).ObjType
                    
                    Case eOBJType.otPuertas 'Es una puerta
                        Call AccionParaPuerta(Map, X + 1, Y, UserIndex)
                    
                End Select
            
            ElseIf MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex > 0 Then
                tempIndex = MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex
                .flags.TargetObj = tempIndex
        
                Select Case ObjData(tempIndex).ObjType
                    Case eOBJType.otPuertas 'Es una puerta
                        Call AccionParaPuerta(Map, X + 1, Y + 1, UserIndex)
                End Select
            
            ElseIf MapData(Map, X, Y + 1).ObjInfo.ObjIndex > 0 Then
                tempIndex = MapData(Map, X, Y + 1).ObjInfo.ObjIndex
                .flags.TargetObj = tempIndex
                
                Select Case ObjData(tempIndex).ObjType
                    Case eOBJType.otPuertas 'Es una puerta
                        Call AccionParaPuerta(Map, X, Y + 1, UserIndex)
                End Select
            End If
        End With
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Accion de Acciones.bas")
End Sub

Public Sub AccionParaForo(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 02/01/2010
'02/01/2010: ZaMa - Agrego foros faccionarios
'***************************************************
On Error GoTo ErrHandler
  

On Error Resume Next

    Dim Pos As WorldPos
    
    Pos.Map = Map
    Pos.X = X
    Pos.Y = Y
    
    If Distancia(Pos, UserList(UserIndex).Pos) > 2 Then
        Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If SendPosts(UserIndex, ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).ForoID) Then
        Call WriteShowForumForm(UserIndex)
    End If
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub AccionParaForo de Acciones.bas")
End Sub

Sub AccionParaPuerta(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)

On Error GoTo ErrHandler

    ' Too far away
    If Distance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, X, Y) > 2 Then
        Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    ' Door locked with a key?
    If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Llave = 1 Then
        Call WriteConsoleMsg(UserIndex, "La puerta esta cerrada con llave.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    Dim ObjIndex As Integer
    ObjIndex = MapData(Map, X, Y).ObjInfo.ObjIndex
    Dim GrhIndex As Long
    
    ' The new object to use is calculated based on the status of the current door
    ObjIndex = IIf(ObjData(ObjIndex).Cerrada = 1, ObjData(ObjIndex).IndexAbierta, ObjData(ObjIndex).IndexCerrada)
    
    ' Set the object
    MapData(Map, X, Y).ObjInfo.ObjIndex = ObjIndex
    MapData(Map, X, Y).ObjInfo.CurrentGrhIndex = ObjData(ObjIndex).GrhIndex

    ' Block or unblock the tile based on the new object status.
    MapData(Map, X, Y).Blocked = ObjData(ObjIndex).Cerrada
    MapData(Map, X - 1, Y).Blocked = ObjData(ObjIndex).Cerrada
        
    'Agregar / Remover block
    Call SendToItemArea(Map, X, Y, PrepareMessageObjectUpdate(X, Y, ObjData(ObjIndex).GrhIndex, ObjData(ObjIndex).ObjType, GetCreateObjectMetadata(ObjIndex, Map, X, Y)))
    Call SendToItemArea(Map, X, Y, PrepareMessagePlayWave(ConstantesSonidos.Puerta, X, Y))

    UserList(UserIndex).flags.TargetObj = MapData(Map, X, Y).ObjInfo.ObjIndex
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub AccionParaPuerta de Acciones.bas")
End Sub

Sub AccionParaCartel(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

On Error Resume Next

If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).ObjType = 8 Then
  
  If Len(ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).texto) > 0 Then
    Call WriteShowSignal(UserIndex, MapData(Map, X, Y).ObjInfo.ObjIndex)
  End If
  
End If

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub AccionParaCartel de Acciones.bas")
End Sub

Sub AccionParaRamita(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

On Error Resume Next

Dim Suerte As Byte
Dim exito As Byte

Dim SkillSupervivencia As Byte

Dim Pos As WorldPos
Pos.Map = Map
Pos.X = X
Pos.Y = Y

With UserList(UserIndex)
    If Distancia(Pos, .Pos) > 2 Then
        Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If MapData(Map, X, Y).Trigger = eTrigger.ZONASEGURA Or MapInfo(Map).Pk = False Then
        Call WriteConsoleMsg(UserIndex, "No puedes hacer fogatas en zona segura.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If HayAgua(Map, X, Y) Then
        Call WriteConsoleMsg(UserIndex, "No puedes encender una fogata en el agua.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If .flags.DueloIndex > 0 Then
        If Not DuelData.Duelo(.flags.DueloIndex).Resucitar Or DuelData.Duelo(.flags.DueloIndex).TipoDuelo = vs1 Then
            Call WriteConsoleMsg(UserIndex, "No puedes encender una fogata élfica durante un duelo en el que Resucitar no está permitido.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    End If

    SkillSupervivencia = GetSkills(UserIndex, eSkill.Supervivencia)
    
    If SkillSupervivencia < 6 Then
        Suerte = 3
        
    ElseIf SkillSupervivencia <= 10 Then
        Suerte = 2
        
    Else
        Suerte = 1
    End If
    
    exito = RandomNumber(1, Suerte)
    
    ' Failed
    If exito <> 1 Then
        Call WriteConsoleMsg(UserIndex, "No has podido hacer fuego.", FontTypeNames.FONTTYPE_INFO)
        Call SubirSkill(UserIndex, eSkill.Supervivencia, False)
        Exit Sub
    End If
    
    ' Cant deploy campfires in cities
    If MapInfo(.Pos.Map).Zona = eTerrainZone.zone_ciudad Then
        Call WriteConsoleMsg(UserIndex, "La ley impide realizar fogatas en las ciudades.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
                    
    Dim ObjIndex As Long
    ObjIndex = ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).CampfireObj

    MapData(Map, X, Y).ObjInfo.ObjIndex = ObjIndex
    MapData(Map, X, Y).ObjInfo.CurrentGrhIndex = ObjData(ObjIndex).GrhIndex

    ' Updates object
    Call SendToItemArea(Map, X, Y, PrepareMessageObjectUpdate(X, Y, ObjData(ObjIndex).GrhIndex, ObjData(ObjIndex).ObjType, GetCreateObjectMetadata(ObjIndex, Map, X, Y)))

    ' Notify object
    Call WriteConsoleMsg(UserIndex, "Has encendido una fogata.", FontTypeNames.FONTTYPE_INFO)
        
    ' Raise survival skill
    Call SubirSkill(UserIndex, eSkill.Supervivencia, True, ConstantesBalance.SkillExpCampfireSuccess)
    
    ' Send a message to the state server to make the campfire disappear
    Call SendAddCampfire(Map, X, Y, ObjData(ObjIndex).DisappearTimeInSec)




End With

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub AccionParaRamita de Acciones.bas")
End Sub

Sub AccionParaTrampa(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)

    Dim TrapPos  As WorldPos
    Dim ObjIndex As Integer
    
    On Error GoTo ErrHandler

    TrapPos.Map = Map
    TrapPos.X = X
    TrapPos.Y = Y
    
    If ConstantesBalance.MaxActiveTrapQty = 0 Then
        Call WriteConsoleMsg(UserIndex, "El servidor tiene las trampas desactivadas.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If Distancia(TrapPos, UserList(UserIndex).Pos) > 2 Then
        Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If UserList(UserIndex).clase <> eClass.Hunter Then
        Call WriteConsoleMsg(UserIndex, "Tu clase no puede activar ni desactivar trampas.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If UserList(UserIndex).flags.Muerto = 1 Then
        Call WriteConsoleMsg(UserIndex, "No puedes activar ni desactivar trampas estando muerto.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If MapInfo(UserList(UserIndex).Pos.Map).Zona = eTerrainZone.zone_ciudad Or MapData(Map, X, Y).Trigger = eTrigger.ZONASEGURA Or MapInfo(Map).Pk = False Then
        Call WriteConsoleMsg(UserIndex, "No se pueden activar trampas en ciudades o zonas seguras.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If HayAgua(Map, X, Y) Then
        Call WriteConsoleMsg(UserIndex, "No puedes activar una trampa en el agua.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If MapData(Map, X, Y).ObjInfo.Amount <> 1 Then
        Call WriteConsoleMsg(UserIndex, "Solo puedes activar 1 trampa a la vez.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
        
    ObjIndex = MapData(Map, X, Y).ObjInfo.ObjIndex
    
    ' If the object is not defined then do nothing
    If ObjIndex <= 0 Then
        Exit Sub
    End If

    With ObjData(ObjIndex)

        ' If the object is not defined then do nothing
        If .TrapActivatedObject <= 0 Then Exit Sub
   
        If .ObjType = eOBJType.otTrampa Then
           
            If .TrapActivable Then
                If (UserList(UserIndex).Stats.ELV < .TrapActivableLevelActivate) Then
                    Call WriteConsoleMsg(UserIndex, "Necesitas ser nivel " & .TrapActivableLevelActivate & " o superior para armar esta trampa..", FontTypeNames.FONTTYPE_INFOBOLD)
                    Exit Sub
                End If
            Else
                If (UserList(UserIndex).Stats.ELV < .TrapActivableLevelDeactivate) Then
                    Call WriteConsoleMsg(UserIndex, "Necesitas ser nivel : " & .TrapActivableLevelDeactivate & " o superior para desactivar esta trampa.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            End If
           
            MapData(Map, X, Y).ObjInfo.ObjIndex = .TrapActivatedObject
            MapData(Map, X, Y).ObjInfo.CurrentGrhIndex = ObjData(.TrapActivatedObject).GrhIndex
            MapData(Map, X, Y).ObjInfo.ActivatedByUser = UserList(UserIndex).ID

            ' Updates object
            Call SendToItemArea(Map, X, Y, PrepareMessageObjectUpdate(X, Y, ObjData(.TrapActivatedObject).GrhIndex, ObjData(.TrapActivatedObject).ObjType, GetCreateObjectMetadata(.TrapActivatedObject, Map, X, Y)))

            If .TrapActivable Then
                Call AddTrapToList(UserIndex, Map, X, Y)
                Call WriteConsoleMsg(UserIndex, "Activaste una trampa.", FontTypeNames.FONTTYPE_INFOBOLD)
            Else
                Call DelTrapFromList(UserIndex, Map, X, Y)
                Call WriteConsoleMsg(UserIndex, "Desactivaste una trampa.", FontTypeNames.FONTTYPE_INFOBOLD)
            End If
        End If
        
    End With
    
    Exit Sub
  
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub AccionParaTrampa de Acciones.bas")

End Sub

Public Sub UserRest(ByVal UserIndex As Integer, Optional ByVal FromClick As Boolean = False)
'***************************************************
'Author: Lex
'Last Modification: 05/04/12
'
'***************************************************
On Error GoTo ErrHandler

    With UserList(UserIndex)
        
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Solo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
                    
        If Not .flags.Descansar Then
            Dim FoundPoss As WorldPos
                        
            If FromClick Then
                'if Distancia(.Pos, .selected
                Dim clickedPos As WorldPos
                clickedPos.Map = .flags.TargetMap
                clickedPos.X = .flags.TargetX
                clickedPos.Y = .flags.TargetY
                
                If Distancia(.Pos, clickedPos) > 2 Then
                    Call WriteConsoleMsg(UserIndex, "La fogata se encuentra muy lejos.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                FoundPoss = clickedPos
                
            Else
                Dim ListObject() As tSpellPosition, LastIndex As Integer, I As Integer
                Dim AllowRest As Boolean
                
                Call ObtainListObjectNearPlayer(.Pos, 2, ListObject, UserIndex)
                
                For I = 1 To VectorWorldPosSize(ListObject)
                    If MapData(ListObject(I).Pos.Map, ListObject(I).Pos.X, ListObject(I).Pos.Y).ObjInfo.ObjIndex > 0 Then
                        If ObjData(MapData(ListObject(I).Pos.Map, ListObject(I).Pos.X, ListObject(I).Pos.Y).ObjInfo.ObjIndex).AllowResting = True Then
                            FoundPoss = ListObject(I).Pos
                            AllowRest = True
                            Exit For
                        End If
                    End If
                Next I
                
                If Not AllowRest = True Then
                    Call WriteConsoleMsg(UserIndex, "No hay ninguna fogata junto a la cual descansar.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            End If
            

            
            .RestObjectCoords = FoundPoss
            
            Call WriteConsoleMsg(UserIndex, "Te acomodás junto a la fogata y comienzas a descansar.", FontTypeNames.FONTTYPE_INFO)
        Else
        
            Call WriteConsoleMsg(UserIndex, "Te levantas.", FontTypeNames.FONTTYPE_INFO)
            .RestObjectCoords.Map = 0
            .RestObjectCoords.X = 0
            .RestObjectCoords.Y = 0
        End If
            
        Call WriteRestOK(UserIndex)
        .flags.Descansar = Not .flags.Descansar
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub UserRest de Acciones.bas")
End Sub
