Attribute VB_Name = "modTriggers"
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

Public Sub CheckTriggerActivation(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, _
    ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal FlushAfterCheck As Boolean)
'***************************************************
'Autor: ZaMa
'Last Modification: 20/10/2012
'Cheks and activate triggers.
'***************************************************

' Pos has trigger Item?
Dim ObjIndex As Integer
Dim TrapActivedBy As Long
Dim OriginIndex As Integer
Dim UserName As String
Dim InstanceId As Long
    
On Error GoTo ErrHandler

    If UserIndex <> 0 Then
        ' Doesn't affect dead users
        If UserList(UserIndex).flags.Muerto = 1 Then Exit Sub
        
        ' It does not affect immune users
        If UserList(UserIndex).flags.Inmunidad = 1 Then Exit Sub
        
        ' Doesn't affect Admins
        If EsGm(UserIndex) Then Exit Sub
    End If
  
    ObjIndex = MapData(Map, X, Y).ObjInfo.ObjIndex
    TrapActivedBy = MapData(Map, X, Y).ObjInfo.ActivatedByUser
    
    ' Check if the object configured is a valid obj index
    If ObjIndex < 1 Or ObjIndex > NumObjDatas Then Exit Sub
    
    ' Check if the object type is a trigger or a trap ('cause both can be activated as traps).
    If ObjData(ObjIndex).ObjType <> eOBJType.otTrigger And ObjData(ObjIndex).ObjType <> eOBJType.otTrampa Then Exit Sub
    
    With ObjData(ObjIndex).Trigger
        ' Affects npc?
        If NpcIndex <> 0 Then
            If .AffectNpc = False Then Exit Sub
        
        ' Special user message?
        Else
            ' Affects users?
            If .AffectUser = False Then Exit Sub
            
            If LenB(.ActivationMessage) <> 0 Then
                Call WriteConsoleMsg(UserIndex, .ActivationMessage, FontTypeNames.FONTTYPE_VENENO)
            End If
        End If
        
         OriginIndex = GetUserIndexFromUserId(TrapActivedBy)
        ' Check if origin user is online
        If OriginIndex > 0 Then
            UserName = UserList(OriginIndex).Name
            InstanceId = UserList(OriginIndex).InstanceId
        Else
            ' The owner of the trap is disconnected.
            ' As this should never happen in a happy path, let's destroy the trap and end here.
            Call EraseObj(MapData(Map, X, Y).ObjInfo.Amount, Map, X, Y)
            Call DelTrapFromList(OriginIndex, Map, X, Y)
            Exit Sub
        End If
        
        If UserIndex > 0 Then
        
        If FriendlyFireProtectionEnabled(UserIndex, OriginIndex) And MapData(Map, X, Y).Trigger <> eTrigger.ZONAPELEA Then Exit Sub
        End If
        
        Dim Died As Boolean
        
        ' Cast spells
        Dim SpellIndex As Long
        Dim IsImmuneToSpell As Boolean
        
        For SpellIndex = 1 To .NumSpells
                
            With .Spells(SpellIndex)
                IsImmuneToSpell = False
                
                If .Index <> 0 Then
                    ' Si el hechizo es de un solo uso, o si es de Damage Over Time
                    ' pero no hay que esperar al primer tick, entonces casteamos el hechizo
                    If Not Hechizos(.Index).DamageOverTime.IsDot Or (Hechizos(.Index).DamageOverTime.IsDot And Not Hechizos(.Index).DamageOverTime.WaitForFirstTick) Then
                        ' Sound & fx
                        Call SendSpellEffects(UserIndex, NpcIndex, .WAV, .FXgrh, .Loops, X, Y)
                    
                        ' Chars effects
                        If UserIndex <> 0 Then
                            IsImmuneToSpell = modMasteries.IsUserImmuneToSpell(TrapActivedBy, UserIndex, .Index, False)
                            
                            If Not IsImmuneToSpell Then Call CastSpellUser(UserIndex, .Index, .Interval, .MaxHit, .MinHit, Died, OriginIndex)
                        Else
                            Call CastSpellNpc(NpcIndex, .Index, .Interval, .MaxHit, .MinHit, Died, OriginIndex)
                        End If
                    End If
                    
                End If
                
                ' Invoke Npcs
                If .InvokeNpcIndex <> 0 Then
                    Dim tmpPos As WorldPos
                    tmpPos.Map = Map
                    tmpPos.X = X
                    tmpPos.Y = Y
                    Call SpawnNpc(.InvokeNpcIndex, tmpPos, True, False)
                End If
                If UserIndex <> 0 Then
                    ' Damage over time
                    If UserList(UserIndex).flags.Muerto = 0 And .DamageOverTime.IsDot And Not IsImmuneToSpell Then
                    
                        ' If the tick count or tick interval is 0 or less, exit
                        If Hechizos(.Index).DamageOverTime.TickCount <= 0 Or Hechizos(.Index).DamageOverTime.TickInterval <= 0 Then Exit Sub
                        
                        ' Send the message to the state server so we can be notified when a damage over time tick ocurred.
                        Call SendDamageOverTimeMessage(eTypeTarget.isUser, TrapActivedBy, InstanceId, UserName, eTypeTarget.isUser, UserList(UserIndex).Id, UserList(UserIndex).InstanceId, UserList(UserIndex).Name, .DamageOverTime.TickCount, .DamageOverTime.TickInterval, .Index, Hechizos(.Index).DamageOverTime.MaxStackEffect)
                    End If
                Else
                    ' Damage over time Npc
                    If Npclist(NpcIndex).flags.NPCActive And .DamageOverTime.IsDot Then
                    
                        ' If the tick count or tick interval is 0 or less, exit
                        If Hechizos(.Index).DamageOverTime.TickCount <= 0 Or Hechizos(.Index).DamageOverTime.TickInterval <= 0 Then Exit Sub
                        
                        ' marks the npc that it has an active dot to remove it from the stateServer at the time of its death
                        Npclist(NpcIndex).flags.isAffectedByDOT = True
                            
                        ' Send the message to the state server so we can be notified when a damage over time tick ocurred.
                        Call SendDamageOverTimeMessage(eTypeTarget.isUser, TrapActivedBy, InstanceId, UserName, eTypeTarget.IsNPC, NpcIndex, Npclist(NpcIndex).InstanceId, Npclist(NpcIndex).Name, .DamageOverTime.TickCount, .DamageOverTime.TickInterval, .Index, Hechizos(.Index).DamageOverTime.MaxStackEffect)
                    End If
                End If
                
            End With
                
            ' No need to keep casting..
            If Died Then Exit For

        Next SpellIndex
        
        ' Dissapears after activating?
        If .Dissapears = 1 Then
            Call EraseObj(MapData(Map, X, Y).ObjInfo.Amount, Map, X, Y)
            Call DelTrapFromList(OriginIndex, Map, X, Y)
        End If

    End With

  Exit Sub

ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CheckTriggerActivation de modTriggers.bas.")
  Call LogError("--> (UserIndex, NpcIndex, Map, X, Y, ObjIndex, TrapActivatedBy, OriginIndex, UserName, InstanceId) " & "(" & UserIndex & "," & NpcIndex & "," & Map & "," & X & "," & Y & "," & ObjIndex & "," & TrapActivedBy & "," & OriginIndex & "," & UserName & "," & InstanceId & ")")

End Sub

Public Sub CastSpellNpc(ByVal NpcIndex As Integer, _
    ByVal SpellIndex As Integer, ByVal SpellInterval As Long, _
    ByVal MinHit As Integer, ByVal MaxHit As Integer, ByRef NpcDied As Boolean, ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 20/10/2012
'Casts spell on npc.
'***************************************************
On Error GoTo ErrHandler
  
    
    Dim Damage As Integer
    Dim Accion As String
    
    With Npclist(NpcIndex)
    
        Dim MaxHp As Integer
        
        If Hechizos(SpellIndex).Tipo = uPropiedades Then
            ' Spell deals damage??
            If Hechizos(SpellIndex).SubeHP = 2 Then
                
                If .flags.Snd2 > 0 Then
                    Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(Npclist(NpcIndex).flags.Snd2, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y, Npclist(NpcIndex).Char.CharIndex))
                End If
            
                Damage = RandomNumber(Hechizos(SpellIndex).MinHp, Hechizos(SpellIndex).MaxHp)
                
                ' Deal damage
                .Stats.MinHp = .Stats.MinHp - Damage
                
                If UserIndex > 0 Then 'This check is because if the user is offline it throws an error
                    Call WriteConsoleMsg(UserIndex, "Has quitado " & Damage & " puntos de vida a la criatura.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                End If
                
                Call CalcularDarExp(UserIndex, NpcIndex, Damage)
                
                ' Muere?
                If .Stats.MinHp < 1 Then
                    .Stats.MinHp = 0
                
                    ' Pet?
                    If .MaestroUser > 0 Then
                        Call MuereNpc(NpcIndex, .MaestroUser)
                    Else
                        Call MuereNpc(NpcIndex, UserIndex)
                    End If
                    
                    NpcDied = True
                    
                End If
            ' Spell recovers health??
            ElseIf Hechizos(SpellIndex).SubeHP = 1 Then
                
                Damage = RandomNumber(MinHit, MaxHit)
            
                ' Recovers health
                .Stats.MinHp = .Stats.MinHp + Damage
                
                If .Stats.MinHp > .Stats.MaxHp Then
                    .Stats.MinHp = .Stats.MaxHp
                End If
    
                If UserIndex > 0 Then 'This check is because if the user is offline it throws an error
                    Call WriteConsoleMsg(UserIndex, "Has restaurado " & Damage & " puntos de vida a la criatura.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                End If
            
            End If
        
        Else
            
            ' Spell Adds/Removes poison?
            If Hechizos(SpellIndex).Envenena = 1 Then
                .flags.Envenenado = 1
            ElseIf Hechizos(SpellIndex).CuraVeneno = 1 Then
                .flags.Envenenado = 0
            End If
    
            ' Spells Adds/Removes Paralisis/Inmobility?
            If Hechizos(SpellIndex).Paraliza = 1 Then
                .flags.Paralizado = 1
                .flags.Inmovilizado = 0
                
                If SpellInterval <> 0 Then
                    .Contadores.Paralisis = SetIntervalEnd(SpellInterval)
                Else
                    .Contadores.Paralisis = SetIntervalEnd(ServerConfiguration.Intervals.IntervaloParalizado)
                End If
                
            ElseIf Hechizos(SpellIndex).Inmoviliza = 1 Then
                If .flags.AfectaParalisis = 0 Then
                
                    .flags.Inmovilizado = 1
                    .flags.Paralizado = 0
                    
                    If SpellInterval <> 0 Then
                        .Contadores.Paralisis = SetIntervalEnd(SpellInterval)
                    Else
                        .Contadores.Paralisis = SetIntervalEnd(ServerConfiguration.Intervals.IntervaloParalizado)
                    End If
                Else
                    Call WriteConsoleMsg(UserIndex, "El NPC es inmune a este hechizo.", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)
                End If
                
            ElseIf Hechizos(SpellIndex).RemoverParalisis = 1 Then
                If .flags.Paralizado = 1 Or .flags.Inmovilizado = 1 Then
                    .flags.Paralizado = 0
                    .flags.Inmovilizado = 0
                    .Contadores.Paralisis = 0
                End If
            End If
        
        End If
    
    End With
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CastSpellNpc de modTriggers.bas")
End Sub

Public Sub CastSpellUser(ByVal UserIndex As Integer, _
    ByVal SpellIndex As Integer, ByVal SpellInterval As Long, _
    ByVal MinHit As Integer, ByVal MaxHit As Integer, _
    ByRef UserDied As Boolean, ByVal OriginIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 20/10/2012
'Casts spell on user.
'***************************************************
On Error GoTo ErrHandler
  
    
    Dim Damage As Integer
    Dim Accion As String
    Dim ReduceMagicDamage As Integer
    
    ReduceMagicDamage = 0
    
    With UserList(UserIndex)
        
        Dim AnilloObjIndex As Integer
        AnilloObjIndex = .Invent.AnilloEqpObjIndex
        
        If Hechizos(SpellIndex).Tipo = uPropiedades Then
        
            ' Health
            If Hechizos(SpellIndex).SubeHP <> 0 Then
                   
                Damage = RandomNumber(Hechizos(SpellIndex).MinHp, Hechizos(SpellIndex).MaxHp)
                
                ' Health (+)
                If Hechizos(SpellIndex).SubeHP = 1 Then
                    .Stats.MinHp = .Stats.MinHp + Damage
                    If .Stats.MinHp > .Stats.MaxHp Then _
                        .Stats.MinHp = .Stats.MaxHp
                    
                    Accion = "restaurado"
                
                ' Health (-)
                ElseIf Hechizos(SpellIndex).SubeHP = 2 Then
                    
                    'cascos antimagia
                    If (.Invent.CascoEqpObjIndex > 0) Then
                       ReduceMagicDamage = ReduceMagicDamage + RandomNumber(ObjData(.Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(.Invent.CascoEqpObjIndex).DefensaMagicaMax)
                    End If
                    
                    'anillos
                    If (.Invent.AnilloEqpObjIndex > 0) Then
                        ReduceMagicDamage = ReduceMagicDamage + RandomNumber(ObjData(.Invent.AnilloEqpObjIndex).DefensaMagicaMin, ObjData(.Invent.AnilloEqpObjIndex).DefensaMagicaMax)
                    End If
                    
                    If .flags.Navegando = 1 And .Invent.BarcoObjIndex > 0 Then
                        ReduceMagicDamage = ReduceMagicDamage + RandomNumber(ObjData(.Invent.BarcoObjIndex).DefensaMagicaMin, ObjData(.Invent.BarcoObjIndex).DefensaMagicaMax)
                    End If
                    
                    ' Ignore a percentage of the magic reduction based on a value configured in the spell (mainly used by traps)
                    ReduceMagicDamage = ReduceMagicDamage - Fix(Porcentaje(ReduceMagicDamage, Hechizos(SpellIndex).IgnoreMagicDefensePerc))
                    
                    ' Maximum magic reduction should be equal to damage
                    If ReduceMagicDamage > Damage Then ReduceMagicDamage = Damage
                    
                    Damage = Max(1, Damage - ReduceMagicDamage)
                
                    .Stats.MinHp = .Stats.MinHp - Damage
                    
                    If .Stats.MinHp < 0 Then .Stats.MinHp = 0
                    
                    Accion = "quitado"
                End If
                
                Call WriteConsoleMsg(UserIndex, UserList(OriginIndex).Name + " te ha " & Accion & " " & Damage & IIf(ReduceMagicDamage > 0, " (" & ReduceMagicDamage & " resistido) ", "") & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                Call WriteConsoleMsg(OriginIndex, "Has " & Accion & " " & Damage & IIf(ReduceMagicDamage > 0, " (" & ReduceMagicDamage & " resistido) ", "") & " puntos de vida a " + .Name, FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                                
                Call WriteUpdateHP(UserIndex)
                
                ' Die?
                If .Stats.MinHp < 1 Then
                    .Stats.MinHp = 0
                    If .flags.AtacablePor <> OriginIndex Then
                        'Store it!
                        Call Statistics.StoreFrag(OriginIndex, UserIndex)
                        Call ContarMuerte(UserIndex, OriginIndex, eDamageType.Spell, Damage, SpellIndex)
                    End If
                    Call ActStats(UserIndex, OriginIndex)
                    Call UserDie(UserIndex)
                    UserDied = True
                    Exit Sub
                End If
            End If
                 
            ' Hunger
            If Hechizos(SpellIndex).SubeHam <> 0 Then
                
                Damage = RandomNumber(Hechizos(SpellIndex).MinHam, Hechizos(SpellIndex).MaxHam)
                
                ' Hunger (+)
                If Hechizos(SpellIndex).SubeHam = 1 Then
                
                    .Stats.MinHam = .Stats.MinHam + Damage
                    If .Stats.MinHam > .Stats.MaxHam Then .Stats.MinHam = .Stats.MaxHam
                    
                    Accion = "restaurado"
                    
                ' Hunger (-)
                ElseIf Hechizos(SpellIndex).SubeHam = 2 Then
                
                    .Stats.MinHam = .Stats.MinHam - Damage
                    If .Stats.MinHam < 1 Then
                        .Stats.MinHam = 0
                        .flags.Hambre = 1
                    End If
                    
                    Accion = "quitado"
                End If
                
                Call WriteConsoleMsg(UserIndex, "Te ha " & Accion & " " & Damage & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                
                Call WriteUpdateHungerAndThirst(UserIndex)
            End If
            
            ' Thirst
            If Hechizos(SpellIndex).SubeSed <> 0 Then
                
                Damage = RandomNumber(Hechizos(SpellIndex).MinSed, Hechizos(SpellIndex).MaxSed)
                
                ' Thirst (+)
                If Hechizos(SpellIndex).SubeSed = 1 Then
                    
                    .Stats.MinAGU = .Stats.MinAGU + Damage
                    If .Stats.MinAGU > .Stats.MaxAGU Then .Stats.MinAGU = .Stats.MaxAGU
                
                    Accion = "restaurado"
                    
                ' Thirst (-)
                ElseIf Hechizos(SpellIndex).SubeSed = 2 Then
                    
                    .Stats.MinAGU = .Stats.MinAGU - Damage
                    
                    If .Stats.MinAGU < 1 Then
                        .Stats.MinAGU = 0
                        .flags.Sed = 1
                    End If
                    
                    Accion = "quitado"
                End If
                
                Call WriteConsoleMsg(UserIndex, "Te ha " & Accion & " " & Damage & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                
                Call WriteUpdateHungerAndThirst(UserIndex)
            End If
            
            ' Dexerity
            If Hechizos(SpellIndex).SubeAgilidad <> 0 Then
                
                Damage = RandomNumber(Hechizos(SpellIndex).MinAgilidad, Hechizos(SpellIndex).MaxAgilidad)
                
                ' Dexerity (+)
                If Hechizos(SpellIndex).SubeAgilidad = 1 Then
                
                    .flags.DuracionEfecto = 1200
                    .Stats.UserAtributos(eAtributos.Agilidad) = .Stats.UserAtributos(eAtributos.Agilidad) + Damage
                    If .Stats.UserAtributos(eAtributos.Agilidad) > MinimoInt(ConstantesBalance.MaxAtributos, .Stats.UserAtributosBackUP(Agilidad) * 2) Then _
                        .Stats.UserAtributos(eAtributos.Agilidad) = MinimoInt(ConstantesBalance.MaxAtributos, .Stats.UserAtributosBackUP(Agilidad) * 2)
                    
                    Accion = "aumentado"
                    
                ' Dexerity (-)
                ElseIf Hechizos(SpellIndex).SubeAgilidad = 2 Then
                
                    .flags.DuracionEfecto = 700
                    .Stats.UserAtributos(eAtributos.Agilidad) = .Stats.UserAtributos(eAtributos.Agilidad) - Damage
                    If .Stats.UserAtributos(eAtributos.Agilidad) < ConstantesBalance.MinAtributos Then .Stats.UserAtributos(eAtributos.Agilidad) = ConstantesBalance.MinAtributos
                    
                    Accion = "disminuído"
                End If
                
                Call WriteConsoleMsg(UserIndex, "Te ha " & Accion & " " & Damage & " puntos de agilidad.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                
                Call WriteUpdateDexterity(UserIndex)
            End If
            
            ' Strenght
            If Hechizos(SpellIndex).SubeFuerza <> 0 Then
                
                Damage = RandomNumber(Hechizos(SpellIndex).MinFuerza, Hechizos(SpellIndex).MaxFuerza)
                
                ' Strenght (+)
                If Hechizos(SpellIndex).SubeFuerza = 1 Then
                    .flags.DuracionEfecto = 1200
                
                    .Stats.UserAtributos(eAtributos.Fuerza) = .Stats.UserAtributos(eAtributos.Fuerza) + Damage
                    If .Stats.UserAtributos(eAtributos.Fuerza) > MinimoInt(ConstantesBalance.MaxAtributos, .Stats.UserAtributosBackUP(Fuerza) * 2) Then _
                        .Stats.UserAtributos(eAtributos.Fuerza) = MinimoInt(ConstantesBalance.MaxAtributos, .Stats.UserAtributosBackUP(Fuerza) * 2)
                    
                    Accion = "aumentado"
                    
                ' Strenght (-)
                ElseIf Hechizos(SpellIndex).SubeFuerza = 2 Then
                
                    .flags.DuracionEfecto = 700
                    .Stats.UserAtributos(eAtributos.Fuerza) = .Stats.UserAtributos(eAtributos.Fuerza) - Damage
                    If .Stats.UserAtributos(eAtributos.Fuerza) < ConstantesBalance.MinAtributos Then .Stats.UserAtributos(eAtributos.Fuerza) = ConstantesBalance.MinAtributos
                
                    Accion = "disminuído"
                End If
                
                Call WriteConsoleMsg(UserIndex, "Te ha " & Accion & " " & Damage & " puntos de fuerza.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                
                Call WriteUpdateStrenght(UserIndex)
            End If
            
            ' Mana
            If Hechizos(SpellIndex).SubeMana <> 0 Then
                
                Damage = RandomNumber(Hechizos(SpellIndex).MinMana, Hechizos(SpellIndex).MaxMana)
                
                ' Mana (+)
                If Hechizos(SpellIndex).SubeMana = 1 Then
                    .Stats.MinMAN = .Stats.MinMAN + Damage
                    If .Stats.MinMAN > .Stats.MaxMan Then .Stats.MinMAN = .Stats.MaxMan
                    
                    Accion = "restaurado"
                    
                ' Mana (-)
                ElseIf Hechizos(SpellIndex).SubeMana = 2 Then
                    
                    .Stats.MinMAN = .Stats.MinMAN - Damage
                    If .Stats.MinMAN < 1 Then .Stats.MinMAN = 0
                    
                    Accion = "quitado"
                End If
                
                Call WriteConsoleMsg(UserIndex, "Te ha " & Accion & " " & Damage & " puntos de maná.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                
                Call WriteUpdateMana(UserIndex)
            End If
            
            ' Stamina
            If Hechizos(SpellIndex).SubeSta = 1 Then
                
                ' Stamina (+)
                If Hechizos(SpellIndex).SubeSta = 1 Then
                    .Stats.MinSta = .Stats.MinSta + Damage
                    If .Stats.MinSta > .Stats.MaxSta Then .Stats.MinSta = .Stats.MaxSta
                    
                    Accion = "restaurado"
                    
                ' Stamina (-)
                ElseIf Hechizos(SpellIndex).SubeSta = 1 Then
                    .Stats.MinSta = .Stats.MinSta - Damage
                    If .Stats.MinSta < 1 Then .Stats.MinSta = 0
                    
                    Accion = "quitado"
                End If
                
                Call WriteUpdateSta(UserIndex)
                
                Call WriteConsoleMsg(UserIndex, "Te ha " & Accion & " " & Damage & " puntos de maná.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
            End If
            
        ElseIf Hechizos(SpellIndex).Tipo = TipoHechizo.uEstado Then
            
            ' Invisibility (Add)
            If Hechizos(SpellIndex).Invisibilidad = 1 Then
               
                UserList(UserIndex).flags.invisible = 1
                
                ' Solo se hace invi para los clientes si no esta navegando
                If UserList(UserIndex).flags.Navegando = 0 Then
                    Call UsUaRiOs.SetInvisible(UserIndex, UserList(UserIndex).Char.CharIndex, True)
                End If
            End If
            
            ' Poison (Add)
            If Hechizos(SpellIndex).Envenena = 1 Then
                UserList(UserIndex).flags.Envenenado = 1
                
            ' Poison (Remove)
            ElseIf Hechizos(SpellIndex).CuraVeneno = 1 Then
                UserList(UserIndex).flags.Envenenado = 0
            End If
            
            ' Curse (Add)
            If Hechizos(SpellIndex).Maldicion = 1 Then
                UserList(UserIndex).flags.Maldicion = 1
            
            ' Curse (Remove)
            ElseIf Hechizos(SpellIndex).RemoverMaldicion = 1 Then
                UserList(UserIndex).flags.Maldicion = 0
            End If
            
            ' Blessing (Add)
            If Hechizos(SpellIndex).Bendicion = 1 Then
                UserList(UserIndex).flags.Bendicion = 1
            End If
            
            Dim Rechaza As Boolean
            
            ' Paralysis/Inmobility (Add)
            If Hechizos(SpellIndex).Paraliza = 1 Or Hechizos(SpellIndex).Inmoviliza = 1 Then
                
                 If UserList(UserIndex).flags.Paralizado = 0 Then
                    
                    If AnilloObjIndex > 0 Then
                        If ObjData(AnilloObjIndex).ImpideParalizar Then
                            Call WriteConsoleMsg(UserIndex, " ¡El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                            Rechaza = True
                        End If
                    End If
                    
                    If Not Rechaza Then
                        UserList(UserIndex).flags.Paralizado = 1
                        If Hechizos(SpellIndex).Inmoviliza = 1 Then UserList(UserIndex).flags.Inmovilizado = 1
                        
                        
                        If SpellInterval <> 0 Then
                            UserList(UserIndex).Counters.Paralisis = SetIntervalEnd(SpellInterval)
                        Else
                            UserList(UserIndex).Counters.Paralisis = SetIntervalEnd(ServerConfiguration.Intervals.IntervaloParalizado)
                        End If
                        
                        Call WriteParalizeOK(UserIndex)
                    End If
                End If
            
            ' Paralysis/Inmobility (Remove)
            ElseIf Hechizos(SpellIndex).RemoverParalisis = 1 Then
                
                ' Remueve si esta en ese estado
                If UserList(UserIndex).flags.Paralizado = 1 Then
                    Call RemoveParalisis(UserIndex)
                End If
            End If
            
            ' Confusion (Add)
            If Hechizos(SpellIndex).Estupidez = 1 Then
                
                If UserList(UserIndex).flags.Estupidez = 0 Then
                    UserList(UserIndex).flags.Estupidez = 1
                    
                    If SpellInterval <> 0 Then
                        UserList(UserIndex).Counters.Ceguera = SetIntervalEnd(SpellInterval)
                    Else
                        UserList(UserIndex).Counters.Ceguera = SetIntervalEnd(ServerConfiguration.Intervals.IntervaloParalizado)
                    End If
                End If
                
                Call WriteDumb(UserIndex)
            
            ' Confusion (Remove)
            ElseIf Hechizos(SpellIndex).RemoverEstupidez = 1 Then
            
                ' Remueve si esta en ese estado
                If UserList(UserIndex).flags.Estupidez = 1 Then
                
                    UserList(UserIndex).flags.Estupidez = 0
                    'no need to crypt this
                    Call WriteDumbNoMore(UserIndex)
                
                End If
            End If
            
            ' Revive
            If Hechizos(SpellIndex).Revivir = eReviveTarget.User Then
                If UserList(UserIndex).flags.Muerto = 1 Then
                    Call RevivirUsuario(UserIndex, True)
                End If
            End If
            
            ' Blind (Add)
            If Hechizos(SpellIndex).Ceguera = 1 Then
                
                If UserList(UserIndex).flags.Ceguera = 0 Then
                    UserList(UserIndex).flags.Ceguera = 1
                    
                    If SpellInterval <> 0 Then
                        UserList(UserIndex).Counters.Ceguera = SetIntervalEnd(SpellInterval)
                    Else
                        UserList(UserIndex).Counters.Ceguera = SetIntervalEnd(ServerConfiguration.Intervals.IntervaloParalizado / 3)
                    End If
            
                    Call WriteBlind(UserIndex)
                End If
                
            ' Blind (Remove)
            ElseIf Hechizos(SpellIndex).Ceguera = 2 Then
                
                If UserList(UserIndex).flags.Ceguera = 1 Then
                    UserList(UserIndex).flags.Ceguera = 0
                    UserList(UserIndex).Counters.Ceguera = 0
            
                    Call WriteBlindNoMore(UserIndex)
                End If
            End If
        End If
    End With

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CastSpellUser de modTriggers.bas")
End Sub

Public Sub SendSpellEffects(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, _
    ByVal iWav As Integer, ByVal iFX As Integer, ByVal Loops As Byte, ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: ZaMa
'Last Modification: 20/10/2012
'Sends spell's wav and fx to users.
'***************************************************
On Error GoTo ErrHandler
  

    Dim FinalTarget As SendTarget
    Dim TargetIndex As Integer
    Dim CharIndex As Integer
    
    If UserIndex <> 0 Then
        TargetIndex = UserIndex
        FinalTarget = SendTarget.ToPCArea
        CharIndex = UserList(UserIndex).Char.CharIndex
    Else
        TargetIndex = NpcIndex
        FinalTarget = SendTarget.ToNPCArea
        CharIndex = Npclist(NpcIndex).Char.CharIndex
    End If

    ' Spell Wav
    If iWav <> 0 Then
        Call SendData(FinalTarget, TargetIndex, _
            PrepareMessagePlayWave(iWav, X, Y, CharIndex))
    End If
    
    ' Spell FX
    If iFX > 0 Then
        Call SendData(FinalTarget, TargetIndex, _
            PrepareMessageCreateFX(CharIndex, iFX, Loops))
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SendSpellEffects de modTriggers.bas")
End Sub
