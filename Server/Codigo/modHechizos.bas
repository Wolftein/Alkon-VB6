Attribute VB_Name = "modHechizos"
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

Public Enum eTargetType
    ieUser = 1
    ieNpc
    ieObject
End Enum

Sub NpcLanzaSpellSobreUser(ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByVal Spell As Integer, _
                           Optional ByVal DecirPalabras As Boolean = False, _
                           Optional ByVal IgnoreVisibilityCheck As Boolean = False, _
                           Optional ByVal IgnoreDistanceCheck As Boolean = False)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 11/11/2010
'13/02/2009: ZaMa - Los npcs que tiren magias, no podran hacerlo en mapas donde no se permita usarla.
'13/07/2010: ZaMa - Ahora no se contabiliza la muerte de un atacable.
'21/09/2010: ZaMa - Amplio los tipos de hechizos que pueden lanzar los npcs.
'21/09/2010: ZaMa - Permito que se ignore el chequeo de visibilidad (pueden atacar a invis u ocultos).
'11/11/2010: ZaMa - No se envian los efectos del hechizo si no lo castea.
'05/09/2016: Anagrama - Agregados muchos tipos de hechizos diseñados para los mini-bosses.
'***************************************************
On Error GoTo ErrHandler
  
    Dim nPos As WorldPos
    Dim UI As Integer
    Dim UserProtected As Boolean
    Dim tmpByte As Integer
    
    With UserList(UserIndex)
    
        ' Doesn't consider if the user is hidden/invisible or not.
        If Not IgnoreVisibilityCheck Then
            If UserList(UserIndex).flags.invisible = 1 Or UserList(UserIndex).flags.Oculto = 1 Then Exit Sub
        End If
        
        If Not IgnoreDistanceCheck Then
            If Not InRangoVision(UserIndex, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y) Then Exit Sub
        End If
        
        ' Si no se peude usar magia en el mapa, no le deja hacerlo.
        If MapInfo(UserList(UserIndex).Pos.Map).MagiaSinEfecto > 0 Then Exit Sub
        
        ' Check if the user is immune to the given spell
        If modMasteries.IsUserImmuneToSpell(0, UserIndex, Spell, False) Then Exit Sub

        Dim Damage As Integer
        Dim AnilloObjIndex As Integer
        AnilloObjIndex = .Invent.AnilloEqpObjIndex
        
        'Atrae al usuario
        If Hechizos(Spell).Atraer = 1 Then
            Dim NpcPos As WorldPos
            NpcPos = Npclist(NpcIndex).Pos
            
            Call SendSpellEffects(UserIndex, NpcIndex, Spell, DecirPalabras)
            Call WriteConsoleMsg(UserIndex, Npclist(NpcIndex).Name & " te ha atraido hacia el.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
            
            If Npclist(NpcIndex).Pos.X < .Pos.X Then
                NpcPos.X = NpcPos.X + 1
            ElseIf Npclist(NpcIndex).Pos.X > .Pos.X Then
                NpcPos.X = NpcPos.X - 1
            ElseIf Npclist(NpcIndex).Pos.Y < .Pos.Y Then
                NpcPos.Y = NpcPos.Y + 1
            ElseIf Npclist(NpcIndex).Pos.Y > .Pos.Y Then
                NpcPos.Y = NpcPos.Y - 1
            End If
            
            Call ClosestLegalPos(NpcPos, nPos, True)
            
            If nPos.X <> 0 And nPos.Y <> 0 Then _
                Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, False, False, True)
        End If
        
        'Baja fuerza
        If Hechizos(Spell).SubeFuerza = 2 Then
            Damage = RandomNumber(Hechizos(Spell).MinFuerza, Hechizos(Spell).MaxFuerza)
            tmpByte = .Stats.UserAtributos(eAtributos.Fuerza)
            tmpByte = tmpByte - Damage
            If tmpByte < ConstantesBalance.MinAtributos Then tmpByte = ConstantesBalance.MinAtributos
            .Stats.UserAtributos(eAtributos.Fuerza) = tmpByte

            .flags.TomoPocion = True
            .flags.DuracionEfecto = 700
           
            Call WriteUpdateStrenght(UserIndex)
            
            ' Disable the Berserk
            If HasPassiveAssigned(UserIndex, ePassiveSpells.Berserk) Then
                If HasPassiveActivated(UserIndex, ePassiveSpells.Berserk) Then
                    Call ActivatePassive(UserIndex, ePassiveSpells.Berserk, False)
                    Call SendBerserkEffect(UserIndex, ePassiveSpells.Berserk, False)
                End If
            End If

            Call SendSpellEffects(UserIndex, NpcIndex, Spell, DecirPalabras)

            Call WriteConsoleMsg(UserIndex, Npclist(NpcIndex).Name & " te ha quitado " & Damage & " puntos de fuerza.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
        End If
        
        'Baja agilidad
        If Hechizos(Spell).SubeAgilidad = 2 Then
            Damage = RandomNumber(Hechizos(Spell).MinAgilidad, Hechizos(Spell).MaxAgilidad)
            tmpByte = .Stats.UserAtributos(eAtributos.Agilidad)
            tmpByte = tmpByte - Damage
            If tmpByte < ConstantesBalance.MinAtributos Then tmpByte = ConstantesBalance.MinAtributos
            .Stats.UserAtributos(eAtributos.Agilidad) = tmpByte
            
            .flags.TomoPocion = True
            .flags.DuracionEfecto = 700
            
            Call WriteUpdateDexterity(UserIndex)
            
            ' Disable the Berserk
            If HasPassiveAssigned(UserIndex, ePassiveSpells.Berserk) Then
                If HasPassiveActivated(UserIndex, ePassiveSpells.Berserk) Then
                    Call ActivatePassive(UserIndex, ePassiveSpells.Berserk, False)
                    Call SendBerserkEffect(UserIndex, ePassiveSpells.Berserk, False)
                End If
            End If
            
            Call SendSpellEffects(UserIndex, NpcIndex, Spell, DecirPalabras)

            Call WriteConsoleMsg(UserIndex, Npclist(NpcIndex).Name & " te ha quitado " & Damage & " puntos de agilidad.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
        End If
        
        'Baja mana
        If Hechizos(Spell).SubeMana = 2 Then
            Damage = RandomNumber(Hechizos(Spell).MinMana, Hechizos(Spell).MaxMana)
            .Stats.MinMAN = .Stats.MinMAN - Damage
            
            Call SendSpellEffects(UserIndex, NpcIndex, Spell, DecirPalabras)
            
            If .Stats.MinMAN < 0 Then .Stats.MinMAN = 0
            
            Call WriteConsoleMsg(UserIndex, Npclist(NpcIndex).Name & " te ha quitado " & Damage & " puntos de mana.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
            Call WriteUpdateUserStats(UserIndex)
        End If
        
        'Damage en area
        If Hechizos(Spell).Area > 0 Then
            Dim X As Byte
            Dim Y As Byte
    
            For X = .Pos.X - Hechizos(Spell).Area To .Pos.X + Hechizos(Spell).Area
                For Y = .Pos.Y - Hechizos(Spell).Area To .Pos.Y + Hechizos(Spell).Area
                    UI = MapData(.Pos.Map, X, Y).UserIndex
                    If UI > 0 Then
                        UserProtected = Not IntervaloPermiteSerAtacado(UI) And UserList(UI).flags.NoPuedeSerAtacado
                        UserProtected = UserProtected Or UserList(UI).flags.Ignorado Or UserList(UI).flags.HelpMode
                            
                        If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.AdminPerseguible And Not UserProtected Then
                            If .flags.Privilegios And PlayerType.User Then
                                Call SendSpellEffects(UI, NpcIndex, Spell, DecirPalabras)
                                
                                Damage = RandomNumber(Hechizos(Spell).MinHp, Hechizos(Spell).MaxHp)
                                
                                If UserList(UI).Invent.CascoEqpObjIndex > 0 Then
                                    Damage = Damage - RandomNumber(ObjData(UserList(UI).Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(UserList(UI).Invent.CascoEqpObjIndex).DefensaMagicaMax)
                                End If
                                
                                If UserList(UI).Invent.AnilloEqpObjIndex > 0 Then
                                    Damage = Damage - RandomNumber(ObjData(UserList(UI).Invent.AnilloEqpObjIndex).DefensaMagicaMin, ObjData(UserList(UI).Invent.AnilloEqpObjIndex).DefensaMagicaMax)
                                End If
                                
                                If UserList(UI).flags.Navegando = 1 And UserList(UI).Invent.BarcoObjIndex > 0 Then
                                    Damage = Damage - RandomNumber(ObjData(UserList(UI).Invent.BarcoObjIndex).DefensaMagicaMin, ObjData(UserList(UI).Invent.BarcoObjIndex).DefensaMagicaMax)
                                End If
                                Damage = Max(1, Damage)
                            
                                UserList(UI).Stats.MinHp = UserList(UI).Stats.MinHp - Damage
                                
                                If UserList(UI).Stats.MinHp < 0 Then UserList(UI).Stats.MinHp = 0
                                
                                Call WriteConsoleMsg(UI, Npclist(NpcIndex).Name & " te ha quitado " & Damage & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                                Call WriteUpdateUserStats(UI)
                                
                                'Muere
                                If UserList(UI).Stats.MinHp < 1 Then
                                    UserList(UI).Stats.MinHp = 0
                                    
                                    Dim MasterIndex As Integer
                                    MasterIndex = Npclist(NpcIndex).MaestroUser
                                    
                                    '[Barrin 1-12-03]
                                    If MasterIndex > 0 Then
                                        
                                        ' No son frags los muertos atacables
                                        If UserList(UI).flags.AtacablePor <> MasterIndex Then
                                            'Store it!
                                            Call Statistics.StoreFrag(MasterIndex, UI)
                                            
                                            Call ContarMuerte(UI, MasterIndex, eDamageType.NpcSpell, Damage, Spell)
                                        End If
                                        
                                        Call ActStats(UI, MasterIndex)
                                    End If
                                    '[/Barrin]
                                    
                                    Call UserDie(UI)
                                    
                                End If
                            End If
                        End If
                    End If
                Next Y
            Next X
        End If
        
        'Salta de user a user
        If Hechizos(Spell).Salta > 0 Then
            Dim I As Integer
            Dim c As Integer
            Dim Saltos As Byte
            Dim SaltoActual As Integer
            Dim Saltando As Byte
            Dim UserTarget() As Integer
            Dim Tmp As Integer
            Dim tmpTarget() As Integer
            Dim tmpSalto As Integer
            Dim FoundSalto As Byte
            
            ReDim UserTarget(1 To 1) As Integer
            ReDim tmpTarget(1 To 1) As Integer

            Saltando = 1
     
            Dim query() As Collision.UUID
            For I = 0 To ModAreas.QueryEntities(NpcIndex, ENTITY_TYPE_NPC, query, ENTITY_TYPE_PLAYER)
                UI = query(I).Name
  
                UserProtected = Not IntervaloPermiteSerAtacado(UI) And UserList(UI).flags.NoPuedeSerAtacado
                UserProtected = UserProtected Or UserList(UI).flags.Ignorado Or UserList(UI).flags.HelpMode
                        
                If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.AdminPerseguible And Not UserProtected Then
                    Tmp = Tmp + 1
                    ReDim Preserve tmpTarget(1 To Tmp) As Integer
                    tmpTarget(Tmp) = UI

                    If RandomNumber(1, UBound(query)) = 1 Or (I = UBound(query) And SaltoActual = 0) Or SaltoActual = 0 Then
                        SaltoActual = Tmp
                        UserTarget(1) = UI
                        Saltos = 1
                    End If
                End If
            Next I
                
            If tmpTarget(1) = 0 Then Exit Sub
            
            Do While Saltando = 1
                tmpSalto = SaltoActual
                For I = 1 To UBound(tmpTarget)
                    UI = tmpTarget(I)
                    FoundSalto = 0
                    For c = 1 To Saltos
                        If UI = UserTarget(c) Then
                            FoundSalto = 1
                            Exit For
                        End If
                    Next c
                    If FoundSalto = 0 Then
                        If Distancia(UserList(UI).Pos, UserList(tmpTarget(SaltoActual)).Pos) <= Hechizos(Spell).DistanciaSalto Then
                            Saltos = Saltos + 1
                            SaltoActual = I
                            ReDim Preserve UserTarget(1 To Saltos) As Integer
                            UserTarget(UBound(UserTarget)) = UI
                        End If
                    End If
                    If Saltos = Hechizos(Spell).Salta Then
                        Saltando = 0
                        Exit For
                    End If
                Next I
                If tmpSalto = SaltoActual Then Saltando = 0
            Loop
            
            If Saltos > 0 Then
                Damage = RandomNumber(Hechizos(Spell).MinHp, Hechizos(Spell).MaxHp)
                Damage = Damage / Saltos
                
                For I = 1 To Saltos
                    If .flags.Privilegios And PlayerType.User Then
                        Call SendSpellEffects(UserTarget(I), NpcIndex, Spell, DecirPalabras)
                        
                        If UserList(UserTarget(I)).Invent.CascoEqpObjIndex > 0 Then
                            Damage = Damage - RandomNumber(ObjData(UserList(UserTarget(I)).Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(UserList(UserTarget(I)).Invent.CascoEqpObjIndex).DefensaMagicaMax)
                        End If
                        
                        If UserList(UserTarget(I)).Invent.AnilloEqpObjIndex > 0 Then
                            Damage = Damage - RandomNumber(ObjData(UserList(UserTarget(I)).Invent.AnilloEqpObjIndex).DefensaMagicaMin, ObjData(UserList(UserTarget(I)).Invent.AnilloEqpObjIndex).DefensaMagicaMax)
                        End If
                        
                        If UserList(UserTarget(I)).flags.Navegando = 1 And UserList(UserTarget(I)).Invent.BarcoObjIndex > 0 Then
                            Damage = Damage - RandomNumber(ObjData(UserList(UserTarget(I)).Invent.BarcoObjIndex).DefensaMagicaMin, ObjData(UserList(UserTarget(I)).Invent.BarcoObjIndex).DefensaMagicaMax)
                        End If
                        
                        Damage = Max(1, Damage)
                                            
                        UserList(UserTarget(I)).Stats.MinHp = UserList(UserTarget(I)).Stats.MinHp - Damage
                        
                        If UserList(UserTarget(I)).Stats.MinHp < 0 Then UserList(UserTarget(I)).Stats.MinHp = 0
                        
                        Call WriteConsoleMsg(UserTarget(I), Npclist(NpcIndex).Name & " te ha quitado " & Damage & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                        Call WriteUpdateUserStats(UserTarget(I))
                        
                        'Muere
                        If UserList(UserTarget(I)).Stats.MinHp < 1 Then
                            UserList(UserTarget(I)).Stats.MinHp = 0
                            Call UserDie(UserTarget(I))
                        End If
                    End If
                Next I
            End If
        End If
        
        'Petrificar
        If Hechizos(Spell).Petrificar = 1 Then
            If .flags.Petrificado = 0 Then
                
                If Hechizos(Spell).ByPassPassive = 0 Then
                    If PassiveConditionMet(UserIndex, ParalysisImmunity) Then
                        Call WriteConsoleMsg(UserIndex, "Tu habilidades te protegen del hechizo petrificante del NPC " & Npclist(NpcIndex).Name & ".", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                        Exit Sub
                    End If
                End If
                
                Call SendSpellEffects(UserIndex, NpcIndex, Spell, DecirPalabras)

                If .flags.Paralizado = 0 And .flags.Inmovilizado = 0 Then Call WriteParalizeOK(UserIndex)
                
                Call WriteConsoleMsg(UserIndex, "¡¡" & Npclist(NpcIndex).Name & " te ha petrificado!!.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                
                .flags.Petrificado = 1
                .flags.Paralizado = 1
                .Counters.Petrificado = SetIntervalEnd(ServerConfiguration.Intervals.IntervaloParalizado / 5)
            End If
        End If
        
        'Teletransportacion
        If Hechizos(Spell).Teletransportacion = 1 Then
            Call ClosestLegalPos(.Pos, nPos)
            
            MapData(Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y).NpcIndex = 0
            MapData(nPos.Map, nPos.X, nPos.Y).NpcIndex = NpcIndex
            Npclist(NpcIndex).Pos = nPos
            
            Call WriteConsoleMsg(UserIndex, "¡¡" & Npclist(NpcIndex).Name & " ha aparecido sobre ti!!.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)

            Call ModAreas.UpdateEntity(NpcIndex, ENTITY_TYPE_NPC, Npclist(NpcIndex).Pos, True)
        End If
        
        'Putrefaccion
        If Hechizos(Spell).Putrefaccion = 1 Then
            If .flags.Putrefaccion = 0 Then
                
                Call SendSpellEffects(UserIndex, NpcIndex, Spell, DecirPalabras)

                If .flags.Paralizado = 0 And .flags.Inmovilizado = 0 Then Call WriteParalizeOK(UserIndex)
                
                Call WriteConsoleMsg(UserIndex, Npclist(NpcIndex).Name & " te ha lanzado Putrefacción.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                
                .flags.Putrefaccion = Spell
                .flags.Paralizado = 1
                .Counters.Putrefaccion = SetIntervalEnd(ServerConfiguration.Intervals.IntervaloParalizado)
                .Counters.PutrefaccionDmg = SetIntervalEnd(ServerConfiguration.Intervals.IntervaloPutrefaccionDmg)
            End If
        End If

        ' Heal HP
        If Hechizos(Spell).SubeHP = 1 Then
        
            Call SendSpellEffects(UserIndex, NpcIndex, Spell, DecirPalabras)
        
            Damage = RandomNumber(Hechizos(Spell).MinHp, Hechizos(Spell).MaxHp)
        
            .Stats.MinHp = .Stats.MinHp + Damage
            If .Stats.MinHp > .Stats.MaxHp Then .Stats.MinHp = .Stats.MaxHp
            
            Call WriteConsoleMsg(UserIndex, Npclist(NpcIndex).Name & " te ha quitado " & Damage & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
            Call WriteUpdateUserStats(UserIndex)
        
        ' Damage
        ElseIf Hechizos(Spell).SubeHP = 2 Then
            
            If .flags.Privilegios And PlayerType.User Then
            
                Call SendSpellEffects(UserIndex, NpcIndex, Spell, DecirPalabras)
            
                Damage = RandomNumber(Hechizos(Spell).MinHp, Hechizos(Spell).MaxHp)
                
                If .Invent.CascoEqpObjIndex > 0 Then
                    Damage = Damage - RandomNumber(ObjData(.Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(.Invent.CascoEqpObjIndex).DefensaMagicaMax)
                End If
                
                If .Invent.AnilloEqpObjIndex > 0 Then
                    Damage = Damage - RandomNumber(ObjData(.Invent.AnilloEqpObjIndex).DefensaMagicaMin, ObjData(.Invent.AnilloEqpObjIndex).DefensaMagicaMax)
                End If
                
                If .flags.Navegando = 1 And .Invent.BarcoObjIndex > 0 Then
                    Damage = Damage - RandomNumber(ObjData(.Invent.BarcoObjIndex).DefensaMagicaMin, ObjData(.Invent.BarcoObjIndex).DefensaMagicaMax)
                End If
                
                Damage = Max(1, Damage)
            
                .Stats.MinHp = .Stats.MinHp - Damage
                                
                If .Stats.MinHp < 0 Then .Stats.MinHp = 0
                
                Call WriteConsoleMsg(UserIndex, Npclist(NpcIndex).Name & " te ha quitado " & Damage & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                
                ' Apply the mana conversion effect to the user receiving the magic damage
                Call ManaConversionEffect(UserIndex, Damage)
                
                Call WriteUpdateUserStats(UserIndex)
                
                'Muere
                If .Stats.MinHp < 1 Then
                    .Stats.MinHp = 0
                    If Npclist(NpcIndex).MaestroUser > 0 Then
                        If UserList(UserIndex).flags.AtacablePor <> Npclist(NpcIndex).MaestroUser Then
                            Call Statistics.StoreFrag(Npclist(NpcIndex).MaestroUser, UserIndex)
                            Call ContarMuerte(UserIndex, Npclist(NpcIndex).MaestroUser, eDamageType.NpcSpell, Damage, Spell)
                        End If
                        Call ActStats(UserIndex, Npclist(NpcIndex).MaestroUser)
                    End If
                    Call UserDie(UserIndex)
                End If
            End If
        End If
        
        ' Paralisis/Inmobilize
        If Hechizos(Spell).Paraliza = 1 Or Hechizos(Spell).Inmoviliza = 1 Then
            If UserList(UserIndex).flags.AdminPerseguible = False Then Exit Sub
            If Hechizos(Spell).Paraliza = 1 Then
                If Hechizos(Spell).ByPassPassive = 0 Then
                    If PassiveConditionMet(UserIndex, ParalysisImmunity) Then
                        Call WriteConsoleMsg(UserIndex, "Tu habilidades te protegen del hechizo paralizante del NPC " & Npclist(NpcIndex).Name & ".", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                        Exit Sub
                    End If
                End If
            End If
            
            If .flags.Paralizado = 0 Then
        
                 If AnilloObjIndex > 0 Then
                    If ObjData(AnilloObjIndex).ImpideParalizar <> 0 Then
                        Call WriteConsoleMsg(UserIndex, "Tu anillo rechaza los efectos de la paralisis.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                        Exit Sub
                    End If
                End If
                
                ' Berserk's protection now applies to both immo and paralysis spells
                If Hechizos(Spell).Inmoviliza = 1 Or Hechizos(Spell).Paraliza = 1 Then
                    If Hechizos(Spell).ByPassPassive = 0 Then
                        If HasPassiveActivated(UserIndex, ePassiveSpells.Berserk) Then
                            Call WriteConsoleMsg(UserIndex, "Tu habilidades te protegen del hechizo " & IIf(Hechizos(Spell).Inmoviliza = 1, "inmovilizante", "paralizante") & " del NPC " & Npclist(NpcIndex).Name & ".", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                            Exit Sub
                        End If
                    End If
                End If
                
                If Hechizos(Spell).Inmoviliza = 1 Then
                    If AnilloObjIndex > 0 Then
                        If ObjData(AnilloObjIndex).ImpideInmobilizar <> 0 Then
                            .flags.Inmovilizado = 0
                            Call WriteConsoleMsg(UserIndex, "Tu anillo rechaza los efectos del hechizo inmovilizar.", FontTypeNames.FONTTYPE_FIGHT)
                            Exit Sub
                        End If
                    End If
                    
                    .flags.Inmovilizado = 1
                End If
                  
                .flags.Paralizado = 1
                .Counters.Paralisis = SetIntervalEnd(ServerConfiguration.Intervals.IntervaloParalizado)
                  
                Call WriteConsoleMsg(UserIndex, Npclist(NpcIndex).Name & " te ha paralizado.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                Call SendSpellEffects(UserIndex, NpcIndex, Spell, DecirPalabras) ' G Toyz: Llamada reposicionada.
                Call WriteParalizeOK(UserIndex)
                
            End If
            
        End If
        
        ' Stupidity
        If Hechizos(Spell).Estupidez = 1 Then
             
            If .flags.Estupidez = 0 Then
            
                Call SendSpellEffects(UserIndex, NpcIndex, Spell, DecirPalabras)
                
                If Hechizos(Spell).ByPassPassive = 0 Then
                    If PassiveConditionMet(UserIndex, IndomitableWill) Then
                        Call WriteConsoleMsg(UserIndex, "Tu habilidades te protegen del efecto de turbación del NPC " & Npclist(NpcIndex).Name & ".", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                        Exit Sub
                    End If
                End If
                
                If AnilloObjIndex > 0 Then
                    If ObjData(AnilloObjIndex).ImpideAturdir <> 0 Then
                        Call WriteConsoleMsg(UserIndex, "Tu anillo rechaza los efectos de la turbación.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                        Exit Sub
                    End If
                End If
                  
                Call WriteConsoleMsg(UserIndex, Npclist(NpcIndex).Name & " te ha lanzado turbación.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                
                .flags.Estupidez = 1
                .Counters.Ceguera = SetIntervalEnd(ServerConfiguration.Intervals.IntervaloInvisible)
                          
                Call WriteDumb(UserIndex)
                
            End If
        End If
        
        ' Blind
        If Hechizos(Spell).Ceguera = 1 Then
             
            If .flags.Ceguera = 0 Then
            
                Call SendSpellEffects(UserIndex, NpcIndex, Spell, DecirPalabras)
            
                If AnilloObjIndex > 0 Then
                    If ObjData(AnilloObjIndex).ImpideCegar <> 0 Then
                        Call WriteConsoleMsg(UserIndex, "Tu anillo rechaza los efectos de la ceguera.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                        Exit Sub
                    End If
                End If
                  
                .flags.Ceguera = 1
                .Counters.Ceguera = SetIntervalEnd(ServerConfiguration.Intervals.IntervaloInvisible)
                          
                Call WriteBlind(UserIndex)
                
            End If
        End If
        
        ' Remove Invisibility/Hidden
        If Hechizos(Spell).RemueveInvisibilidadParcial = 1 Then
                 
            Call SendSpellEffects(UserIndex, NpcIndex, Spell, DecirPalabras)
                 
            'Sacamos el efecto de ocultarse
            If .flags.Oculto = 1 Then
                .flags.Oculto = 0
                .Counters.TiempoOculto = 0
                Call UsUaRiOs.SetInvisible(UserIndex, .Char.CharIndex, False)
                Call WriteConsoleMsg(UserIndex, "¡Has sido detectado!", FontTypeNames.FONTTYPE_VENENO, eMessageType.Combate)
            Else
                'sino, solo lo "iniciamos" en la sacada de invisibilidad.
                Call WriteConsoleMsg(UserIndex, "Comienzas a hacerte visible.", FontTypeNames.FONTTYPE_VENENO, eMessageType.Combate)
                .Counters.Invisibilidad = ServerConfiguration.Intervals.IntervaloInvisible - 1
            End If
        
        End If
        
    End With
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub NpcLanzaSpellSobreUser de modHechizos.bas")
End Sub

Private Sub SendSpellEffects(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal Spell As Integer, _
                             ByVal DecirPalabras As Boolean)
'***************************************************
'Author: ZaMa
'Last Modification: 11/11/2010
'Sends spell's wav, fx and mgic words to users.
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)
        ' Spell Wav
        Call SendData(SendTarget.ToPCArea, UserIndex, _
            PrepareMessagePlayWave(Hechizos(Spell).WAV, .Pos.X, .Pos.Y, Npclist(NpcIndex).Char.CharIndex))
            
        ' Spell FX
        Call SendData(SendTarget.ToPCArea, UserIndex, _
            PrepareMessageCreateFX(.Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).Loops))
    
        ' Spell Words
        If DecirPalabras Then
            Call SendData(SendTarget.ToNPCArea, NpcIndex, _
                PrepareMessageChatOverHead(Hechizos(Spell).PalabrasMagicas, Npclist(NpcIndex).Char.CharIndex, vbCyan))
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SendSpellEffects de modHechizos.bas")
End Sub

Public Sub NpcLanzaSpellSobreNpc(ByVal NpcIndex As Integer, ByVal TargetNpc As Integer, _
                                 ByVal SpellIndex As Integer, Optional ByVal DecirPalabras As Boolean = False)
'***************************************************
'Author: Unknown
'Last Modification: 21/09/2010
'21/09/2010: ZaMa - Now npcs can cast a wider range of spells.
'***************************************************
On Error GoTo ErrHandler
  

    Dim Danio As Integer
    
    With Npclist(TargetNpc)
    
    
        ' Spell sound and FX
        Call SendData(SendTarget.ToNPCArea, TargetNpc, _
            PrepareMessagePlayWave(Hechizos(SpellIndex).WAV, .Pos.X, .Pos.Y, .Char.CharIndex))
            
        Call SendData(SendTarget.ToNPCArea, TargetNpc, _
            PrepareMessageCreateFX(.Char.CharIndex, Hechizos(SpellIndex).FXgrh, Hechizos(SpellIndex).Loops))
    
        ' Decir las palabras magicas?
        If DecirPalabras Then
            Call SendData(SendTarget.ToNPCArea, NpcIndex, _
                PrepareMessageChatOverHead(Hechizos(SpellIndex).PalabrasMagicas, Npclist(NpcIndex).Char.CharIndex, vbCyan))
        End If

    
        ' Spell deals damage??
        If Hechizos(SpellIndex).SubeHP = 2 Then
            
            Danio = RandomNumber(Hechizos(SpellIndex).MinHp, Hechizos(SpellIndex).MaxHp)
            
            ' Deal damage
            .Stats.MinHp = .Stats.MinHp - Danio
            
            'Muere?
            If .Stats.MinHp < 1 Then
                .Stats.MinHp = 0
                If Npclist(NpcIndex).MaestroUser > 0 Then
                    Call MuereNpc(TargetNpc, Npclist(NpcIndex).MaestroUser)
                Else
                    Call MuereNpc(TargetNpc, 0)
                End If
            End If
            
        ' Spell recovers health??
        ElseIf Hechizos(SpellIndex).SubeHP = 1 Then
            
            Danio = RandomNumber(Hechizos(SpellIndex).MinHp, Hechizos(SpellIndex).MaxHp)
            
            ' Recovers health
            .Stats.MinHp = .Stats.MinHp + Danio
            
            If .Stats.MinHp > .Stats.MaxHp Then
                .Stats.MinHp = .Stats.MaxHp
            End If
            
        End If
        
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
            .Contadores.Paralisis = SetIntervalEnd(ServerConfiguration.Intervals.IntervaloNPCParalizado)
            
        ElseIf Hechizos(SpellIndex).Inmoviliza = 1 Then
            .flags.Inmovilizado = 1
            .flags.Paralizado = 0
            .Contadores.Paralisis = SetIntervalEnd(ServerConfiguration.Intervals.IntervaloNPCParalizado)
            
        ElseIf Hechizos(SpellIndex).RemoverParalisis = 1 Then
            If .flags.Paralizado = 1 Or .flags.Inmovilizado = 1 Then
                .flags.Paralizado = 0
                .flags.Inmovilizado = 0
                .Contadores.Paralisis = 0
            End If
        End If
    
    End With

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub NpcLanzaSpellSobreNpc de modHechizos.bas")
End Sub

Function TieneHechizo(ByVal I As Integer, ByVal UserIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo ErrHandler
    
    Dim J As Integer
    For J = 1 To MAXUSERHECHIZOS
        If UserList(UserIndex).Stats.UserHechizos(J).SpellNumber = I Then
            TieneHechizo = True
            Exit Function
        End If
    Next

Exit Function
ErrHandler:

End Function

Sub AgregarHechizo(ByVal UserIndex As Integer, ByVal Slot As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

Dim hIndex As Integer

With UserList(UserIndex)
    hIndex = ObjData(.Invent.Object(Slot).ObjIndex).HechizoIndex
    
    If AddSpellByIndex(UserIndex, hIndex) Then
        'Quitamos del inv el item
        Call QuitarUserInvItem(UserIndex, CByte(Slot), 1)
    End If
    
    'If Not TieneHechizo(hIndex, UserIndex) Then
    '    'Buscamos un slot vacio
    '    For J = 1 To MAXUSERHECHIZOS
    '        If .Stats.UserHechizos(J) = 0 Then Exit For
    '    Next J
    '
    '    If .Stats.UserHechizos(J) <> 0 Then
    '        Call WriteConsoleMsg(UserIndex, "No tienes espacio para más hechizos.", FontTypeNames.FONTTYPE_INFO)
    '    Else
    '        .Stats.UserHechizos(J) = hIndex
    '        Call UpdateUserHechizos(False, UserIndex, CByte(J))
    '        'Quitamos del inv el item
    '        Call QuitarUserInvItem(UserIndex, CByte(Slot), 1)
    '    End If
    'Else
    '    Call WriteConsoleMsg(UserIndex, "Ya tienes ese hechizo.", FontTypeNames.FONTTYPE_INFO)
    'End If
End With

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub AgregarHechizo de modHechizos.bas")
End Sub

Public Function AddSpellByIndex(ByVal nUserIndex As Integer, ByVal nSpellIndex As Integer) As Boolean
'***************************************************
'Author: D'Artagnan (Taken from AgregarHechizo)
'Last Modification: 31/05/2015
'
'***************************************************
On Error GoTo ErrHandler
  
    Dim J As Integer
    
    With UserList(nUserIndex)
        If Not TieneHechizo(nSpellIndex, nUserIndex) Then
            'Buscamos un slot vacio
            For J = 1 To MAXUSERHECHIZOS
                If .Stats.UserHechizos(J).SpellNumber = 0 Then Exit For
            Next J
                
            If .Stats.UserHechizos(J).SpellNumber <> 0 Then
                Call WriteConsoleMsg(nUserIndex, "No tienes espacio para más hechizos.", FontTypeNames.FONTTYPE_INFO)
            Else
                .Stats.UserHechizos(J).SpellNumber = nSpellIndex
                .Stats.UserHechizos(J).LastUsedAt = 0
                .Stats.UserHechizos(J).LastUsedSuccessfully = False
                Call UpdateUserHechizos(False, nUserIndex, CByte(J))
                AddSpellByIndex = True
            End If
        Else
            Call WriteConsoleMsg(nUserIndex, "Ya tienes ese hechizo.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function AddSpellByIndex de modHechizos.bas")
End Function
            
Sub DecirPalabrasMagicas(ByVal SpellWords As String, ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 17/11/2009
'25/07/2009: ZaMa - Invisible admins don't say any word when casting a spell
'17/11/2009: ZaMa - Now the user become visible when casting a spell, if it is hidden
'11/06/2011: CHOTS - Color de dialogos customizables
'***************************************************
On Error GoTo ErrHandler

    With UserList(UserIndex)
        If .flags.AdminInvisible <> 1 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatPersonalizado(SpellWords, .Char.CharIndex, 5))
            
            ' Si estaba oculto, se vuelve visible
            If .flags.Oculto = 1 Then
                .flags.Oculto = 0
                .Counters.TiempoOculto = 0
                
                If .flags.invisible = 0 Then
                    Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)
                    Call UsUaRiOs.SetInvisible(UserIndex, .Char.CharIndex, False)
                End If
            End If
        End If
    End With
    
    Exit Sub
    
ErrHandler:
    Call LogError("Error en DecirPalabrasMagicas. Error: " & Err.Number & " - " & Err.Description)
End Sub

''
' Check if an user can cast a certain spell
'
' @param UserIndex Specifies reference to user
' @param HechizoIndex Specifies reference to spell
' @return   True if the user can cast the spell, otherwise returns false
Function PuedeLanzar(ByVal UserIndex As Integer, ByVal HechizoIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 05/04/2015 (D'Artagnan)
'06/11/09 - Corregida la bonificación de maná del mimetismo en el druida con flauta mágica equipada.
'19/11/2009: ZaMa - Validacion de mana para el Invocar Mascotas
'12/01/2010: ZaMa - Validacion de mana para hechizos lanzados por druida.
'05/04/2015: D'Artagnan - New requirements for casting spells.
'23/05/2015: D'Artagnan - Minimum level.
'08/07/2016: Anagrama - Invocar mascota ahora requiere amuleto ligero o flauta elfica.
'                       Reduccion de mana para bardo con laud elfico y mago con engarzado al invocar es ahora de 26%.
'***************************************************
On Error GoTo ErrHandler
    Dim sRequirements As String

    With UserList(UserIndex)
        If .flags.HelpMode Then
            Call WriteConsoleMsg(UserIndex, "No puedes lanzar hechizos si estás en consulta.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
    
        If .flags.Muerto Then
            Call WriteConsoleMsg(UserIndex, "No puedes lanzar hechizos estando muerto.", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)
            Exit Function
        End If
        
        ' Requirements for casting the spell.
        
        If .Stats.ELV < Hechizos(HechizoIndex).MinLevel Then
            Call WriteConsoleMsg(UserIndex, "Debes ser nivel " & Hechizos(HechizoIndex).MinLevel & _
                                 " o mayor para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
                
        If GetSkills(UserIndex, eSkill.Magia) < Hechizos(HechizoIndex).MinSkill Then
            Call WriteConsoleMsg(UserIndex, "No tienes suficientes puntos de magia para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)
            Exit Function
        End If
        
        If .Stats.MinSta < Hechizos(HechizoIndex).StaRequerido Then
            If .Genero = eGenero.Hombre Then
                Call WriteConsoleMsg(UserIndex, "Estás muy cansado para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)
            Else
                Call WriteConsoleMsg(UserIndex, "Estás muy cansada para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)
            End If
            Exit Function
        End If
        
        If Not CanCastSpellByMagicPower(UserIndex, HechizoIndex) Then
            Call WriteConsoleMsg(UserIndex, "No tienes suficiente poder de casteo para lanzar este hechizo", FontTypeNames.FONTTYPE_INFO, Combate)
            Exit Function
        End If
            
        If .clase = eClass.Druid And Hechizos(HechizoIndex).Warp = 1 Then
            ' Si no tiene mascotas, no tiene sentido que lo use
            If .TammedPetsCount = 0 Then
                Call WriteConsoleMsg(UserIndex, "Debes poseer alguna mascota para poder lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)
                Exit Function
            End If
        End If
        
        If .Stats.MinMAN < modHechizos.GetSpellRequiredMana(HechizoIndex, UserIndex) Then
            Call WriteConsoleMsg(UserIndex, "No tienes suficiente maná.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        
        If Hechizos(HechizoIndex).RequireFullMana = 1 And (.Stats.MinMAN < .Stats.MaxMan) Then
            Call WriteConsoleMsg(UserIndex, "Este hechizo requiere tu barra de maná completa.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        
        ' Check if thespell can be used during a duel
        If .flags.DueloIndex > 0 Then
            If Not CanUseSpellInDuel(HechizoIndex) Then
                Call WriteConsoleMsg(UserIndex, "No puedes usar este hechizo durante un duelo.", FontTypeNames.FONTTYPE_INFO)
                Exit Function
            End If
        End If
        
    End With
    
    PuedeLanzar = True
    
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function PuedeLanzar de modHechizos.bas")
End Function


Public Function CanUseSpellInDuel(ByVal SpellNumber As Integer)
On Error GoTo ErrHandler:

    Dim I As Integer
    
    For I = 1 To ConstantesBalance.DuelProhibitedSpellsQty
        If ConstantesBalance.DuelProhibitedSpells(I) = SpellNumber Then
            CanUseSpellInDuel = False
            Exit Function
        End If
    Next I
    
    CanUseSpellInDuel = True
    
    Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function PuedeLanzar de modHechizos.bas")
End Function

Function HechizoInvocacion(ByVal UserIndex As Integer, ByRef TargetPos As WorldPos, ByVal SpellIndex As Integer) As Boolean
On Error GoTo ErrHandler
    Dim invokedPet As Byte
    Dim NroNpcs As Integer, NpcIndex As Integer, PetIndex As Integer
    Dim ActiveTammedPetsQty As Integer
    Dim I As Integer
    Dim isWaterTile As Boolean
    Dim TammedPetIndex As Integer
    Dim OldRemainingHp As Integer
    
With UserList(UserIndex)
    
    'No permitimos se invoquen criaturas en zonas seguras
    If MapInfo(TargetPos.Map).Pk = False Or MapData(TargetPos.Map, TargetPos.X, TargetPos.Y).Trigger = eTrigger.ZONASEGURA Then
        Call WriteConsoleMsg(UserIndex, "No puedes invocar criaturas en zona segura.", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)
        Exit Function
    End If
    
    'No permitimos se invoquen criaturas en mapas donde esta prohibido hacerlo
    If MapInfo(TargetPos.Map).InvocarSinEfecto = 1 Then
        Call WriteConsoleMsg(UserIndex, "Invocar no está permitido aquí! Retirate de la Zona si deseas utilizar el Hechizo.", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)
        Exit Function
    End If
    
    ' not allowed to invoke pets on TPs
    Dim ObjIndex As Integer
    ObjIndex = MapData(TargetPos.Map, TargetPos.X, TargetPos.Y).ObjInfo.ObjIndex
    If ObjIndex <> 0 Then
        If ObjData(ObjIndex).ObjType = eOBJType.otTeleport Then
            Call WriteConsoleMsg(UserIndex, "No puedes invocar una mascota sobre un teleport", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)
            Exit Function
        End If
    End If
    
    'No usar elementales en desafios si no lo permite
    'If .Challenge.InSand > 0 And .Pos.Map = SandsChallenge(.Challenge.InSand).Event_map Then
    '    If SandsChallenge(.Challenge.InSand).Elementary > 0 Then
    '        Call WriteConsoleMsg(UserIndex, "¡Invocar no está permitido aquí!", FontTypeNames.FONTTYPE_INFO)
    '        HechizoCasteado = False
    '        Exit Sub
    '    End If
    'End If
    
    For I = 1 To Classes(.clase).ClassMods.MaxTammedPets
        If .TammedPets(I).NpcIndex <> 0 Then
            ActiveTammedPetsQty = ActiveTammedPetsQty + 1
            TammedPetIndex = I
        End If
    Next I
    
    ' Warp de mascotas
    If Hechizos(SpellIndex).Warp = 1 Then
    
        If .SelectedPet = 0 Then
            Call WriteConsoleMsg(UserIndex, "No tenes ninguna mascota seleccionada.", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)
            Exit Function
        End If

        If (.InvokedPetsCount >= Classes(.clase).ClassMods.MaxInvokedPets Or _
           (.InvokedPetsCount + ActiveTammedPetsQty) >= Classes(.clase).ClassMods.MaxActivePets) And _
            .TammedPets(.SelectedPet).NpcIndex = 0 Then
            If ActiveTammedPetsQty > 0 Then
                OldRemainingHp = .TammedPets(TammedPetIndex).RemainingLife
                
                Call QuitarNPC(.TammedPets(TammedPetIndex).NpcIndex)
                
                .TammedPets(TammedPetIndex).RemainingLife = OldRemainingHp
            Else
                Call WriteConsoleMsg(UserIndex, "Has superado la cantidad máxima de mascotas invocadas.", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)
                Exit Function
            End If
        End If
        
        If .TammedPets(.SelectedPet).NpcNumber = 0 Then
            Call WriteConsoleMsg(UserIndex, "No tenes ninguna mascota seleccionada.", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)
            Exit Function
        End If
        
        If .TammedPets(.SelectedPet).RemainingLife = 0 Then
            Call WriteConsoleMsg(UserIndex, "Este hechizo solo funciona con mascotas que estén vivas.", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)
            Exit Function
        End If
        
        isWaterTile = HayAgua(TargetPos.Map, TargetPos.X, TargetPos.Y)
        
        ' Check whether the Pet can be invoked or not based on his ability to walk over the water or ground.
        If isWaterTile And NpcData(.TammedPets(.SelectedPet).NpcNumber).flags.AguaValida = 0 Then
            Call WriteConsoleMsg(UserIndex, "Tu mascota " & NpcData(.TammedPets(.SelectedPet).NpcNumber).Name & " no puede transitar sobre el agua. Intenta invocarla en otra parte.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        If Not isWaterTile And NpcData(.TammedPets(.SelectedPet).NpcNumber).flags.TierraInvalida = 1 Then
            Call WriteConsoleMsg(UserIndex, "Tu mascota " & NpcData(.TammedPets(.SelectedPet).NpcNumber).Name & " no puede transitar sobre la tierra. Intenta invocarla en otra parte.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        
        ' Select the pet as the target NPC.
        .flags.LastNpcInvoked = .TammedPets(.SelectedPet).NpcIndex
 
        PetIndex = .TammedPets(.SelectedPet).NpcIndex
                
        ' La invoco cerca mio
        If PetIndex > 0 Then
            
                    
            If Not WarpMascota(UserIndex, .SelectedPet) Then
                Exit Function
            End If
            Call SubirSkill(UserIndex, eSkill.Domar, True)
        Else ' Si la mascota no fue creada, entonces la creamos
            
            'Validamos si la mascota está muerta
            If .TammedPets(.SelectedPet).RemainingLife = 0 Then
                Call WriteConsoleMsg(UserIndex, "Este hechizo solo funciona con mascotas que estén vivas.", FontTypeNames.FONTTYPE_INFO)
                Exit Function
            End If
        
            Dim MascotasIndex As Integer
            MascotasIndex = SpawnNpc(.TammedPets(.SelectedPet).NpcNumber, TargetPos, False, True)
            .TammedPets(.SelectedPet).NpcIndex = MascotasIndex
            
            .flags.LastNpcInvoked = MascotasIndex
            
            If MascotasIndex > 0 Then
                Npclist(MascotasIndex).MaestroUser = UserIndex
                Npclist(MascotasIndex).MenuIndex = eMenues.ieMascota
                If (.TammedPets(.SelectedPet).RemainingLife <> 0) Then
                    Npclist(MascotasIndex).Stats.MinHp = .TammedPets(.SelectedPet).RemainingLife
                End If
                
                Call FollowAmo(MascotasIndex)
            Else
                .TammedPets(.SelectedPet).NpcIndex = 0
                Exit Function
            End If
        End If
        
    ' Invocacion normal
    Else
        
        If (.InvokedPetsCount >= Classes(.clase).ClassMods.MaxInvokedPets Or _
           (.InvokedPetsCount + ActiveTammedPetsQty) >= Classes(.clase).ClassMods.MaxActivePets) Then
            If ActiveTammedPetsQty > 0 Then
                Call WriteConsoleMsg(UserIndex, "Has superado la cantidad máxima de mascotas invocadas.", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)
                Exit Function
            Else
                invokedPet = GetOlderInvokedPetIndex(UserIndex)
                If (invokedPet > 0) Then
                    Call QuitarInvocacion(UserIndex, invokedPet)
                End If
            End If
        End If
     
        For NroNpcs = 1 To Hechizos(SpellIndex).cant
            
            If .InvokedPetsCount < Classes(.clase).ClassMods.MaxInvokedPets Then
                NpcIndex = SpawnNpc(Hechizos(SpellIndex).NumNpc, TargetPos, False, False)
                .flags.LastNpcInvoked = NpcIndex
                If NpcIndex > 0 Then
                    .InvokedPetsCount = .InvokedPetsCount + 1
                    
                    PetIndex = FreeInvokedPetIndex(UserIndex)
                    
                    .InvokedPets(PetIndex).NpcIndex = NpcIndex
                    .InvokedPets(PetIndex).NpcNumber = Npclist(NpcIndex).Numero
                    
                    With Npclist(NpcIndex)
                        .MaestroUser = UserIndex
                        .Contadores.TiempoExistencia = SetIntervalEnd(ServerConfiguration.Intervals.IntervaloInvocacion)
                        .GiveGLD = 0
                        .MenuIndex = eMenues.ieMascota
                    End With
                    
                    Call FollowAmo(NpcIndex)
                Else
                    Exit Function
                End If
            Else
                Exit For
            End If
        
        Next NroNpcs
    End If
End With

HechizoInvocacion = True

Exit Function

ErrHandler:
    With UserList(UserIndex)
        LogError ("[" & Err.Number & "] " & Err.Description & " por el usuario " & .Name & "(" & UserIndex & _
                ") en (" & .Pos.Map & ", " & .Pos.X & ", " & .Pos.Y & "). Tratando de tirar el hechizo " & _
                Hechizos(SpellIndex).Nombre & "(" & SpellIndex & ") en la posicion ( " & .flags.TargetX & ", " & .flags.TargetY & ")")
    End With

End Function

Function CastSpellToTerrain(ByVal UserIndex As Integer, ByRef TargetPos As WorldPos, ByVal SpellIndex As Integer) As Boolean
On Error GoTo ErrHandler
    Dim IsAttackSuccessful As Boolean
    
    Select Case Hechizos(SpellIndex).Tipo
        Case TipoHechizo.uInvocacion
            IsAttackSuccessful = HechizoInvocacion(UserIndex, TargetPos, SpellIndex)

    End Select
    
    If IsAttackSuccessful Then
        Call CastSpellToTerrainFX(UserIndex, SpellIndex)
    End If
    
    CastSpellToTerrain = IsAttackSuccessful
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function CastSpellToTerrain de modHechizos.bas")
End Function

Function CastSpellToUser(ByVal UserIndex As Integer, ByVal TargetUser As Integer, ByVal SpellIndex As Integer, ByVal TargetDistance As Byte) As Boolean
On Error GoTo ErrHandler:
    Dim IsAttackSuccessful As Boolean

    With UserList(UserIndex)

        If Not HandleHechizoUsuarioValidate(UserIndex, TargetUser, SpellIndex) Then
            Exit Function
        End If
        
        Select Case Hechizos(SpellIndex).Tipo
            Case TipoHechizo.uEstado
                ' Afectan estados (por ejem : Envenenamiento)
                IsAttackSuccessful = HechizoEstadoUsuario(UserIndex, TargetUser, SpellIndex, TargetDistance)
            
            Case TipoHechizo.uPropiedades
                ' Afectan HP,MANA,STAMINA,ETC
                IsAttackSuccessful = HechizoPropUsuario(UserIndex, TargetUser, SpellIndex, TargetDistance)

            Case TipoHechizo.uInvocacion
                If Hechizos(SpellIndex).Revivir = Pet Then
                    IsAttackSuccessful = RevivePet(UserIndex)
                End If
        End Select

        If IsAttackSuccessful Then
            Call CastSpellToUserFX(UserIndex, TargetUser, SpellIndex)
            
            ' Update the stats of the target
            Call WriteUpdateUserStats(TargetUser)
        End If
               
        CastSpellToUser = IsAttackSuccessful
        
    End With
    
    Exit Function
    
ErrHandler:
    Call LogError("Error en CastSpellToUser. Error " & Err.Number & " : " & Err.Description & _
        " Hechizo: " & Hechizos(SpellIndex).Nombre & "(" & SpellIndex & _
        "). Casteado por: " & UserList(UserIndex).Name & "(" & UserIndex & "). TargetUser: " & UserList(TargetUser).Name & "(" & TargetUser & ")")
End Function

Function CastSpellToNpc(ByVal UserIndex As Integer, ByVal TargetNpc As Integer, ByVal SpellIndex As Integer, ByVal TargetDistance As Byte) As Boolean
On Error GoTo ErrHandler
    
    Dim IsAttackSuccessful As Boolean
    
    If Not HandleHechizoNpcValidate(UserIndex, TargetNpc) Then
        Exit Function
    End If
    Dim TargetKilled As Boolean

    With UserList(UserIndex)
        Select Case Hechizos(SpellIndex).Tipo
            Case TipoHechizo.uEstado
                ' Afectan estados (por ejem : Envenenamiento)
                IsAttackSuccessful = HechizoEstadoNPC(UserIndex, TargetNpc, SpellIndex, TargetDistance)
                
            Case TipoHechizo.uPropiedades
                ' Afectan HP,MANA,STAMINA,ETC
                IsAttackSuccessful = HechizoPropNPC(UserIndex, TargetNpc, SpellIndex, TargetDistance, TargetKilled)
        End Select
        
        If IsAttackSuccessful Then
            ' Bonificación para druidas.
            If .clase = eClass.Druid Then
                ' Será ignorado hasta que pierda el efecto del mimetismo o ataque un npc
                .flags.Ignorado = (Hechizos(SpellIndex).Mimetiza = 1)
            End If
        End If
        
        CastSpellToNpc = IsAttackSuccessful
    End With
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CastSpellToNPC de modHechizos.bas")
End Function

Function CastSpellToObj(ByVal UserIndex As Integer, ByVal TargetObj As Integer, ByVal SpellIndex As Integer) As Boolean
On Error GoTo ErrHandler
      With UserList(UserIndex)
      
        Dim IsAttackSuccessful As Boolean
        
        If Not HandleHechizoObjValidate(UserIndex) Then
            Exit Function
        End If

        IsAttackSuccessful = HechizoEstadoObj(UserIndex, TargetObj, SpellIndex)
            
        If IsAttackSuccessful Then
            ' Bonificación para druidas.
            If .clase = eClass.Druid Then
                ' Será ignorado hasta que pierda el efecto del mimetismo o ataque un npc
                .flags.Ignorado = (Hechizos(SpellIndex).Mimetiza = 1)
            End If
        End If
        
        CastSpellToObj = IsAttackSuccessful
    End With
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CastSpellToObj de modHechizos.bas")
End Function

Public Function HandleHechizoNpcValidate(ByVal UserIndex As Integer, ByVal TargetNpc As Integer) As Boolean
    
    With UserList(UserIndex)
        If Abs(Npclist(TargetNpc).Pos.Y - .Pos.Y) > RANGO_VISION_Y Or Abs(Npclist(TargetNpc).Pos.X - .Pos.X) > RANGO_VISION_X Then
            Exit Function
        End If
        
        
        If TargetNpc = 0 Then
            Call WriteConsoleMsg(UserIndex, "Este hechizo sólo afecta a los npcs.", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)
            Exit Function
        End If
    End With
    
    HandleHechizoNpcValidate = True
    
End Function

Public Function HandleHechizoUsuarioValidate(ByVal UserIndex As Integer, ByVal TargetUser As Integer, ByVal SpellIndex As Integer) As Boolean
    With UserList(UserIndex)
        If Abs(UserList(TargetUser).Pos.Y - .Pos.Y) > RANGO_VISION_Y Or Abs(UserList(TargetUser).Pos.X - .Pos.X) > RANGO_VISION_X Then
            Exit Function
        End If
        
         If IsUserImmuneToSpell(UserIndex, TargetUser, SpellIndex, True) Then Exit Function
    End With
    
    HandleHechizoUsuarioValidate = True

End Function

Public Function HandleHechizoObjValidate(ByVal UserIndex As Integer) As Boolean
    With UserList(UserIndex)
        ' The spell only works with Resources
        If (ObjData(.flags.TargetObj).ObjType <> otResource) Then
            Exit Function
        End If
    End With
    
    HandleHechizoObjValidate = True
End Function

Private Function GetTarget(ByRef Pos As WorldPos, ByVal X As Integer, ByVal Y As Integer, Optional ByVal DistanceFromTarget As Byte = 0) As tSpellPosition
    Dim Target As tSpellPosition
    Target.Pos = Pos
    Target.Pos.X = Target.Pos.X + X
    Target.Pos.Y = Target.Pos.Y + Y
    Target.DistanceFromTarget = DistanceFromTarget
    
    GetTarget = Target
End Function

Public Function GetAttackVector(ByVal UserIndex As Integer, ByVal SpellIndex As Integer) As tSpellPosition()
'***************************************************
' Generate a vector of different attack positions
' if the weapon can generate a splash damage
'***************************************************
    Dim AttackVector() As tSpellPosition
    Dim Pos As WorldPos

    Dim Spell As tHechizo
    Dim Area As Byte
    Spell = Hechizos(SpellIndex)
    Area = Spell.Area
    
    With UserList(UserIndex)
        Pos.Map = .Pos.Map
        Pos.X = .flags.TargetX
        Pos.Y = .flags.TargetY
        
        If Area Then
            Dim I As Integer
            Dim k As Integer
            Dim J As Integer
            Dim TargetedTiles As Integer
            
            ' Tiles affected by Spell
            TargetedTiles = (2 + 2 * Area) * Area + 1
            ReDim AttackVector(1 To TargetedTiles) As tSpellPosition
            
            AttackVector(1).Pos = Pos

            For I = Area To 1 Step -1
                AttackVector(TargetedTiles) = GetTarget(Pos, 0, I, I)
                AttackVector(TargetedTiles - 1) = GetTarget(Pos, 0, -I, I)
                AttackVector(TargetedTiles - 2) = GetTarget(Pos, I, 0, I)
                AttackVector(TargetedTiles - 3) = GetTarget(Pos, -I, 0, I)
                
                TargetedTiles = TargetedTiles - 4
                
                For k = 1 To I
                    If k < I Then
                        J = I - k
                        AttackVector(TargetedTiles) = GetTarget(Pos, k, J, I)
                        AttackVector(TargetedTiles - 1) = GetTarget(Pos, k, -J, I)
                        AttackVector(TargetedTiles - 2) = GetTarget(Pos, -k, J, I)
                        AttackVector(TargetedTiles - 3) = GetTarget(Pos, -k, -J, I)
                        
                        TargetedTiles = TargetedTiles - 4
                    End If
                Next k
            Next I
            
        Else
            ReDim AttackVector(1 To 1) As tSpellPosition
            
            Dim WP As MapBlock
            Dim wp2 As MapBlock
            
            WP = MapData(Pos.Map, Pos.X, Pos.Y)
            wp2 = MapData(Pos.Map, Pos.X, Pos.Y + 1) ' Character Hitbox is [x,y, x,y+1]
            
            ' No area damage, so we are going to hit only to the base target
            AttackVector(1).Pos = Pos
            
            If wp2.UserIndex > 0 Or wp2.NpcIndex > 0 Then
                AttackVector(1) = GetTarget(Pos, 0, 1)
            End If
            
        End If
    End With
       
    GetAttackVector = AttackVector
End Function

Sub LanzarHechizo(ByVal SpellNumber As Integer, ByVal UserIndex As Integer)
On Error GoTo ErrHandler

    Dim Spell As tHechizo
    
    With UserList(UserIndex)
        If PuedeLanzar(UserIndex, SpellNumber) Then
            
            Spell = Hechizos(SpellNumber)
            
             'Substract energy anyways because no valid target was found
            Call SubstractStamina(UserIndex, Spell.StaRequerido, False)
            
            Dim AttackVector() As tSpellPosition
            Dim TargetPos As MapBlock
            Dim TargetDistance As Byte
            AttackVector = GetAttackVector(UserIndex, SpellNumber)
            
            Dim I As Integer
            Dim IsAttackSuccessful As Boolean
            
            Dim CasterIsGM As Boolean
            Dim TargetIsGM As Boolean
            Dim UserProtected As Boolean
            
            CasterIsGM = EsGm(UserIndex)
            
            For I = 1 To UBound(AttackVector)
                TargetPos = MapData(AttackVector(I).Pos.Map, AttackVector(I).Pos.X, AttackVector(I).Pos.Y)
                TargetDistance = AttackVector(I).DistanceFromTarget
                
                If TargetPos.UserIndex > 0 And Spell.TargetUser Then
                    TargetIsGM = EsGm(TargetPos.UserIndex)
                    UserProtected = (.Id = UserList(TargetPos.UserIndex).Id And Not Hechizos(SpellNumber).CasterAffected)

                    If CasterIsGM Or (Not CasterIsGM And (Not TargetIsGM Or (TargetIsGM And UserList(TargetPos.UserIndex).flags.AdminInvisible = False))) Then
                        If Not UserProtected Then
                            If CastSpellToUser(UserIndex, TargetPos.UserIndex, SpellNumber, TargetDistance) Then
                                IsAttackSuccessful = True
                            End If
                        End If
                    End If
               
                ElseIf TargetPos.NpcIndex > 0 And Spell.TargetNpc Then
                    If CastSpellToNpc(UserIndex, TargetPos.NpcIndex, SpellNumber, TargetDistance) Then
                        IsAttackSuccessful = True
                    End If
                    
                ElseIf TargetPos.ObjInfo.ObjIndex > 0 And Spell.TargetObj Then
                    If CastSpellToObj(UserIndex, TargetPos.ObjInfo.ObjIndex, SpellNumber) Then
                        IsAttackSuccessful = True
                    End If

                ElseIf Spell.TargetTerrain Then
                    If CastSpellToTerrain(UserIndex, AttackVector(I).Pos, SpellNumber) Then
                        IsAttackSuccessful = True
                    End If
                    
                End If
            Next I
            
            .Stats.UserHechizos(.flags.CastedSpellIndex).LastUsedSuccessfully = IsAttackSuccessful
                    
            If IsAttackSuccessful Then
                Dim ManaRequerida As Long
            
                ManaRequerida = GetSpellRequiredMana(SpellNumber, UserIndex)
                Call SubirSkill(UserIndex, eSkill.Magia, True)
                
                If Hechizos(SpellNumber).RequireFullMana = 1 Then
                    ManaRequerida = .Stats.MaxMan
                End If

                ' Quito la mana requerida
                .Stats.MinMAN = .Stats.MinMAN - ManaRequerida
                If .Stats.MinMAN < 0 Then .Stats.MinMAN = 0
                
                .flags.TargetNpc = 0
                .flags.TargetUser = 0
                
                Call DecirPalabrasMagicas(Hechizos(SpellNumber).PalabrasMagicas, UserIndex)
            'Else
                'Call WriteConsoleMsg(UserIndex, "Tu hechizo no tuvo efecto.", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)
            End If
            
            Call WriteSpellAttackResult(UserIndex, IsAttackSuccessful, UserList(UserIndex).flags.CastedSpellIndex)
            
            Call WriteUpdateUserStats(UserIndex)
                    
        End If
    End With

Exit Sub

ErrHandler:
    Call LogError("Error en LanzarHechizo. Error " & Err.Number & " : " & Err.Description & _
        " Hechizo: " & Hechizos(SpellNumber).Nombre & "(" & SpellNumber & _
        "). Casteado por: " & UserList(UserIndex).Name & "(" & UserIndex & ").")
    
End Sub

Function HechizoEstadoUsuario(ByVal UserIndex As Integer, ByVal TargetIndex As Integer, ByVal SpellIndex As Integer, ByVal TargetDistance As Byte) As Boolean
On Error GoTo ErrHandler:

With UserList(UserIndex)
    Dim AnilloObjIndex As Integer
    AnilloObjIndex = UserList(TargetIndex).Invent.AnilloEqpObjIndex

    ' <-------- Agrega Invisibilidad ---------->
    If Hechizos(SpellIndex).Invisibilidad = 1 Then
        If UserList(TargetIndex).flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡El usuario está muerto!", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)

            Exit Function
        End If
        
        If UserList(TargetIndex).flags.invisible = 1 Then
            If UserIndex <> TargetIndex Then
                Call WriteConsoleMsg(UserIndex, "¡El usuario ya se encuentra invisible!", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)
            Else
                Call WriteConsoleMsg(UserIndex, "¡Ya eres invisible!", FontTypeNames.FONTTYPE_WARNING)
            End If

            Exit Function
        End If
        
        If UserList(TargetIndex).Counters.Saliendo Then
            If UserIndex <> TargetIndex Then
                Call WriteConsoleMsg(UserIndex, "¡El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)
            Else
                Call WriteConsoleMsg(UserIndex, "¡No puedes hacerte invisible mientras te encuentras saliendo!", FontTypeNames.FONTTYPE_WARNING)
            End If

            Exit Function
        End If
        
        'No usar invi mapas InviSinEfecto
        If MapInfo(UserList(TargetIndex).Pos.Map).InviSinEfecto > 0 Then
            Call WriteConsoleMsg(UserIndex, "¡La invisibilidad no funciona aquí!", FontTypeNames.FONTTYPE_INFO)

            Exit Function
        End If
        
        'No usar invi desafios si no lo permite
        If .Challenge.InSand > 0 Then
            If .Pos.Map = SandsChallenge(.Challenge.InSand).Event_map And SandsChallenge(.Challenge.InSand).Invisibility > 0 Then
                Call WriteConsoleMsg(UserIndex, "¡La invisibilidad no funciona aquí!", FontTypeNames.FONTTYPE_INFO)
    
                Exit Function
            End If
        End If
        
        ' Chequea si el status permite ayudar al otro usuario
        If Not CanSupportUser(UserIndex, TargetIndex, True) Then Exit Function
        
        'Si sos user, no uses este hechizo con GMS.
        If .flags.Privilegios And PlayerType.User Then
            If Not UserList(TargetIndex).flags.Privilegios And PlayerType.User Then
    
                Exit Function
            End If
        End If
        
        ' Disable Berserk of the target if it has it enabled
        If HasPassiveActivated(TargetIndex, ePassiveSpells.Berserk) Then
            Call ActivatePassive(TargetIndex, ePassiveSpells.Berserk, False)
            Call modPassiveSkills.SendBerserkEffect(TargetIndex, ePassiveSpells.Berserk, False)
        End If
       
        UserList(TargetIndex).flags.invisible = 1
        
        ' Adds the duration here instead of in the timer. If the user has a mastery enabled then we add that to the minimum and maximum duration
        UserList(TargetIndex).Counters.Invisibilidad = SetIntervalEnd(CalculateAreaEfficacy(ServerConfiguration.Intervals.IntervaloInvisible + RandomNumber(-100, 100) + RandomNumber(.Masteries.Boosts.AddInviMinDuration, .Masteries.Boosts.AddInviMaxDuration), SpellIndex, TargetDistance))
         
        ' Solo se hace invi para los clientes si no esta navegando
        If UserList(TargetIndex).flags.Navegando = 0 Then
            Call UsUaRiOs.SetInvisible(TargetIndex, UserList(TargetIndex).Char.CharIndex, True)
        End If
                
    End If
    
    ' <-------- Remueven Invisibilidad ---------->
    If Hechizos(SpellIndex).RemueveInvisibilidadParcial = 1 Then
        If UserList(TargetIndex).flags.invisible = False Or UserList(TargetIndex).flags.AdminInvisible = 1 Then
            Exit Function
        End If
    End If
    
    ' <-------- Agrega Mimetismo ---------->
    If Hechizos(SpellIndex).Mimetiza = 1 Then
        If Not DoMimetizar(UserIndex, TargetIndex, eTargetType.ieUser) Then
            Exit Function
        End If
    End If
    
    ' <-------- Agrega Envenenamiento ---------->
    If Hechizos(SpellIndex).Envenena = 1 Then
        If UserIndex = TargetIndex Then
            Call WriteConsoleMsg(UserIndex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
            Exit Function
        End If
        

        If FriendlyFireProtectionEnabled(UserIndex, TargetIndex) And TriggerZonaPelea(UserIndex, TargetIndex) <> TRIGGER6_PERMITE Then
            Call WriteConsoleMsg(UserIndex, "No puedes atacar a un miembro de tu mismo Clan.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
            Exit Function
        End If

        If Not PuedeAtacar(UserIndex, TargetIndex) Then Exit Function
        If UserIndex <> TargetIndex Then
            Call UsuarioAtacadoPorUsuario(UserIndex, TargetIndex)
        End If
        UserList(TargetIndex).flags.Envenenado = 1
        
        
    End If
    
    ' <-------- Cura Envenenamiento ---------->
    If Hechizos(SpellIndex).CuraVeneno = 1 Then
    
        'Verificamos que el usuario no este muerto
        If UserList(TargetIndex).flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡El usuario está muerto!", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)

            Exit Function
        End If
        
        ' Chequea si el status permite ayudar al otro usuario
        If Not CanSupportUser(UserIndex, TargetIndex) Then Exit Function
            
        'Si sos user, no uses este hechizo con GMS.
        If .flags.Privilegios And PlayerType.User Then
            If Not UserList(TargetIndex).flags.Privilegios And PlayerType.User Then
                Exit Function
            End If
        End If
            
        UserList(TargetIndex).flags.Envenenado = 0
        
    End If
    
    ' <-------- Agrega Maldicion ---------->
    If Hechizos(SpellIndex).Maldicion = 1 Then
        If UserIndex = TargetIndex Then
            Call WriteConsoleMsg(UserIndex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
            Exit Function
        End If
        
        If Not PuedeAtacar(UserIndex, TargetIndex) Then Exit Function
        If FriendlyFireProtectionEnabled(UserIndex, TargetIndex) And TriggerZonaPelea(UserIndex, TargetIndex) <> TRIGGER6_PERMITE Then
            Call WriteConsoleMsg(UserIndex, "No puedes atacar a un miembro de tu mismo Clan.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
            Exit Function
        End If

        If UserIndex <> TargetIndex Then
            Call UsuarioAtacadoPorUsuario(UserIndex, TargetIndex)
        End If
        UserList(TargetIndex).flags.Maldicion = 1
        
    End If
    
    ' <-------- Remueve Maldicion ---------->
    If Hechizos(SpellIndex).RemoverMaldicion = 1 Then
            UserList(TargetIndex).flags.Maldicion = 0
    End If
    
    ' <-------- Agrega Bendicion ---------->
    If Hechizos(SpellIndex).Bendicion = 1 Then
            UserList(TargetIndex).flags.Bendicion = 1
    End If
    
    ' <-------- Agrega Paralisis/Inmobilidad ---------->
    If Hechizos(SpellIndex).Paraliza = 1 Or Hechizos(SpellIndex).Inmoviliza = 1 Then

        If UserIndex = TargetIndex Then
            Call WriteConsoleMsg(UserIndex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
            Exit Function
        End If
                
        If UserList(TargetIndex).flags.Paralizado Or UserList(TargetIndex).flags.Inmovilizado Then Exit Function
        If Not PuedeAtacar(UserIndex, TargetIndex) Then Exit Function
        
        If MapInfo(UserList(UserIndex).Pos.Map).InmovilizarSinEfecto = 1 Then
            Call WriteConsoleMsg(UserIndex, "No puedes lanzar este hechizo en el mapa.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
            Exit Function
        End If
            
        If UserIndex <> TargetIndex Then
            Call UsuarioAtacadoPorUsuario(UserIndex, TargetIndex)
        End If

        If FriendlyFireProtectionEnabled(UserIndex, TargetIndex) And TriggerZonaPelea(UserIndex, TargetIndex) <> TRIGGER6_PERMITE Then
            Call WriteConsoleMsg(UserIndex, "No puedes atacar a un miembro de tu mismo Clan.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
            Exit Function
        End If
        
        If Hechizos(SpellIndex).Paraliza = 1 Then
            If PassiveConditionMet(TargetIndex, ParalysisImmunity) Then
                Call WriteConsoleMsg(UserIndex, "El usuario es inmune a tu hechizo.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                Call WriteConsoleMsg(TargetIndex, "Tu habilidades te protegen del hechizo paralizante de " & UserList(UserIndex).Name & ".", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                Exit Function
            End If
            
            If AnilloObjIndex > 0 Then
                If ObjData(AnilloObjIndex).ImpideParalizar <> 0 Then
                    Call WriteConsoleMsg(TargetIndex, "Tu anillo rechaza los efectos de la paralisis.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                    Call WriteConsoleMsg(UserIndex, " ¡El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                    Exit Function
                End If
            End If
        End If
        
        If Hechizos(SpellIndex).Inmoviliza = 1 Or Hechizos(SpellIndex).Paraliza = 1 Then
            If Hechizos(SpellIndex).ByPassPassive = 0 Then
                If HasPassiveActivated(TargetIndex, ePassiveSpells.Berserk) Then
                    Call WriteConsoleMsg(TargetIndex, "Tu habilidades te protegen del hechizo " & IIf(Hechizos(SpellIndex).Inmoviliza = 1, "inmovilizante", "paralizante") & " de " & UserList(UserIndex).Name & ".", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                    Call WriteConsoleMsg(UserIndex, " ¡El usuario es inmune a tu hechizo!", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                    Exit Function
                End If
            End If
        End If
        
        
        If Hechizos(SpellIndex).Inmoviliza = 1 Then
            If AnilloObjIndex > 0 Then
                If ObjData(AnilloObjIndex).ImpideInmobilizar <> 0 Then
                    Call WriteConsoleMsg(UserIndex, "Tu anillo rechaza los efectos del hechizo inmobilizar.", FontTypeNames.FONTTYPE_FIGHT)
                    Call WriteConsoleMsg(UserIndex, "El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_FIGHT)
                    Exit Function
                End If
            End If
            
            UserList(TargetIndex).flags.Inmovilizado = 1
        End If
        
        UserList(TargetIndex).flags.Paralizado = 1
        UserList(TargetIndex).Counters.Paralisis = SetIntervalEnd(CalculateAreaEfficacy(ServerConfiguration.Intervals.IntervaloParalizado, SpellIndex, TargetDistance))
        
        UserList(TargetIndex).flags.ParalizedByIndex = UserIndex
        UserList(TargetIndex).flags.ParalizedBy = UserList(UserIndex).Name
        
        Call WriteParalizeOK(TargetIndex)
    End If
    
    ' <-------- Remueve Paralisis/Inmobilidad ---------->
    If Hechizos(SpellIndex).RemoverParalisis = 1 Then
        If UserList(TargetIndex).flags.Paralizado = 0 Then Exit Function
        
        ' Remueve si esta paralizado
        ' Chequea si el status permite ayudar al otro usuario
        If Not CanSupportUser(UserIndex, TargetIndex, True) Then Exit Function
        Call RemoveParalisis(TargetIndex)

    End If
    
    ' <-------- Remueve Estupidez (Aturdimiento) ---------->
    If Hechizos(SpellIndex).RemoverEstupidez = 1 Then
    
        ' Remueve si esta en ese estado
        If UserList(TargetIndex).flags.Estupidez = 0 Then Exit Function
        
        ' Chequea si el status permite ayudar al otro usuario
        If Not CanSupportUser(UserIndex, TargetIndex) Then Exit Function
    
        UserList(TargetIndex).flags.Estupidez = 0
    
        'no need to crypt this
        Call WriteDumbNoMore(TargetIndex)
    End If
    
    ' <-------- Revive ---------->
    If Hechizos(SpellIndex).Revivir = eReviveTarget.User Then
    
        If UserList(TargetIndex).flags.Muerto = 0 Then Exit Function
            
        'Seguro de resurreccion (solo afecta a los hechizos, no al sacerdote ni al comando de GM)
        If UserList(TargetIndex).flags.SeguroResu Then
            Call WriteConsoleMsg(UserIndex, "¡El espíritu no tiene intenciones de regresar al mundo de los vivos!", FontTypeNames.FONTTYPE_INFO)

            Exit Function
        End If
    
        'No usar resu en mapas con ResuSinEfecto
        If MapInfo(UserList(TargetIndex).Pos.Map).ResuSinEfecto > 0 Then
            Call WriteConsoleMsg(UserIndex, "¡Revivir no está permitido aquí! Retirate de la Zona si deseas utilizar el Hechizo.", FontTypeNames.FONTTYPE_INFO)

            Exit Function
        End If
        
        If EnMapaDuelos(UserIndex) Then
            If UserList(UserIndex).flags.DueloIndex > 0 Then
                If Not DuelData.Duelo(UserList(UserIndex).flags.DueloIndex).Resucitar Then
                    Call WriteConsoleMsg(UserIndex, "¡No está permitido resucitar en este duelo!", FontTypeNames.FONTTYPE_INFO)
        
                    Exit Function
                End If
            End If
        End If
         
        ' Chequea si el status permite ayudar al otro usuario
        If Not CanSupportUser(UserIndex, TargetIndex, True) Then Exit Function
        
        'Agregado para quitar la penalización de vida en el ring y cambio de ecuacion. (NicoNZ)
        If (TriggerZonaPelea(UserIndex, TargetIndex) <> TRIGGER6_PERMITE) Then
            'Solo saco vida si es User. no quiero que exploten GMs por ahi.
            If .flags.Privilegios And PlayerType.User Then
                .Stats.MinHp = .Stats.MinHp * (1 - UserList(TargetIndex).Stats.ELV * 0.015)
            End If
        End If
        
        If (.Stats.MinHp <= 0) Then
            Call UserDie(UserIndex)
            Call WriteConsoleMsg(UserIndex, "El esfuerzo de resucitar fue demasiado grande.", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, "El esfuerzo de resucitar te ha debilitado.", FontTypeNames.FONTTYPE_INFO)
        End If
        
        If UserList(TargetIndex).flags.Traveling = 1 Then
            Call EndTravel(TargetIndex, True)
        End If
        
        Call RevivirUsuario(TargetIndex, True)
    End If
    
    ' <-------- Agrega Ceguera ---------->
    If Hechizos(SpellIndex).Ceguera = 1 Then
        If UserIndex = TargetIndex Then
            Call WriteConsoleMsg(UserIndex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
            Exit Function
        End If
        
        If FriendlyFireProtectionEnabled(UserIndex, TargetIndex) And TriggerZonaPelea(UserIndex, TargetIndex) <> TRIGGER6_PERMITE Then
            Call WriteConsoleMsg(UserIndex, "No puedes atacar a un miembro de tu mismo Clan.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
            Exit Function
        End If
        
        If Not PuedeAtacar(UserIndex, TargetIndex) Then Exit Function
        
        If UserIndex <> TargetIndex Then
            Call UsuarioAtacadoPorUsuario(UserIndex, TargetIndex)
        End If
                
        If AnilloObjIndex > 0 Then
            If ObjData(AnilloObjIndex).ImpideCegar <> 0 Then
                Call WriteConsoleMsg(UserIndex, "Tu anillo rechaza los efectos de la ceguera.", FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(UserIndex, " ¡El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_FIGHT)
                Exit Function
            End If
        End If

        UserList(TargetIndex).flags.Ceguera = 1
        
        UserList(TargetIndex).Counters.Ceguera = SetIntervalEnd(CalculateAreaEfficacy(ServerConfiguration.Intervals.IntervaloParalizado / 3, SpellIndex, TargetDistance))

        Call WriteBlind(TargetIndex)
        
    End If
    
    ' <-------- Agrega Estupidez (Aturdimiento) ---------->
    If Hechizos(SpellIndex).Estupidez = 1 Then
    
        If UserIndex = TargetIndex Then
            Call WriteConsoleMsg(UserIndex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
            Exit Function
        End If
        
        If Not PuedeAtacar(UserIndex, TargetIndex) Then Exit Function
        
        If FriendlyFireProtectionEnabled(UserIndex, TargetIndex) And TriggerZonaPelea(UserIndex, TargetIndex) <> TRIGGER6_PERMITE Then
            Call WriteConsoleMsg(UserIndex, "No puedes atacar a un miembro de tu mismo Clan.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
            Exit Function
        End If
        
        If UserIndex <> TargetIndex Then
            Call UsuarioAtacadoPorUsuario(UserIndex, TargetIndex)
        End If

        If AnilloObjIndex > 0 Then
            If ObjData(AnilloObjIndex).ImpideAturdir <> 0 Then
                Call WriteConsoleMsg(UserIndex, "Tu anillo rechaza los efectos de la turbación.", FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(UserIndex, " ¡El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_FIGHT)
                Exit Function
            End If
        End If
        
        If PassiveConditionMet(TargetIndex, IndomitableWill) Then

            Call WriteConsoleMsg(TargetIndex, "Tu habilidades te protegen del hechizo aturdidor de " & UserList(UserIndex).Name & ".", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
            Call WriteConsoleMsg(UserIndex, " ¡El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_FIGHT)
            Exit Function
        End If

        
        If UserList(TargetIndex).flags.Estupidez = 0 Then
            UserList(TargetIndex).flags.Estupidez = 1
            UserList(TargetIndex).Counters.Ceguera = SetIntervalEnd(CalculateAreaEfficacy(ServerConfiguration.Intervals.IntervaloParalizado, SpellIndex, TargetDistance))
        End If
        
        Call WriteDumb(TargetIndex)
     
    End If
End With

HechizoEstadoUsuario = True

Exit Function
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HechizoEstadoUsuario de modHechizos.bas")

End Function

Function HechizoEstadoNPC(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal SpellIndex As Integer, ByVal TargetDistance As Byte) As Boolean
On Error GoTo ErrHandler

With Npclist(NpcIndex)
    If Hechizos(SpellIndex).Invisibilidad = 1 Then
        .flags.invisible = 1
    End If
    
    If Hechizos(SpellIndex).Envenena = 1 Then
        If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
            Exit Function
        End If
        Call NPCAtacado(NpcIndex, UserIndex)
        
        .flags.Envenenado = 1
    End If
    
    If Hechizos(SpellIndex).CuraVeneno = 1 Then
        .flags.Envenenado = 0
    End If
    
    If Hechizos(SpellIndex).Maldicion = 1 Then
        If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
            Exit Function
        End If
        
        Call NPCAtacado(NpcIndex, UserIndex)
        .flags.Maldicion = 1
    End If
    
    If Hechizos(SpellIndex).RemoverMaldicion = 1 Then
        .flags.Maldicion = 0
    End If
    
    If Hechizos(SpellIndex).Bendicion = 1 Then
        .flags.Bendicion = 1
    End If
    
    If Hechizos(SpellIndex).Paraliza = 1 Then
        If .flags.AfectaParalisis = 0 Then
            If Not PuedeAtacarNPC(UserIndex, NpcIndex, True) Then
                Exit Function
            End If
            Call NPCAtacado(NpcIndex, UserIndex)
            
            .flags.Paralizado = 1
            .flags.Inmovilizado = 0
            .Contadores.Paralisis = SetIntervalEnd(CalculateAreaEfficacy(ServerConfiguration.Intervals.IntervaloNPCParalizado, SpellIndex, TargetDistance))
        Else
            Call WriteConsoleMsg(UserIndex, "El NPC es inmune a este hechizo.", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)
            Exit Function
        End If
    End If
    
    If Hechizos(SpellIndex).RemoverParalisis = 1 Then
        If .flags.Paralizado = 1 Or .flags.Inmovilizado = 1 Then
            If .MaestroUser = UserIndex Then
                .flags.Paralizado = 0
                .Contadores.Paralisis = 0
            Else
                If .NPCtype = eNPCType.GuardiaReal Then
                    If esArmada(UserIndex) Then
                        .flags.Paralizado = 0
                        .Contadores.Paralisis = 0
                        Exit Function
                    Else
                        Call WriteConsoleMsg(UserIndex, "Sólo puedes remover la parálisis de los Guardias si perteneces a su facción.", FontTypeNames.FONTTYPE_INFO)
                        Exit Function
                    End If

                    Call WriteConsoleMsg(UserIndex, "Solo puedes remover la parálisis de los NPCs que te consideren su amo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Function
                Else
                    If .NPCtype = eNPCType.GuardiasCaos Then
                        If esCaos(UserIndex) Then
                            .flags.Paralizado = 0
                            .Contadores.Paralisis = 0
                            Exit Function
                        Else
                            Call WriteConsoleMsg(UserIndex, "Solo puedes remover la parálisis de los Guardias si perteneces a su facción.", FontTypeNames.FONTTYPE_INFO)
                            Exit Function
                        End If
                    End If
                End If
            End If
       Else
          Call WriteConsoleMsg(UserIndex, "Este NPC no está paralizado", FontTypeNames.FONTTYPE_INFO)
          Exit Function
       End If
    End If
     
    If Hechizos(SpellIndex).Inmoviliza = 1 Then
        If MapInfo(UserList(UserIndex).Pos.Map).InmovilizarSinEfecto = 1 Then
            Call WriteConsoleMsg(UserIndex, "No puedes lanzar este hechizo en el mapa.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
            Exit Function
        End If
        
        If .flags.AfectaParalisis = 0 Then
            If Not PuedeAtacarNPC(UserIndex, NpcIndex, True) Then
                Exit Function
            End If
            Call NPCAtacado(NpcIndex, UserIndex)
            .flags.Inmovilizado = 1
            .flags.Paralizado = 0
            .Contadores.Paralisis = SetIntervalEnd(CalculateAreaEfficacy(ServerConfiguration.Intervals.IntervaloNPCParalizado, SpellIndex, TargetDistance))
        Else
            Call WriteConsoleMsg(UserIndex, "El NPC es inmune al hechizo.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
    End If
    
    If Hechizos(SpellIndex).Mimetiza = 1 Then
        If Not DoMimetizar(UserIndex, NpcIndex, eTargetType.ieNpc) Then
            Exit Function
        End If
    End If
End With

Call CastSpellToNpcFX(UserIndex, NpcIndex, SpellIndex)

HechizoEstadoNPC = True
    
Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HechizoEstadoNPC de modHechizos.bas")
End Function

Function HechizoPropNPC(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal SpellIndex As Integer, ByVal TargetDistance As Byte, Optional ByRef TargetKilled As Boolean) As Boolean
On Error GoTo ErrHandler

Dim Damage As Long

With Npclist(NpcIndex)
    'Salud
    If Hechizos(SpellIndex).SubeHP = 1 Then
        If CanSupportNpc(UserIndex, NpcIndex) Then
            Damage = CalculateSpellDamageForUserHP(UserIndex, SpellIndex, TargetDistance)
            
            .Stats.MinHp = .Stats.MinHp + Damage
            If .Stats.MinHp > .Stats.MaxHp Then _
                .Stats.MinHp = .Stats.MaxHp
            Call WriteConsoleMsg(UserIndex, "Has curado " & Damage & " puntos de vida a la criatura.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
        End If
        
    ElseIf Hechizos(SpellIndex).SubeHP = 2 Then
        If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
            Exit Function
        End If
        Call NPCAtacado(NpcIndex, UserIndex)
        
        Damage = CalculateSpellDamageForUserHP(UserIndex, SpellIndex, TargetDistance)
                
        If .flags.Snd2 > 0 Then
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(.flags.Snd2, .Pos.X, .Pos.Y, .Char.CharIndex))
        End If
        
        'Quizas tenga defenza magica el NPC. Pablo (ToxicWaste)
        Damage = Max(1, Damage - .Stats.DefM)
                    
        .Stats.MinHp = .Stats.MinHp - Damage
        
        Call WriteConsoleMsg(UserIndex, "¡Le has quitado " & Damage & " puntos de vida a la criatura!", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
        
        Call LifeLeechEffect(ieUser, UserIndex, Damage, Hechizos(SpellIndex).LifeLeechPerc)
        
        Call CalcularDarExp(UserIndex, NpcIndex, Damage)
        
        Call CheckPets(NpcIndex, UserIndex, False)
        
        Call CastSpellToNpcFX(UserIndex, NpcIndex, SpellIndex)
    
        If .Stats.MinHp < 1 Then
            .Stats.MinHp = 0
            Call MuereNpc(NpcIndex, UserIndex)
            TargetKilled = True
        Else
            If .ExtraBodies > 0 Then
                If .ActualBody < .ExtraBodies Then
                    If .Stats.MinHp < .Stats.MaxHp * ((.ExtraBodies - .ActualBody) / (.ExtraBodies + 1)) Then
                        .ActualBody = .ActualBody + 1
                        .Char.body = .ExtraBody(.ActualBody)
                        Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCharacterChange(.ExtraBody(.ActualBody), _
                                        .Char.head, .Char.heading, .Char.CharIndex, .Char.WeaponAnim, .Char.ShieldAnim, _
                                        .Char.FX, .Char.Loops, .Char.CascoAnim, CBool(Npclist(NpcIndex).flags.TierraInvalida = 1), False, 0, eCharacterAlignment.Neutral))
                    End If
                End If
            End If
        End If
    End If
End With

HechizoPropNPC = True
Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HechizoPropNPC de modHechizos.bas")
End Function

Function HechizoEstadoObj(ByVal UserIndex As Integer, ObjIndex As Integer, ByVal SpellIndex As Integer) As Boolean
On Error GoTo ErrHandler
  
    If Hechizos(SpellIndex).Mimetiza = 1 Then
        If Not DoMimetizar(UserIndex, ObjIndex, eTargetType.ieObject) Then
            Exit Function
        End If
    End If

  HechizoEstadoObj = True
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HechizoEstadoObj de modHechizos.bas")
End Function

Sub CastSpellToUserFX(ByVal UserIndex As Integer, TargetIndex As Integer, ByVal SpellIndex As Integer)
    With UserList(UserIndex)
        If .flags.AdminInvisible = 1 And UserIndex = TargetIndex Then
            ' Los admins invisibles no producen sonidos ni fx's
            Call SendData(ToUser, UserIndex, PrepareMessageCreateFX(UserList(TargetIndex).Char.CharIndex, Hechizos(SpellIndex).FXgrh, Hechizos(SpellIndex).Loops))
            Call SendData(ToUser, UserIndex, PrepareMessagePlayWave(Hechizos(SpellIndex).WAV, UserList(TargetIndex).Pos.X, UserList(TargetIndex).Pos.Y, UserList(TargetIndex).Char.CharIndex))
        Else
            Call SendData(SendTarget.ToPCArea, TargetIndex, PrepareMessageCreateFX(UserList(TargetIndex).Char.CharIndex, Hechizos(SpellIndex).FXgrh, Hechizos(SpellIndex).Loops))
            Call SendData(SendTarget.ToPCArea, TargetIndex, PrepareMessagePlayWave(Hechizos(SpellIndex).WAV, UserList(TargetIndex).Pos.X, UserList(TargetIndex).Pos.Y, UserList(TargetIndex).Char.CharIndex))
        End If
        
        If UserIndex <> TargetIndex Then
            If .ShowName Then
                Call WriteConsoleMsg(UserIndex, Hechizos(SpellIndex).HechizeroMsg & " " & UserList(TargetIndex).Name, FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
            Else
                Call WriteConsoleMsg(UserIndex, Hechizos(SpellIndex).HechizeroMsg & " alguien.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
            End If
            Call WriteConsoleMsg(TargetIndex, UserList(UserIndex).Name & " " & Hechizos(SpellIndex).TargetMsg, FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
        Else
            Call WriteConsoleMsg(UserIndex, Hechizos(SpellIndex).PropioMsg, FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
        End If
    End With
End Sub

Sub CastSpellToNpcFX(ByVal UserIndex As Integer, TargetIndex As Integer, ByVal SpellIndex As Integer)
    With UserList(UserIndex)
        Call SendData(SendTarget.ToNPCArea, TargetIndex, PrepareMessageCreateFX(Npclist(TargetIndex).Char.CharIndex, Hechizos(SpellIndex).FXgrh, Hechizos(SpellIndex).Loops))
        Call SendData(SendTarget.ToNPCArea, TargetIndex, PrepareMessagePlayWave(Hechizos(SpellIndex).WAV, Npclist(TargetIndex).Pos.X, Npclist(TargetIndex).Pos.Y, Npclist(TargetIndex).Char.CharIndex))
        
        Call WriteConsoleMsg(UserIndex, Hechizos(SpellIndex).HechizeroMsg & " " & "la criatura.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
    End With
End Sub

Sub CastSpellToTerrainFX(ByVal UserIndex As Integer, ByVal SpellIndex As Integer)
    Dim TargetIndex As Integer
    
    With UserList(UserIndex)
        ' Si estamos clieckeando sobre el suelo, el hechizo es de invocación y el último npc invocado es > 0
        If Hechizos(SpellIndex).Tipo = uInvocacion And UserList(UserIndex).flags.LastNpcInvoked > 0 Then
            TargetIndex = UserList(UserIndex).flags.LastNpcInvoked
            
            Call SendData(SendTarget.ToNPCArea, TargetIndex, PrepareMessageCreateFX(Npclist(TargetIndex).Char.CharIndex, Hechizos(SpellIndex).FXgrh, Hechizos(SpellIndex).Loops))
            Call SendData(SendTarget.ToNPCArea, TargetIndex, PrepareMessagePlayWave(Hechizos(SpellIndex).WAV, Npclist(TargetIndex).Pos.X, Npclist(TargetIndex).Pos.Y, Npclist(TargetIndex).Char.CharIndex))
            
            Call WriteConsoleMsg(UserIndex, Hechizos(SpellIndex).HechizeroMsg, FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
        
        End If
    End With
End Sub

Sub CastSpellToObjFX(ByVal UserIndex As Integer, ByVal SpellIndex As Integer)
    With UserList(UserIndex)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, Hechizos(SpellIndex).FXgrh, Hechizos(SpellIndex).Loops))
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(SpellIndex).WAV, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, UserList(UserIndex).Char.CharIndex))
        
        Call WriteConsoleMsg(UserIndex, Hechizos(SpellIndex).HechizeroMsg & " " & "el objeto.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
    End With
End Sub

Public Function HechizoPropUsuario(ByVal UserIndex As Integer, ByVal TargetIndex As Integer, ByVal SpellIndex As Integer, ByVal TargetDistance As Byte) As Boolean
On Error GoTo ErrHandler

Dim Damage As Long

With UserList(TargetIndex)
    If .flags.Muerto Then
        Call WriteConsoleMsg(UserIndex, "No puedes lanzar este hechizo a un muerto.", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)
        Exit Function
    End If
              
    ' <-------- Aumenta Hambre ---------->
    If Hechizos(SpellIndex).SubeHam = 1 Then
        ' Chequea si el status permite ayudar al otro usuario
        If Not CanSupportUser(UserIndex, TargetIndex) Then Exit Function
    
        Damage = CalculateSpellDamageForUser(UserIndex, SpellIndex, Hechizos(SpellIndex).MinHam, Hechizos(SpellIndex).MaxHam, TargetDistance)

        .Stats.MinHam = .Stats.MinHam + Damage
        If .Stats.MinHam > .Stats.MaxHam Then _
            .Stats.MinHam = .Stats.MaxHam
        
        If UserIndex <> TargetIndex Then
            Call WriteConsoleMsg(UserIndex, "Le has restaurado " & Damage & " puntos de hambre a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
            Call WriteConsoleMsg(TargetIndex, UserList(UserIndex).Name & " te ha restaurado " & Damage & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
        Else
            Call WriteConsoleMsg(UserIndex, "Te has restaurado " & Damage & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
        End If
        
        Call WriteUpdateHungerAndThirst(TargetIndex)
    
    ' <-------- Quita Hambre ---------->
    ElseIf Hechizos(SpellIndex).SubeHam = 2 Then
        If Not PuedeAtacar(UserIndex, TargetIndex) Then Exit Function
        
        If UserIndex <> TargetIndex Then
            Call UsuarioAtacadoPorUsuario(UserIndex, TargetIndex)
        Else
            Exit Function
        End If
        
        Damage = CalculateSpellDamageForUser(UserIndex, SpellIndex, Hechizos(SpellIndex).MinHam, Hechizos(SpellIndex).MaxHam, TargetDistance)
        
        .Stats.MinHam = .Stats.MinHam - Damage
        
        If UserIndex <> TargetIndex Then
            Call WriteConsoleMsg(UserIndex, "Le has quitado " & Damage & " puntos de hambre a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
            Call WriteConsoleMsg(TargetIndex, UserList(UserIndex).Name & " te ha quitado " & Damage & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
        Else
            Call WriteConsoleMsg(UserIndex, "Te has quitado " & Damage & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
        End If
        
        If .Stats.MinHam < 1 Then
            .Stats.MinHam = 0
            .flags.Hambre = 1
        End If
        
        Call WriteUpdateHungerAndThirst(TargetIndex)
    End If
    
    ' <-------- Aumenta Sed ---------->
    If Hechizos(SpellIndex).SubeSed = 1 Then
        ' Chequea si el status permite ayudar al otro usuario
        If Not CanSupportUser(UserIndex, TargetIndex) Then Exit Function
        
        Damage = CalculateSpellDamageForUser(UserIndex, SpellIndex, Hechizos(SpellIndex).MinSed, Hechizos(SpellIndex).MaxSed, TargetDistance)
        
        .Stats.MinAGU = .Stats.MinAGU + Damage
        If .Stats.MinAGU > .Stats.MaxAGU Then _
            .Stats.MinAGU = .Stats.MaxAGU
        
        Call WriteUpdateHungerAndThirst(TargetIndex)
             
        If UserIndex <> TargetIndex Then
          Call WriteConsoleMsg(UserIndex, "Le has restaurado " & Damage & " puntos de sed a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
          Call WriteConsoleMsg(TargetIndex, UserList(UserIndex).Name & " te ha restaurado " & Damage & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
        Else
          Call WriteConsoleMsg(UserIndex, "Te has restaurado " & Damage & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
        End If
        
    
    ' <-------- Quita Sed ---------->
    ElseIf Hechizos(SpellIndex).SubeSed = 2 Then
        
        If Not PuedeAtacar(UserIndex, TargetIndex) Then Exit Function
        
        If UserIndex <> TargetIndex Then
            Call UsuarioAtacadoPorUsuario(UserIndex, TargetIndex)
        End If
                
        Damage = CalculateSpellDamageForUser(UserIndex, SpellIndex, Hechizos(SpellIndex).MinSed, Hechizos(SpellIndex).MaxSed, TargetDistance)
        
        .Stats.MinAGU = .Stats.MinAGU - Damage
        
        If UserIndex <> TargetIndex Then
            Call WriteConsoleMsg(UserIndex, "Le has quitado " & Damage & " puntos de sed a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
            Call WriteConsoleMsg(TargetIndex, UserList(UserIndex).Name & " te ha quitado " & Damage & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
        Else
            Call WriteConsoleMsg(UserIndex, "Te has quitado " & Damage & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
        End If
        
        If .Stats.MinAGU < 1 Then
            .Stats.MinAGU = 0
            .flags.Sed = 1
        End If
        
        Call WriteUpdateHungerAndThirst(TargetIndex)
        
    End If
    
    ' <-------- Aumenta Agilidad ---------->
    If Hechizos(SpellIndex).SubeAgilidad = 1 Then
        ' Chequea si el status permite ayudar al otro usuario
        If Not CanSupportUser(UserIndex, TargetIndex) Then Exit Function
        
        ' Chequea si el status permite ayudar al otro usuario
        If Not CanSupportUser(UserIndex, TargetIndex) Then Exit Function
        
        Damage = CalculateSpellDamageForUser(UserIndex, SpellIndex, Hechizos(SpellIndex).MinAgilidad, Hechizos(SpellIndex).MaxAgilidad, TargetDistance)
        Damage = .Stats.UserAtributos(eAtributos.Agilidad) + Damage
        
        If Damage > MinimoInt(ConstantesBalance.MaxAtributos, .Stats.UserAtributosBackUP(Agilidad) * 2) Then
            Damage = MinimoInt(ConstantesBalance.MaxAtributos, .Stats.UserAtributosBackUP(Agilidad) * 2)
        End If
        
        .Stats.UserAtributos(eAtributos.Agilidad) = Damage

        .flags.DuracionEfecto = 1200

        .flags.TomoPocion = True
        Call WriteUpdateDexterity(TargetIndex)
        
        ' Enable Berserk
        If HasPassiveAssigned(TargetIndex, ePassiveSpells.Berserk) And HasPassiveActivated(TargetIndex, ePassiveSpells.Berserk) = False Then
            If BerzerkConditionMet(TargetIndex) Then
                Call ActivatePassive(TargetIndex, ePassiveSpells.Berserk)
            End If
        End If
        
        ' Enable Indomitable Will
        If HasPassiveAssigned(TargetIndex, IndomitableWill) Then
            If PassiveConditionMet(TargetIndex, IndomitableWill) Then
                Call ActivatePassive(TargetIndex, IndomitableWill)
            End If
        End If
        
        ' Enable Indomitable Will
        If HasPassiveAssigned(TargetIndex, ParalysisImmunity) Then
            If PassiveConditionMet(TargetIndex, ParalysisImmunity) Then
                Call ActivatePassive(TargetIndex, ParalysisImmunity)
            End If
        End If
        
    ' <-------- Quita Agilidad ---------->
    ElseIf Hechizos(SpellIndex).SubeAgilidad = 2 Then
        If Not PuedeAtacar(UserIndex, TargetIndex) Then Exit Function
        
        If UserIndex <> TargetIndex Then
            Call UsuarioAtacadoPorUsuario(UserIndex, TargetIndex)
        End If
                
        .flags.TomoPocion = True
        Damage = CalculateSpellDamageForUser(UserIndex, SpellIndex, Hechizos(SpellIndex).MinAgilidad, Hechizos(SpellIndex).MaxAgilidad, TargetDistance)
        
        .flags.DuracionEfecto = 700
        .Stats.UserAtributos(eAtributos.Agilidad) = MaximoInt(.Stats.UserAtributos(eAtributos.Agilidad) - Damage, ConstantesBalance.MinAtributos)
        
        Call WriteUpdateDexterity(TargetIndex)
        
        ' Disable Berserk
        If HasPassiveAssigned(TargetIndex, ePassiveSpells.Berserk) Then
            If HasPassiveActivated(TargetIndex, ePassiveSpells.Berserk) Then
                Call ActivatePassive(TargetIndex, ePassiveSpells.Berserk, False)
                Call SendBerserkEffect(TargetIndex, ePassiveSpells.Berserk, False)
            End If
        End If
        
        ' Disable the Indomitable Will
        'If HasPassiveActivated(TargetIndex, IndomitableWill) Then
            'Call ActivatePassive(TargetIndex, ePassiveSpells.Berserk, False)
        'End If
        
    End If
    
    ' <-------- Aumenta Fuerza ---------->
    If Hechizos(SpellIndex).SubeFuerza = 1 Then
        ' Chequea si el status permite ayudar al otro usuario
        If Not CanSupportUser(UserIndex, TargetIndex) Then Exit Function
        
        Damage = CalculateSpellDamageForUser(UserIndex, SpellIndex, Hechizos(SpellIndex).MinFuerza, Hechizos(SpellIndex).MaxFuerza, TargetDistance)
        Damage = .Stats.UserAtributos(eAtributos.Fuerza) + Damage
        
        If Damage > MinimoInt(ConstantesBalance.MaxAtributos, .Stats.UserAtributosBackUP(Fuerza) * 2) Then
            Damage = MinimoInt(ConstantesBalance.MaxAtributos, .Stats.UserAtributosBackUP(Fuerza) * 2)
        End If
        
        .Stats.UserAtributos(eAtributos.Fuerza) = Damage
        
        .flags.DuracionEfecto = 1200

        .flags.TomoPocion = True
        Call WriteUpdateStrenght(TargetIndex)
        
        ' Enable Berserk
        If HasPassiveAssigned(TargetIndex, ePassiveSpells.Berserk) And HasPassiveActivated(TargetIndex, ePassiveSpells.Berserk) = False Then
            If BerzerkConditionMet(TargetIndex) Then
                Call ActivatePassive(TargetIndex, ePassiveSpells.Berserk)
            End If
        End If
        
        ' Enable Indomitable Will
        If HasPassiveAssigned(TargetIndex, IndomitableWill) Then
            If PassiveConditionMet(TargetIndex, IndomitableWill) Then
                Call ActivatePassive(TargetIndex, IndomitableWill)
            End If
        End If
        
        ' Enable Indomitable Will
        If HasPassiveAssigned(TargetIndex, ParalysisImmunity) Then
            If PassiveConditionMet(TargetIndex, ParalysisImmunity) Then
                Call ActivatePassive(TargetIndex, ParalysisImmunity)
            End If
        End If
        
    ' <-------- Quita Fuerza ---------->
    ElseIf Hechizos(SpellIndex).SubeFuerza = 2 Then
        
        If FriendlyFireProtectionEnabled(UserIndex, TargetIndex) And TriggerZonaPelea(UserIndex, TargetIndex) <> TRIGGER6_PERMITE Then
            Call WriteConsoleMsg(UserIndex, "No puedes atacar a un miembro de tu mismo Clan.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
            Exit Function
        End If
    
        If Not PuedeAtacar(UserIndex, TargetIndex) Then Exit Function
        
        If UserIndex <> TargetIndex Then
            Call UsuarioAtacadoPorUsuario(UserIndex, TargetIndex)
        End If
                
        .flags.TomoPocion = True
        
        Damage = CalculateSpellDamageForUser(UserIndex, SpellIndex, Hechizos(SpellIndex).MinFuerza, Hechizos(SpellIndex).MaxFuerza, TargetDistance)
        
        .flags.DuracionEfecto = 700
        .Stats.UserAtributos(eAtributos.Fuerza) = MaximoInt(.Stats.UserAtributos(eAtributos.Fuerza) - Damage, ConstantesBalance.MinAtributos)
                
        Call WriteUpdateStrenght(TargetIndex)
        
        ' Disable the Berserk
        If HasPassiveAssigned(TargetIndex, ePassiveSpells.Berserk) Then
            If HasPassiveActivated(TargetIndex, ePassiveSpells.Berserk) Then
                Call ActivatePassive(TargetIndex, ePassiveSpells.Berserk, False)
                Call SendBerserkEffect(TargetIndex, ePassiveSpells.Berserk, False)
            End If
        End If
        
        ' Disable the Indomitable Will
        'If HasPassiveActivated(TargetIndex, IndomitableWill) Then
            'Call EnablePassive(TargetIndex, ePassiveSpells.Berserk, False)
        'End If
        
    End If
    
    ' <-------- Cura salud ---------->
    If Hechizos(SpellIndex).SubeHP = 1 Then
        
        'Verifica que el usuario no este muerto
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡El usuario está muerto!", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        
        ' Chequea si el status permite ayudar al otro usuario
        If Not CanSupportUser(UserIndex, TargetIndex) Then Exit Function
           
        Damage = CalculateSpellDamageForUserHP(UserIndex, SpellIndex, TargetDistance)
            
        .Stats.MinHp = .Stats.MinHp + Damage
        If .Stats.MinHp > .Stats.MaxHp Then _
            .Stats.MinHp = .Stats.MaxHp
        
        Call WriteUpdateHP(TargetIndex)
        
        If UserIndex <> TargetIndex Then
            Call WriteConsoleMsg(UserIndex, "Le has restaurado " & Damage & " puntos de vida a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
            Call WriteConsoleMsg(TargetIndex, UserList(UserIndex).Name & " te ha restaurado " & Damage & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
        Else
            Call WriteConsoleMsg(UserIndex, "Te has restaurado " & Damage & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
        End If
        
    ' <-------- Quita salud (Daña) ---------->
    ElseIf Hechizos(SpellIndex).SubeHP = 2 Then
        Dim DamageLeech As Integer
        
        If UserIndex = TargetIndex Then
            Call WriteConsoleMsg(UserIndex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
            Exit Function
        End If
        
        If FriendlyFireProtectionEnabled(UserIndex, TargetIndex) And TriggerZonaPelea(UserIndex, TargetIndex) <> TRIGGER6_PERMITE Then
            Call WriteConsoleMsg(UserIndex, "No puedes atacar a un miembro de tu mismo Clan.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
            Exit Function
        End If
        
        If Not PuedeAtacar(UserIndex, TargetIndex) Then Exit Function
        
        ' Calculate damage
        Damage = CalculateSpellDamageForUserHP(UserIndex, SpellIndex, TargetDistance)

        ' Substract from the damage based on the magic defense from
        If (.Invent.CascoEqpObjIndex > 0) Then
            Damage = Damage - RandomNumber(ObjData(.Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(.Invent.CascoEqpObjIndex).DefensaMagicaMax)
        End If
        
        'anillos
        If (.Invent.AnilloEqpObjIndex > 0) Then
            Damage = Damage - RandomNumber(ObjData(.Invent.AnilloEqpObjIndex).DefensaMagicaMin, ObjData(.Invent.AnilloEqpObjIndex).DefensaMagicaMax)
        End If
        
        ' barcos
        If .flags.Navegando = 1 And .Invent.BarcoObjIndex > 0 Then
            Damage = Damage - RandomNumber(ObjData(.Invent.BarcoObjIndex).DefensaMagicaMin, ObjData(.Invent.BarcoObjIndex).DefensaMagicaMax)
        End If
        
        Damage = Max(1, Damage)
             
        .Stats.MinHp = .Stats.MinHp - Damage
        
        Call WriteUpdateHP(TargetIndex)
        
        Call WriteConsoleMsg(UserIndex, "Le has quitado " & Damage & " puntos de vida a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
        Call WriteConsoleMsg(TargetIndex, UserList(UserIndex).Name & " te ha quitado " & Damage & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
        
        ' Calculate Life Leech
        Call LifeLeechEffect(ieUser, UserIndex, Damage, Hechizos(SpellIndex).LifeLeechPerc)
        
        ' Calculate the mana conversion of the target
        Call ManaConversionEffect(TargetIndex, Damage)
            
        'Muere
        If .Stats.MinHp < 1 Then
        
            If .flags.AtacablePor <> UserIndex Then
                'Store it!
                Call Statistics.StoreFrag(UserIndex, TargetIndex)
                ' TODO: NIGHTW, Fix the "Challenge mode"
                Call ContarMuerte(TargetIndex, UserIndex, eDamageType.Spell, Damage, SpellIndex)
            End If
            
            .Stats.MinHp = 0
            
            Call ActStats(TargetIndex, UserIndex)
            Call UserDie(TargetIndex)
            
            Call AllFollowAmo(UserIndex)
        End If
        
        If UserIndex <> TargetIndex Then
            Call UsuarioAtacadoPorUsuario(UserIndex, TargetIndex)
        End If
        
    End If
    
    ' <-------- Aumenta Mana ---------->
    If Hechizos(SpellIndex).SubeMana = 1 Then
        ' Chequea si el status permite ayudar al otro usuario
        If Not CanSupportUser(UserIndex, TargetIndex) Then Exit Function
        Damage = CalculateSpellDamageForUser(UserIndex, SpellIndex, Hechizos(SpellIndex).MinMana, Hechizos(SpellIndex).MaxMana, TargetDistance)
        
        .Stats.MinMAN = .Stats.MinMAN + Damage
        If .Stats.MinMAN > .Stats.MaxMan Then _
            .Stats.MinMAN = .Stats.MaxMan
        
        Call WriteUpdateMana(TargetIndex)
        
        If UserIndex <> TargetIndex Then
            Call WriteConsoleMsg(UserIndex, "Le has restaurado " & Damage & " puntos de maná a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
            Call WriteConsoleMsg(TargetIndex, UserList(UserIndex).Name & " te ha restaurado " & Damage & " puntos de maná.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
        Else
            Call WriteConsoleMsg(UserIndex, "Te has restaurado " & Damage & " puntos de maná.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
        End If
        
    ' <-------- Quita Mana ---------->
    ElseIf Hechizos(SpellIndex).SubeMana = 2 Then
        If Not PuedeAtacar(UserIndex, TargetIndex) Then Exit Function
        
        Damage = CalculateSpellDamageForUser(UserIndex, SpellIndex, Hechizos(SpellIndex).MinMana, Hechizos(SpellIndex).MaxMana, TargetDistance)
        
        If UserIndex <> TargetIndex Then
            Call UsuarioAtacadoPorUsuario(UserIndex, TargetIndex)
        End If
                
        If UserIndex <> TargetIndex Then
            Call WriteConsoleMsg(UserIndex, "Le has quitado " & Damage & " puntos de maná a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
            Call WriteConsoleMsg(TargetIndex, UserList(UserIndex).Name & " te ha quitado " & Damage & " puntos de maná.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
        Else
            Call WriteConsoleMsg(UserIndex, "Te has quitado " & Damage & " puntos de maná.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
        End If
        
        .Stats.MinMAN = .Stats.MinMAN - Damage
        If .Stats.MinMAN < 1 Then .Stats.MinMAN = 0
        
        Call WriteUpdateMana(TargetIndex)
        
    End If
    
    ' <-------- Aumenta Stamina ---------->
    If Hechizos(SpellIndex).SubeSta = 1 Then
        ' Chequea si el status permite ayudar al otro usuario
        If Not CanSupportUser(UserIndex, TargetIndex) Then Exit Function
        Damage = CalculateSpellDamageForUser(UserIndex, SpellIndex, Hechizos(SpellIndex).MinSta, Hechizos(SpellIndex).MaxSta, TargetDistance)
    
        .Stats.MinSta = .Stats.MinSta + Damage
        If .Stats.MinSta > .Stats.MaxSta Then _
            .Stats.MinSta = .Stats.MaxSta
        
        Call WriteUpdateSta(TargetIndex)
        
        If UserIndex <> TargetIndex Then
            Call WriteConsoleMsg(UserIndex, "Le has restaurado " & Damage & " puntos de energía a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
            Call WriteConsoleMsg(TargetIndex, UserList(UserIndex).Name & " te ha restaurado " & Damage & " puntos de energía.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
        Else
            Call WriteConsoleMsg(UserIndex, "Te has restaurado " & Damage & " puntos de energía.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
        End If
        
    ' <-------- Quita Stamina ---------->
    ElseIf Hechizos(SpellIndex).SubeSta = 2 Then
        If Not PuedeAtacar(UserIndex, TargetIndex) Then Exit Function
        
        Damage = CalculateSpellDamageForUser(UserIndex, SpellIndex, Hechizos(SpellIndex).MinSta, Hechizos(SpellIndex).MaxSta, TargetDistance)
        
        If UserIndex <> TargetIndex Then
            Call UsuarioAtacadoPorUsuario(UserIndex, TargetIndex)
        End If
                
        If UserIndex <> TargetIndex Then
            Call WriteConsoleMsg(UserIndex, "Le has quitado " & Damage & " puntos de energía a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
            Call WriteConsoleMsg(TargetIndex, UserList(UserIndex).Name & " te ha quitado " & Damage & " puntos de energía.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
        Else
            Call WriteConsoleMsg(UserIndex, "Te has quitado " & Damage & " puntos de energía.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
        End If
        
        .Stats.MinSta = .Stats.MinSta - Damage
        
        If .Stats.MinSta < 1 Then .Stats.MinSta = 0
        
        Call WriteUpdateSta(TargetIndex)
        
    End If
End With
    
    HechizoPropUsuario = True
    
    Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function HechizoPropUsuario de modHechizos.bas. UserIndex: " & UserIndex & ", TargetIndex: " & TargetIndex & ", SpellIndex: " & SpellIndex & ", TargetDistance: " & TargetDistance)
End Function

Public Function CanSupportUser(ByVal CasterIndex As Integer, ByVal TargetIndex As Integer, _
                               Optional ByVal DoCriminal As Boolean = False) As Boolean
'***************************************************
'Author: ZaMa
'Last Modification: 28/04/2010
'Checks if caster can cast support magic on target user.
'***************************************************
     
 On Error GoTo ErrHandler
 
    With UserList(CasterIndex)
        
        ' Te podes curar a vos mismo
        If CasterIndex = TargetIndex Then
            CanSupportUser = True
            Exit Function
        End If
        
         ' No podes ayudar si estas en consulta
        If .flags.HelpMode Then
            Call WriteConsoleMsg(CasterIndex, "No puedes ayudar usuarios mientras estas en consulta.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        
        ' Si estas en la arena, esta todo permitido
        If TriggerZonaPelea(CasterIndex, TargetIndex) = TRIGGER6_PERMITE Then
            CanSupportUser = True
            Exit Function
        End If
     
        If Not CanHelpByAlignment(CasterIndex, TargetIndex) Then
            Call WriteConsoleMsg(CasterIndex, "No puedes ayudar a usuarios de esa alineación.", FontTypeNames.FONTTYPE_INFO)
            CanSupportUser = False
            Exit Function
        End If
        
    End With
    
    CanSupportUser = True

    Exit Function
    
ErrHandler:
    Call LogError("Error en CanSupportUser, Error: " & Err.Number & " - " & Err.Description & _
                  " CasterIndex: " & CasterIndex & ", TargetIndex: " & TargetIndex)

End Function

Sub UpdateUserHechizos(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal Slot As Byte)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

Dim LoopC As Byte

With UserList(UserIndex)
    'Actualiza un solo slot
    If Not UpdateAll Then
        'Actualiza el inventario
        If .Stats.UserHechizos(Slot).SpellNumber > 0 Then
            Call ChangeUserHechizo(UserIndex, Slot, .Stats.UserHechizos(Slot).SpellNumber)
        Else
            Call ChangeUserHechizo(UserIndex, Slot, 0)
        End If
    Else
        'Actualiza todos los slots
        For LoopC = 1 To MAXUSERHECHIZOS
            'Actualiza el inventario
            If .Stats.UserHechizos(LoopC).SpellNumber > 0 Then
                Call ChangeUserHechizo(UserIndex, LoopC, .Stats.UserHechizos(LoopC).SpellNumber)
            Else
                Call ChangeUserHechizo(UserIndex, LoopC, 0)
            End If
        Next LoopC
    End If
End With

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub UpdateUserHechizos de modHechizos.bas")
End Sub

Public Function CanSupportNpc(ByVal CasterIndex As Integer, ByVal TargetIndex As Integer) As Boolean
'***************************************************
'Author: ZaMa
'Last Modification: 18/09/2010
'Checks if caster can cast support magic on target Npc.
'***************************************************
     
 On Error GoTo ErrHandler
 
    Dim OwnerIndex As Integer
 
    With UserList(CasterIndex)
        
        OwnerIndex = Npclist(TargetIndex).Owner
        
        ' Si no tiene dueño puede
        If OwnerIndex = 0 Then
            CanSupportNpc = True
            Exit Function
        End If
        
        ' Puede hacerlo si es su propio npc
        If CasterIndex = OwnerIndex Then
            CanSupportNpc = True
            Exit Function
        End If
        
         ' No podes ayudar si estas en consulta
        If .flags.HelpMode Then
            Call WriteConsoleMsg(CasterIndex, "No puedes ayudar npcs mientras estas en consulta.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        
        ' Si estas en la arena, esta todo permitido
        If TriggerZonaPelea(CasterIndex, OwnerIndex) = TRIGGER6_PERMITE Then
            CanSupportNpc = True
            Exit Function
        End If
        
        ' If the character is not neutral, it can't help NPCs fighting against people from the same faction.
        ' If it's neutral, it can do whatever they want
        If .Faccion.Alignment <> eCharacterAlignment.Neutral And .Faccion.Alignment = UserList(OwnerIndex).Faccion.Alignment Then
            Call WriteConsoleMsg(CasterIndex, "No puedes ayudar npcs que están luchando contra un miembro de tu facción.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If

    End With
    
    CanSupportNpc = True

    Exit Function
    
ErrHandler:
    Call LogError("Error en CanSupportNpc, Error: " & Err.Number & " - " & Err.Description & _
                  " CasterIndex: " & CasterIndex & ", OwnerIndex: " & OwnerIndex)

End Function
Sub ChangeUserHechizo(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Hechizo As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  
    
    UserList(UserIndex).Stats.UserHechizos(Slot).SpellNumber = Hechizo
    
    If Hechizo > 0 And Hechizo < NumeroHechizos + 1 Then
        Call WriteChangeSpellSlot(UserIndex, Slot)
    Else
        Call WriteChangeSpellSlot(UserIndex, Slot)
    End If

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ChangeUserHechizo de modHechizos.bas")
End Sub


Public Sub DesplazarHechizo(ByVal UserIndex As Integer, ByVal Dire As Integer, ByVal HechizoDesplazado As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

If (Dire <> 1 And Dire <> -1) Then Exit Sub
If Not (HechizoDesplazado >= 1 And HechizoDesplazado <= MAXUSERHECHIZOS) Then Exit Sub

Dim TempHechizo As Integer
Dim TempLastUsedAt As Double
Dim LastUsedSuccessfully As Boolean

With UserList(UserIndex)
    If Dire = 1 Then 'Mover arriba
        If HechizoDesplazado = 1 Then
            Call WriteConsoleMsg(UserIndex, "No puedes mover el hechizo en esa dirección.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        Else
            TempHechizo = .Stats.UserHechizos(HechizoDesplazado).SpellNumber
            TempLastUsedAt = .Stats.UserHechizos(HechizoDesplazado).LastUsedAt
            LastUsedSuccessfully = .Stats.UserHechizos(HechizoDesplazado).LastUsedSuccessfully
            
            .Stats.UserHechizos(HechizoDesplazado).SpellNumber = .Stats.UserHechizos(HechizoDesplazado - 1).SpellNumber
            .Stats.UserHechizos(HechizoDesplazado).LastUsedAt = .Stats.UserHechizos(HechizoDesplazado - 1).LastUsedAt
            .Stats.UserHechizos(HechizoDesplazado).LastUsedSuccessfully = .Stats.UserHechizos(HechizoDesplazado - 1).LastUsedSuccessfully
            
            
            .Stats.UserHechizos(HechizoDesplazado - 1).SpellNumber = TempHechizo
            .Stats.UserHechizos(HechizoDesplazado - 1).LastUsedAt = TempLastUsedAt
            .Stats.UserHechizos(HechizoDesplazado - 1).LastUsedSuccessfully = LastUsedSuccessfully
        End If
    Else 'mover abajo
        If HechizoDesplazado = MAXUSERHECHIZOS Then
            Call WriteConsoleMsg(UserIndex, "No puedes mover el hechizo en esa dirección.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        Else
            TempHechizo = .Stats.UserHechizos(HechizoDesplazado).SpellNumber
            TempLastUsedAt = .Stats.UserHechizos(HechizoDesplazado).LastUsedAt
            LastUsedSuccessfully = .Stats.UserHechizos(HechizoDesplazado).LastUsedSuccessfully
            
            .Stats.UserHechizos(HechizoDesplazado).SpellNumber = .Stats.UserHechizos(HechizoDesplazado + 1).SpellNumber
            .Stats.UserHechizos(HechizoDesplazado).LastUsedAt = .Stats.UserHechizos(HechizoDesplazado + 1).LastUsedAt
            .Stats.UserHechizos(HechizoDesplazado).LastUsedSuccessfully = .Stats.UserHechizos(HechizoDesplazado + 1).LastUsedSuccessfully
            
            .Stats.UserHechizos(HechizoDesplazado + 1).SpellNumber = TempHechizo
            .Stats.UserHechizos(HechizoDesplazado + 1).LastUsedAt = TempLastUsedAt
            .Stats.UserHechizos(HechizoDesplazado + 1).LastUsedSuccessfully = LastUsedSuccessfully
        End If
    End If
End With

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub DesplazarHechizo de modHechizos.bas")
End Sub

Public Function DoMimetizar(ByVal UserIndex As Integer, ByVal TargetIndex As Integer, _
    ByVal TargetType As Byte) As Boolean
On Error GoTo ErrHandler
  
    
    ' If already mimetized can't do it again
    If UserList(UserIndex).flags.Mimetizado <> 0 Then Exit Function
    
    Dim IsDruid As Boolean
    IsDruid = (UserList(UserIndex).clase = eClass.Druid)
    
    Select Case TargetType
        ' Users validations
        Case eTargetType.ieUser
            ' Not with Admins
            If EsGm(TargetIndex) Then
                ' Only if visible..
                If UserList(TargetIndex).flags.AdminInvisible = 0 Then
                    WriteConsoleMsg UserIndex, "¡No puedes mimetizarte con administradores!", FontTypeNames.FONTTYPE_INFO, eMessageType.info
                End If
                
                Exit Function
        
            ' Only with visible chars
            ElseIf UserList(TargetIndex).flags.invisible = 1 Then
                WriteConsoleMsg UserIndex, "¡No puedes mimetizarte con usuarios invisibles!", FontTypeNames.FONTTYPE_INFO, eMessageType.info
                Exit Function
                
            ' Only with living chars
            ElseIf UserList(TargetIndex).flags.Muerto = 1 Then
                ' Except for druids..
                If Not IsDruid Then
                    WriteConsoleMsg UserIndex, "¡No puedes mimetizarte con usuarios muertos!", FontTypeNames.FONTTYPE_INFO, eMessageType.info
                    Exit Function
                End If
            End If
                
            ' Both have to be either above water or ground.
            If (UserList(UserIndex).flags.Navegando Xor UserList(TargetIndex).flags.Navegando) <> 0 Then
                
                If UserList(TargetIndex).flags.Navegando = 1 Then
                    WriteConsoleMsg UserIndex, "¡No puedes mimetizarte con usuarios que navegan mientras no navegas!", FontTypeNames.FONTTYPE_INFO, eMessageType.info
                Else
                    WriteConsoleMsg UserIndex, "¡No puedes mimetizarte con usuarios que no navegan mientras navegas!", FontTypeNames.FONTTYPE_INFO, eMessageType.info
                End If
                
                Exit Function
                
            ' Can't do it with npc-mimetized users
            ElseIf UserList(TargetIndex).flags.Mimetizado = 2 Then
                ' Except for druids..
                If Not IsDruid Then
                    WriteConsoleMsg UserIndex, "¡No puedes mimetizarte con usuarios mimetizados de criaturas o árboles!", FontTypeNames.FONTTYPE_INFO, eMessageType.info
                    Exit Function
                End If
            End If
            
        ' Npc validations
        Case eTargetType.ieNpc
            ' Only allowed to druids
            If Not IsDruid Then
                WriteConsoleMsg UserIndex, "¡No puedes mimetizarte con criaturas!", FontTypeNames.FONTTYPE_INFO, eMessageType.info
                Exit Function
                
            ' Some restrictions to no-hostile npcs
            ElseIf Npclist(TargetIndex).Hostile = 0 Then
                ' Can't mimetize with gobernators or nobles
                If Npclist(TargetIndex).NPCtype = eNPCType.Gobernador Or Npclist(TargetIndex).NPCtype = eNPCType.Noble Then
                    WriteConsoleMsg UserIndex, "¡No puedes mimetizarte con este personaje!", FontTypeNames.FONTTYPE_INFO, eMessageType.info
                    Exit Function
                End If
            End If
            
            ' if sailing, npc must be an aquatic type (or amphibian)
            If UserList(UserIndex).flags.Navegando = 1 And Npclist(TargetIndex).flags.AguaValida = 0 Then
                WriteConsoleMsg UserIndex, "¡No puedes mimetizarte con criaturas terrestres mientras navegas!", FontTypeNames.FONTTYPE_INFO, eMessageType.info
                Exit Function
                
            ' if not sailing, npc must be a land type (or amphibian)
            ElseIf UserList(UserIndex).flags.Navegando = 0 And Npclist(TargetIndex).flags.TierraInvalida = 1 Then
                WriteConsoleMsg UserIndex, "¡No puedes mimetizarte con criaturas acuáticas mientras no navegas!", FontTypeNames.FONTTYPE_INFO, eMessageType.info
                Exit Function
            End If
            
        ' Objects validations
        Case eTargetType.ieObject
            ' Only allowed to druids
            If Not IsDruid Then
                WriteConsoleMsg UserIndex, "¡No puedes mimetizarte con objetos!", FontTypeNames.FONTTYPE_INFO, eMessageType.info
                Exit Function
            
            ' Not allowed while sailing.
            ElseIf UserList(UserIndex).flags.Navegando = 1 Then
                Call WriteConsoleMsg(UserIndex, "¡Estás navegando!", FontTypeNames.FONTTYPE_INFO)
                Exit Function
            
            ' Only with trees
            ElseIf (ObjData(TargetIndex).ObjType <> otResource) Then
                WriteConsoleMsg UserIndex, "¡No puedes mimetizarte con ese objeto!", FontTypeNames.FONTTYPE_INFO, eMessageType.info
                Exit Function
            End If
    End Select
    
    ' Mimetize
    With UserList(UserIndex)
        ' Save original char
        If .flags.Navegando = 0 Then
            .OrigChar.body = .Char.body
            .OrigChar.head = .Char.head
            .OrigChar.CascoAnim = .Char.CascoAnim
            .OrigChar.ShieldAnim = .Char.ShieldAnim
            .OrigChar.WeaponAnim = .Char.WeaponAnim
        End If
        
        ' User
        If TargetType = eTargetType.ieUser Then
            ' Copy
            .Char.body = UserList(TargetIndex).Char.body
            .Char.head = UserList(TargetIndex).Char.head
            .Char.CascoAnim = UserList(TargetIndex).Char.CascoAnim
            .Char.ShieldAnim = UserList(TargetIndex).Char.ShieldAnim
            .Char.WeaponAnim = UserList(TargetIndex).Char.WeaponAnim
            
            .flags.Mimetizado = 1
            
            ' Type
            If .flags.Navegando = 1 Then
                .flags.MimetizadoType = eMimeType.ieAquatic
            Else
                .flags.MimetizadoType = eMimeType.ieTerrain
            End If
            
        ' Npc
        ElseIf TargetType = eTargetType.ieNpc Then
        
            .Char.body = Npclist(TargetIndex).Char.body
            .Char.head = Npclist(TargetIndex).Char.head
            .Char.CascoAnim = Npclist(TargetIndex).Char.CascoAnim
            .Char.ShieldAnim = Npclist(TargetIndex).Char.ShieldAnim
            .Char.WeaponAnim = Npclist(TargetIndex).Char.WeaponAnim
            .ShowName = False
            Call RefreshCharStatus(UserIndex, False)
            
            .flags.Mimetizado = 2
            
            ' Type
            If (Npclist(TargetIndex).flags.AguaValida = 1) Then
                If (Npclist(TargetIndex).flags.TierraInvalida = 1) Then
                    .flags.MimetizadoType = eMimeType.ieAquatic
                Else
                    .flags.MimetizadoType = eMimeType.ieBoth
                End If
            Else
                .flags.MimetizadoType = eMimeType.ieTerrain
            End If
            
        ' Object
        Else
            .Char.body = Not ObjData(TargetIndex).GrhIndex
            .Char.head = 0
            .Char.CascoAnim = ConstantesGRH.NingunCasco
            .Char.ShieldAnim = ConstantesGRH.NingunEscudo
            .Char.WeaponAnim = ConstantesGRH.NingunArma
            
            .flags.Mimetizado = 2
        End If
        
        Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)

        .Counters.Mimetismo = SetIntervalEnd(ServerConfiguration.Intervals.IntervaloMimetismo)
        
    End With
    
    DoMimetizar = True
    
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function DoMimetizar de modHechizos.bas")
End Function


Private Sub SubstractStamina(ByVal UserIndex As Integer, ByVal QtyToSubstract As Integer, ByVal WriteTobuffer As Boolean)

    With UserList(UserIndex)
    
        If .Stats.MinSta < QtyToSubstract Then
            .Stats.MinSta = 0
        Else
            .Stats.MinSta = .Stats.MinSta - QtyToSubstract
        End If
        
        ' Update the stats of the caster.
        If WriteTobuffer Then Call WriteUpdateUserStats(UserIndex)
        
    End With
End Sub

Private Function CalculateAreaEfficacy(ByVal Value As Long, ByVal SpellIndex As Integer, ByVal TargetDistance As Byte) As Long
     If TargetDistance > 0 Then
            CalculateAreaEfficacy = Porcentaje(value, Hechizos(SpellIndex).AreaEfficacy(TargetDistance))
            Exit Function
    End If
    CalculateAreaEfficacy = value
End Function

Private Function CalculateSpellDamageForUserHP(ByVal UserIndex As Integer, ByVal SpellIndex As Integer, Optional ByVal TargetDistance As Byte = 0) As Integer
    Dim Damage As Integer
    Dim DamageBoost As Integer

    With UserList(UserIndex)
    
        ' Add the magic damage boost from the weapon modifier
        If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then DamageBoost = DamageBoost + ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).MagicDamageBonus
        
        ' Add the magic damage boost from the accessory modifier
        If UserList(UserIndex).Invent.AnilloEqpObjIndex > 0 Then DamageBoost = DamageBoost + ObjData(UserList(UserIndex).Invent.AnilloEqpObjIndex).MagicDamageBonus
        
        ' Add the magic damage boost from the class modifier
        DamageBoost = DamageBoost + Classes(.clase).ClassMods.MagicDamageBonus
                
    End With
    
    ' Calculate the base and final damage
    With Hechizos(SpellIndex)
        ' Base damage
        Damage = RandomNumber(.MinHp, .MaxHp)
        
        ' Plus character damage
        Damage = Damage + Porcentaje(Damage, 3 * (UserList(UserIndex).Stats.ELV * 100 / ConstantesBalance.MaxLvl * 0.4))
        
        ' Plus boost based on class and equipment
        Damage = Damage + Porcentaje(Damage, DamageBoost)
        
        ' Plus spell mastery bonus
        Damage = Damage + modMasteries.GetMasterySpellPowerBonus(UserIndex, SpellIndex, Damage)
        
        ' Efficacy reduction
        Damage = CalculateAreaEfficacy(Damage, SpellIndex, TargetDistance)
        
        CalculateSpellDamageForUserHP = Damage
    End With
End Function

Private Function CalculateSpellDamageForUser(ByVal UserIndex As Integer, ByVal SpellIndex As Integer, ByVal Min As Integer, ByVal Max As Integer, Optional ByVal TargetDistance As Byte = 0) As Integer

    Dim Damage As Integer
               
    ' Calculate the base and final damage for all spells which do not affect target health
    With Hechizos(SpellIndex)
    
        ' Base damage
        Damage = RandomNumber(Min, Max)
        
        ' Plus character damage
        'Damage = Damage + Porcentaje(Damage, 3 * (UserList(UserIndex).Stats.ELV * 100 / ConstantesBalance.MaxLvl * 0.4))
        
        ' Plus boost based on class and equipment
        'Damage = Damage + Porcentaje(Damage, DamageBoost)
        
        ' Plus spell mastery bonus
        Damage = Damage + modMasteries.GetMasterySpellPowerBonus(UserIndex, SpellIndex, Damage)
        
        ' Efficacy reduction
        Damage = CalculateAreaEfficacy(Damage, SpellIndex, TargetDistance)
        
        CalculateSpellDamageForUser = Damage
    End With
    
End Function

Public Function CanCastSpellByMagicPower(ByVal UserIndex As Integer, ByVal SpellIndex As Integer) As Boolean
    
    Dim CurrentMagicCastPower As Long
    
    With UserList(UserIndex)
        ' Add the magic power of the weapon
        If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then CurrentMagicCastPower = CurrentMagicCastPower + ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).MagicCastPower
        
        ' Add the magic power of the accessory
        If UserList(UserIndex).Invent.AnilloEqpObjIndex > 0 Then CurrentMagicCastPower = CurrentMagicCastPower + ObjData(UserList(UserIndex).Invent.AnilloEqpObjIndex).MagicCastPower
        
        ' Add the magic power of the class
        CurrentMagicCastPower = CurrentMagicCastPower + Classes(.clase).ClassMods.MagicCastPower + .Masteries.Boosts.AddMagicCastPower
    End With
    
    With Hechizos(SpellIndex)
        If Hechizos(SpellIndex).MagicCastPowerRequired > 0 Then
            CanCastSpellByMagicPower = CurrentMagicCastPower >= .MagicCastPowerRequired
        Else
            CanCastSpellByMagicPower = True
        End If
    End With
    
End Function

Public Sub LifeLeechEffect(ByVal SourceType As eTargetType, ByVal SourceIndex As Integer, ByVal DamageDone As Integer, ByVal DamageSourceLifeLeech As Integer)

    Dim LeechedLife As Integer

    ' USER LifeLeech
    If SourceType = eTargetType.ieUser Then
        With UserList(SourceIndex)
            If .Stats.MinHp = .Stats.MaxHp Then Exit Sub
            
            Dim LifeLeechPerc As Integer
            LifeLeechPerc = DamageSourceLifeLeech + .Masteries.Boosts.AddMagicLifeLeechPerc
            LeechedLife = Porcentaje(DamageDone, Min(LifeLeechPerc, 100))
            LeechedLife = Min(LeechedLife, .Stats.MaxHp - .Stats.MinHp)
                    
            .Stats.MinHp = .Stats.MinHp + LeechedLife
        End With
        
        If LeechedLife > 0 Then
            Call WriteConsoleMsg(SourceIndex, "Te has curado " & LeechedLife & " punto" & IIf(LeechedLife > 1, "s", "") & " de vida basado en tu daño", FontTypeNames.FONTTYPE_TALK)
            Call WriteUpdateHP(SourceIndex)
        End If
        
    ' NPC LifeLeech
    ElseIf SourceType = eTargetType.ieNpc Then
        With Npclist(SourceIndex)
            If .Stats.MinHp = .Stats.MaxHp Then Exit Sub
            
            LeechedLife = Porcentaje(DamageDone, Max(100, DamageSourceLifeLeech))
            LeechedLife = Max(.Stats.MaxHp - .Stats.MinHp, LeechedLife)
            
            .Stats.MinHp = .Stats.MinHp + LeechedLife
        End With
    End If
    
End Sub


Public Function GetSpellRequiredMana(ByVal SpellIndex As Integer, ByVal UserIndex As Integer) As Integer

    ' Some spells require a combination of fixe mana + a percentage of the max mana of the player.
    GetSpellRequiredMana = Hechizos(SpellIndex).ManaRequerido + Porcentaje(UserList(UserIndex).Stats.MaxMan, Hechizos(SpellIndex).ManaRequeridoPerc)
       
    ' Substract the mana reduction for the given spell based on the masteries aquired.
    GetSpellRequiredMana = GetSpellRequiredMana - Porcentaje(GetSpellRequiredMana, modMasteries.GetMasteryManaReductionPercentForSpell(UserIndex, SpellIndex))
       
End Function

Public Sub ManaConversionEffect(ByVal UserIndex As Integer, ByVal DamageReceived As Integer)
    Dim ManaConverted As Integer
    
    ' The user will convert a portion of the damage received into mana.
    With UserList(UserIndex)
    
        If .Masteries.Boosts.MagicSpellManaConversionPerc <= 0 Then Exit Sub
        
        ManaConverted = Porcentaje(DamageReceived, .Masteries.Boosts.MagicSpellManaConversionPerc)
        
        .Stats.MinMAN = Min(.Stats.MaxMan, .Stats.MinMAN + ManaConverted)
        
        Call WriteConsoleMsg(UserIndex, "Recuperaste " & ManaConverted & " de mana gracias a una de tus maestrias", FontTypeNames.FONTTYPE_TALK)
        
    End With
    
End Sub

