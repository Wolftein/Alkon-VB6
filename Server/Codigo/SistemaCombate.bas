Attribute VB_Name = "SistemaCombate"

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
'
'Diseño y corrección del modulo de combate por
'Gerardo Saiz, gerardosaiz@yahoo.com
'

'9/01/2008 Pablo (ToxicWaste) - Ahora TODOS los modificadores de Clase se controlan desde Balance.dat


Option Explicit

Public Const MAXDISTANCIAARCO As Byte = 18

Public Function MinimoInt(ByVal a As Integer, ByVal b As Integer) As Integer
On Error GoTo ErrHandler
  
    If a > b Then
        MinimoInt = b
    Else
        MinimoInt = a
    End If
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function MinimoInt de SistemaCombate.bas")
End Function

Public Function MaximoInt(ByVal a As Integer, ByVal b As Integer) As Integer
On Error GoTo ErrHandler
  
    If a > b Then
        MaximoInt = a
    Else
        MaximoInt = b
    End If
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function MaximoInt de SistemaCombate.bas")
End Function

Private Function PoderEvasionEscudo(ByVal UserIndex As Integer) As Long
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)
        PoderEvasionEscudo = .Stats.UserAtributos(eAtributos.Agilidad) * 4 * Classes(.clase).ClassMods.Escudo + GetSkills(UserIndex, eSkill.Defensa)
    End With
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function PoderEvasionEscudo de SistemaCombate.bas")
End Function

Private Function PoderEvasion(ByVal UserIndex As Integer) As Long
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  
    Dim lTemp As Long
    With UserList(UserIndex)
        PoderEvasion = .Stats.UserAtributos(eAtributos.Agilidad) * 4 * Classes(.clase).ClassMods.Evasion + GetSkills(UserIndex, eSkill.Tacticas)
    End With
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function PoderEvasion de SistemaCombate.bas")
End Function

Private Function PoderAtaqueArma(ByVal UserIndex As Integer) As Long
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    Dim PoderAtaqueTemp As Long
    
    With UserList(UserIndex)
        PoderAtaqueArma = .Stats.UserAtributos(eAtributos.Agilidad) * 4 * Classes(.clase).ClassMods.AtaqueArmas + GetSkills(UserIndex, eSkill.Armas) * 2
    End With
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function PoderAtaqueArma de SistemaCombate.bas")
End Function

Private Function PoderAtaqueProyectil(ByVal UserIndex As Integer) As Long
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    Dim PoderAtaqueTemp As Long
    Dim SkillProyectiles As Integer
    
    With UserList(UserIndex)
        PoderAtaqueProyectil = .Stats.UserAtributos(eAtributos.Agilidad) * 4 * Classes(.clase).ClassMods.AtaqueProyectiles + GetSkills(UserIndex, eSkill.Proyectiles) * 2
    End With
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function PoderAtaqueProyectil de SistemaCombate.bas")
End Function

Private Function PoderAtaqueWrestling(ByVal UserIndex As Integer) As Long
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    Dim PoderAtaqueTemp As Long
    Dim WrestlingSkill As Integer
    
    With UserList(UserIndex)
        PoderAtaqueWrestling = .Stats.UserAtributos(eAtributos.Agilidad) * 5 * Classes(.clase).ClassMods.AtaqueWrestling + GetSkills(UserIndex, eSkill.Wrestling)
    End With
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function PoderAtaqueWrestling de SistemaCombate.bas")
End Function

Public Function UserImpactoNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    Dim PoderAtaque As Long
    Dim Arma As Integer
    Dim skill As eSkill
    Dim ProbExito As Long
    
    Arma = UserList(UserIndex).Invent.WeaponEqpObjIndex
    
    If Arma > 0 Then 'Usando un arma
        If ObjData(Arma).proyectil = 1 Then
            PoderAtaque = PoderAtaqueProyectil(UserIndex)
            skill = eSkill.Proyectiles
        Else
            PoderAtaque = PoderAtaqueArma(UserIndex)
            skill = eSkill.Armas
        End If
    Else 'Peleando con puños
        PoderAtaque = PoderAtaqueWrestling(UserIndex)
        skill = eSkill.Wrestling
    End If
    
    ' Chances are rounded
    ProbExito = MaximoInt(10, MinimoInt(100, 50 + ((PoderAtaque - Npclist(NpcIndex).PoderEvasion) * 0.4)))
    
    UserImpactoNpc = (RandomNumber(1, 100) <= ProbExito)
    
    If UserImpactoNpc Then
        Call SubirSkill(UserIndex, skill, True)
    Else
        Call SubirSkill(UserIndex, skill, False)
    End If
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function UserImpactoNpc de SistemaCombate.bas")
End Function

Public Function NpcImpacto(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Revisa si un NPC logra impactar a un user o no
'03/15/2006 Maraxus - Evité una división por cero que eliminaba NPCs
'*************************************************
On Error GoTo ErrHandler
  
    Dim Rechazo As Boolean
    Dim ProbRechazo As Long
    Dim ProbExito As Long
    Dim UserEvasion As Long
    Dim NpcPoderAtaque As Long
    Dim PoderEvasioEscudo As Long
    Dim SkillTacticas As Long
    Dim SkillDefensa As Long
    
    UserEvasion = PoderEvasion(UserIndex)
    NpcPoderAtaque = Npclist(NpcIndex).PoderAtaque
    PoderEvasioEscudo = PoderEvasionEscudo(UserIndex)
    
    SkillTacticas = GetSkills(UserIndex, eSkill.Tacticas)
    SkillDefensa = GetSkills(UserIndex, eSkill.Defensa)
    
    'Esta usando un escudo ???
    If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then UserEvasion = UserEvasion + PoderEvasioEscudo
    
    ' Chances are rounded
    ProbExito = MaximoInt(10, MinimoInt(90, 50 + ((NpcPoderAtaque - UserEvasion) * 0.4)))
    
    NpcImpacto = (RandomNumber(1, 100) <= ProbExito)
    
    ' el usuario esta usando un escudo ???
    If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
        If Not NpcImpacto Then
            If SkillDefensa + SkillTacticas > 0 Then  'Evitamos división por cero
                ' Chances are rounded
                ProbRechazo = MaximoInt(10, MinimoInt(90, 100 * SkillDefensa / (SkillDefensa + SkillTacticas)))
            Else
                ProbRechazo = 10 'Si no tiene skills le dejamos el 10% mínimo
            End If
            
            Rechazo = (RandomNumber(1, 100) <= ProbRechazo)
                
            If Rechazo Then
                'Se rechazo el ataque con el escudo
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(ConstantesSonidos.Escudo, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, UserList(UserIndex).Char.CharIndex))
                Call WriteMultiMessage(UserIndex, eMessages.BlockedWithShieldUser) 'Call WriteBlockedWithShieldUser(UserIndex)
                Call SubirSkill(UserIndex, eSkill.Defensa, True)
            Else
                Call SubirSkill(UserIndex, eSkill.Defensa, False)
            End If
        End If
    End If
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function NpcImpacto de SistemaCombate.bas")
End Function

Public Function CalculateDamage(ByVal UserIndex As Integer, Optional ByVal NpcIndex As Integer = 0, Optional ByVal IsFirstHitFromArea As Boolean = True) As Long

On Error GoTo ErrHandler
  
    Dim WeaponDamage As Long
    Dim UserDamage As Long
    Dim Weapon As ObjData
    Dim ClassModifier As Single
    Dim proyectil As ObjData
    Dim MaxWeaponDamage As Long
    Dim MinWeaponDamage As Long
    Dim ObjIndex As Integer
    Dim RangedMod As Double
    Dim DistToTarget As Byte
    Dim ReduceBaseDamagePerc As Integer
    
    DistToTarget = 1
    RangedMod = 1
    
    ' Remove this if we want the dragon to be killed in one hit.
    Dim OneShotKill As Boolean
    OneShotKill = False
    
    With UserList(UserIndex)
        If .Invent.WeaponEqpObjIndex > 0 Then
            Weapon = ObjData(.Invent.WeaponEqpObjIndex)
            
            If NpcIndex > 0 Then
                If Weapon.proyectil = 1 Then
                    DistToTarget = Distancia(.Pos, Npclist(NpcIndex).Pos)
                    ClassModifier = Classes(.clase).ClassMods.DamageProjectiles
                    WeaponDamage = RandomNumber(Weapon.MinHit, Weapon.MaxHit)
                    MaxWeaponDamage = Weapon.MaxHit
                    
                    If Weapon.Municion = 1 Then
                        proyectil = ObjData(.Invent.MunicionEqpObjIndex)
                        WeaponDamage = WeaponDamage + RandomNumber(proyectil.MinHit, proyectil.MaxHit)
                    End If
                Else
                    ClassModifier = Classes(.clase).ClassMods.DamageWeapons
                    
                    ' The Great Dragon should be killed with one shot
                    If .Invent.WeaponEqpObjIndex = ConstantesItems.EspadaMataDragones Then ' Usa la mata Dragones?
                        If Npclist(NpcIndex).NPCtype = DRAGON Then 'Ataca Dragon?
                            WeaponDamage = RandomNumber(Weapon.MinHit, Weapon.MaxHit)
                            MaxWeaponDamage = Weapon.MaxHit
                            OneShotKill = True
                        Else
                            WeaponDamage = 1
                            MaxWeaponDamage = 1
                        End If
                    Else
                        WeaponDamage = RandomNumber(Weapon.MinHit, Weapon.MaxHit)
                        MaxWeaponDamage = Weapon.MaxHit
                    End If
                End If
            Else
            
                If Weapon.proyectil <> 1 Then
                    
                    ClassModifier = Classes(.clase).ClassMods.DamageWeapons
                    
                    If .Invent.WeaponEqpObjIndex = ConstantesItems.EspadaMataDragones Then
                        ClassModifier = Classes(.clase).ClassMods.DamageWeapons
                        WeaponDamage = 1 ' Si usa la espada mataDragones daño es 1
                        MaxWeaponDamage = 1
                    Else
                        WeaponDamage = RandomNumber(Weapon.MinHit, Weapon.MaxHit)
                        MaxWeaponDamage = Weapon.MaxHit
                    End If
                Else
                    DistToTarget = Distancia(.Pos, UserList(.flags.TargetUser).Pos)
                End If
                
            End If
                
            If Weapon.proyectil = 1 Then
                ClassModifier = Classes(.clase).ClassMods.DamageProjectiles
                WeaponDamage = RandomNumber(Weapon.MinHit, Weapon.MaxHit)
                MaxWeaponDamage = Weapon.MaxHit
                
                ' Sharing the same functionallity for both bows+arrows and throwing knifes
                If Weapon.Municion = 1 Or Weapon.Acuchilla = 1 Then
                    If Weapon.Municion = 1 Then proyectil = ObjData(.Invent.MunicionEqpObjIndex)
                    If Weapon.Acuchilla = 1 Then proyectil = Weapon
                    
                    WeaponDamage = WeaponDamage + RandomNumber(proyectil.MinHit, proyectil.MaxHit)
                    
                    Dim ClassDamageReduction As Byte
                    ClassDamageReduction = Classes(.clase).ClassMods.DistanceDmgReduction
                    
                    ' Calculate the ranged mod
                    If ClassDamageReduction = 0 Or (ClassDamageReduction > 0 And DistToTarget < 3) Then
                        RangedMod = 1
                    Else
                        RangedMod = 1 - IIf(ClassDamageReduction = 0, 1, (DistToTarget - 3) * (0.01 * ClassDamageReduction))
                    End If
                    
                End If
            End If
            
            ' We should apply a reduction in the base damage if we're using a weapon with SplashDamage
            ' And this is not the first hit from the Damage Vector
            If Weapon.SplashDamage = True And Not IsFirstHitFromArea Then
                ReduceBaseDamagePerc = Weapon.SplashDamageReduction
            End If
            
        Else
            ' The user is Wrestling (hitting without a weapon)
            
            ClassModifier = Classes(.clase).ClassMods.DamageWrestling
            
            MinWeaponDamage = Classes(.clase).ClassMods.DamageWrestlingMin
            MaxWeaponDamage = Classes(.clase).ClassMods.DamageWrestlingMax
            
            ObjIndex = .Invent.AnilloEqpObjIndex
            If ObjIndex > 0 Then
                If ObjData(ObjIndex).Guante = 1 Then
                    MinWeaponDamage = MinWeaponDamage + ObjData(ObjIndex).MinHit
                    MaxWeaponDamage = MaxWeaponDamage + ObjData(ObjIndex).MaxHit
                End If
            End If
            
            WeaponDamage = RandomNumber(MinWeaponDamage, MaxWeaponDamage)
            
        End If
        
        UserDamage = .Stats.ELV * Classes(.Clase).ClassMods.PhysicalDamage + Classes(.Clase).ClassMods.BaseDamage
        
        If OneShotKill Then
            CalculateDamage = Npclist(NpcIndex).Stats.MinHp + Npclist(NpcIndex).Stats.Def
        Else
            CalculateDamage = ((3 * WeaponDamage + ((MaxWeaponDamage / 5) * MaximoInt(0, .Stats.UserAtributos(eAtributos.Fuerza) - 15))) * ClassModifier + UserDamage) * RangedMod
        End If
        
        CalculateDamage = CalculateDamage - Porcentaje(CalculateDamage, ReduceBaseDamagePerc)
        
    End With
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function CalculateDamage de SistemaCombate.bas")
End Function

Public Sub UserDanioNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal IsFirstHitFromArea As Boolean)
'***************************************************
'Author: Unknown
'Last Modification: 07/04/2010 (Pato)
'25/01/2010: ZaMa - Agrego poder acuchillar npcs.
'07/04/2010: ZaMa - Los asesinos apuñalan acorde al Danio base sin descontar la defensa del npc.
'07/04/2010: Pato - Si se mata al dragón en party se loguean los miembros de la misma.
'11/07/2010: ZaMa - Ahora la defensa es solo ignorada para asesinos.
'***************************************************
On Error GoTo ErrHandler
  

    Dim Danio As Long
    Dim DanioBase As Long
    Dim FinalDefense As Long
    Dim Perforation As Integer
    Dim PI As Integer
    Dim MembersOnline() As Integer
    ReDim MembersOnline(1 To Constantes.MaxPartyMembers) As Integer
    Dim Text As String
    Dim I As Integer
    
    Dim BoatIndex As Integer
    
    DanioBase = CalculateDamage(UserIndex, NpcIndex, IsFirstHitFromArea)
       
    With UserList(UserIndex)
        'esta navegando? si es asi le sumamos el Danio del barco
        If .flags.Navegando = 1 Then
        
            BoatIndex = .Invent.BarcoObjIndex
            If BoatIndex > 0 Then
                DanioBase = DanioBase + RandomNumber(ObjData(BoatIndex).MinHit, ObjData(BoatIndex).MaxHit)
            End If
        End If
        
         ' Perforation of the attacker's weapon
        If .Invent.WeaponEqpObjIndex > 0 Then
            Perforation = Perforation + ObjData(.Invent.WeaponEqpObjIndex).Perforation
        End If
        
        ' Perforation of the attacker's ammunition
        If .Invent.MunicionEqpObjIndex > 0 Then
            Perforation = Perforation + ObjData(.Invent.MunicionEqpObjIndex).Perforation
        End If
    End With
   
    With Npclist(NpcIndex)
        FinalDefense = MaximoInt(.Stats.Def - Perforation, 0)
        Danio = Max(1, DanioBase - FinalDefense)
                
        Call WriteMultiMessage(UserIndex, eMessages.UserHitNPC, Danio)
        Call CalcularDarExp(UserIndex, NpcIndex, Danio)
        .Stats.MinHp = .Stats.MinHp - Danio
        
       'Flecha en NPC
        If UserList(UserIndex).Invent.WeaponEqpObjIndex <> 0 Then
            'This entire block should be refactored because now we can drop not only arrows, but also knifes
            If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Municion = 1 Then
                Select Case UserList(UserIndex).Invent.MunicionEqpObjIndex
                    Case ConstantesItems.Flecha
                        .TengoFlechas(1) = .TengoFlechas(1) + 1
                    Case ConstantesItems.Flecha1
                        .TengoFlechas(2) = .TengoFlechas(2) + 1
                    Case ConstantesItems.Flecha2
                        .TengoFlechas(3) = .TengoFlechas(3) + 1
                    Case ConstantesItems.Flecha3
                        .TengoFlechas(4) = .TengoFlechas(4) + 1
                    Case ConstantesItems.FlechaNewbie
                        .TengoFlechas(5) = .TengoFlechas(5) + 1
                End Select
            End If
            
            'This entire block should be refactored because now we can drop not only arrows, but also knifes
            If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).proyectil = 1 Then
                Select Case UserList(UserIndex).Invent.WeaponEqpObjIndex
                    Case ConstantesItems.Cuchillas
                        .TengoFlechas(6) = .TengoFlechas(6) + 1
                End Select
            End If
            
        End If
        
        If .Stats.MinHp > 0 Then
            'Trata de apuñalar por la espalda al enemigo
            If PuedeApuñalar(UserIndex) Then
                
                ' La defensa se ignora solo en asesinos
                If UserList(UserIndex).clase <> eClass.Assasin Then
                    DanioBase = Danio
                End If
                
                Call DoStab(UserIndex, NpcIndex, 0, DanioBase)
                
            End If

            'trata de dar golpe crítico
            Call DoGolpeCritico(UserIndex, NpcIndex, 0, Danio)
            
            If PuedeAcuchillar(UserIndex) Then
                Call DoAcuchillar(UserIndex, NpcIndex, 0, Danio)
            End If
        End If
        
        If .Stats.MinHp <= 0 Then
            ' Si era un Dragon perdemos la espada mataDragones
            If .NPCtype = DRAGON Then
                'Si tiene equipada la matadracos se la sacamos
                If UserList(UserIndex).Invent.WeaponEqpObjIndex = ConstantesItems.EspadaMataDragones Then
                    Call QuitarObjetos(ConstantesItems.EspadaMataDragones, 1, UserIndex)
                End If
                If .Stats.MaxHp > 100000 Then
                    Text = UserList(UserIndex).Name & " mató un dragón"
                    PI = UserList(UserIndex).PartyIndex
                    
                    If PI > 0 Then
                        Call Parties(PI).ObtenerMiembrosOnline(MembersOnline())
                        Text = Text & " estando en party "
                        
                        For I = 1 To Constantes.MaxPartyMembers
                            If MembersOnline(I) > 0 Then
                                Text = Text & UserList(MembersOnline(I)).Name & ", "
                            End If
                        Next I
                        
                        Text = Left$(Text, Len(Text) - 2) & ")"
                    End If
                    
                    Call LogDesarrollo(Text & ".")
                End If
            End If
            
            ' Para que las mascotas no sigan intentando luchar y
            ' comiencen a seguir al amo
            For I = 1 To Classes(UserList(UserIndex).clase).ClassMods.MaxTammedPets
                
                ' Tammed pets
                If UserList(UserIndex).TammedPets(I).NpcIndex > 0 Then
                    If Npclist(UserList(UserIndex).TammedPets(I).NpcIndex).TargetNPC = NpcIndex Then
                        Npclist(UserList(UserIndex).TammedPets(I).NpcIndex).TargetNPC = 0
                        Npclist(UserList(UserIndex).TammedPets(I).NpcIndex).Movement = TipoAI.SigueAmo
                    End If
                End If
                
            Next I
            
            For I = 1 To Classes(UserList(UserIndex).clase).ClassMods.MaxInvokedPets
                
                ' Invoked pets
                If UserList(UserIndex).InvokedPets(I).NpcIndex > 0 Then
                    If Npclist(UserList(UserIndex).InvokedPets(I).NpcIndex).TargetNPC = NpcIndex Then
                        Npclist(UserList(UserIndex).InvokedPets(I).NpcIndex).TargetNPC = 0
                        Npclist(UserList(UserIndex).InvokedPets(I).NpcIndex).Movement = TipoAI.SigueAmo
                    End If
                End If
                
            Next I
            
            Call MuereNpc(NpcIndex, UserIndex)
        Else
            If .ExtraBodies > 0 Then
                If .ActualBody < .ExtraBodies Then
                    If .Stats.MinHp < .Stats.MaxHp * ((.ExtraBodies - .ActualBody) / (.ExtraBodies + 1)) Then
                        .ActualBody = .ActualBody + 1
                        .Char.body = .ExtraBody(.ActualBody)
                        Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCharacterChange(.ExtraBody(.ActualBody), _
                                        .Char.head, .Char.heading, .Char.CharIndex, .Char.WeaponAnim, .Char.ShieldAnim, _
                                        .Char.FX, .Char.Loops, .Char.CascoAnim, CBool(.flags.TierraInvalida = 1), False, 0, eCharacterAlignment.Neutral))
                    End If
                End If
            End If
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub UserDanioNpc de SistemaCombate.bas")
End Sub

Public Sub NpcDanio(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 03/06/2011 (Amraphen)
'18/09/2010: ZaMa - Ahora se considera siempre la defensa del barco y el escudo.
'03/06/2011: Amraphen - Agrego defensa adicional de armadura de segunda jerarquía.
'***************************************************
On Error GoTo ErrHandler
  

    Dim Damage As Integer
    Dim Lugar As Integer
    Dim Obj As ObjData
    
    Dim BoatDefense As Integer
    Dim HeadDefense As Integer
    Dim BodyDefense As Integer
    
    Dim BoatIndex As Integer
    Dim HelmetIndex As Integer
    Dim ArmourIndex As Integer
    Dim ShieldIndex As Integer
    
    Damage = RandomNumber(Npclist(NpcIndex).Stats.MinHit, Npclist(NpcIndex).Stats.MaxHit)
    
    With Npclist(NpcIndex)
        If .Movement = TipoAI.BossDI Then
            If .Stats.MinHp < (.Stats.MaxHp - (.Stats.MaxHp / 3 * 2)) Then 'Si tiene dos tercio de vida
                Damage = Damage * 3
            ElseIf .Stats.MinHp < (.Stats.MaxHp - (.Stats.MaxHp / 3)) Then  'Si tiene un tercios de vida
                Damage = Damage * 2
            End If
        End If
    End With
    
    With UserList(UserIndex)
        ' Navega?
        If .flags.Navegando = 1 Then
            ' En barca suma defensa
            BoatIndex = .Invent.BarcoObjIndex
            If BoatIndex > 0 Then
                Obj = ObjData(BoatIndex)
                BoatDefense = RandomNumber(Obj.MinDef, Obj.MaxDef)
            End If
        End If
        
        Lugar = RandomNumber(PartesCuerpo.bCabeza, PartesCuerpo.bTorso)
        
        Select Case Lugar
        
            Case PartesCuerpo.bCabeza
            
                'Si tiene casco absorbe el golpe
                HelmetIndex = .Invent.CascoEqpObjIndex
                If HelmetIndex > 0 Then
                   Obj = ObjData(HelmetIndex)
                   HeadDefense = RandomNumber(Obj.MinDef, Obj.MaxDef)
                End If
                
            Case Else
                
                Dim MinDef As Integer
                Dim MaxDef As Integer
            
                'Si tiene armadura absorbe el golpe
                ArmourIndex = .Invent.ArmourEqpObjIndex
                If ArmourIndex > 0 Then
                    Obj = ObjData(ArmourIndex)
                    MinDef = Obj.MinDef
                    MaxDef = Obj.MaxDef
                End If
                
                'Si tiene armadura de segunda jerarquía obtiene un porcentaje de defensa adicional.
                If .Invent.FactionArmourEqpObjIndex > 0 Then
                    MinDef = MinDef * ConstantesBalance.ModDefSegJerarquia
                    MaxDef = MaxDef * ConstantesBalance.ModDefSegJerarquia
                End If
                
                ' Si tiene escudo absorbe el golpe
                ShieldIndex = .Invent.EscudoEqpObjIndex
                If ShieldIndex > 0 Then
                    Obj = ObjData(ShieldIndex)
                    MinDef = MinDef + Obj.MinDef
                    MaxDef = MaxDef + Obj.MaxDef
                End If
                
                BodyDefense = RandomNumber(MinDef, MaxDef)
        
        End Select
        
        ' Damage final
        Damage = Max(1, Damage - HeadDefense - BodyDefense - BoatDefense)
        
        Call WriteMultiMessage(UserIndex, eMessages.NPCHitUser, Lugar, Damage)
        
        If .flags.Privilegios And PlayerType.User Then .Stats.MinHp = .Stats.MinHp - Damage
        
        If .flags.Meditando Then
            If Damage > Fix(.Stats.MinHp / 100 * .Stats.UserAtributos(eAtributos.Inteligencia) * _
               GetSkills(UserIndex, eSkill.Meditar) / 100 * 12 / (RandomNumber(0, 5) + 7)) Then
                .flags.Meditando = False
                Call WriteMeditateToggle(UserIndex)
                Call WriteConsoleMsg(UserIndex, "Dejas de meditar.", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)
                .Char.FX = 0
                .Char.Loops = 0
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0, , 0))
            End If
        End If
        
        'Muere el usuario
        If .Stats.MinHp <= 0 Then
            Call WriteMultiMessage(UserIndex, eMessages.NPCKillUser)  'Le informamos que ha muerto ;)
            
            Call UserDie(UserIndex)
         
            If Npclist(NpcIndex).MaestroUser > 0 Then
                Call AllFollowAmo(Npclist(NpcIndex).MaestroUser)
                
                If .flags.AtacablePor <> Npclist(NpcIndex).MaestroUser Then
                    Call Statistics.StoreFrag(Npclist(NpcIndex).MaestroUser, UserIndex)
                    Call ContarMuerte(UserIndex, Npclist(NpcIndex).MaestroUser, eDamageType.NpcDamage, Damage)
                End If
                
                Call ActStats(UserIndex, Npclist(NpcIndex).MaestroUser)
            Else
                'Al matarlo no lo sigue mas
                With Npclist(NpcIndex)
                    If .Stats.Alineacion = 0 Then
                        .Movement = .flags.OldMovement
                        .Hostile = .flags.OldHostil
                        .flags.AttackedBy = vbNullString
                    End If
                End With
                
            End If
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub NpcDanio de SistemaCombate.bas")
End Sub


Public Sub CheckPets(ByVal NpcIndex As Integer, ByVal UserIndex As Integer, Optional ByVal CheckElementales As Boolean = True)
'***************************************************
'Author: Unknown
'Last Modification: 15/04/2010
'15/04/2010: ZaMa - Las mascotas no se apropian de npcs.
'***************************************************
On Error GoTo ErrHandler
  

    Dim J As Integer

    ' TODO: Make sure TammedPetsCount and InvokedPetsCount holds the right values
    ' and use these properties to prevent looping through both arrays.
    '
    ' Si no tengo mascotas, para que cheaquear lo demas?
    'If UserList(UserIndex).TammedPetsCount = 0 Then Exit Sub
    
    If Not PuedeAtacarNPC(UserIndex, NpcIndex, , True) Then Exit Sub
    
    With UserList(UserIndex)
        For J = 1 To Classes(.clase).ClassMods.MaxTammedPets
            If .TammedPets(J).NpcIndex > 0 Then
                If .TammedPets(J).NpcIndex <> NpcIndex Then
                    If Npclist(.TammedPets(J).NpcIndex).TargetNPC = 0 Then Npclist(.TammedPets(J).NpcIndex).TargetNPC = NpcIndex
                    Npclist(.TammedPets(J).NpcIndex).Movement = TipoAI.NpcAtacaNpc
                End If
            End If
        Next J
        
        For J = 1 To Classes(.clase).ClassMods.MaxInvokedPets
            If .InvokedPets(J).NpcIndex > 0 Then
               If .InvokedPets(J).NpcIndex <> NpcIndex Then
                If CheckElementales Or (Npclist(.InvokedPets(J).NpcIndex).Numero <> ConstantesNPCs.EleFuego And Npclist(.InvokedPets(J).NpcIndex).Numero <> ConstantesNPCs.EleTierra) Then
                    
                    If Npclist(.InvokedPets(J).NpcIndex).TargetNPC = 0 Then Npclist(.InvokedPets(J).NpcIndex).TargetNPC = NpcIndex
                    Npclist(.InvokedPets(J).NpcIndex).Movement = TipoAI.NpcAtacaNpc
                End If
               End If
            End If
        Next J
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CheckPets de SistemaCombate.bas")
End Sub

Public Sub AllFollowAmo(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    Dim J As Integer
    
    With UserList(UserIndex)
        For J = 1 To Classes(.clase).ClassMods.MaxTammedPets
            If .TammedPets(J).NpcIndex > 0 Then
                Call FollowAmo(.TammedPets(J).NpcIndex)
            End If
        Next J
        
        For J = 1 To Classes(.clase).ClassMods.MaxInvokedPets
            If .InvokedPets(J).NpcIndex > 0 Then
                Call FollowAmo(.InvokedPets(J).NpcIndex)
            End If
        Next J
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub AllFollowAmo de SistemaCombate.bas")
End Sub

Public Function NpcAtacaUser(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean
'*************************************************
'Author: Unknown
'Last modified: -
'
'*************************************************
On Error GoTo ErrHandler
  

    With UserList(UserIndex)
        If .flags.AdminInvisible = 1 Then Exit Function
        If (Not .flags.Privilegios And PlayerType.User) <> 0 And Not .flags.AdminPerseguible Then Exit Function
    End With
    
    With Npclist(NpcIndex)
        ' El npc puede atacar ???
        NpcAtacaUser = True
        Call CheckPets(NpcIndex, UserIndex, False)
        
        If .Target = 0 Then .Target = UserIndex
        
        If UserList(UserIndex).flags.AtacadoPorNpc = 0 And UserList(UserIndex).flags.AtacadoPorUser = 0 Then
            UserList(UserIndex).flags.AtacadoPorNpc = NpcIndex
        End If
        
        If .flags.Snd1 > 0 Then
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(.flags.Snd1, .Pos.X, .Pos.Y, .Char.CharIndex))
        End If
    End With
    
    Dim MasterIndex As Integer
    MasterIndex = Npclist(NpcIndex).MaestroUser
    
    If NpcImpacto(NpcIndex, UserIndex) Then
        With UserList(UserIndex)
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(ConstantesSonidos.Impacto, .Pos.X, .Pos.Y, .Char.CharIndex))
            
            If .flags.Meditando = False Then
                If .flags.Navegando = 0 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, ConstantesFX.FxSangre, 0))
                End If
            End If
            
            Call NpcDanio(NpcIndex, UserIndex)
            Call WriteUpdateHP(UserIndex)
            
            '¿Puede envenenar?
            If Npclist(NpcIndex).Veneno = 1 Then Call NpcEnvenenarUser(UserIndex)
            
        End With
        
        Call SubirSkill(UserIndex, eSkill.Tacticas, False)
        
    Else
        Call WriteMultiMessage(UserIndex, eMessages.NPCSwing)
        Call SubirSkill(UserIndex, eSkill.Tacticas, True)
        
    End If
    
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function NpcAtacaUser de SistemaCombate.bas")
End Function

Private Function NpcImpactoNpc(ByVal Atacante As Integer, ByVal Victima As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    Dim PoderAtt As Long
    Dim PoderEva As Long
    Dim ProbExito As Long
    
    PoderAtt = Npclist(Atacante).PoderAtaque
    PoderEva = Npclist(Victima).PoderEvasion
    
    ' Chances are rounded
    ProbExito = MaximoInt(10, MinimoInt(90, 50 + (PoderAtt - PoderEva) * 0.4))
    NpcImpactoNpc = (RandomNumber(1, 100) <= ProbExito)
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function NpcImpactoNpc de SistemaCombate.bas")
End Function

Public Sub NpcDanioNpc(ByVal Atacante As Integer, ByVal Victima As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    Dim Damage As Integer
    Dim MasterIndex As Integer
    
    With Npclist(Atacante)
        Damage = RandomNumber(.Stats.MinHit, .Stats.MaxHit)
        Npclist(Victima).Stats.MinHp = Npclist(Victima).Stats.MinHp - Damage
        
        If .MaestroUser > 0 Then
            Call CalcularDarExp(.MaestroUser, Victima, Damage)
        End If
                
        If Npclist(Victima).Stats.MinHp < 1 Then
            .Movement = .flags.OldMovement
            
            If LenB(.flags.AttackedBy) <> 0 Then
                .Hostile = .flags.OldHostil
            End If
            
            MasterIndex = .MaestroUser
            If MasterIndex > 0 Then
                Call FollowAmo(Atacante)
            End If
            
            Call MuereNpc(Victima, MasterIndex)
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub NpcDanioNpc de SistemaCombate.bas")
End Sub

Public Sub NpcAtacaNpc(ByVal Atacante As Integer, ByVal Victima As Integer, Optional ByVal cambiarMOvimiento As Boolean = True)
'*************************************************
'Author: Unknown
'Last modified: 01/03/2009
'01/03/2009: ZaMa - Las mascotas no pueden atacar al rey si quedan pretorianos vivos.
'23/05/2010: ZaMa - Ahora los elementales renuevan el tiempo de pertencia del npc que atacan si pertenece a su amo.
'*************************************************
On Error GoTo ErrHandler
  
    
    Dim MasterIndex As Integer
    
    With Npclist(Atacante)
        
        'Es el Rey Preatoriano?
        If Npclist(Victima).NPCtype = eNPCType.Pretoriano Then
            If Not ClanPretoriano(Npclist(Victima).ClanIndex).CanAtackMember(Victima) Then
                Call WriteConsoleMsg(.MaestroUser, "Debes matar al resto del ejército antes de atacar al rey!", FontTypeNames.FONTTYPE_FIGHT)
                .TargetNPC = 0
                Exit Sub
            End If
        End If
        
        If cambiarMOvimiento Then
            Npclist(Victima).TargetNPC = Atacante
            Npclist(Victima).Movement = TipoAI.NpcAtacaNpc
        End If

        If .flags.Snd1 > 0 Then
            Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessagePlayWave(.flags.Snd1, .Pos.X, .Pos.Y, .Char.CharIndex))
        End If
        
        MasterIndex = .MaestroUser
        
        ' Tiene maestro?
        If MasterIndex > 0 Then
            ' Su maestro es dueño del npc al que ataca?
            If Npclist(Victima).Owner = MasterIndex Then
                ' Renuevo el timer de pertenencia
                Call IntervaloPerdioNpc(MasterIndex, True)
            End If
        End If
        
        
        If NpcImpactoNpc(Atacante, Victima) Then
            If Npclist(Victima).flags.Snd2 > 0 Then
                Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(Npclist(Victima).flags.Snd2, Npclist(Victima).Pos.X, Npclist(Victima).Pos.Y, Npclist(Victima).Char.CharIndex))
            Else
                Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(ConstantesSonidos.Impacto2, Npclist(Victima).Pos.X, Npclist(Victima).Pos.Y, Npclist(Victima).Char.CharIndex))
            End If
            
            If MasterIndex > 0 Then
                
                Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessagePlayWave(ConstantesSonidos.Impacto, .Pos.X, .Pos.Y, .Char.CharIndex))
            Else
            Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(ConstantesSonidos.Impacto, Npclist(Victima).Pos.X, Npclist(Victima).Pos.Y, Npclist(Victima).Char.CharIndex))
            End If
            
            Call NpcDanioNpc(Atacante, Victima)
        Else
            If MasterIndex > 0 Then
                Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessagePlayWave(ConstantesSonidos.Swing, .Pos.X, .Pos.Y, .Char.CharIndex))
            Else
                Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(ConstantesSonidos.Swing, Npclist(Victima).Pos.X, Npclist(Victima).Pos.Y, Npclist(Victima).Char.CharIndex))
            End If
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub NpcAtacaNpc de SistemaCombate.bas")
End Sub

Public Function UsuarioAtacaNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal IsFirstHitFromArea As Boolean) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 03/08/2012 (Amraphen)
'12/01/2010: ZaMa - Los druidas pierden la inmunidad de ser atacados por npcs cuando los atacan.
'14/01/2010: ZaMa - Lo transformo en función, para que no se pierdan municiones al atacar targets inválidos.
'13/02/2011: Amraphen - Ahora la stamina es quitada cuando efectivamente se ataca al NPC.
'24/05/2011: Amraphen - Ahora se envía la animación del pj al golpear.
'03/08/2012: ZaMa - Ahora apuñlar sube skills como fallido al fallar el golpe.
'***************************************************

On Error GoTo ErrHandler

    If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then Exit Function
    
    Call NPCAtacado(NpcIndex, UserIndex)
    
    Call CheckPets(NpcIndex, UserIndex, False)
    
    If UserImpactoNpc(UserIndex, NpcIndex) Then
        'Send animation
        Call SendCharacterSwing(UserIndex)
            
        If Npclist(NpcIndex).flags.Snd2 > 0 Then
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(Npclist(NpcIndex).flags.Snd2, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y, Npclist(NpcIndex).Char.CharIndex))
        Else
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(ConstantesSonidos.Impacto2, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y, Npclist(NpcIndex).Char.CharIndex))
        End If
        
        'y ahora, el ladrón puede llegar a paralizar con el golpe.
        If UserList(UserIndex).clase = eClass.Thief Then
            Call DoHandInmoNpc(UserIndex, NpcIndex)
        End If
            
        Call UserDanioNpc(UserIndex, NpcIndex, IsFirstHitFromArea)
    Else
        ' Si no impacta, apuñalar sube como fallido
        If PuedeApuñalar(UserIndex) Then
            Call SubirSkill(UserIndex, eSkill.Apuñalar, False)
        End If

        'Send animation
        Call SendCharacterSwing(UserIndex)
            
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(ConstantesSonidos.Swing, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, UserList(UserIndex).Char.CharIndex))
    
        Call WriteMultiMessage(UserIndex, eMessages.UserSwing)
    End If
    
    ' Reveló su condición de usuario al atacar, los npcs lo van a atacar
    UserList(UserIndex).flags.Ignorado = False
    
    UsuarioAtacaNpc = True
    
    Exit Function
    
ErrHandler:
    Dim UserName As String
    
    If UserIndex > 0 Then UserName = UserList(UserIndex).Name
    
    Call LogError("Error en UsuarioAtacaNpc. Error " & Err.Number & " : " & Err.Description & ". User: " & _
                   UserIndex & "-> " & UserName & ". NpcIndex: " & NpcIndex & ".")
    
End Function

Public Function UsuarioAtaca(ByVal UserIndex As Integer) As Boolean
On Error GoTo ErrHandler
    
    If Not EsGm(UserIndex) Then
        
        'Check bow's interval
        If Not IntervaloPermiteUsarArcos(UserIndex, False) Then Exit Function
        
        'Check Spell-Magic interval
        If Not IntervaloPermiteMagiaGolpe(UserIndex) Then
            'Check Attack interval
            If Not IntervaloPermiteAtacar(UserIndex) Then
                Exit Function
            End If
        End If
            
    End If
    
    Dim TargetIndex As Integer
    Dim AttackPos As WorldPos
        
    With UserList(UserIndex)
        Dim RequiredStamina As Integer
        RequiredStamina = GetStaminaRequiredToAttack(UserIndex)
        
        'Chequeamos que tenga la energía necesaria para atacar
        If .Stats.MinSta < RequiredStamina Then
            Call WriteConsoleMsg(UserIndex, "No tenés suficiente energía para luchar.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
    
        Dim AttackVector() As tAttackPosition
        AttackVector = GetAttackVector(UserIndex, .Char.heading, .Invent.WeaponEqpObjIndex)
        
        Dim I As Integer
        Dim CanHitToTarget As Boolean
        Dim AttackSucceeded As Boolean
        
        For I = 1 To UBound(AttackVector)
            CanHitToTarget = True
            
            AttackPos = AttackVector(I).Pos
            
            If AttackPos.X < XMinMapSize Or AttackPos.X > XMaxMapSize Or AttackPos.Y <= YMinMapSize Or AttackPos.Y > YMaxMapSize Then
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(ConstantesSonidos.Swing, .Pos.X, .Pos.Y, .Char.CharIndex))
                CanHitToTarget = False
            End If
            
            If CanHitToTarget Then
            
                TargetIndex = MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).UserIndex
                
                'Look for user
                If TargetIndex > 0 Then
                    AttackSucceeded = UsuarioAtacaUsuario(UserIndex, TargetIndex, Not AttackVector(I).ReducedDamageFromSplash)
                    Call WriteUpdateUserStats(TargetIndex)
                    
                    'Exit Sub
                Else
                
                    TargetIndex = MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).NpcIndex
                    
                     'Look for NPC
                    If TargetIndex > 0 Then
                        If Npclist(TargetIndex).Attackable And (Npclist(TargetIndex).MaestroUser = 0 Or (Npclist(TargetIndex).MaestroUser > 0 And MapInfo(Npclist(TargetIndex).Pos.Map).Pk = True)) Then
                            AttackSucceeded = UsuarioAtacaNpc(UserIndex, TargetIndex, Not AttackVector(I).ReducedDamageFromSplash)
                        End If
                    End If
                End If

            End If
        Next I
        
        Call QuitarSta(UserIndex, RequiredStamina)
                
        'Send animation
        Call SendCharacterSwing(UserIndex)
        
        'Send sound
        If (Not AttackSucceeded) Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(ConstantesSonidos.Swing, .Pos.X, .Pos.Y, .Char.CharIndex))
        End If
        
        Call WriteUpdateUserStats(UserIndex)
        
        UsuarioAtaca = True

    End With
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub UsuarioAtaca de SistemaCombate.bas")
End Function

Public Function UsuarioImpacto(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 21/05/2010
'21/05/2010: ZaMa - Evito division por cero.
'29/07/2016: Anagrama - Arreglado un error que causaba que el minimo de punteria fuera 1%.
'***************************************************

On Error GoTo ErrHandler

    Dim ProbRechazo As Long
    Dim Rechazo As Boolean
    Dim ProbExito As Long
    Dim PoderAtaque As Long
    Dim UserPoderEvasion As Long
    Dim UserPoderEvasionEscudo As Long
    Dim Arma As Integer
    Dim SkillTacticas As Long
    Dim SkillDefensa As Long
    Dim ProbEvadir As Long
    Dim skill As eSkill
    
    With UserList(VictimaIndex)
        SkillTacticas = GetSkills(VictimaIndex, eSkill.Tacticas)
        SkillDefensa = GetSkills(VictimaIndex, eSkill.Defensa)
        
        Arma = UserList(AtacanteIndex).Invent.WeaponEqpObjIndex
        
        'Calculamos el poder de evasion...
        UserPoderEvasion = PoderEvasion(VictimaIndex)
        
        If .Invent.EscudoEqpObjIndex > 0 Then
           UserPoderEvasionEscudo = PoderEvasionEscudo(VictimaIndex)
           UserPoderEvasion = UserPoderEvasion + UserPoderEvasionEscudo
        Else
            UserPoderEvasionEscudo = 0
        End If
        
        'Esta usando un arma ???
        If UserList(AtacanteIndex).Invent.WeaponEqpObjIndex > 0 Then
            If ObjData(Arma).proyectil = 1 Then
                PoderAtaque = PoderAtaqueProyectil(AtacanteIndex)
                skill = eSkill.Proyectiles
            Else
                PoderAtaque = PoderAtaqueArma(AtacanteIndex)
                skill = eSkill.Armas
            End If
        Else
            PoderAtaque = PoderAtaqueWrestling(AtacanteIndex)
            skill = eSkill.Wrestling
        End If
        
        ' Chances are rounded
        ProbExito = MaximoInt(10, MinimoInt(100, 50 + (PoderAtaque - UserPoderEvasion) * 0.4))
        
        ' Se reduce la evasion un 25%
        If .flags.Meditando Then
            ProbEvadir = (100 - ProbExito) * 0.75
            ProbExito = MinimoInt(100, 100 - ProbEvadir)
        End If
        
        UsuarioImpacto = (RandomNumber(1, 100) <= ProbExito)
        
        ' el usuario esta usando un escudo ???
        If .Invent.EscudoEqpObjIndex > 0 Then
            'Fallo ???
            If Not UsuarioImpacto Then
                
                Dim SumaSkills As Integer
                
                ' Para evitar division por 0
                SumaSkills = MaximoInt(1, SkillDefensa + SkillTacticas)
                
                ' Chances are rounded
                ProbRechazo = MaximoInt(10, MinimoInt(90, 100 * SkillDefensa / SumaSkills))
                Rechazo = (RandomNumber(1, 100) <= ProbRechazo)
                If Rechazo Then
                    'Se rechazo el ataque con el escudo
                    Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessagePlayWave(ConstantesSonidos.Escudo, .Pos.X, .Pos.Y, .Char.CharIndex))
                      
                    Call WriteMultiMessage(AtacanteIndex, eMessages.BlockedWithShieldother)
                    Call WriteMultiMessage(VictimaIndex, eMessages.BlockedWithShieldUser)
                    
                    Call SubirSkill(VictimaIndex, eSkill.Defensa, True)
                Else
                    Call SubirSkill(VictimaIndex, eSkill.Defensa, False)
                End If
            End If
        End If
        
        If Not UsuarioImpacto Then
            Call SubirSkill(AtacanteIndex, skill, False)
        End If
    End With
    
    Exit Function
    
ErrHandler:
    Dim AtacanteNick As String
    Dim VictimaNick As String
    
    If AtacanteIndex > 0 Then AtacanteNick = UserList(AtacanteIndex).Name
    If VictimaIndex > 0 Then VictimaNick = UserList(VictimaIndex).Name
    
    Call LogError("Error en UsuarioImpacto. Error " & Err.Number & " : " & Err.Description & " AtacanteIndex: " & _
             AtacanteIndex & " Nick: " & AtacanteNick & " VictimaIndex: " & VictimaIndex & " Nick: " & VictimaNick)
End Function

Public Function UsuarioAtacaUsuario(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer, ByVal IsFirstHitFromArea As Boolean) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 03/08/2012
'14/01/2010: ZaMa - Lo transformo en función, para que no se pierdan municiones al atacar targets
'                    inválidos, y evitar un doble chequeo innecesario
'24/05/2011: Amraphen - Ahora se envía la animación del user al golpear.
'03/08/2012: ZaMa - Ahora apuñlar sube skills como fallido al fallar el golpe.
'***************************************************

On Error GoTo ErrHandler

    If Not PuedeAtacar(AtacanteIndex, VictimaIndex) Then Exit Function
    
    With UserList(AtacanteIndex)
        If Distancia(.Pos, UserList(VictimaIndex).Pos) > MAXDISTANCIAARCO Then
           Call WriteConsoleMsg(AtacanteIndex, "Estás muy lejos para disparar.", FontTypeNames.FONTTYPE_FIGHT)
           Exit Function
        End If
        
        Call UsuarioAtacadoPorUsuario(AtacanteIndex, VictimaIndex)
        
        If UsuarioImpacto(AtacanteIndex, VictimaIndex) Then
            'Send animation
            Call SendCharacterSwing(AtacanteIndex)
        
            Call SendData(SendTarget.ToPCArea, AtacanteIndex, PrepareMessagePlayWave(ConstantesSonidos.Impacto, .Pos.X, .Pos.Y, .Char.CharIndex))
            
            If UserList(VictimaIndex).flags.Navegando = 0 Then
                Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessageCreateFX(UserList(VictimaIndex).Char.CharIndex, ConstantesFX.FxSangre, 0))
            End If
               
            'y ahora, el ladrón puede llegar a paralizar con el golpe.
            If .Clase = eClass.Thief Then
                Call DoHandInmo(AtacanteIndex, VictimaIndex)
            End If
            
            Call SubirSkill(VictimaIndex, eSkill.Tacticas, False)
            Call UserDanioUser(AtacanteIndex, VictimaIndex, IsFirstHitFromArea)
        Else
            
            ' Si apuñala sube como fallido
            If PuedeApuñalar(AtacanteIndex) Then
                Call SubirSkill(AtacanteIndex, eSkill.Apuñalar, False)
            End If
                    
            ' Invisible admins doesn't make sound to other clients except itself
            If .flags.AdminInvisible = 1 Then
                Call SendData(ToUser, AtacanteIndex, PrepareMessagePlayWave(ConstantesSonidos.Swing, .Pos.X, .Pos.Y, .Char.CharIndex))
            Else
                'Send animation
                Call SendCharacterSwing(AtacanteIndex)
                
                Call SendData(SendTarget.ToPCArea, AtacanteIndex, PrepareMessagePlayWave(ConstantesSonidos.Swing, .Pos.X, .Pos.Y, .Char.CharIndex))
            End If
            
            Call WriteMultiMessage(AtacanteIndex, eMessages.UserSwing)
            Call WriteMultiMessage(VictimaIndex, eMessages.UserAttackedSwing, AtacanteIndex)
            Call SubirSkill(VictimaIndex, eSkill.Tacticas, True)
        End If
               
    End With
    
    UsuarioAtacaUsuario = True
    
    Exit Function
    
ErrHandler:
    Call LogError("Error en UsuarioAtacaUsuario. Error " & Err.Number & " : " & Err.Description)
End Function

Public Sub UserDanioUser(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer, ByVal IsFirstHitFromArea As Boolean)
'***************************************************
'Author: Unknown
'Last Modification: 03/06/2011 (Amraphen)
'12/01/2010: ZaMa - Implemento armas arrojadizas y probabilidad de acuchillar.
'11/03/2010: ZaMa - Ahora no cuenta la muerte si estaba en estado atacable, y no se vuelve criminal.
'18/09/2010: ZaMa - Ahora se cosidera la defensa de los barcos siempre.
'03/06/2011: Amraphen - Agrego defensa adicional de armadura de segunda jerarquía.
'***************************************************
    
On Error GoTo ErrHandler

    Dim Danio As Long
    Dim Lugar As Byte
    Dim Obj As ObjData
    
    Dim BoatDefense As Integer
    Dim BodyDefense As Integer
    Dim HeadDefense As Integer
    Dim WeaponBoost As Integer
    Dim FinalDefense As Integer
    Dim Perforation As Integer
    
    Dim BoatIndex As Integer
    Dim WeaponIndex As Integer
    Dim HelmetIndex As Integer
    Dim ArmourIndex As Integer
    Dim ShieldIndex As Integer
    
    Danio = CalculateDamage(AtacanteIndex, 0, IsFirstHitFromArea)
    Perforation = 0
    
    Call UserEnvenena(AtacanteIndex, VictimaIndex)
    
    With UserList(AtacanteIndex)
        
        ' Aumento de danio por barca (atacante)
        If .flags.Navegando = 1 Then
            
            BoatIndex = .Invent.BarcoObjIndex
            
            If BoatIndex > 0 Then
                Obj = ObjData(BoatIndex)
                Danio = Danio + RandomNumber(Obj.MinHit, Obj.MaxHit)
                Perforation = Perforation + Obj.Perforation
            End If
            
        End If
        
        ' Aumento de defensa por barca (victima)
        If UserList(VictimaIndex).flags.Navegando = 1 Then
            
            BoatIndex = UserList(VictimaIndex).Invent.BarcoObjIndex
            
            If BoatIndex > 0 Then
                Obj = ObjData(BoatIndex)
                BoatDefense = RandomNumber(Obj.MinDef, Obj.MaxDef)
            End If
            
        End If
        
        ' Perforation of the attacker's weapon
        WeaponIndex = .Invent.WeaponEqpObjIndex
        If WeaponIndex > 0 Then
            Perforation = Perforation + ObjData(WeaponIndex).Perforation
        End If
        
        ' Perforation of the attacker's ammunition
        If .Invent.MunicionEqpObjIndex > 0 Then
            Perforation = Perforation + ObjData(.Invent.MunicionEqpObjIndex).Perforation
        End If
        
        
        ' Now we have a 10% chance of hitting on the head.
        If RandomNumber(1, 10) = 10 Then
            Lugar = PartesCuerpo.bCabeza
        Else
            Lugar = RandomNumber(PartesCuerpo.bPiernaIzquierda, PartesCuerpo.bTorso)
        End If
        
        Select Case Lugar
        
            Case PartesCuerpo.bCabeza
            
                'Si tiene casco absorbe el golpe
                HelmetIndex = UserList(VictimaIndex).Invent.CascoEqpObjIndex
                If HelmetIndex > 0 Then
                    Obj = ObjData(HelmetIndex)
                    HeadDefense = RandomNumber(Obj.MinDef, Obj.MaxDef)
                End If
            
            Case Else
                
                Dim MinDef As Integer
                Dim MaxDef As Integer
                
                'Si tiene armadura absorbe el golpe
                ArmourIndex = UserList(VictimaIndex).Invent.ArmourEqpObjIndex
                If ArmourIndex > 0 Then
                    Obj = ObjData(ArmourIndex)
                    MinDef = Obj.MinDef
                    MaxDef = Obj.MaxDef
                End If
                
                'Si tiene armadura de segunda jerarquía obtiene un porcentaje de defensa adicional.
                If UserList(VictimaIndex).Invent.FactionArmourEqpObjIndex > 0 Then
                    MinDef = MinDef * ConstantesBalance.ModDefSegJerarquia
                    MaxDef = MaxDef * ConstantesBalance.ModDefSegJerarquia
                End If
                
                ' Si tiene escudo, tambien absorbe el golpe
                ShieldIndex = UserList(VictimaIndex).Invent.EscudoEqpObjIndex
                If ShieldIndex > 0 Then
                    Obj = ObjData(ShieldIndex)
                    MinDef = MinDef + Obj.MinDef
                    MaxDef = MaxDef + Obj.MaxDef
                End If
                
                BodyDefense = RandomNumber(MinDef, MaxDef)
        End Select
        
        ' Substract the perforation and make the final defense value.
        FinalDefense = MaximoInt((HeadDefense + BodyDefense + BoatDefense) - Perforation, 0)
        
        Danio = Max(1, Danio - FinalDefense)
              
        Call WriteMultiMessage(AtacanteIndex, eMessages.UserHittedUser, UserList(VictimaIndex).Char.CharIndex, Lugar, Danio)
        Call WriteMultiMessage(VictimaIndex, eMessages.UserHittedByUser, .Char.CharIndex, Lugar, Danio)
        
        UserList(VictimaIndex).Stats.MinHp = UserList(VictimaIndex).Stats.MinHp - Danio
        
        ' Stab the enemy
        If PuedeApuñalar(AtacanteIndex) Then
            Call DoStab(AtacanteIndex, 0, VictimaIndex, Danio, (Lugar = PartesCuerpo.bCabeza))
        End If
        
        ' Si acuchilla
        If PuedeAcuchillar(AtacanteIndex) Then
            Call DoAcuchillar(AtacanteIndex, 0, VictimaIndex, Danio)
        End If
        
        'e intenta dar un golpe crítico [Pablo (ToxicWaste)]
        Call DoGolpeCritico(AtacanteIndex, 0, VictimaIndex, Danio)
        
        If .flags.Hambre = 0 And .flags.Sed = 0 Then
            'Si usa un arma quizas suba "Combate con armas"
            If WeaponIndex > 0 Then
                If ObjData(WeaponIndex).proyectil Then
                    'es un Arco. Sube Armas a Distancia
                    Call SubirSkill(AtacanteIndex, eSkill.Proyectiles, True)
                Else
                    'Sube combate con armas.
                    Call SubirSkill(AtacanteIndex, eSkill.Armas, True)
                End If
            Else
                'sino tal vez lucha libre
                Call SubirSkill(AtacanteIndex, eSkill.Wrestling, True)
            End If

        End If
        
        If UserList(VictimaIndex).Stats.MinHp <= 0 Then
            
            ' No cuenta la muerte si estaba en estado atacable
            If UserList(VictimaIndex).flags.AtacablePor <> AtacanteIndex Then
                'Store it!
                Call Statistics.StoreFrag(AtacanteIndex, VictimaIndex)
                Call ContarMuerte(VictimaIndex, AtacanteIndex, IIf(WeaponIndex > 0, eDamageType.Weapon, eDamageType.BareHand), Danio, WeaponIndex)
            End If
            
            ' Para que las mascotas no sigan intentando luchar y
            ' comiencen a seguir al amo
            Dim J As Integer
            For J = 1 To Classes(.clase).ClassMods.MaxTammedPets
                ' Tammed pets
                If .TammedPets(J).NpcIndex > 0 Then
                    If Npclist(.TammedPets(J).NpcIndex).Target = VictimaIndex Then
                        Npclist(.TammedPets(J).NpcIndex).Target = 0
                        Call FollowAmo(.TammedPets(J).NpcIndex)
                    End If
                End If
            Next J
            
            For J = 1 To Classes(.clase).ClassMods.MaxInvokedPets
                ' Invoked Pets
                If .InvokedPets(J).NpcIndex > 0 Then
                    If Npclist(.InvokedPets(J).NpcIndex).Target = VictimaIndex Then
                        Npclist(.InvokedPets(J).NpcIndex).Target = 0
                        Call FollowAmo(.InvokedPets(J).NpcIndex)
                    End If
                End If
            Next J
            
            Call ActStats(VictimaIndex, AtacanteIndex)
            Call UserDie(VictimaIndex)
        Else
            'Está vivo - Actualizamos el HP
            Call WriteUpdateHP(VictimaIndex)
        End If
    End With

    Exit Sub
    
ErrHandler:
    Dim AtacanteNick As String
    Dim VictimaNick As String
    
    If AtacanteIndex > 0 Then AtacanteNick = UserList(AtacanteIndex).Name
    If VictimaIndex > 0 Then VictimaNick = UserList(VictimaIndex).Name
    
    Call LogError("Error en UserDanioUser. Error " & Err.Number & " : " & Err.Description & " AtacanteIndex: " & _
             AtacanteIndex & " Nick: " & AtacanteNick & " VictimaIndex: " & VictimaIndex & " Nick: " & VictimaNick)
End Sub
Sub UsuarioAtacadoPorUsuario(ByVal AttackerIndex As Integer, ByVal VictimIndex As Integer)
'***************************************************
'Autor: Unknown
'Last Modification: 05/05/2010
'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
'10/01/2008: Tavo - Se cancela la salida del juego si el user esta saliendo
'05/05/2010: ZaMa - Ahora no suma puntos de bandido al atacar a alguien en estado atacable.
'***************************************************
On Error GoTo ErrHandler
  
    If TriggerZonaPelea(AttackerIndex, VictimIndex) = TRIGGER6_PERMITE Then Exit Sub
    
    With UserList(VictimIndex)
        If .flags.Meditando Then
            .flags.Meditando = False
            Call WriteMeditateToggle(VictimIndex)
            Call WriteConsoleMsg(VictimIndex, "Dejas de meditar.", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)
            .Char.FX = 0
            .Char.Loops = 0
            Call SendData(SendTarget.ToPCArea, VictimIndex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0, , 0))
        End If
    End With
    
    If UserList(VictimIndex).flags.Muerto = 0 Then
        Call AllMascotasAtacanUser(AttackerIndex, VictimIndex)
        Call AllMascotasAtacanUser(VictimIndex, AttackerIndex)
    End If
    
    'Si la victima esta saliendo se cancela la salida
    Call CancelExit(VictimIndex)

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub UsuarioAtacadoPorUsuario de SistemaCombate.bas")
End Sub

Sub AllMascotasAtacanUser(ByVal victim As Integer, ByVal Maestro As Integer)
On Error GoTo ErrHandler
  
    Dim iCount As Integer
    
    With UserList(Maestro)
        For iCount = 1 To Classes(.clase).ClassMods.MaxTammedPets
            ' Tammed pets
            If .TammedPets(iCount).NpcIndex > 0 Then
                Npclist(.TammedPets(iCount).NpcIndex).flags.AttackedBy = UserList(victim).Name
                Npclist(.TammedPets(iCount).NpcIndex).Movement = TipoAI.NPCDEFENSA
                Npclist(.TammedPets(iCount).NpcIndex).Hostile = 1
            End If
        Next iCount
        
        For iCount = 1 To Classes(.clase).ClassMods.MaxInvokedPets
            'Invoked pets
            If .InvokedPets(iCount).NpcIndex > 0 Then
                Npclist(.InvokedPets(iCount).NpcIndex).flags.AttackedBy = UserList(victim).Name
                Npclist(.InvokedPets(iCount).NpcIndex).Movement = TipoAI.NPCDEFENSA
                Npclist(.InvokedPets(iCount).NpcIndex).Hostile = 1
            End If
        Next iCount
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub AllMascotasAtacanUser de SistemaCombate.bas")
End Sub

Public Function PuedeAtacar(ByVal AttackerIndex As Integer, ByVal VictimIndex As Integer) As Boolean
On Error GoTo ErrHandler
    'MUY importante el orden de estos "IF"...
    
    'Estas muerto no podes atacar
    If UserList(AttackerIndex).flags.Muerto = 1 Then
        Call WriteConsoleMsg(AttackerIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
        PuedeAtacar = False
        Exit Function
    End If
    
    'No podes atacar a alguien muerto
    If UserList(VictimIndex).flags.Muerto = 1 Then
        Call WriteConsoleMsg(AttackerIndex, "No puedes atacar a un espíritu.", FontTypeNames.FONTTYPE_INFO)
        PuedeAtacar = False
        Exit Function
    End If
    
    ' No podes atacar si estas en consulta
    If UserList(AttackerIndex).flags.HelpMode Then
        Call WriteConsoleMsg(AttackerIndex, "No puedes atacar usuarios mientras estas en consulta.", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    
    ' No podes atacar si esta en consulta
    If UserList(VictimIndex).flags.HelpMode Then
        Call WriteConsoleMsg(AttackerIndex, "No puedes atacar usuarios mientras estan en consulta.", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If

    If EnMapaDuelos(AttackerIndex) Then
        If UserList(AttackerIndex).flags.DueloIndex > 0 Then
            If DuelData.Duelo(UserList(AttackerIndex).flags.DueloIndex).estado = eDuelState.Esperando_Final Then
                Call WriteConsoleMsg(AttackerIndex, "¡El duelo ya ha finalizado!", FontTypeNames.FONTTYPE_INFO)
                Exit Function
            End If
            
            If Not DuelData.Duelo(UserList(AttackerIndex).flags.DueloIndex).estado = eDuelState.Iniciado Then
                Call WriteConsoleMsg(AttackerIndex, "¡El duelo aun no ha iniciado!", FontTypeNames.FONTTYPE_INFO)
                Exit Function
            End If
            
            If Not GetTipoDuelo(UserList(AttackerIndex).flags.DueloIndex) = eDuelType.vs1 Then
                If GetUserTeam(UserList(AttackerIndex).flags.DueloIndex, AttackerIndex) = GetUserTeam(UserList(VictimIndex).flags.DueloIndex, VictimIndex) Then
                    Call WriteConsoleMsg(AttackerIndex, "¡No puedes atacar a alguien de tu equipo!", FontTypeNames.FONTTYPE_INFO)
                    Exit Function
                End If
            End If
        End If
    End If
    
    'Estamos en una Arena? o un trigger zona segura?
    Select Case TriggerZonaPelea(AttackerIndex, VictimIndex)
        Case eTrigger6.TRIGGER6_PERMITE
            PuedeAtacar = (UserList(VictimIndex).flags.AdminInvisible = 0)
            Exit Function
        
        Case eTrigger6.TRIGGER6_PROHIBE
            PuedeAtacar = False
            Exit Function
        
        Case eTrigger6.TRIGGER6_AUSENTE
            'Si no estamos en el Trigger 6 entonces es imposible atacar un gm
            If (UserList(VictimIndex).flags.Privilegios And PlayerType.User) = 0 Then
                If UserList(VictimIndex).flags.AdminInvisible = 0 Then Call WriteConsoleMsg(AttackerIndex, "El ser es demasiado poderoso.", FontTypeNames.FONTTYPE_WARNING)
                PuedeAtacar = False
                Exit Function
            End If
    End Select
    
    If Not CanAttackOrStealByAlignment(AttackerIndex, VictimIndex) Then
        Call WriteConsoleMsg(AttackerIndex, "No puedes atacar a un usuario de esa alineación.", FontTypeNames.FONTTYPE_WARNING)
        Exit Function
    End If
    
    If UserList(AttackerIndex).Faccion.Alignment = UserList(VictimIndex).Faccion.Alignment And MapInfo(UserList(VictimIndex).Pos.Map).MismoBando = 1 Then
        Call WriteConsoleMsg(AttackerIndex, "No puedes atacar usuarios de tu misma alineación en este mapa.", FontTypeNames.FONTTYPE_WARNING)
        Exit Function
    End If
        
    'Estas en un Mapa Seguro?
    If MapInfo(UserList(VictimIndex).Pos.Map).Pk = False Then
        If esArmada(AttackerIndex) Then
            If UserList(AttackerIndex).Faccion.RecompensasReal > 11 Then
                If UserList(VictimIndex).Pos.Map = 58 Or UserList(VictimIndex).Pos.Map = 59 Or UserList(VictimIndex).Pos.Map = 60 Then
                Call WriteConsoleMsg(VictimIndex, "¡Huye de la ciudad! Estás siendo atacado y no podrás defenderte.", FontTypeNames.FONTTYPE_WARNING)
                PuedeAtacar = True 'Beneficio de Armadas que atacan en su ciudad.
                Exit Function
                End If
            End If
        End If
        If esCaos(AttackerIndex) Then
            If UserList(AttackerIndex).Faccion.RecompensasCaos > 11 Then
                If UserList(VictimIndex).Pos.Map = 151 Or UserList(VictimIndex).Pos.Map = 156 Then
                Call WriteConsoleMsg(VictimIndex, "¡Huye de la ciudad! Estás siendo atacado y no podrás defenderte.", FontTypeNames.FONTTYPE_WARNING)
                PuedeAtacar = True 'Beneficio de Caos que atacan en su ciudad.
                Exit Function
                End If
            End If
        End If
        Call WriteConsoleMsg(AttackerIndex, "Esta es una zona segura, aquí no puedes atacar a otros usuarios.", FontTypeNames.FONTTYPE_WARNING)
        PuedeAtacar = False
        Exit Function
    End If
    
    'Estas atacando desde un trigger seguro? o tu victima esta en uno asi?
    If MapData(UserList(VictimIndex).Pos.Map, UserList(VictimIndex).Pos.X, UserList(VictimIndex).Pos.Y).Trigger = eTrigger.ZONASEGURA Or _
        MapData(UserList(AttackerIndex).Pos.Map, UserList(AttackerIndex).Pos.X, UserList(AttackerIndex).Pos.Y).Trigger = eTrigger.ZONASEGURA Then
        Call WriteConsoleMsg(AttackerIndex, "No puedes pelear aquí.", FontTypeNames.FONTTYPE_WARNING)
        PuedeAtacar = False
        Exit Function
    End If
    
    If FriendlyFireProtectionEnabled(AttackerIndex, VictimIndex) And TriggerZonaPelea(AttackerIndex, VictimIndex) <> TRIGGER6_PERMITE Then
        Call WriteConsoleMsg(AttackerIndex, "No puedes atacar a un miembro de tu mismo Clan.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
        PuedeAtacar = False
        Exit Function
    End If
    
    PuedeAtacar = True
Exit Function

ErrHandler:
    Call LogError("Error en PuedeAtacar. Error " & Err.Number & " : " & Err.Description)
End Function

Public Function PuedeAtacarNPC(ByVal AttackerIndex As Integer, ByVal NpcIndex As Integer, _
                Optional ByVal Paraliza As Boolean = False, Optional ByVal IsPet As Boolean = False) As Boolean
On Error GoTo ErrHandler

    With Npclist(NpcIndex)
    
        'Estas muerto?
        If UserList(AttackerIndex).flags.Muerto = 1 Then
            Call WriteConsoleMsg(AttackerIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        
        'Sos consejero?
        If UserList(AttackerIndex).flags.Privilegios And PlayerType.Consejero Then
            'No pueden atacar NPC los Consejeros.
            Exit Function
        End If
        
        ' No podes atacar si estas en consulta
        If UserList(AttackerIndex).flags.HelpMode Then
            Call WriteConsoleMsg(AttackerIndex, "No puedes atacar npcs mientras estas en consulta.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        
        'Es una criatura atacable?
        If .Attackable = 0 Then
            Call WriteConsoleMsg(AttackerIndex, "No puedes atacar esta criatura.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        
        'Es valida la distancia a la cual estamos atacando?
        If Distancia(UserList(AttackerIndex).Pos, .Pos) >= MAXDISTANCIAARCO Then
           Call WriteConsoleMsg(AttackerIndex, "Estás muy lejos para disparar.", FontTypeNames.FONTTYPE_FIGHT)
           Exit Function
        End If
        
        'Es una criatura No-Hostil?
        If .Hostile = 0 Then
            'Es Guardia del Caos?
            If .NPCtype = eNPCType.GuardiasCaos And esCaos(AttackerIndex) Then
                'Lo quiere atacar un caos?
                If esCaos(AttackerIndex) Then
                    Call WriteConsoleMsg(AttackerIndex, "No puedes atacar Guardias del Caos siendo de la legión oscura.", FontTypeNames.FONTTYPE_INFO)
                    Exit Function
                End If
            'Es guardia Real?
            ElseIf .NPCtype = eNPCType.GuardiaReal Then
                'Lo quiere atacar un Armada?
                If esArmada(AttackerIndex) Then
                    Call WriteConsoleMsg(AttackerIndex, "No puedes atacar Guardias Reales siendo del ejército real.", FontTypeNames.FONTTYPE_INFO)
                    Exit Function
                End If
            End If
        End If
    
    
        Dim MasterIndex As Integer
        MasterIndex = .MaestroUser
        
        'Es el NPC mascota de alguien?
        If MasterIndex > 0 Then
            
            'La mascotas es de un miembro del clan y tiene el upgrade activo
            'AttackerIndex
            If FriendlyFireProtectionEnabled(AttackerIndex, MasterIndex) And TriggerZonaPelea(AttackerIndex, MasterIndex) <> TRIGGER6_PERMITE Then
                Call WriteConsoleMsg(AttackerIndex, "No puedes atacar a una mascota o invocacion de un miembro de tu mismo Clan.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                Exit Function
            End If
            
            If UserList(AttackerIndex).Faccion.Alignment <> eCharacterAlignment.Neutral And (UserList(AttackerIndex).Faccion.Alignment = UserList(MasterIndex).Faccion.Alignment) Then
                ' Members of the same faction cannot attack each other pets.
                Call WriteConsoleMsg(AttackerIndex, "No puedes atacar una mascota de una persona de tu misma facción.", FontTypeNames.FONTTYPE_INFO)
                Exit Function
            End If

            Dim OwnerUserIndex As Integer
            
          
         
        End If
    End With
    
    'Es el Rey Preatoriano?
    If Npclist(NpcIndex).NPCtype = eNPCType.Pretoriano Then
        If Not ClanPretoriano(Npclist(NpcIndex).ClanIndex).CanAtackMember(NpcIndex) Then
            Call WriteConsoleMsg(AttackerIndex, "Debes matar al resto del ejército antes de atacar al rey.", FontTypeNames.FONTTYPE_FIGHT)
            Exit Function
        End If
    End If
    
    PuedeAtacarNPC = True
        
    Exit Function
        
ErrHandler:
    
    Dim AtckName As String
    Dim OwnerName As String

    If AttackerIndex > 0 Then AtckName = UserList(AttackerIndex).Name
    If OwnerUserIndex > 0 Then OwnerName = UserList(OwnerUserIndex).Name
    
    Call LogError("Error en PuedeAtacarNpc. Erorr: " & Err.Number & " - " & Err.Description & " Atacante: " & _
                   AttackerIndex & "-> " & AtckName & ". Owner: " & OwnerUserIndex & "-> " & OwnerName & _
                   ". NpcIndex: " & NpcIndex & ".")
End Function

Private Function SameClan(ByVal UserIndex As Integer, ByVal OtherUserIndex As Integer) As Boolean
'***************************************************
'Autor: ZaMa
'Returns True if both players belong to the same clan.
'Last Modification: 16/11/2009
'***************************************************
On Error GoTo ErrHandler
  
    SameClan = (UserList(UserIndex).Guild.IdGuild = UserList(OtherUserIndex).Guild.IdGuild) And _
                UserList(UserIndex).Guild.IdGuild <> 0
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function SameClan de SistemaCombate.bas")
End Function

Private Function SameParty(ByVal UserIndex As Integer, ByVal OtherUserIndex As Integer) As Boolean
'***************************************************
'Autor: ZaMa
'Returns True if both players belong to the same party.
'Last Modification: 16/11/2009
'***************************************************
On Error GoTo ErrHandler
  
    SameParty = UserList(UserIndex).PartyIndex = UserList(OtherUserIndex).PartyIndex And _
                UserList(UserIndex).PartyIndex <> 0
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function SameParty de SistemaCombate.bas")
End Function

Sub CalcularDarExp(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal ElDaño As Long)
'***************************************************
'Autor: Nacho (Integer)
'Last Modification: 03/09/06 Nacho
'Reescribi gran parte del Sub
'Ahora, da toda la experiencia del npc mientras este vivo.
'***************************************************
On Error GoTo ErrHandler
  
    Dim ExpaDar As Long
    Dim nExperience As Long
    
    '[Nacho] Chekeamos que las variables sean validas para las operaciones
    If ElDaño <= 0 Then ElDaño = 0
    If Npclist(NpcIndex).Stats.MaxHp <= 0 Then Exit Sub
    If ElDaño > Npclist(NpcIndex).Stats.MinHp Then ElDaño = Npclist(NpcIndex).Stats.MinHp
    
    ' Pets should not give exp.
    If Npclist(NpcIndex).MaestroUser > 0 Then Exit Sub
    
    '[Nacho] La experiencia a dar es la porcion de vida quitada * toda la experiencia
    ExpaDar = CLng(ElDaño * (Npclist(NpcIndex).GiveEXP / Npclist(NpcIndex).Stats.MaxHp))
    If ExpaDar <= 0 Then Exit Sub
    
    '[Nacho] Vamos contando cuanta experiencia sacamos, porque se da toda la que no se dio al user que mata al NPC
            'Esto es porque cuando un elemental ataca, no se da exp, y tambien porque la cuenta que hicimos antes
            'Podria dar un numero fraccionario, esas fracciones se acumulan hasta formar enteros ;P
    If ExpaDar > Npclist(NpcIndex).flags.ExpCount Then
        ExpaDar = Npclist(NpcIndex).flags.ExpCount
        Npclist(NpcIndex).flags.ExpCount = 0
    Else
        Npclist(NpcIndex).flags.ExpCount = Npclist(NpcIndex).flags.ExpCount - ExpaDar
    End If
    
    '[Nacho] Le damos la exp al user
    If ExpaDar > 0 Then
        ExpaDar = ApplyExperienceModifier(UserIndex, NpcIndex, ExpaDar)
        
        If UserList(UserIndex).PartyIndex > 0 Then
            Call mdParty.ObtenerExito(UserIndex, ExpaDar, Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y)
        Else
            UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + ExpaDar
            Call WriteConsoleMsg(UserIndex, "Has ganado " & ExpaDar & " puntos de experiencia.", _
                                 FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
        End If
        
        Call CheckUserLevel(UserIndex)
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CalcularDarExp de SistemaCombate.bas")
End Sub

Public Function TriggerZonaPelea(ByVal Origen As Integer, ByVal Destino As Integer) As eTrigger6
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

'TODO: Pero que rebuscado!!
'Nigo:  Te lo rediseñe, pero no te borro el TODO para que lo revises.
On Error GoTo ErrHandler
    Dim tOrg As eTrigger
    Dim tDst As eTrigger
    
    tOrg = MapData(UserList(Origen).Pos.Map, UserList(Origen).Pos.X, UserList(Origen).Pos.Y).Trigger
    tDst = MapData(UserList(Destino).Pos.Map, UserList(Destino).Pos.X, UserList(Destino).Pos.Y).Trigger
    
    If tOrg = eTrigger.ZONAPELEA Or tDst = eTrigger.ZONAPELEA Then
        If tOrg = tDst Then
            TriggerZonaPelea = TRIGGER6_PERMITE
        Else
            TriggerZonaPelea = TRIGGER6_PROHIBE
        End If
    Else
        TriggerZonaPelea = TRIGGER6_AUSENTE
    End If

Exit Function
ErrHandler:
    TriggerZonaPelea = TRIGGER6_AUSENTE
    LogError ("Error en TriggerZonaPelea - " & Err.Description)
End Function

Sub UserEnvenena(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    Dim ObjInd As Integer
    
    ObjInd = UserList(AtacanteIndex).Invent.WeaponEqpObjIndex
    
    If ObjInd > 0 Then
        If ObjData(ObjInd).proyectil = 1 Then
            ObjInd = UserList(AtacanteIndex).Invent.MunicionEqpObjIndex
        End If
        
        If ObjInd > 0 Then
            If ObjData(ObjInd).Envenena = 1 Then
                
                If RandomNumber(1, 100) < 60 Then
                    UserList(VictimaIndex).flags.Envenenado = 1
                    Call WriteConsoleMsg(VictimaIndex, "¡¡" & UserList(AtacanteIndex).Name & " te ha envenenado!!", _
                                         FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                    Call WriteConsoleMsg(AtacanteIndex, "¡¡Has envenenado a " & UserList(VictimaIndex).Name & "!!", _
                                         FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                End If
            End If
        End If
    End If

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub UserEnvenena de SistemaCombate.bas")
End Sub

Public Sub LanzarProyectil(ByVal UserIndex As Integer, ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Autor: ZaMa
'Last Modification: 10/07/2010
'Throws an arrow or knive to target user/npc.
'***************************************************
On Error GoTo ErrHandler

    Dim MunicionSlot As Byte
    Dim MunicionIndex As Integer
    Dim WeaponSlot As Byte
    Dim WeaponIndex As Integer

    Dim TargetUserIndex As Integer
    Dim TargetNpcIndex As Integer

    Dim DummyInt As Integer
    
    Dim Threw As Boolean
    Threw = True
    
    'Make sure the item is valid and there is ammo equipped.
    With UserList(UserIndex)
        
        With .Invent
            MunicionSlot = .MunicionEqpSlot
            MunicionIndex = .MunicionEqpObjIndex
            WeaponSlot = .WeaponEqpSlot
            WeaponIndex = .WeaponEqpObjIndex
        End With
        
        ' Tiene arma equipada?
        If WeaponIndex = 0 Then
            DummyInt = 1
            Call WriteConsoleMsg(UserIndex, "No tienes un arco o cuchilla equipada.", FontTypeNames.FONTTYPE_INFO)
            
        ' En un slot válido?
        ElseIf WeaponSlot < 1 Or WeaponSlot > .CurrentInventorySlots Then
            DummyInt = 1
            Call WriteConsoleMsg(UserIndex, "No tienes un arco o cuchilla equipada.", FontTypeNames.FONTTYPE_INFO)
            
        ' Usa munición? (Si no la usa, puede ser un arma arrojadiza)
        ElseIf ObjData(WeaponIndex).Municion = 1 Then
        
            ' La municion esta equipada en un slot valido?
            If MunicionSlot < 1 Or MunicionSlot > .CurrentInventorySlots Then
                DummyInt = 1
                Call WriteConsoleMsg(UserIndex, "No tienes municiones equipadas.", FontTypeNames.FONTTYPE_INFO)
                
            ' Tiene munición?
            ElseIf MunicionIndex = 0 Then
                DummyInt = 1
                Call WriteConsoleMsg(UserIndex, "No tienes municiones equipadas.", FontTypeNames.FONTTYPE_INFO)
                
            ' Son flechas?
            ElseIf ObjData(MunicionIndex).ObjType <> eOBJType.otFlechas Then
                DummyInt = 1
                Call WriteConsoleMsg(UserIndex, "No tienes municiones.", FontTypeNames.FONTTYPE_INFO)
                
            ' Tiene suficientes?
            ElseIf .Invent.Object(MunicionSlot).Amount < 1 Then
                DummyInt = 1
                Call WriteConsoleMsg(UserIndex, "No tienes municiones.", FontTypeNames.FONTTYPE_INFO)
            End If
            
        ' Es un arma de proyectiles?
        ElseIf ObjData(WeaponIndex).proyectil <> 1 Then
            DummyInt = 2
        End If
        
        If DummyInt <> 0 Then
            If DummyInt = 1 Then
                Call Desequipar(UserIndex, WeaponSlot, True)
            End If
            
            Call Desequipar(UserIndex, MunicionSlot, True)
            Exit Sub
        End If
        
        Dim RequiredStamina As Integer
        RequiredStamina = GetStaminaRequiredToAttack(UserIndex)
    
        ' Substract the stamina required for attacking
        If .Stats.MinSta < RequiredStamina Then
            Call WriteConsoleMsg(UserIndex, "No tenés suficiente energía para luchar.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Call QuitarSta(UserIndex, RequiredStamina)
        
        ' Saling?
        If .flags.Navegando = 1 Then
            'Hidden?
            If .flags.Oculto > 0 And .flags.AdminInvisible = 0 Then
                ' Thief/Pirate?
                'If .clase = eClass.Thief Then
                '    ' Pierde la apariencia de fragata fantasmal
                '    .flags.Oculto = 0
                '    .Counters.TiempoOculto = 0
                '    Call ToggleBoatBody(UserIndex)
                '    Call WriteConsoleMsg(UserIndex, "¡Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                '    Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.heading, NingunArma, _
                '                        NingunEscudo, NingunCasco)
                'End If
            End If
        End If
        
        Call LookatTile(UserIndex, .Pos.Map, X, Y)
        
        TargetUserIndex = .flags.TargetUser
        TargetNpcIndex = .flags.TargetNPC
        
        'Validate target
        If TargetUserIndex > 0 Then
            'Only allow to atack if the other one can retaliate (can see us)
            If Abs(UserList(TargetUserIndex).Pos.Y - .Pos.Y) > RANGO_VISION_Y Or Abs(UserList(TargetUserIndex).Pos.X - .Pos.X) > RANGO_VISION_X Then
                Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos para atacar.", FontTypeNames.FONTTYPE_WARNING)
                Exit Sub
            End If
            
            'Prevent from hitting self
            If TargetUserIndex = UserIndex Then
                Call WriteConsoleMsg(UserIndex, "¡No puedes atacarte a vos mismo!", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)
                Exit Sub
            End If
            
            'Attack!
            Threw = UsuarioAtacaUsuario(UserIndex, TargetUserIndex, True)
            
        ElseIf TargetNpcIndex > 0 Then
            'Only allow to atack if the other one can retaliate (can see us)
            If Abs(Npclist(TargetNpcIndex).Pos.Y - .Pos.Y) > RANGO_VISION_Y Or Abs(Npclist(TargetNpcIndex).Pos.X - .Pos.X) > RANGO_VISION_X Then
                Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos para atacar.", FontTypeNames.FONTTYPE_WARNING)
                Exit Sub
            End If
            
            'Is it attackable???
            If Npclist(TargetNpcIndex).Attackable <> 0 Then
                'Attack!
                Threw = UsuarioAtacaNpc(UserIndex, TargetNpcIndex, True)
            End If
        End If
        
        ' Solo pierde la munición si pudo atacar al target, o tiro al aire
        If Threw Then
            
            Dim Slot As Byte
            
            ' Tiene equipado arco y flecha?
            If ObjData(WeaponIndex).Municion = 1 Then
                Slot = MunicionSlot
            ' Tiene equipado un arma arrojadiza
            Else
                Slot = WeaponSlot
            End If
            
            'Take 1 knife/arrow away
            Call QuitarUserInvItem(UserIndex, Slot, 1)
        End If
        
    End With
    
    Exit Sub

ErrHandler:

    Dim UserName As String
    Dim Map As Integer
    If UserIndex > 0 Then
        UserName = UserList(UserIndex).Name
        Map = UserList(UserIndex).Pos.Map
    End If

    Call LogError("Error en LanzarProyectil " & Err.Number & ": " & Err.Description & _
                   ". User: " & UserName & "(" & UserIndex & "). Map: " & Map)

End Sub

Public Sub SendCharacterSwing(ByVal UserIndex As Integer)
'***************************************************
'Autor: Amraphen
'Last Modification: 24/05/2011
'Sends the CharacterAttackMovement message to the PC Area
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)
        If Not (.flags.Navegando Or .flags.invisible Or .flags.AdminInvisible) Then _
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterAttackMovement(UserList(UserIndex).Char.CharIndex))
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SendCharacterSwing de SistemaCombate.bas")
End Sub

Public Function GetStaminaRequiredToAttack(ByVal UserIndex As Integer) As Integer
    ' Consumes energy based on the weapon stamina requirement.
    If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
        GetStaminaRequiredToAttack = ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).RequiredStamina
    Else
        ' Melee attacks require 2 stamina points.
        GetStaminaRequiredToAttack = 2
    End If
End Function


Public Function GetAttackVector(ByVal UserIndex As Integer, ByVal heading As eHeading, ByVal WeaponIndex As Integer) As tAttackPosition()
'***************************************************
' Generate a vector of different attack positions
' if the weapon can generate a splash damage
'***************************************************
    Dim AttackVector() As tAttackPosition
    Dim BaseAttackPos As WorldPos
    Dim SplashDamage As Boolean
    Dim SplashDamageType As Byte
    
    SplashDamage = True
    
    BaseAttackPos = UserList(UserIndex).Pos
    Call HeadtoPos(heading, BaseAttackPos)
    
    ' Check if the user is hitting with its bare hand, or using a weapon
    ' If it's using a weapon, then take the splash damage values
    If WeaponIndex = 0 Or WeaponIndex > UBound(ObjData) Then
        SplashDamage = False
        SplashDamageType = eSplashDamageType.None
    Else
        SplashDamage = ObjData(WeaponIndex).SplashDamage
        SplashDamageType = ObjData(WeaponIndex).SplashDamageType
    End If
    
    ' Splash damage Attack.
    Select Case SplashDamageType
        Case eSplashDamageType.Swing
            ' This example shows how the vector will be calculated if the user is heading to
            ' the East, where X is the attacker and 1, 2 and 3 will be the target positions
            '   O 2
            '   X 1
            '   O 3
            '
            ReDim AttackVector(1 To 3) As tAttackPosition
            AttackVector(1).Pos = BaseAttackPos
            
            AttackVector(2).Pos = HeadToPosLateral(heading, -1, BaseAttackPos)
            AttackVector(2).ReducedDamageFromSplash = True
            
            AttackVector(3).Pos = HeadToPosLateral(heading, 1, BaseAttackPos)
            AttackVector(3).ReducedDamageFromSplash = True
            
        Case eSplashDamageType.Lance
            ReDim AttackVector(1 To 2) As tAttackPosition
            ' This example shows how the vector will be calculated if the user is heading to
            ' the East, where X is the attacker and 1 and 2 will be the target positions
            '   O O O
            '   X 1 2
            '   O O O
            AttackVector(1).Pos = BaseAttackPos
            
            AttackVector(2).Pos = BaseAttackPos
            AttackVector(2).ReducedDamageFromSplash = True
            
            Call HeadtoPos(heading, AttackVector(2).Pos)
        Case eSplashDamageType.Pike
            ' This example shows how the vector will be calculated if the user is heading to
            ' the East, where X is the attacker and 1, 2, 3 and 4 will be the target positions
            '   O 2 O
            '   X 1 4
            '   O 3 O
            ReDim AttackVector(1 To 4) As tAttackPosition
            AttackVector(1).Pos = BaseAttackPos
            
            AttackVector(2).Pos = HeadToPosLateral(heading, -1, BaseAttackPos)
            AttackVector(2).ReducedDamageFromSplash = True
            
            AttackVector(3).Pos = HeadToPosLateral(heading, 1, BaseAttackPos)
            AttackVector(3).ReducedDamageFromSplash = True
            
            AttackVector(4).Pos = BaseAttackPos
            AttackVector(4).ReducedDamageFromSplash = True
            
            Call HeadtoPos(heading, AttackVector(4).Pos)
        Case Else
            ' No splash damage, so we are going to hit only to the base target
            ReDim AttackVector(1 To 1) As tAttackPosition
            AttackVector(1).Pos = BaseAttackPos
            
    End Select

    
    GetAttackVector = AttackVector
End Function

Public Sub DoStab(ByVal UserIndex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal Damage As Long, Optional ByVal head As Byte = 0)
On Error GoTo ErrHandler
  
    Dim StabChance As Integer
    Dim skill As Integer
    Dim TmpDamage As Integer
    Dim StabDamageReduction As Integer
    
    skill = GetSkills(UserIndex, eSkill.Apuñalar)
    
    If VictimUserIndex <> 0 Then
        StabChance = Int(Classes(UserList(UserIndex).clase).ClassMods.StabChance / 2 + skill * Classes(UserList(UserIndex).clase).ClassMods.StabChance / 200)
    Else
        StabChance = Classes(UserList(UserIndex).clase).ClassMods.StabChance
    End If
   
    ' Add a stab chance bonus if the user is invisible and has this mastery enabled
    If UserList(UserIndex).flags.invisible And UserList(UserIndex).Masteries.Boosts.AddStabChanceWhenInviPerc > 0 Then
        StabChance = StabChance + Porcentaje(StabChance, UserList(UserIndex).Masteries.Boosts.AddStabChanceWhenInviPerc)
    End If
    
    ' If there's no luck hitting the enemy, we exit.
    If RandomNumber(1, 100) > StabChance Then
        Call WriteConsoleMsg(UserIndex, "¡No has logrado apuñalar a tu enemigo!", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
        Call SubirSkill(UserIndex, eSkill.Apuñalar, False)
        Exit Sub
    End If
    
    If VictimUserIndex <> 0 Then
        Damage = Round(Damage * Classes(UserList(UserIndex).clase).ClassMods.StabDamageMultiplier, 0)
        
        ' If the Attacker is facing the same direction as the enemy we can safely assume it is doing a backstab
        If UserList(UserIndex).Char.heading = UserList(VictimUserIndex).Char.heading Then
            Damage = Damage + Porcentaje(Damage, UserList(UserIndex).Masteries.Boosts.AddBackstabDamageBonusPerc)
        End If
        
        With UserList(VictimUserIndex)
        
            ' Barco
            If .Invent.BarcoObjIndex > 0 Then
                StabDamageReduction = StabDamageReduction + ObjData(.Invent.BarcoObjIndex).StabDamageReduction
            End If
        
            If head = 0 Then
                ' Escudo
                If .Invent.EscudoEqpObjIndex > 0 Then
                    StabDamageReduction = StabDamageReduction + ObjData(.Invent.EscudoEqpObjIndex).StabDamageReduction
                End If
                
                ' Armadura
                If .Invent.ArmourEqpObjIndex > 0 Then
                    StabDamageReduction = StabDamageReduction + ObjData(.Invent.ArmourEqpObjIndex).StabDamageReduction
                End If
                
                ' Armadura Faccionaria
                If .Invent.FactionArmourEqpObjIndex > 0 Then
                    StabDamageReduction = StabDamageReduction + ObjData(.Invent.FactionArmourEqpObjIndex).StabDamageReduction
                End If
                
            Else
                 ' Casco
                If .Invent.CascoEqpObjIndex > 0 Then
                    StabDamageReduction = StabDamageReduction + ObjData(.Invent.CascoEqpObjIndex).StabDamageReduction
                End If
            End If
            
            Damage = Max(1, Damage - StabDamageReduction)
            
            .Stats.MinHp = .Stats.MinHp - Damage
            
            Call WriteConsoleMsg(UserIndex, "Has apuñalado a " & .Name & " por " & Damage, FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
            Call WriteConsoleMsg(VictimUserIndex, "Te ha apuñalado " & UserList(UserIndex).Name & " por " & Damage, FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
        End With
        
    Else

        Damage = Int(Damage * 2)
        
        ' If the Attacker is facing the same direction as the enemy we can safely assume it is doing a backstab
        If Npclist(VictimNpcIndex).Char.heading = Npclist(VictimNpcIndex).Char.heading Then
            Damage = Damage + Porcentaje(Damage, UserList(UserIndex).Masteries.Boosts.AddBackstabDamageBonusPerc)
        End If
                
        Npclist(VictimNpcIndex).Stats.MinHp = Npclist(VictimNpcIndex).Stats.MinHp - Damage
        Call WriteConsoleMsg(UserIndex, "Has apuñalado a la criatura por " & Damage, FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)

        Call CalcularDarExp(UserIndex, VictimNpcIndex, Damage)
    End If
    
    Call SubirSkill(UserIndex, eSkill.Apuñalar, True)
 
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub DoStab de Trabajo.bas")
End Sub

Public Function CanAttackOrStealByAlignment(ByVal AttackerIndex As Integer, ByVal VictimIndex As Integer) As Boolean
    If UserList(AttackerIndex).Faccion.Alignment < 0 Or UserList(VictimIndex).Faccion.Alignment < 0 Then Exit Function
    
    CanAttackOrStealByAlignment = ConstantesBalance.AlignmentAttackActionMatrix(UserList(AttackerIndex).Faccion.Alignment, UserList(VictimIndex).Faccion.Alignment)
End Function

Public Function CanHelpByAlignment(ByVal AttackerIndex As Integer, ByVal VictimIndex As Integer) As Boolean
    If UserList(AttackerIndex).Faccion.Alignment < 0 Or UserList(VictimIndex).Faccion.Alignment < 0 Then Exit Function
    
    CanHelpByAlignment = ConstantesBalance.AlignmentHelpActionMatrix(UserList(AttackerIndex).Faccion.Alignment, UserList(VictimIndex).Faccion.Alignment)
End Function


