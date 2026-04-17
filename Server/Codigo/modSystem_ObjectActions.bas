Attribute VB_Name = "modSystem_ObjectActions"
Option Explicit
Option Base 1


' This module is intended to hold all actions triggered by using or intacting with an object.
' Some examples could be:
'   * Using potions
'   * Using a bow
'   * Using a tool
'
'
' The orchestration/usage of this functions should be outisde of this module


''
' Uses a potion object
'
' @param UserIndex The UserIndex that triggered the action
' @param Object The Object information
' @param Slot The inventory slot where the item was located
' @return Returns if the object was successfully consumed/used
' @remarks  Must be done after creating the timer and before using it, otherwise, Interval will be 0
Public Function Potions_Use(ByVal UserIndex As Integer, ByRef Obj As ObjData) As Boolean
    Dim ItemConsumed As Boolean
    Dim MinValue As Integer
    Dim MaxValue As Integer
    Dim ObjectConsumed As Integer
    
    With UserList(UserIndex)
        
        .flags.TomoPocion = True
        .flags.TipoPocion = Obj.TipoPocion
        
        ' Affects mana
        If Obj.AffectsMana.Min > 0 Or Obj.AffectsMana.Max > 0 Then
            If Obj.AffectsMana.IsPercent Then
                MinValue = Porcentaje(.Stats.MaxMan, Obj.AffectsMana.Min)
                MaxValue = Porcentaje(.Stats.MaxMan, Obj.AffectsMana.Max)
            Else
                MinValue = Obj.AffectsMana.Min
                MaxValue = Obj.AffectsMana.Max
            End If
            
            .Stats.MinMAN = .Stats.MinMAN + RandomNumber(MinValue, MaxValue)
            If .Stats.MinMAN > .Stats.MaxMan Then _
                .Stats.MinMAN = .Stats.MaxMan
            
            ObjectConsumed = True
        End If

        ' Affects health
        If Obj.AffectsHealth.Min > 0 Or Obj.AffectsHealth.Max > 0 Then
            If Obj.AffectsHealth.IsPercent Then
                MinValue = Porcentaje(.Stats.MaxHp, Obj.AffectsHealth.Min)
                MaxValue = Porcentaje(.Stats.MaxHp, Obj.AffectsHealth.Max)
            Else
                MinValue = Obj.AffectsHealth.Min
                MaxValue = Obj.AffectsHealth.Max
            End If
            
            .Stats.MinHp = .Stats.MinHp + RandomNumber(MinValue, MaxValue)
            If .Stats.MinHp > .Stats.MaxHp Then _
                .Stats.MinHp = .Stats.MaxHp
            
            ObjectConsumed = True
        End If
        
        ' Affects strength
        If Obj.AffectsStrength.Min > 0 Or Obj.AffectsStrength.Max > 0 Then
            .flags.DuracionEfecto = Obj.DuracionEfecto
            
            Dim MaxCharacterStrength As Byte
            MaxCharacterStrength = 2 * .Stats.UserAtributosBackUP(Fuerza)
            
            If Obj.AffectsStrength.IsPercent Then
                MinValue = Porcentaje(MaxCharacterStrength, Obj.AffectsStrength.Min)
                MaxValue = Porcentaje(MaxCharacterStrength, Obj.AffectsStrength.Max)
            Else
                MinValue = Obj.AffectsStrength.Min
                MaxValue = Obj.AffectsStrength.Max
            End If
    
            .Stats.UserAtributos(eAtributos.Fuerza) = .Stats.UserAtributos(eAtributos.Fuerza) + RandomNumber(MinValue, MaxValue)
            If .Stats.UserAtributos(eAtributos.Fuerza) > ConstantesBalance.MaxAtributos Then _
                .Stats.UserAtributos(eAtributos.Fuerza) = ConstantesBalance.MaxAtributos
                
            If .Stats.UserAtributos(eAtributos.Fuerza) > MaxCharacterStrength Then .Stats.UserAtributos(eAtributos.Fuerza) = MaxCharacterStrength
                        
            Call WriteUpdateStrenght(UserIndex)
            Call Potion_BerserkEffect(UserIndex)
            
            ObjectConsumed = True
        End If
        
        ' Affects agility
        If Obj.AffectsAgility.Min > 0 Or Obj.AffectsAgility.Max > 0 Then
            .flags.DuracionEfecto = Obj.DuracionEfecto
            
            Dim MaxCharacterAgility As Byte
            MaxCharacterAgility = 2 * .Stats.UserAtributosBackUP(Agilidad)
            
            If Obj.AffectsAgility.IsPercent Then
                MinValue = Porcentaje(MaxCharacterAgility, Obj.AffectsAgility.Min)
                MaxValue = Porcentaje(MaxCharacterAgility, Obj.AffectsAgility.Max)
            Else
                MinValue = Obj.AffectsAgility.Min
                MaxValue = Obj.AffectsAgility.Max
            End If
        
            'Usa el item
            .Stats.UserAtributos(eAtributos.Agilidad) = .Stats.UserAtributos(eAtributos.Agilidad) + RandomNumber(MinValue, MaxValue)
            If .Stats.UserAtributos(eAtributos.Agilidad) > ConstantesBalance.MaxAtributos Then _
                .Stats.UserAtributos(eAtributos.Agilidad) = ConstantesBalance.MaxAtributos
                
            If .Stats.UserAtributos(eAtributos.Agilidad) > MaxCharacterAgility Then .Stats.UserAtributos(eAtributos.Agilidad) = MaxCharacterAgility
            
            Call WriteUpdateDexterity(UserIndex)
            Call Potion_BerserkEffect(UserIndex)

            ObjectConsumed = True
        End If
                
        If ObjectConsumed = True Then
            'Call QuitarUserInvItem(UserIndex, Slot, 1)
            
            ' Los admin invisibles solo producen sonidos a si mismos
            If .flags.AdminInvisible = 1 Then
                Call SendData(ToUser, UserIndex, PrepareMessagePlayWave(ConstantesSonidos.Tomar, .Pos.X, .Pos.Y, .Char.CharIndex))
            Else
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(ConstantesSonidos.Tomar, .Pos.X, .Pos.Y, .Char.CharIndex))
            End If
        End If
    End With
    
    Potions_Use = ObjectConsumed
End Function


Private Sub Potion_BerserkEffect(ByVal UserIndex As Integer)
    ' Enable the berserk if it mets the requirements.
    If HasPassiveAssigned(UserIndex, ePassiveSpells.Berserk) And Not HasPassiveActivated(UserIndex, ePassiveSpells.Berserk) Then
        If BerzerkConditionMet(UserIndex) Then
            Call ActivatePassive(UserIndex, ePassiveSpells.Berserk, True)
            Call SendBerserkEffect(UserIndex, ePassiveSpells.Berserk, True)
        End If
    End If
    
    ' Enable the Indomitable Will if it mets the requirements.
    If HasPassiveAssigned(UserIndex, IndomitableWill) And Not HasPassiveActivated(UserIndex, IndomitableWill) Then
        If PassiveConditionMet(UserIndex, IndomitableWill) Then
            Call ActivatePassive(UserIndex, IndomitableWill, True)
        End If
    End If
End Sub


