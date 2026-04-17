Attribute VB_Name = "modMasteries"
Option Explicit

Public Enum eSendMasteryType
    ClassMasteries = 1
    CharacterMasteries = 2
End Enum

Public Type tMasterySpellValuePercentPair
    Spell As Integer
    ValuePercent As Integer
End Type

Public Type tMasteryBoost
    Id As Integer
    Name As String
    Description As String
    Enabled As Boolean
    
    IconGrh As Integer
    PointsRequired As Integer
    GoldRequired As Long
    MasteryRequired As Integer
    
    EnableBerserkWhileSailing As Boolean
    
    ' Stats Masteries
    AddMaxHealth As Integer
    AddMaxMana As Integer
    AddMaxManaPerc As Integer
    AddBaseWeaponDamagePercent As Integer
    AddBaseWrestlingDamagePercent As Integer
    AddBaseRangedDamagePercent As Integer
    AddBaseMagicDamagePercent As Integer
    AddEnergyRegeneration As Integer
    AddExtraHitChance As Integer
    '
    
    AddInviMinDuration  As Integer
    AddInviMaxDuration  As Integer
    AddTamingPoints  As Integer
    MagicSpellDamageLeechPerc  As Integer
    MagicSpellManaConversionPerc  As Integer
    AddStabChanceWhenInviPerc As Integer
    AddBackstabDamageBonusPerc As Integer
    AddMagicLifeLeechPerc As Integer
    AddMagicCastPower As Integer
    
    BypassProhibitedClassesQty As Integer
    BypassProhibitedClassesObjs() As Integer
    
    SpellManaCostReductionQty As Integer
    SpellManaCostReduction() As tMasterySpellValuePercentPair
    
    MagicBonusForSpellQty As Integer
    MagicBonusForSpell() As tMasterySpellValuePercentPair
    
    ImmunityToSpellQty As Integer
    ImmunityToSpell() As Integer
        
    CanDissarmWithItemQty As Integer
    CanDissarmWithItem() As Integer

End Type

Public Type tUserMastery
    Id As Integer
    DateAdded As Date
End Type

Public Type tMasteryBoostGroupConfig
    MasteriesQty As Integer
    Masteries() As Integer
End Type

Public Type tUserMasteryBoostGroup
    GroupId As Integer
    MasteriesQty As Integer
    Masteries() As tUserMastery
End Type

Public Type tUserMasteryBoost
    GroupsQty As Integer
    Groups() As tUserMasteryBoostGroup
    
    PendingGroupsQty As Integer
    PendingGroups As tUserMasteryBoostGroup
    
    Boosts As tMasteryBoost
End Type

Public EmptyMasteryElement As tUserMasteryBoost

Public Sub AssignMasteryPropertiesToUser(ByVal UserIndex As Integer, ByVal MasteryID As Integer)

    With UserList(UserIndex)
        If Masteries(MasteryID).Enabled Then
                        
            ' Masteries that impacts based on a list of criterias
            Dim TempValue As Integer
            Dim X As Integer
                        
            ' CanDissarmWithItem list calculation
            If Masteries(MasteryID).CanDissarmWithItemQty > 0 Then
                TempValue = .Masteries.Boosts.CanDissarmWithItemQty
                .Masteries.Boosts.CanDissarmWithItemQty = .Masteries.Boosts.CanDissarmWithItemQty + Masteries(MasteryID).CanDissarmWithItemQty
                ReDim Preserve .Masteries.Boosts.CanDissarmWithItem(1 To .Masteries.Boosts.CanDissarmWithItemQty)
                
                For X = 1 To Masteries(MasteryID).CanDissarmWithItemQty
                    .Masteries.Boosts.CanDissarmWithItem(TempValue + X) = Masteries(MasteryID).CanDissarmWithItem(X)
                Next X
                
            End If
            
            
            ' Finished masteries
            .Masteries.Boosts.AddInviMinDuration = .Masteries.Boosts.AddInviMinDuration + Masteries(MasteryID).AddInviMinDuration
            .Masteries.Boosts.AddInviMaxDuration = .Masteries.Boosts.AddInviMaxDuration + Masteries(MasteryID).AddInviMaxDuration
            
            .Masteries.Boosts.AddMagicLifeLeechPerc = .Masteries.Boosts.AddMagicLifeLeechPerc + Masteries(MasteryID).AddMagicLifeLeechPerc
            .Masteries.Boosts.AddTamingPoints = .Masteries.Boosts.AddTamingPoints + Masteries(MasteryID).AddTamingPoints
            .Masteries.Boosts.AddStabChanceWhenInviPerc = .Masteries.Boosts.AddStabChanceWhenInviPerc + Masteries(MasteryID).AddStabChanceWhenInviPerc
            .Masteries.Boosts.AddBackstabDamageBonusPerc = .Masteries.Boosts.AddBackstabDamageBonusPerc + Masteries(MasteryID).AddBackstabDamageBonusPerc
            .Masteries.Boosts.AddMagicCastPower = .Masteries.Boosts.AddMagicCastPower + Masteries(MasteryID).AddMagicCastPower
            .Masteries.Boosts.MagicSpellDamageLeechPerc = .Masteries.Boosts.MagicSpellDamageLeechPerc + Masteries(MasteryID).MagicSpellDamageLeechPerc
            .Masteries.Boosts.MagicSpellManaConversionPerc = .Masteries.Boosts.MagicSpellManaConversionPerc + Masteries(MasteryID).MagicSpellManaConversionPerc
            
            .Masteries.Boosts.AddMaxHealth = .Masteries.Boosts.AddMaxHealth + Masteries(MasteryID).AddMaxHealth
            .Masteries.Boosts.AddMaxMana = .Masteries.Boosts.AddMaxMana + Masteries(MasteryID).AddMaxMana
            .Masteries.Boosts.AddMaxManaPerc = .Masteries.Boosts.AddMaxManaPerc + Masteries(MasteryID).AddMaxManaPerc
            .Masteries.Boosts.AddBaseWeaponDamagePercent = .Masteries.Boosts.AddBaseWeaponDamagePercent + Masteries(MasteryID).AddBaseWeaponDamagePercent
            .Masteries.Boosts.AddBaseWrestlingDamagePercent = .Masteries.Boosts.AddBaseWrestlingDamagePercent + Masteries(MasteryID).AddBaseWrestlingDamagePercent
            .Masteries.Boosts.AddBaseRangedDamagePercent = .Masteries.Boosts.AddBaseRangedDamagePercent + Masteries(MasteryID).AddBaseRangedDamagePercent
            .Masteries.Boosts.AddBaseMagicDamagePercent = .Masteries.Boosts.AddBaseMagicDamagePercent + Masteries(MasteryID).AddBaseMagicDamagePercent
            .Masteries.Boosts.AddEnergyRegeneration = .Masteries.Boosts.AddEnergyRegeneration + Masteries(MasteryID).AddEnergyRegeneration
            .Masteries.Boosts.AddExtraHitChance = .Masteries.Boosts.AddExtraHitChance + Masteries(MasteryID).AddExtraHitChance
                
            
            ' Masteries that enables certain features
            If Masteries(MasteryID).EnableBerserkWhileSailing Then
                .Masteries.Boosts.EnableBerserkWhileSailing = True
            End If
         
            ' BypassProhibitedClasses list calculation
            If Masteries(MasteryID).BypassProhibitedClassesQty > 0 Then
                TempValue = .Masteries.Boosts.BypassProhibitedClassesQty
                .Masteries.Boosts.BypassProhibitedClassesQty = .Masteries.Boosts.BypassProhibitedClassesQty + Masteries(MasteryID).BypassProhibitedClassesQty
                ReDim Preserve .Masteries.Boosts.BypassProhibitedClassesObjs(1 To .Masteries.Boosts.BypassProhibitedClassesQty)
                
                For X = 1 To Masteries(MasteryID).BypassProhibitedClassesQty
                    .Masteries.Boosts.BypassProhibitedClassesObjs(TempValue + X) = Masteries(MasteryID).BypassProhibitedClassesObjs(X)
                Next X

            End If
            
            If Masteries(MasteryID).SpellManaCostReductionQty > 0 Then
                TempValue = .Masteries.Boosts.SpellManaCostReductionQty
                .Masteries.Boosts.SpellManaCostReductionQty = .Masteries.Boosts.SpellManaCostReductionQty + Masteries(MasteryID).SpellManaCostReductionQty
                ReDim Preserve .Masteries.Boosts.SpellManaCostReduction(1 To .Masteries.Boosts.SpellManaCostReductionQty)
                
                For X = 1 To Masteries(MasteryID).SpellManaCostReductionQty
                    .Masteries.Boosts.SpellManaCostReduction(TempValue + X) = Masteries(MasteryID).SpellManaCostReduction(X)
                Next X

            End If
            
            ' MagicBonusForSpell list calculation
            If Masteries(MasteryID).MagicBonusForSpellQty > 0 Then
                TempValue = .Masteries.Boosts.MagicBonusForSpellQty
                .Masteries.Boosts.MagicBonusForSpellQty = .Masteries.Boosts.MagicBonusForSpellQty + Masteries(MasteryID).MagicBonusForSpellQty
                ReDim Preserve .Masteries.Boosts.MagicBonusForSpell(1 To .Masteries.Boosts.MagicBonusForSpellQty)
                
                For X = 1 To Masteries(MasteryID).MagicBonusForSpellQty
                    .Masteries.Boosts.MagicBonusForSpell(TempValue + X).Spell = Masteries(MasteryID).MagicBonusForSpell(X).Spell
                    .Masteries.Boosts.MagicBonusForSpell(TempValue + X).ValuePercent = Masteries(MasteryID).MagicBonusForSpell(X).ValuePercent
                Next X
        
            End If
            
            ' ImmunityToSpell list calculation
            If Masteries(MasteryID).ImmunityToSpellQty > 0 Then
                TempValue = .Masteries.Boosts.ImmunityToSpellQty
                .Masteries.Boosts.ImmunityToSpellQty = .Masteries.Boosts.ImmunityToSpellQty + Masteries(MasteryID).ImmunityToSpellQty
                ReDim Preserve .Masteries.Boosts.ImmunityToSpell(1 To .Masteries.Boosts.ImmunityToSpellQty)
                
                For X = 1 To Masteries(MasteryID).ImmunityToSpellQty
                    .Masteries.Boosts.ImmunityToSpell(TempValue + X) = Masteries(MasteryID).ImmunityToSpell(X)
                Next X
        
            End If
            
                
        End If
    End With

End Sub

Public Function HasMasteryAssigned(ByVal UserIndex As Integer, ByVal MasteryGroup As Integer, ByVal MasteryID As Integer) As Boolean
On Error GoTo ErrHandler:

    Dim I As Integer
    Dim J As Integer
    
    With UserList(UserIndex)
        For J = 1 To .Masteries.GroupsQty
            For I = 1 To UserList(UserIndex).Masteries.Groups(J).MasteriesQty
                If UserList(UserIndex).Masteries.Groups(J).Masteries(I).Id = MasteryID Then
                    HasMasteryAssigned = True
                    Exit Function
                End If
            Next I
        Next J
    
    End With
    
    
    HasMasteryAssigned = False
        
    Exit Function
    
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HasMasteryAssigned de modMasteries.bas")

End Function


Public Sub AquireMastery(ByVal UserIndex As Integer, ByVal MasteryGroup As Integer, ByVal MasteryID As Integer, Optional ByVal SubstractPoints As Boolean = True)
On Error GoTo ErrHandler:
    Dim MasteryInserted As Boolean
    
    ' Insert the mastery in the database.
    Call AddMasteryDB(UserList(UserIndex).Id, MasteryGroup, MasteryID, Masteries(MasteryID).PointsRequired)
       
    With UserList(UserIndex).Masteries
    
        Dim NewMasteryIndex As Integer
        NewMasteryIndex = .Groups(MasteryGroup).MasteriesQty + 1
        
        'Add the mastery to the player and recalculate everything
        .Groups(MasteryGroup).MasteriesQty = NewMasteryIndex
        ReDim Preserve .Groups(MasteryGroup).Masteries(1 To NewMasteryIndex)
        .Groups(MasteryGroup).Masteries(NewMasteryIndex).Id = MasteryID
        
        ' Add the effects from this mastery
        Call modMasteries.AssignMasteryPropertiesToUser(UserIndex, MasteryID)
    End With
    
    If SubstractPoints Then
        ' Substract mastery points
        If UserList(UserIndex).Stats.MasteryPoints >= Masteries(MasteryID).PointsRequired Then
            UserList(UserIndex).Stats.MasteryPoints = UserList(UserIndex).Stats.MasteryPoints - Masteries(MasteryID).PointsRequired
        Else
            UserList(UserIndex).Stats.MasteryPoints = 0
        End If
        
        ' Substract gold
        If UserList(UserIndex).Stats.GLD >= Masteries(MasteryID).GoldRequired Then
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Masteries(MasteryID).GoldRequired
        Else
            UserList(UserIndex).Stats.GLD = 0
        End If
        
    End If
    
    With Masteries(MasteryID)
    
        If .BypassProhibitedClassesQty Then
            Call UpdateUserInv(True, UserIndex, 0)
        End If
        
        If .AddMaxHealth Then
            ' We need to recalculate the max health
            UserList(UserIndex).Stats.MaxHp = RecalculateCharacterMaxHealth(UserIndex)
        End If
        
        If .AddMaxMana Or .AddMaxManaPerc Then
            ' We need to recalculate the max mana
            UserList(UserIndex).Stats.MaxMan = RecalculateCharacterMaxMana(UserIndex)
        End If
    
    End With
    
    Call WriteUpdateUserStats(UserIndex)
    
    Call WriteConsoleMsg(UserIndex, "Adquiriste la maestría " & Masteries(MasteryID).Name, FontTypeNames.FONTTYPE_TALK)
    
    Exit Sub
    
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub AquireMastery de modMasteries.bas")
  
End Sub



Public Function MasteryAllowToEquipItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean
On Error GoTo ErrHandler:
    
    Dim I As Integer
    
    With UserList(UserIndex)
    
        If .Masteries.Boosts.BypassProhibitedClassesQty = 0 Then
            MasteryAllowToEquipItem = False
            Exit Function
        End If
        
        For I = 1 To .Masteries.Boosts.BypassProhibitedClassesQty
            ' If the class and obj id are specified in one of the masteries, then the user can use the object
            If ObjIndex = .Masteries.Boosts.BypassProhibitedClassesObjs(I) Then
                MasteryAllowToEquipItem = True
                Exit Function
            End If
        Next I
    End With
    
    MasteryAllowToEquipItem = False
    Exit Function
    
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub MasteryAllowToEquipItem de modMasteries.bas")
  
End Function

Public Sub ResetMasteries(ByVal UserIndex As Integer)
On Error GoTo ErrHandler:
    
    
    UserList(UserIndex).Masteries = EmptyMasteryElement
    
    Exit Sub
        
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ResetMasteries de modMasteries.bas")
    
End Sub

Public Function GetMasteryManaReductionPercentForSpell(ByVal UserIndex As Integer, ByVal SpellIndex As Integer) As Integer
        
    With UserList(UserIndex).Masteries.Boosts
        
        ' If there's no mastery assigned, then the reduction is 0%
        If .SpellManaCostReductionQty = 0 Then
            GetMasteryManaReductionPercentForSpell = 0
            Exit Function
        End If
        
        Dim I As Integer
        For I = 1 To .SpellManaCostReductionQty
            If .SpellManaCostReduction(I).Spell = SpellIndex Then
                GetMasteryManaReductionPercentForSpell = GetMasteryManaReductionPercentForSpell + .SpellManaCostReduction(I).ValuePercent
            End If
        Next I
               
    End With
End Function


Public Function GetMasterySpellPowerBonus(ByVal UserIndex As Integer, ByVal SpellIndex As Integer, ByVal BaseDamage As Integer) As Integer
    Dim BonusPerc As Integer
    
     With UserList(UserIndex).Masteries.Boosts
        
        ' If there's no mastery assigned, then the bonus is 0%
        If .MagicBonusForSpellQty = 0 Then
            BonusPerc = 0
            Exit Function
        End If
        
        Dim I As Integer
        For I = 1 To .MagicBonusForSpellQty
            If .MagicBonusForSpell(I).Spell = SpellIndex Then
                BonusPerc = GetMasterySpellPowerBonus + .MagicBonusForSpell(I).ValuePercent
            End If
        Next I
        
        GetMasterySpellPowerBonus = Porcentaje(BaseDamage, BonusPerc)
               
    End With

End Function

Public Function IsUserImmuneToSpell(ByVal CasterIndex As Integer, ByVal UserIndex As Integer, ByVal SpellIndex As Integer, ByVal NotifyCaster As Boolean) As Boolean
    
    With UserList(UserIndex).Masteries.Boosts
        
        ' If there's no mastery assigned, then the bonus is 0%
        If .ImmunityToSpellQty = 0 Then
            IsUserImmuneToSpell = False
            Exit Function
        End If
        
        Dim I As Integer
        For I = 1 To .ImmunityToSpellQty
            If .ImmunityToSpell(I) = SpellIndex Then
                IsUserImmuneToSpell = True
                Exit For
            End If
        Next I
        
        If IsUserImmuneToSpell And NotifyCaster Then
            Call WriteConsoleMsg(CasterIndex, "El usuario es inmune a este hechizo", FontTypeNames.FONTTYPE_INFO)
        End If
               
    End With
    
End Function

