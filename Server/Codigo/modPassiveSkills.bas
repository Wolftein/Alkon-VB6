Attribute VB_Name = "modPassiveSkills"
Option Explicit

Private Const REGENERATION_SKILL As Byte = 20
Private Const PARALYSIS_IMMUNITY_SKILL As Byte = 40
Private Const INDOMITABLE_WILL_SKILL As Byte = 60
Private Const VITAL_RESTORATION_SKILL As Byte = 80
Private Const BERSERK_SKILL As Byte = 100


Public Function HasParalysisImmunity(ByVal UserIndex As Integer) As Boolean
'************************************************************************
'Author: Lucas Figelj(Luke)
'Last Modification: 16/04/2015
'Checks if the user has immobilization immunity when the caster is an NPC
'************************************************************************
On Error GoTo ErrHandler
    Dim HasPassive As Boolean

    With UserList(UserIndex).Stats
        Dim I As Integer
         For I = 1 To MAXUSERPASSIVES
            If .UserPassives(I).ID = ePassiveSpells.ParalysisImmunity Then
                HasPassive = True
            End If
         Next I
         
        'For Each Item In .UserPassives
        '
        '    If Item = ePassiveSpells.ParalysisImmunity Then
        '        HasParalysisImmunity = True
        '    End If
        '
        'Next
        
        If HasPassive = False Then
            HasParalysisImmunity = False
        End If
            'End Function
        
        If HasMaxAttributes(UserIndex) Then
            If HasPassive = True Then
                HasParalysisImmunity = True
            End If
        Else
            HasParalysisImmunity = False
        End If
        
    End With
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function HasParalysisImmunity de modPassiveSkills.bas")
End Function

Public Function BerzerkConditionMet(ByVal TargetIndex As Integer) As Boolean
On Error GoTo ErrHandler
  
    Dim conditionMet As Boolean
    BerzerkConditionMet = False
    With UserList(TargetIndex)
    
        If Not HasPassiveAssigned(TargetIndex, ePassiveSpells.Berserk) Then Exit Function
        
        If .clase <> Warrior And .clase <> Worker Then Exit Function
        
        If .Invent.EscudoEqpSlot > 0 Then Exit Function
        
        If .flags.Navegando And Not .Masteries.Boosts.EnableBerserkWhileSailing Then Exit Function
        
        If Not HasMaxAttributes(TargetIndex) Then Exit Function
         
        If .flags.invisible = 1 Or .flags.Oculto = 1 Then Exit Function
        
        If .Invent.WeaponEqpSlot = 0 Then
            BerzerkConditionMet = True
        Else
            BerzerkConditionMet = ObjData(.Invent.WeaponEqpObjIndex).proyectil = 0
        End If

    
    End With
    
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function BerzerkConditionMet de modPassiveSkills.bas")
End Function

Public Function PassiveConditionMet(ByVal TargetIndex As Integer, ByVal Passive As ePassiveSpells) As Boolean
On Error GoTo ErrHandler
  
    PassiveConditionMet = True
    
    With UserList(TargetIndex)

        Select Case Passive
            ' Berserker Conditions
            Case ePassiveSpells.Berserk
                If .Invent.EscudoEqpSlot = 0 And .flags.Navegando = False And HasMaxAttributes(TargetIndex) = True And .flags.invisible = False Then
                    If .Invent.WeaponEqpSlot = 0 Then
                        PassiveConditionMet = True
                    Else
                        PassiveConditionMet = ObjData(.Invent.WeaponEqpObjIndex).proyectil = 0
                    End If
                Else
                    PassiveConditionMet = False
                End If
                
            Case Else
                PassiveConditionMet = HasPassiveAssigned(TargetIndex, Passive) And HasMaxAttributes(TargetIndex)
            
        End Select
    
    End With
    
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function PassiveConditionMet de modPassiveSkills.bas")
End Function

Public Function ImmunityConditionMet(ByVal TargetIndex As Integer) As Boolean
On Error GoTo ErrHandler
  
    Dim conditionMet As Boolean
    
    With UserList(TargetIndex)
        
        If .Invent.EscudoEqpSlot <> 0 And .flags.Navegando = False And HasMaxAttributes(TargetIndex) = True Then
            If .Invent.WeaponEqpSlot = 0 Then
                conditionMet = True
            Else
                conditionMet = ObjData(.Invent.WeaponEqpObjIndex).proyectil = 0
            End If
        End If
    
    End With

  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function ImmunityConditionMet de modPassiveSkills.bas")
End Function

Function HasMaxAttributes(ByVal TargetIndex As Integer) As Boolean
On Error GoTo ErrHandler
  

    HasMaxAttributes = UserList(TargetIndex).Stats.UserAtributos(eAtributos.Agilidad) = MinimoInt(ConstantesBalance.MaxAtributos, UserList(TargetIndex).Stats.UserAtributosBackUP(Agilidad) * 2) And UserList(TargetIndex).Stats.UserAtributos(eAtributos.Fuerza) = MinimoInt(ConstantesBalance.MaxAtributos, UserList(TargetIndex).Stats.UserAtributosBackUP(Fuerza) * 2)
    
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function HasMaxAttributes de modPassiveSkills.bas")
End Function

Public Function HasIndomitableWill(ByVal TargetIndex As Integer) As Boolean
'************************************************************************
'Author: Lucas Figelj(Luke)
'Last Modification: 22/04/2015
'Checks if the user has indomitable will
'************************************************************************
On Error GoTo ErrHandler
  

Dim HasPassive As Boolean

    With UserList(TargetIndex).Stats
  
        'For Each Item In .UserPassives
        '
        '    If Item = ePassiveSpells.IndomitableWill Then
        '        HasPassive = True
        '    End If
        '
        'Next
        Dim I As Integer
         For I = 1 To MAXUSERPASSIVES
            If .UserPassives(I).ID = ePassiveSpells.ParalysisImmunity Then
                HasPassive = True
            End If
         Next I
    
        If HasPassive = False Then
            HasIndomitableWill = False
        End If
            'End Function
        
        If HasMaxAttributes(TargetIndex) Then
            If HasPassive = True Then
                HasIndomitableWill = True
            End If
        Else
            HasIndomitableWill = False
        End If

    End With
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function HasIndomitableWill de modPassiveSkills.bas")
End Function


Public Function HasPassiveAssigned(ByVal UserIndex As Integer, ByVal Passive As ePassiveSpells) As Boolean
On Error GoTo ErrHandler
  
    If ServerConfiguration.PassiveSkillsQty < 1 Then Exit Function
    
    HasPassiveAssigned = UserList(UserIndex).Stats.UserPassives(Passive).Enabled

  Exit Function
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function HasPassiveAssigned de modPassiveSkills.bas")
End Function

Public Function HasPassiveActivated(ByVal UserIndex As Integer, ByVal Passive As ePassiveSpells) As Boolean
On Error GoTo ErrHandler

    If ServerConfiguration.PassiveSkillsQty < 1 Then Exit Function
  
    HasPassiveActivated = UserList(UserIndex).Stats.UserPassives(Passive).Active
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function HasPassiveActivated de modPassiveSkills.bas")
End Function

Public Sub ActivatePassive(ByVal UserIndex As Integer, ByVal Passive As ePassiveSpells, Optional ByVal Enabled As Boolean = True)
On Error GoTo ErrHandler
  
    UserList(UserIndex).Stats.UserPassives(Passive).Active = Enabled
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub EnablePassive de modPassiveSkills.bas")
End Sub

Public Sub SendBerserkEffect(ByVal UserIndex As Integer, ByVal Passive As ePassiveSpells, ByVal Enabled As Boolean)
On Error GoTo ErrHandler
  
    If Enabled Then
        Call WriteConsoleMsg(UserIndex, "La habilidad pasiva Berserk se activó automáticamente.", FontTypeNames.FONTTYPE_INFO)
        'TODO: Nightw - Change the FX to the corresponding glow effect.
        'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, FXIDs.FXMEDITARCHICO, INFINITE_LOOPS, True))
    Else
        Call WriteConsoleMsg(UserIndex, "La habilidad pasiva Berserk se desactivó.", FontTypeNames.FONTTYPE_INFO)
        'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, 0, INFINITE_LOOPS, True))
    End If
    
    Call WriteBerserkerEnabled(UserIndex, Enabled)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SendBerserkEffect de modPassiveSkills.bas")
End Sub


' NEW SYSTEM

Public Sub SetUserPassiveDefaults(ByVal UserIndex As Integer)
On Error GoTo ErrHandler:

    If ServerConfiguration.PassiveSkillsQty < 1 Then
        Erase UserList(UserIndex).Stats.UserPassives
        Exit Sub
    End If
        
    ReDim UserList(UserIndex).Stats.UserPassives(1 To ServerConfiguration.PassiveSkillsQty)
    
    Exit Sub
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function SetUserPassiveDefaults de modPassiveSkills.bas")
End Sub

Public Sub RecalculateUserPassives(ByVal UserIndex As Integer, ByVal NotifyNewUnlock As Boolean)
On Error GoTo ErrHandler:

    Dim I As Integer
    Dim J As Integer
    Dim PassiveWasEnabled As Boolean
    
    If ServerConfiguration.PassiveSkillsQty < 1 Then Exit Sub
    
    With UserList(UserIndex)
        
        For I = 1 To ServerConfiguration.PassiveSkillsQty
            PassiveWasEnabled = .Stats.UserPassives(I).Enabled
        
            .Stats.UserPassives(I).AllowedByClass = ClassCanUnlockPassive(.Clase, I)
            .Stats.UserPassives(I).Enabled = .Stats.ELV >= ServerConfiguration.PassiveSkills(I).UnlockLevel And .Stats.UserPassives(I).AllowedByClass
            .Stats.UserPassives(I).Name = ServerConfiguration.PassiveSkills(I).Name
            
            If PassiveWasEnabled = False And .Stats.UserPassives(I).Enabled = True And NotifyNewUnlock Then
                Call WriteConsoleMsg(UserIndex, "¡Desbloqueaste la habilidad pasiva " & .Stats.UserPassives(I).Name & "!", FontTypeNames.FONTTYPE_INFO)
            End If

        Next I
    End With
 
    Exit Sub
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function RecalculateUserPassives de modPassiveSkills.bas")
End Sub


Private Function ClassCanUnlockPassive(ByVal Clase As eClass, ByVal PassiveIndex As Integer) As Boolean
On Error GoTo ErrHandler:
    Dim I As Integer
    ClassCanUnlockPassive = False
    
    With ServerConfiguration.PassiveSkills(PassiveIndex)
    
        ' If the passive is disabled, then the user can not unlock it.
        If Not .Enabled Then Exit Function
        
        ' Loop through the list of allowed classes.
        For I = 1 To .AllowedClassesQty
            If .AllowedClasses(I) = Clase Then
                ClassCanUnlockPassive = True
                Exit Function
            End If
        Next I
    End With
        
    Exit Function
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function ClassCanUnlockPassive de modPassiveSkills.bas")
End Function


Public Sub PrintUserPassives(ByVal UserIndex As Integer)
    Dim I As Integer
    
    For I = 1 To ServerConfiguration.PassiveSkillsQty
        Debug.Print UserList(UserIndex).Stats.UserPassives(I).Name, UserList(UserIndex).Stats.UserPassives(I).Enabled
    Next I
    
End Sub
