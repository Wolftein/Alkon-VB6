Attribute VB_Name = "Trabajo"
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

Public Sub DoPermanecerOculto(ByVal UserIndex As Integer)
'********************************************************
'Autor: Nacho (Integer)
'Last Modif: 11/19/2009
'Chequea si ya debe mostrarse
'Pablo (ToxicWaste): Cambie los ordenes de prioridades porque sino no andaba.
'13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
'13/01/2010: ZaMa - Arreglo condicional para que el bandido camine oculto.
'17/03/2015: Luke - Hunter no longer needs Armor to hide permanently. Only needs 100 skill points.
'********************************************************
On Error GoTo ErrHandler

    With UserList(UserIndex)
        
        If .clase = eClass.Hunter Then Exit Sub
        If Not IsIntervalReached(.Counters.TiempoOculto) Then Exit Sub
    
        .flags.Oculto = 0

        If Not ThiefRestoreBoatAppearance(UserIndex) And .flags.invisible = 0 Then
            Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
            
            'Si está en el oscuro no lo hacemos visible
            If MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger <> eTrigger.zonaOscura Then
                Call UsUaRiOs.SetInvisible(UserIndex, .Char.CharIndex, False)
            End If
        End If
        
        If BerzerkConditionMet(UserIndex) Then
            Call ActivatePassive(UserIndex, ePassiveSpells.Berserk, True)
            Call SendBerserkEffect(UserIndex, ePassiveSpells.Berserk, True)
        End If
    End With
    
    Exit Sub

ErrHandler:
    Call LogError("Error en Sub DoPermanecerOculto")
End Sub

Public Sub DoOcultarse(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 13/01/2010 (ZaMa)
'Modifique la fórmula y ahora anda bien.
'13/01/2010: ZaMa - El pirata se transforma en galeon fantasmal cuando se oculta en agua.
'***************************************************

On Error GoTo ErrHandler

    Dim Suerte As Double
    Dim res As Integer
    Dim skill As Integer
    
    With UserList(UserIndex)
        Suerte = 5 + GetSkills(UserIndex, eSkill.Ocultarse) * Classes(.Clase).ClassMods.HidingChance

        If Suerte > RandomNumber(1, 100) Then
            .flags.Oculto = 1
            Suerte = (GetSkills(UserIndex, eSkill.Ocultarse) * ServerConfiguration.Intervals.IntervaloOculto / 125 + ServerConfiguration.Intervals.IntervaloOculto / 5) * Classes(.Clase).ClassMods.HidingDuration
            .Counters.TiempoOculto = SetIntervalEnd(CLng(Suerte))
            
            ' Disable berserk
            If HasPassiveActivated(UserIndex, ePassiveSpells.Berserk) Then
                Call ActivatePassive(UserIndex, ePassiveSpells.Berserk, False)
                Call SendBerserkEffect(UserIndex, ePassiveSpells.Berserk, False)
            End If
            
            
            ' No es pirata o es uno sin barca
            If .flags.Navegando = 0 Then
                Call UsUaRiOs.SetInvisible(UserIndex, .Char.CharIndex, True)
        
                Call WriteConsoleMsg(UserIndex, "¡Te has escondido entre las sombras!", FontTypeNames.FONTTYPE_INFO)
            ' Es un pirata navegando
            Else
                ' Le cambiamos el body a galeon fantasmal
                .Char.body = ConstantesGRH.FragataFantasmal
                ' Actualizamos clientes
                Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.heading, ConstantesGRH.NingunArma, _
                                    ConstantesGRH.NingunEscudo, ConstantesGRH.NingunCasco)
                
                Call WriteConsoleMsg(UserIndex, "¡Has adquirido una apariencia fantasmal!", FontTypeNames.FONTTYPE_INFO)
            End If
            
            Call SubirSkill(UserIndex, eSkill.Ocultarse, True)
        Else
            '[CDT 17-02-2004]
            If Not .flags.UltimoMensaje = 4 Then
                Call WriteConsoleMsg(UserIndex, "¡No has logrado esconderte!", FontTypeNames.FONTTYPE_INFO)
                .flags.UltimoMensaje = 4
            End If
            '[/CDT]
            
            Call SubirSkill(UserIndex, eSkill.Ocultarse, False)
        End If

    End With
    
    Exit Sub

ErrHandler:
    Call LogError("Error en Sub DoOcultarse")

End Sub

Public Sub DoNavega(ByVal UserIndex As Integer, ByVal Slot As Byte)
'***************************************************
'Author: Unknown
'Last Modification: 07/02/2015 (D'Artagnan)
'13/01/2010: ZaMa - El pirata pierde el ocultar si desequipa barca.
'16/09/2010: ZaMa - Ahora siempre se va el invi para los clientes al equipar la barca (Evita cortes de cabeza).
'10/12/2010: Pato - Limpio las variables del inventario que hacen referencia a la barca, sino el pirata que la última barca que equipo era el galeón no explotaba(Y capaz no la tenía equipada :P).
'07/02/2015: D'Artagnan - Restore secondary armour.
'***************************************************
On Error GoTo ErrHandler
  

    Dim ModNave As Single
        Dim Barco As ObjData
    
    With UserList(UserIndex)
        ModNave = ModNavegacion(.clase, UserIndex)
                Barco = ObjData(.Invent.Object(Slot).ObjIndex)
        
        If .Stats.ELV < 25 Then
            Call WriteConsoleMsg(UserIndex, "No tienes suficiente nivel para navegar.", FontTypeNames.FONTTYPE_INFO)
            If ModNave = 2 And Barco.MinSkill <> 35 Then
                Call WriteConsoleMsg(UserIndex, "Tu clase no puede usar esta embarcación.", FontTypeNames.FONTTYPE_INFO)
            End If
            Exit Sub
        End If
        
        ' No estaba navegando
        If .flags.Navegando = 0 Then
            .Invent.BarcoObjIndex = .Invent.Object(Slot).ObjIndex
            .Invent.BarcoSlot = Slot
            
            .Char.head = 0
            
            ' No esta muerto
            If .flags.Muerto = 0 Then
                
                ' Cheks if remains mimetized
                If .flags.Mimetizado <> 0 Then
                    If .flags.MimetizadoType = eMimeType.ieTerrain Then
                        Call EndMimic(UserIndex, False, True)
                    End If
                End If
                
                Call ToggleBoatBody(UserIndex)
                
                ' Pierde el ocultar
                If .flags.Oculto = 1 Then
                    .flags.Oculto = 0
                    Call UsUaRiOs.SetInvisible(UserIndex, .Char.CharIndex, False)
                    Call WriteConsoleMsg(UserIndex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
                    
                    ' If the users was hidden and we made it visible again, let's check the berserk passive
                    'If BerzerkConditionMet(UserIndex) Then
                    '    Call ActivatePassive(UserIndex, ePassiveSpells.Berserk, True)
                    '    Call SendBerserkEffect(UserIndex, ePassiveSpells.Berserk, True)
                    'End If
                End If
                
                ' Disable the Berzerk
                If HasPassiveAssigned(UserIndex, ePassiveSpells.Berserk) Then
                    If HasPassiveActivated(UserIndex, ePassiveSpells.Berserk) And Not .Masteries.Boosts.EnableBerserkWhileSailing Then
                        Call ActivatePassive(UserIndex, ePassiveSpells.Berserk, False)
                        Call SendBerserkEffect(UserIndex, ePassiveSpells.Berserk, False)
                    End If
                End If
               
                ' Siempre se ve la barca (Nunca esta invisible), pero solo para el cliente.
                If .flags.invisible = 1 Then
                    Call UsUaRiOs.SetInvisible(UserIndex, .Char.CharIndex, False)
                End If
                
            ' Esta muerto
            Else
                .Char.body = ConstantesGRH.FragataFantasmal
                .Char.ShieldAnim = ConstantesGRH.NingunEscudo
                .Char.WeaponAnim = ConstantesGRH.NingunArma
                .Char.CascoAnim = ConstantesGRH.NingunCasco
            End If
            
            ' Comienza a navegar
            .flags.Navegando = 1
        
        ' Estaba navegando
        Else
            .Invent.BarcoObjIndex = 0
            .Invent.BarcoSlot = 0
            
            ' Remueve mimetismo
            If .flags.Mimetizado = 1 Then
                .Counters.Mimetismo = 0
                Call EfectoMimetismo(UserIndex)
            End If
            
        
            ' No esta muerto
            If .flags.Muerto = 0 Then
                .Char.head = .OrigChar.head
                
                ' Cheks if remains mimetized
                If (.flags.Mimetizado <> 0) Then
                    If .flags.MimetizadoType = eMimeType.ieAquatic Then
                        Call EndMimic(UserIndex, False, True)
                    End If
                End If
                
                If .clase = eClass.Thief Then
                    If .flags.Oculto = 1 Then
                        ' Al desequipar barca, perdió el ocultar
                        .flags.Oculto = 0
                        Call WriteConsoleMsg(UserIndex, "¡Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
                
                If .Invent.ArmourEqpObjIndex > 0 Then
                    ' Restore secondary armour.
                    If .Invent.FactionArmourEqpObjIndex > 0 Then
                        .Char.body = GetBodyForUser(UserIndex, .Invent.FactionArmourEqpObjIndex)
                    Else
                        .Char.body = GetBodyForUser(UserIndex, .Invent.ArmourEqpObjIndex)
                    End If
                Else
                    Call DarCuerpoDesnudo(UserIndex)
                End If
                
                If .Invent.EscudoEqpObjIndex > 0 Then _
                    .Char.ShieldAnim = ObjData(.Invent.EscudoEqpObjIndex).ShieldAnim
                If .Invent.WeaponEqpObjIndex > 0 Then _
                    .Char.WeaponAnim = GetWeaponAnim(UserIndex, .Invent.WeaponEqpObjIndex)
                If .Invent.CascoEqpObjIndex > 0 Then _
                    .Char.CascoAnim = ObjData(.Invent.CascoEqpObjIndex).CascoAnim
                
                
                ' Al dejar de navegar, si estaba invisible actualizo los clientes
                If .flags.invisible = 1 Then
                    Call UsUaRiOs.SetInvisible(UserIndex, .Char.CharIndex, True)
                End If
                

            ' Esta muerto
            Else
                .Char.body = ConstantesGRH.CuerpoMuerto
                .Char.head = ConstantesGRH.CabezaMuerto
                .Char.ShieldAnim = ConstantesGRH.NingunEscudo
                .Char.WeaponAnim = ConstantesGRH.NingunArma
                .Char.CascoAnim = ConstantesGRH.NingunCasco
            End If
            
            ' Termina de navegar
            .flags.Navegando = 0
            
            ' Enable Berserk
            If HasPassiveAssigned(UserIndex, ePassiveSpells.Berserk) And Not HasPassiveActivated(UserIndex, ePassiveSpells.Berserk) Then
                If BerzerkConditionMet(UserIndex) Then
                    Call ActivatePassive(UserIndex, ePassiveSpells.Berserk)
                    Call SendBerserkEffect(UserIndex, ePassiveSpells.Berserk, True)
                End If
            End If
        End If
        
        ' Actualizo clientes
        Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
    End With

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub DoNavega de Trabajo.bas")
End Sub

Public Sub FundirMineral(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo ErrHandler

    With UserList(UserIndex)
        If .flags.TargetObjInvIndex > 0 Then
           If ObjData(.flags.TargetObjInvIndex).ObjType = eOBJType.otMinerales And _
                ObjData(.flags.TargetObjInvIndex).MinSkill <= GetSkills(UserIndex, eSkill.Mineria) / ModFundicion(.clase) Then
                Call DoLingotes(UserIndex)
           Else
                Call WriteConsoleMsg(UserIndex, "No tienes conocimientos de minería suficientes para trabajar este mineral.", FontTypeNames.FONTTYPE_INFO)
           End If
        
        End If
    End With

    Exit Sub

ErrHandler:
    Call LogError("Error en FundirMineral. Error " & Err.Number & " : " & Err.Description)

End Sub

Public Sub FundirArmas(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)
        If .flags.TargetObjInvIndex > 0 Then
            If ObjData(.flags.TargetObjInvIndex).ObjType = eOBJType.otWeapon Then
                If ObjData(.flags.TargetObjInvIndex).SkHerreria <= GetSkills(UserIndex, eSkill.Herreria) / ModHerreriA(.clase) Then
                    Call DoFundir(UserIndex)
                Else
                    Call WriteConsoleMsg(UserIndex, "No tienes los conocimientos suficientes en herrería para fundir este objeto.", FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                Call WriteConsoleMsg(UserIndex, "Solo puedes fundir armas.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
    End With
    
    Exit Sub
ErrHandler:
    Call LogError("Error en FundirArmas. Error " & Err.Number & " : " & Err.Description)
End Sub

Function TieneObjetos(ByVal ItemIndex As Integer, ByVal cant As Long, ByVal UserIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 10/07/2010
'10/07/2010: ZaMa - Ahora cant es long para evitar un overflow.
'***************************************************
On Error GoTo ErrHandler
  

    Dim I As Integer
    Dim Total As Long
    For I = 1 To UserList(UserIndex).CurrentInventorySlots
        If UserList(UserIndex).Invent.Object(I).ObjIndex = ItemIndex Then
            Total = Total + UserList(UserIndex).Invent.Object(I).Amount
        End If
    Next I
    
    If cant <= Total Then
        TieneObjetos = True
        Exit Function
    End If
        
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function TieneObjetos de Trabajo.bas")
End Function

Public Sub QuitarObjetos(ByVal ItemIndex As Integer, ByVal cant As Long, ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 05/08/09
'05/08/09: Pato - Cambie la funcion a procedimiento ya que se usa como procedimiento siempre, y fixie el bug 2788199
'***************************************************
On Error GoTo ErrHandler
  

    Dim I As Integer
    For I = 1 To UserList(UserIndex).CurrentInventorySlots
        With UserList(UserIndex).Invent.Object(I)
            If .ObjIndex = ItemIndex Then
                If .Amount <= cant And .Equipped = 1 Then Call Desequipar(UserIndex, I, True)
                
                .Amount = .Amount - cant
                If .Amount <= 0 Then
                    cant = Abs(.Amount)
                    UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
                    .Amount = 0
                    .ObjIndex = 0
                Else
                    cant = 0
                End If
                
                Call UpdateUserInv(False, UserIndex, I)
                
                If cant = 0 Then Exit Sub
            End If
        End With
    Next I

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub QuitarObjetos de Trabajo.bas")
End Sub

Private Function MineralesParaLingote(ByVal Lingote As iMinerales) As Integer
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  
    Select Case Lingote
        Case iMinerales.HierroCrudo
            MineralesParaLingote = 7
        Case iMinerales.PlataCruda
            MineralesParaLingote = 20
        Case iMinerales.OroCrudo
            MineralesParaLingote = 35
        Case Else
            MineralesParaLingote = 10000
    End Select
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function MineralesParaLingote de Trabajo.bas")
End Function


Public Sub DoLingotes(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 16/11/2009
'16/11/2009: ZaMa - Implementado nuevo sistema de construccion de items
'***************************************************

On Error GoTo ErrHandler
  
    Dim Slot As Integer
    Dim obji As Integer
    Dim CantidadItems As Integer
    Dim TieneMinerales As Boolean
    Dim OtroUserIndex As Integer
    
    With UserList(UserIndex)
        If isTradingWithUser(UserIndex) Then
            OtroUserIndex = getTradingUser(UserIndex)
                
            If (OtroUserIndex > 0) And (OtroUserIndex <= MaxUsers) Then
                Call WriteConsoleMsg(UserIndex, "¡¡Comercio cancelado, no puedes comerciar mientras trabajas!!", FontTypeNames.FONTTYPE_TALK)
                Call WriteConsoleMsg(OtroUserIndex, "¡¡Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
                
                Call LimpiarComercioSeguro(UserIndex)
            End If
        End If
        
        CantidadItems = MaximoInt(1, CInt((.Stats.ELV - 4) / 5))

        Slot = .flags.TargetObjInvSlot
        obji = .Invent.Object(Slot).ObjIndex
        
        Dim NumLingotes As Integer
        NumLingotes = MineralesParaLingote(obji)
        While (CantidadItems > 0) And (Not TieneMinerales)
            If .Invent.Object(Slot).Amount >= (NumLingotes * CantidadItems) Then
                TieneMinerales = True
            Else
                CantidadItems = CantidadItems - 1
            End If
        Wend
        
        If (Not TieneMinerales) Or (ObjData(obji).ObjType <> eOBJType.otMinerales) Then
            Call WriteConsoleMsg(UserIndex, "No tienes suficientes minerales para hacer un lingote.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Dim MiObj As Obj
        
        MiObj.Amount = CantidadItems
        MiObj.ObjIndex = ObjData(.flags.TargetObjInvIndex).LingoteIndex
              
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call WriteConsoleMsg(UserIndex, "No tienes suficiente espacio en el inventario para hacer un lingote.", FontTypeNames.FONTTYPE_INFO)
            Call WriteStopWorking(UserIndex)
            Exit Sub
        End If
        
        Call QuitarUserInvItem(UserIndex, Slot, NumLingotes * CantidadItems)
        
        Call WriteConsoleMsg(UserIndex, "¡Has obtenido " & CantidadItems & " lingote" & _
                            IIf(CantidadItems = 1, "", "s") & "!", FontTypeNames.FONTTYPE_INFO)
    
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub DoLingotes de Trabajo.bas")
End Sub

Public Sub DoFundir(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 03/06/2010
'03/06/2010 - Pato: Si es el último ítem a fundir y está equipado lo desequipamos.
'11/03/2010 - ZaMa: Reemplazo división por producto para uan mejor performanse.
'***************************************************
On Error GoTo ErrHandler
  
Dim I As Integer
Dim Num As Integer
Dim Slot As Byte
Dim Lingotes(2) As Integer
Dim OtroUserIndex As Integer
Dim ItemIndex As Integer
 
    With UserList(UserIndex)
        If isTradingWithUser(UserIndex) Then
            OtroUserIndex = getTradingUser(UserIndex)
                
            If (OtroUserIndex > 0) And (OtroUserIndex <= MaxUsers) Then
                Call WriteConsoleMsg(UserIndex, "¡¡Comercio cancelado, no puedes comerciar mientras trabajas!!", FontTypeNames.FONTTYPE_TALK)
                Call WriteConsoleMsg(OtroUserIndex, "¡¡Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
                
                Call LimpiarComercioSeguro(UserIndex)
            End If
        End If
        
        Slot = .flags.TargetObjInvSlot
        
        If Slot > 1 And Slot < .CurrentInventorySlots Then _
            ItemIndex = .Invent.Object(Slot).ObjIndex
        
        Call QuitarUserInvItem(UserIndex, Slot, 1)
        
        Num = RandomNumber(10, 25)
        
        Lingotes(0) = Int((ObjData(.flags.TargetObjInvIndex).LingH * Num) * 0.01)
        Lingotes(1) = Int((ObjData(.flags.TargetObjInvIndex).LingP * Num) * 0.01)
        Lingotes(2) = Int((ObjData(.flags.TargetObjInvIndex).LingO * Num) * 0.01)
    
        Dim MiObj(2) As Obj
        
        For I = 0 To 2
            MiObj(I).Amount = Lingotes(I)
            MiObj(I).ObjIndex = ConstantesItems.LingoteHierro + I 'Una gran negrada pero práctica
            
            If MiObj(I).Amount > 0 Then
                If Not MeterItemEnInventario(UserIndex, MiObj(I)) Then
                    Call TirarItemAlPiso(.Pos, MiObj(I))
                End If
            End If
        Next I
        
        Call WriteConsoleMsg(UserIndex, "¡Has obtenido el " & Num & "% de los lingotes utilizados para la construcción del objeto!", FontTypeNames.FONTTYPE_INFO)
        
        If ItemIndex > 0 Then
            If ObjData(ItemIndex).Log = 1 Then _
                Call LogDesarrollo(.Name & " ha fundido el ítem " & ObjData(ItemIndex).Name)
        End If
        
    End With

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub DoFundir de Trabajo.bas")
End Sub

Function ModNavegacion(ByVal clase As eClass, ByVal UserIndex As Integer) As Single
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 27/11/2009
'27/11/2009: ZaMa - A worker can navigate before only if it's an expert fisher
'12/04/2010: ZaMa - Arreglo modificador de pescador, para que navegue con 60 skills.
'17/03/2015: Luke - Los skills ya no se checkean. Solo Worker y Thief/Pirate usan galeon/galera.
'***************************************************
On Error GoTo ErrHandler
  
Select Case clase
    Case eClass.Thief
        ModNavegacion = 1
    Case eClass.Worker
        ModNavegacion = 1
    Case Else
        ModNavegacion = 2
End Select

  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function ModNavegacion de Trabajo.bas")
End Function


Function ModFundicion(ByVal clase As eClass) As Single
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

Select Case clase
    Case eClass.Worker
        ModFundicion = 1
    Case Else
        ModFundicion = 3
End Select

  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function ModFundicion de Trabajo.bas")
End Function

Function ModCarpinteria(ByVal clase As eClass) As Integer
On Error GoTo ErrHandler
' TODO: Move this to the Classes.dat file

Select Case clase
    Case eClass.Worker
        ModCarpinteria = 1
    Case Else
        ModCarpinteria = 2
End Select

  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function ModCarpinteria de Trabajo.bas")
End Function

Function ModHerreriA(ByVal clase As eClass) As Single
On Error GoTo ErrHandler
' TODO: Move this to the Classes.dat file
  
Select Case clase
    Case eClass.Worker
        ModHerreriA = 1
    Case Else
        ModHerreriA = 2
End Select

  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function ModHerreriA de Trabajo.bas")
End Function

Function FreeInvokedPetIndex(ByVal UserIndex As Integer) As Integer
'***************************************************
'Author: Unknown
'Last Modification: 02/03/09
'02/03/09: ZaMa - Busca un indice libre de mascotas, revisando los types y no los indices de los npcs
'***************************************************
On Error GoTo ErrHandler
  
    Dim J As Integer
    With UserList(UserIndex)
        For J = 1 To Classes(.clase).ClassMods.MaxInvokedPets
            If .InvokedPets(J).NpcIndex = 0 Then
                FreeInvokedPetIndex = J
                Exit Function
            End If
        Next J
    End With
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function FreeMascotaIndex de Trabajo.bas")
End Function

Function FreeMascotaIndex(ByVal UserIndex As Integer) As Integer
'***************************************************
'Author: Unknown
'Last Modification: 02/03/09
'02/03/09: ZaMa - Busca un indice libre de mascotas, revisando los types y no los indices de los npcs
'***************************************************
On Error GoTo ErrHandler
  
    Dim J As Integer
    With UserList(UserIndex)
        For J = 1 To Classes(.clase).ClassMods.MaxTammedPets
            If .TammedPets(J).NpcNumber = 0 Then
                FreeMascotaIndex = J
                Exit Function
            End If
        Next J
    End With
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function FreeMascotaIndex de Trabajo.bas")
End Function

Sub DoTameNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
On Error GoTo ErrHandler

    Dim puntosDomar As Integer
    Dim puntosRequeridos As Integer
    Dim CanStay As Boolean
    Dim petType As Integer
    Dim NroPets As Integer
    Dim TammingSuccessful As Boolean
    Dim OldRemainingLife As Integer
    
    If Npclist(NpcIndex).MaestroUser = UserIndex Then
        Call WriteConsoleMsg(UserIndex, "Ya domaste a esa criatura.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    With UserList(UserIndex)
    
        If .TammedPetsCount >= Classes(.clase).ClassMods.MaxTammedPets And _
               (.TammedPetsCount + .InvokedPetsCount) >= Classes(.clase).ClassMods.MaxActivePets Then
            Call WriteConsoleMsg(UserIndex, "No puedes controlar más criaturas.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        If Npclist(NpcIndex).flags.ItemToTame <> 0 And Npclist(NpcIndex).flags.ItemToTame <> .Invent.AnilloEqpObjIndex Then
            Call WriteConsoleMsg(UserIndex, "No posees el objeto necesario para domar esta criatura.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    
        If Npclist(NpcIndex).MaestroNpc > 0 Or Npclist(NpcIndex).MaestroUser > 0 Then
            Call WriteConsoleMsg(UserIndex, "La criatura ya tiene amo.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Not PuedeDomarMascota(UserIndex, NpcIndex) Then
            Call WriteConsoleMsg(UserIndex, "No puedes domar más criaturas del mismo tipo.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
                    
        'Nueva fórmula para doma que usa un modificador de clase: [Skills en doma*Carisma*Modificador]: Puntos de doma
        puntosDomar = (CInt(GetSkills(UserIndex, eSkill.Domar)) * CInt(.Stats.UserAtributos(eAtributos.Carisma)) * _
                      Classes(.clase).ClassMods.Taming) + .Masteries.Boosts.AddTamingPoints
                                    
        puntosRequeridos = Npclist(NpcIndex).flags.Domable
        
        If puntosRequeridos > puntosDomar Then
            Call WriteConsoleMsg(UserIndex, "Todavía no puedes domar esta criatura.", FontTypeNames.FONTTYPE_INFO)
            Call SubirSkill(UserIndex, eSkill.Domar, False)
            Exit Sub
        End If
        
        If RandomNumber(1, 5) = 1 Then
            
            Dim Index As Integer
            
            TammingSuccessful = True
            
            'El user que lo domo tiene otras mascotas?
            If .TammedPetsCount > 0 Then
                For Index = 1 To Classes(.clase).ClassMods.MaxTammedPets
                    If .TammedPets(Index).NpcIndex > 0 Then
                                           
                        OldRemainingLife = .TammedPets(Index).RemainingLife
                        
                        Call QuitarNPC(.TammedPets(Index).NpcIndex)
                        
                        .TammedPets(Index).RemainingLife = OldRemainingLife
                    End If
                Next Index
            End If
            
            If .flags.AtacadoPorNpc = NpcIndex Then .flags.AtacadoPorNpc = 0
            .TammedPetsCount = .TammedPetsCount + 1
            
            Index = FreeMascotaIndex(UserIndex)
            .TammedPets(Index).NpcIndex = NpcIndex
            .TammedPets(Index).NpcNumber = Npclist(NpcIndex).Numero
            .TammedPets(Index).RemainingLife = Npclist(NpcIndex).Stats.MinHp
            
            Npclist(NpcIndex).MaestroUser = UserIndex
            Npclist(NpcIndex).MenuIndex = eMenues.ieMascota
      
            Call FollowAmo(NpcIndex)
            Call ReSpawnNpc(Npclist(NpcIndex))
            
            ' Bosses can also be tammed, so we need to clear the flags that make the NPC
            ' behave like a boss
            If Npclist(NpcIndex).flags.Boss > 0 Then
                Call modBosses.RestartBossSpawn(Npclist(NpcIndex).flags.Boss)
                Npclist(NpcIndex).flags.Boss = 0
            End If
            
            Call WriteConsoleMsg(UserIndex, "La criatura te ha aceptado como su amo.", FontTypeNames.FONTTYPE_INFO)
            
            ' Set the last tamed npc
            .flags.LastTamedPet = Npclist(NpcIndex).Numero
            
            ' Send the list of pets to the client.
            Call WriteSendPetList(UserIndex)
            
            ' Es zona segura?
            CanStay = (MapInfo(.Pos.Map).Pk = True)
            
            If Not CanStay Then
                petType = Npclist(NpcIndex).Numero
                NroPets = .TammedPetsCount
                
                Call QuitarNPC(NpcIndex)
                
                .TammedPets(Index).NpcNumber = petType
                .TammedPetsCount = NroPets
                
                Call WriteConsoleMsg(UserIndex, "No se permiten mascotas en zona segura. Éstas te esperarán afuera.", FontTypeNames.FONTTYPE_INFO)
            End If
        Else
            Call WriteConsoleMsg(UserIndex, "No has logrado domar la criatura.", FontTypeNames.FONTTYPE_INFO)
        End If
        
        Call SubirSkill(UserIndex, eSkill.Domar, TammingSuccessful)
    
    End With
    
    Exit Sub

ErrHandler:
    Call LogError("Error en DoTameNpc. Error " & Err.Number & " : " & Err.Description)

End Sub

''
' Checks if the user can tames a pet.
'
' @param integer userIndex The user id from who wants tame the pet.
' @param integer NPCindex The index of the npc to tome.
' @return boolean True if can, false if not.
Private Function PuedeDomarMascota(ByVal UserIndex As Integer, ByVal NpcIndex As Integer) As Boolean
'***************************************************
'Author: ZaMa
'This function checks how many NPCs of the same type have
'been tamed by the user.
'Returns True if that amount is less than two.
'***************************************************
On Error GoTo ErrHandler
  
    Dim I As Long
    Dim numMascotas As Long
    
    With UserList(UserIndex)
        For I = 1 To Classes(.clase).ClassMods.MaxTammedPets
            If .TammedPets(I).NpcNumber = Npclist(NpcIndex).Numero Then
                numMascotas = numMascotas + 1
            End If
        Next I
    End With
    
    If numMascotas < 1 Then PuedeDomarMascota = True
    
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function PuedeDomarMascota de Trabajo.bas")
End Function

Sub DoAdminInvisible(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 12/01/2010 (ZaMa)
'Makes an admin invisible o visible.
'13/07/2009: ZaMa - Now invisible admins' chars are erased from all clients, except from themselves.
'12/01/2010: ZaMa - Los druidas pierden la inmunidad de ser atacados cuando pierden el efecto del mimetismo.
'***************************************************
On Error GoTo ErrHandler
  
    
    With UserList(UserIndex)
        If .flags.AdminInvisible = 0 Then
            ' Sacamos el mimetizmo
            If .flags.Mimetizado <> 0 Then
                Call EndMimic(UserIndex, False, False)
            End If
            
            .flags.AdminInvisible = 1
            .flags.invisible = 1
            .flags.Oculto = 1
            
            ' Solo el admin sabe que se hace invi
            Call SendData(ToUser, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, True))
            'Le mandamos el mensaje para que borre el personaje a los clientes que estén cerca
            Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageCharacterRemove(.Char.CharIndex))
        Else
            .flags.AdminInvisible = 0
            .flags.invisible = 0
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0

            ' Solo el admin sabe que se hace visible
            Call SendData(ToUser, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
             
            'Le mandamos el mensaje para crear el personaje a los clientes que estén cerca
            Call MakeUserChar(True, UserIndex, UserIndex)
        End If
    End With
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub DoAdminInvisible de Trabajo.bas")
End Sub

Public Sub DoPescar(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 24/01/2015
'16/11/2009: ZaMa - Implementado nuevo sistema de extraccion.
'11/05/2010: ZaMa - Arreglo formula de maximo de items contruibles/extraibles.
'05/13/2010: Pato - Refix a la formula de maximo de items construibles/extraibles.
'22/05/2010: ZaMa - Los caos ya no suben plebe al trabajar.
'28/05/2010: ZaMa - Los pks no suben plebe al trabajar.
'24/01/2015: D'Artagnan - Fishing rod and code optimization.
'09/07/2016: Anagrama - Ahora la plebe no sube solo si es criminal y tiene seguro activado.
'***************************************************
On Error GoTo ErrHandler

Dim Suerte As Integer
Dim res As Integer
Dim CantidadItems As Integer
Dim skill As Integer
Dim nAmount As Integer

With UserList(UserIndex)
    If .clase = eClass.Worker Then
        Call QuitarSta(UserIndex, ConstantesTrabajo.EsfuerzoPescarPescador)
    Else
        Call QuitarSta(UserIndex, ConstantesTrabajo.EsfuerzoPescarGeneral)
    End If
    
    skill = GetSkills(UserIndex, eSkill.Pesca)
    Suerte = Int(-0.00125 * skill * skill - 0.3 * skill + 49)
    
    res = RandomNumber(1, Suerte)
    
    If res <= 6 Then
        If .clase = eClass.Worker Then
            CantidadItems = MaxItemsExtraibles(.Stats.ELV)
            nAmount = RandomNumber(1, CantidadItems)
        Else
            nAmount = 1
        End If
        
        Call AddInventoryItem(UserIndex, Pescado, nAmount)
        Call WriteConsoleMsg(UserIndex, "¡Has pescado un lindo pez!", FontTypeNames.FONTTYPE_INFO, eMessageType.Trabajo)
        
        ' Fishing rod.
        If RandomNumber(1, 10000) = 5000 Then
            Call AddInventoryItem(UserIndex, RED_PESCA, 1)
            Call WriteConsoleMsg(UserIndex, "¡Has obtenido una red de pesca!", FontTypeNames.FONTTYPE_INFO, _
                                 eMessageType.Trabajo)
        End If
        
        Call SubirSkill(UserIndex, eSkill.Pesca, True)
    Else
        '[CDT 17-02-2004]
        If Not .flags.UltimoMensaje = 6 Then
          Call WriteConsoleMsg(UserIndex, "¡No has pescado nada!", FontTypeNames.FONTTYPE_INFO, eMessageType.Trabajo)
          .flags.UltimoMensaje = 6
        End If
        '[/CDT]
        
        Call SubirSkill(UserIndex, eSkill.Pesca, False)
    End If
    
End With

Exit Sub

ErrHandler:
    Call LogError("Error en DoPescar. Error " & Err.Number & " : " & Err.Description)
End Sub

Public Sub DoPescarRed(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 24/01/2015
'12/10/2014: D'Artagnan - Surprise items.
'24/01/2015: D'Artagnan - New logic and code optimization.
'09/07/2016: Anagrama - Ahora la plebe no sube solo si es criminal y tiene seguro activado.
'***************************************************
On Error GoTo ErrHandler

    Dim iSkill As Integer
    Dim Suerte As Integer
    Dim res As Integer
    Dim EsPescador As Boolean
    Dim CantidadItems As Integer
    Dim nAmount As Integer
    Dim nItemID As Integer
    
    With UserList(UserIndex)
    
        If .clase = eClass.Worker Then
            Call QuitarSta(UserIndex, ConstantesTrabajo.EsfuerzoPescarPescador)
            EsPescador = True
        Else
            Call QuitarSta(UserIndex, ConstantesTrabajo.EsfuerzoPescarGeneral)
            EsPescador = False
        End If
        
        iSkill = GetSkills(UserIndex, eSkill.Pesca)
        
        ' m = (60-11)/(1-10)
        ' y = mx - m*10 + 11
        
        Suerte = Int(-0.00125 * iSkill * iSkill - 0.3 * iSkill + 49)

        If Suerte > 0 Then
            res = RandomNumber(1, Suerte)
            
            If res <= 6 Then
                If EsPescador Then
                    CantidadItems = MaxItemsExtraibles(.Stats.ELV)
                    nAmount = RandomNumber(1, CantidadItems)
                Else
                    nAmount = 1
                End If
                
                If RandomNumber(0, 100) < 50 Then
                    nItemID = 544  ' "Pez dorado"
                ElseIf RandomNumber(0, 100) < 25 Then
                    nItemID = 546  ' "Merluza"
                ElseIf RandomNumber(0, 100) < 12 Then
                    nItemID = 545  ' "Pez espada"
                ElseIf RandomNumber(1, 10000) = 5000 Then
                    nItemID = 775  ' "Hipocampo"
                Else
                    nItemID = Pescado
                End If
                
                Call AddInventoryItem(UserIndex, nItemID, nAmount)
                Call WriteConsoleMsg(UserIndex, "¡Has pescado algunos peces!", FontTypeNames.FONTTYPE_INFO)
                
                ' Surprise items.
                If .flags.Navegando = 1 Then
                    If RandomNumber(0, 20000) < 5 Then
                        If .Pos.Map = 201 Or .Pos.Map = 283 Or .Pos.Map = 284 Or _
                           .Pos.Map = 285 Or .Pos.Map = 132 Then
                            ' Pirate chest.
                            nItemID = 1127
                        Else
                            ' Simple chest.
                            nItemID = 1126
                        End If
                
                        Call AddInventoryItem(UserIndex, nItemID, 1)
                        Call WriteConsoleMsg(UserIndex, _
                            "¡Has obtenido un cofre" & IIf(nItemID = 1127, " pirata", "") & "!", _
                            FontTypeNames.FONTTYPE_INFOBOLD)
                    End If
                End If
                
                Call SubirSkill(UserIndex, eSkill.Pesca, True)
            Else
                If Not .flags.UltimoMensaje = 6 Then
                  Call WriteConsoleMsg(UserIndex, "¡No has pescado nada!", FontTypeNames.FONTTYPE_INFO)
                  .flags.UltimoMensaje = 6
                End If
                
                Call SubirSkill(UserIndex, eSkill.Pesca, False)
            End If
        End If
        
    End With
    
    Exit Sub

ErrHandler:
    Call LogError("Error en DoPescarRed")
End Sub

''
' Try to steal an item / gold to another character
'
' @param LadrOnIndex Specifies reference to user that stoles
' @param VictimaIndex Specifies reference to user that is being stolen

Public Sub DoRobar(ByVal LadrOnIndex As Integer, ByVal VictimaIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 05/04/2010
'Last Modification By: ZaMa
'24/07/08: Marco - Now it calls to WriteUpdateGold(VictimaIndex and LadrOnIndex) when the thief stoles gold. (MarKoxX)
'27/11/2009: ZaMa - Optimizacion de codigo.
'18/12/2009: ZaMa - Los ladrones ciudas pueden robar a pks.
'01/04/2010: ZaMa - Los ladrones pasan a robar oro acorde a su nivel.
'05/04/2010: ZaMa - Los armadas no pueden robarle a ciudadanos jamas.
'23/04/2010: ZaMa - No se puede robar mas sin energia.
'23/04/2010: ZaMa - El alcance de robo pasa a ser de 1 tile.
'08/11/2015: D'Artagnan - On failure, earn skill experience as it would on success.
'*************************************************

On Error GoTo ErrHandler

    Dim OtroUserIndex As Integer

    If Not MapInfo(UserList(VictimaIndex).Pos.Map).Pk Then Exit Sub
    
    If UserList(VictimaIndex).flags.HelpMode Then
        Call WriteConsoleMsg(LadrOnIndex, "¡¡¡No puedes robar a usuarios en consulta!!!", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    With UserList(LadrOnIndex)
    
        If Not CanAttackOrStealByAlignment(LadrOnIndex, VictimaIndex) Then
            Call WriteConsoleMsg(LadrOnIndex, "No puedes robarle a alguien de esa alineación.", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub
        End If
        
        If TriggerZonaPelea(LadrOnIndex, VictimaIndex) <> TRIGGER6_AUSENTE Then Exit Sub
    
        ' Tiene energia?
        If .Stats.MinSta < 15 Then
            If .Genero = eGenero.Hombre Then
                Call WriteConsoleMsg(LadrOnIndex, "Estás muy cansado para robar.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(LadrOnIndex, "Estás muy cansada para robar.", FontTypeNames.FONTTYPE_INFO)
            End If
            
            Exit Sub
        End If
        
        ' Quito energia
        Call QuitarSta(LadrOnIndex, 15)
        
        Dim GuantesHurto As Boolean
    
        If .Invent.AnilloEqpObjIndex = ConstantesItems.GuanteHurto Then GuantesHurto = True
        
        If UserList(VictimaIndex).flags.Privilegios And PlayerType.User Then
            
            Dim Suerte As Integer
            
            Suerte = 10 + GetSkills(LadrOnIndex, eSkill.Robar) / 2 * Classes(.clase).ClassMods.StealingChance
                
            If Suerte > RandomNumber(1, 100) Then 'Exito robo
                If isTradingWithUser(VictimaIndex) Then
                    OtroUserIndex = getTradingUser(VictimaIndex)
                    
                    If (OtroUserIndex > 0) And (OtroUserIndex <= MaxUsers) Then
                        Call WriteConsoleMsg(VictimaIndex, "¡¡Comercio cancelado, te están robando!!", FontTypeNames.FONTTYPE_TALK)
                        Call WriteConsoleMsg(OtroUserIndex, "¡¡Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
                        
                        Call LimpiarComercioSeguro(VictimaIndex)
                    End If
                End If
               
                If (RandomNumber(1, 50) < 0) And (.clase = eClass.Thief) Then
                    If TieneObjetosRobables(VictimaIndex) Then
                        Call RobarObjeto(LadrOnIndex, VictimaIndex)
                    Else
                        Call WriteConsoleMsg(LadrOnIndex, UserList(VictimaIndex).Name & " no tiene objetos.", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else 'Roba oro
                    If UserList(VictimaIndex).Stats.GLD > 0 Then
                        Dim N As Long

                        If GuantesHurto Then
                            N = RandomNumber(.Stats.ELV * 50, .Stats.ELV * 100) * Classes(.clase).ClassMods.StealingAmount
                        Else
                            N = RandomNumber(.Stats.ELV * 25, .Stats.ELV * 50) * Classes(.clase).ClassMods.StealingAmount
                        End If

                        If N > UserList(VictimaIndex).Stats.GLD Then N = UserList(VictimaIndex).Stats.GLD
                        UserList(VictimaIndex).Stats.GLD = UserList(VictimaIndex).Stats.GLD - N
                        
                        .Stats.GLD = .Stats.GLD + N
                        If .Stats.GLD > ConstantesBalance.MaxOro Then _
                            .Stats.GLD = ConstantesBalance.MaxOro
                        
                        Call WriteConsoleMsg(LadrOnIndex, "Le has robado " & N & " monedas de oro a " & UserList(VictimaIndex).Name, FontTypeNames.FONTTYPE_INFO)
                        Call WriteUpdateGold(LadrOnIndex) 'Le actualizamos la billetera al ladron
                        
                        Call WriteUpdateGold(VictimaIndex) 'Le actualizamos la billetera a la victima
                    Else
                        Call WriteConsoleMsg(LadrOnIndex, UserList(VictimaIndex).Name & " no tiene oro.", FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
            Else
                Call WriteConsoleMsg(LadrOnIndex, "¡No has logrado robar nada!", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(VictimaIndex, "¡" & .Name & " ha intentado robarte!", FontTypeNames.FONTTYPE_INFO)
            End If
            
            Call SubirSkill(LadrOnIndex, eSkill.Robar, True)

        End If
    End With

Exit Sub

ErrHandler:
    Call LogError("Error en DoRobar. Error " & Err.Number & " : " & Err.Description)

End Sub

''
' Check if one item is stealable
'
' @param VictimaIndex Specifies reference to victim
' @param Slot Specifies reference to victim's inventory slot
' @return If the item is stealable
Public Function ObjEsRobable(ByVal VictimaIndex As Integer, ByVal Slot As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
' Agregué los barcos
' Esta funcion determina qué objetos son robables.
' 22/05/2010: Los items newbies ya no son robables.
' 28/02/2015: D'Artagnan - Factional items can be stolen.
'***************************************************
On Error GoTo ErrHandler
  

    Dim OI As Integer
    
    OI = UserList(VictimaIndex).Invent.Object(Slot).ObjIndex
    
    With ObjData(OI)
        ObjEsRobable = _
            .ObjType <> eOBJType.otLlaves And _
            UserList(VictimaIndex).Invent.Object(Slot).Equipped = 0 And _
            .ObjType <> eOBJType.otBarcos And _
            Not ItemNewbie(OI) And _
            .Intransferible = 0 And _
            .NoRobable = 0 And _
            Not IsSecondaryArmour(OI)
    End With
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function ObjEsRobable de Trabajo.bas")
End Function

Public Function ThiefRestoreBoatAppearance(ByVal UserIndex As Integer) As Boolean
'***************************************************
'Author: D'Artagnan
'Last Modification: 13/01/2010
'Switch from ghostly boat to normal appearance.
'Retrun True if succeed, False otherwise.
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)
        ' Must be Thief and sailing.
        If .flags.Navegando = 1 And .clase = eClass.Thief Then
            ' Switch to normal boat body.
            Call ToggleBoatBody(UserIndex)
            ' Update client.
            Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.heading, ConstantesGRH.NingunArma, _
                                ConstantesGRH.NingunEscudo, ConstantesGRH.NingunCasco)
            Call WriteConsoleMsg(UserIndex, "¡Has recuperado tu apariencia normal!", _
                                 FontTypeNames.FONTTYPE_INFO)
            ThiefRestoreBoatAppearance = True
        End If
    End With
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function ThiefRestoreBoatAppearance de Trabajo.bas")
End Function

''
' Try to steal an item to another character
'
' @param LadrOnIndex Specifies reference to user that stoles
' @param VictimaIndex Specifies reference to user that is being stolen
Public Sub RobarObjeto(ByVal LadrOnIndex As Integer, ByVal VictimaIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 02/04/2010
'02/04/2010: ZaMa - Modifico la cantidad de items robables por el ladron.
'***************************************************
On Error GoTo ErrHandler
  

Dim Flag As Boolean
Dim I As Integer

Flag = False

With UserList(VictimaIndex)
    If RandomNumber(1, 12) < 6 Then 'Comenzamos por el principio o el final?
        I = 1
        Do While Not Flag And I <= .CurrentInventorySlots
            'Hay objeto en este slot?
            If .Invent.Object(I).ObjIndex > 0 Then
               If ObjEsRobable(VictimaIndex, I) Then
                     If RandomNumber(1, 10) < 4 Then Flag = True
               End If
            End If
            If Not Flag Then I = I + 1
        Loop
    Else
        I = .CurrentInventorySlots
        Do While Not Flag And I > 0
          'Hay objeto en este slot?
          If .Invent.Object(I).ObjIndex > 0 Then
             If ObjEsRobable(VictimaIndex, I) Then
                   If RandomNumber(1, 10) < 4 Then Flag = True
             End If
          End If
          If Not Flag Then I = I - 1
        Loop
    End If
    
    If Flag Then
        Dim MiObj As Obj
        Dim Num As Integer
        Dim ObjAmount As Integer
        
        ObjAmount = .Invent.Object(I).Amount
        
        'Cantidad al azar entre el 5% y el 10% del total, con minimo 1.
        Num = MaximoInt(1, RandomNumber(ObjAmount * 0.05, ObjAmount * 0.1))
                                    
        MiObj.Amount = Num
        MiObj.ObjIndex = .Invent.Object(I).ObjIndex
        
        Call QuitarUserInvItem(VictimaIndex, I, Num)
                    
        If Not MeterItemEnInventario(LadrOnIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(LadrOnIndex).Pos, MiObj)
        End If
        
        If UserList(LadrOnIndex).clase = eClass.Thief Then
            Call WriteConsoleMsg(LadrOnIndex, "Has robado " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name, FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(LadrOnIndex, "Has hurtado " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name, FontTypeNames.FONTTYPE_INFO)
        End If
    Else
        Call WriteConsoleMsg(LadrOnIndex, "No has logrado robar ningún objeto.", FontTypeNames.FONTTYPE_INFO)
    End If

    'If exiting, cancel de quien es robado
    Call CancelExit(VictimaIndex)
End With

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub RobarObjeto de Trabajo.bas")
End Sub


Public Sub DoAcuchillar(ByVal UserIndex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal Damage As Integer)
'***************************************************
'Autor: ZaMa
'Last Modification: 12/01/2010
'***************************************************
On Error GoTo ErrHandler
  

    If RandomNumber(1, 100) <= ConstantesCombate.ProbAcuchillar Then
        Damage = Int(Damage * ConstantesCombate.DañoAcuchillar)
        
        If VictimUserIndex <> 0 Then
        
            With UserList(VictimUserIndex)
                .Stats.MinHp = .Stats.MinHp - Damage
                Call WriteConsoleMsg(UserIndex, "Has acuchillado a " & .Name & " por " & Damage, FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                Call WriteConsoleMsg(VictimUserIndex, UserList(UserIndex).Name & " te ha acuchillado por " & Damage, FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
            End With
            
        Else
        
            Npclist(VictimNpcIndex).Stats.MinHp = Npclist(VictimNpcIndex).Stats.MinHp - Damage
            Call WriteConsoleMsg(UserIndex, "Has acuchillado a la criatura por " & Damage, FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
            Call CalcularDarExp(UserIndex, VictimNpcIndex, Damage)
        
        End If
    End If
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub DoAcuchillar de Trabajo.bas")
End Sub

Public Sub DoGolpeCritico(ByVal UserIndex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal Damage As Long)
'***************************************************
'Autor: Pablo (ToxicWaste)
'Last Modification: 28/01/2007
'01/06/2010: ZaMa - Valido si tiene arma equipada antes de preguntar si es vikinga.
'***************************************************
On Error GoTo ErrHandler
  
    Dim Suerte As Integer
    Dim skill As Integer
    Dim WeaponIndex As Integer
    
    With UserList(UserIndex)
        ' Es bandido?
        If .clase <> eClass.Bandit Then Exit Sub
        
        WeaponIndex = .Invent.WeaponEqpObjIndex
    End With
    
    ' Only allowed to do critical hits with a weapon
    If WeaponIndex <= 0 Then Exit Sub
      
    ' Si no es 2 filos, mazo de juicio o espada de plata no hay critico
    If ObjData(WeaponIndex).Critical Then
    
        skill = GetSkills(UserIndex, eSkill.Wrestling)
        Suerte = Int((((0.00000003 * skill + 0.000006) * skill + 0.000107) * skill + 0.0893) * 100)
    
        If RandomNumber(1, 100) <= Suerte Then
    
            Damage = Int(Damage * 0.75)
        
            If VictimUserIndex <> 0 Then
            
                With UserList(VictimUserIndex)
                    .Stats.MinHp = .Stats.MinHp - Damage
                    Call WriteConsoleMsg(UserIndex, "Has golpeado críticamente a " & .Name & " por " & Damage & ".", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                    Call WriteConsoleMsg(VictimUserIndex, UserList(UserIndex).Name & " te ha golpeado críticamente por " & Damage & ".", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                End With
            
            Else
        
                Npclist(VictimNpcIndex).Stats.MinHp = Npclist(VictimNpcIndex).Stats.MinHp - Damage
                Call WriteConsoleMsg(UserIndex, "Has golpeado críticamente a la criatura por " & Damage & ".", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                Call CalcularDarExp(UserIndex, VictimNpcIndex, Damage)
            
            End If
        
        End If
    End If

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub DoGolpeCritico de Trabajo.bas")
End Sub

Public Sub QuitarSta(ByVal UserIndex As Integer, ByVal Cantidad As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo ErrHandler

    UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - Cantidad
    If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
    Call WriteUpdateSta(UserIndex)
    
Exit Sub

ErrHandler:
    Call LogError("Error en QuitarSta. Error " & Err.Number & " : " & Err.Description)
    
End Sub

Public Sub DoExtractResource(ByVal UserIndex As Integer, ByRef Pos As WorldPos, ByRef ProfessionIndex As Byte, ByRef ToolIndex As Integer)
On Error GoTo ErrHandler

Dim ModClass, Probability As Double
Dim ObjIndex, res, skill, MinQty, MaxQty, SkillNumber As Integer
Dim CanExtractFromResource As Boolean
Dim HasExtracted As Boolean

HasExtracted = False

With UserList(UserIndex)
    If .clase = eClass.Worker Then
        Call QuitarSta(UserIndex, Professions(ProfessionIndex).RequiredStaminaWorker)
    Else
        Call QuitarSta(UserIndex, Professions(ProfessionIndex).RequiredStaminaOther)
    End If
    
    SkillNumber = Professions(ProfessionIndex).SkillNumber
    skill = GetSkills(UserIndex, SkillNumber)
    
    'Old Formula
    'Suerte = Int(-0.00125 * skill * skill - 0.3 * skill + 49)
    'res = RandomNumber(1, Suerte)
    
    'New Formula
    ModClass = Classes(.clase).ClassMods.Work
    'With low skill level, probability is REALLY low
    Probability = 20 + skill * ModClass
    res = RandomNumber(1, 100)

    If res <= Probability Then
    
        With MapData(Pos.Map, Pos.X, Pos.Y).ObjInfo
            ObjIndex = .ObjIndex
            
            Dim ResourceIndex As Integer
            For ResourceIndex = 1 To UBound(.Resources)
                Dim MiObj As Obj
                
                
                CanExtractFromResource = ((ObjData(ObjIndex).Resources(ResourceIndex).UnlimitedResource) Or _
                                        (Not ObjData(ObjIndex).Resources(ResourceIndex).UnlimitedResource And .Resources(ResourceIndex).MaxAvailableQuantity > 0)) And _
                                        ObjData(ToolIndex).ToolPower >= .Resources(ResourceIndex).MinToolPower
                                        
                ' Calculate the probability
                If RandomDecimalNumber(0, 100) <= (.Resources(ResourceIndex).ExtractionProbability) And CanExtractFromResource Then
                    
                    If UserList(UserIndex).clase = eClass.Worker Then
                        MinQty = .Resources(ResourceIndex).MinPerTickWorker
                        MaxQty = .Resources(ResourceIndex).MaxPerTickWorker
                    Else
                        MinQty = .Resources(ResourceIndex).MinPerTickOther
                        MaxQty = .Resources(ResourceIndex).MaxPerTickOther
                    End If
                    
                    MiObj.Amount = RandomNumber(MinQty, MaxQty)
                    
                    If MiObj.Amount > 0 Then
                        HasExtracted = True
                        Call ResourceExtract(UserIndex, Pos, ResourceIndex, MiObj.Amount)
                    
                        MiObj.ObjIndex = .Resources(ResourceIndex).ObjIndex
                    
                        If Not MeterItemEnInventario(UserIndex, MiObj) Then
                            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
                        End If
                    End If
                    
                End If
                
            Next ResourceIndex
            
            If .PendingQty <= 0 Then
                Call ExhaustResource(UserIndex, Pos)
            End If
        End With
    End If
    
    If HasExtracted Then
        'Verify if exp is gained with each sub resource or the total extraction
        Call SubirSkill(UserIndex, SkillNumber, True, Professions(ProfessionIndex).SkillExpSuccess)
        Call WriteConsoleMsg(UserIndex, "¡Has conseguido recursos!", FontTypeNames.FONTTYPE_INFO, eMessageType.Trabajo)
    Else
        '[CDT 17-02-2004]
        If Not .flags.UltimoMensaje = 8 Then
            Call WriteConsoleMsg(UserIndex, "No has obtenido nada", FontTypeNames.FONTTYPE_INFO, eMessageType.Trabajo)
            .flags.UltimoMensaje = 8
        End If
        '[/CDT]
        Call SubirSkill(UserIndex, SkillNumber, False, Professions(ProfessionIndex).SkillExpFailure)
    End If
        
End With

Exit Sub

ErrHandler:
    Call LogError("Error en DoExtractResource: " & Err.Description)
End Sub

Public Sub DoMineria(ByVal UserIndex As Integer)
'***************************************************
'Autor: Unknown
'Last Modification: 28/05/2010
'16/11/2009: ZaMa - Implementado nuevo sistema de extraccion.
'11/05/2010: ZaMa - Arreglo formula de maximo de items contruibles/extraibles.
'05/13/2010: Pato - Refix a la formula de maximo de items construibles/extraibles.
'22/05/2010: ZaMa - Los caos ya no suben plebe al trabajar.
'28/05/2010: ZaMa - Los pks no suben plebe al trabajar.
'09/07/2016: Anagrama - Ahora la plebe no sube solo si es criminal y tiene seguro activado.
'***************************************************
On Error GoTo ErrHandler

Dim Suerte As Integer
Dim res As Integer
Dim CantidadItems As Integer

With UserList(UserIndex)
    If .clase = eClass.Worker Then
        Call QuitarSta(UserIndex, ConstantesTrabajo.EsfuerzoExcavarMinero)
    Else
        Call QuitarSta(UserIndex, ConstantesTrabajo.EsfuerzoExcavarGeneral)
    End If
    
    Dim skill As Integer
    skill = GetSkills(UserIndex, eSkill.Mineria)
    Suerte = Int(-0.00125 * skill * skill - 0.3 * skill + 49)
    
    res = RandomNumber(1, Suerte)
    
    If res <= 5 Then
        Dim MiObj As Obj
        
        If .flags.TargetObj = 0 Then Exit Sub
        
        MiObj.ObjIndex = ObjData(.flags.TargetObj).MineralIndex
        
        If .clase = eClass.Worker Then
            CantidadItems = MaxItemsExtraibles(.Stats.ELV)
            
            MiObj.Amount = RandomNumber(1, CantidadItems)
        Else
            MiObj.Amount = 1
        End If
        
        If Not MeterItemEnInventario(UserIndex, MiObj) Then _
            Call TirarItemAlPiso(.Pos, MiObj)
        
        Call WriteConsoleMsg(UserIndex, "¡Has extraido algunos minerales!", FontTypeNames.FONTTYPE_INFO, eMessageType.Trabajo)
        
        Call SubirSkill(UserIndex, eSkill.Mineria, True)
    Else
        '[CDT 17-02-2004]
        If Not .flags.UltimoMensaje = 9 Then
            Call WriteConsoleMsg(UserIndex, "¡No has conseguido nada!", FontTypeNames.FONTTYPE_INFO, eMessageType.Trabajo)
            .flags.UltimoMensaje = 9
        End If
        '[/CDT]
        Call SubirSkill(UserIndex, eSkill.Mineria, False)
    End If
    
End With

Exit Sub

ErrHandler:
    Call LogError("Error en Sub DoMineria")

End Sub

Public Sub DoMeditar(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 26/04/2015
'26/04/2015: D'Artagnan - Interval changed.
'***************************************************
On Error GoTo ErrHandler
  

    With UserList(UserIndex)
        Dim Suerte As Integer
        Dim cant As Integer
    
        'Barrin 3/10/03
        'Esperamos a que se termine de concentrar
        Dim TActual As Long
        TActual = GetTickCount()
        
        Dim iInterval As Integer
        
        iInterval = 3500 - .Stats.ELV * 40
               
        If GetInterval(TActual, .Counters.tInicioMeditar) < iInterval Then
            Exit Sub
        End If
        
        If .Counters.bPuedeMeditar = False Then
            .Counters.bPuedeMeditar = True
        End If
            
        If .Stats.MinMAN >= .Stats.MaxMan Then
            Call WriteConsoleMsg(UserIndex, "Has terminado de meditar.", FontTypeNames.FONTTYPE_INFO)
            Call WriteMeditateToggle(UserIndex)
            .flags.Meditando = False
            .Char.FX = 0
            .Char.Loops = 0
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0, , 0))
            Exit Sub
        End If
         
        Suerte = 7.5 + GetSkills(UserIndex, eSkill.Meditar) * 0.25
        
        If Suerte > RandomNumber(1, 100) Then
           ' Succeeded
           Cant = Porcentaje(.Stats.MaxMan, PorcentajeRecuperoMana)
           If Cant <= 0 Then Cant = 1
           .Stats.MinMAN = .Stats.MinMAN + Cant
           If .Stats.MinMAN > .Stats.MaxMan Then _
               .Stats.MinMAN = .Stats.MaxMan
           
           If Not .flags.UltimoMensaje = 22 Then
               Call WriteConsoleMsg(UserIndex, "¡Has recuperado " & cant & " puntos de maná!", FontTypeNames.FONTTYPE_INFO)
               .flags.UltimoMensaje = 22
           End If
           
           Call WriteUpdateMana(UserIndex)
           Call SubirSkill(UserIndex, eSkill.Meditar, True)
        Else
            Call SubirSkill(UserIndex, eSkill.Meditar, False)
            Exit Sub
        End If

    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub DoMeditar de Trabajo.bas")
End Sub

Public Sub DoDesequipar(ByVal UserIndex As Integer, ByVal VictimIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modif: 15/04/2010
'Unequips either shield, weapon or helmet from target user.
'***************************************************
On Error GoTo ErrHandler
  

    Dim Probabilidad As Integer
    Dim Resultado As Integer
    Dim WrestlingSkill As Byte
    Dim AlgoEquipado As Boolean
    
    With UserList(UserIndex)
        ' Si no tiene guantes de hurto no desequipa.
        If .Invent.AnilloEqpObjIndex <> ConstantesItems.GuanteHurto Then Exit Sub
        
        ' Si no esta solo con manos, no desequipa tampoco.
        If .Invent.WeaponEqpObjIndex > 0 Then Exit Sub
        
        WrestlingSkill = GetSkills(UserIndex, eSkill.Wrestling)
        
        Probabilidad = WrestlingSkill * 0 + .Stats.ELV * 0
   End With
   
   With UserList(VictimIndex)
        ' Si tiene escudo, intenta desequiparlo
        If .Invent.EscudoEqpObjIndex > 0 Then
            
            Resultado = RandomNumber(1, 100)
            
            If Resultado <= Probabilidad Then
                ' Se lo desequipo
                Call Desequipar(VictimIndex, .Invent.EscudoEqpSlot, True)
                
                Call WriteConsoleMsg(UserIndex, "Has logrado desequipar el escudo de tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
                
                If .Stats.ELV < 20 Then
                    Call WriteConsoleMsg(VictimIndex, "¡Tu oponente te ha desequipado el escudo!", FontTypeNames.FONTTYPE_FIGHT)
                End If

                Exit Sub
            End If
            
            AlgoEquipado = True
        End If
        
        ' No tiene escudo, o fallo desequiparlo, entonces trata de desequipar arma
        If .Invent.WeaponEqpObjIndex > 0 Then
            
            Resultado = RandomNumber(1, 100)
            
            If Resultado <= Probabilidad Then
                ' Se lo desequipo
                Call Desequipar(VictimIndex, .Invent.WeaponEqpSlot, True)
                
                Call WriteConsoleMsg(UserIndex, "¡Has logrado desarmar a tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
                
                If .Stats.ELV < 20 Then
                    Call WriteConsoleMsg(VictimIndex, "¡Tu oponente te ha desarmado!", FontTypeNames.FONTTYPE_FIGHT)
                End If

                Exit Sub
            End If
            
            AlgoEquipado = True
        End If
        
        ' No tiene arma, o fallo desequiparla, entonces trata de desequipar casco
        If .Invent.CascoEqpObjIndex > 0 Then
            
            Resultado = RandomNumber(1, 100)
            
            If Resultado <= Probabilidad Then
                ' Se lo desequipo
                Call Desequipar(VictimIndex, .Invent.CascoEqpSlot, True)
                
                Call WriteConsoleMsg(UserIndex, "Has logrado desequipar el casco de tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
                
                If .Stats.ELV < 20 Then
                    Call WriteConsoleMsg(VictimIndex, "¡Tu oponente te ha desequipado el casco!", FontTypeNames.FONTTYPE_FIGHT)
                End If

                Exit Sub
            End If
            
            AlgoEquipado = True
        End If
    
        If AlgoEquipado Then
            Call WriteConsoleMsg(UserIndex, "Tu oponente no tiene equipado items!", FontTypeNames.FONTTYPE_FIGHT)
        Else
            Call WriteConsoleMsg(UserIndex, "No has logrado desequipar ningún item a tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
        End If
    
    End With


  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub DoDesequipar de Trabajo.bas")
End Sub

Public Sub DoHurtar(ByVal UserIndex As Integer, ByVal VictimaIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modif: 03/03/2010
'Implements the pick pocket skill of the Bandit :)
'03/03/2010 - Pato: Sólo se puede hurtar si no está en trigger 6 :)
'***************************************************
On Error GoTo ErrHandler
  
Dim OtroUserIndex As Integer

If TriggerZonaPelea(UserIndex, VictimaIndex) <> TRIGGER6_AUSENTE Then Exit Sub

If UserList(UserIndex).clase <> eClass.Bandit Then Exit Sub
'Esto es precario y feo, pero por ahora no se me ocurrió nada mejor.
'Uso el slot de los anillos para "equipar" los guantes.
'Y los reconozco porque les puse DefensaMagicaMin y Max = 0
If UserList(UserIndex).Invent.AnilloEqpObjIndex <> ConstantesItems.GuanteHurto Then Exit Sub

Dim res As Integer
res = RandomNumber(1, 100)
If (res < 20) Then
    If TieneObjetosRobables(VictimaIndex) Then
    
        If isTradingWithUser(VictimaIndex) Then
            OtroUserIndex = getTradingUser(VictimaIndex)
                
            If (OtroUserIndex > 0) And (OtroUserIndex <= MaxUsers) Then
                Call WriteConsoleMsg(VictimaIndex, "¡¡Comercio cancelado, te están robando!!", FontTypeNames.FONTTYPE_TALK)
                Call WriteConsoleMsg(OtroUserIndex, "¡¡Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
                
                Call LimpiarComercioSeguro(VictimaIndex)
            End If
        End If
                
        Call RobarObjeto(UserIndex, VictimaIndex)
        Call WriteConsoleMsg(VictimaIndex, "¡" & UserList(UserIndex).Name & " es un Bandido!", FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(UserIndex, UserList(VictimaIndex).Name & " no tiene objetos.", FontTypeNames.FONTTYPE_INFO)
    End If
End If

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub DoHurtar de Trabajo.bas")
End Sub

Public Sub DoHandInmo(ByVal UserIndex As Integer, ByVal VictimaIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modif: 17/02/2007
'Implements the special Skill of the Thief
'***************************************************
On Error GoTo ErrHandler
  
    If UserList(VictimaIndex).flags.Paralizado = 1 Then Exit Sub
    
    If Not PuedeAcuchillar(UserIndex) Then Exit Sub
    
    Dim res As Integer
    res = RandomNumber(0, 100)
        
    If res < (GetSkills(UserIndex, eSkill.Wrestling) * 0.25 + UserList(UserIndex).Stats.ELV * 0) Then
    
        If HasPassiveAssigned(VictimaIndex, ePassiveSpells.Berserk) And HasPassiveActivated(VictimaIndex, ePassiveSpells.Berserk) Then
            Call WriteConsoleMsg(UserIndex, "Las habilidades de " & UserList(VictimaIndex).Name & " lo protegen de tu habilidad inmovilizante.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
            Call WriteConsoleMsg(VictimaIndex, "Tu habilidades te protegen del hechizo inmovilizante de " & UserList(UserIndex).Name & ".", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
            Exit Sub
        End If
    
        UserList(VictimaIndex).flags.Paralizado = 1
        UserList(VictimaIndex).Counters.Paralisis = SetIntervalEnd(ServerConfiguration.Intervals.IntervaloParalizado / 2)
        
        UserList(VictimaIndex).flags.ParalizedByIndex = UserIndex
        UserList(VictimaIndex).flags.ParalizedBy = UserList(UserIndex).Name
        
        Call WriteParalizeOK(VictimaIndex)
        Call WriteConsoleMsg(UserIndex, "Tu golpe ha dejado inmóvil a tu oponente", FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(VictimaIndex, "¡El golpe te ha dejado inmóvil!", FontTypeNames.FONTTYPE_INFO)
    End If

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub DoHandInmo de Trabajo.bas")
End Sub

Public Sub DoHandInmoNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modif: 05/09/2012
'Implements the special Skill of the Thief
'***************************************************
On Error GoTo ErrHandler
  
    
    With Npclist(NpcIndex)
        If .flags.Paralizado = 1 Then Exit Sub
        If .flags.Inmovilizado = 1 Then Exit Sub
        If .flags.AfectaParalisis = 1 Then Exit Sub
        
        If Not PuedeAcuchillar(UserIndex) Then Exit Sub
            
        Dim res As Integer
        res = RandomNumber(0, 100)
        If res < (GetSkills(UserIndex, eSkill.Wrestling) / 4) Then
        
            .flags.Inmovilizado = 1
            .flags.Paralizado = 0
            .Contadores.Paralisis = SetIntervalEnd(ServerConfiguration.Intervals.IntervaloNPCParalizado)
            
            Call WriteConsoleMsg(UserIndex, "Tu golpe ha dejado inmóvil a la criatura", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub DoHandInmoNpc de Trabajo.bas")
End Sub

Public Sub Desarmar(ByVal UserIndex As Integer, ByVal VictimIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 29/09/2015
'02/04/2010: ZaMa - Nueva formula para desarmar.
'29/09/2015: D'Artagnan - Abort if victim has no weapon.
'***************************************************
On Error GoTo ErrHandler
  

    Dim Probabilidad As Integer
    Dim Resultado As Integer
    Dim WrestlingSkill As Byte
    
    If Not PuedeAcuchillar(UserIndex) Or _
       UserList(VictimIndex).Invent.WeaponEqpObjIndex = 0 Then Exit Sub
    
    With UserList(UserIndex)
        WrestlingSkill = GetSkills(UserIndex, eSkill.Wrestling)
        
        Probabilidad = WrestlingSkill * 0.2 + .Stats.ELV * 0.66
        
        Resultado = RandomNumber(1, 100)
        
        If Resultado <= Probabilidad Then
            Call Desequipar(VictimIndex, UserList(VictimIndex).Invent.WeaponEqpSlot, True)
            Call WriteConsoleMsg(UserIndex, "¡Has logrado desarmar a tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
            If UserList(VictimIndex).Stats.ELV < 20 Then
                Call WriteConsoleMsg(VictimIndex, "¡Tu oponente te ha desarmado!", FontTypeNames.FONTTYPE_FIGHT)
            End If
        End If
    End With
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Desarmar de Trabajo.bas")
End Sub

Public Function MaxItemsConstruibles(ByVal UserIndex As Integer) As Integer
'***************************************************
'Author: ZaMa
'Last Modification: 29/01/2010
'11/05/2010: ZaMa - Arreglo formula de maximo de items contruibles/extraibles.
'05/13/2010: Pato - Refix a la formula de maximo de items construibles/extraibles.
'***************************************************
On Error GoTo ErrHandler
  
    
    With UserList(UserIndex)
        If .clase = eClass.Worker Then
            MaxItemsConstruibles = MaximoInt(1, CInt((.Stats.ELV - 2) * 0.2))
        Else
            MaxItemsConstruibles = 1
        End If
    End With
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function MaxItemsConstruibles de Trabajo.bas")
End Function

Public Function MaxItemsExtraibles(ByVal UserLevel As Integer) As Integer
'***************************************************
'Author: ZaMa
'Last Modification: 14/05/2010
'***************************************************
On Error GoTo ErrHandler
  
    MaxItemsExtraibles = MaximoInt(1, CInt((UserLevel - 2) * 0.2)) + 1
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function MaxItemsExtraibles de Trabajo.bas")
End Function

Public Sub ImitateNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 20/11/2010
'Copies body, head and desc from previously clicked npc.
'13/07/2016: Anagrama - Ahora se copia casco, escudo y arma tambien.
'***************************************************
On Error GoTo ErrHandler
  
    
    With UserList(UserIndex)
        
        ' Copy desc
        .DescRM = Npclist(NpcIndex).Name
        
          ' Save original char
        If .flags.Navegando = 0 And .flags.Mimetizado = 0 Then
            .OrigChar.body = .Char.body
            .OrigChar.head = .Char.head
            .OrigChar.CascoAnim = .Char.CascoAnim
            .OrigChar.ShieldAnim = .Char.ShieldAnim
            .OrigChar.WeaponAnim = .Char.WeaponAnim
        End If
        
        .flags.Mimetizado = 1
        .Counters.Mimetismo = -1
        
        ' Remove Anims (Npcs don't use equipment anims yet)
        ' 13/07/2016: They do now.
        .Char.CascoAnim = Npclist(NpcIndex).Char.CascoAnim
        .Char.ShieldAnim = Npclist(NpcIndex).Char.ShieldAnim
        .Char.WeaponAnim = Npclist(NpcIndex).Char.WeaponAnim
        
        ' If admin is invisible the store it in old char
        .Char.body = Npclist(NpcIndex).Char.body
        .Char.head = Npclist(NpcIndex).Char.head
            
        Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
    
    End With
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ImitateNpc de Trabajo.bas")
End Sub

Public Sub PerformMenuAction(ByVal UserIndex As Integer, ByVal iAction As Integer, ByVal Slot As Byte)
'***************************************************
'Author: ZaMa
'Last Modification: 23/03/2012
'Performs menu action.
'***************************************************
On Error GoTo ErrHandler

    With UserList(UserIndex)
        
        ' Dead is invalid
        If (.flags.Muerto = 1) Then Exit Sub
    
        If Slot > 0 And Slot <= .CurrentInventorySlots Then
            .flags.TargetObjInvSlot = Slot
            .flags.TargetObjInvIndex = .Invent.Object(Slot).ObjIndex
        End If

        Select Case iAction
            
            Case eMenuAction.ieBlacksmith
                If CanBlacksmith(UserIndex) Then
                    Call WriteCraftableRecipes(UserIndex, eProfessions.Blacksmithing)
                    Call WriteShowCraftForm(UserIndex)
                End If
                
            Case eMenuAction.ieMakeLingot
                If CanMelt(UserIndex, True) Then
                    Call FundirMineral(UserIndex)
                End If
                
            Case eMenuAction.ieMeltDown
                'If CanMelt(UserIndex, False) Then
                '    If ObjData(.Invent.Object(.flags.TargetObjInvSlot).ObjIndex).SkHerreria > 0 Then
                '        Call FundirArmas(UserIndex)
                '    End If
                'End If
                
            Case eMenuAction.ieTameNpc
                
                Dim NpcIndex As Integer
                NpcIndex = .flags.TargetNpc
                If NpcIndex > 0 Then
                    If CanTameNpc(UserIndex, NpcIndex) Then
                        Call DoTameNpc(UserIndex, NpcIndex)
                    End If
                End If
                
            Case eMenuAction.ieLightFire
                If .flags.TargetObj = ConstantesItems.FogataApagada Or .flags.TargetObj = ConstantesItems.RamitaElfica Then
                    'If the passiveskill is not assigned, do nothing
                    If .flags.TargetObj = ConstantesItems.RamitaElfica And Not HasPassiveAssigned(UserIndex, VitalRestoration) Then
                             If (.clase = eClass.Thief Or .clase = eClass.Worker Or .clase = eClass.Warrior Or .clase = Hunter) Then
                                Call WriteConsoleMsg(UserIndex, "Debes tener la habilidad pasiva restauración vital.", FontTypeNames.FONTTYPE_INFO)
                            Else
                                Call WriteConsoleMsg(UserIndex, "Tu clase no te permite realizar fogatas élficas", FontTypeNames.FONTTYPE_INFO)
                            End If
                        Exit Sub
                    End If
                
                    Call AccionParaRamita(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY, UserIndex)
                End If
        End Select
    End With
    
    Exit Sub
    
ErrHandler:
    Call LogError("Error en PerformMenuAction. Userindex: " & UserIndex & " Action: " & iAction & " Slot: " & Slot)
End Sub

Public Function CanTameNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer) As Boolean
'***************************************************
'Author: ZaMa
'Last Modification: 23/03/2012
'Checks if user can tame npc.
'***************************************************
On Error GoTo ErrHandler
  

    With Npclist(NpcIndex)
        If .flags.Domable = 0 Then
            Call WriteConsoleMsg(UserIndex, "No puedes domar a esa criatura.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
    
        If Abs(UserList(UserIndex).Pos.X - .Pos.X) + Abs(UserList(UserIndex).Pos.Y - .Pos.Y) > 2 Then
            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        
        If LenB(.flags.AttackedBy) <> 0 Then
            Call WriteConsoleMsg(UserIndex, "No puedes domar una criatura que está luchando con un jugador.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
    End With
    
    CanTameNpc = True
    
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function CanTameNpc de Trabajo.bas")
End Function

Public Function CanMelt(ByVal UserIndex As Integer, ByVal ValidarMinerales As Boolean) As Boolean
'***************************************************
'Author: ZaMa
'Last Modification: 23/03/2012
'Checks if user can melt weapon or mineral.
'***************************************************
On Error GoTo ErrHandler
  

    'Check interval
    If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Function
    
    Dim ObjIndex As Integer
    
    'Check there is a proper item there
    With UserList(UserIndex)
        
        ObjIndex = .flags.TargetObj
    
        If ObjIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, "Ahí no hay ninguna fragua.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        
        If ObjData(ObjIndex).ObjType <> eOBJType.otFragua Then
            Call WriteConsoleMsg(UserIndex, "Ahí no hay ninguna fragua.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
            
        Dim TargetObjInvSlot As Integer
        TargetObjInvSlot = .flags.TargetObjInvSlot
    
        'Validate other items
        If TargetObjInvSlot < 1 Or TargetObjInvSlot > .CurrentInventorySlots Then
            Exit Function
        End If
        
        Dim ObjSlotIndex As Integer
        ObjSlotIndex = .Invent.Object(TargetObjInvSlot).ObjIndex
        
        If ObjSlotIndex = 0 Then
            Exit Function
        End If
        
        ''chequeamos que no se zarpe duplicando oro
        If ValidarMinerales Then
            If (ObjData(ObjSlotIndex).ObjType = eOBJType.otMinerales) Then
                If .Invent.Object(TargetObjInvSlot).ObjIndex <> .flags.TargetObjInvIndex Then
                    If .Invent.Object(TargetObjInvSlot).ObjIndex = 0 Or .Invent.Object(TargetObjInvSlot).Amount = 0 Then
                        Call WriteConsoleMsg(UserIndex, "No tienes más minerales.", FontTypeNames.FONTTYPE_INFO)
                        Exit Function
                    End If
                    
                    ''FUISTE
                    Call DisconnectWithMessage(UserIndex, "Has sido expulsado por el sistema anti cheats.")

                    Exit Function
                End If
            Else
                Call WriteConsoleMsg(UserIndex, "Debes seleccionar algún mineral para hacer lingotes.", FontTypeNames.FONTTYPE_INFO)
                Exit Function
            End If
        End If
    End With
    
    CanMelt = True
    
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function CanMelt de Trabajo.bas")
End Function

Public Function CanBlacksmith(ByRef UserIndex As Integer) As Boolean
'***************************************************
'Author: ZaMa
'Last Modification: 23/03/2012
'Checks if user can blackmith.
'***************************************************
On Error GoTo ErrHandler
  

    Dim Slot As Integer

    With UserList(UserIndex)
        
        Slot = .flags.TargetObjInvSlot
        
        If Slot < 1 Or Slot > .CurrentInventorySlots Then Exit Function
    
        If .flags.TargetObj = 0 Then
            Call WriteConsoleMsg(UserIndex, "Ahí no hay ningún yunque.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        
        If ObjData(.flags.TargetObj).ObjType <> eOBJType.otYunque Then
            Call WriteConsoleMsg(UserIndex, "Ahí no hay ningún yunque.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        
        ' Using hammer?
        If (.Invent.Object(Slot).ObjIndex <> ConstantesItems.MartilloHerrero And .Invent.Object(Slot).ObjIndex <> ConstantesItems.MartilloHerreroNW) Then
            Call WriteConsoleMsg(UserIndex, "Debes seleccionar el martillo de herrero para usar el yunque.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        
        ' Equiped?
        If (.Invent.Object(Slot).Equipped = 0) Then
            Call WriteConsoleMsg(UserIndex, "Antes de usar la herramienta deberías equipartela.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If

    End With
    
    CanBlacksmith = True
    
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function CanBlacksmith de Trabajo.bas")
End Function

Public Function CanMakeFireWood(ByVal UserIndex As Integer, ByVal Slot As Integer) As Boolean
'***************************************************
'Author: ZaMa
'Last Modification: 06/04/2017
'Checks if user can make FireWood.
'06/04/2017 - G Toyz - Los chequeos para la leña élfica los hago acá.
'***************************************************
On Error GoTo ErrHandler
  

    With UserList(UserIndex)
    
        If (.flags.Meditando) Then
            Exit Function    'The error message should have been provided by the client.
        End If

        ' Target is wood?
        If (.flags.TargetObj <> ConstantesItems.Leña And .flags.TargetObj <> ConstantesItems.LeñaElfica) Then
            Call WriteConsoleMsg(UserIndex, "Debes clickear sobre los troncos.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        
        ' Using dagger?
        If (.Invent.Object(Slot).ObjIndex <> ConstantesItems.Daga) Then
            Call WriteConsoleMsg(UserIndex, "Debes seleccionar una daga común no newbie para hacer ramitas.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        
        ' Valid interval?
        If Not IntervaloPermiteUsar(UserIndex) Then
            Exit Function
        End If
        
        ' Equiped?
        If (.Invent.Object(Slot).Equipped = 0) Then
            Call WriteConsoleMsg(UserIndex, "Antes de usar la herramienta deberías equipartela.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        
        If .flags.TargetObj = ConstantesItems.LeñaElfica Then
            If Not HasPassiveAssigned(UserIndex, ePassiveSpells.VitalRestoration) Then
                If (.clase = eClass.Thief Or .clase = eClass.Worker Or .clase = eClass.Warrior Or .clase = Hunter) Then
                    Call WriteConsoleMsg(UserIndex, "Debes tener la habilidad pasiva restauración vital.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, "Tu clase no te permite realizar fogatas élficas", FontTypeNames.FONTTYPE_INFO)
                End If
                Exit Function
            End If
        End If
    End With
    
    CanMakeFireWood = True
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function CanMakeFireWood de Trabajo.bas")
End Function

Public Sub ResourceExtract(UserIndex As Integer, Pos As WorldPos, ByVal ResourceIndex As Integer, ByRef Amount As Integer)
On Error GoTo ErrHandler

    With MapData(Pos.Map, Pos.X, Pos.Y).ObjInfo
                        
        If ObjData(.ObjIndex).Resources(ResourceIndex).MaxAvailableQuantity > 0 Then
            If .Resources(ResourceIndex).MaxAvailableQuantity < Amount Then
            Amount = .Resources(ResourceIndex).MaxAvailableQuantity
        End If
            ' Substact the extacted amount from this specific resource available cuantity
            .Resources(ResourceIndex).MaxAvailableQuantity = .Resources(ResourceIndex).MaxAvailableQuantity - Amount
        End If
        
        ' Substract the extracted amount from the Extraction Point available quantity
        
        If .PendingQty < Amount Then
            Amount = .PendingQty
        End If
    
        .PendingQty = .PendingQty - Amount
    End With
    
    Exit Sub
        
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function ResourceExtract de Trabajo.bas")
  
End Sub

Public Sub ExhaustResource(ByVal UserIndex As Integer, ByRef Pos As WorldPos)
    
    Dim ResourceObj As Integer
    
    With MapData(Pos.Map, Pos.X, Pos.Y).ObjInfo
        ResourceObj = .ObjIndex
        
        'Sends state server message to respawn this tree after the respawn time
        Call SendResourceToSpawn(Pos.Map, Pos.X, Pos.Y, .ObjIndex, ObjData(.ObjIndex).RespawnCooldown)
                        
        'Play WAV
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(ObjData(.ObjIndex).SoundNumber, Pos.X, Pos.Y))
            
        'Empty slot
        .PendingQty = 0
        .CurrentGrhIndex = ObjData(.ObjIndex).DepletedGrhIndex
        
        ' Remove the resource from the group
        Call ExhaustResourceFromGroup(ResourceObj, Pos.Map, Pos.X, Pos.Y)
                
        'Refresh
        If .ObjIndex > 0 Then
            If MapData(Pos.Map, Pos.X, Pos.Y).Trigger <> eTrigger.zonaOscura Then
                Call SendToItemArea(Pos.Map, Pos.X, Pos.Y, PrepareMessageObjectCreate(ObjData(.ObjIndex).DepletedGrhIndex, Pos.X, Pos.Y, ObjData(.ObjIndex).ObjType, 0, ObjData(.ObjIndex).Luminous, ObjData(.ObjIndex).LightOffsetX, ObjData(.ObjIndex).LightOffsetY, ObjData(.ObjIndex).LightSize, ObjData(.ObjIndex).CanBeTransparent))
            Else
                Call SendToItemAreaButCounselors(Pos.Map, Pos.X, Pos.Y, PrepareMessageObjectCreate(ObjData(.ObjIndex).DepletedGrhIndex, Pos.X, Pos.Y, ObjData(.ObjIndex).ObjType, 0, ObjData(.ObjIndex).Luminous, ObjData(.ObjIndex).LightOffsetX, ObjData(.ObjIndex).LightOffsetY, ObjData(.ObjIndex).LightSize, ObjData(.ObjIndex).CanBeTransparent))
            End If
        End If
    End With
    
    Exit Sub

End Sub
