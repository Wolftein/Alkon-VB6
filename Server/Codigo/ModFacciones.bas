Attribute VB_Name = "ModFacciones"
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

Public ArmaduraImperial1 As Integer
Public ArmaduraImperial2 As Integer
Public ArmaduraImperial3 As Integer
Public TunicaMagoImperial As Integer
Public TunicaMagoImperialEnanos As Integer
Public ArmaduraCaos1 As Integer
Public ArmaduraCaos2 As Integer
Public ArmaduraCaos3 As Integer
Public TunicaMagoCaos As Integer
Public TunicaMagoCaosEnanos As Integer

Public VestimentaImperialHumano As Integer
Public VestimentaImperialEnano As Integer
Public TunicaConspicuaHumano As Integer
Public TunicaConspicuaEnano As Integer
Public ArmaduraNobilisimaHumano As Integer
Public ArmaduraNobilisimaEnano As Integer
Public ArmaduraGranSacerdote As Integer

Public VestimentaLegionHumano As Integer
Public VestimentaLegionEnano As Integer
Public TunicaLobregaHumano As Integer
Public TunicaLobregaEnano As Integer
Public TunicaEgregiaHumano As Integer
Public TunicaEgregiaEnano As Integer
Public SacerdoteDemoniaco As Integer

Public Const NUM_RANGOS_FACCION As Integer = 15
Private Const NUM_DEF_FACCION_ARMOURS As Byte = 4

Public Enum eTipoDefArmors
    ieBaja
    ieMedia
    ieAlta
    ieMax
End Enum

Public Type tFaccionArmaduras
    Armada(NUM_DEF_FACCION_ARMOURS - 1) As Integer
    Caos(NUM_DEF_FACCION_ARMOURS - 1) As Integer
End Type

' Matriz que contiene las armaduras faccionarias segun raza, clase, faccion y defensa de armadura
Public ArmadurasFaccion(1 To NUMCLASES, 1 To NUMRAZAS) As tFaccionArmaduras

' Contiene la cantidad de exp otorgada cada vez que aumenta el rango
Public RecompensaFacciones(NUM_RANGOS_FACCION) As Long

Private Function GetArmourAmount(ByVal Rango As Integer, ByVal TipoDef As eTipoDefArmors) As Integer
'***************************************************
'Autor: ZaMa
'Last Modification: 15/04/2010
'Returns the amount of armours to give, depending on the specified rank
'***************************************************
On Error GoTo ErrHandler
  

    Select Case TipoDef
        
        Case eTipoDefArmors.ieBaja
            GetArmourAmount = 20 / (Rango + 1)
            
        Case eTipoDefArmors.ieMedia
            GetArmourAmount = Rango * 2 / MaximoInt((Rango - 4), 1)
            
        Case eTipoDefArmors.ieAlta
            GetArmourAmount = Rango * 1.35
            
    End Select
    
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetArmourAmount de ModFacciones.bas")
End Function

Private Sub GiveFactionArmours(ByVal UserIndex As Integer, ByVal IsCaos As Boolean)
'***************************************************
'Autor: ZaMa
'Last Modification: 15/04/2010
'Gives faction armours to user
'***************************************************
On Error GoTo ErrHandler
  
    
    Dim ObjArmour As Obj
    Dim Rango As Integer
    
    With UserList(UserIndex)
    
        Rango = Val(IIf(IsCaos, .Faccion.RecompensasCaos, .Faccion.RecompensasReal)) + 1
    
    
        ' Entrego armaduras de defensa baja
        ObjArmour.Amount = GetArmourAmount(Rango, eTipoDefArmors.ieBaja)
        
        If IsCaos Then
            ObjArmour.ObjIndex = ArmadurasFaccion(.clase, .raza).Caos(eTipoDefArmors.ieBaja)
        Else
            ObjArmour.ObjIndex = ArmadurasFaccion(.clase, .raza).Armada(eTipoDefArmors.ieBaja)
        End If
        
        If Not MeterItemEnInventario(UserIndex, ObjArmour) Then
            Call TirarItemAlPiso(.Pos, ObjArmour)
        End If
        
        
        ' Entrego armaduras de defensa media
        ObjArmour.Amount = GetArmourAmount(Rango, eTipoDefArmors.ieMedia)
        
        If IsCaos Then
            ObjArmour.ObjIndex = ArmadurasFaccion(.clase, .raza).Caos(eTipoDefArmors.ieMedia)
        Else
            ObjArmour.ObjIndex = ArmadurasFaccion(.clase, .raza).Armada(eTipoDefArmors.ieMedia)
        End If
        
        If Not MeterItemEnInventario(UserIndex, ObjArmour) Then
            Call TirarItemAlPiso(.Pos, ObjArmour)
        End If

    
        ' Entrego armaduras de defensa alta
        ObjArmour.Amount = GetArmourAmount(Rango, eTipoDefArmors.ieAlta)
        
        If IsCaos Then
            ObjArmour.ObjIndex = ArmadurasFaccion(.clase, .raza).Caos(eTipoDefArmors.ieAlta)
        Else
            ObjArmour.ObjIndex = ArmadurasFaccion(.clase, .raza).Armada(eTipoDefArmors.ieAlta)
        End If
        
        If Not MeterItemEnInventario(UserIndex, ObjArmour) Then
            Call TirarItemAlPiso(.Pos, ObjArmour)
        End If

    End With

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub GiveFactionArmours de ModFacciones.bas")
End Sub

Public Sub GiveExpReward(ByVal UserIndex As Integer, ByVal Rango As Long)
'***************************************************
'Autor: ZaMa
'Last Modification: 15/04/2010
'Gives reward exp to user
'***************************************************
On Error GoTo ErrHandler
  
    
    Dim GivenExp As Long
    
    With UserList(UserIndex)
        
        GivenExp = RecompensaFacciones(Rango)
        
        .Stats.Exp = .Stats.Exp + GivenExp
        
        Call WriteConsoleMsg(UserIndex, "Has sido recompensado con " & GivenExp & " puntos de experiencia.", FontTypeNames.FONTTYPE_FIGHT)

        Call CheckUserLevel(UserIndex)
        
    End With
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub GiveExpReward de ModFacciones.bas")
End Sub

Public Sub EnlistarArmadaReal(ByVal UserIndex As Integer)
'***************************************************
'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
'Last Modification: 15/04/2010
'Handles the entrance of users to the "Armada Real"
'15/03/2009: ZaMa - No se puede enlistar el fundador de un clan con alineación neutral.
'27/11/2009: ZaMa - Ahora no se puede enlistar un miembro de un clan neutro, por ende saque la antifaccion.
'15/04/2010: ZaMa - Cambio en recompensas iniciales.
'08/07/2016: Anagrama - Reducido el costo de nobleza para ingresar a 500.000 y eliminado requisito de frags para trabajadores.
'                       Reducida la cantidad de criminales matados necesarios a 15 y el siguiente rango a 50.
'***************************************************
On Error GoTo ErrHandler
  

With UserList(UserIndex)
    
    If .Faccion.ArmadaReal = 1 Or .Faccion.Alignment = eCharacterAlignment.FactionRoyal Then
        Call WriteChatOverHead(UserIndex, "Ya perteneces a las tropas reales", str(Npclist(.flags.TargetNpc).Char.CharIndex), vbWhite)
        Exit Sub
    End If
    
    If .Faccion.Alignment <> eCharacterAlignment.Neutral Then
        Call WriteChatOverHead(UserIndex, "Solo personajes Neutrales pueden unirse a las tropas reales.", str(Npclist(.flags.TargetNpc).Char.CharIndex), vbWhite)
        Exit Sub
    End If
    
    If .Stats.ELV < ConstantesBalance.FactionMinLevel Then
        Call WriteChatOverHead(UserIndex, "Para unirte a nuestras fuerzas debes ser al menos nivel 20.", str(Npclist(.flags.TargetNpc).Char.CharIndex), vbWhite)
        Exit Sub
    End If
         
    If .Faccion.Reenlistadas >= ConstantesBalance.FactionMaxRejoins Then
        Call WriteChatOverHead(UserIndex, "¡Has sido expulsado de las fuerzas reales demasiadas veces!", str(Npclist(.flags.TargetNpc).Char.CharIndex), vbWhite)
        Exit Sub
    End If
    
    ' If the user is in a guild, then we need to make sure he's the only member
    If .Guild.IdGuild > 0 Then
        If GuildList(.Guild.GuildIndex).Alignment <> eCharacterAlignment.FactionRoyal And GuildList(.Guild.GuildIndex).MemberCount > 1 Then
            Call WriteChatOverHead(UserIndex, "No puedes ingresar a las fuerzas reales perteneciendo a un clan que cuente con otros miembros además de tí.", str(Npclist(.flags.TargetNpc).Char.CharIndex), vbWhite)
            Exit Sub
        End If
        
        ' Set the guild's alignment to be the same one as the character's alignment.
        GuildList(.Guild.GuildIndex).Alignment = eCharacterAlignment.FactionRoyal
        Call modGuild_DB.UpdateGuildAlignment(GuildList(.Guild.GuildIndex).IdGuild, eCharacterAlignment.FactionRoyal)
    End If
  
    .Faccion.ArmadaReal = 1
    .Faccion.Alignment = eCharacterAlignment.FactionRoyal
    .Faccion.Reenlistadas = .Faccion.Reenlistadas + 1
    
    Call WriteChatOverHead(UserIndex, "¡¡¡Bienvenido al ejército real!!! Aquí tienes tus vestimentas. Cumple bien tu labor exterminando legionarios y me encargaré de recompensarte.", str(Npclist(.flags.TargetNpc).Char.CharIndex), vbWhite)
    
    ' TODO: Dejo esta variable por ahora, pero con chequear las reenlistadas deberia ser suficiente :S
    If .Faccion.RecibioArmaduraReal = 0 Then
        
        Call GiveExpReward(UserIndex, 0)
        
        .Faccion.RecibioArmaduraReal = 1
        .Faccion.NivelIngreso = .Stats.ELV
        .Faccion.FechaIngreso = Now
        'Esto por ahora es inútil, siempre va a ser cero, pero bueno, despues va a servir.
        .Faccion.MatadosIngreso = .Faccion.CiudadanosMatados
        
        .Faccion.RecibioExpInicialReal = 1
        .Faccion.RecompensasReal = 0
        .Faccion.NextRecompensa = 50
        
    End If
    
    Call RefreshCharStatus(UserIndex, False)
    
    Call LogEjercitoReal(.Name & " ingresó el " & Date & " cuando era nivel " & .Stats.ELV)
End With

    Call UpdateFactionaryItems(UserIndex)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub EnlistarArmadaReal de ModFacciones.bas")
End Sub

Public Sub RecompensaArmadaReal(ByVal UserIndex As Integer)
'***************************************************
'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
'Last Modification: 15/04/2010
'Handles the way of gaining new ranks in the "Armada Real"
'15/04/2010: ZaMa - Agrego recompensas de oro y armaduras
'08/07/2016: Anagrama - Modificados los requisitos para cada rango.
'***************************************************
On Error GoTo ErrHandler
  
Dim Crimis As Long
Dim Lvl As Byte
Dim NextRecom As Long
Dim Nobleza As Long

With UserList(UserIndex)
    Lvl = .Stats.ELV
    Crimis = .Faccion.CriminalesMatados
    NextRecom = .Faccion.NextRecompensa

    If Crimis < NextRecom Then
        Call WriteChatOverHead(UserIndex, "Mata " & NextRecom - Crimis & " criminales más para recibir la próxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
        Exit Sub
    End If
    
    Select Case NextRecom
        Case 50:
            If Lvl < 27 Then
                Call WriteChatOverHead(UserIndex, "Mataste suficientes criminales, pero te faltan " & 27 - Lvl & " niveles para poder recibir la próxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                Exit Sub
            End If
            .Faccion.RecompensasReal = 1
            .Faccion.NextRecompensa = 100
        
        Case 100:
            If Lvl < 29 Then
                Call WriteChatOverHead(UserIndex, "Mataste suficientes criminales, pero te faltan " & 29 - Lvl & " niveles para poder recibir la próxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                Exit Sub
            End If
            .Faccion.RecompensasReal = 2
            .Faccion.NextRecompensa = 150
        
        Case 150:
            If Lvl < 31 Then
                Call WriteChatOverHead(UserIndex, "Mataste suficientes criminales, pero te faltan " & 31 - Lvl & " niveles para poder recibir la próxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                Exit Sub
            End If
            .Faccion.RecompensasReal = 3
            .Faccion.NextRecompensa = 200
        
        Case 200:
            .Faccion.RecompensasReal = 4
            .Faccion.NextRecompensa = 300
        
        Case 300:
            .Faccion.RecompensasReal = 5
            .Faccion.NextRecompensa = 450
        
        Case 450:
            If Lvl < 33 Then
                Call WriteChatOverHead(UserIndex, "Mataste suficientes criminales, pero te faltan " & 27 - Lvl & " niveles para poder recibir la próxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                Exit Sub
            End If
            .Faccion.RecompensasReal = 6
            .Faccion.NextRecompensa = 600
        
        Case 600:
            .Faccion.RecompensasReal = 7
            .Faccion.NextRecompensa = 800
        
        Case 800:
            If Lvl < 35 Then
                Call WriteChatOverHead(UserIndex, "Mataste suficientes criminales, pero te faltan " & 35 - Lvl & " niveles para poder recibir la próxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                Exit Sub
            End If
            .Faccion.RecompensasReal = 8
            .Faccion.NextRecompensa = 1100
        
        Case 1100:
            .Faccion.RecompensasReal = 9
            .Faccion.NextRecompensa = 1400
        
        Case 1400:
            If Lvl < 36 Then
                Call WriteChatOverHead(UserIndex, "Mataste suficientes criminales, pero te faltan " & 36 - Lvl & " niveles para poder recibir la próxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                Exit Sub
            End If
            .Faccion.RecompensasReal = 10
            .Faccion.NextRecompensa = 1800
        
        Case 1800:
            If Lvl < 37 Then
                Call WriteChatOverHead(UserIndex, "Mataste suficientes criminales, pero te faltan " & 37 - Lvl & " niveles para poder recibir la próxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                Exit Sub
            End If
            .Faccion.RecompensasReal = 11
            .Faccion.NextRecompensa = 2200
        
        Case 2200:
            If Lvl < 38 Then
                Call WriteChatOverHead(UserIndex, "Mataste suficientes criminales, pero te faltan " & 38 - Lvl & " niveles para poder recibir la próxima recompensa.", str(Npclist(.flags.TargetNpc).Char.CharIndex), vbWhite)
                Exit Sub
            End If
            .Faccion.RecompensasReal = 12
            .Faccion.NextRecompensa = 2600
        
        Case 2600:
            If Lvl < 39 Then
                Call WriteChatOverHead(UserIndex, "Mataste suficientes criminales, pero te faltan " & 39 - Lvl & " niveles para poder recibir la próxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                Exit Sub
            End If
            .Faccion.RecompensasReal = 13
            .Faccion.NextRecompensa = 3000
        
        Case 3000:
            If Lvl < 40 Then
                Call WriteChatOverHead(UserIndex, "Mataste suficientes criminales, pero te faltan " & 40 - Lvl & " niveles para poder recibir la próxima recompensa.", str(Npclist(.flags.TargetNpc).Char.CharIndex), vbWhite)
                Exit Sub
            End If
            .Faccion.RecompensasReal = 14
            .Faccion.NextRecompensa = 10000
        
        Case 10000:
            Call WriteChatOverHead(UserIndex, "Eres uno de mis mejores soldados. Mataste " & Crimis & " criminales, sigue así. Ya no tengo más recompensa para darte que mi agradecimiento. ¡Felicidades!", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        
        Case Else:
            Exit Sub
    End Select
    
    Call WriteChatOverHead(UserIndex, "¡¡¡Aquí tienes tu recompensa " & TituloReal(UserIndex) & "!!!", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)

    ' Recompensas de armaduras y exp
    Call GiveExpReward(UserIndex, .Faccion.RecompensasReal)

End With

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub RecompensaArmadaReal de ModFacciones.bas")
End Sub

Public Sub ExpulsarFaccionReal(ByVal UserIndex As Integer, Optional Expulsado As Boolean = True)
    On Error GoTo ErrHandler

    With UserList(UserIndex)
    
        If .Guild.IdGuild > 0 Then
            If GuildList(.Guild.GuildIndex).MemberCount > 1 Then
                 If Not Expulsado Then
                    Call WriteChatOverHead(UserIndex, "No puedes salir del Ejército Real siendo parte de un clan que tiene más miembros además de ti. Sal del clan antes.", str(Npclist(.flags.TargetNpc).Char.CharIndex), vbWhite)
                End If
                Exit Sub
            End If
            
            ' The player is the only guild member on the guild, so we can update it's alignment to match the new aligment of the character
            GuildList(.Guild.GuildIndex).Alignment = eCharacterAlignment.Neutral
            Call modGuild_DB.UpdateGuildAlignment(GuildList(.Guild.GuildIndex).IdGuild, eCharacterAlignment.Neutral)
        End If
    
        .Faccion.ArmadaReal = 0
        .Faccion.Alignment = eCharacterAlignment.Neutral
        
        If Expulsado Then
            .Faccion.Reenlistadas = 200
            Call WriteConsoleMsg(UserIndex, "¡¡¡Has sido expulsado del Ejército Real!!!", FontTypeNames.FONTTYPE_FIGHT)
        Else
            Call WriteConsoleMsg(UserIndex, "¡¡¡Te has retirado del Ejército Real!!!", FontTypeNames.FONTTYPE_FIGHT)
        End If
    
        Dim bRefresh As Boolean
    
        If .Invent.ArmourEqpObjIndex <> 0 Then
            'Desequipamos la armadura real si está equipada
            If ObjData(.Invent.ArmourEqpObjIndex).Real = 1 Then Call Desequipar(UserIndex, .Invent.ArmourEqpSlot, False)
            bRefresh = True
        End If
    
        If .Invent.EscudoEqpObjIndex <> 0 Then
            'Desequipamos el escudo de caos si está equipado
            If ObjData(.Invent.EscudoEqpObjIndex).Real = 1 Then Call Desequipar(UserIndex, .Invent.EscudoEqpSlot, False)
            bRefresh = True
        End If
        
        ' Actualizamos solo el slot que tenga un item faccionario.
        Call UpdateFactionaryItems(UserIndex)
    
        If bRefresh Then
            With .Char
                Call ChangeUserChar(UserIndex, .body, .head, .heading, .WeaponAnim, .ShieldAnim, .CascoAnim)
            End With
            Call WriteUpdateUserStats(UserIndex)

        End If
    
        If .flags.Navegando Then Call RefreshCharStatus(UserIndex, False) 'Actualizamos la barca si esta navegando (NicoNZ)
    End With
  
    Exit Sub
  
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ExpulsarFaccionReal de ModFacciones.bas")

End Sub

Public Sub ExpulsarFaccionCaos(ByVal UserIndex As Integer, Optional Expulsado As Boolean = True)
    On Error GoTo ErrHandler

    With UserList(UserIndex)
    
        If .Guild.IdGuild > 0 Then
            If GuildList(.Guild.GuildIndex).MemberCount > 1 Then
                 If Not Expulsado Then
                    Call WriteChatOverHead(UserIndex, "No puedes salir de la Legión Oscura siendo parte de un clan que tiene más miembros además de ti. Sal del clan antes.", str(Npclist(.flags.TargetNpc).Char.CharIndex), vbWhite)
                End If
                Exit Sub
            End If
            
            ' The player is the only guild member on the guild, so we can update it's alignment to match the new aligment of the character
            GuildList(.Guild.GuildIndex).Alignment = eCharacterAlignment.Neutral
            Call modGuild_DB.UpdateGuildAlignment(GuildList(.Guild.GuildIndex).IdGuild, eCharacterAlignment.Neutral)
        End If
    
        .Faccion.FuerzasCaos = 0
        .Faccion.Alignment = eCharacterAlignment.Neutral

        If Expulsado Then
            .Faccion.Reenlistadas = 200
            Call WriteConsoleMsg(UserIndex, "¡¡¡Has sido expulsado de la Legión Oscura!!!", FontTypeNames.FONTTYPE_FIGHT)
        Else
            Call WriteConsoleMsg(UserIndex, "¡¡¡Te has retirado de la Legión Oscura!!!", FontTypeNames.FONTTYPE_FIGHT)
        End If
    
        Dim bRefresh As Boolean
    
        If .Invent.ArmourEqpObjIndex <> 0 Then
            'Desequipamos la armadura de caos si está equipada
            If ObjData(.Invent.ArmourEqpObjIndex).Caos = 1 Then Call Desequipar(UserIndex, .Invent.ArmourEqpSlot, False)
            bRefresh = True
        End If
    
        If .Invent.EscudoEqpObjIndex <> 0 Then
            'Desequipamos el escudo de caos si está equipado
            If ObjData(.Invent.EscudoEqpObjIndex).Caos = 1 Then Call Desequipar(UserIndex, .Invent.EscudoEqpSlot, False)
            bRefresh = True
        End If
        
        ' Actualizamos solo el slot que tenga un item faccionario.
        Call UpdateFactionaryItems(UserIndex)
    
        If bRefresh Then
            With .Char
                Call ChangeUserChar(UserIndex, .body, .head, .heading, .WeaponAnim, .ShieldAnim, .CascoAnim)
            End With
            Call WriteUpdateUserStats(UserIndex)
        End If
    
        If .flags.Navegando Then Call RefreshCharStatus(UserIndex, False) 'Actualizamos la barca si esta navegando (NicoNZ)
    End With
  
    Exit Sub
  
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ExpulsarFaccionCaos de ModFacciones.bas")

End Sub

Private Sub UpdateFactionaryItems(ByVal UserIndex As Integer)

    Dim I        As Long
    Dim ObjIndex As Long
                 
    With UserList(UserIndex)

        For I = 1 To .CurrentInventorySlots

            With .Invent.Object(I)
                ObjIndex = .ObjIndex
                
                If ObjIndex > 0 Then
                
                    If ObjData(ObjIndex).Caos > 0 Or ObjData(ObjIndex).Real > 0 Then
           
                        Call UpdateUserInv(False, UserIndex, I)
                    End If

                End If

            End With
               
        Next I

    End With

End Sub

Public Function TituloReal(ByVal UserIndex As Integer) As String
'***************************************************
'Autor: Unknown
'Last Modification: 23/01/2007 Pablo (ToxicWaste)
'08/07/2016: Anagrama - Trabajadores tienen un titulo fijo.
'                       Modificados nombres de los rangos y otorgada distinción por género.
'Handles the titles of the members of the "Armada Real"
'***************************************************
On Error GoTo ErrHandler
  

If UserList(UserIndex).clase = eClass.Worker Then
    If UserList(UserIndex).Genero = eGenero.Hombre Then
        TituloReal = "Trabajador Real"
    Else
        TituloReal = "Trabajadora Real"
    End If
    Exit Function
End If

Select Case UserList(UserIndex).Faccion.RecompensasReal
'Rango 1: Aprendiz (30 Criminales)
'Rango 2: Escudero (70 Criminales)
'Rango 3: Soldado (130 Criminales)
'Rango 4: Sargento (210 Criminales)
'Rango 5: Caballero (320 Criminales)
'Rango 6: Comandante (460 Criminales)
'Rango 7: Capitán (640 Criminales + > lvl 27)
'Rango 8: Senescal (870 Criminales)
'Rango 9: Mariscal (1160 Criminales)
'Rango 10: Condestable (2000 Criminales + > lvl 30)
'Rangos de Honor de la Armada Real: (Consejo de Bander)
'Rango 11: Ejecutor Imperial (2500 Criminales + 2.000.000 Nobleza)
'Rango 12: Protector del Reino (3000 Criminales + 3.000.000 Nobleza)
'Rango 13: Avatar de la Justicia (3500 Criminales + 4.000.000 Nobleza + > lvl 35)
'Rango 14: Guardián del Bien (4000 Criminales + 5.000.000 Nobleza + > lvl 36)
'Rango 15: Campeón de la Luz (5000 Criminales + 6.000.000 Nobleza + > lvl 37)
    
    Case 0
        TituloReal = "Recluta"
    Case 1
        TituloReal = "Soldado"
    Case 2
        TituloReal = "Escudero"
    Case 3
        TituloReal = "Teniente"
    Case 4
        TituloReal = "Capitán"
    Case 5
        TituloReal = "Comandante"
    Case 6
        TituloReal = "Senescal"
    Case 7
        TituloReal = "Mariscal"
    Case 8
        TituloReal = "Ejecutor Imperial"
    Case 9
        If UserList(UserIndex).Genero = eGenero.Hombre Then
            TituloReal = "Emisario de la Paz"
        Else
            TituloReal = "Emisaria de la Paz"
        End If
    Case 10
        If UserList(UserIndex).Genero = eGenero.Hombre Then
            TituloReal = "Protector del Reino"
        Else
            TituloReal = "Protectora del Reino"
        End If
    Case 11
        If UserList(UserIndex).Genero = eGenero.Hombre Then
            TituloReal = "Defensor de la Ley"
        Else
            TituloReal = "Defensora de la Ley"
        End If
    Case 12
        If UserList(UserIndex).Genero = eGenero.Hombre Then
            TituloReal = "Guardián del Bien"
        Else
            TituloReal = "Guardiana del Bien"
        End If
    Case 13
        TituloReal = "Lider de la Justicia"
    Case Else
        If UserList(UserIndex).Genero = eGenero.Hombre Then
            TituloReal = "Campeón de la Luz"
        Else
            TituloReal = "Campeona de la Luz"
        End If
End Select


  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function TituloReal de ModFacciones.bas")
End Function

Public Sub EnlistarCaos(ByVal UserIndex As Integer)
'***************************************************
'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
'Last Modification: 27/11/2009
'15/03/2009: ZaMa - No se puede enlistar el fundador de un clan con alineación neutral.
'27/11/2009: ZaMa - Ahora no se puede enlistar un miembro de un clan neutro, por ende saque la antifaccion.
'08/07/2016: Anagrama - Los trabajadores ya no necesitan ciudadanos matados para ingresar.
'                       Reducidos los frags requeridos a 35 y el siguiente rango a 100.
'Handles the entrance of users to the "Legión Oscura"
'***************************************************
On Error GoTo ErrHandler
  

With UserList(UserIndex)

    If .Faccion.FuerzasCaos = 1 Or .Faccion.Alignment = eCharacterAlignment.FactionLegion Then
        Call WriteChatOverHead(UserIndex, "Ya perteneces a la legión oscura.", str(Npclist(.flags.TargetNpc).Char.CharIndex), vbWhite)
        Exit Sub
    End If
    
    If .Faccion.Alignment <> eCharacterAlignment.Neutral Then
        Call WriteChatOverHead(UserIndex, "Solo personajes Neutrales pueden unirse a la legión oscura.", str(Npclist(.flags.TargetNpc).Char.CharIndex), vbWhite)
        Exit Sub
    End If

    If .Stats.ELV < ConstantesBalance.FactionMinLevel Then
        Call WriteChatOverHead(UserIndex, "Para unirte a nuestras fuerzas debes ser al menos nivel 20.", str(Npclist(.flags.TargetNpc).Char.CharIndex), vbWhite)
        Exit Sub
    End If
    
    If .Faccion.Reenlistadas > 4 Then
        If .Faccion.Reenlistadas = 200 Then
            Call WriteChatOverHead(UserIndex, "Has sido expulsado de las fuerzas oscuras y durante tu rebeldía has atacado a mi ejército. ¡Vete de aquí!", str(Npclist(.flags.TargetNpc).Char.CharIndex), vbWhite)
        Else
            Call WriteChatOverHead(UserIndex, "¡Has sido expulsado de las fuerzas oscuras demasiadas veces!", str(Npclist(.flags.TargetNpc).Char.CharIndex), vbWhite)
        End If
        Exit Sub
    End If
    
    ' If the user is in a guild, then we need to make sure he's the only member
    If .Guild.IdGuild > 0 Then
        If GuildList(.Guild.GuildIndex).Alignment <> eCharacterAlignment.FactionLegion And GuildList(.Guild.GuildIndex).MemberCount > 1 Then
            Call WriteChatOverHead(UserIndex, "No puedes ingresar a las fuerzas oscuras perteneciendo a un clan de otra facción que cuente con otros miembros además de tí.", str(Npclist(.flags.TargetNpc).Char.CharIndex), vbWhite)
            Exit Sub
        End If
        
        ' Set the guild's alignment to be the same one as the character's alignment.
        GuildList(.Guild.GuildIndex).Alignment = eCharacterAlignment.FactionLegion
        Call modGuild_DB.UpdateGuildAlignment(GuildList(.Guild.GuildIndex).IdGuild, eCharacterAlignment.FactionLegion)
    End If
    
    .Faccion.Reenlistadas = .Faccion.Reenlistadas + 1
    .Faccion.FuerzasCaos = 1
    .Faccion.Alignment = eCharacterAlignment.FactionLegion
    
    Call WriteChatOverHead(UserIndex, "¡¡¡Bienvenido al lado oscuro!!! Aquí tienes tus armaduras. Derrama sangre real, y serás recompensado, lo prometo.", str(Npclist(.flags.TargetNpc).Char.CharIndex), vbWhite)
    
    If .Faccion.RecibioArmaduraCaos = 0 Then
                
        Call GiveExpReward(UserIndex, 0)
        
        .Faccion.RecibioArmaduraCaos = 1
        .Faccion.NivelIngreso = .Stats.ELV
        .Faccion.FechaIngreso = Now
    
        .Faccion.RecibioExpInicialCaos = 1
        .Faccion.RecompensasCaos = 0
        .Faccion.NextRecompensa = 100
    End If
    
    Call RefreshCharStatus(UserIndex, False)

    Call LogEjercitoCaos(.Name & " ingresó el " & Date & " cuando era nivel " & .Stats.ELV)
End With

    Call UpdateFactionaryItems(UserIndex)
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub EnlistarCaos de ModFacciones.bas")
End Sub

Public Sub RecompensaCaos(ByVal UserIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste) & Unknown (orginal version)
'Last Modification: 15/04/2010
'Handles the way of gaining new ranks in the "Legión Oscura"
'15/04/2010: ZaMa - Agrego recompensas de oro y armaduras
'08/07/2016: Anagrama - Modificados los requisitos de cada rango.
'***************************************************
On Error GoTo ErrHandler
  
Dim Ciudas As Long
Dim Lvl As Byte
Dim NextRecom As Long

With UserList(UserIndex)
    Lvl = .Stats.ELV
    Ciudas = .Faccion.CiudadanosMatados
    NextRecom = .Faccion.NextRecompensa
    
    If Ciudas < NextRecom Then
        Call WriteChatOverHead(UserIndex, "Mata " & NextRecom - Ciudas & " cuidadanos más para recibir la próxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
        Exit Sub
    End If
    
    Select Case NextRecom
        Case 100:
            If Lvl < 27 Then
                Call WriteChatOverHead(UserIndex, "Mataste suficientes ciudadanos, pero te faltan " & 27 - Lvl & " niveles para poder recibir la próxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                Exit Sub
            End If
            .Faccion.RecompensasCaos = 1
            .Faccion.NextRecompensa = 200
        
        Case 200:
            If Lvl < 29 Then
                Call WriteChatOverHead(UserIndex, "Mataste suficientes ciudadanos, pero te faltan " & 29 - Lvl & " niveles para poder recibir la próxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                Exit Sub
            End If
            .Faccion.RecompensasCaos = 2
            .Faccion.NextRecompensa = 300
        
        Case 300:
            If Lvl < 31 Then
                Call WriteChatOverHead(UserIndex, "Mataste suficientes ciudadanos, pero te faltan " & 31 - Lvl & " niveles para poder recibir la próxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                Exit Sub
            End If
            .Faccion.RecompensasCaos = 3
            .Faccion.NextRecompensa = 500
        
        Case 500:
            .Faccion.RecompensasCaos = 4
            .Faccion.NextRecompensa = 800
        
        Case 800:
            .Faccion.RecompensasCaos = 5
            .Faccion.NextRecompensa = 1100
        
        Case 1100:
            If Lvl < 33 Then
                Call WriteChatOverHead(UserIndex, "Mataste suficientes ciudadanos, pero te faltan " & 33 - Lvl & " niveles para poder recibir la próxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                Exit Sub
            End If
            .Faccion.RecompensasCaos = 6
            .Faccion.NextRecompensa = 1400
        
        Case 1400:
            .Faccion.RecompensasCaos = 7
            .Faccion.NextRecompensa = 1700
        
        Case 1700:
            If Lvl < 35 Then
                Call WriteChatOverHead(UserIndex, "Mataste suficientes ciudadanos, pero te faltan " & 35 - Lvl & " niveles para poder recibir la próxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                Exit Sub
            End If
            .Faccion.RecompensasCaos = 8
            .Faccion.NextRecompensa = 2000
        
        Case 2000:
            .Faccion.RecompensasCaos = 9
            .Faccion.NextRecompensa = 2400
        
        Case 2400:
            If Lvl < 36 Then
                Call WriteChatOverHead(UserIndex, "Mataste suficientes ciudadanos, pero te faltan " & 36 - Lvl & " niveles para poder recibir la próxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                Exit Sub
            End If
            .Faccion.RecompensasCaos = 10
            .Faccion.NextRecompensa = 2800
        
        Case 2800:
            If Lvl < 37 Then
                Call WriteChatOverHead(UserIndex, "Mataste suficientes ciudadanos, pero te faltan " & 37 - Lvl & " niveles para poder recibir la próxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                Exit Sub
            End If
            .Faccion.RecompensasCaos = 11
            .Faccion.NextRecompensa = 3200
        
        Case 3200:
            If Lvl < 38 Then
                Call WriteChatOverHead(UserIndex, "Mataste suficientes ciudadanos, pero te faltan " & 38 - Lvl & " niveles para poder recibir la próxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                Exit Sub
            End If
            .Faccion.RecompensasCaos = 12
            .Faccion.NextRecompensa = 3600
        
        Case 3600:
            If Lvl < 39 Then
                Call WriteChatOverHead(UserIndex, "Mataste suficientes ciudadanos, pero te faltan " & 39 - Lvl & " niveles para poder recibir la próxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                Exit Sub
            End If
            .Faccion.RecompensasCaos = 13
            .Faccion.NextRecompensa = 4000
        
        Case 4000:
            If Lvl < 40 Then
                Call WriteChatOverHead(UserIndex, "Mataste suficientes ciudadanos, pero te faltan " & 40 - Lvl & " niveles para poder recibir la próxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                Exit Sub
            End If
            .Faccion.RecompensasCaos = 14
            .Faccion.NextRecompensa = 23000
        
        Case 23000:
            Call WriteChatOverHead(UserIndex, "Eres uno de mis mejores soldados. Mataste " & Ciudas & " ciudadanos . Tu única recompensa será la sangre derramada. ¡¡Continúa así!!", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        
        Case Else:
            Exit Sub
            
    End Select
    
    Call WriteChatOverHead(UserIndex, "¡¡¡Bien hecho " & TituloCaos(UserIndex) & ", aquí tienes tu recompensa!!!", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
    
    ' Recompensas de armaduras y exp
    Call GiveExpReward(UserIndex, .Faccion.RecompensasCaos)
    
End With

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub RecompensaCaos de ModFacciones.bas")
End Sub

Public Function TituloCaos(ByVal UserIndex As Integer) As String
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 23/01/2007 Pablo (ToxicWaste)
'08/07/2016: Anagrama - El titulo de los trabajadores es fijo.
'                       Modificados los nombres de los rangos y otorgada distinción por género.
'Handles the titles of the members of the "Legión Oscura"
'***************************************************
'Rango 1: Acólito (70)
'Rango 2: Alma Corrupta (160)
'Rango 3: Paria (300)
'Rango 4: Condenado (490)
'Rango 5: Esbirro (740)
'Rango 6: Sanguinario (1100)
'Rango 7: Corruptor (1500 + lvl 27)
'Rango 8: Heraldo Impio (2010)
'Rango 9: Caballero de la Oscuridad (2700)
'Rango 10: Señor del Miedo (4600 + lvl 30)
'Rango 11: Ejecutor Infernal (5800 + lvl 31)
'Rango 12: Protector del Averno (6990 + lvl 33)
'Rango 13: Avatar de la Destrucción (8100 + lvl 35)
'Rango 14: Guardián del Mal (9300 + lvl 36)
'Rango 15: Campeón de la Oscuridad (11500 + lvl 37)
On Error GoTo ErrHandler
  

If UserList(UserIndex).clase = eClass.Worker Then
    If UserList(UserIndex).Genero = eGenero.Hombre Then
        TituloCaos = "Trabajador Caótico"
    Else
        TituloCaos = "Trabajadora Caótica"
    End If
    Exit Function
End If

Select Case UserList(UserIndex).Faccion.RecompensasCaos
    Case 0
        TituloCaos = "Acólito"
    Case 1
        TituloCaos = "Alma Siniestra"
    Case 2
        If UserList(UserIndex).Genero = eGenero.Hombre Then
            TituloCaos = "Servidor del Mal"
        Else
            TituloCaos = "Servidora del Mal"
        End If
    Case 3
        If UserList(UserIndex).Genero = eGenero.Hombre Then
            TituloCaos = "Acechador Sombrío"
        Else
            TituloCaos = "Acechadora Sombría"
        End If
    Case 4
        TituloCaos = "Esbirro"
    Case 5
        TituloCaos = "Verdugo"
    Case 6
        TituloCaos = "Vigía Espectral"
    Case 7
        If UserList(UserIndex).Genero = eGenero.Hombre Then
            TituloCaos = "Caballero Oscuro"
        Else
            TituloCaos = "Dama Oscura"
        End If
    Case 8
        TituloCaos = "Ejecutor Infernal"
    Case 9
        If UserList(UserIndex).Genero = eGenero.Hombre Then
            TituloCaos = "Emisario del Caos"
        Else
            TituloCaos = "Emisaria del Caos"
        End If
    Case 10
        If UserList(UserIndex).Genero = eGenero.Hombre Then
            TituloCaos = "Protector del Averno"
        Else
            TituloCaos = "Protectora del Averno"
        End If
    Case 11
        If UserList(UserIndex).Genero = eGenero.Hombre Then
            TituloCaos = "Señor del Miedo"
        Else
            TituloCaos = "Señora del Miedo"
        End If
    Case 12
        TituloCaos = "Bastión de la Destrucción"
    Case 13
        If UserList(UserIndex).Genero = eGenero.Hombre Then
            TituloCaos = "Devorador de Almas"
        Else
            TituloCaos = "Devoradora de Almas"
        End If
    Case Else
        If UserList(UserIndex).Genero = eGenero.Hombre Then
            TituloCaos = "Campeón de la Oscuridad"
        Else
            TituloCaos = "Campeona de la Oscuridad"
        End If
End Select

  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function TituloCaos de ModFacciones.bas")
End Function

Public Sub KickFromFactionByName(ByVal UserIndex As Integer, ByRef UserName As String)
On Error GoTo ErrHandler:
    Dim tUser As Integer
   
    tUser = NameIndex(UserName)
    
    Dim CurrentAlignment As Byte
    If tUser > 0 Then
        Dim CurrentAlignmentName As String
        CurrentAlignment = UserList(tUser).Faccion.Alignment
        
        If CurrentAlignment = eCharacterAlignment.Newbie Or CurrentAlignment = eCharacterAlignment.Neutral Then
            Call WriteConsoleMsg(UserIndex, "No puedes expulsar al usuario " & UserName & " de su facción porque se encuentra en una facción protegida (Newbie/Neutral).", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
       
        If UserList(tUser).Guild.IdGuild > 0 Then
            ' Can't kick a user from the faction because it will break the guild.
            If GuildList(UserList(tUser).Guild.GuildIndex).MemberCount > 1 Then
                Call WriteConsoleMsg(UserIndex, "No puedes expulsar al usuario " & UserName & " de su facción porque se encuentra en un clan que posee otros miembros. Hacer esto dejaría al clan en una alineación diferente a la del personaje.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            GuildList(UserList(tUser).Guild.GuildIndex).Alignment = eCharacterAlignment.Neutral
        End If
        
        ' This is dirty. Eventually we will get rid of the CHAOS and ARMY fields in the database
        ' so the faction system is generic and we only use
        If CurrentAlignment = eCharacterAlignment.FactionRoyal Then
            Call ExpulsarFaccionReal(tUser, True)
            CurrentAlignmentName = "la Legión Oscura"
        ElseIf CurrentAlignment = eCharacterAlignment.FactionLegion Then
            Call ExpulsarFaccionCaos(tUser, True)
            CurrentAlignmentName = "la Armada Real"
        End If
        
        Call WriteConsoleMsg(tUser, UserList(UserIndex).Name & " te ha expulsado en forma definitiva de tu facción.", FontTypeNames.FONTTYPE_FIGHT)
        Call RefreshCharStatus(tUser, True)
    Else
    
        Dim UserId As Long, charName As String, GuildId As Long, GuildIndex As Integer
        
        UserId = GetUserID(UserName)
        
        Call GetCharInformationForFactionKick(UserId, charName, CurrentAlignment, GuildId)
        
        If UserId <= 0 Then
             Call WriteConsoleMsg(UserIndex, "Personaje " & UserName & " inexistente.", FontTypeNames.FONTTYPE_INFO)
             Exit Sub
        End If
        
        If GuildId > 0 Then
            GuildIndex = GuildIndexOf(GuildId)
            If GuildList(UserList(tUser).Guild.GuildIndex).MemberCount > 1 Then
                Call WriteConsoleMsg(UserIndex, "No puedes expulsar al usuario " & UserName & " de su facción porque se encuentra en un clan que posee otros miembros. Hacer esto dejaría al clan en una alineación diferente a la del personaje.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            GuildList(GuildIndex).Alignment = eCharacterAlignment.Neutral
        End If
    
        Call ExpellUserFromFactionDB(UserId, eCharacterAlignment.Neutral, 200, UserList(UserIndex).Id)
    End If
        
    Call WriteConsoleMsg(UserIndex, UserName & " expulsado de su facción y prohibida su reenlistada.", FontTypeNames.FONTTYPE_INFO)
    


    Exit Sub
  
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function KickFromFactionByName de ModFacciones.bas")
End Sub

Public Sub ForgiveCharacter(ByVal UserIndex As Integer, ByRef TargetUserName As String)
On Error GoTo ErrHandler:

    Dim TargetIndex As Integer
    
    If TargetUserName = vbNullString Then Exit Sub
    
    TargetIndex = NameIndex(TargetUserName)
    If TargetIndex <= 0 Then
        Call WriteConsoleMsg(UserIndex, "El usuario se encuentra offline. Solo puedes perdonar personajes que estén presentes.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If UserList(TargetIndex).Faccion.Alignment <> eCharacterAlignment.Neutral Then
        Call WriteConsoleMsg(UserIndex, "Solo puedes perdonar personajes neutrales.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    UserList(TargetIndex).Faccion.Reenlistadas = ConstantesBalance.FactionMaxRejoins - 1
    
    Call WriteConsoleMsg(UserIndex, "Has perdonado al usuario " & TargetUserName, FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(TargetIndex, "Los Dioses han perdonado tus errores faccionarios, pero no olvidan.", FontTypeNames.FONTTYPE_INFOBOLD)
    
    Call LogGM(UserList(UserIndex).Name, "Perdonó a " & TargetUserName & " (" & UserList(TargetIndex).Id & ")")
    
    Call AddPunishmentDB(UserList(TargetIndex).Id, UserList(UserIndex).Id, PUNISHMENT_TYPE_RECORD, "Perdón divino", "")
    
    
    Exit Sub

ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function ForgiveCharacter de ModFacciones.bas")
End Sub
    
