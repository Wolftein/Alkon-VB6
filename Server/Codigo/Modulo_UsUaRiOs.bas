Attribute VB_Name = "UsUaRiOs"
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

'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                        Modulo Usuarios
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Rutinas de los usuarios
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

Public Sub ActStats(ByVal VictimIndex As Integer, ByVal AttackerIndex As Integer)
On Error GoTo ErrHandler
  
    Dim VictimELV As Integer
    
    With UserList(AttackerIndex)
        
        VictimELV = CInt(UserList(VictimIndex).Stats.ELV)
        
        Call CheckUserLevel(AttackerIndex)
                
        'Lo mata
        Call WriteMultiMessage(AttackerIndex, eMessages.HaveKilledUser, VictimIndex, VictimELV)
        Call WriteMultiMessage(VictimIndex, eMessages.UserKill, AttackerIndex)
        
        'Log
        Call LogAsesinato(.Name & " asesino a " & UserList(VictimIndex).Name)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ActStats de Modulo_UsUaRiOs.bas")
End Sub

Public Sub RevivirUsuario(ByVal UserIndex As Integer, Optional ByVal MakeItHungryAndThirsty As Boolean = True)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    With UserList(UserIndex)
        .flags.Muerto = 0
        .Stats.MinHp = .Stats.UserAtributos(eAtributos.Constitucion)
        
        .Stats.MinMAN = 0
        .Stats.MinSta = 0
                
        If MakeItHungryAndThirsty Then
            .Stats.MinAGU = 0
            .flags.Sed = 1
            .Stats.MinHam = 0
            .flags.Hambre = 1
        End If
        
        If .Stats.MinHp > .Stats.MaxHp Then
            .Stats.MinHp = .Stats.MaxHp
        End If
        
        If .flags.Navegando = 1 Then
            Call ToggleBoatBody(UserIndex)
        Else
            Call DarCuerpoDesnudo(UserIndex)
            
            .Char.head = .OrigChar.head
        End If
        
        If .flags.Traveling Then
            Call EndTravel(UserIndex, True)
        End If
        
        Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
        Call WriteUpdateUserStats(UserIndex)
        Call WriteUpdateHungerAndThirst(UserIndex)
        
        If EnMapaDuelos(UserIndex) Then
            If UserList(UserIndex).flags.DueloIndex > 0 Then
                DuelData.Duelo(UserList(UserIndex).flags.DueloIndex).Team(GetUserTeam(UserList(UserIndex).flags.DueloIndex, UserIndex)).Muerto(GetTeamSlot(UserList(UserIndex).flags.DueloIndex, GetUserTeam(UserList(UserIndex).flags.DueloIndex, UserIndex), UserIndex)) = False
            End If
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub RevivirUsuario de Modulo_UsUaRiOs.bas")
End Sub

Public Sub ToggleBoatBody(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 25/07/2010
'Gives boat body depending on user alignment.
'25/07/2010: ZaMa - Now makes difference depending on faccion and atacable status.
'***************************************************
On Error GoTo ErrHandler
  
    Dim NewBody As Integer
    
    With UserList(UserIndex)
 
        If .Invent.BarcoObjIndex = 0 Then Exit Sub
        .Char.head = 0
                        
        Select Case .Faccion.Alignment
            Case eCharacterAlignment.Neutral
                NewBody = ObjData(.Invent.BarcoObjIndex).NumBodyNeutral
            Case eCharacterAlignment.FactionRoyal
                NewBody = ObjData(.Invent.BarcoObjIndex).NumBodyRoyal
            Case eCharacterAlignment.FactionLegion
                NewBody = ObjData(.Invent.BarcoObjIndex).NumBodyLegion
        
        End Select
        
        If NewBody = 0 Then NewBody = ConstantesGRH.Barca
        
        .Char.body = NewBody
        .Char.ShieldAnim = ConstantesGRH.NingunEscudo
        .Char.WeaponAnim = ConstantesGRH.NingunArma
        .Char.CascoAnim = ConstantesGRH.NingunCasco
    End With

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ToggleBoatBody de Modulo_UsUaRiOs.bas")
End Sub

Public Sub ChangeUserChar(ByVal UserIndex As Integer, ByVal body As Integer, ByVal head As Integer, ByVal heading As Byte, _
                    ByVal Arma As Integer, ByVal Escudo As Integer, ByVal casco As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex).Char
        .body = body
        .head = head
        .heading = heading
        .WeaponAnim = Arma
        .ShieldAnim = Escudo
        .CascoAnim = casco
    End With
    
    With UserList(UserIndex)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChange(body, head, heading, .Char.CharIndex, Arma, Escudo, .Char.FX, .Char.Loops, casco, CBool(.flags.Navegando), .flags.Muerto = 1, .OverHeadIcon, .Faccion.Alignment))
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ChangeUserChar de Modulo_UsUaRiOs.bas")
End Sub

Public Function GetWeaponAnim(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Integer
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 03/29/10
'
'***************************************************
On Error GoTo ErrHandler
  
    Dim Tmp As Integer

    With UserList(UserIndex)
        Tmp = ObjData(ObjIndex).WeaponRazaEnanaAnim
            
        If Tmp > 0 Then
            If .raza = eRaza.Enano Or .raza = eRaza.Gnomo Then
                GetWeaponAnim = Tmp
                Exit Function
            End If
        End If
        
        GetWeaponAnim = ObjData(ObjIndex).WeaponAnim
    End With
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetWeaponAnim de Modulo_UsUaRiOs.bas")
End Function


Public Sub EraseUserChar(ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 08/01/2009
'08/01/2009: ZaMa - No se borra el char de un admin invisible en todos los clientes excepto en su mismo cliente.
'*************************************************

On Error GoTo ErrorHandler
    
    With UserList(UserIndex)
        
        If .Char.CharIndex > 0 And .Char.CharIndex <= LastChar Then
            charList(.Char.CharIndex) = 0
            
            If .Char.CharIndex = LastChar Then
                Do Until charList(LastChar) > 0
                    LastChar = LastChar - 1
                    If LastChar <= 1 Then Exit Do
                Loop
            End If
        End If
        
        Call ModAreas.DeleteEntity(UserIndex, ENTITY_TYPE_PLAYER)
        
        If MapaValido(.Pos.Map) Then
            MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex = 0
        End If
        
        .Char.CharIndex = 0
    End With
    
    NumChars = NumChars - 1
Exit Sub
    
ErrorHandler:
    
    Dim UserName As String
    Dim CharIndex As Integer
    
    If UserIndex > 0 Then
        UserName = UserList(UserIndex).Name
        CharIndex = UserList(UserIndex).Char.CharIndex
    End If

    Call LogError("Error en EraseUserchar " & Err.Number & ": " & Err.Description & _
                  ". User: " & UserName & "(UI: " & UserIndex & " - CI: " & CharIndex & ")")
End Sub

Public Sub RefreshCharStatus(ByVal UserIndex As Integer, ByVal CheckGuildAlignment As Boolean)
'*************************************************
'Author: Tararira
'Last modified: 14/03/2011
'Refreshes the status and tag of UserIndex.
'04/07/2009: ZaMa - Ahora mantenes la fragata fantasmal si estas muerto.
'14/03/2011: ZaMa - Now checks guild alignment.
'*************************************************
On Error GoTo ErrHandler
  
    Dim ClanTag As String
    Dim NickColor As Byte
    
    With UserList(UserIndex)
        
        NickColor = GetNickColor(UserIndex)
        If .Guild.IdGuild > 0 Then
            ClanTag = " <" & GuildList(.Guild.GuildIndex).Name & ">"
        End If
        
        If .ShowName Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageUpdateTagAndStatus(UserIndex, NickColor, .secName & ClanTag))
        Else
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageUpdateTagAndStatus(UserIndex, NickColor, vbNullString))
        End If
        
        'Si esta navengando, se cambia la barca.
        If .flags.Navegando Then
            If .flags.Muerto = 1 Then
                .Char.body = ConstantesGRH.FragataFantasmal
            Else
                Call ToggleBoatBody(UserIndex)
            End If
        End If
        Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
        
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub RefreshCharStatus de Modulo_UsUaRiOs.bas")
End Sub

Public Function GetNickColor(ByVal UserIndex As Integer) As Byte
'*************************************************
'Author: ZaMa
'Last modified: 15/01/2010
'
'*************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)
        
        Select Case .Faccion.Alignment
            Case eCharacterAlignment.Neutral
                GetNickColor = eNickColor.ieNeutral
            Case eCharacterAlignment.FactionRoyal
                GetNickColor = eNickColor.ieCiudadano
            Case eCharacterAlignment.FactionLegion
                GetNickColor = eNickColor.ieCriminal
        End Select
    End With
    
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetNickColor de Modulo_UsUaRiOs.bas")
End Function

Public Function NewUserChar(ByVal UserIndex As Integer) As Boolean
On Error GoTo ErrHandler
    
    Dim Dirty As Boolean
    
    With UserList(UserIndex)
    
        If (Not InMapBounds(.Pos.Map, .Pos.X, .Pos.Y)) Then
            Exit Function
        End If
                        
        ' We only need to send it once!
        If .Char.CharIndex = 0 Then
            .Char.CharIndex = NextOpenCharIndex
            charList(.Char.CharIndex) = UserIndex
            
            Dirty = True
        End If
                
        'Place character on map if needed
        MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex = UserIndex
              
        ' Update character in self
        Call MakeUserChar(False, UserIndex, UserIndex)
        
        ' Create it on the grid
        Call ModAreas.CreateEntity(UserIndex, ENTITY_TYPE_PLAYER, UserList(UserIndex).Pos, ModAreas.DEFAULT_PLAYER_WIDTH, ModAreas.DEFAULT_PLAYER_HEIGHT)
        
        If (Dirty) Then
            Call WriteUserCharIndexInServer(UserIndex)
        End If
        
    End With
    
    NewUserChar = True
    Exit Function
    
ErrHandler:

    Call CloseSocket(UserIndex)
End Function

Public Sub MakeUserChar(ByVal ToArea As Boolean, ByVal UserIndex As Integer, ByVal Observer As Integer)
On Error GoTo ErrHandler

    Dim ClanTag As String
    Dim NickColor As Byte
    Dim UserName As String
    Dim Privileges As Byte
    
    With UserList(UserIndex)
        If .Guild.IdGuild > 0 Then
            ClanTag = GuildList(.Guild.GuildIndex).Name
        End If
                
        NickColor = GetNickColor(UserIndex)
        Privileges = .flags.Privilegios
                
        'Preparo el nick
        If .ShowName Then
            UserName = .secName
                    
            If .flags.HelpMode Then
                UserName = UserName & " " & TAG_CONSULT_MODE
            Else
                If LenB(ClanTag) <> 0 Then _
                    UserName = UserName & " <" & ClanTag & ">"
                End If
            End If
            
            Dim Target As SendTarget
            Target = IIf(ToArea, SendTarget.ToPCAreaButIndex, SendTarget.ToUser)

            Call SendData(Target, Observer, _
                PrepareMessageCharacterCreate(.Char.body, .Char.head, .Char.heading, _
                    .Char.CharIndex, .Pos.X, .Pos.Y, .Char.WeaponAnim, .Char.ShieldAnim, _
                    .Char.FX, INFINITE_LOOPS, .Char.CascoAnim, UserName, NickColor, .Faccion.Alignment, Privileges, _
                    False, False, CBool(.flags.Navegando), OverHeadIcon:=.OverHeadIcon))
    End With

    Exit Sub

ErrHandler:

    Dim UserMap As Integer

    If UserIndex > 0 Then
        UserName = UserList(UserIndex).Name
        UserMap = UserList(UserIndex).Pos.Map
    End If

    Call LogError("Error en la subrutina MakeUserChar - Error : " & Err.Number & _
        " - Description : " & Err.Description & ". User: " & UserName & "(" & UserIndex & "). Map: " & UserMap)
End Sub

''
' Checks if the user gets the next level.
'
' @param UserIndex Specifies reference to user

Public Sub CheckUserLevel(ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 14/04/2015
'Chequea que el usuario no halla alcanzado el siguiente nivel,
'de lo contrario le da la vida, mana, etc, correspodiente.
'07/08/2006 Integer - Modificacion de los valores
'01/10/2007 Tavo - Corregido el BUG de STAT_MAXELV
'24/01/2007 Pablo (ToxicWaste) - Agrego modificaciones en ELU al subir de nivel.
'24/01/2007 Pablo (ToxicWaste) - Agrego modificaciones de la subida de mana de los magos por lvl.
'13/03/2007 Pablo (ToxicWaste) - Agrego diferencias entre el 18 y el 19 en Constitución.
'09/01/2008 Pablo (ToxicWaste) - Ahora el incremento de vida por Consitución se controla desde Balance.dat
'12/09/2008 Marco Vanotti (Marco) - Ahora si se llega a nivel 25 y está en un clan, se lo expulsa para no sumar antifacción
'02/03/2009 ZaMa - Arreglada la validacion de expulsion para miembros de clanes faccionarios que llegan a 25.
'11/19/2009 Pato - Modifico la nueva fórmula de maná ganada para el bandido y se la limito a 499
'02/04/2010: ZaMa - Modifico la ganancia de hit por nivel del ladron.
'08/04/2011: Amraphen - Arreglada la distribución de probabilidades para la vida en el caso de promedio entero.
'07/08/2014: D'Artagnan - Level up animation.
'14/04/2015: D'Artagnan - Free skills for each class.
'21/02/2016: Nightw - Added the missing funcionallity for the mastery points after reaching lvl max.
'*************************************************
    Dim Pts As Integer
    Dim NewHp As Integer
    Dim NewStamina As Integer
    Dim NewMana As Integer
    
    Dim NewMinHit As Integer
    Dim NewMaxHit As Integer
    Dim WasNewbie As Boolean
    Dim Promedio As Double
    Dim aux As Integer
    Dim DistVida(1 To 5) As Integer
    Dim GI As Integer 'Guild Index
    Dim SubioNivel As Boolean
    
On Error GoTo ErrHandler
    
    WasNewbie = EsNewbie(UserIndex)
    
    With UserList(UserIndex)
    
        If .Stats.Exp > ConstantesBalance.MaxExp Then .Stats.Exp = ConstantesBalance.MaxExp
        
        Do While .Stats.Exp >= .Stats.ELU And .Stats.Exp > 0 And .Stats.ELU > 0
            'Checkea si alcanzó el máximo nivel
            If .Stats.ELV > ConstantesBalance.MaxLvl Then
                .Stats.Exp = 0
                .Stats.ELU = TablaExperiencia(ConstantesBalance.MaxLvl)
                .Stats.ELV = ConstantesBalance.MaxLvl
                Exit Sub
            End If
            
            SubioNivel = True

            ' Level up animation.
            If Not (.flags.invisible = 1) And Not (.flags.AdminInvisible = 1) And _
               Not (.flags.Oculto = 1) Then
                Call SendData(SendTarget.ToPCArea, UserIndex, _
                              PrepareMessageCreateFX(.Char.CharIndex, 43, 0))
            End If
            
            ' If the user reached the max level, then compute the mastery points
            If .Stats.ELV = ConstantesBalance.MaxLvl Then
                ' Insert Mastery shit here
                .Stats.Exp = .Stats.Exp - .Stats.ELU
                .Stats.ELU = TablaExperiencia(ConstantesBalance.MaxLvl)
                .Stats.MasteryPoints = .Stats.MasteryPoints + 1
                    
                Call WriteConsoleMsg(UserIndex, "Has ganado un punto de maestría. Tienes " & .Stats.MasteryPoints & IIf(.Stats.MasteryPoints > 0, " puntos disponibles.", " punto disponible"), FontTypeNames.FONTTYPE_INFO)
                
                ' Log the mastery raise.
                Call Statistics.UserLevelUp(UserIndex, True)
                                
            Else
                .Stats.ELV = .Stats.ELV + 1
                .Stats.Exp = .Stats.Exp - .Stats.ELU
                
                            
                Pts = Pts + Classes(.clase).ClassMods.SkillsPerLevel
                
                'Store it!
                Call Statistics.UserLevelUp(UserIndex)
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(ConstantesSonidos.Nivel, .Pos.X, .Pos.Y, .Char.CharIndex))
                Call WriteConsoleMsg(UserIndex, "¡Has subido de nivel!", FontTypeNames.FONTTYPE_INFO)
                
                .Stats.ELU = TablaExperiencia(.Stats.ELV)
  
                Dim HpGained As Integer, MpGained As Integer
                HpGained = .Stats.MaxHp
                MpGained = .Stats.MaxMan
                
                ' Recalculate the Max HP
                NewHp = RecalculateCharacterMaxHealth(UserIndex)
                HpGained = NewHp - .Stats.MaxHp
                .Stats.MaxHp = NewHp
                .Stats.MinHp = NewHp
                
                'Actualizamos Mana
                NewMana = RecalculateCharacterMaxMana(UserIndex)
                MpGained = NewMana - .Stats.MaxMan
                .Stats.MaxMan = NewMana
                .Stats.MinMAN = NewMana
                
                'Actualizamos Stamina
                NewStamina = Classes(.clase).ClassMods.StaminaPerLevel
                .Stats.MaxSta = .Stats.MaxSta + NewStamina
                If .Stats.MaxSta > ConstantesBalance.MaxSta Then .Stats.MaxSta = ConstantesBalance.MaxSta
                                              
                'Notificamos al user
                If HpGained > 0 Then Call WriteConsoleMsg(UserIndex, "Has ganado " & HpGained & " puntos de vida.", FontTypeNames.FONTTYPE_INFO)
                
                If NewStamina > 0 Then Call WriteConsoleMsg(UserIndex, "Has ganado " & NewStamina & " puntos de energía.", FontTypeNames.FONTTYPE_INFO)

                If MpGained > 0 Then Call WriteConsoleMsg(UserIndex, "Has ganado " & MpGained & " puntos de maná.", FontTypeNames.FONTTYPE_INFO)

                Call LogDesarrollo(.Name & " paso a nivel " & .Stats.ELV & " gano HP: " & NewHp)
    
                'If user is in a party, we modify the variable p_sumaniveleselevados
                Call mdParty.ActualizarSumaNivelesElevados(UserIndex)
            End If
        Loop
        
        If SubioNivel Then
            If Not EsNewbie(UserIndex) And WasNewbie Then
                ' Change the alignment
                .Faccion.Alignment = eCharacterAlignment.Neutral
                
                ' Remove the newbie items and move the character to another map if needed.
                Call QuitarNewbieObj(UserIndex)
                If MapInfo(.Pos.Map).Restringir = eRestrict.restrict_newbie Then
                    Call WarpUserChar(UserIndex, 1, 50, 50, True)
                    Call WriteConsoleMsg(UserIndex, "Debes abandonar el Dungeon Newbie.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
            
            'Send all gained skill points at once (if any)
            If Pts > 0 Then
                .Stats.SkillPts = .Stats.SkillPts + Pts
                Call WriteLevelUp(UserIndex, .Stats.SkillPts)
                Call WriteConsoleMsg(UserIndex, "Has ganado un total de " & Pts & " skillpoints.", FontTypeNames.FONTTYPE_INFO)
            End If
            
            ' Recalculate the passive habilities to see if there's a new passive that should be enabled
            Call RecalculateUserPassives(UserIndex, True)
            
            Call WriteUpdateUserStats(UserIndex)
            
        Else
            Call WriteUpdateExp(UserIndex)
        End If

    End With
    
    Exit Sub

ErrHandler:
    Dim UserName As String
    Dim UserMap As Integer

    If UserIndex > 0 Then
        UserName = UserList(UserIndex).Name
        UserMap = UserList(UserIndex).Pos.Map
    End If

    Call LogError("Error en la subrutina CheckUserLevel - Error : " & Err.Number & _
        " - Description : " & Err.Description & ". User: " & UserName & "(" & UserIndex & "). Map: " & UserMap)
End Sub

Public Function PuedeAtravesarAgua(ByVal UserIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    PuedeAtravesarAgua = UserList(UserIndex).flags.Navegando = 1 _
                    Or UserList(UserIndex).flags.Vuela = 1
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function PuedeAtravesarAgua de Modulo_UsUaRiOs.bas")
End Function

Function MoveUserChar(ByVal UserIndex As Integer, ByVal nHeading As eHeading) As Boolean
'*************************************************
'Author: Unknown
'Last modified: 13/07/2009
'Moves the char, sending the message to everyone in range.
'30/03/2009: ZaMa - Now it's legal to move where a casper is, changing its pos to where the moving char was.
'28/05/2009: ZaMa - When you are moved out of an Arena, the resurrection safe is activated.
'13/07/2009: ZaMa - Now all the clients don't know when an invisible admin moves, they force the admin to move.
'13/07/2009: ZaMa - Invisible admins aren't allowed to force dead characater to move
'*************************************************
On Error GoTo ErrHandler
  
    Dim nPos As WorldPos
    Dim sailing As Boolean
    Dim CasperIndex As Integer
    Dim CasperHeading As eHeading
    Dim isAdminInvi As Boolean
    Dim isZonaOscura As Boolean
    Dim isZonaOscuraNewPos As Boolean
    Dim UserMoved As Boolean
    
    sailing = PuedeAtravesarAgua(UserIndex)
    nPos = UserList(UserIndex).Pos
    isZonaOscura = (MapData(nPos.Map, nPos.X, nPos.Y).Trigger = eTrigger.zonaOscura)
    
    Call HeadtoPos(nHeading, nPos)
    
    isZonaOscuraNewPos = (MapData(nPos.Map, nPos.X, nPos.Y).Trigger = eTrigger.zonaOscura)
    
    isAdminInvi = (UserList(UserIndex).flags.AdminInvisible = 1)
    
    UserList(UserIndex).Char.heading = nHeading
    
    If MoveToLegalPos(UserList(UserIndex).Pos.Map, nPos.X, nPos.Y, sailing, Not sailing) Then
        UserMoved = True
    
        'si no estoy solo en el mapa...
        If MapInfo(UserList(UserIndex).Pos.Map).NumUsers > 1 Then
            CasperIndex = MapData(UserList(UserIndex).Pos.Map, nPos.X, nPos.Y).UserIndex
                        
            'Si hay un usuario, y paso la validacion, entonces es un casper
            If CasperIndex > 0 Then
                ' Los admins invisibles no pueden patear caspers
                If Not isAdminInvi Then
                    
                    If TriggerZonaPelea(UserIndex, CasperIndex) = TRIGGER6_PROHIBE Then
                        If UserList(CasperIndex).flags.SeguroResu = False Then
                            UserList(CasperIndex).flags.SeguroResu = True
                            Call WriteMultiMessage(CasperIndex, eMessages.ResuscitationSafeOn)
                        End If
                    End If
    
                    With UserList(CasperIndex)
                        CasperHeading = InvertHeading(nHeading)
                        Call HeadtoPos(CasperHeading, .Pos)
                    
                        ' Si es un admin invisible, no se avisa a los demas clientes
                        If Not (.flags.AdminInvisible = 1) Then
                             'Los valores de visible o invisible están invertidos porque estos flags son del UserIndex, por lo tanto si el UserIndex entra, el casper sale y viceversa :P
                            If isZonaOscura Then
                                If Not isZonaOscuraNewPos Then
                                    Call SendData(SendTarget.ToUsersAndRmsAndCounselorsAreaButGMs, CasperIndex, PrepareMessageSetInvisible(.Char.CharIndex, True))
                                End If
                            Else
                                If isZonaOscuraNewPos Then
                                    Call SendData(SendTarget.ToUsersAndRmsAndCounselorsAreaButGMs, CasperIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
                                End If
                            End If
                        End If
                             
                        'Update map and char
                        .Char.heading = CasperHeading
                        MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex = CasperIndex
                        
                        Call WriteForceCharMove(CasperIndex, CasperHeading)
                   
                        'Actualizamos las áreas de ser necesario
                        Call ModAreas.UpdateEntity(CasperIndex, ENTITY_TYPE_PLAYER, .Pos, False)
                    End With
                End If
            End If
        End If
        
        ' Los admins invisibles no pueden patear caspers
        If (Not isAdminInvi) Or (CasperIndex = 0) Then
            With UserList(UserIndex)
                ' Si no hay intercambio de pos con nadie
                If CasperIndex = 0 Then
                    MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex = 0
                End If
                
                .Pos = nPos
                MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex = UserIndex
                
                If isZonaOscura Then
                    If Not isZonaOscuraNewPos Then
                        If (.flags.invisible Or .flags.Oculto) = 0 Then
                            Call SendData(SendTarget.ToUsersAndRmsAndCounselorsAreaButGMs, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
                        End If
                    End If
                Else
                    If isZonaOscuraNewPos Then
                        If (.flags.invisible Or .flags.Oculto) = 0 Then
                            Call SendData(SendTarget.ToUsersAndRmsAndCounselorsAreaButGMs, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, True))
                        End If
                    End If
                End If
                Dim intObj As Integer
                intObj = MapData(.Pos.Map, .Pos.X, .Pos.Y).ObjInfo.ObjIndex
                
                
                If intObj = ConstantesItems.FogataElfica Then
                    ' If the user is dead, revive him!
                    If .flags.Muerto Then
                        Call RevivirUsuario(UserIndex, True)
                        .Stats.MinHp = Int(.Stats.MaxHp / 2)
                        Call WriteUpdateHP(UserIndex)
                        
                        ' Send a message to him
                        Call WriteConsoleMsg(UserIndex, "Los poderes curativos de la fogata élfica te han devuelto a la vida.", FontTypeNames.FONTTYPE_INFOBOLD)
                    End If
                    
                    ' We should remove the campfire after the user steps on it.
                    Call EraseObj(intObj, .Pos.Map, .Pos.X, .Pos.Y)
                End If
                
                'Actualizamos las áreas de ser necesario
                Call ModAreas.UpdateEntity(UserIndex, ENTITY_TYPE_PLAYER, .Pos, False)
                
                Call DoTileEvents(UserIndex, .Pos.Map, .Pos.X, .Pos.Y)
            End With
        Else
            Call WritePosUpdate(UserIndex)
        End If
    Else
        Call WritePosUpdate(UserIndex)
    End If
        
    MoveUserChar = UserMoved
    
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function MoveUserChar de Modulo_UsUaRiOs.bas")
End Function

Public Function InvertHeading(ByVal nHeading As eHeading) As eHeading
'*************************************************
'Author: ZaMa
'Last modified: 30/03/2009
'Returns the heading opposite to the one passed by val.
'*************************************************
On Error GoTo ErrHandler
  
    Select Case nHeading
        Case eHeading.EAST
            InvertHeading = WEST
        Case eHeading.WEST
            InvertHeading = EAST
        Case eHeading.SOUTH
            InvertHeading = NORTH
        Case eHeading.NORTH
            InvertHeading = SOUTH
    End Select
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function InvertHeading de Modulo_UsUaRiOs.bas")
End Function
Function NextOpenUser() As Integer
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    Dim LoopC As Long
    
    For LoopC = 1 To MaxUsers + 1
        If LoopC > MaxUsers Then Exit For
        If (UserList(LoopC).ConnIDValida = False And UserList(LoopC).flags.UserLogged = False) Then Exit For
    Next LoopC
    
    NextOpenUser = LoopC
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function NextOpenUser de Modulo_UsUaRiOs.bas")
End Function

Function NextOpenCharIndex() As Integer
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    Dim LoopC As Long
    
    For LoopC = 1 To MAXCHARS
        If charList(LoopC) = 0 Then
            NextOpenCharIndex = LoopC
            NumChars = NumChars + 1
            
            If LoopC > LastChar Then _
                LastChar = LoopC
            
            Exit Function
        End If
    Next LoopC
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function NextOpenCharIndex de Modulo_UsUaRiOs.bas")
End Function

Public Sub FreeSlot(ByVal UserIndex As Integer)
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 01/10/2012
'
'***************************************************
On Error GoTo ErrHandler

If UserIndex = LastUser Then
    Do While (LastUser > 0)
        If (UserList(LastUser).flags.UserLogged) Then Exit Do
        LastUser = LastUser - 1
    Loop
End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub FreeSlot de Modulo_UsUaRiOs.bas")
End Sub

Public Sub SendUserStatsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 26/05/2011 (Amraphen)
'26/05/2011: Amraphen - Ahora envía la defensa adicional de la armadura de segunda jerarquía
'***************************************************
On Error GoTo ErrHandler
  

    Dim GuildI As Integer
    Dim ModificadorDefensa As Single 'Por las armaduras de segunda jerarquía.
    
    With UserList(UserIndex)
        Call WriteConsoleMsg(sendIndex, "Estadísticas de: " & .Name, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Nivel: " & .Stats.ELV & "  EXP: " & .Stats.Exp & "/" & .Stats.ELU, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Salud: " & .Stats.MinHp & "/" & .Stats.MaxHp & "  Maná: " & .Stats.MinMAN & "/" & .Stats.MaxMan & "  Energía: " & .Stats.MinSta & "/" & .Stats.MaxSta, FontTypeNames.FONTTYPE_INFO)
                
        If .Invent.ArmourEqpObjIndex > 0 Then
            If .Invent.FactionArmourEqpObjIndex > 0 Then ModificadorDefensa = ConstantesBalance.ModDefSegJerarquia Else ModificadorDefensa = 1
            If .Invent.EscudoEqpObjIndex > 0 Then
                Call WriteConsoleMsg(sendIndex, "(CUERPO) Mín Def/Máx Def: " & CInt(ObjData(.Invent.ArmourEqpObjIndex).MinDef * ModificadorDefensa) + ObjData(.Invent.EscudoEqpObjIndex).MinDef & "/" & CInt(ObjData(.Invent.ArmourEqpObjIndex).MaxDef * ModificadorDefensa) + ObjData(.Invent.EscudoEqpObjIndex).MaxDef, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(sendIndex, "(CUERPO) Mín Def/Máx Def: " & CInt(ObjData(.Invent.ArmourEqpObjIndex).MinDef * ModificadorDefensa) & "/" & CInt(ObjData(.Invent.ArmourEqpObjIndex).MaxDef * ModificadorDefensa), FontTypeNames.FONTTYPE_INFO)
            End If
        Else
            Call WriteConsoleMsg(sendIndex, "(CUERPO) Mín Def/Máx Def: 0", FontTypeNames.FONTTYPE_INFO)
        End If
        
        If .Invent.CascoEqpObjIndex > 0 Then
            Call WriteConsoleMsg(sendIndex, "(CABEZA) Mín Def/Máx Def: " & ObjData(.Invent.CascoEqpObjIndex).MinDef & "/" & ObjData(.Invent.CascoEqpObjIndex).MaxDef, FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(sendIndex, "(CABEZA) Mín Def/Máx Def: 0", FontTypeNames.FONTTYPE_INFO)
        End If
        
        'GuildI = .IdGuild
        'If GuildI > 0 Then
        '    Call WriteConsoleMsg(sendIndex, "Clan: " & modGuilds.GuildSpecialName(GuildI), FontTypeNames.FONTTYPE_INFO)
        '    If modGuilds.IsLeaderName(.Name, GuildI) Then
        '        Call WriteConsoleMsg(sendIndex, "Status: Líder", FontTypeNames.FONTTYPE_INFO)
        '    End If
        '    'guildpts no tienen objeto
        'End If
        
#If ConUpTime Then
        Dim TempDate As Date
        Dim TempSecs As Long
        Dim TempStr As String
        TempDate = Now - .LogOnTime
        TempSecs = (.UpTime + (Abs(Day(TempDate) - 30) * 24 * 3600) + (Hour(TempDate) * 3600) + (Minute(TempDate) * 60) + Second(TempDate))
        TempStr = (TempSecs \ 86400) & " Dias, " & ((TempSecs Mod 86400) \ 3600) & " Horas, " & ((TempSecs Mod 86400) Mod 3600) \ 60 & " Minutos, " & (((TempSecs Mod 86400) Mod 3600) Mod 60) & " Segundos."
        Call WriteConsoleMsg(sendIndex, "Logeado hace: " & Hour(TempDate) & ":" & Minute(TempDate) & ":" & Second(TempDate), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Total: " & TempStr, FontTypeNames.FONTTYPE_INFO)
#End If
        If .flags.Traveling = 1 Then
            Call WriteConsoleMsg(sendIndex, "Tiempo restante para llegar a tu hogar: " & GetHomeArrivalTime(UserIndex) & " segundos.", FontTypeNames.FONTTYPE_INFO)
        End If
        
        Call WriteConsoleMsg(sendIndex, "Oro: " & .Stats.GLD & "  Posición: " & .Pos.X & "," & .Pos.Y & " en mapa " & .Pos.Map, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Dados: " & .Stats.UserAtributos(eAtributos.Fuerza) & ", " & .Stats.UserAtributos(eAtributos.Agilidad) & ", " & .Stats.UserAtributos(eAtributos.Inteligencia) & ", " & .Stats.UserAtributos(eAtributos.Carisma) & ", " & .Stats.UserAtributos(eAtributos.Constitucion), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Hogar: " & ListaCiudades(.Hogar), FontTypeNames.FONTTYPE_INFO)
        
        If .Stats.ELV = ConstantesBalance.MaxLvl Then
            Call WriteConsoleMsg(sendIndex, "Puntos de maestría: " & .Stats.MasteryPoints, FontTypeNames.FONTTYPE_INFO)
        End If
        
        ' Ranking points of the character
        If .Stats.ELV >= ConstantesBalance.RankingMinLevel Then
            Call WriteConsoleMsg(sendIndex, "Puntos de ranking: " & .Stats.RankingPoints, FontTypeNames.FONTTYPE_INFO)
        End If
        
        ' Ranking points of the Guild
        If .Guild.IdGuild > 0 Then
            Call WriteConsoleMsg(sendIndex, "Puntos de ranking de clan: " & GuildList(.Guild.GuildIndex).RankingPoints, FontTypeNames.FONTTYPE_INFO)
        End If
        
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SendUserStatsTxt de Modulo_UsUaRiOs.bas")
End Sub

Sub SendUserMiniStatsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 23/01/2007
'Shows the users Stats when the user is online.
'23/01/2007 Pablo (ToxicWaste) - Agrego de funciones y mejora de distribución de parámetros.
'*************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)
        Call WriteConsoleMsg(sendIndex, "Pj: " & .Name, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Neutrales matados: " & .Faccion.CiudadanosMatados & "| Armada matados: " & .Faccion.CiudadanosMatados & " | Caos matados: " & .Faccion.CriminalesMatados & " usuarios matados: " & .Stats.UsuariosMatados, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "NPCs matados: " & .Stats.NPCsMuertos, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Clase: " & ListaClases(.clase), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Raza: " & ListaRazas(.raza), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Pena: " & .Counters.Pena, FontTypeNames.FONTTYPE_INFO)
        
        If .Faccion.ArmadaReal = 1 Then
            Call WriteConsoleMsg(sendIndex, "Ejército real desde: " & Format$(.Faccion.FechaIngreso, "yyyy-mm-dd h:mm:ss"), FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Ingresó en nivel: " & .Faccion.NivelIngreso & " con " & .Faccion.MatadosIngreso & " ciudadanos matados.", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Veces que ingresó: " & .Faccion.Reenlistadas, FontTypeNames.FONTTYPE_INFO)
        
        ElseIf .Faccion.FuerzasCaos = 1 Then
            Call WriteConsoleMsg(sendIndex, "Legión oscura desde: " & Format$(.Faccion.FechaIngreso, "yyyy-mm-dd h:mm:ss"), FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Ingresó en nivel: " & .Faccion.NivelIngreso, FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Veces que ingresó: " & .Faccion.Reenlistadas, FontTypeNames.FONTTYPE_INFO)
        
        ElseIf .Faccion.RecibioExpInicialReal = 1 Then
            Call WriteConsoleMsg(sendIndex, "Fue ejército real", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Veces que ingresó: " & .Faccion.Reenlistadas, FontTypeNames.FONTTYPE_INFO)
        
        ElseIf .Faccion.RecibioExpInicialCaos = 1 Then
            Call WriteConsoleMsg(sendIndex, "Fue legión oscura", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Veces que ingresó: " & .Faccion.Reenlistadas, FontTypeNames.FONTTYPE_INFO)
        End If
        
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SendUserMiniStatsTxt de Modulo_UsUaRiOs.bas")
End Sub

Sub SendUserInvTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    Dim J As Long
    
    With UserList(UserIndex)
        Call WriteConsoleMsg(sendIndex, "Items en inventario de " & .Name, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Tiene " & .Invent.NroItems & " objetos.", FontTypeNames.FONTTYPE_INFO)
                
        Dim sMensaje As String
        For J = 1 To .CurrentInventorySlots
            If .Invent.Object(J).ObjIndex > 0 Then
                
                sMensaje = "Objeto " & J & " " & ObjData(.Invent.Object(J).ObjIndex).Name & " Cantidad:" & .Invent.Object(J).Amount
                If .Invent.Object(J).Equipped = 1 Then
                    sMensaje = sMensaje & " (E)"
                End If
                
                Call WriteConsoleMsg(sendIndex, sMensaje, FontTypeNames.FONTTYPE_INFO)
            End If
        Next J
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SendUserInvTxt de Modulo_UsUaRiOs.bas")
End Sub

Sub SendUserInvTxtFromChar(ByVal sendIndex As Integer, ByVal charName As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    Dim J As Long
    Dim Tmp As String
    Dim ObjInd As Long
    Dim ObjCant As Long
    
    Dim UserID As Long
    UserID = GetUserID(charName)
    
    If UserID <> 0 Then
        Call SendUserInvTxtFromDB(sendIndex, UserID, charName)
    Else
        Call WriteConsoleMsg(sendIndex, "Usuario inexistente: " & charName, FontTypeNames.FONTTYPE_INFO)
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SendUserInvTxtFromChar de Modulo_UsUaRiOs.bas")
End Sub

Sub SendUserSkillsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

On Error Resume Next
    Dim J As Integer
    
    Call WriteConsoleMsg(sendIndex, UserList(UserIndex).Name, FontTypeNames.FONTTYPE_INFO)
    
    For J = 1 To NUMSKILLS
        Call WriteConsoleMsg(sendIndex, SkillsNames(J) & " = " & GetSkills(UserIndex, J), _
                             FontTypeNames.FONTTYPE_INFO)
    Next J
    
    Call WriteConsoleMsg(sendIndex, "SkillLibres:" & UserList(UserIndex).Stats.SkillPts, FontTypeNames.FONTTYPE_INFO)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SendUserSkillsTxt de Modulo_UsUaRiOs.bas")
End Sub


Public Sub ExitSecureCommerce(ByVal nUserIndex As Integer)
'**********************************************
'Author: D'Artagnan (built from repeated code around the project)
'Last Modification: 11/12/2014
'Exit and reset secure commerce flags.
'**********************************************
On Error GoTo ErrHandler
  
    Dim tUser As Integer
    
    With UserList(nUserIndex)
        tUser = getTradingUser(nUserIndex)
        
        ' Already trading.
        If tUser > 0 Then
            ' Received trade petition.
            If .flags.TargetUser > 0 Then
                Call WriteConsoleMsg(.flags.TargetUser, .Name & " ha cancelado la operación.", FontTypeNames.FONTTYPE_TALK)
                Call FinComerciarUsu(.flags.TargetUser, True)
            End If
            
            If UserList(tUser).flags.UserLogged Then
                If getTradingUser(tUser) = nUserIndex Then
                    Call WriteConsoleMsg(tUser, "Comercio cancelado por el otro usuario.", FontTypeNames.FONTTYPE_TALK)
                    Call FinComerciarUsu(tUser)
                End If
            End If
        ' Not accepted yet.
        ElseIf .flags.nCommerceSourceUser > 0 Then
            Call WriteConsoleMsg(.flags.nCommerceSourceUser, .Name & " ha cancelado la operación.", _
                                 FontTypeNames.FONTTYPE_TALK)
            Call FinComerciarUsu(.flags.nCommerceSourceUser)
        End If
        
        Call FinComerciarUsu(nUserIndex)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ExitSecureCommerce de Modulo_UsUaRiOs.bas")
End Sub

Sub NPCAtacado(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
'**********************************************
'Author: Unknown
'Last Modification: 02/04/2010
'24/01/2007 -> Pablo (ToxicWaste): Agrego para que se actualize el tag si corresponde.
'24/07/2007 -> Pablo (ToxicWaste): Guardar primero que ataca NPC y el que atacas ahora.
'06/28/2008 -> NicoNZ: Los elementales al atacarlos por su amo no se paran más al lado de él sin hacer nada.
'02/04/2010: ZaMa: Un ciuda no se vuelve mas criminal al atacar un npc no hostil.
'**********************************************
On Error GoTo ErrHandler
  
    Dim EraCriminal As Boolean
    
    'Las mascotas no atacan a sus owners
    If Npclist(NpcIndex).MaestroUser = UserIndex Then Exit Sub
    
    'Guardamos el usuario que ataco el npc.
    Npclist(NpcIndex).flags.AttackedBy = UserList(UserIndex).Name
    
    'Npc que estabas atacando.
    Dim LastNpcHit As Integer
    LastNpcHit = UserList(UserIndex).flags.NPCAtacado
    'Guarda el NPC que estas atacando ahora.
    UserList(UserIndex).flags.NPCAtacado = NpcIndex
    
    'Revisamos robo de npc.
    'Guarda el primer nick que lo ataca.
    If Npclist(NpcIndex).flags.AttackedFirstBy = vbNullString Then
        'El que le pegabas antes ya no es tuyo
        If LastNpcHit <> 0 Then
            If Npclist(LastNpcHit).flags.AttackedFirstBy = UserList(UserIndex).Name Then
                Npclist(LastNpcHit).flags.AttackedFirstBy = vbNullString
            End If
        End If
        Npclist(NpcIndex).flags.AttackedFirstBy = UserList(UserIndex).Name
    ElseIf Npclist(NpcIndex).flags.AttackedFirstBy <> UserList(UserIndex).Name Then
        'Estas robando NPC
        'El que le pegabas antes ya no es tuyo
        If LastNpcHit <> 0 Then
            If Npclist(LastNpcHit).flags.AttackedFirstBy = UserList(UserIndex).Name Then
                Npclist(LastNpcHit).flags.AttackedFirstBy = vbNullString
            End If
        End If
    End If
    
    ' If it's a boss NPC, it won't get into defense mode, as they have their own AI.
    If Npclist(NpcIndex).flags.Boss = 1 Then
        Exit Sub
    End If
    
    ' Admins are not attacked back by default.
    If UserList(UserIndex).flags.AdminPerseguible = False Then
        Exit Sub
    End If
        
    ' If it's a pet, send a message to the owner to let them know
    If Npclist(NpcIndex).MaestroUser > 0 And Npclist(NpcIndex).MaestroUser <> UserIndex Then
        If Npclist(NpcIndex).MaestroUser <> UserIndex Then
            Call AllMascotasAtacanUser(UserIndex, Npclist(NpcIndex).MaestroUser)
            Call WriteConsoleMsg(Npclist(NpcIndex).MaestroUser, "¡¡" & UserList(UserIndex).Name & " esta atacando tu mascota!!", FontTypeNames.FONTTYPE_INFO)
        End If
    End If
    
    ' Bosses has their own AI and should not be replaced.
    If Npclist(NpcIndex).flags.Boss > 0 Then Exit Sub
    
    
    Npclist(NpcIndex).Movement = TipoAI.NPCDEFENSA
    Npclist(NpcIndex).Hostile = 1
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub NPCAtacado de Modulo_UsUaRiOs.bas")
End Sub

Public Function PuedeApuñalar(ByVal UserIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 03/08/2012
'03/08/2012: ZaMa - Ya no se necesitan 10 skills iniciales
'***************************************************
On Error GoTo ErrHandler
  
    
    Dim WeaponIndex As Integer
    WeaponIndex = UserList(UserIndex).Invent.WeaponEqpObjIndex
        
    If WeaponIndex > 0 Then
        PuedeApuñalar = (ObjData(WeaponIndex).Apuñala = 1)
    End If
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function PuedeApuñalar de Modulo_UsUaRiOs.bas")
End Function

Public Function PuedeAcuchillar(ByVal UserIndex As Integer) As Boolean
'***************************************************
'Author: ZaMa
'Last Modification: 25/01/2010 (ZaMa)
'
'***************************************************
On Error GoTo ErrHandler
  
    
    Dim WeaponIndex As Integer
    
    With UserList(UserIndex)
        If .clase = eClass.Thief Then
        
            WeaponIndex = .Invent.WeaponEqpObjIndex
            If WeaponIndex > 0 Then
                PuedeAcuchillar = (ObjData(WeaponIndex).Acuchilla = 1)
            End If
        End If
    End With
    
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function PuedeAcuchillar de Modulo_UsUaRiOs.bas")
End Function

Sub SubirSkill(ByVal UserIndex As Integer, ByVal skill As Integer, ByVal Acerto As Boolean, Optional ByVal Exp As Integer = 0)
'*************************************************
'Author: Unknown
'Last modified: 26/04/2015
'11/19/2009: Pato - Implement the new system to train the skills.
'26/04/2015: D'Artagnan - Experience bonus removed.
'*************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)
        If .flags.Hambre = 0 And .flags.Sed = 0 Then
            With .Stats
                If GetSkills(UserIndex, skill) >= MAX_SKILL_POINTS Then Exit Sub
                
                If Not NaturalSkillsAvailable(UserIndex, skill) Then Exit Sub
                
                ' Added new profession skill exp constants. TODO: Rework sub to avoid mixing up with regular skills
                If Acerto Then
                    If Exp = 0 Then Exp = ConstantesBalance.ExpAciertoSkill
                    .ExpSkills(skill) = .ExpSkills(skill) + (Exp * ConstantesBalance.ModTrainingExpMultiplier)
                Else
                    If Exp = 0 Then Exp = ConstantesBalance.ExpFalloSkill
                    .ExpSkills(skill) = .ExpSkills(skill) + (Exp * ConstantesBalance.ModTrainingExpMultiplier)
                End If
                
                ' Let's exhaust all the exp and give all the appropiate skills to the user.
                While GetSkills(UserIndex, skill) < MAX_SKILL_POINTS And .ExpSkills(skill) >= .EluSkills(skill) And NaturalSkillsAvailable(UserIndex, skill)
                    
                    Call AddNaturalSkills(UserIndex, skill, 1)
                    
                    Call WriteConsoleMsg(UserIndex, "¡Has mejorado tu skill " & SkillsNames(skill) & _
                                       " en un punto! Ahora tienes " & GetSkills(UserIndex, skill) & " pts.", _
                                       FontTypeNames.FONTTYPE_INFO, eMessageType.info)
                                        
                    Call CheckEluSkill(UserIndex, skill, False)
                Wend
            End With
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SubirSkill de Modulo_UsUaRiOs.bas")
End Sub

''
' Muere un usuario
'
' @param UserIndex  Indice del usuario que muere
'

Public Sub UserDie(ByVal UserIndex As Integer)
'************************************************
'Author: Uknown
'Last Modified: 12/01/2010 (ZaMa)
'04/15/2008: NicoNZ - Ahora se resetea el counter del invi
'13/02/2009: ZaMa - Ahora se borran las mascotas cuando moris en agua.
'27/05/2009: ZaMa - El seguro de resu no se activa si estas en una arena.
'21/07/2009: Marco - Al morir se desactiva el comercio seguro.
'16/11/2009: ZaMa - Al morir perdes la criatura que te pertenecia.
'27/11/2009: Budi - Al morir envia los atributos originales.
'12/01/2010: ZaMa - Los druidas pierden la inmunidad de ser atacados cuando mueren.
'************************************************
On Error GoTo ErrorHandler
    Dim I As Long
    Dim aN As Integer
    
    Dim iSoundDeath As Integer
    
    With UserList(UserIndex)
        'Sonido
        If .Genero = eGenero.Mujer Then
            If HayAgua(.Pos.Map, .Pos.X, .Pos.Y) Then
                iSoundDeath = e_SoundIndex.MUERTE_MUJER_AGUA
            Else
                iSoundDeath = e_SoundIndex.MUERTE_MUJER
            End If
        Else
            If HayAgua(.Pos.Map, .Pos.X, .Pos.Y) Then
                iSoundDeath = e_SoundIndex.MUERTE_HOMBRE_AGUA
            Else
                iSoundDeath = e_SoundIndex.MUERTE_HOMBRE
            End If
        End If
        
        Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, iSoundDeath)
        
        'Quitar el dialogo del user muerto
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageRemoveCharDialog(.Char.CharIndex))
        
        .Stats.MinHp = 0
        .Stats.MinSta = 0
        .flags.AtacadoPorUser = 0
        .flags.Envenenado = 0
        .flags.Muerto = 1
        
        ' No se activa en arenas
        If TriggerZonaPelea(UserIndex, UserIndex) <> TRIGGER6_PERMITE Then
            .flags.SeguroResu = True
            Call WriteMultiMessage(UserIndex, eMessages.ResuscitationSafeOn) 'Call WriteResuscitationSafeOn(UserIndex)
        Else
            .flags.SeguroResu = False
            Call WriteMultiMessage(UserIndex, eMessages.ResuscitationSafeOff) 'Call WriteResuscitationSafeOff(UserIndex)
        End If
        
        aN = .flags.AtacadoPorNpc
        If aN > 0 Then
            Npclist(aN).Movement = Npclist(aN).flags.OldMovement
            Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
            Npclist(aN).flags.AttackedBy = vbNullString
        End If
        
        aN = .flags.NPCAtacado
        If aN > 0 Then
            If Npclist(aN).flags.AttackedFirstBy = .Name Then
                Npclist(aN).flags.AttackedFirstBy = vbNullString
            End If
        End If
        .flags.AtacadoPorNpc = 0
        .flags.NPCAtacado = 0
        
        Call PerdioNpc(UserIndex, False)
        
        '<<<< Atacable >>>>
        If .flags.AtacablePor > 0 Then
            .flags.AtacablePor = 0
            Call RefreshCharStatus(UserIndex, False)
        End If
        
        '<<<< Paralisis & Inmo >>>>
        If .flags.Paralizado = 1 Or .flags.Inmovilizado Then
            .flags.Paralizado = 0
            .flags.Inmovilizado = 0
            .flags.Putrefaccion = 0
            .flags.Petrificado = 0
            Call WriteParalizeOK(UserIndex)
        End If
        
        '<<< Estupidez >>>
        If .flags.Estupidez = 1 Then
            .flags.Estupidez = 0
            Call WriteDumbNoMore(UserIndex)
        End If
        
        '<<<< Descansando >>>>
        If .flags.Descansar Then
            .flags.Descansar = False
            Call WriteRestOK(UserIndex)
        End If
        
        '<<<< Meditando >>>>
        If .flags.Meditando Then
            .flags.Meditando = False
            Call WriteMeditateToggle(UserIndex)
        End If
        
        '<<<< Invisible >>>>
        If .flags.invisible = 1 Or .flags.Oculto = 1 Then
            .flags.Oculto = 0
            .flags.invisible = 0
            .Counters.TiempoOculto = 0
            .Counters.Invisibilidad = 0
            
            'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
            Call UsUaRiOs.SetInvisible(UserIndex, UserList(UserIndex).Char.CharIndex, False)
        End If
        
        ' Disable the Berzerk
        If HasPassiveAssigned(UserIndex, ePassiveSpells.Berserk) Then
            If HasPassiveActivated(UserIndex, ePassiveSpells.Berserk) Then
                Call ActivatePassive(UserIndex, ePassiveSpells.Berserk, False)
                Call SendBerserkEffect(UserIndex, ePassiveSpells.Berserk, False)
            End If
        End If
        
        If .CraftingStore.IsOpen Then
            Call modCrafting.CloseWorkerStore(UserIndex)
        End If
        
        If TriggerZonaPelea(UserIndex, UserIndex) <> eTrigger6.TRIGGER6_PERMITE Then
            ' << Si es newbie no pierde el inventario >>
            
            If Not EsNewbie(UserIndex) Then
                Call TirarTodo(UserIndex)
            Else
                Call TirarTodosLosItemsNoNewbies(UserIndex)
            End If
        Else
            If EnMapaDuelos(UserIndex) Then
                If UserList(UserIndex).flags.DueloIndex > 0 Then
                    DuelData.Duelo(UserList(UserIndex).flags.DueloIndex).Team(GetUserTeam(UserList(UserIndex).flags.DueloIndex, UserIndex)).Muerto(GetTeamSlot(UserList(UserIndex).flags.DueloIndex, GetUserTeam(UserList(UserIndex).flags.DueloIndex, UserIndex), UserIndex)) = True
                    Call CheckDueloPlayersState(UserList(UserIndex).flags.DueloIndex)
                End If
            End If
        End If
        
        ' DESEQUIPA TODOS LOS OBJETOS
        'desequipar armadura
        If .Invent.ArmourEqpObjIndex > 0 Then
            If .Invent.FactionArmourEqpObjIndex > 0 Then
                Call Desequipar(UserIndex, .Invent.FactionArmourEqpSlot, False)
            End If
            
            Call Desequipar(UserIndex, .Invent.ArmourEqpSlot, False)
        End If
        
        'desequipar arma
        If .Invent.WeaponEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.WeaponEqpSlot, False)
        End If
        
        'desequipar casco
        If .Invent.CascoEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.CascoEqpSlot, False)
        End If
        
        'desequipar herramienta
        If .Invent.AnilloEqpSlot > 0 Then
            Call Desequipar(UserIndex, .Invent.AnilloEqpSlot, False)
        End If
        'TODO_TORNEO: revivir etc si faltan rounds por ganar bla bla..
        'desequipar municiones
        If .Invent.MunicionEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.MunicionEqpSlot, False)
        End If
        
        'desequipar escudo
        If .Invent.EscudoEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.EscudoEqpSlot, False)
        End If
        
        ' << Reseteamos los posibles FX sobre el personaje >>
        If .Char.Loops = INFINITE_LOOPS Then
            .Char.FX = 0
            .Char.Loops = 0
        End If
        
        ' << Restauramos el mimetismo
        If .flags.Mimetizado <> 0 Then
            Call EndMimic(UserIndex, False, False)
        End If
        
        ' << Restauramos los atributos >>
        If .flags.TomoPocion = True Then
            For I = 1 To 5
                .Stats.UserAtributos(I) = .Stats.UserAtributosBackUP(I)
            Next I
        End If
        
        '<< Cambiamos la apariencia del char >>
        If .flags.Navegando = 0 Then
            .Char.body = ConstantesGRH.CuerpoMuerto
            .Char.head = ConstantesGRH.CabezaMuerto
            .Char.ShieldAnim = ConstantesGRH.NingunEscudo
            .Char.WeaponAnim = ConstantesGRH.NingunArma
            .Char.CascoAnim = ConstantesGRH.NingunCasco
        Else
            .Char.body = ConstantesGRH.FragataFantasmal
        End If
        
        For I = 1 To Classes(.clase).ClassMods.MaxInvokedPets
        
            ' Remove invoked pets.
            If .InvokedPets(I).NpcIndex > 0 Then
                Call MuereNpc(.InvokedPets(I).NpcIndex, 0)
            ' Si estan en agua o zona segura
            Else
                .InvokedPets(I).NpcNumber = 0
            End If
            .InvokedPets(I).RemainingLife = 0
            
        Next I
        
        For I = 1 To Classes(.clase).ClassMods.MaxTammedPets

            ' Remove tammed pets, but don't "release" the slot.
            If .TammedPets(I).NpcIndex > 0 Then
                Call MuereNpc(.TammedPets(I).NpcIndex, 0)
            End If
            .TammedPets(I).RemainingLife = 0
            
        Next I
        
        '<< Actualizamos clientes >>
        Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.heading, ConstantesGRH.NingunArma, ConstantesGRH.NingunEscudo, ConstantesGRH.NingunCasco)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0, , 0))
        Call WriteUpdateUserStats(UserIndex)
        Call WriteUpdateStrenghtAndDexterity(UserIndex)
        '<<Castigos por party>>
        If .PartyIndex > 0 Then
            Call mdParty.ObtenerExito(UserIndex, .Stats.ELV * -10 * mdParty.CantMiembros(UserIndex), .Pos.Map, .Pos.X, .Pos.Y)
        End If
        
        '<<Cerramos comercio seguro>>
        Call LimpiarComercioSeguro(UserIndex)
        
        ' Hay que teletransportar?
        Dim mapa As Integer
        mapa = .Pos.Map
        Dim MapaTelep As Integer
        MapaTelep = MapInfo(mapa).OnDeathGoTo.Map
        
        If MapaTelep <> 0 Then
            Call WriteConsoleMsg(UserIndex, "¡¡¡Tu estado no te permite permanecer en el mapa!!!", FontTypeNames.FONTTYPE_INFOBOLD)
            Call WarpUserChar(UserIndex, MapaTelep, MapInfo(mapa).OnDeathGoTo.X, _
                MapInfo(mapa).OnDeathGoTo.Y, True, True)
        End If
        
    End With
Exit Sub

ErrorHandler:
    Call LogError("Error en SUB USERDIE. Error: " & Err.Number & " Descripción: " & Err.Description)
End Sub

Public Sub ContarMuerte(ByVal VictimIndex As Integer, ByVal AttackerIndex As Integer, ByVal DamageType As eDamageType, ByVal DamageValue As Long, Optional ByVal DamageWeaponIndex As Integer = 0)
'***************************************************
'Author: Unknown
'Last Modification: 26/02/15
'13/07/2010: ZaMa - Los matados en estado atacable ya no suman frag.
'25/01/2011: ZaMa - Now frags are stored in quest stats.
'24/02/14: Mithrandir - Agregé para contar los frags en desafio
'31/01/15: D'Artagnan - Update database fields.
'26/02/15: D'Artagnan - Fixed runtime error 6 overflow.
'***************************************************


On Error GoTo ErrHandler:
    If EsNewbie(VictimIndex) Then Exit Sub
    
    If TriggerZonaPelea(VictimIndex, AttackerIndex) = TRIGGER6_PERMITE Then Exit Sub
    
    With UserList(AttackerIndex)
      
        Select Case UserList(VictimIndex).Faccion.Alignment
            Case eCharacterAlignment.Neutral
                If .Faccion.NeutralsKilled < ConstantesBalance.MaxUsersMatados Then
                    .Faccion.NeutralsKilled = .Faccion.NeutralsKilled + 1
                End If
            
            Case eCharacterAlignment.FactionRoyal
                If .Faccion.CiudadanosMatados < ConstantesBalance.MaxUsersMatados Then
                    .Faccion.CiudadanosMatados = .Faccion.CiudadanosMatados + 1
                End If
            
            Case eCharacterAlignment.FactionLegion
                If .Faccion.CriminalesMatados < ConstantesBalance.MaxUsersMatados Then
                    .Faccion.CriminalesMatados = .Faccion.CriminalesMatados + 1
                End If
        
        End Select
              
        If .Stats.UsuariosMatados < ConstantesBalance.MaxUsersMatados Then _
            .Stats.UsuariosMatados = .Stats.UsuariosMatados + 1

        ' Here we're calculating both the RankingPOints for then player and the RankingPoints for the guild
        ' PlayerPoints / RankingPoints is only available after certain level, for both players (Victim and Attacker)
        If .Stats.ELV >= ConstantesBalance.RankingMinLevel And UserList(VictimIndex).Stats.ELV >= ConstantesBalance.RankingMinLevel Then
        
            Dim PreviousVictimPoints As Long
            Dim PreviousAttackerPoints As Long
            Dim AttackerPointsWon As Long
            Dim AttackerGuildPointsWon As Long
            Dim VictimPointsLost As Long
            Dim VictimGuildPointsLost As Long
            Dim AttackerMessage As String
            Dim VictimMessage As String
                        
            ' PlayerPoints calculation
            PreviousAttackerPoints = .Stats.RankingPoints
            PreviousVictimPoints = UserList(VictimIndex).Stats.RankingPoints
            
            Dim CombatEloResult As tEloCombatResult
            
            CombatEloResult = CalculateEloPoints(.Stats.RankingPoints, UserList(VictimIndex).Stats.RankingPoints)
            
            .Stats.RankingPoints = CombatEloResult.AttackerNewPoints
            UserList(VictimIndex).Stats.RankingPoints = CombatEloResult.VictimNewPoints
            
            AttackerMessage = "Has ganado " & CombatEloResult.AttackerPointsDifference & " puntos de ranking"
            VictimMessage = "Has perdido " & (CombatEloResult.VictimPointsDifference * -1) & " puntos de ranking"
            
            Dim GuildCombatEloResult As tEloCombatResult
            Dim AttackerGuildName As String, VictimGuildName As String
            
            If .Guild.IdGuild > 0 And UserList(VictimIndex).Guild.IdGuild > 0 And .Guild.IdGuild <> UserList(VictimIndex).Guild.IdGuild Then
                ' RankingPoints calculation
                PreviousAttackerPoints = GuildList(.Guild.GuildIndex).RankingPoints
                PreviousVictimPoints = GuildList(UserList(VictimIndex).Guild.GuildIndex).RankingPoints
                
                GuildCombatEloResult = CalculateEloPoints(GuildList(.Guild.GuildIndex).RankingPoints, GuildList(UserList(VictimIndex).Guild.GuildIndex).RankingPoints)
                
                GuildList(.Guild.GuildIndex).RankingPoints = GuildCombatEloResult.AttackerNewPoints
                GuildList(UserList(VictimIndex).Guild.GuildIndex).RankingPoints = GuildCombatEloResult.VictimNewPoints
                
                AttackerMessage = AttackerMessage & " y " & GuildCombatEloResult.AttackerPointsDifference & " puntos de ranking de clan"
                VictimMessage = VictimMessage & " y " & (GuildCombatEloResult.VictimPointsDifference * -1) & " puntos de ranking de clan"
                                
                AttackerGuildName = GuildList(.Guild.GuildIndex).Name
                VictimGuildName = GuildList(UserList(VictimIndex).Guild.GuildIndex).Name
                
            End If
            
            Call WriteConsoleMsg(AttackerIndex, "¡" & AttackerMessage & "!", FontTypeNames.FONTTYPE_FIGHT)
            Call WriteConsoleMsg(VictimIndex, "¡" & VictimMessage & "!", FontTypeNames.FONTTYPE_FIGHT)
            
            Dim AttackerPos As String, VictimPos As String
            AttackerPos = .Pos.Map & "-" & .Pos.X & "-" & .Pos.Y
            VictimPos = UserList(VictimIndex).Pos.Map & "-" & UserList(VictimIndex).Pos.X & "-" & UserList(VictimIndex).Pos.Y
            
            ' Send the event data to the online ranking
            Call modMessageQueueProxy.SendDeathEventUserKilledUser(.ID, .Name, AttackerPos, .Stats.ELV, CombatEloResult.AttackerPreviousPoints, _
                                                                    CombatEloResult.AttackerNewPoints, UserList(VictimIndex).ID, UserList(VictimIndex).Name, _
                                                                    VictimPos, UserList(VictimIndex).Stats.ELV, CombatEloResult.VictimPreviousPoints, _
                                                                    CombatEloResult.VictimNewPoints, DamageType, DamageValue, DamageWeaponIndex, _
                                                                    .Guild.IdGuild, AttackerGuildName, UserList(VictimIndex).Guild.IdGuild, VictimGuildName, _
                                                                    GuildCombatEloResult.AttackerPreviousPoints, GuildCombatEloResult.AttackerNewPoints, _
                                                                    GuildCombatEloResult.VictimPreviousPoints, GuildCombatEloResult.VictimNewPoints)
                                                                    
            
        End If
        
        ' Update guild quest status
        If GuildHasQuest(UserList(AttackerIndex).Guild.GuildIndex) Then
            Call modQuestSystem.GuildQuestUpdateStatus(UserList(AttackerIndex).Guild.GuildIndex, AttackerIndex, VictimIndex, eQuestRequirement.UserKill, 1, 1)
        End If
        
    End With
    
    Exit Sub
    
ErrHandler:
    Call LogError("Error en ContarMuerte. Error: " & Err.Number & " - " & Err.Description & " - Atacante: " & UserList(AttackerIndex).Name & " - Victima: " & UserList(VictimIndex).Name)
End Sub

Public Function CalculateEloPoints(ByRef AttackerPoints As Long, ByRef VictimPoints As Long) As tEloCombatResult
    
    'Dim PreviousAttackerPoints As Double, PreviousVictimPoints As Double
    Dim ExpectedAttackerOutcome As Double, ExpectedVictimOutcome As Double
    Dim CombatResult As tEloCombatResult
    
    With CombatResult
        .AttackerPreviousPoints = AttackerPoints
        .VictimPreviousPoints = VictimPoints
        
        .SkewDistanceUsed = ConstantesBalance.RankingSkewDistance
        
        'PreviousAttackerPoints = AttackerPoints
        'PreviousVictimPoints = VictimPoints
        
        ExpectedAttackerOutcome = 10 ^ (.AttackerPreviousPoints / 400)
        ExpectedVictimOutcome = 10 ^ (.VictimPreviousPoints / 400)
        
        'AttackerPoints = Round(.AttackerPreviousPoints + (ConstantesBalance.RankingSkewDistance * (1 - (ExpectedAttackerOutcome / (ExpectedAttackerOutcome + ExpectedVictimOutcome)))))
        'VictimPoints = Round(.VictimPreviousPoints + (ConstantesBalance.RankingSkewDistance * (0 - (ExpectedVictimOutcome / (ExpectedAttackerOutcome + ExpectedVictimOutcome)))))
        .AttackerNewPoints = Round(.AttackerPreviousPoints + (ConstantesBalance.RankingSkewDistance * (1 - (ExpectedAttackerOutcome / (ExpectedAttackerOutcome + ExpectedVictimOutcome)))))
        .VictimNewPoints = Round(.VictimPreviousPoints + (ConstantesBalance.RankingSkewDistance * (0 - (ExpectedVictimOutcome / (ExpectedAttackerOutcome + ExpectedVictimOutcome)))))
        
        .AttackerPointsDifference = .AttackerNewPoints - .AttackerPreviousPoints
        .VictimPointsDifference = .VictimNewPoints - .VictimPreviousPoints

    End With
    
    CalculateEloPoints = CombatResult
    
End Function

Sub Tilelibre(ByRef Pos As WorldPos, ByRef nPos As WorldPos, ByRef Obj As Obj, _
              ByRef PuedeAgua As Boolean, ByRef PuedeTierra As Boolean)
'**************************************************************
'Author: Unknown
'Last Modify Date: 18/09/2010
'23/01/2007 -> Pablo (ToxicWaste): El agua es ahora un TileLibre agregando las condiciones necesarias.
'18/09/2010: ZaMa - Aplico optimizacion de busqueda de tile libre en forma de rombo.
'**************************************************************
On Error GoTo ErrHandler

    Dim Found As Boolean
    Dim LoopC As Integer
    Dim tX As Long
    Dim tY As Long
    
    nPos = Pos
    tX = Pos.X
    tY = Pos.Y
    
    LoopC = 1
    
    ' La primera posicion es valida?
    If LegalPos(Pos.Map, nPos.X, nPos.Y, PuedeAgua, PuedeTierra, True) Then
        
        If Not HayObjeto(Pos.Map, nPos.X, nPos.Y, Obj.ObjIndex, Obj.Amount) Then
            Found = True
        End If
        
    End If
    
    ' Busca en las demas posiciones, en forma de "rombo"
    If Not Found Then
        While (Not Found) And LoopC <= 16
            If RhombLegalTilePos(Pos, tX, tY, LoopC, Obj.ObjIndex, Obj.Amount, PuedeAgua, PuedeTierra) Then
                nPos.X = tX
                nPos.Y = tY
                Found = True
            End If
        
            LoopC = LoopC + 1
        Wend
        
    End If
    
    If Not Found Then
        nPos.X = 0
        nPos.Y = 0
    End If
    
    Exit Sub
    
ErrHandler:
    Call LogError("Error en Tilelibre. Error: " & Err.Number & " - " & Err.Description)
End Sub

Function SearchObjectNearPlayer(ByVal ObjIndex As Long, ByRef UserPos As WorldPos, ByRef foundPoss As WorldPos) As Boolean
    Dim LoopC As Integer
    Dim tX As Long
    Dim tY As Long
    Dim tMap As Long
    
    tX = UserPos.X
    tY = UserPos.Y
    
    While LoopC <= 2
        If RhombObjExists(UserPos, tX, tY, LoopC, ObjIndex) Then
            foundPoss.Map = UserPos.Map
            foundPoss.X = tX
            foundPoss.Y = tY
            
            SearchObjectNearPlayer = True
            Exit Function
        End If
    
        LoopC = LoopC + 1
    Wend
        
    SearchObjectNearPlayer = False
    Exit Function
    
ErrHandler:
    Call LogError("Error en SearchObjectNearPlayer. Error: " & Err.Number & " - " & Err.Description)
End Function

Function GetItemTypeSlot(ByVal UserIndex As Integer, ItemType As eOBJType) As Byte
On Error GoTo ErrHandler
    Dim I

    With UserList(UserIndex)
         For I = 1 To .CurrentInventorySlots
            If .Invent.Object(I).ObjIndex Then
                If ObjData(.Invent.Object(I).ObjIndex).ObjType = ItemType Then
                    If .Invent.Object(I).ObjIndex Then
                        GetItemTypeSlot = I
                        Exit Function
                    End If
                End If
            End If
        Next I
        Exit Function
    End With
    
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub GetItemTypeSlot de Modulo_UsUaRiOs.bas")
End Function

Function GetBoatSlot(ByVal UserIndex As Integer) As Byte
On Error GoTo ErrHandler
    
    With UserList(UserIndex)
        If Not .Invent.BarcoSlot Then
            .Invent.BarcoSlot = GetItemTypeSlot(UserIndex, eOBJType.otBarcos)
        ElseIf Not .Invent.Object(.Invent.BarcoSlot).ObjIndex Then
            .Invent.BarcoSlot = GetItemTypeSlot(UserIndex, eOBJType.otBarcos)
        End If
        GetBoatSlot = .Invent.BarcoSlot
        Exit Function
    End With

ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub GetBoatSlot de Modulo_UsUaRiOs.bas")
End Function

Sub WarpUserChar(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, _
ByVal FX As Boolean, Optional ByVal Teletransported As Boolean, Optional ByVal Forced As Boolean)
On Error GoTo ErrHandler
  
    Dim OldMap As Integer
    Dim OldX As Integer
    Dim OldY As Integer
    
    With UserList(UserIndex)
    
        'Quitar el dialogo solo si no es GM.
        If .flags.AdminInvisible = 0 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageRemoveCharDialog(.Char.CharIndex))
        End If
        
        OldMap = .Pos.Map
        OldX = .Pos.X
        OldY = .Pos.Y

        If OldMap <> Map Then
            Call EraseUserChar(UserIndex)

            .flags.Inmunidad = 1
            .Counters.Inmunidad = SetIntervalEnd(ServerConfiguration.Intervals.IntervaloInmunidad)
            Call WriteChangeMap(UserIndex, Map, MapInfo(.Pos.Map).MapVersion)

            If .flags.Privilegios And PlayerType.User Then 'El chequeo de invi/ocultar solo afecta a Usuarios (C4b3z0n)
                Dim AhoraVisible As Boolean 'Para enviar el mensaje de invi y hacer visible (C4b3z0n)
                Dim WasInvi As Boolean
                'Chequeo de flags de mapa por invisibilidad (C4b3z0n)
                If MapInfo(Map).InviSinEfecto > 0 And .flags.invisible = 1 Then
                    .flags.invisible = 0
                    .Counters.Invisibilidad = 0
                    AhoraVisible = True
                    WasInvi = True 'si era invi, para el string
                End If
                'Chequeo de flags de mapa por ocultar (C4b3z0n)
                If MapInfo(Map).OcultarSinEfecto > 0 And .flags.Oculto = 1 Then
                    AhoraVisible = True
                    .flags.Oculto = 0
                    .Counters.TiempoOculto = 0
                End If
                
                If AhoraVisible Then 'Si no era visible y ahora es, le avisa. (C4b3z0n)
                    Call UsUaRiOs.SetInvisible(UserIndex, .Char.CharIndex, False)
                    If WasInvi Then 'era invi
                        Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible ya que no esta permitida la invisibilidad en este mapa.", FontTypeNames.FONTTYPE_INFO)
                    Else 'estaba oculto
                        Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible ya que no esta permitido ocultarse en este mapa.", FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
            End If
            
            Call WritePlayMusic(UserIndex, Map)
            
            'Update new Map Users
            MapInfo(Map).NumUsers = MapInfo(Map).NumUsers + 1

            'Update old Map Users
            MapInfo(OldMap).NumUsers = MapInfo(OldMap).NumUsers - 1

            If MapInfo(OldMap).NumUsers < 0 Then
                MapInfo(OldMap).NumUsers = 0
            End If

            .flags.lastMap = .Pos.Map
            
            If .flags.Privilegios = PlayerType.User Or .flags.Privilegios = PlayerType.RoyalCouncil Or .flags.Privilegios = PlayerType.ChaosCouncil Then
                Call WriteRemoveAllDialogs(UserIndex)
            End If
            
             ' TODO: have the collision library handle different map to remove this conditional
            
            .Pos.X = X
            .Pos.Y = Y
            .Pos.Map = Map
            
            If (Not NewUserChar(UserIndex)) Then
                Exit Sub
            End If

            If (.flags.AdminInvisible) Then
                Call WriteSetInvisible(UserIndex, .Char.CharIndex, True, True)
            End If
                        
            If Forced = True And Teletransported = True And .flags.DueloIndex > 0 Then
                Call modDuelos.CerrarDuelo(.flags.DueloIndex)
            End If
            
        Else
            MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex = 0
            MapData(.Pos.Map, X, Y).UserIndex = UserIndex
                        
            .Pos.X = X
            .Pos.Y = Y
            .Pos.Map = Map

            Call WritePosUpdate(UserIndex)
            Call ModAreas.UpdateEntity(UserIndex, ENTITY_TYPE_PLAYER, .Pos, True)
        End If

        Call DoTileEvents(UserIndex, Map, X, Y)
        
        ' Step on trigger?
        Call CheckTriggerActivation(UserIndex, 0, Map, X, Y, False)

        'Seguis invisible al pasar de mapa, excepto si es mapa de desafio
        If (.flags.invisible = 1 Or .flags.Oculto = 1) And (Not .flags.AdminInvisible = 1) Then
            If .Challenge.InSand > 0 Then
                If SandsChallenge(.Challenge.InSand).Invisibility = 1 Then
                    .flags.Oculto = 0
                    .flags.invisible = 0
                    .Counters.TiempoOculto = 0
                    .Counters.Invisibilidad = 0
                    Call SetInvisible(UserIndex, .Char.CharIndex, False)
                End If
            Else
                ' No si estas navegando
                If .flags.Navegando = 0 Then
                    Call SetInvisible(UserIndex, .Char.CharIndex, True)
                End If
            End If
        End If
        
        If Teletransported Then
            If .flags.Traveling = 1 Then
                Call EndTravel(UserIndex, True)
            End If
        End If
        
        If FX And .flags.AdminInvisible = 0 Then 'FX
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(ConstantesSonidos.Warp, X, Y, .Char.CharIndex))
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, ConstantesFX.FxWarp, 0))
        End If
        
        Call DisableAllTrapsForUser(UserIndex)
        
        If .TammedPetsCount > 0 Then
            Call WarpMascotas(UserIndex)
        End If
        
        If .InvokedPetsCount > 0 Then
            Dim I As Integer
            For I = 1 To Classes(.clase).ClassMods.MaxInvokedPets
                Call QuitarInvocacion(UserIndex, I)
            Next I
            
        End If

        If Forced = False Then 'Si fue forzado hacia esa posicion no tomo el intervalo
            ' No puede ser atacado cuando cambia de mapa, por cierto tiempo
            Call IntervaloPermiteSerAtacado(UserIndex, True)
        End If
        
        ' Perdes el npc al cambiar de mapa
        Call PerdioNpc(UserIndex, False)
        
        ' Automatic toggle navigate
        If HayAgua(.Pos.Map, .Pos.X, .Pos.Y) Then
            If .flags.Navegando = 0 Then
                If EsGm(UserIndex) Or PlayerType.User = 0 Then
                    .flags.Navegando = 1
                    
                    'Tell the client that we are navigating.
                    Call WriteNavigateChange(UserIndex, True)
                ElseIf GetBoatSlot(UserIndex) Then
                    Call DoNavega(UserIndex, .Invent.BarcoSlot)
                End If
            End If
        Else
            If .flags.Navegando = 1 Then
                If GetBoatSlot(UserIndex) Then
                    Call DoNavega(UserIndex, .Invent.BarcoSlot)
                Else
                    .flags.Navegando = 0
                    Call WriteNavigateChange(UserIndex, False)
                End If
            End If
        End If
      
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WarpUserChar de Modulo_UsUaRiOs.bas")
End Sub

Public Sub WarpMascotas(ByVal UserIndex As Integer)
'************************************************
'Author: Uknown
'Last Modified: 26/10/2010
'13/02/2009: ZaMa - Arreglado respawn de mascotas al cambiar de mapa.
'13/02/2009: ZaMa - Las mascotas no regeneran su vida al cambiar de mapa (Solo entre mapas inseguros).
'11/05/2009: ZaMa - Chequeo si la mascota pueden spwnear para asiganrle los stats.
'26/10/2010: ZaMa - Ahora las mascotas rapswnean de forma aleatoria.
'************************************************
On Error GoTo ErrHandler
  
    Dim I As Integer
    Dim petType As Integer
    Dim PetRespawn As Boolean
    Dim PetTiempoDeVida As Integer
    Dim NroPets As Integer
    Dim InvocadosMatados As Integer
    Dim canWarp As Boolean
    Dim canSummon As Boolean
    Dim Index As Integer
    Dim iMinHP As Integer
    Dim SpawnPos As WorldPos
    Dim IsWaterTile As Boolean
    Dim canMarpIndividualPet As Boolean
    
    With UserList(UserIndex)
        SpawnPos.Map = .Pos.Map
        SpawnPos.X = .Pos.X + RandomNumber(-3, 3)
        SpawnPos.Y = .Pos.Y + RandomNumber(-3, 3)
    
        IsWaterTile = HayAgua(SpawnPos.Map, SpawnPos.X, SpawnPos.Y)
       
        'If HayAgua(SpawnPos.Map, SpawnPos.X, SpawnPos.Y) Then
        '    Call WriteConsoleMsg(UserIndex, "Tus mascotas no pueden transitar este mapa.", FontTypeNames.FONTTYPE_INFO)
        '    Exit Sub
        'End If
    
        NroPets = .TammedPetsCount
        canWarp = (MapInfo(.Pos.Map).Pk = True)
        'recupero el valor de InvocarSinEfecto (Aclaracion: si el mapa no posee la propiedad devuelve False)
        canSummon = Not (MapInfo(.Pos.Map).InvocarSinEfecto = 1)

        For I = 1 To Classes(.clase).ClassMods.MaxTammedPets
            Index = .TammedPets(I).NpcIndex
   
            If Index > 0 Then
                ' si la mascota tiene tiempo de vida > 0 significa q fue invocada => we kill it
                If Npclist(Index).Contadores.TiempoExistencia > 0 Then
                    Call QuitarNPC(Index)
                    .TammedPets(I).NpcIndex = 0
                    InvocadosMatados = InvocadosMatados + 1
                    NroPets = NroPets - 1
                    .TammedPetsCount = NroPets
                
                    petType = 0
                Else
                    'Store data and remove NPC to recreate it after warp
                    'PetRespawn = Npclist(index).flags.Respawn = 0
                    petType = .TammedPets(I).NpcNumber
                    'PetTiempoDeVida = Npclist(index).Contadores.TiempoExistencia
                
                    ' Guardamos el hp, para restaurarlo uando se cree el npc
                    iMinHP = Npclist(Index).Stats.MinHp
                    .TammedPets(I).RemainingLife = iMinHP
                
                    Call QuitarNPC(Index)
                
                    ' Restauramos el valor de la variable
                    .TammedPets(I).NpcNumber = petType
                    .TammedPets(I).RemainingLife = iMinHP
                
                End If
            Else
                petType = 0
            End If
                
            If petType > 0 And (canWarp And canSummon) Then
                canMarpIndividualPet = True
            
                 ' Check whether the Pet can be invoked or not based on his ability to walk over the water or ground.
                If IsWaterTile And NpcData(petType).flags.AguaValida = 0 Then
                    Call WriteConsoleMsg(UserIndex, "Tu mascota " & NpcData(petType).Name & " no puede transitar sobre el agua. Intenta invocarla en otra parte.", FontTypeNames.FONTTYPE_INFO)
                    canMarpIndividualPet = False
                End If
                If Not IsWaterTile And NpcData(petType).flags.TierraInvalida = 1 Then
                    Call WriteConsoleMsg(UserIndex, "Tu mascota " & NpcData(petType).Name & " no puede transitar sobre la tierra. Intenta invocarla en otra parte.", FontTypeNames.FONTTYPE_INFO)
                    canMarpIndividualPet = False
                End If
        
                If canMarpIndividualPet Then
                    Index = SpawnNpc(petType, SpawnPos, False, PetRespawn)
                    .TammedPets(I).NpcIndex = Index
    
                    ' Nos aseguramos de que conserve el hp, si estaba dañado
                    If (.TammedPets(I).RemainingLife <> 0) Then
                        Npclist(Index).Stats.MinHp = .TammedPets(I).RemainingLife
                    End If
                    Npclist(Index).MenuIndex = eMenues.ieMascota
                    Npclist(Index).MaestroUser = UserIndex
                    Npclist(Index).Contadores.TiempoExistencia = PetTiempoDeVida
                    Call FollowAmo(Index)
                End If
            End If
        Next I
    
        If InvocadosMatados > 0 Then
            Call WriteConsoleMsg(UserIndex, "Pierdes el control de tus mascotas invocadas.", FontTypeNames.FONTTYPE_INFO)
        End If
    
        If Not canWarp Then
            Call WriteConsoleMsg(UserIndex, "No se permiten mascotas en zona segura. Éstas te esperarán afuera.", FontTypeNames.FONTTYPE_INFO)
        End If
                        'evito que escriba 2 mensajes
        If Not canSummon And canWarp Then
            Call WriteConsoleMsg(UserIndex, "No se permiten invocar mascotas en esta zona. Éstas te esperarán afuera.", FontTypeNames.FONTTYPE_INFO)
        End If
        .PetAliveCount = NroPets
        
    End With

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WarpMascotas de Modulo_UsUaRiOs.bas")
End Sub

Public Function RevivePet(ByVal UserIndex As Integer) As Boolean
On Error GoTo ErrHandler
  
    Dim Index As Integer
    Dim IsWaterTile As Boolean
    Dim I As Integer
    Dim ActiveTammedPetsQty As Byte
    
    RevivePet = False
    
    With UserList(UserIndex)
        
        If .flags.TargetUser <> UserIndex Then
            Call WriteConsoleMsg(UserIndex, "No puedes revivir las mascotas de otros usuarios.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        
        ' Validate if its a safe zone or not. Pet's can't be invoked or revived in safe zones
        Dim safeZone As Boolean
        safeZone = (MapInfo(.Pos.Map).Pk = True)
        If Not safeZone Then
            Call WriteConsoleMsg(UserIndex, "No puedes resucitar mascotas en zona segura.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        
        ' If the user doesn't have any tamed pet, exit
        If .TammedPetsCount = 0 Then
            Call WriteConsoleMsg(UserIndex, "No tienes ninguna mascota para revivir.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        
        ' If there's no selected slot, exit
        If .SelectedPet = 0 Then
            Call WriteConsoleMsg(UserIndex, "No tienes ninguna mascota seleccionada para revivir.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        
        'If the selected slot doesn't have any tamed pet, exit
        If .TammedPets(.SelectedPet).NpcNumber = 0 Then
            Call WriteConsoleMsg(UserIndex, "No tienes ninguna mascota en el slot seleccionado.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        
        Dim PetIndex As Integer
        PetIndex = .TammedPets(.SelectedPet).NpcIndex
        If PetIndex > 0 Then
            Call WriteConsoleMsg(UserIndex, "Tu mascota se encuentra viva.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        
        ' if the pet is already alive, exit
        If .TammedPets(.SelectedPet).RemainingLife <> 0 Then
            Call WriteConsoleMsg(UserIndex, "Tu mascota se encuentra viva.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        
        ' validate if maps allows to invoke pets
        If MapInfo(.Pos.Map).InvocarSinEfecto Then
            Call WriteConsoleMsg(UserIndex, "¡Revivir mascotas no está permitido aquí! Retirate del mapa si deseas utilizar el Hechizo.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
 
        Dim NpcIndex As Integer
        Dim petPoss As WorldPos
        
        Call ClosestLegalPos(.Pos, petPoss, NpcData(.TammedPets(.SelectedPet).NpcNumber).flags.AguaValida = 1, Npclist(.TammedPets(.SelectedPet).NpcNumber).flags.TierraInvalida = 0)
        
        If petPoss.X = 0 Or petPoss.Y = 0 Then
            Call WriteConsoleMsg(UserIndex, "Tu mascota no puede transitar este mapa.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        
        IsWaterTile = HayAgua(petPoss.Map, petPoss.X, petPoss.Y)
        
        ' Check whether the Pet can be invoked or not based on his ability to walk over the water or ground.
        If IsWaterTile And NpcData(.TammedPets(.SelectedPet).NpcNumber).flags.AguaValida = 0 Then
            Call WriteConsoleMsg(UserIndex, "Tu mascota " & NpcData(.TammedPets(.SelectedPet).NpcNumber).Name & " no puede transitar sobre el agua. Intenta resucitarla en otra parte.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        If Not IsWaterTile And NpcData(.TammedPets(.SelectedPet).NpcNumber).flags.TierraInvalida = 1 Then
            Call WriteConsoleMsg(UserIndex, "Tu mascota " & NpcData(.TammedPets(.SelectedPet).NpcNumber).Name & " no puede transitar sobre la tierra. Intenta resucitarla en otra parte.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        
        
        For I = 1 To Classes(.Clase).ClassMods.MaxTammedPets
            If .TammedPets(I).NpcIndex <> 0 Then
                ActiveTammedPetsQty = ActiveTammedPetsQty + 1
            End If
        Next I
                
        If (.InvokedPetsCount >= Classes(.Clase).ClassMods.MaxInvokedPets Or _
           (.InvokedPetsCount + ActiveTammedPetsQty) >= Classes(.Clase).ClassMods.MaxActivePets) Then
            Call WriteConsoleMsg(UserIndex, "Has superado la cantidad máxima de mascotas invocadas.", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)
            Exit Function
        End If
                
        NpcIndex = CrearNPC(.TammedPets(.SelectedPet).NpcNumber, .Pos.Map, petPoss, , True)
        .flags.LastNpcInvoked = NpcIndex
        
        If NpcIndex = 0 Then Exit Function
        
        .TammedPets(.SelectedPet).NpcIndex = NpcIndex
        .TammedPets(.SelectedPet).NpcNumber = Npclist(NpcIndex).Numero
        .TammedPets(.SelectedPet).RemainingLife = Npclist(NpcIndex).Stats.MaxHp
        
        Npclist(NpcIndex).MaestroUser = UserIndex
        Npclist(NpcIndex).MenuIndex = eMenues.ieMascota
        
        Call FollowAmo(NpcIndex)
        
        Call WriteConsoleMsg(UserIndex, "Has revivido a tu mascota.", FontTypeNames.FONTTYPE_INFO)
                
        'If Not CanStay Then
        '    Call QuitarNPC(NpcIndex)
        '
        '    .MascotasType(.SelectedPet) = Npclist(NpcIndex).Numero
        '    '.NroMascotas = .NroMascotas
        '
        '    Call WriteConsoleMsg(UserIndex, "No se permiten mascotas en zona segura. Éstas te esperarán afuera.", FontTypeNames.FONTTYPE_INFO)
        'End If
        
        ' We need to add skill exp after reviving the pet.
        Call SubirSkill(UserIndex, eSkill.Domar, True)

        RevivePet = True
    End With
                
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function RevivePet de Modulo_UsUaRiOs.bas")
End Function

Function WarpMascota(ByVal UserIndex As Integer, ByVal PetIndex As Integer) As Boolean
'************************************************
'Author: ZaMa
'Last Modified: 18/11/2009
'Warps a pet without changing its stats
'************************************************
On Error GoTo ErrHandler
  
    WarpMascota = False
  
    Dim petType As Integer
    Dim NpcIndex As Integer
    Dim iMinHP As Integer
    Dim TargetPos As WorldPos
    Dim IsWaterTile As Boolean
    
    With UserList(UserIndex)
        
        TargetPos.Map = .flags.TargetMap
        TargetPos.X = .flags.TargetX
        TargetPos.Y = .flags.TargetY
                
        NpcIndex = .TammedPets(PetIndex).NpcIndex
        
        IsWaterTile = HayAgua(TargetPos.Map, TargetPos.X, TargetPos.Y)
        
        ' Check wether the Pet can be invoked or not based on his ability to walk over the water or ground.
        If IsWaterTile And Npclist(NpcIndex).flags.AguaValida = 0 Then
            Call WriteConsoleMsg(UserIndex, "Tu mascota " & Npclist(NpcIndex).Name & " no puede ser invocada en el lugar seleccionado. Intenta invocarla en otra parte.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        If Not IsWaterTile And Npclist(NpcIndex).flags.TierraInvalida = 1 Then
            Call WriteConsoleMsg(UserIndex, "Tu mascota " & Npclist(NpcIndex).Name & " no puede ser invocada en el lugar seleccionado. Intenta invocarla en otra parte.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
            
        'Store data and remove NPC to recreate it after warp
        petType = .TammedPets(PetIndex).NpcNumber
        
        ' Guardamos el hp, para restaurarlo cuando se cree el npc
        iMinHP = Npclist(NpcIndex).Stats.MinHp
        
        Call QuitarNPC(NpcIndex)
        
        ' Restauramos el valor de la variable
        .TammedPets(PetIndex).NpcNumber = petType
        .TammedPets(PetIndex).RemainingLife = iMinHP
        
        NpcIndex = SpawnNpc(petType, TargetPos, False, False)
        
        .TammedPets(PetIndex).NpcIndex = NpcIndex

        With Npclist(NpcIndex)
            ' Nos aseguramos de que conserve el hp, si estaba dañado
            If UserList(UserIndex).TammedPets(PetIndex).RemainingLife <> 0 Then
                .Stats.MinHp = UserList(UserIndex).TammedPets(PetIndex).RemainingLife
            End If
            
            .MaestroUser = UserIndex
            .Movement = TipoAI.SigueAmo
            .MenuIndex = eMenues.ieMascota
            .Target = 0
            .TargetNPC = 0
        End With
            
        Call FollowAmo(NpcIndex)
    End With
  
  WarpMascota = True
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WarpMascota de Modulo_UsUaRiOs.bas")
End Function


''
' Se inicia la salida de un usuario.
'
' @param    UserIndex   El index del usuario que va a salir

Sub Cerrar_Usuario(ByVal UserIndex As Integer, Optional ByVal bShowAccountForm As Boolean = False)
On Error GoTo ErrHandler
  
    Dim IsNotVisible As Boolean
    Dim HiddenPirat As Boolean
    
    With UserList(UserIndex)
        If .flags.UserLogged And Not .Counters.Saliendo Then
            .Counters.Saliendo = True
            .Counters.Salir = IIf((.flags.Privilegios And PlayerType.User) And MapInfo(.Pos.Map).Pk, ServerConfiguration.Intervals.IntervaloCerrarConexion, 0)
            
            .bShowAccountForm = bShowAccountForm
            
            IsNotVisible = (.flags.Oculto Or .flags.invisible)
            If IsNotVisible Then
                
                If .flags.Oculto Then
                    If .flags.Navegando = 1 Then
                        If .clase = eClass.Thief Then
                            ' Pierde la apariencia de fragata fantasmal
                            Call ToggleBoatBody(UserIndex)
                            Call WriteConsoleMsg(UserIndex, "¡Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                            Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.heading, ConstantesGRH.NingunArma, _
                                                ConstantesGRH.NingunEscudo, ConstantesGRH.NingunCasco)
                            HiddenPirat = True
                        End If
                    End If
                End If
                
                ' Due to some unwanted mechanics with the Berserk effect and invisibility, we need to make sure that we
                ' prevent the user to use any mechanic like that after trying to close the game
                If .flags.invisible = 1 Or .flags.Oculto = 1 Then
                    ' We make him visibile again.
                    .flags.invisible = False
                    Call UsUaRiOs.SetInvisible(UserIndex, .Char.CharIndex, False)
                    Call WriteConsoleMsg(UserIndex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
                    
                End If
                           
            End If
            
            ' The sailing flag will cause no head the next time, so remove it
            ' if we have privileges.
            If EsGm(UserIndex) Then
                If .flags.Navegando = 1 And Not BodyIsBoat(.Char.body) Then
                    .flags.Navegando = 0
                End If
                
                ' If the GM was in HelpingMode with another user, then we set the flags back and end the HelpMode.
                If .flags.HelpingUser > 0 Then
                    If UserList(.flags.HelpingUser).Name = .flags.HelpingUserName And UserList(.flags.HelpingUser).flags.HelpMode Then
                        Call SetHelpModeToUser(UserIndex, .flags.HelpingUser, False)
                    End If
                End If
            Else
                ' if the user was being helped by a Game Master, then we reset the flags
                If .flags.HelpMode = True And .flags.HelpedBy > 0 Then
                    If EsGm(.flags.HelpedBy) Then
                        Call SetHelpModeToUser(.flags.HelpedBy, UserIndex, False)
                    End If
                End If
            End If
            
            If .flags.Traveling = 1 Then
                Call EndTravel(UserIndex, True)
            End If
            
            Call WriteConsoleMsg(UserIndex, "Cerrando...Se cerrará el juego en " & .Counters.Salir & " segundos...", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Cerrar_Usuario de Modulo_UsUaRiOs.bas")
End Sub

''
' Cancels the exit of a user. If it's disconnected it's reset.
'
' @param    UserIndex   The index of the user whose exit is being reset.

Public Sub CancelExit(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/02/08
'
'***************************************************
On Error GoTo ErrHandler
  
    If UserList(UserIndex).Counters.Saliendo Then
        ' Is the user still connected?
        If UserList(UserIndex).ConnIDValida Then
            UserList(UserIndex).Counters.Saliendo = False
            UserList(UserIndex).Counters.Salir = 0
            Call WriteConsoleMsg(UserIndex, "/salir cancelado.", FontTypeNames.FONTTYPE_WARNING)
        Else
            'Simply reset
            UserList(UserIndex).Counters.Salir = IIf((UserList(UserIndex).flags.Privilegios And PlayerType.User) And MapInfo(UserList(UserIndex).Pos.Map).Pk, ServerConfiguration.Intervals.IntervaloCerrarConexion, 0)
        End If
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CancelExit de Modulo_UsUaRiOs.bas")
End Sub

Sub SendUserOROTxtFromChar(ByVal sendIndex As Integer, ByVal charName As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    Dim dOro As Double
    Dim UserID As Long
    
    UserID = GetUserID(charName)
    
    If UserID <> 0 Then
        dOro = Val(GetCharData("USER_STATS", "ORO_BANCO", UserID))
    
        Call WriteConsoleMsg(sendIndex, charName, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Tiene " & dOro & " en el banco.", FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(sendIndex, "Usuario inexistente: " & charName, FontTypeNames.FONTTYPE_INFO)
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SendUserOROTxtFromChar de Modulo_UsUaRiOs.bas")
End Sub

''
'Checks if a given body index is a boat or not.
'
'@param body    The body index to bechecked.
'@return    True if the body is a boat, false otherwise.

Public Function BodyIsBoat(ByVal body As Integer) As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 10/07/2008
'Checks if a given body index is a boat
'**************************************************************
'TODO : This should be checked somehow else. This is nasty....
On Error GoTo ErrHandler
  
    BodyIsBoat = body = ConstantesGRH.FragataReal Or body = ConstantesGRH.FragataCaos Or body = ConstantesGRH.BarcaPk Or _
                 body = ConstantesGRH.GaleraPk Or body = ConstantesGRH.GaleonPk Or body = ConstantesGRH.BarcaCiuda Or _
                 body = ConstantesGRH.GaleraCiuda Or body = ConstantesGRH.GaleonCiuda Or body = ConstantesGRH.FragataFantasmal Or _
                 body = ConstantesGRH.Barca Or body = ConstantesGRH.Galera Or body = ConstantesGRH.Galeon Or _
                 body = ConstantesGRH.BarcaCiudaAtacable Or body = ConstantesGRH.GaleraCiudaAtacable Or _
                 body = ConstantesGRH.GaleonCiudaAtacable Or body = ConstantesGRH.BarcaReal Or _
                 body = ConstantesGRH.BarcaRealAtacable Or body = ConstantesGRH.GaleraReal Or _
                 body = ConstantesGRH.GaleraReal Or body = ConstantesGRH.GaleraRealAtacable Or _
                 body = ConstantesGRH.GaleonReal Or body = ConstantesGRH.GaleonRealAtacable
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function BodyIsBoat de Modulo_UsUaRiOs.bas")
End Function

Public Sub SetInvisible(ByVal UserIndex As Integer, ByVal userCharIndex As Integer, ByVal invisible As Boolean)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

Dim sndNick As String, ClanTag As String

Dim UseInvisibilityTransparency As Boolean

With UserList(UserIndex)
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(userCharIndex, invisible))
    
    If (.Guild.GuildIndex > 0) Then
        Call SendData(SendTarget.ToClanArea, UserIndex, PrepareMessageSetInvisible(userCharIndex, invisible, GuildList(.Guild.GuildIndex).UpgradeEffect.IsSeeInvisibleGuildMember))
    End If

End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SetInvisible de Modulo_UsUaRiOs.bas")
End Sub

Public Sub SetHelpModeToUser(ByVal GMIndex As Integer, ByVal UserIndex As Integer, ByVal Enabled As Boolean)
'***************************************************
'Author: Alejandro Masolini (Nightw)
'Last Modification: 18/02/2018
'
'***************************************************
On Error GoTo ErrHandler

    Dim sndNick As String
    
    With UserList(UserIndex)
        .flags.HelpMode = Enabled
        sndNick = .Name
        
        If Enabled Then
        
            ' Set the properties of the user
            With UserList(UserIndex)
                .flags.HelpedBy = GMIndex
                .flags.HelpedByUserName = UserList(GMIndex).Name
            End With
            
            ' Set the properties of the GM
            UserList(GMIndex).flags.HelpingUser = UserIndex
            UserList(GMIndex).flags.HelpingUserName = UserList(UserIndex).Name
        
            sndNick = sndNick & " " & TAG_CONSULT_MODE
        Else
            
            ' Set the properties of the user
            With UserList(UserIndex)
                .flags.HelpMode = False
                .flags.HelpedBy = 0
                .flags.HelpedByUserName = vbNullString
            End With
            
            ' Set the properties of the GM
            UserList(GMIndex).flags.HelpingUser = 0
            UserList(GMIndex).flags.HelpingUserName = vbNullString
            
            If .Guild.IdGuild > 0 Then
                sndNick = sndNick & " <" & GuildList(.Guild.GuildIndex).Name & ">"
            End If
        End If
        
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChangeNick(.Char.CharIndex, sndNick))
    End With


    Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SetHelpModeToUser de Modulo_UsUaRiOs.bas")
End Sub


Public Function IsArena(ByVal UserIndex As Integer) As Boolean
'**************************************************************
'Author: ZaMa
'Last Modify Date: 10/11/2009
'Returns true if the user is in an Arena
'**************************************************************
On Error GoTo ErrHandler
  
    IsArena = (TriggerZonaPelea(UserIndex, UserIndex) = TRIGGER6_PERMITE)
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function IsArena de Modulo_UsUaRiOs.bas")
End Function

Public Sub PerdioNpc(ByVal UserIndex As Integer, Optional ByVal CheckPets As Boolean = True)
'**************************************************************
'Author: ZaMa
'Last Modify Date: 11/07/2010 (ZaMa)
'The user loses his owned npc
'18/01/2010: ZaMa - Las mascotas dejan de atacar al npc que se perdió.
'11/07/2010: ZaMa - Coloco el indice correcto de las mascotas y ahora siguen al amo si existen.
'13/07/2010: ZaMa - Ahora solo dejan de atacar las mascotas si estan atacando al npc que pierde su amo.
'**************************************************************
On Error GoTo ErrHandler
  

    Dim PetCounter As Long
    Dim PetIndex As Integer
    Dim NpcIndex As Integer
    
    With UserList(UserIndex)
        
        NpcIndex = .flags.OwnedNpc
        If NpcIndex > 0 Then
            
            If CheckPets Then
                ' Dejan de atacar las mascotas
                If .TammedPetsCount > 0 Then
                    For PetCounter = 1 To Classes(.clase).ClassMods.MaxTammedPets
                    
                        ' Tammed pets
                        PetIndex = .TammedPets(PetCounter).NpcIndex
                        
                        If PetIndex > 0 Then
                            ' Si esta atacando al npc deja de hacerlo
                            If Npclist(PetIndex).TargetNPC = NpcIndex Then
                                Call FollowAmo(PetIndex)
                            End If
                        End If
                        
                    Next PetCounter
                End If
                
                If .InvokedPetsCount > 0 Then
                    For PetCounter = 1 To Classes(.clase).ClassMods.MaxInvokedPets
                    
                        ' Invoked pets
                        PetIndex = .InvokedPets(PetCounter).NpcIndex
                        
                        If PetIndex > 0 Then
                            ' Si esta atacando al npc deja de hacerlo
                            If Npclist(PetIndex).TargetNPC = NpcIndex Then
                                Call FollowAmo(PetIndex)
                            End If
                        End If
                        
                    Next PetCounter
                End If
            End If
            
            ' Reset flags
            Npclist(NpcIndex).Owner = 0
            .flags.OwnedNpc = 0

        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub PerdioNpc de Modulo_UsUaRiOs.bas")
End Sub

Public Sub ApropioNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
'**************************************************************
'Author: ZaMa
'Last Modify Date: 27/07/2010 (zaMa)
'The user owns a new npc
'18/01/2010: ZaMa - El sistema no aplica a zonas seguras.
'19/04/2010: ZaMa - Ahora los admins no se pueden apropiar de npcs.
'27/07/2010: ZaMa - El sistema no aplica a mapas seguros.
'**************************************************************
On Error GoTo ErrHandler
  

    With UserList(UserIndex)
        ' Los admins no se pueden apropiar de npcs
        If EsGm(UserIndex) Then Exit Sub
        
        Dim mapa As Integer
        mapa = .Pos.Map
        
        ' No aplica a triggers seguras
        If MapData(mapa, .Pos.X, .Pos.Y).Trigger = eTrigger.ZONASEGURA Then Exit Sub
        
        ' No se aplica a mapas seguros
        If MapInfo(mapa).Pk = False Then Exit Sub
        
        ' No aplica a algunos mapas que permiten el robo de npcs
        If MapInfo(mapa).RoboNpcsPermitido = 1 Then Exit Sub
        
        ' No se puede apropiar a bosses
        If Npclist(NpcIndex).flags.Boss > 0 Then Exit Sub
        
        ' Pierde el npc anterior
        If .flags.OwnedNpc > 0 Then Npclist(.flags.OwnedNpc).Owner = 0
        
        ' Si tenia otro dueño, lo perdio aca
        Npclist(NpcIndex).Owner = UserIndex
        .flags.OwnedNpc = NpcIndex
    End With
    
    ' Inicializo o actualizo el timer de pertenencia
    Call IntervaloPerdioNpc(UserIndex, True)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ApropioNpc de Modulo_UsUaRiOs.bas")
End Sub

Public Function FarthestPet(ByVal UserIndex As Integer) As Integer
'**************************************************************
'Author: ZaMa
'Last Modify Date: 18/11/2009
'Devuelve el indice de la mascota mas lejana.
'**************************************************************
On Error GoTo ErrHandler
    
    Dim PetIndex As Integer
    Dim Distancia As Integer
    Dim OtraDistancia As Integer
    
    With UserList(UserIndex)
        If .TammedPetsCount = 0 Then Exit Function
    
        For PetIndex = 1 To Classes(.clase).ClassMods.MaxTammedPets
            ' Solo pos invocar criaturas que exitan!
            If .TammedPets(PetIndex).NpcIndex > 0 Then
                ' Solo aplica a mascota, nada de elementales..
                If Npclist(.TammedPets(PetIndex).NpcIndex).Contadores.TiempoExistencia = 0 Then
                    If FarthestPet = 0 Then
                        ' Por si tiene 1 sola mascota
                        FarthestPet = PetIndex
                        Distancia = Abs(.Pos.X - Npclist(.TammedPets(PetIndex).NpcIndex).Pos.X) + _
                                    Abs(.Pos.Y - Npclist(.TammedPets(PetIndex).NpcIndex).Pos.Y)
                    Else
                        ' La distancia de la proxima mascota
                        OtraDistancia = Abs(.Pos.X - Npclist(.TammedPets(PetIndex).NpcIndex).Pos.X) + _
                                        Abs(.Pos.Y - Npclist(.TammedPets(PetIndex).NpcIndex).Pos.Y)
                        ' Esta mas lejos?
                        If OtraDistancia > Distancia Then
                            Distancia = OtraDistancia
                            FarthestPet = PetIndex
                        End If
                    End If
                End If
            End If
        Next PetIndex
    End With

    Exit Function
    
ErrHandler:
    Call LogError("Error en FarthestPet")
End Function

''
' Set the EluSkill value at the skill.
'
' @param UserIndex  Specifies reference to user
' @param Skill      Number of the skill to check
' @param Allocation True If the motive of the modification is the allocation, False if the skill increase by training

Public Sub CheckEluSkill(ByVal UserIndex As Integer, ByVal skill As Byte, ByVal Allocation As Boolean)
'*************************************************
'Author: Torres Patricio (Pato)
'Last modified: 11/20/2009
'
'*************************************************
On Error GoTo ErrHandler
  

With UserList(UserIndex).Stats
    If GetSkills(UserIndex, skill) < MAX_SKILL_POINTS Then
        If Allocation Then
            .ExpSkills(skill) = 0
        Else
            .ExpSkills(skill) = .ExpSkills(skill) - .EluSkills(skill)
        End If
        
        .EluSkills(skill) = ConstantesBalance.EluSkillInicial * 1.05 ^ GetSkills(UserIndex, skill)
    Else
        .ExpSkills(skill) = 0
        .EluSkills(skill) = 0
    End If
End With

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CheckEluSkill de Modulo_UsUaRiOs.bas")
End Sub


Public Sub CalculateNaturalAndAssignSkills(ByVal SkillQuantity As Byte, ByRef NaturalSkillPoints As Byte, ByRef AssignedSkillPoints As Byte)
On Error GoTo ErrHandler
    
    Dim TotalSkills As Byte
    Dim Diff As Byte

    If SkillQuantity > MAX_SKILL_POINTS Then
        SkillQuantity = MAX_SKILL_POINTS
    End If
    
    TotalSkills = NaturalSkillPoints + AssignedSkillPoints
    
    'check if I have to add skills
    If SkillQuantity > TotalSkills Then
        'add to assigned skills
        AssignedSkillPoints = (SkillQuantity - NaturalSkillPoints)
    Else
        'calculate skills point to remove from user
        Diff = TotalSkills - SkillQuantity
        
        'check if this difference can be taken from assigned skills
        If AssignedSkillPoints >= Diff Then
            AssignedSkillPoints = AssignedSkillPoints - Diff
        Else
            Diff = Diff - AssignedSkillPoints
            AssignedSkillPoints = 0
            
            'check if can be taken from natural skill points
            If Diff >= NaturalSkillPoints Then
                NaturalSkillPoints = 0
            Else
                NaturalSkillPoints = NaturalSkillPoints - Diff
            End If
        End If
    End If
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CalculateNaturalAndAssignSkills de Modulo_UsUaRiOs.bas")
End Sub

Public Function HasEnoughItems(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, ByVal Amount As Long) As Boolean
'**************************************************************
'Author: ZaMa
'Last Modify Date: 25/11/2009
'Cheks Wether the user has the required amount of items in the inventory or not
'**************************************************************
On Error GoTo ErrHandler
  

    Dim Slot As Long
    Dim ItemInvAmount As Long
    
    With UserList(UserIndex)
        For Slot = 1 To .CurrentInventorySlots
            ' Si es el item que busco
            If .Invent.Object(Slot).ObjIndex = ObjIndex Then
                ' Lo sumo a la cantidad total
                ItemInvAmount = ItemInvAmount + .Invent.Object(Slot).Amount
            End If
        Next Slot
    End With
    
    HasEnoughItems = Amount <= ItemInvAmount
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function HasEnoughItems de Modulo_UsUaRiOs.bas")
End Function

Public Function TotalOfferItems(ByVal ObjIndex As Integer, ByVal UserIndex As Integer) As Long
'**************************************************************
'Author: ZaMa
'Last Modify Date: 25/11/2009
'Cheks the amount of items the user has in offerSlots.
'**************************************************************
On Error GoTo ErrHandler
  
    Dim Slot As Byte
    
    For Slot = 1 To MAX_OFFER_SLOTS
            ' Si es el item que busco
        If UserList(UserIndex).ComUsu.Objeto(Slot) = ObjIndex Then
            ' Lo sumo a la cantidad total
            TotalOfferItems = TotalOfferItems + UserList(UserIndex).ComUsu.cant(Slot)
        End If
    Next Slot

  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function TotalOfferItems de Modulo_UsUaRiOs.bas")
End Function

Public Sub goHome(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Budi
    'Last Modification: 01/06/2010
    '01/06/2010: ZaMa - Ahora usa otro tipo de intervalo (lo saque de tPiquetec)
    '***************************************************
    On Error GoTo ErrHandler

    Dim Distance As Long

    With UserList(UserIndex)

        If .flags.Muerto <> 1 Then
            Call WriteConsoleMsg(UserIndex, "Debes estar muerto para poder utilizar este comando.", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub
        End If
        
        Call IntervaloGoHome(UserIndex, ConstantesBalance.HomeWaitingTime * 1000, True)

        If .flags.Navegando = 1 Then
            .Char.FX = AnimHogarNavegando(.Char.heading)
        Else
            .Char.FX = AnimHogar(.Char.heading)

        End If
        
        Call WriteMultiMessage(UserIndex, eMessages.Home, ConstantesBalance.HomeWaitingTime, , , MapInfo(Ciudades(.Hogar).Map).Name)
        
        .Char.Loops = INFINITE_LOOPS
        
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, .Char.FX, INFINITE_LOOPS, , 0))
  
    End With
  
    Exit Sub
  
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub goHome de Modulo_UsUaRiOs.bas")

End Sub

Public Function ToogleToAtackable(ByVal UserIndex As Integer, ByVal OwnerIndex As Integer, Optional ByVal StealingNpc As Boolean = True) As Boolean
'***************************************************
'Author: ZaMa
'Last Modification: 15/01/2010
'Change to Atackable mode.
'***************************************************
On Error GoTo ErrHandler
  
    
    Dim AtacablePor As Integer
    
    With UserList(UserIndex)
        
        If MapInfo(.Pos.Map).Pk = False Then
            Call WriteConsoleMsg(UserIndex, "No puedes robar npcs en zonas seguras.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        
        AtacablePor = .flags.AtacablePor
            
        If AtacablePor > 0 Then
            ' Intenta robar un npc
            If StealingNpc Then
                ' Puede atacar el mismo npc que ya estaba robando, pero no una nuevo.
                If AtacablePor <> OwnerIndex Then
                    Call WriteConsoleMsg(UserIndex, "No puedes atacar otra criatura con dueño hasta que haya terminado tu castigo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Function
                End If
            ' Esta atacando a alguien en estado atacable => Se renueva el timer de atacable
            Else
                ' Renovar el timer
                Call IntervaloEstadoAtacable(UserIndex, True)
                ToogleToAtackable = True
                Exit Function
            End If
        End If
        
        .flags.AtacablePor = OwnerIndex
    
        ' Actualizar clientes
        Call RefreshCharStatus(UserIndex, False)
        
        ' Inicializar el timer
        Call IntervaloEstadoAtacable(UserIndex, True)
        
        ToogleToAtackable = True
        
    End With
    
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function ToogleToAtackable de Modulo_UsUaRiOs.bas")
End Function

Public Sub setHome(ByVal UserIndex As Integer, ByVal newHome As eCiudad, ByVal NpcIndex As Integer)
'***************************************************
'Author: Budi
'Last Modification: 01/06/2010
'30/04/2010: ZaMa - Ahora el npc avisa que se cambio de hogar.
'01/06/2010: ZaMa - Ahora te avisa si ya tenes ese hogar.
'***************************************************
On Error GoTo ErrHandler
  
    If newHome < eCiudad.cUllathorpe Or newHome > eCiudad.cLastCity - 1 Then Exit Sub
    
    If UserList(UserIndex).Hogar <> newHome Then
        UserList(UserIndex).Hogar = newHome
    
        Call WriteChatOverHead(UserIndex, "¡¡¡Bienvenido a nuestra humilde comunidad, este es ahora tu nuevo hogar!!!", Npclist(NpcIndex).Char.CharIndex, vbWhite)
    Else
        Call WriteChatOverHead(UserIndex, "¡¡¡Ya eres miembro de nuestra humilde comunidad!!!", Npclist(NpcIndex).Char.CharIndex, vbWhite)
    End If

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub setHome de Modulo_UsUaRiOs.bas")
End Sub

Public Function GetHomeArrivalTime(ByVal UserIndex As Integer) As Long
'**************************************************************
'Author: ZaMa
'Last Modify by: ZaMa
'Last Modify Date: 01/06/2010
'Calculates the time left to arrive home.
'**************************************************************
On Error GoTo ErrHandler
  
    Dim TActual As Long

    TActual = GetTickCount()
    
    With UserList(UserIndex)
        GetHomeArrivalTime = getInterval(.Counters.goHome, TActual) * 0.001
    End With

  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetHomeArrivalTime de Modulo_UsUaRiOs.bas")
End Function

Public Sub HomeArrival(ByVal UserIndex As Integer)
'**************************************************************
'Author: ZaMa
'Last Modify by: ZaMa
'Last Modify Date: 01/06/2010
'Teleports user to its home.
'**************************************************************
On Error GoTo ErrHandler
  
    
    Dim tX As Integer
    Dim tY As Integer
    Dim tMap As Integer

    With UserList(UserIndex)

        'Antes de que el pj llegue a la ciudad, lo hacemos dejar de navegar para que no se buguee.
        If .flags.Navegando = 1 Then
            .Char.body = ConstantesGRH.CuerpoMuerto
            .Char.head = ConstantesGRH.CabezaMuerto
            .Char.ShieldAnim = ConstantesGRH.NingunEscudo
            .Char.WeaponAnim = ConstantesGRH.NingunArma
            .Char.CascoAnim = ConstantesGRH.NingunCasco
            
            .flags.Navegando = 0
            
            Call WriteNavigateChange(UserIndex, False)
            'Le sacamos el navegando, pero no le mostramos a los demás porque va a ser sumoneado hasta ulla.
        End If
        
        tX = Ciudades(.Hogar).X
        tY = Ciudades(.Hogar).Y
        tMap = Ciudades(.Hogar).Map
        
        Call FindLegalPos(UserIndex, tMap, tX, tY)
        Call WarpUserChar(UserIndex, tMap, tX, tY, True)
        
        Call WriteMultiMessage(UserIndex, eMessages.FinishHome)
        
        Call EndTravel(UserIndex, False)
        
    End With
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HomeArrival de Modulo_UsUaRiOs.bas")
End Sub

Public Sub EndTravel(ByVal UserIndex As Integer, ByVal Cancelado As Boolean)
'**************************************************************
'Author: ZaMa
'Last Modify Date: 11/06/2011
'Ends travel.
'**************************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)
        .Counters.goHome = 0
        .flags.Traveling = 0
       
        Call WriteMultiMessage(UserIndex, eMessages.CancelHome, Cancelado)
        
        .Char.FX = 0
        .Char.Loops = 0
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0, , 0))
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub EndTravel de Modulo_UsUaRiOs.bas")
End Sub

Public Sub EndMimic(ByVal UserIndex As Integer, ByVal ValidateSaling As Boolean, ByVal ShowMessage As Boolean)
'**************************************************************
'Author: ZaMa
'Last Modify Date: 25/11/2014 (D'Artagnan)
'Ends user mimic and returns char to original
'21/02/2014: D'Artagnan - Restore nickname.
'**************************************************************
On Error GoTo ErrHandler
  

    With UserList(UserIndex)
        
        If ValidateSaling And (.flags.Navegando = 1) Then
            Call ToggleBoatBody(UserIndex)
        Else
            .Char.body = .OrigChar.body
            .Char.head = .OrigChar.head
            .Char.CascoAnim = .OrigChar.CascoAnim
            .Char.ShieldAnim = .OrigChar.ShieldAnim
            .Char.WeaponAnim = .OrigChar.WeaponAnim
        End If
        
        
        .flags.Mimetizado = 0
        .Counters.Mimetismo = 0
        .flags.MimetizadoType = 0
        .flags.Ignorado = False
        
        ' Restore nickname.
        If Not .ShowName Then
            .ShowName = True
            Call RefreshCharStatus(UserIndex, False)
        End If
    End With

    If ShowMessage Then
        Call WriteConsoleMsg(UserIndex, "Recuperas tu apariencia normal.", FontTypeNames.FONTTYPE_INFO)
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub EndMimic de Modulo_UsUaRiOs.bas")
End Sub

Public Function FreeInventorySlots(ByVal UserIndex As Integer) As Integer
'**************************************************************
'Author: ZaMa
'Last Modify Date: 09/01/2011
'Returns inventory free slots.
'**************************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)
        FreeInventorySlots = .CurrentInventorySlots - .Invent.NroItems
    End With
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function FreeInventorySlots de Modulo_UsUaRiOs.bas")
End Function

Public Function GetPromedioReputacion(ByVal NobleRep As Long, _
    ByVal BurguesRep As Long, _
    ByVal PlebeRep As Long, _
    ByVal LadronesRep As Long, _
    ByVal BandidoRep As Long, _
    ByVal AsesinoRep As Long) As Long
'**************************************************************
'Author: ZaMa
'Last Modify Date: 17/01/2014
'Returns Reputation average.
'**************************************************************
On Error GoTo ErrHandler
  
    
    Dim lPromedio As Long
    lPromedio = (-AsesinoRep) + _
            (-BandidoRep) + _
            BurguesRep + _
            (-LadronesRep) + _
            NobleRep + _
            PlebeRep
            
    GetPromedioReputacion = Round(lPromedio / 6)
    
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetPromedioReputacion de Modulo_UsUaRiOs.bas")
End Function

Public Function CommerceAllowed(ByVal nUserIndex As Integer) As Boolean
'***************************************************
'Author: D'Artagnan
'Last Modification: 28/10/2014
'Return True if the specified user can commerce. False otherwise.
'***************************************************
On Error GoTo ErrHandler
  
    Dim nTargetUserIndex As Integer
    
    With UserList(nUserIndex)
        'Dead people can't commerce.
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(nUserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        ' Cancel previous opertaion.
        ElseIf isTradingWithUser(nUserIndex) Then
            nTargetUserIndex = getTradingUser(nUserIndex)
                      
            If nTargetUserIndex > -1 Then
                If Not UserList(nTargetUserIndex).ComUsu.Acepto Then
                    If .flags.TargetUser = nTargetUserIndex Then
                        ' Same target.
                        Call WriteConsoleMsg(nUserIndex, "Ya has enviado una solicitud de comercio.", _
                                             FontTypeNames.FONTTYPE_TALK)
                        Exit Function
                    Else
                        Call WriteConsoleMsg(nUserIndex, _
                            IIf( _
                                LenB(UserList(nTargetUserIndex).Name) > 0, _
                                "Has cancelado la operación con " & UserList(nTargetUserIndex).Name & ".", _
                                "Has cancelado la operación." _
                            ), _
                            FontTypeNames.FONTTYPE_TALK)
                                             
                        Call WriteConsoleMsg(nTargetUserIndex, .Name & " ha cancelado la operación.", _
                                             FontTypeNames.FONTTYPE_TALK)
                        
                        ' Clean up both users.
                        Call FinComerciarUsu(nUserIndex)
                        Call FinComerciarUsu(nTargetUserIndex, True)
                    End If
                End If
            End If
            
        ' Can't commerce while sailing.
        ElseIf .flags.Navegando = 1 Then
            Call WriteConsoleMsg(nUserIndex, "¡Estás navegando!", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        ' Already trading.
        ElseIf isTrading(nUserIndex) Then
            Call WriteConsoleMsg(nUserIndex, "Ya estás comerciando.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        
        If .flags.DueloIndex > 0 Then
            Call WriteConsoleMsg(nUserIndex, "No se puede comerciar durante un duelo.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
    End With
    
    CommerceAllowed = True
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function CommerceAllowed de Modulo_UsUaRiOs.bas")
End Function


Public Function UserAreaHasWater(ByVal nUserIndex As Integer) As Boolean
    UserAreaHasWater = IsValidUserPositionArea(nUserIndex, True, False)
End Function

Public Function UserAreaHasLand(ByVal nUserIndex As Integer) As Boolean
    UserAreaHasLand = IsValidUserPositionArea(nUserIndex, False, True)
End Function

Private Function IsValidUserPositionArea(ByVal nUserIndex As Integer, ByVal AllowWater As Boolean, ByVal AllowLand As Boolean) As Boolean
On Error GoTo ErrHandler

    With UserList(nUserIndex)
        IsValidUserPositionArea = RhombLegalPos(.Pos.Map, .Pos.X, .Pos.Y, 1, AllowWater, AllowLand, UpdatePos:=False)
    End With
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function IsValidUserPositionArea de Modulo_UsUaRiOs.bas")
End Function

Public Function GetUsersCount(Optional ByVal bIncludePremium As Boolean = True) As Long
'***************************************************
'Author: D'Artagnan (taken from HandleOnline)
'Last Modification: 29/04/2015
'Included bIncludePremium parameter.
'***************************************************
On Error GoTo ErrHandler
  
    Dim I As Long
    
    GetUsersCount = 0
    
    For I = 1 To LastUser
        If LenB(UserList(I).Name) <> 0 Then
            If IIf(bIncludePremium, True, Not UserList(I).bIsPremium) And UserList(I).flags.Privilegios And _
                (PlayerType.User Or PlayerType.Consejero) Then _
                GetUsersCount = GetUsersCount + 1
        End If
    Next I
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetUsersCount de Modulo_UsUaRiOs.bas")
End Function

Public Function GetStartingHealth(ByVal Class As Byte, ByVal race As Byte) As Integer
    GetStartingHealth = Classes(Class).ClassMods.StartingHealth + Classes(Class).RaceMods(race).StartingHealth
End Function

Public Function GetStartingMana(ByVal UserIndex As Integer) As Integer
    With UserList(UserIndex)
        GetStartingMana = Fix(CSng(CSng(.Stats.UserAtributos(eAtributos.Inteligencia)) * Classes(.clase).ClassMods.ManaStarterMultiplier))
    End With
End Function

Public Function RecalculateCharacterMaxMana(ByVal UserIndex As Integer) As Integer
On Error GoTo ErrHandler
    Dim AccumulatedMana As Integer
    Dim I As Integer
    Dim BaseHealth As Integer

    With UserList(UserIndex)
    
        ' Starting mana
        RecalculateCharacterMaxMana = GetStartingMana(UserIndex)
        
        ' Mana per level
        RecalculateCharacterMaxMana = RecalculateCharacterMaxMana + (Fix(CSng(.Stats.UserAtributos(eAtributos.Inteligencia) * Classes(.clase).ClassMods.ManaPerLevelMultiplier)) * (.Stats.ELV - 1))
        
        ' Add the extra mana based on masteries and Boosts
        RecalculateCharacterMaxMana = RecalculateCharacterMaxMana + Porcentaje(RecalculateCharacterMaxMana, .Masteries.Boosts.AddMaxManaPerc)
        RecalculateCharacterMaxMana = RecalculateCharacterMaxMana + .Masteries.Boosts.AddMaxMana
                    
    End With
    
    Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function RecalculateCharacterMaxMana de Modulo_UsUaRiOs.bas")
End Function

Public Function RecalculateCharacterMaxHealth(ByVal UserIndex As Integer) As Integer
On Error GoTo ErrHandler
    Dim AccumulatedHp As Integer
    Dim I As Integer
    Dim BaseHealth As Integer
    With UserList(UserIndex)
        
        ' Starting health
        RecalculateCharacterMaxHealth = GetStartingHealth(.clase, .raza)
        
        ' Health per level
        RecalculateCharacterMaxHealth = RecalculateCharacterMaxHealth + RandomNumber(Classes(.clase).RaceMods(.raza).HealthPerLevelMin, Classes(.clase).RaceMods(.raza).HealthPerLevelMax) * (.Stats.ELV - 1)
        
        ' Add extra health based on different levels
        For I = 1 To .Stats.ELV
            ' Get the extra health given on certain extra levels
             If Classes(.clase).RaceMods(.raza).ExtraHealthAtLevel(I) > 0 Then
                RecalculateCharacterMaxHealth = RecalculateCharacterMaxHealth + Classes(.clase).RaceMods(.raza).ExtraHealthAtLevel(I)
            End If
        Next I
        
        ' Add the extra health based on masteries
        RecalculateCharacterMaxHealth = RecalculateCharacterMaxHealth + .Masteries.Boosts.AddMaxHealth
    
    End With

    
    Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function RecalculateCharacterMaxHealth de Modulo_UsUaRiOs.bas")
End Function

Public Sub RecalculateUserAttributes(ByVal UserIndex As Integer)
On Error GoTo ErrHandler

    With UserList(UserIndex)
        .Stats.UserAtributos(eAtributos.Agilidad) = ModRaza(.raza).Agilidad
        .Stats.UserAtributos(eAtributos.Carisma) = ModRaza(.raza).Carisma
        .Stats.UserAtributos(eAtributos.Constitucion) = ModRaza(.raza).Constitucion
        .Stats.UserAtributos(eAtributos.Fuerza) = ModRaza(.raza).Fuerza
        .Stats.UserAtributos(eAtributos.Inteligencia) = ModRaza(.raza).Inteligencia
    End With
    
    Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function RecalculateUserAttributes de Modulo_UsUaRiOs.bas")
End Sub
