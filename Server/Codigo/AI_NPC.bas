Attribute VB_Name = "AI"
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

Public Enum TipoAI
    ESTATICO = 1
    MueveAlAzar = 2
    NpcMaloAtacaUsersBuenos = 3
    NPCDEFENSA = 4
    GuardiasAtacanCriminales = 5
    NpcObjeto = 6
    SigueAmo = 8
    NpcAtacaNpc = 9
    NpcPathfinding = 10
    
    'Pretorianos
    SacerdotePretorianoAi = 20
    GuerreroPretorianoAi = 21
    MagoPretorianoAi = 22
    CazadorPretorianoAi = 23
    ReyPretoriano = 24
    
    'MiniBosses
    BossDM = 30
    BossDV = 31
    BossDI = 32
    BossDA = 33
    BossDATentaculo = 34
End Enum

'Damos a los NPCs el mismo rango de visión que un PJ
Public Const RANGO_VISION_X As Byte = 8
Public Const RANGO_VISION_Y As Byte = 6

'****************************************************************************************************************
'****************************************************************************************************************
'****************************************************************************************************************
'                        Modulo AI_NPC
'****************************************************************************************************************
'****************************************************************************************************************
'****************************************************************************************************************
'AI de los NPC
'****************************************************************************************************************
'****************************************************************************************************************
'****************************************************************************************************************

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''Nuevas IAs''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub BossDATentaculoAttack(ByVal NpcIndex As Integer)
'---------------------------------------------------------------------------------------
' Module    : AI
' Author    : Anagrama
' Date      : 03/09/2016
' Purpose   : Ataque del tentaculo invocado del boss de DA (Este no se mueve).
'---------------------------------------------------------------------------------------
On Error GoTo ErrHandler
  
    Dim nPos As WorldPos
    Dim headingloop As Byte
    Dim FirstHeading As Byte
    Dim I As Long
    Dim UI As Integer
    Dim NPCI As Integer
    Dim UserTarget(1 To 4) As Integer
    Dim NpcTarget(1 To 4) As Integer
    Dim CurrentUser As Byte
    Dim CurrentNpc As Byte
    Dim UserProtected As Boolean

    With Npclist(NpcIndex)
        
        If Not .Timers.Check(TimersIndex.Hit, False) Then Exit Sub
        
        For headingloop = eHeading.NORTH To eHeading.WEST
            nPos = .Pos
            If .flags.Inmovilizado = 0 Or .Char.heading = headingloop Then
                Call HeadtoPos(headingloop, nPos)
                If InMapBounds(nPos.Map, nPos.X, nPos.Y) Then
                    UI = MapData(nPos.Map, nPos.X, nPos.Y).UserIndex
                    NPCI = MapData(nPos.Map, nPos.X, nPos.Y).NpcIndex
                    If UI > 0 Then
                        UserProtected = Not IntervaloPermiteSerAtacado(UI) And UserList(UI).flags.NoPuedeSerAtacado
                        UserProtected = UserProtected Or UserList(UI).flags.HelpMode
                        
                        If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.AdminPerseguible And (Not UserProtected) Then
                            If CurrentUser = 0 Then FirstHeading = headingloop
                            CurrentUser = CurrentUser + 1
                            UserTarget(CurrentUser) = UI
                        End If
                    ElseIf NPCI > 0 Then
                        If Npclist(NPCI).MaestroUser > 0 And Npclist(NPCI).flags.Paralizado = 0 Then
                            CurrentNpc = CurrentNpc + 1
                            NpcTarget(CurrentNpc) = NPCI
                        End If
                    End If
                End If
            End If  'inmo
        Next headingloop
        
        If CurrentUser Then
            Call ChangeNPCChar(NpcIndex, .Char.body, .Char.head, FirstHeading)
            For I = 1 To CurrentUser
                Call NpcLanzaSpellSobreUser(NpcIndex, UserTarget(CurrentUser), ConstantesBosses.BossDASpellTentaculo, , True)
                Call NpcAtacaUser(NpcIndex, UserTarget(I))
            Next I
            Call .Timers.Restart(TimersIndex.Hit)
        End If
        
        If CurrentNpc Then
            For I = 1 To CurrentNpc
                Call SistemaCombate.NpcAtacaNpc(NpcIndex, NpcTarget(I), False)
            Next I
            Call .Timers.Restart(TimersIndex.Hit)
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub BossDATentaculoAttack de AI_NPC.bas")
End Sub

Private Sub BossDAAttack(ByVal NpcIndex As Integer)
'---------------------------------------------------------------------------------------
' Module    : AI
' Author    : Anagrama
' Date      : 03/09/2016
' Purpose   : Ataque del boss de DA (Este no se mueve).
'---------------------------------------------------------------------------------------
On Error GoTo ErrHandler
  
    Dim nPos As WorldPos
    Dim headingloop As Byte
    Dim FirstHeading As Byte
    Dim I As Long
    Dim UI As Integer
    Dim NPCI As Integer
    Dim AtraerTarget() As Integer
    Dim TargetsAtraidos As Byte
    Dim TentaculoTarget As Integer
    Dim AplastarTarget As Integer
    Dim UserProtected As Boolean
    Dim X As Byte
    Dim Y As Byte
    
    With Npclist(NpcIndex)
        
        If Not .Timers.Check(TimersIndex.Hit) Then Exit Sub
        
        ReDim AtraerTarget(1 To 1) As Integer
        
        Dim Query() As Collision.UUID
        For I = 0 To ModAreas.QueryEntities(NpcIndex, ENTITY_TYPE_NPC, Query, ENTITY_TYPE_PLAYER, eGridSortClosest)
            UI = Query(I).Name

                    UserProtected = Not IntervaloPermiteSerAtacado(UI) And UserList(UI).flags.NoPuedeSerAtacado
                    UserProtected = UserProtected Or UserList(UI).flags.HelpMode
                    
                    If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.AdminPerseguible And Not UserProtected Then
                        If Distancia(.Pos, UserList(UI).Pos) > 1 Then
                        TargetsAtraidos = TargetsAtraidos + 1
                        ReDim Preserve AtraerTarget(1 To TargetsAtraidos) As Integer
                        AtraerTarget(TargetsAtraidos) = UI
                        End If
                        
                        If (RandomNumber(1, UBound(Query)) = 1 Or _
                            I = UBound(Query)) And TentaculoTarget = 0 Then
                            
                            TentaculoTarget = UI
                        End If
                        
                        If UserList(UI).flags.Paralizado Or UserList(UI).flags.Inmovilizado Then
                            If RandomNumber(1, 100) <= ConstantesBosses.BossDACastTorrente Then
                                Call NpcLanzaSpellSobreUser(NpcIndex, UI, ConstantesBosses.BossDASpellTorrente, , True)
                            End If
                        End If
                    End If

        Next I

        If RandomNumber(1, 100) <= ConstantesBosses.BossDACastAtraer Then
            If TargetsAtraidos > 0 Then
                For I = 1 To UBound(AtraerTarget)
                    Call NpcLanzaSpellSobreUser(NpcIndex, AtraerTarget(I), ConstantesBosses.BossDASpellAtraer, , True)
                Next I
            End If
        End If
        
        If RandomNumber(1, 100) <= ConstantesBosses.BossDACastTentaculo Then
            If TentaculoTarget > 0 Then
                Call DoNpcInvocacion(NpcIndex, UserList(TentaculoTarget).Pos)
            End If
        End If
        
        For headingloop = eHeading.NORTH To eHeading.WEST
            nPos = .Pos
            If .flags.Inmovilizado = 0 Or .Char.heading = headingloop Then
                Call HeadtoPos(headingloop, nPos)
                If InMapBounds(nPos.Map, nPos.X, nPos.Y) Then
                    UI = MapData(nPos.Map, nPos.X, nPos.Y).UserIndex
                    NPCI = MapData(nPos.Map, nPos.X, nPos.Y).NpcIndex
                    If UI > 0 Then
                        UserProtected = Not IntervaloPermiteSerAtacado(UI) And UserList(UI).flags.NoPuedeSerAtacado
                        UserProtected = UserProtected Or UserList(UI).flags.HelpMode
                        
                        If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.AdminPerseguible And (Not UserProtected) And _
                            AplastarTarget = 0 Then
                            FirstHeading = headingloop
                            Call ChangeNPCChar(NpcIndex, .Char.body, .Char.head, FirstHeading)
                            AplastarTarget = UI
                        End If
                    ElseIf NPCI > 0 Then
                        If Npclist(NPCI).MaestroUser > 0 And Npclist(NPCI).flags.Paralizado = 0 Then
                            Call SistemaCombate.NpcAtacaNpc(NpcIndex, NPCI, False)
                        End If
                    End If
                End If
            End If  'inmo
        Next headingloop

        If RandomNumber(1, 100) <= ConstantesBosses.BossDACastAplastar Then
            If AplastarTarget > 0 Then
                For X = UserList(AplastarTarget).Pos.X - ConstantesBosses.BossDAAplastarArea _
                        To UserList(AplastarTarget).Pos.X + ConstantesBosses.BossDAAplastarArea
                    For Y = UserList(AplastarTarget).Pos.Y - ConstantesBosses.BossDAAplastarArea _
                        To UserList(AplastarTarget).Pos.Y + ConstantesBosses.BossDAAplastarArea
                        
                        UI = MapData(.Pos.Map, X, Y).UserIndex
                        
                        If UI > 0 Then
                            UserProtected = Not IntervaloPermiteSerAtacado(UI) And UserList(UI).flags.NoPuedeSerAtacado
                            UserProtected = UserProtected Or UserList(UI).flags.HelpMode
                        
                            If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.AdminPerseguible And (Not UserProtected) Then
                                Call NpcAtacaUser(NpcIndex, UI)
                            End If
                        End If
                    Next Y
                Next X
            End If
        End If

    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub BossDAAttack de AI_NPC.bas")
End Sub

Private Sub BossDIAttack(ByVal NpcIndex As Integer)
'---------------------------------------------------------------------------------------
' Module    : AI
' Author    : Anagrama
' Date      : 03/09/2016
' Purpose   : Ataque del boss de DI.
'---------------------------------------------------------------------------------------
On Error GoTo ErrHandler
  
    Dim nPos As WorldPos
    Dim headingloop As Byte
    Dim I As Long
    Dim UI As Integer
    Dim NPCI As Integer
    Dim DebuffTarget As Integer
    Dim SpellTarget As Integer
    Dim UserProtected As Boolean
    Dim Spell As Byte
    
    With Npclist(NpcIndex)
        
        If Not .Timers.Check(TimersIndex.Hit) Then Exit Sub
        
        Dim Query() As Collision.UUID
        For I = 0 To ModAreas.QueryEntities(NpcIndex, ENTITY_TYPE_NPC, Query, ENTITY_TYPE_PLAYER, eGridSortClosest)
            UI = Query(I).Name

                    UserProtected = Not IntervaloPermiteSerAtacado(UI) And UserList(UI).flags.NoPuedeSerAtacado
                    UserProtected = UserProtected Or UserList(UI).flags.HelpMode
                    
                    If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.AdminPerseguible And Not UserProtected Then
                        If RandomNumber(1, UBound(Query)) = 1 Or _
                            (I = UBound(Query) And DebuffTarget = 0) Then
                                
                            DebuffTarget = UI
                        End If
                        If RandomNumber(1, UBound(Query)) = 1 Or _
                            I = UBound(Query) Then
                                
                            If UserList(UI).flags.Paralizado > 0 Or UserList(UI).flags.Inmovilizado > 0 Or _
                                ((I = UBound(Query) Or _
                                RandomNumber(1, UBound(Query)) = 1) And SpellTarget = 0) Then
                                SpellTarget = UI
                            End If
                        End If
                    End If

        Next I
        
        If .Stats.MinHp >= (.Stats.MaxHp - (.Stats.MaxHp / 3 * 2)) Then
            If DebuffTarget > 0 Then
                Spell = RandomNumber(1, ConstantesBosses.BossDINumDebuff)
                Call NpcLanzaSpellSobreUser(NpcIndex, DebuffTarget, ConstantesBosses.BossDISpellDebuff(Spell), , True)
            End If
        End If
        
        If .Stats.MinHp >= (.Stats.MaxHp - .Stats.MaxHp / 3) Then
            If SpellTarget > 0 Then
                Call NpcLanzaSpellSobreUser(NpcIndex, SpellTarget, ConstantesBosses.BossDISpellBola, , True)
            End If
        End If
        
        For headingloop = eHeading.NORTH To eHeading.WEST
            nPos = .Pos
            If .flags.Inmovilizado = 0 Or .Char.heading = headingloop Then
                Call HeadtoPos(headingloop, nPos)
                If InMapBounds(nPos.Map, nPos.X, nPos.Y) Then
                    UI = MapData(nPos.Map, nPos.X, nPos.Y).UserIndex
                    NPCI = MapData(nPos.Map, nPos.X, nPos.Y).NpcIndex
                    If UI > 0 Then
                        UserProtected = Not IntervaloPermiteSerAtacado(UI) And UserList(UI).flags.NoPuedeSerAtacado
                        UserProtected = UserProtected Or UserList(UI).flags.HelpMode
                        
                        If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.AdminPerseguible And (Not UserProtected) Then
                            Call ChangeNPCChar(NpcIndex, .Char.body, .Char.head, headingloop)
                            Call NpcAtacaUser(NpcIndex, UI)
                        End If
                    ElseIf NPCI > 0 Then
                        If Npclist(NPCI).MaestroUser > 0 And Npclist(NPCI).flags.Paralizado = 0 Then
                            Call ChangeNPCChar(NpcIndex, .Char.body, .Char.head, headingloop)
                            Call SistemaCombate.NpcAtacaNpc(NpcIndex, NPCI, False)
                        End If
                    End If
                End If
            End If  'inmo
        Next headingloop
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub BossDIAttack de AI_NPC.bas")
End Sub

Private Sub BossDIMovement(ByVal NpcIndex As Integer)
'---------------------------------------------------------------------------------------
' Module    : AI
' Author    : Anagrama
' Date      : 03/09/2016
' Purpose   : Movimiento del boss de DI.
'---------------------------------------------------------------------------------------
On Error GoTo ErrHandler
  
    Dim UserIndex As Long
    Dim I As Long
    Dim UserProtected As Boolean
    Dim ClosestUser As Integer
    Dim tmpDistance As Byte
    Dim ClosestDistance As Byte
    
    With Npclist(NpcIndex)
        
        Dim Query() As Collision.UUID
        For I = 0 To ModAreas.QueryEntities(NpcIndex, ENTITY_TYPE_NPC, Query, ENTITY_TYPE_PLAYER, eGridSortClosest)
            UserIndex = Query(I).Name

                    UserProtected = Not IntervaloPermiteSerAtacado(UserIndex) And UserList(UserIndex).flags.NoPuedeSerAtacado
                    UserProtected = UserProtected Or UserList(UserIndex).flags.HelpMode
                    
                    If UserList(UserIndex).flags.Muerto = 0 And UserList(UserIndex).flags.AdminPerseguible And Not UserProtected Then
                        
                        tmpDistance = Distance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, .Pos.X, .Pos.Y)
                        If ClosestDistance = 0 Or tmpDistance < ClosestDistance Then
                            ClosestDistance = tmpDistance
                            ClosestUser = UserIndex
                        End If
                    End If

        Next I
        
        If ClosestUser Then
            Call GeneralPathFinder(NpcIndex, ClosestUser)
            Exit Sub
        End If
        
        'Si llega aca es que no había ningún usuario cercano vivo.
        'A bailar. Pablo (ToxicWaste)
        If RandomNumber(0, 10) = 0 Then
            Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
        End If
    End With

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub BossDIMovement de AI_NPC.bas")
End Sub

Private Sub BossDVAttack(ByVal NpcIndex As Integer)
'---------------------------------------------------------------------------------------
' Module    : AI
' Author    : Anagrama
' Date      : 03/09/2016
' Purpose   : Ataque del boss de DV.
'---------------------------------------------------------------------------------------
On Error GoTo ErrHandler
  
    Dim nPos As WorldPos
    Dim headingloop As Byte
    Dim NewHeading As Byte
    Dim I As Long
    Dim UI As Integer
    Dim NPCI As Integer
    Dim UserTarget(1 To 4) As Integer
    Dim NpcTarget(1 To 4) As Integer
    Dim CurrentUser As Byte
    Dim CurrentNpc As Byte
    Dim UserProtected As Boolean
    
    With Npclist(NpcIndex)
        
        If Not .Timers.Check(TimersIndex.Hit, False) Then Exit Sub
        
        If .Target = 0 Then
            Exit Sub
        Else
            If (UserList(.Target).flags.UserLogged = False) Or _
                (UserList(.Target).flags.Muerto = 1) Or _
                (.Pos.Map <> UserList(.Target).Pos.Map) Or _
                (UserList(.Target).flags.HelpMode = True) Or _
                (Abs(UserList(.Target).Pos.X - .Pos.X) > RANGO_VISION_X + 2) Or _
                (Abs(UserList(.Target).Pos.Y - .Pos.Y) > RANGO_VISION_Y + 2) Then
                    .Target = 0
                    Exit Sub
            End If
        End If
        
        If .Pos.X < UserList(.Target).Pos.X Then
            NewHeading = eHeading.EAST
        ElseIf .Pos.X > UserList(.Target).Pos.X Then
            NewHeading = eHeading.WEST
        ElseIf .Pos.Y < UserList(.Target).Pos.Y Then
            NewHeading = eHeading.NORTH
        ElseIf .Pos.Y > UserList(.Target).Pos.Y Then
            NewHeading = eHeading.SOUTH
        End If
        
        Call ChangeNPCChar(NpcIndex, .Char.body, .Char.head, NewHeading)
        
        If RandomNumber(1, 100) <= ConstantesBosses.BossDVCastDescarga Then
            Call NpcLanzaSpellSobreUser(NpcIndex, .Target, ConstantesBosses.BossDVSpellDescarga, , True)
        End If

        If RandomNumber(1, 100) <= ConstantesBosses.BossDVCastTormenta Then
            Call NpcLanzaSpellSobreUser(NpcIndex, .Target, ConstantesBosses.BossDVSpellTormenta, , True)
        End If
        
        If RandomNumber(1, 100) <= ConstantesBosses.BossDVCastPetrificar Then
            
            Dim Query() As Collision.UUID
            For I = 0 To ModAreas.QueryEntities(NpcIndex, ENTITY_TYPE_NPC, Query, ENTITY_TYPE_PLAYER, eGridSortClosest)
                UI = Query(I).Name

                        UserProtected = Not IntervaloPermiteSerAtacado(UI) And UserList(UI).flags.NoPuedeSerAtacado
                        UserProtected = UserProtected Or UserList(UI).flags.HelpMode
                        
                        If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.AdminPerseguible And Not UserProtected Then
                            If (.Char.heading = eHeading.EAST And UserList(UI).Pos.X > .Pos.X) Or _
                                (.Char.heading = eHeading.WEST And UserList(UI).Pos.X < .Pos.X) Or _
                                (.Char.heading = eHeading.SOUTH And UserList(UI).Pos.Y > .Pos.Y) Or _
                                (.Char.heading = eHeading.NORTH And UserList(UI).Pos.Y < .Pos.Y) Then
                                
                                Call NpcLanzaSpellSobreUser(NpcIndex, UI, ConstantesBosses.BossDVSpellPetrificar, , True)
                            End If
                        End If

            Next I
        End If
        
        For headingloop = eHeading.NORTH To eHeading.WEST
            nPos = .Pos
            If .flags.Inmovilizado = 0 Or .Char.heading = headingloop Then
                Call HeadtoPos(headingloop, nPos)
                If InMapBounds(nPos.Map, nPos.X, nPos.Y) Then
                    UI = MapData(nPos.Map, nPos.X, nPos.Y).UserIndex
                    NPCI = MapData(nPos.Map, nPos.X, nPos.Y).NpcIndex
                    If UI > 0 Then
                        UserProtected = Not IntervaloPermiteSerAtacado(UI) And UserList(UI).flags.NoPuedeSerAtacado
                        UserProtected = UserProtected Or UserList(UI).flags.HelpMode
                        
                        If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.AdminPerseguible And (Not UserProtected) Then
                            CurrentUser = CurrentUser + 1
                            UserTarget(CurrentUser) = UI
                        End If
                    ElseIf NPCI > 0 Then
                        If Npclist(NPCI).MaestroUser > 0 And Npclist(NPCI).flags.Paralizado = 0 Then
                            CurrentNpc = CurrentNpc + 1
                            NpcTarget(CurrentNpc) = NPCI
                        End If
                    End If
                End If
            End If  'inmo
        Next headingloop
        
        If CurrentUser Then
            For I = 1 To CurrentUser
                Call NpcAtacaUser(NpcIndex, UserTarget(I))
            Next I
        End If
        
        If CurrentNpc Then
            For I = 1 To CurrentNpc
                Call SistemaCombate.NpcAtacaNpc(NpcIndex, NpcTarget(I), False)
            Next I
        End If
        
        Call .Timers.Restart(TimersIndex.Hit)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub BossDVAttack de AI_NPC.bas")
End Sub

Private Sub BossDVMovement(ByVal NpcIndex As Integer)
'---------------------------------------------------------------------------------------
' Module    : AI
' Author    : Anagrama
' Date      : 03/09/2016
' Purpose   : Movimiento del boss de DV.
'---------------------------------------------------------------------------------------
On Error GoTo ErrHandler
  
    Dim UserIndex As Integer
    Dim I As Long
    Dim UserProtected As Boolean
    
    With Npclist(NpcIndex)
        
        If .Target > 0 Then 'Me fijo si el target sigue siendo valido
            If (UserList(.Target).flags.UserLogged = False) Or _
                (UserList(.Target).flags.Muerto = 1) Or _
                (.Pos.Map <> UserList(.Target).Pos.Map) Or _
                (UserList(.Target).flags.HelpMode = True) Or _
                (Abs(UserList(.Target).Pos.X - .Pos.X) > RANGO_VISION_X + 2) Or _
                (Distancia(.Pos, UserList(.Target).Pos) > ConstantesBosses.BossDVDistance) Or _
                (Abs(UserList(.Target).Pos.Y - .Pos.Y) > RANGO_VISION_Y + 2) Then
                    
                .Target = 0
            End If
        End If
        
        If .Target = 0 Then 'Si no tengo target busco uno
            
            Dim Query() As Collision.UUID
            For I = 0 To ModAreas.QueryEntities(NpcIndex, ENTITY_TYPE_NPC, Query, ENTITY_TYPE_PLAYER, eGridSortClosest)
                UserIndex = Query(I).Name

                        UserProtected = Not IntervaloPermiteSerAtacado(UserIndex) And UserList(UserIndex).flags.NoPuedeSerAtacado
                        UserProtected = UserProtected Or UserList(UserIndex).flags.HelpMode
                        
                        If UserList(UserIndex).flags.Muerto = 0 And UserList(UserIndex).flags.AdminPerseguible And Not UserProtected Then
                            
                            If Distancia(.Pos, UserList(UserIndex).Pos) <= ConstantesBosses.BossDVDistance Then
                                .Target = UserIndex
                                Exit For
                            ElseIf RandomNumber(1, UBound(Query)) = 1 Or _
                                (I = UBound(Query) And .Target = 0) Then
                                
                                .Target = UserIndex
                            End If
                        End If

            Next I
        End If
        
        If .Target > 0 Then 'Si tengo target actuo si estoy a menos de la distancia maxima
            If Distance(.Pos.X, .Pos.Y, UserList(.Target).Pos.X, UserList(.Target).Pos.Y) > ConstantesBosses.BossDVDistance Then
                Call GeneralPathFinder(NpcIndex, .Target)
            End If
            Exit Sub
        End If
        
        'Si llega aca es que no había ningún usuario cercano vivo.
        'A bailar. Pablo (ToxicWaste)
        If RandomNumber(0, 10) = 0 Then
            Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
        End If
    End With

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub BossDVMovement de AI_NPC.bas")
End Sub

Private Sub BossDMAttack(ByVal NpcIndex As Integer)
'---------------------------------------------------------------------------------------
' Module    : AI
' Author    : Anagrama
' Date      : 03/09/2016
' Purpose   : Ataque del boss de DM.
'---------------------------------------------------------------------------------------
On Error GoTo ErrHandler
  
    Dim nPos As WorldPos
    Dim headingloop As Byte
    Dim FirstHeading As Byte
    Dim I As Long
    Dim UI As Integer
    Dim NPCI As Integer
    Dim UserTarget(1 To 4) As Integer
    Dim NpcTarget(1 To 4) As Integer
    Dim CurrentUser As Byte
    Dim CurrentNpc As Byte
    Dim UserProtected As Boolean
    Dim FurthestUser As Integer
    Dim FurthestDistance As Byte
    Dim tmpDistance As Byte
    Dim PutreTarget As Integer
    
    With Npclist(NpcIndex)
        
        If Not .Timers.Check(TimersIndex.Hit) Then Exit Sub
        
        'Esto es multifuncion, encuentra el usuario mas lejano y seleccion el target para putrefaccion.
        Dim Query() As Collision.UUID
        For I = 0 To ModAreas.QueryEntities(NpcIndex, ENTITY_TYPE_NPC, Query, ENTITY_TYPE_PLAYER)
            UI = Query(I).Name

                    UserProtected = Not IntervaloPermiteSerAtacado(UI) And UserList(UI).flags.NoPuedeSerAtacado
                    UserProtected = UserProtected Or UserList(UI).flags.HelpMode
                    
                    If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.AdminPerseguible And Not UserProtected Then
                        
                        tmpDistance = Distance(UserList(UI).Pos.X, UserList(UI).Pos.Y, .Pos.X, .Pos.Y)
                        If (FurthestDistance = 0 Or tmpDistance > FurthestDistance) And tmpDistance > 1 Then
                            FurthestDistance = tmpDistance
                            FurthestUser = UI
                        End If
                        If (RandomNumber(1, UBound(Query)) = 1) Or _
                            (I = UBound(Query)) Then
                            
                            If PutreTarget = 0 Then PutreTarget = UI
                        End If
                    End If

        Next I
        
        If PutreTarget Then
            If RandomNumber(1, 100) <= ConstantesBosses.BossDMCastPutrefaccion Then
                Call NpcLanzaSpellSobreUser(NpcIndex, PutreTarget, ConstantesBosses.BossDMSpellPutrefaccion, , True)
            End If
        End If
        
        If FurthestUser Then
            If RandomNumber(1, 100) <= ConstantesBosses.BossDMCastAparicion Then
                Call NpcLanzaSpellSobreUser(NpcIndex, FurthestUser, ConstantesBosses.BossDMSpellAparicion, , True)
            End If
        End If
        
        For headingloop = eHeading.NORTH To eHeading.WEST
            nPos = .Pos
            If .flags.Inmovilizado = 0 Or .Char.heading = headingloop Then
                Call HeadtoPos(headingloop, nPos)
                If InMapBounds(nPos.Map, nPos.X, nPos.Y) Then
                    UI = MapData(nPos.Map, nPos.X, nPos.Y).UserIndex
                    NPCI = MapData(nPos.Map, nPos.X, nPos.Y).NpcIndex
                    If UI > 0 Then
                        UserProtected = Not IntervaloPermiteSerAtacado(UI) And UserList(UI).flags.NoPuedeSerAtacado
                        UserProtected = UserProtected Or UserList(UI).flags.HelpMode
                        
                        If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.AdminPerseguible And (Not UserProtected) Then
                            If CurrentUser = 0 Then FirstHeading = headingloop
                            CurrentUser = CurrentUser + 1
                            UserTarget(CurrentUser) = UI
                        End If
                    ElseIf NPCI > 0 Then
                        If Npclist(NPCI).MaestroUser > 0 And Npclist(NPCI).flags.Paralizado = 0 Then
                            CurrentNpc = CurrentNpc + 1
                            NpcTarget(CurrentNpc) = NPCI
                        End If
                    End If
                End If
            End If  'inmo
        Next headingloop
        
        If CurrentUser Then
            Call ChangeNPCChar(NpcIndex, .Char.body, .Char.head, FirstHeading)
            For I = 1 To CurrentUser
                Call NpcAtacaUser(NpcIndex, UserTarget(I))
            Next I
        End If
        
        If CurrentNpc Then
            For I = 1 To CurrentNpc
                Call SistemaCombate.NpcAtacaNpc(NpcIndex, NpcTarget(I), False)
            Next I
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub BossDMAttack de AI_NPC.bas")
End Sub

Private Sub BossDMMovement(ByVal NpcIndex As Integer)
'---------------------------------------------------------------------------------------
' Module    : AI
' Author    : Anagrama
' Date      : 03/09/2016
' Purpose   : Movimiento del boss de DM.
'---------------------------------------------------------------------------------------
On Error GoTo ErrHandler
  
    Dim UserIndex As Integer
    Dim I As Long
    Dim UserProtected As Boolean
    Dim ClosestUser As Integer
    Dim tmpDistance As Byte
    Dim ClosestDistance As Byte
    
    With Npclist(NpcIndex)
        Dim Query() As Collision.UUID
        For I = 0 To ModAreas.QueryEntities(NpcIndex, ENTITY_TYPE_NPC, Query, ENTITY_TYPE_PLAYER)
            UserIndex = Query(I).Name

                    UserProtected = Not IntervaloPermiteSerAtacado(UserIndex) And UserList(UserIndex).flags.NoPuedeSerAtacado
                    UserProtected = UserProtected Or UserList(UserIndex).flags.HelpMode
                    
                    If UserList(UserIndex).flags.Muerto = 0 And UserList(UserIndex).flags.AdminPerseguible And Not UserProtected Then
                        
                        tmpDistance = Distance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, .Pos.X, .Pos.Y)
                        If ClosestDistance = 0 Or tmpDistance < ClosestDistance Then
                            ClosestDistance = tmpDistance
                            ClosestUser = UserIndex
                        End If
                    End If

        Next I
        
        If ClosestUser Then
            Call GeneralPathFinder(NpcIndex, ClosestUser)
            Exit Sub
        End If
        
        'Si llega aca es que no había ningún usuario cercano vivo.
        'A bailar. Pablo (ToxicWaste)
        If RandomNumber(0, 10) = 0 Then
            Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
        End If
    End With

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub BossDMMovement de AI_NPC.bas")
End Sub

Private Sub GeneralPathFinder(ByVal NpcIndex As Integer, ByVal TargetIndex As Integer, _
                                Optional ByVal TargetType As Byte = 0, Optional ByVal IsPet As Boolean = False)
'---------------------------------------------------------------------------------------
' Module    : AI
' Author    : Anagrama
' Date      : 17/07/2016
' Purpose   : Busca el mejor camino hacia el usuario sobre la marcha.
' Aclaro que la lectura de esto es una tortura porque son puros if con el
' fin de ahorrar la mayor cantidad de recursos posibles.
' Todavia se puede mejorar muchisimo esto para optimizarlo mas.
' 07/10/2016: Anagrama - Ahora funciona contra NPCs.
' 17/11/2016: Anagrama - Modificado para funcionar bien con npcs de agua y tomar como valor si es una mascota o no.
'---------------------------------------------------------------------------------------
On Error GoTo ErrHandler
  
    Dim AxisX As Byte
    Dim AxisY As Byte
    Dim MyHeading As Byte
    
    If TargetType = 0 Then
        If Distancia(Npclist(NpcIndex).Pos, UserList(TargetIndex).Pos) = 1 Then Exit Sub
    Else
        If Distancia(Npclist(NpcIndex).Pos, Npclist(TargetIndex).Pos) = 1 Then Exit Sub
    End If
    
    With Npclist(NpcIndex)
        If TargetType = 0 Then 'El target es un user.
            'Aca se encuentra la posicion relativa en los ejes X e Y
            If .Pos.Y < UserList(TargetIndex).Pos.Y Then
                AxisY = 0 'Npc arriba
            ElseIf .Pos.Y > UserList(TargetIndex).Pos.Y Then
                AxisY = 1 'Npc abajo
            Else
                AxisY = 2 'Npc vertical
            End If
            
            If .Pos.X > UserList(TargetIndex).Pos.X Then
                AxisX = 0 'Npc derecha
            ElseIf .Pos.X < UserList(TargetIndex).Pos.X Then
                AxisX = 1 'Npc izquierda
            Else
                AxisX = 2 'Npc horizontal
            End If
        Else 'El target es un npc.
            'Aca se encuentra la posicion relativa en los ejes X e Y
            If .Pos.Y < Npclist(TargetIndex).Pos.Y Then
                AxisY = 0 'Npc arriba
            ElseIf .Pos.Y > Npclist(TargetIndex).Pos.Y Then
                AxisY = 1 'Npc abajo
            Else
                AxisY = 2 'Npc vertical
            End If
            
            If .Pos.X > Npclist(TargetIndex).Pos.X Then
                AxisX = 0 'Npc derecha
            ElseIf .Pos.X < Npclist(TargetIndex).Pos.X Then
                AxisX = 1 'Npc izquierda
            Else
                AxisX = 2 'Npc horizontal
            End If
        End If
        
        'Asigna la direccion a la que revisar basandose en si es la primera vez que intenta buscar o no.
        If .flags.KeepHeading = 1 Then
            MyHeading = .Char.heading
        Else
            If AxisY = 0 And AxisX <> 1 Then
                MyHeading = eHeading.SOUTH
            ElseIf AxisY <> 0 And AxisX = 0 Then
                MyHeading = eHeading.WEST
            ElseIf AxisY <> 1 And AxisX = 1 Then
                MyHeading = eHeading.EAST
            ElseIf AxisY = 1 And AxisX <> 0 Then
                MyHeading = eHeading.NORTH
            End If
            .flags.KeepHeading = 1
        End If
        
        'Segun la direccion de npc revisa su posicion en relacion a la pos del target,
        'seguido a eso intenta moverse eliminando la diferencia en el eje contrario
        'al que esta apuntando.
        Select Case MyHeading
            Case eHeading.NORTH
                If AxisX = 1 Then
                    If LegalPosNPC(.Pos.Map, .Pos.X + 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                        Call MoveNPCChar(NpcIndex, eHeading.EAST)
                        Exit Sub
                    ElseIf LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y - 1, .flags.AguaValida, IsPet) Then
                        Call MoveNPCChar(NpcIndex, eHeading.NORTH)
                        Exit Sub
                    Else
                        If RandomNumber(1, 2) = 1 Then
                            If LegalPosNPC(.Pos.Map, .Pos.X - 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.WEST)
                                Exit Sub
                            ElseIf LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y + 1, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.SOUTH)
                                Exit Sub
                            End If
                        Else
                            If LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y + 1, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.SOUTH)
                                Exit Sub
                            ElseIf LegalPosNPC(.Pos.Map, .Pos.X - 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.WEST)
                                Exit Sub
                            End If
                        End If
                    End If
                ElseIf AxisX = 0 Then
                    If LegalPosNPC(.Pos.Map, .Pos.X - 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                        Call MoveNPCChar(NpcIndex, eHeading.WEST)
                        Exit Sub
                    ElseIf LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y - 1, .flags.AguaValida, IsPet) Then
                        Call MoveNPCChar(NpcIndex, eHeading.NORTH)
                        Exit Sub
                    Else
                        If RandomNumber(1, 2) = 1 Then
                            If LegalPosNPC(.Pos.Map, .Pos.X + 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.EAST)
                                Exit Sub
                            ElseIf LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y + 1, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.SOUTH)
                                Exit Sub
                            End If
                        Else
                            If LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y + 1, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.SOUTH)
                                Exit Sub
                            ElseIf LegalPosNPC(.Pos.Map, .Pos.X + 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.EAST)
                                Exit Sub
                            End If
                        End If
                    End If
                Else
                    If AxisY = 0 Then
                        If LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y + 1, .flags.AguaValida, IsPet) Then
                            Call MoveNPCChar(NpcIndex, eHeading.SOUTH)
                            Exit Sub
                        End If
                    Else
                        If LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y - 1, .flags.AguaValida, IsPet) Then
                            Call MoveNPCChar(NpcIndex, eHeading.NORTH)
                            Exit Sub
                        End If
                    End If
                    If RandomNumber(1, 2) = 1 Then
                        If LegalPosNPC(.Pos.Map, .Pos.X - 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                            Call MoveNPCChar(NpcIndex, eHeading.WEST)
                            Exit Sub
                        ElseIf LegalPosNPC(.Pos.Map, .Pos.X + 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                            Call MoveNPCChar(NpcIndex, eHeading.EAST)
                            Exit Sub
                        End If
                    Else
                        If LegalPosNPC(.Pos.Map, .Pos.X + 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                            Call MoveNPCChar(NpcIndex, eHeading.WEST)
                            Exit Sub
                        ElseIf LegalPosNPC(.Pos.Map, .Pos.X - 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                            Call MoveNPCChar(NpcIndex, eHeading.EAST)
                            Exit Sub
                        End If
                    End If
                End If
            Case eHeading.WEST
                If AxisY = 1 Then
                    If LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y - 1, .flags.AguaValida, IsPet) Then
                        Call MoveNPCChar(NpcIndex, eHeading.NORTH)
                        Exit Sub
                    ElseIf LegalPosNPC(.Pos.Map, .Pos.X - 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                        Call MoveNPCChar(NpcIndex, eHeading.WEST)
                        Exit Sub
                    Else
                        If RandomNumber(1, 2) = 1 Then
                            If LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y + 1, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.SOUTH)
                                Exit Sub
                            ElseIf LegalPosNPC(.Pos.Map, .Pos.X + 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.EAST)
                                Exit Sub
                            End If
                        Else
                            If LegalPosNPC(.Pos.Map, .Pos.X + 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.EAST)
                                Exit Sub
                            ElseIf LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y + 1, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.SOUTH)
                                Exit Sub
                            End If
                        End If
                    End If
                ElseIf AxisY = 0 Then
                    If LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y + 1, .flags.AguaValida, IsPet) Then
                        Call MoveNPCChar(NpcIndex, eHeading.SOUTH)
                        Exit Sub
                    ElseIf LegalPosNPC(.Pos.Map, .Pos.X - 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                        Call MoveNPCChar(NpcIndex, eHeading.WEST)
                        Exit Sub
                    Else
                        If RandomNumber(1, 2) = 1 Then
                            If LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y - 1, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.NORTH)
                                Exit Sub
                            ElseIf LegalPosNPC(.Pos.Map, .Pos.X + 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.EAST)
                                Exit Sub
                            End If
                        Else
                            If LegalPosNPC(.Pos.Map, .Pos.X + 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.EAST)
                                Exit Sub
                            ElseIf LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y - 1, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.NORTH)
                                Exit Sub
                            End If
                        End If
                    End If
                Else
                    If AxisX = 0 Then
                        If LegalPosNPC(.Pos.Map, .Pos.X - 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                            Call MoveNPCChar(NpcIndex, eHeading.WEST)
                            Exit Sub
                        End If
                    Else
                        If LegalPosNPC(.Pos.Map, .Pos.X + 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                            Call MoveNPCChar(NpcIndex, eHeading.EAST)
                            Exit Sub
                        End If
                    End If
                    
                    If RandomNumber(1, 2) = 1 Then
                        If LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y - 1, .flags.AguaValida, IsPet) Then
                            Call MoveNPCChar(NpcIndex, eHeading.NORTH)
                            Exit Sub
                        ElseIf LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y + 1, .flags.AguaValida, IsPet) Then
                            Call MoveNPCChar(NpcIndex, eHeading.SOUTH)
                            Exit Sub
                        End If
                    Else
                        If LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y + 1, .flags.AguaValida, IsPet) Then
                            Call MoveNPCChar(NpcIndex, eHeading.SOUTH)
                            Exit Sub
                        ElseIf LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y - 1, .flags.AguaValida, IsPet) Then
                            Call MoveNPCChar(NpcIndex, eHeading.NORTH)
                            Exit Sub
                        End If
                    End If
                End If
            Case eHeading.SOUTH
                If AxisX = 1 Then
                    If LegalPosNPC(.Pos.Map, .Pos.X + 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                        Call MoveNPCChar(NpcIndex, eHeading.EAST)
                        Exit Sub
                    ElseIf LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y + 1, .flags.AguaValida, IsPet) Then
                        Call MoveNPCChar(NpcIndex, eHeading.SOUTH)
                        Exit Sub
                    Else
                        If RandomNumber(1, 2) = 1 Then
                            If LegalPosNPC(.Pos.Map, .Pos.X - 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.WEST)
                                Exit Sub
                            ElseIf LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y - 1, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.NORTH)
                                Exit Sub
                            End If
                        Else
                            If LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y - 1, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.NORTH)
                                Exit Sub
                            ElseIf LegalPosNPC(.Pos.Map, .Pos.X - 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.WEST)
                                Exit Sub
                            End If
                        End If
                    End If
                ElseIf AxisX = 0 Then
                    If LegalPosNPC(.Pos.Map, .Pos.X - 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                        Call MoveNPCChar(NpcIndex, eHeading.WEST)
                        Exit Sub
                    ElseIf LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y + 1, .flags.AguaValida, IsPet) Then
                        Call MoveNPCChar(NpcIndex, eHeading.SOUTH)
                        Exit Sub
                    Else
                        If RandomNumber(1, 2) = 1 Then
                            If LegalPosNPC(.Pos.Map, .Pos.X + 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.EAST)
                                Exit Sub
                            ElseIf LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y - 1, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.NORTH)
                                Exit Sub
                            End If
                        Else
                            If LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y - 1, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.NORTH)
                                Exit Sub
                            ElseIf LegalPosNPC(.Pos.Map, .Pos.X + 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.EAST)
                                Exit Sub
                            End If
                        End If
                    End If
                Else
                    If AxisY = 1 Then
                        If LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y - 1, .flags.AguaValida, IsPet) Then
                            Call MoveNPCChar(NpcIndex, eHeading.NORTH)
                            Exit Sub
                        End If
                    Else
                        If LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y + 1, .flags.AguaValida, IsPet) Then
                            Call MoveNPCChar(NpcIndex, eHeading.SOUTH)
                            Exit Sub
                        End If
                    End If
                    If RandomNumber(1, 2) = 1 Then
                        If LegalPosNPC(.Pos.Map, .Pos.X - 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                            Call MoveNPCChar(NpcIndex, eHeading.WEST)
                            Exit Sub
                        ElseIf LegalPosNPC(.Pos.Map, .Pos.X + 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                            Call MoveNPCChar(NpcIndex, eHeading.EAST)
                            Exit Sub
                        End If
                    Else
                        If LegalPosNPC(.Pos.Map, .Pos.X + 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                            Call MoveNPCChar(NpcIndex, eHeading.WEST)
                            Exit Sub
                        ElseIf LegalPosNPC(.Pos.Map, .Pos.X - 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                            Call MoveNPCChar(NpcIndex, eHeading.EAST)
                            Exit Sub
                        End If
                    End If
                End If
            Case eHeading.EAST
                If AxisY = 1 Then
                    If LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y - 1, .flags.AguaValida, IsPet) Then
                        Call MoveNPCChar(NpcIndex, eHeading.NORTH)
                        Exit Sub
                    ElseIf LegalPosNPC(.Pos.Map, .Pos.X + 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                        Call MoveNPCChar(NpcIndex, eHeading.EAST)
                        Exit Sub
                    Else
                        If RandomNumber(1, 2) = 1 Then
                            If LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y + 1, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.SOUTH)
                                Exit Sub
                            ElseIf LegalPosNPC(.Pos.Map, .Pos.X - 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.WEST)
                                Exit Sub
                            End If
                        Else
                            If LegalPosNPC(.Pos.Map, .Pos.X - 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.WEST)
                                Exit Sub
                            ElseIf LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y + 1, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.SOUTH)
                                Exit Sub
                            End If
                        End If
                    End If
                ElseIf AxisY = 0 Then
                    If LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y + 1, .flags.AguaValida, IsPet) Then
                        Call MoveNPCChar(NpcIndex, eHeading.SOUTH)
                        Exit Sub
                    ElseIf LegalPosNPC(.Pos.Map, .Pos.X + 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                        Call MoveNPCChar(NpcIndex, eHeading.EAST)
                        Exit Sub
                    Else
                        If RandomNumber(1, 2) = 1 Then
                            If LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y - 1, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.NORTH)
                                Exit Sub
                            ElseIf LegalPosNPC(.Pos.Map, .Pos.X - 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.WEST)
                                Exit Sub
                            End If
                        Else
                            If LegalPosNPC(.Pos.Map, .Pos.X - 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.WEST)
                                Exit Sub
                            ElseIf LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y - 1, .flags.AguaValida, IsPet) Then
                                Call MoveNPCChar(NpcIndex, eHeading.NORTH)
                                Exit Sub
                            End If
                        End If
                    End If
                Else
                    If AxisX = 1 Then
                        If LegalPosNPC(.Pos.Map, .Pos.X + 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                            Call MoveNPCChar(NpcIndex, eHeading.EAST)
                            Exit Sub
                        End If
                    Else
                        If LegalPosNPC(.Pos.Map, .Pos.X - 1, .Pos.Y, .flags.AguaValida, IsPet) Then
                            Call MoveNPCChar(NpcIndex, eHeading.WEST)
                            Exit Sub
                        End If
                    End If
                    If RandomNumber(1, 2) = 1 Then
                        If LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y - 1, .flags.AguaValida, IsPet) Then
                            Call MoveNPCChar(NpcIndex, eHeading.NORTH)
                            Exit Sub
                        ElseIf LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y + 1, .flags.AguaValida, IsPet) Then
                            Call MoveNPCChar(NpcIndex, eHeading.SOUTH)
                            Exit Sub
                        End If
                    Else
                        If LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y + 1, .flags.AguaValida, IsPet) Then
                            Call MoveNPCChar(NpcIndex, eHeading.SOUTH)
                            Exit Sub
                        ElseIf LegalPosNPC(.Pos.Map, .Pos.X, .Pos.Y - 1, .flags.AguaValida, IsPet) Then
                            Call MoveNPCChar(NpcIndex, eHeading.NORTH)
                            Exit Sub
                        End If
                    End If
                End If
        End Select
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub GeneralPathFinder de AI_NPC.bas")
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''Viejas, feas y aburridas IAs :(''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub GuardiasAI(ByVal NpcIndex As Integer, ByVal DelCaos As Boolean)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 12/01/2010 (ZaMa)
'14/09/2009: ZaMa - Now npcs don't atack protected users.
'12/01/2010: ZaMa - Los npcs no atacan druidas mimetizados con npcs
'***************************************************
On Error GoTo ErrHandler
  
    Dim nPos As WorldPos
    Dim headingloop As Byte
    Dim UI As Integer
    Dim UserProtected As Boolean
    
    If Not Npclist(NpcIndex).Timers.Check(TimersIndex.Hit, False) Then Exit Sub
    
    With Npclist(NpcIndex)
        For headingloop = eHeading.NORTH To eHeading.WEST
            nPos = .Pos
            If .flags.Inmovilizado = 0 Or headingloop = .Char.heading Then
                Call HeadtoPos(headingloop, nPos)
                If InMapBounds(nPos.Map, nPos.X, nPos.Y) Then
                    UI = MapData(nPos.Map, nPos.X, nPos.Y).UserIndex
                    If UI > 0 Then
                        UserProtected = Not IntervaloPermiteSerAtacado(UI) And UserList(UI).flags.NoPuedeSerAtacado
                        UserProtected = UserProtected Or UserList(UI).flags.Ignorado Or UserList(UI).flags.HelpMode
                        
                        If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.AdminPerseguible And Not UserProtected Then
                            If Not DelCaos Then
                                If UserList(UI).Faccion.Alignment = eCharacterAlignment.FactionLegion Then
                                    If NpcAtacaUser(NpcIndex, UI) Then
                                        Call ChangeNPCChar(NpcIndex, .Char.body, .Char.head, headingloop)
                                        Call Npclist(NpcIndex).Timers.Restart(TimersIndex.Hit)
                                    End If
                                    Exit Sub
                                ElseIf .flags.AttackedBy = UserList(UI).Name And Not .flags.Follow Then
                                    
                                    If NpcAtacaUser(NpcIndex, UI) Then
                                        Call ChangeNPCChar(NpcIndex, .Char.body, .Char.head, headingloop)
                                        Call Npclist(NpcIndex).Timers.Restart(TimersIndex.Hit)
                                    End If
                                    Exit Sub
                                End If
                            Else
                                If UserList(UI).Faccion.Alignment = eCharacterAlignment.FactionRoyal Then
                                    If NpcAtacaUser(NpcIndex, UI) Then
                                        Call ChangeNPCChar(NpcIndex, .Char.body, .Char.head, headingloop)
                                        Call Npclist(NpcIndex).Timers.Restart(TimersIndex.Hit)
                                    End If
                                    Exit Sub
                                ElseIf .flags.AttackedBy = UserList(UI).Name And Not .flags.Follow Then
                                      
                                    If NpcAtacaUser(NpcIndex, UI) Then
                                        Call ChangeNPCChar(NpcIndex, .Char.body, .Char.head, headingloop)
                                        Call Npclist(NpcIndex).Timers.Restart(TimersIndex.Hit)
                                    End If
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                End If
            End If  'not inmovil
        Next headingloop
    End With
    
    'Call RestoreOldMovement(NpcIndex)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub GuardiasAI de AI_NPC.bas")
End Sub

''
' Handles the evil npcs' artificial intelligency.
'
' @param NpcIndex Specifies reference to the npc
Private Sub HostilMalvadoAI(ByVal NpcIndex As Integer)
'**************************************************************
'Author: Unknown
'Last Modify Date: 12/01/2010 (ZaMa)
'28/04/2009: ZaMa - Now those NPCs who doble attack, have 50% of posibility of casting a spell on user.
'14/09/200*: ZaMa - Now npcs don't atack protected users.
'12/01/2010: ZaMa - Los npcs no atacan druidas mimetizados con npcs
'**************************************************************
On Error GoTo ErrHandler
  
    Dim nPos As WorldPos
    Dim headingloop As Byte
    Dim UI As Integer
    Dim NPCI As Integer
    Dim atacoPJ As Boolean
    Dim UserProtected As Boolean
    
    If Not Npclist(NpcIndex).Timers.Check(TimersIndex.Hit, False) Then Exit Sub
    
    atacoPJ = False
    
    With Npclist(NpcIndex)
        For headingloop = eHeading.NORTH To eHeading.WEST
            nPos = .Pos
            If .flags.Inmovilizado = 0 Or .Char.heading = headingloop Then
                Call HeadtoPos(headingloop, nPos)
                If InMapBounds(nPos.Map, nPos.X, nPos.Y) Then
                    UI = MapData(nPos.Map, nPos.X, nPos.Y).UserIndex
                    NPCI = MapData(nPos.Map, nPos.X, nPos.Y).NpcIndex
                    If UI > 0 And Not atacoPJ Then
                        UserProtected = Not IntervaloPermiteSerAtacado(UI) And UserList(UI).flags.NoPuedeSerAtacado
                        UserProtected = UserProtected Or UserList(UI).flags.Ignorado Or UserList(UI).flags.HelpMode
                        
                        If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.AdminPerseguible And (Not UserProtected) Then
                            
                            atacoPJ = True
                            If .Movement = NpcObjeto Then
                                ' Los npc objeto no atacan siempre al mismo usuario
                                If RandomNumber(1, 3) = 3 Then atacoPJ = False
                            End If
                            
                            If atacoPJ Then
                                If .flags.LanzaSpells Then
                                    If .flags.AtacaDoble Then
                                        If (RandomNumber(0, 1)) Then
                                            If NpcAtacaUser(NpcIndex, UI) Then
                                                Call ChangeNPCChar(NpcIndex, .Char.body, .Char.head, headingloop)
                                                Call Npclist(NpcIndex).Timers.Restart(TimersIndex.Hit)
                                            End If
                                            Exit Sub
                                        End If
                                    End If
                                    
                                    If UserList(UI).flags.invisible = 0 And UserList(UI).flags.Oculto = 0 Then
                                        Call ChangeNPCChar(NpcIndex, .Char.body, .Char.head, headingloop)
                                        Call NpcLanzaUnSpell(NpcIndex, UI)
                                        Call Npclist(NpcIndex).Timers.Restart(TimersIndex.Hit)
                                        Exit Sub
                                    End If
                                End If
                            End If
                            If NpcAtacaUser(NpcIndex, UI) Then
                                Call ChangeNPCChar(NpcIndex, .Char.body, .Char.head, headingloop)
                                Call Npclist(NpcIndex).Timers.Restart(TimersIndex.Hit)
                            End If
                            Exit Sub

                        End If
                    ElseIf NPCI > 0 Then
                        If Npclist(NPCI).MaestroUser > 0 And Npclist(NPCI).flags.Paralizado = 0 Then
                            Call ChangeNPCChar(NpcIndex, .Char.body, .Char.head, headingloop)
                            Call SistemaCombate.NpcAtacaNpc(NpcIndex, NPCI, False)
                            Call Npclist(NpcIndex).Timers.Restart(TimersIndex.Hit)
                            Exit Sub
                        End If
                    End If
                End If
            End If  'inmo
        Next headingloop
    End With
    
    'Call RestoreOldMovement(NpcIndex)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HostilMalvadoAI de AI_NPC.bas")
End Sub

Private Sub HostilBuenoAI(ByVal NpcIndex As Integer)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 12/01/2010 (ZaMa)
'14/09/2009: ZaMa - Now npcs don't atack protected users.
'12/01/2010: ZaMa - Los npcs no atacan druidas mimetizados con npcs
'***************************************************
On Error GoTo ErrHandler
  
    Dim nPos As WorldPos
    Dim headingloop As eHeading
    Dim UI As Integer
    Dim UserProtected As Boolean
    
    If Not Npclist(NpcIndex).Timers.Check(TimersIndex.Hit, False) Then Exit Sub
    
    With Npclist(NpcIndex)
        For headingloop = eHeading.NORTH To eHeading.WEST
            nPos = .Pos
            If .flags.Inmovilizado = 0 Or .Char.heading = headingloop Then
                Call HeadtoPos(headingloop, nPos)
                If InMapBounds(nPos.Map, nPos.X, nPos.Y) Then
                    UI = MapData(nPos.Map, nPos.X, nPos.Y).UserIndex
                    If UI > 0 Then
                        If UserList(UI).Name = .flags.AttackedBy Then
                        
                            UserProtected = Not IntervaloPermiteSerAtacado(UI) And UserList(UI).flags.NoPuedeSerAtacado
                            UserProtected = UserProtected Or UserList(UI).flags.Ignorado Or UserList(UI).flags.HelpMode
                            
                            If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.AdminPerseguible And Not UserProtected Then
                                If .flags.LanzaSpells > 0 Then
                                    If UserList(UI).flags.invisible = 0 And UserList(UI).flags.Oculto = 0 Then
                                        Call NpcLanzaUnSpell(NpcIndex, UI)
                                        Call Npclist(NpcIndex).Timers.Restart(TimersIndex.Hit)
                                        Exit Sub
                                    End If
                                End If
                                
                                If NpcAtacaUser(NpcIndex, UI) Then
                                    Call ChangeNPCChar(NpcIndex, .Char.body, .Char.head, headingloop)
                                    Call Npclist(NpcIndex).Timers.Restart(TimersIndex.Hit)
                                End If
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            End If
        Next headingloop
    End With
    
    'Call RestoreOldMovement(NpcIndex)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HostilBuenoAI de AI_NPC.bas")
End Sub

Private Sub IrUsuarioCercano(ByVal NpcIndex As Integer)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 25/07/2010 (ZaMa)
'14/09/2009: ZaMa - Now npcs don't follow protected users.
'12/01/2010: ZaMa - Los npcs no atacan druidas mimetizados con npcs
'25/07/2010: ZaMa - Agrego una validacion temporal para evitar que los npcs ataquen a usuarios de mapas difernetes.
'***************************************************
On Error GoTo ErrHandler
  
    Dim tHeading As Byte
    Dim UserIndex As Integer
    Dim SignoNS As Integer
    Dim SignoEO As Integer
    Dim I As Long
    Dim UserProtected As Boolean
    
    With Npclist(NpcIndex)
        If .flags.Inmovilizado = 1 Then
            Select Case .Char.heading
                Case eHeading.NORTH
                    SignoNS = -1
                    SignoEO = 0
                
                Case eHeading.EAST
                    SignoNS = 0
                    SignoEO = 1
                
                Case eHeading.SOUTH
                    SignoNS = 1
                    SignoEO = 0
                
                Case eHeading.WEST
                    SignoEO = -1
                    SignoNS = 0
            End Select
            
            Dim Query() As Collision.UUID
            For I = 0 To ModAreas.QueryEntities(NpcIndex, ENTITY_TYPE_NPC, Query, ENTITY_TYPE_PLAYER, eGridSortClosest)
                UserIndex = Query(I).Name
                
                'Is it in it's range of vision??
                If Sgn(UserList(UserIndex).Pos.X - .Pos.X) = SignoEO Then
                    If Sgn(UserList(UserIndex).Pos.Y - .Pos.Y) = SignoNS Then
                        
                        UserProtected = Not IntervaloPermiteSerAtacado(UserIndex) And UserList(UserIndex).flags.NoPuedeSerAtacado
                        UserProtected = UserProtected Or UserList(UserIndex).flags.Ignorado Or UserList(UserIndex).flags.HelpMode
                        
                        If UserList(UserIndex).flags.Muerto = 0 Then
                            If Not UserProtected Then
                                Call NpcAttacksAndSpellUser(UserIndex, NpcIndex)
                                Exit Sub
                            End If
                        End If
                        
                    End If
                End If
            Next I
            
        ' No esta inmobilizado
        Else
            
            If .flags.VolviendoOrig Then
                tHeading = FindDirection(Npclist(NpcIndex).Pos, Npclist(NpcIndex).Orig)
                Call MoveNPCChar(NpcIndex, tHeading)
                Exit Sub
            End If
            
            ' Tiene prioridad de seguir al usuario al que le pertenece si esta en el rango de vision
            Dim OwnerIndex As Integer
            
            OwnerIndex = .Owner
            If OwnerIndex > 0 Then
                
                ' TODO: Es temporal hatsa reparar un bug que hace que ataquen a usuarios de otros mapas
                If UserList(OwnerIndex).Pos.Map = .Pos.Map Then
                    
                    'Is it in it's range of vision??
                    If Abs(UserList(OwnerIndex).Pos.X - .Pos.X) <= RANGO_VISION_X Then
                        If Abs(UserList(OwnerIndex).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                            
                            ' va hacia el si no esta invi ni oculto
                            If UserList(OwnerIndex).flags.invisible = 0 And UserList(OwnerIndex).flags.Oculto = 0 And Not UserList(OwnerIndex).flags.HelpMode And Not UserList(OwnerIndex).flags.Ignorado Then
                               
                                Call NpcAttacksAndSpellUser(OwnerIndex, NpcIndex)
                                
                                If Npclist(NpcIndex).PathFinding = 1 Then
                                    Call GeneralPathFinder(NpcIndex, OwnerIndex)
                                    Exit Sub
                                Else
                                    tHeading = FindDirection(Npclist(NpcIndex).Pos, UserList(OwnerIndex).Pos)
                                    Call MoveNPCChar(NpcIndex, tHeading)
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                
                ' Esto significa que esta bugueado.. Lo logueo, y "reparo" el error a mano (Todo temporal)
                Else
                    Call LogError("El npc: " & .Name & "(" & NpcIndex & "), intenta atacar a " & _
                                  UserList(OwnerIndex).Name & "(Index: " & OwnerIndex & ", Mapa: " & _
                                  UserList(OwnerIndex).Pos.Map & ") desde el mapa " & .Pos.Map)
                    .Owner = 0
                End If
                
            End If
            
            ' No le pertenece a nadie o el dueño no esta en el rango de vision, sigue a cualquiera
            For I = 0 To ModAreas.QueryEntities(NpcIndex, ENTITY_TYPE_NPC, Query, ENTITY_TYPE_PLAYER, eGridSortClosest)
                UserIndex = Query(I).Name

                        With UserList(UserIndex)
                            
                            UserProtected = Not IntervaloPermiteSerAtacado(UserIndex) And .flags.NoPuedeSerAtacado
                            UserProtected = UserProtected Or .flags.Ignorado Or .flags.HelpMode
                            
                            If .flags.Muerto = 0 And .flags.invisible = 0 And .flags.Oculto = 0 And _
                                .flags.AdminPerseguible And Not UserProtected Then
                                
                                Call NpcAttacksAndSpellUser(UserIndex, NpcIndex)
                                
                                If Npclist(NpcIndex).PathFinding = 1 Then
                                    Call GeneralPathFinder(NpcIndex, UserIndex)
                                    Exit Sub
                                Else
                                    tHeading = FindDirection(Npclist(NpcIndex).Pos, UserList(UserIndex).Pos)
                                    Call MoveNPCChar(NpcIndex, tHeading)
                                    Exit Sub
                                End If
                            End If
                            
                        End With

            Next I
            
            'Si llega aca es que no habia ningun usuario cercano vivo.
            'A bailar. Pablo (ToxicWaste)
            If RandomNumber(0, 10) = 0 Then
                Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
            End If
            
        End If
    End With
    
    Call RestoreOldMovement(NpcIndex)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub IrUsuarioCercano de AI_NPC.bas")
End Sub

Public Function GetHeading(ByVal UserIndex As Integer, ByVal NpcIndex As Integer) As Integer
    Dim NpcX As Integer
    Dim NpcY As Integer
    Dim UserX As Integer
    Dim UserY As Integer
    
    NpcX = Npclist(NpcIndex).Pos.X
    NpcY = Npclist(NpcIndex).Pos.Y
    UserX = UserList(UserIndex).Pos.X
    UserY = UserList(UserIndex).Pos.Y
    
    If UserX < NpcX Then
        GetHeading = eHeading.WEST
        Exit Function
    End If
    
    If UserX > NpcX Then
        GetHeading = eHeading.EAST
        Exit Function
    End If
    
    If UserY > NpcY Then
        GetHeading = eHeading.SOUTH
        Exit Function
    End If
    
    If UserY < NpcY Then
        GetHeading = eHeading.NORTH
        Exit Function
    End If
    
    GetHeading = eHeading.EAST
End Function
''
' Makes a Pet / Summoned Npc to Follow an enemy
'
' @param NpcIndex Specifies reference to the npc
Private Sub SeguirAgresor(ByVal NpcIndex As Integer)
'**************************************************************
'Author: Unknown
'Last Modify by: Marco Vanotti (MarKoxX)
'Last Modify Date: 08/16/2008
'08/16/2008: MarKoxX - Now pets that do melee attacks have to be near the enemy to attack.
'**************************************************************
On Error GoTo ErrHandler
  
    Dim tHeading As Byte
    Dim UI As Integer
    
    Dim I As Long
    
    Dim SignoNS As Integer
    Dim SignoEO As Integer
    
    Dim UserProtected As Boolean
    
    With Npclist(NpcIndex)
        If .flags.Paralizado = 1 Or .flags.Inmovilizado = 1 Then
            Select Case .Char.heading
                Case eHeading.NORTH
                    SignoNS = -1
                    SignoEO = 0
                
                Case eHeading.EAST
                    SignoNS = 0
                    SignoEO = 1
                
                Case eHeading.SOUTH
                    SignoNS = 1
                    SignoEO = 0
                
                Case eHeading.WEST
                    SignoEO = -1
                    SignoNS = 0
            End Select
            
            Dim Query() As Collision.UUID
            For I = 0 To ModAreas.QueryEntities(NpcIndex, ENTITY_TYPE_NPC, Query, ENTITY_TYPE_PLAYER, eGridSortClosest)
                UI = Query(I).Name
                
                'Is it in it's range of vision??
                If Sgn(UserList(UI).Pos.X - .Pos.X) = SignoEO Then
                    If Sgn(UserList(UI).Pos.Y - .Pos.Y) = SignoNS Then

                        UserProtected = Not IntervaloPermiteSerAtacado(UI) And UserList(UI).flags.NoPuedeSerAtacado
                        UserProtected = UserProtected Or UserList(UI).flags.Ignorado Or UserList(UI).flags.HelpMode

                        If UserList(UI).Name = .flags.AttackedBy And (Not UserProtected) Then
                            If .MaestroUser > 0 Then
                                If UserList(.MaestroUser).Faccion.Alignment = eCharacterAlignment.FactionRoyal And UserList(UI).Faccion.Alignment = eCharacterAlignment.FactionRoyal And (UserList(.MaestroUser).flags.Seguro Or UserList(.MaestroUser).Faccion.ArmadaReal = 1) Then
                                    Call WriteConsoleMsg(.MaestroUser, "La mascota no atacará a ciudadanos si eres miembro del ejército real o tienes el seguro activado.", FontTypeNames.FONTTYPE_INFO)
                                    .flags.AttackedBy = vbNullString
                                    Exit Sub
                                End If
                            End If

                            If UserList(UI).flags.Muerto = 0 Then
                                 If .flags.LanzaSpells > 0 Then
                                    If Npclist(NpcIndex).Timers.Check(TimersIndex.Hit) And UserList(UI).flags.invisible = 0 And UserList(UI).flags.Oculto = 0 Then Call NpcLanzaUnSpell(NpcIndex, UI)
                                 Else
                                    If Distancia(UserList(UI).Pos, Npclist(NpcIndex).Pos) <= 1 Then
                                        If Npclist(NpcIndex).Timers.Check(TimersIndex.Hit) Then Call NpcAtacaUser(NpcIndex, UI)
                                    End If
                                 End If
                                 Exit Sub
                            End If
                        End If
                        
                    End If
                End If
                
            Next I
        Else
            If .flags.VolviendoOrig Then
                tHeading = FindDirection(Npclist(NpcIndex).Pos, Npclist(NpcIndex).Orig)
                Call MoveNPCChar(NpcIndex, tHeading)
                Exit Sub
            End If
            
            For I = 0 To ModAreas.QueryEntities(NpcIndex, ENTITY_TYPE_NPC, Query, ENTITY_TYPE_PLAYER, eGridSortClosest)
                UI = Query(I).Name

                        UserProtected = Not IntervaloPermiteSerAtacado(UI) And UserList(UI).flags.NoPuedeSerAtacado
                        UserProtected = UserProtected Or UserList(UI).flags.Ignorado Or UserList(UI).flags.HelpMode
                        
                        If UserList(UI).Name = .flags.AttackedBy And (Not UserProtected) Then
                            If .MaestroUser > 0 Then
                                If UserList(.MaestroUser).Faccion.Alignment = eCharacterAlignment.FactionRoyal And UserList(UI).Faccion.Alignment = eCharacterAlignment.FactionRoyal And (UserList(.MaestroUser).flags.Seguro Or UserList(.MaestroUser).Faccion.Alignment = FactionRoyal) Then
                                    Call WriteConsoleMsg(.MaestroUser, "La mascota no atacará a ciudadanos si eres miembro del ejército real o tienes el seguro activado.", FontTypeNames.FONTTYPE_INFO)
                                    .flags.AttackedBy = vbNullString
                                    Call FollowAmo(NpcIndex)
                                    Exit Sub
                                End If
                            End If
                            
                            If UserList(UI).flags.Muerto = 0 Then
                                 If .flags.LanzaSpells > 0 Then
                                 
                                    If UserList(UI).flags.invisible = 0 And UserList(UI).flags.Oculto = 0 Then
                                        Call NpcAttacksAndSpellUser(UI, NpcIndex)
                                    End If
                                 Else
                                    If Distancia(UserList(UI).Pos, Npclist(NpcIndex).Pos) <= 1 Then
                                        If Npclist(NpcIndex).Timers.Check(TimersIndex.Hit) Then Call NpcAtacaUser(NpcIndex, UI)
                                    End If
                                 End If
                                 
                                 If UserList(UI).flags.invisible = 0 And UserList(UI).flags.Oculto = 0 Then
                                     If Npclist(NpcIndex).PathFinding = 1 Then
                                        Call GeneralPathFinder(NpcIndex, UI)
                                        Exit Sub
                                    Else
                                        tHeading = FindDirection(Npclist(NpcIndex).Pos, UserList(UI).Pos)
                                        Call MoveNPCChar(NpcIndex, tHeading)
                                        Exit Sub
                                    End If
                                 End If
                            End If
                        End If

            Next I
        End If
    End With
    
    Call RestoreOldMovement(NpcIndex)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SeguirAgresor de AI_NPC.bas")
End Sub

Private Sub NpcAttacksAndSpellUser(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)

    With Npclist(NpcIndex)
        If .Timers.Check(TimersIndex.Hit, False) Then
            If .flags.LanzaSpells Then
            
                If UserList(UserIndex).flags.invisible <> 1 And UserList(UserIndex).flags.Oculto <> 1 Then
                    Call NpcLanzaUnSpell(NpcIndex, UserIndex)
                    .Timers.Restart (TimersIndex.Hit)
                    If .flags.AtacaDoble = 0 Then
                        Exit Sub
                    End If
                End If
                
                If Abs(UserList(UserIndex).Pos.Y - .Pos.Y) + Abs(UserList(UserIndex).Pos.X - .Pos.X) <= 1 Then
                    Call NpcAtacaUser(NpcIndex, UserIndex)
                    .Timers.Restart (TimersIndex.Hit)
                    Call ChangeNPCChar(NpcIndex, .Char.body, .Char.head, GetHeading(UserIndex, NpcIndex))
                End If
                
            Else
                If Abs(UserList(UserIndex).Pos.Y - .Pos.Y) + Abs(UserList(UserIndex).Pos.X - .Pos.X) <= 1 Then
                    Call NpcAtacaUser(NpcIndex, UserIndex)
                    .Timers.Restart (TimersIndex.Hit)
                    Call ChangeNPCChar(NpcIndex, .Char.body, .Char.head, GetHeading(UserIndex, NpcIndex))
                End If
            End If
        End If
    End With
    
    
End Sub

Private Sub RestoreOldMovement(ByVal NpcIndex As Integer)
On Error GoTo ErrHandler
  
    With Npclist(NpcIndex)
        If .MaestroUser = 0 Then
            .Movement = .flags.OldMovement
            .Hostile = .flags.OldHostil
            .flags.AttackedBy = vbNullString
            .flags.KeepHeading = 0
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub RestoreOldMovement de AI_NPC.bas")
End Sub

Private Sub FollowRoyal(ByVal NpcIndex As Integer)
On Error GoTo ErrHandler
  
    Dim UserIndex As Integer
    Dim tHeading As Byte
    Dim I As Long
    Dim UserProtected As Boolean
    
    With Npclist(NpcIndex)
        If .flags.VolviendoOrig Then
            tHeading = FindDirection(Npclist(NpcIndex).Pos, Npclist(NpcIndex).Orig)
            Call MoveNPCChar(NpcIndex, tHeading)
            Exit Sub
        End If
            
        Dim Query() As Collision.UUID
        For I = 0 To ModAreas.QueryEntities(NpcIndex, ENTITY_TYPE_NPC, Query, ENTITY_TYPE_PLAYER, eGridSortClosest)
            UserIndex = Query(I).Name

                    If UserList(UserIndex).Faccion.Alignment = FactionRoyal Then
                    
                        UserProtected = Not IntervaloPermiteSerAtacado(UserIndex) And UserList(UserIndex).flags.NoPuedeSerAtacado
                        UserProtected = UserProtected Or UserList(UserIndex).flags.Ignorado Or UserList(UserIndex).flags.HelpMode
                        
                        If UserList(UserIndex).flags.Muerto = 0 And UserList(UserIndex).flags.invisible = 0 And _
                            UserList(UserIndex).flags.Oculto = 0 And UserList(UserIndex).flags.AdminPerseguible And Not UserProtected Then
                            
                            If .flags.LanzaSpells > 0 Then
                                If Npclist(NpcIndex).Timers.Check(TimersIndex.Hit) Then Call NpcLanzaUnSpell(NpcIndex, UserIndex)
                            End If
                            If Npclist(NpcIndex).PathFinding = 1 Then
                                Call GeneralPathFinder(NpcIndex, UserIndex)
                                Exit Sub
                            Else
                                tHeading = FindDirection(Npclist(NpcIndex).Pos, UserList(UserIndex).Pos)
                                Call MoveNPCChar(NpcIndex, tHeading)
                                Exit Sub
                            End If
                            Exit Sub
                        End If
                    End If

        Next I
    End With
    
    Call RestoreOldMovement(NpcIndex)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub FollowRoyal de AI_NPC.bas")
End Sub

Private Sub FollowLegion(ByVal NpcIndex As Integer)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 12/01/2010 (ZaMa)
'14/09/2009: ZaMa - Now npcs don't follow protected users.
'12/01/2010: ZaMa - Los npcs no atacan druidas mimetizados con npcs.
'***************************************************
On Error GoTo ErrHandler
  
    Dim UserIndex As Integer
    Dim tHeading As Byte
    Dim I As Long
    Dim SignoNS As Integer
    Dim SignoEO As Integer
    Dim UserProtected As Boolean
    
    With Npclist(NpcIndex)
        If .flags.Inmovilizado = 1 Then
            Select Case .Char.heading
                Case eHeading.NORTH
                    SignoNS = -1
                    SignoEO = 0
                
                Case eHeading.EAST
                    SignoNS = 0
                    SignoEO = 1
                
                Case eHeading.SOUTH
                    SignoNS = 1
                    SignoEO = 0
                
                Case eHeading.WEST
                    SignoEO = -1
                    SignoNS = 0
            End Select
                
            Dim Query() As Collision.UUID
            For I = 0 To ModAreas.QueryEntities(NpcIndex, ENTITY_TYPE_NPC, Query, ENTITY_TYPE_PLAYER, eGridSortClosest)
                UserIndex = Query(I).Name
                    
                'Is it in it's range of vision??
                If Sgn(UserList(UserIndex).Pos.X - .Pos.X) = SignoEO Then
                    If Sgn(UserList(UserIndex).Pos.Y - .Pos.Y) = SignoNS Then
                        
                        If UserList(UserIndex).Faccion.Alignment = eCharacterAlignment.FactionLegion Then
                            With UserList(UserIndex)
                                 
                                 UserProtected = Not IntervaloPermiteSerAtacado(UserIndex) And .flags.NoPuedeSerAtacado
                                 UserProtected = UserProtected Or UserList(UserIndex).flags.Ignorado Or UserList(UserIndex).flags.HelpMode
                                 
                                 If .flags.Muerto = 0 And .flags.invisible = 0 And _
                                    .flags.Oculto = 0 And .flags.AdminPerseguible And Not UserProtected Then
                                     
                                     If Npclist(NpcIndex).flags.LanzaSpells > 0 Then
                                        If Npclist(NpcIndex).Timers.Check(TimersIndex.Hit) Then Call NpcLanzaUnSpell(NpcIndex, UserIndex)
                                     End If
                                     Exit Sub
                                End If
                            End With
                        End If
                        
                   End If
                End If
            Next I
        Else
            If .flags.VolviendoOrig Then
                tHeading = FindDirection(Npclist(NpcIndex).Pos, Npclist(NpcIndex).Orig)
                Call MoveNPCChar(NpcIndex, tHeading)
                Exit Sub
            End If
                
            For I = 0 To ModAreas.QueryEntities(NpcIndex, ENTITY_TYPE_NPC, Query, ENTITY_TYPE_PLAYER, eGridSortClosest)
                UserIndex = Query(I).Name

                If UserList(UserIndex).Faccion.Alignment = eCharacterAlignment.FactionLegion Then
                    
                    UserProtected = Not IntervaloPermiteSerAtacado(UserIndex) And UserList(UserIndex).flags.NoPuedeSerAtacado
                    UserProtected = UserProtected Or UserList(UserIndex).flags.Ignorado Or UserList(UserIndex).flags.HelpMode
                    
                    If UserList(UserIndex).flags.Muerto = 0 And UserList(UserIndex).flags.invisible = 0 And _
                       UserList(UserIndex).flags.Oculto = 0 And UserList(UserIndex).flags.AdminPerseguible And Not UserProtected Then
                        If .flags.LanzaSpells > 0 Then
                            If Npclist(NpcIndex).Timers.Check(TimersIndex.Hit) Then Call NpcLanzaUnSpell(NpcIndex, UserIndex)
                        End If
                        If .flags.Inmovilizado = 1 Then Exit Sub
                        If Npclist(NpcIndex).PathFinding = 1 Then
                            Call GeneralPathFinder(NpcIndex, UserIndex)
                            Exit Sub
                        Else
                            tHeading = FindDirection(Npclist(NpcIndex).Pos, UserList(UserIndex).Pos)
                            Call MoveNPCChar(NpcIndex, tHeading)
                            Exit Sub
                        End If
                        Exit Sub
                   End If
                End If

            Next I
        End If
    End With
    
    Call RestoreOldMovement(NpcIndex)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub FollowLegion de AI_NPC.bas")
End Sub

Private Sub SeguirAmo(ByVal NpcIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    Dim tHeading As Byte
    Dim UI As Integer
    
    With Npclist(NpcIndex)
        If .Target = 0 And .TargetNPC = 0 Then
            UI = .MaestroUser
            
            If UI > 0 Then
                'Is it in it's range of vision??
                If Abs(UserList(UI).Pos.X - .Pos.X) <= RANGO_VISION_X Then
                    If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                        If UserList(UI).flags.Muerto = 0 _
                                And UserList(UI).flags.invisible = 0 _
                                And UserList(UI).flags.Oculto = 0 _
                                And Distancia(.Pos, UserList(UI).Pos) > 3 Then
                            If Npclist(NpcIndex).PathFinding = 1 Then
                                Call GeneralPathFinder(NpcIndex, UI, , True)
                                Exit Sub
                            Else
                                tHeading = FindDirection(Npclist(NpcIndex).Pos, UserList(UI).Pos)
                                Call MoveNPCChar(NpcIndex, tHeading)
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End With
    
    Call RestoreOldMovement(NpcIndex)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SeguirAmo de AI_NPC.bas")
End Sub

Private Sub MovimientoNpcAtacaNpc(ByVal NpcIndex As Integer)
On Error GoTo ErrHandler
  

'***************************************************
'Author: Anagrama
'Last Modification: 17/08/2016
'Movimiento de eles y mascotas.
'***************************************************
    Dim tHeading As Byte
    Dim I As Long
    Dim X As Long
    Dim Y As Long
    Dim NI As Integer
    Dim bNoEsta As Boolean
    
    With Npclist(NpcIndex)
        Dim Results() As Collision.UUID
        

    
        If .flags.Inmovilizado = 0 Then
        
            For I = 0 To ModAreas.QueryEntities(NpcIndex, ENTITY_TYPE_NPC, Results, ENTITY_TYPE_NPC, eGridSortClosest)
                NI = Results(I).Name
                
                If NI > 0 Then
                     If .TargetNpc = NI Then
                          bNoEsta = True
                         If Npclist(NpcIndex).PathFinding = 1 Then
                             Call GeneralPathFinder(NpcIndex, NI, 1, True)
                             Exit Sub
                         Else
                             tHeading = FindDirection(Npclist(NpcIndex).Pos, Npclist(NI).Pos)
                             Call MoveNPCChar(NpcIndex, tHeading)
                             Exit Sub
                         End If
                         Exit Sub
                     End If
                End If
            Next I
        End If
        
        If Not bNoEsta Then
            If .MaestroUser > 0 Then
                Call FollowAmo(NpcIndex)
            Else
                .Movement = .flags.OldMovement
                .Hostile = .flags.OldHostil
            End If
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub MovimientoNpcAtacaNpc de AI_NPC.bas")
End Sub

Private Sub AiNpcAtacaNpc(ByVal NpcIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    Dim X As Long
    Dim Y As Long
    Dim NI As Integer
    Dim bNoEsta As Boolean
    
    Dim SignoNS As Integer
    Dim SignoEO As Integer
    
    If Not Npclist(NpcIndex).Timers.Check(TimersIndex.Hit) Then Exit Sub
    
    With Npclist(NpcIndex)
        If .flags.Inmovilizado = 1 Then
            Select Case .Char.heading
                Case eHeading.NORTH
                    SignoNS = -1
                    SignoEO = 0
                
                Case eHeading.EAST
                    SignoNS = 0
                    SignoEO = 1
                
                Case eHeading.SOUTH
                    SignoNS = 1
                    SignoEO = 0
                
                Case eHeading.WEST
                    SignoEO = -1
                    SignoNS = 0
            End Select
            
            For Y = .Pos.Y To .Pos.Y + SignoNS * RANGO_VISION_Y Step IIf(SignoNS = 0, 1, SignoNS)
                For X = .Pos.X To .Pos.X + SignoEO * RANGO_VISION_X Step IIf(SignoEO = 0, 1, SignoEO)
                    If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
                        NI = MapData(.Pos.Map, X, Y).NpcIndex
                        If NI > 0 Then
                            If .TargetNPC = NI Then
                                bNoEsta = True
                                If .Numero = ConstantesNPCs.EleFuego Then
                                    Call NpcLanzaUnSpellSobreNpc(NpcIndex, NI)
                                    If Npclist(NI).NPCtype = DRAGON Then
                                        Call NpcLanzaUnSpellSobreNpc(NI, NpcIndex)
                                     End If
                                 Else
                                    'aca verificamosss la distancia de ataque
                                    If Distancia(.Pos, Npclist(NI).Pos) <= 1 Then
                                        Call SistemaCombate.NpcAtacaNpc(NpcIndex, NI)
                                    End If
                                 End If
                                 Exit Sub
                            End If
                       End If
                    End If
                Next X
            Next Y
        Else
            For Y = .Pos.Y - RANGO_VISION_Y To .Pos.Y + RANGO_VISION_Y
                For X = .Pos.X - RANGO_VISION_Y To .Pos.X + RANGO_VISION_Y
                    If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
                       NI = MapData(.Pos.Map, X, Y).NpcIndex
                       If NI > 0 Then
                            If .TargetNPC = NI Then
                                 bNoEsta = True
                                 If .Numero = ConstantesNPCs.EleFuego Then
                                     Call NpcLanzaUnSpellSobreNpc(NpcIndex, NI)
                                     If Npclist(NI).NPCtype = DRAGON Then
                                        Call NpcLanzaUnSpellSobreNpc(NI, NpcIndex)
                                     End If
                                 Else
                                    'aca verificamosss la distancia de ataque
                                    If Distancia(.Pos, Npclist(NI).Pos) <= 1 Then
                                        Call SistemaCombate.NpcAtacaNpc(NpcIndex, NI)
                                    End If
                                 End If
                                 Exit Sub
                            End If
                       End If
                    End If
                Next X
            Next Y
        End If
        
        If Not bNoEsta Then
            If .MaestroUser > 0 Then
                Call FollowAmo(NpcIndex)
            Else
                .Movement = .flags.OldMovement
                .Hostile = .flags.OldHostil
            End If
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub AiNpcAtacaNpc de AI_NPC.bas")
End Sub

Public Sub AiNpcObjeto(ByVal NpcIndex As Integer)
'***************************************************
'Autor: ZaMa
'Last Modification: 14/09/2009 (ZaMa)
'14/09/2009: ZaMa - Now npcs don't follow protected users.
'***************************************************
On Error GoTo ErrHandler
  
    Dim UserIndex As Integer
    Dim I As Long
    Dim UserProtected As Boolean
    
    If Not Npclist(NpcIndex).Timers.Check(TimersIndex.Hit) Then Exit Sub
    
    With Npclist(NpcIndex)
               
        Dim Query() As Collision.UUID
        For I = 0 To ModAreas.QueryEntities(NpcIndex, ENTITY_TYPE_NPC, Query, ENTITY_TYPE_PLAYER, eGridSortClosest)
            UserIndex = Query(I).Name

                    With UserList(UserIndex)
                        UserProtected = Not IntervaloPermiteSerAtacado(UserIndex) And .flags.NoPuedeSerAtacado
                        
                        If .flags.Muerto = 0 And .flags.invisible = 0 And _
                            .flags.Oculto = 0 And .flags.AdminPerseguible And Not UserProtected Then
                            
                            ' No quiero que ataque siempre al primero
                            If RandomNumber(1, 3) < 3 Then
                                If Npclist(NpcIndex).flags.LanzaSpells > 0 Then
                                     Call NpcLanzaUnSpell(NpcIndex, UserIndex)
                                End If
                            
                                Exit Sub
                            End If
                        End If
                    End With

        Next I
    End With

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub AiNpcObjeto de AI_NPC.bas")
End Sub

Sub AIAtaque(ByVal NpcIndex As Integer)
'**************************************************************
'Author: Anagrama
'Last Modify Date: 17/08/2016
'Maneja el ataque de los npcs.
'**************************************************************
On Error GoTo ErrorHandler
    With Npclist(NpcIndex)
        '<<<<<<<<<<< Ataques >>>>>>>>>>>>>>>>
        If .Timers.Check(TimersIndex.MoveAttack, False) = False Then Exit Sub
        
        If .MaestroUser = 0 Then
            'Busca a alguien para atacar
            '¿Es un guardia?
            
            If .Movement = TipoAI.BossDM Then
                Call BossDMAttack(NpcIndex)
            ElseIf .Movement = TipoAI.BossDV Then
                Call BossDVAttack(NpcIndex)
            ElseIf .Movement = TipoAI.BossDI Then
                Call BossDIAttack(NpcIndex)
            ElseIf .Movement = TipoAI.BossDA Then
                Call BossDAAttack(NpcIndex)
            ElseIf .Movement = TipoAI.BossDATentaculo Then
                Call BossDATentaculoAttack(NpcIndex)
            End If
            
            If .NPCtype = eNPCType.GuardiaReal Then
                Call GuardiasAI(NpcIndex, False)
            ElseIf .NPCtype = eNPCType.GuardiasCaos Then
                Call GuardiasAI(NpcIndex, True)
            ElseIf .Hostile And .Stats.Alineacion <> 0 Then
                Call HostilMalvadoAI(NpcIndex)
            ElseIf .Hostile And .Stats.Alineacion = 0 Then
                Call HostilBuenoAI(NpcIndex)
            End If
        Else
            If .Movement = TipoAI.NpcAtacaNpc Then
                Call AiNpcAtacaNpc(NpcIndex)
            End If
            'Evitamos que ataque a su amo, a menos
            'que el amo lo ataque.
            'Call HostilBuenoAI(NpcIndex)
        End If
    End With
    
    Exit Sub
    
ErrorHandler:
    With Npclist(NpcIndex)
        Call LogError("Error en NPCAI. Error: " & Err.Number & " - " & Err.Description & ". " & _
        "Npc: " & .Name & ", Index: " & NpcIndex & ", MaestroUser: " & .MaestroUser & _
        ", MaestroNpc: " & .MaestroNpc & ", Mapa: " & .Pos.Map & " x:" & .Pos.X & " y:" & _
        .Pos.Y & " Mov:" & .Movement & " TargU:" & _
        .Target & " TargN:" & .TargetNPC)
    End With
    
    Dim MiNPC As npc
    MiNPC = Npclist(NpcIndex)
    Call QuitarNPC(NpcIndex)
    Call ReSpawnNpc(MiNPC)
End Sub

Sub AIMovimiento(ByVal NpcIndex As Integer)
'**************************************************************
'Author: Anagrama
'Last Modify Date: 17/08/2016
'Maneja el movimiento de los npcs.
'**************************************************************
On Error GoTo ErrorHandler
    With Npclist(NpcIndex)
        
        Select Case .Movement

            Case TipoAI.MueveAlAzar
                If .flags.Inmovilizado = 1 Then Exit Sub
                If .NPCtype = eNPCType.GuardiaReal Then
                    If RandomNumber(1, 12) = 3 Then
                        Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
                    End If
                    
                    Call FollowLegion(NpcIndex)
                    
                ElseIf .NPCtype = eNPCType.GuardiasCaos Then
                    If RandomNumber(1, 12) = 3 Then
                        Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
                    End If
                    
                    Call FollowRoyal(NpcIndex)
                    
                Else
                    If RandomNumber(1, 12) = 3 Then
                        Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
                    End If
                End If

            'Va hacia el usuario cercano
            Case TipoAI.NpcMaloAtacaUsersBuenos
                Call IrUsuarioCercano(NpcIndex)
            
            'Va hacia el usuario que lo ataco(FOLLOW)
            Case TipoAI.NPCDEFENSA
                Call SeguirAgresor(NpcIndex)
            
            Case TipoAI.GuardiasAtacanCriminales
                Call FollowLegion(NpcIndex)
            
            Case TipoAI.SigueAmo
                If .flags.Inmovilizado = 1 Then Exit Sub
                Call SeguirAmo(NpcIndex)
                If RandomNumber(1, 12) = 3 Then
                    Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
                End If
            
            Case TipoAI.NpcAtacaNpc
                Call MovimientoNpcAtacaNpc(NpcIndex)
                
            Case TipoAI.NpcObjeto
                Call AiNpcObjeto(NpcIndex)
                
            Case TipoAI.NpcPathfinding
                If .flags.Inmovilizado = 1 Then Exit Sub
                If ReCalculatePath(NpcIndex) Then
                    Call PathFindingAI(NpcIndex)
                    'Existe el camino?
                    If .PFINFO.NoPath Then 'Si no existe nos movemos al azar
                        'Move randomly
                        Call MoveNPCChar(NpcIndex, RandomNumber(eHeading.NORTH, eHeading.WEST))
                    End If
                Else
                    If Not PathEnd(NpcIndex) Then
                        Call FollowPath(NpcIndex)
                    Else
                        .PFINFO.PathLenght = 0
                    End If
                End If
            
            Case TipoAI.BossDM
                Call BossDMMovement(NpcIndex)
                
            Case TipoAI.BossDV
                Call BossDVMovement(NpcIndex)
            
            Case TipoAI.BossDI
                Call BossDIMovement(NpcIndex)
        End Select
    End With
Exit Sub

ErrorHandler:
    With Npclist(NpcIndex)
        Call LogError("Error en NPCAI. Error: " & Err.Number & " - " & Err.Description & ". " & _
        "Npc: " & .Name & ", Index: " & NpcIndex & ", MaestroUser: " & .MaestroUser & _
        ", MaestroNpc: " & .MaestroNpc & ", Mapa: " & .Pos.Map & " x:" & .Pos.X & " y:" & _
        .Pos.Y & " Mov:" & .Movement & " TargU:" & _
        .Target & " TargN:" & .TargetNPC)
    End With
    
    Dim MiNPC As npc
    MiNPC = Npclist(NpcIndex)
    Call QuitarNPC(NpcIndex)
    Call ReSpawnNpc(MiNPC)
End Sub

Function UserNear(ByVal NpcIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'Returns True if there is an user adjacent to the npc position.
'***************************************************
On Error GoTo ErrHandler
  

    With Npclist(NpcIndex)
        UserNear = Not Int(Distance(.Pos.X, .Pos.Y, UserList(.PFINFO.TargetUser).Pos.X, _
                    UserList(.PFINFO.TargetUser).Pos.Y)) > 1
    End With
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function UserNear de AI_NPC.bas")
End Function

Function ReCalculatePath(ByVal NpcIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'Returns true if we have to seek a new path
'***************************************************
On Error GoTo ErrHandler
  

    If Npclist(NpcIndex).PFINFO.PathLenght = 0 Then
        ReCalculatePath = True
    ElseIf Not UserNear(NpcIndex) And Npclist(NpcIndex).PFINFO.PathLenght = Npclist(NpcIndex).PFINFO.CurPos - 1 Then
        ReCalculatePath = True
    End If
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function ReCalculatePath de AI_NPC.bas")
End Function

Function PathEnd(ByVal NpcIndex As Integer) As Boolean
'***************************************************
'Author: Gulfas Morgolock
'Last Modification: -
'Returns if the npc has arrived to the end of its path
'***************************************************
On Error GoTo ErrHandler
  
    PathEnd = Npclist(NpcIndex).PFINFO.CurPos = Npclist(NpcIndex).PFINFO.PathLenght
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function PathEnd de AI_NPC.bas")
End Function

Function FollowPath(ByVal NpcIndex As Integer) As Boolean
'***************************************************
'Author: Gulfas Morgolock
'Last Modification: -
'Moves the npc.
'***************************************************
On Error GoTo ErrHandler
  
    Dim tmpPos As WorldPos
    Dim tHeading As Byte
    
    With Npclist(NpcIndex)
        tmpPos.Map = .Pos.Map
        tmpPos.X = .PFINFO.Path(.PFINFO.CurPos).Y ' invertí las coordenadas
        tmpPos.Y = .PFINFO.Path(.PFINFO.CurPos).X
        
        'Debug.Print "(" & tmpPos.X & "," & tmpPos.Y & ")"
        
        tHeading = FindDirection(.Pos, tmpPos)
        
        MoveNPCChar NpcIndex, tHeading
        
        .PFINFO.CurPos = .PFINFO.CurPos + 1
    End With
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function FollowPath de AI_NPC.bas")
End Function

Function PathFindingAI(ByVal NpcIndex As Integer) As Boolean
'***************************************************
'Author: Gulfas Morgolock
'Last Modification: -
'This function seeks the shortest path from the Npc
'to the user's location.
'***************************************************
On Error GoTo ErrHandler
  
    Dim Y As Long
    Dim X As Long
    
    With Npclist(NpcIndex)
        For Y = .Pos.Y - 10 To .Pos.Y + 10    'Makes a loop that looks at
             For X = .Pos.X - 10 To .Pos.X + 10   '5 tiles in every direction
                
                 'Make sure tile is legal
                 If X > MinXBorder And X < MaxXBorder And Y > MinYBorder And Y < MaxYBorder Then
                    
                     'look for a user
                     If MapData(.Pos.Map, X, Y).UserIndex > 0 Then
                         'Move towards user
                          Dim tmpUserIndex As Integer
                          tmpUserIndex = MapData(.Pos.Map, X, Y).UserIndex
                          With UserList(tmpUserIndex)
                            If .flags.Muerto = 0 And .flags.invisible = 0 And .flags.Oculto = 0 And .flags.AdminPerseguible Then
                                'We have to invert the coordinates, this is because
                                'ORE refers to maps in converse way of my pathfinding
                                'routines.
                                Npclist(NpcIndex).PFINFO.Target.X = .Pos.Y
                                Npclist(NpcIndex).PFINFO.Target.Y = .Pos.X 'ops!
                                Npclist(NpcIndex).PFINFO.TargetUser = tmpUserIndex
                                Call SeekPath(NpcIndex)
                                Exit Function
                            End If
                        End With
                    End If
                End If
            Next X
        Next Y
    End With
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function PathFindingAI de AI_NPC.bas")
End Function

Sub NpcLanzaUnSpell(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
'**************************************************************
'Author: Unknown
'Last Modify by: -
'Last Modify Date: -
'**************************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)
        If .flags.invisible = 1 Or .flags.Oculto = 1 Then Exit Sub
    End With
    
    Dim k As Integer
    k = RandomNumber(1, Npclist(NpcIndex).flags.LanzaSpells)
    Call NpcLanzaSpellSobreUser(NpcIndex, UserIndex, Npclist(NpcIndex).Spells(k))
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub NpcLanzaUnSpell de AI_NPC.bas")
End Sub

Sub NpcLanzaUnSpellSobreNpc(ByVal NpcIndex As Integer, ByVal TargetNPC As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    Dim k As Integer
    k = RandomNumber(1, Npclist(NpcIndex).flags.LanzaSpells)
    Call NpcLanzaSpellSobreNpc(NpcIndex, TargetNPC, Npclist(NpcIndex).Spells(k))
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub NpcLanzaUnSpellSobreNpc de AI_NPC.bas")
End Sub
