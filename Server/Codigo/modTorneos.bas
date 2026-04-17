Attribute VB_Name = "modTorneos"
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

Public Enum eTournamentState
    ieNone
    ieRegistered
    ieWaitingForFight
    ieFighting
End Enum

Public Enum eTournamentEdit
    ieMaxCompetitor
    ieMaxLevel
    ieMinLevel
    ieRequiredGold
    ieForbiddenItems
    iePermitedClass
    ieNumRoundsToWin
    ieKillAfterLoose
    ieWaitingMap
    ieArenaPosition
    ieFinalMap
    ieSaveConfig
End Enum

Public Enum eTournamentExpellMotive
    ieAbandon
    ieExpelled
    ieLose
    ieMassiveExpell
End Enum

Public Const MAX_ARENAS As Byte = 5

Private Type tArena
    ' Default positions
    Map As Integer
    UserPos1 As WorldPos
    UserPos2 As WorldPos
    
    CompetitorIndex1 As Integer
    CompetitorIndex2 As Integer
    
    Active As Boolean
End Type

Private Type tTournament
    
    ' Competitors
    MaxCompetitors As Byte
    CompetitorsList As cCola
    
    ' Restrictions
    MinLevel As Byte
    MaxLevel As Byte
    RequiredGold As Long
    
    NumForbiddenItems As Byte
    ForbiddenItem() As Integer
    
    NumPermitedClass As Byte
    PermitedClass() As Byte
    
    ' Aditionals
    NumRoundsToWin As Byte
    
    ' Counters
    RegistrationCountdown As Byte
    FightCountdown As Byte
    
    ' Flags
    RegistrationOpen As Boolean
    CountdownActivated As Boolean
    Active As Boolean
    PreparingArena As Byte
    KillAfterLoose As Byte
    
    ' Positions
    WaitingMap As WorldPos
    FinalMap As WorldPos
    Arenas(1 To MAX_ARENAS) As tArena
End Type

' Public holder
Public Tournament As tTournament

Public Sub TournamentCountdownCheck()
'***************************************************
'Author: ZaMa
'Last Modification: 31/05/2012
'Update clients' countdown when automatic tournament activated
'***************************************************
On Error GoTo ErrHandler
  
    
    With Tournament
        If .FightCountdown <> 0 Then
            If .FightCountdown > 1 Then
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("> " & CStr(.FightCountdown - 1), FontTypeNames.FONTTYPE_CITIZEN))
            Else
                ' Update users tournament state
                UserList(.Arenas(.PreparingArena).CompetitorIndex1).flags.TournamentState = eTournamentState.ieFighting
                UserList(.Arenas(.PreparingArena).CompetitorIndex2).flags.TournamentState = eTournamentState.ieFighting
            
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("YA!!", FontTypeNames.FONTTYPE_CITIZEN))
                .CountdownActivated = False
                .PreparingArena = 0
            End If
            
            .FightCountdown = .FightCountdown - 1
        End If
            
        If .RegistrationCountdown <> 0 Then
            If .RegistrationCountdown > 1 Then
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Inscripciones al Torneo abiertas en ... " & CStr(.RegistrationCountdown - 1), FontTypeNames.FONTTYPE_CITIZEN))
            Else
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Inscripciones abiertas.", FontTypeNames.FONTTYPE_CITIZEN))
                .RegistrationOpen = True
                .CountdownActivated = False
            End If
            .RegistrationCountdown = .RegistrationCountdown - 1
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub TournamentCountdownCheck de modTorneos.bas")
End Sub

Public Function UserCanRegisterIntoTorunament(ByVal UserIndex As Integer) As Boolean
'***************************************************
'Author: ZaMa
'Last Modification: 31/05/2012
'Validates if user can register into tournament
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex)
        
        ' Tournament available?
        If Not Tournament.Active Then
            WriteConsoleMsg UserIndex, "¡¡¡No hay ningún torneo activo!!!", FontTypeNames.FONTTYPE_INFO
            Exit Function
        End If
        
        ' Registration available?
        If Not Tournament.RegistrationOpen Then
            WriteConsoleMsg UserIndex, "¡¡¡Las inscripciones están cerradas!!!", FontTypeNames.FONTTYPE_INFO
            Exit Function
        End If
        
        ' Registration countdown?
        If Tournament.RegistrationCountdown <> 0 Then
            WriteConsoleMsg UserIndex, "¡¡¡Las inscripciones no se han abierto aún!!!", FontTypeNames.FONTTYPE_INFO
            Exit Function
        End If
        
        ' Is alive?
        If .flags.Muerto = 1 Then
            WriteConsoleMsg UserIndex, "¡¡¡Estas muerto!!!", FontTypeNames.FONTTYPE_INFO
            Exit Function
        End If
     
        ' Level min?
        If Tournament.MinLevel <> 0 Then
            If .Stats.ELV < Tournament.MinLevel Then
                WriteConsoleMsg UserIndex, "¡¡¡Tu nivel es menor al requerido: " & Tournament.MinLevel & ", no puedes participar!!!", FontTypeNames.FONTTYPE_INFO
                Exit Function
            End If
        End If
        
        ' Level max?
        If Tournament.MaxLevel <> 0 Then
            If .Stats.ELV > Tournament.MaxLevel Then
                WriteConsoleMsg UserIndex, "¡¡¡Tu nivel es mayor al requerido: " & Tournament.MaxLevel & ", no puedes participar!!!", FontTypeNames.FONTTYPE_INFO
                Exit Function
            End If
        End If
        
        ' Required gold?
        If .Stats.GLD < Tournament.RequiredGold Then
            If .Stats.Banco < Tournament.RequiredGold Then
                WriteConsoleMsg UserIndex, "¡¡¡No tienes oro suficiente, se requiere: " & Tournament.RequiredGold & ", no puedes participar!!!", FontTypeNames.FONTTYPE_INFO
                Exit Function
            End If
        End If
        
        ' Has a forbidden item?
        Dim Counter As Long
        Dim SlotCounter As Long
        Dim HasForbiddenItem As Boolean
        
        For Counter = 1 To Tournament.NumForbiddenItems
            For SlotCounter = 1 To .CurrentInventorySlots
                If .Invent.Object(SlotCounter).ObjIndex = Tournament.ForbiddenItem(Counter) Then
                    HasForbiddenItem = True
                    Exit For
                End If
            Next SlotCounter
            
            If HasForbiddenItem Then Exit For
        Next Counter
        
        If HasForbiddenItem Then
            WriteConsoleMsg UserIndex, "¡¡¡No puedes participar si tienes un item prohibido!!!", FontTypeNames.FONTTYPE_INFO
            Exit Function
        End If
        
        ' Permited class?
        Dim UserClass As Byte
        UserClass = .clase
        
        For Counter = 1 To Tournament.NumPermitedClass
            If UserClass = Tournament.PermitedClass(Counter) Then Exit For
        Next Counter
        
        If Counter > Tournament.NumPermitedClass Then
            WriteConsoleMsg UserIndex, "¡¡¡Tu clase no puede participar del torneo!!!", FontTypeNames.FONTTYPE_INFO
            Exit Function
        End If
        
        
        ' In Prision?
        If .Pos.Map = Prision.Map Then
            WriteConsoleMsg UserIndex, "¡¡¡No puedes participar si estas en la carcel!!!", FontTypeNames.FONTTYPE_INFO
            Exit Function
        End If
       
        ' Already registered?
        If Tournament.CompetitorsList.Existe(.Name) Then
            WriteConsoleMsg UserIndex, "¡¡¡Ya te has registrado en el torneo!!!", FontTypeNames.FONTTYPE_INFO
            Exit Function
        End If
        
    End With

    UserCanRegisterIntoTorunament = True
    
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function UserCanRegisterIntoTorunament de modTorneos.bas")
End Function

Public Sub RegisterUserToTournament(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 31/05/2012
'Register user into tournament and warps it to wating map
'***************************************************
    ' Register
On Error GoTo ErrHandler
  
    Call WriteConsoleMsg(UserIndex, "Te has registrado en el torneo.", FontTypeNames.FONTTYPE_VENENO)
    Call Tournament.CompetitorsList.Push(UserList(UserIndex).Name)
    
    ' Warp user
    With Tournament.WaitingMap
        Call WarpUserChar(UserIndex, .Map, .X, .Y, True)
    End With
    
    With UserList(UserIndex)
        ' Update state
        .flags.TournamentState = eTournamentState.ieRegistered
    
        ' Required gold?
        If Tournament.RequiredGold > 0 Then
            If .Stats.GLD >= Tournament.RequiredGold Then
                .Stats.GLD = .Stats.GLD - Tournament.RequiredGold
                Call WriteUpdateGold(UserIndex)
            Else
                .Stats.Banco = .Stats.Banco - Tournament.RequiredGold
                Call WriteUpdateBankGold(UserIndex)
            End If
        End If
    End With
    
    ' Close registrations
    If Tournament.CompetitorsList.Longitud = Tournament.MaxCompetitors Then
        Tournament.RegistrationOpen = False
    End If
        
    Call WriteConsoleMsg(UserIndex, "Has sido teletransportado a la Sala de Espera, aguarda tu turno.", FontTypeNames.FONTTYPE_INFO)

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub RegisterUserToTournament de modTorneos.bas")
End Sub

Public Sub SendCompetitorsList(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 31/05/2012
'Sends remaining competitor list to given user
'***************************************************
On Error GoTo ErrHandler
  
    
    ' Active?
    If Not Tournament.Active Then
        Call WriteConsoleMsg(UserIndex, "No hay ningun torneo activo.", FontTypeNames.FONTTYPE_SERVER)
        Exit Sub
    End If
    
    If Tournament.CompetitorsList.Longitud = 0 Then
        Call WriteConsoleMsg(UserIndex, "No hay ningun inscripto en el torneo.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    Call WriteTournamentCompetitorList(UserIndex)

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SendCompetitorsList de modTorneos.bas")
End Sub

Public Sub TournamentFightBegin(ByVal UserIndex As Integer, ByRef Competitor1 As String, ByRef Competitor2 As String, _
    ByVal ArenaIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 31/05/2012
'Checks if fight can start and reset fight countdown.
'***************************************************
On Error GoTo ErrHandler
  
    
    With Tournament
        ' Torunament active?
        If .Active Then
        
            ' Valid index?
            If ArenaIndex <= 0 Or ArenaIndex > MAX_ARENAS Then
                Call WriteConsoleMsg(UserIndex, "Índice de la arena inválido. Rango: 1-" & MAX_ARENAS & ".", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
            ' In use?
            If .Arenas(ArenaIndex).Active Then
                Call WriteConsoleMsg(UserIndex, "Ya hay usuarios utilizando la arena " & ArenaIndex & ".", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            ' Any arena countdown?
            If .PreparingArena <> 0 Then
                Call WriteConsoleMsg(UserIndex, "Se esta preparando la arena " & .PreparingArena & ", aguarda un momento.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            ' Is competing user1?
            If Not .CompetitorsList.Existe(Competitor1) Then
                Call WriteConsoleMsg(UserIndex, "El usuario " & Competitor1 & " no esta compitiendo en este torneo.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            ' Conected?
            Dim CompetitorIndex1 As Integer
            CompetitorIndex1 = NameIndex(Competitor1)
            
            If CompetitorIndex1 = 0 Then
                Call WriteConsoleMsg(UserIndex, "El usuario " & Competitor1 & " no esta conectado.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            ' Already fighting?
            If UserList(CompetitorIndex1).flags.TournamentState <> eTournamentState.ieRegistered Then
                Call WriteConsoleMsg(UserIndex, "El usuario " & Competitor1 & " ya esta peleando.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            ' Is competing user2?
            If Not .CompetitorsList.Existe(Competitor2) Then
                Call WriteConsoleMsg(UserIndex, "El usuario " & Competitor2 & " no esta compitiendo en este torneo.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            ' Conected?
            Dim CompetitorIndex2 As Integer
            CompetitorIndex2 = NameIndex(Competitor2)
            
            If CompetitorIndex2 = 0 Then
                Call WriteConsoleMsg(UserIndex, "El usuario " & Competitor2 & " no esta conectado.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            ' Already fighting?
            If UserList(CompetitorIndex2).flags.TournamentState <> eTournamentState.ieRegistered Then
                Call WriteConsoleMsg(UserIndex, "El usuario " & Competitor2 & " ya esta peleando.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            ' Teleport to arena
            With .Arenas(ArenaIndex)
                Call WarpUserChar(CompetitorIndex1, .Map, .UserPos1.X, .UserPos1.Y, True)
                Call WarpUserChar(CompetitorIndex2, .Map, .UserPos2.X, .UserPos2.Y, True)
                .CompetitorIndex1 = CompetitorIndex1
                .CompetitorIndex2 = CompetitorIndex2
            End With
            
            ' Update users tournament state
            UserList(CompetitorIndex1).flags.TournamentState = eTournamentState.ieWaitingForFight
            UserList(CompetitorIndex2).flags.TournamentState = eTournamentState.ieWaitingForFight
            
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Torneo>> " & Competitor1 & " VS " & Competitor2 & " en la arena " & ArenaIndex & ".", FontTypeNames.FONTTYPE_CONSEJO))
            .FightCountdown = 5
            .CountdownActivated = True
            .PreparingArena = ArenaIndex
        Else
            Call WriteConsoleMsg(UserIndex, "No hay ningún torneo activo.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub TournamentFightBegin de modTorneos.bas")
End Sub

Public Sub TournamentUserExpell(ByVal UserIndex As Integer, ByVal Motive As Byte)
'***************************************************
'Author: ZaMa
'Last Modification: 31/05/2012
'Expelled user from tournament if closes, looses or the torunament is cancelled.
'Also drop user's items in final map if that's how is configured.
'***************************************************
On Error GoTo ErrHandler
  

'TODO_TORNEO:

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub TournamentUserExpell de modTorneos.bas")
End Sub

Private Function GetArena(ByVal UserIndex As Integer) As Integer
'***************************************************
'Author: ZaMa
'Last Modification: 31/05/2012
'Returns the arena wich correspond to given userindex
'***************************************************
On Error GoTo ErrHandler
  
    
    With Tournament
        Dim ArenaIndex As Long
        For ArenaIndex = 1 To MAX_ARENAS
            If .Arenas(ArenaIndex).Active Then
                If .Arenas(ArenaIndex).CompetitorIndex1 = UserIndex Or .Arenas(ArenaIndex).CompetitorIndex2 = UserIndex Then
                    GetArena = CInt(ArenaIndex)
                    Exit Function
                End If
            End If
        Next ArenaIndex
    End With
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetArena de modTorneos.bas")
End Function

