Attribute VB_Name = "modDuelos"

Option Explicit

Public MapArenas() As Integer

Public Type tDuelTeams
    Player(1 To 4) As Integer
    Muerto(1 To 4) As Boolean
End Type

Public Enum eDuelType
    vs1
    vs2
    vs3
    vs4
End Enum

Public Enum eDuelState
    Vacio
    Esperando_Jugadores
    Esperando_Inicio
    Iniciado
    Esperando_Final
    Finalizado
End Enum

Public Type tArena
    EnUso As Byte
    Map As Integer
    X1 As Byte
    X2 As Byte
    Y1 As Byte
    Y2 As Byte
End Type

Public Arena1v1() As tArena
Public Arena2v2() As tArena
Public Arena3v3() As tArena
Public Arena4v4() As tArena

Public Type tDuel
    Arena As Byte
    TipoDuelo As eDuelType
    estado As eDuelState
    Oro As Long
    Drop As Boolean
    Resucitar As Boolean
    Counter As Integer
    Team(1 To 2) As tDuelTeams
    Ganador As Byte
End Type

Public Type tDuelData
    Duelo() As tDuel
    On As Boolean
    MinLevel As Byte
    MinGold As Long
    MinGoldPublic As Long
    DuelDuration As Long
End Type

Public DuelData As tDuelData

Public PublicDuel() As Integer

Public Sub LoadDuelData()
On Error GoTo ErrHandler

    Dim I As Byte
    Dim NumArenas As Byte
    Dim tmpArenas() As String
    
    DuelData.On = CBool(GetVar(DatPath & "Duelos.dat", "INIT", "Activados"))
    DuelData.MinLevel = Val(GetVar(DatPath & "Duelos.dat", "INIT", "MinLevel"))
    DuelData.MinGold = Val(GetVar(DatPath & "Duelos.dat", "INIT", "MinGold"))
    DuelData.MinGoldPublic = Val(GetVar(DatPath & "Duelos.dat", "INIT", "MinGoldPublic"))
    DuelData.DuelDuration = Val(GetVar(DatPath & "Duelos.dat", "INIT", "DuelDuration"))
    
    NumArenas = Val(GetVar(DatPath & "Duelos.dat", "INIT", "NumArenas1v1"))
    ReDim Arena1v1(1 To NumArenas) As tArena
    
    For I = 1 To NumArenas
        Arena1v1(I).Map = Val(GetVar(DatPath & "Duelos.dat", "Arena1v1-" & I, "Mapa"))
        Arena1v1(I).X1 = Val(GetVar(DatPath & "Duelos.dat", "Arena1v1-" & I, "X1"))
        Arena1v1(I).X2 = Val(GetVar(DatPath & "Duelos.dat", "Arena1v1-" & I, "X2"))
        Arena1v1(I).Y1 = Val(GetVar(DatPath & "Duelos.dat", "Arena1v1-" & I, "Y1"))
        Arena1v1(I).Y2 = Val(GetVar(DatPath & "Duelos.dat", "Arena1v1-" & I, "Y2"))
    Next I
    
    NumArenas = Val(GetVar(DatPath & "Duelos.dat", "INIT", "NumArenas2v2"))
    ReDim Arena2v2(1 To NumArenas) As tArena
    
    For I = 1 To NumArenas
        Arena2v2(I).Map = Val(GetVar(DatPath & "Duelos.dat", "Arena2v2-" & I, "Mapa"))
        Arena2v2(I).X1 = Val(GetVar(DatPath & "Duelos.dat", "Arena2v2-" & I, "X1"))
        Arena2v2(I).X2 = Val(GetVar(DatPath & "Duelos.dat", "Arena2v2-" & I, "X2"))
        Arena2v2(I).Y1 = Val(GetVar(DatPath & "Duelos.dat", "Arena2v2-" & I, "Y1"))
        Arena2v2(I).Y2 = Val(GetVar(DatPath & "Duelos.dat", "Arena2v2-" & I, "Y2"))
    Next I
    
    NumArenas = Val(GetVar(DatPath & "Duelos.dat", "INIT", "NumArenas3v3"))
    ReDim Arena3v3(1 To NumArenas) As tArena
    
    For I = 1 To NumArenas
        Arena3v3(I).Map = Val(GetVar(DatPath & "Duelos.dat", "Arena3v3-" & I, "Mapa"))
        Arena3v3(I).X1 = Val(GetVar(DatPath & "Duelos.dat", "Arena3v3-" & I, "X1"))
        Arena3v3(I).X2 = Val(GetVar(DatPath & "Duelos.dat", "Arena3v3-" & I, "X2"))
        Arena3v3(I).Y1 = Val(GetVar(DatPath & "Duelos.dat", "Arena3v3-" & I, "Y1"))
        Arena3v3(I).Y2 = Val(GetVar(DatPath & "Duelos.dat", "Arena3v3-" & I, "Y2"))
    Next I
    
    NumArenas = Val(GetVar(DatPath & "Duelos.dat", "INIT", "NumArenas4v4"))
    ReDim Arena4v4(1 To NumArenas) As tArena
    
    For I = 1 To NumArenas
        Arena4v4(I).Map = Val(GetVar(DatPath & "Duelos.dat", "Arena4v4-" & I, "Mapa"))
        Arena4v4(I).X1 = Val(GetVar(DatPath & "Duelos.dat", "Arena4v4-" & I, "X1"))
        Arena4v4(I).X2 = Val(GetVar(DatPath & "Duelos.dat", "Arena4v4-" & I, "X2"))
        Arena4v4(I).Y1 = Val(GetVar(DatPath & "Duelos.dat", "Arena4v4-" & I, "Y1"))
        Arena4v4(I).Y2 = Val(GetVar(DatPath & "Duelos.dat", "Arena4v4-" & I, "Y2"))
    Next I
    
    ReDim PublicDuel(1 To 1) As Integer
    ReDim DuelData.Duelo(1 To UBound(Arena1v1) + UBound(Arena2v2) + UBound(Arena3v3) + UBound(Arena4v4)) As tDuel
    
    tmpArenas = Split(GetVar(DatPath & "Duelos.dat", "INIT", "MapArenas"), "-")
    ReDim MapArenas(LBound(tmpArenas) To UBound(tmpArenas)) As Integer
    
    For I = LBound(tmpArenas) To UBound(tmpArenas)
        MapArenas(I) = Val(tmpArenas(I))
    Next I
    
    frmMain.TimerDuelos.Enabled = True
    
    Exit Sub
    
ErrHandler:
    DuelData.On = False
    frmMain.TimerDuelos.Enabled = False
    Call LogError("Error " & Err.Number & " (" & Err.Description & ") en Sub LoadDuelData del Módulo modDuelos")
End Sub

Public Sub CancelarEsperaDuelo(ByVal UserIndex As Integer)
On Error GoTo ErrHandler
  
    If UserList(UserIndex).flags.DueloPublico = 0 Then Exit Sub
    
    PublicDuel(UserList(UserIndex).flags.DueloPublico) = 0
    UserList(UserIndex).flags.DueloPublico = 0
    
    Call WriteOkDueloPublico(UserIndex)
    Call WriteParalizeOK(UserIndex)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CancelarEsperaDuelo de modDuelos.bas")
End Sub

Public Sub EmpezarDueloPublico(ByVal Duelo1 As Byte, ByVal Duelo2 As Byte)
On Error GoTo ErrHandler
  
    Dim MySlot As Byte
    
    MySlot = GetNewDueloSlot(eDuelType.vs1)
    If Not MySlot > 0 Then
        Call WriteConsoleMsg(PublicDuel(Duelo1), "Todas las Arenas para Duelos 1vs1 están ocupadas, por favor espera.", FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(PublicDuel(Duelo2), "Todas las Arenas para Duelos 1vs1 están ocupadas, por favor espera.", FontTypeNames.FONTTYPE_INFO)
        UserList(PublicDuel(Duelo1)).flags.DueloPublico = 1
        UserList(PublicDuel(Duelo2)).flags.DueloPublico = 1
        Exit Sub
    End If
    
    UserList(PublicDuel(Duelo2)).flags.DueloPublico = 0
    Call GiveDueloSlot(eDuelType.vs1, MySlot, PublicDuel(Duelo1))
    Call SetDueloConfig(MySlot, eDuelType.vs1, False, True, DuelData.MinGoldPublic)
    
    UserList(PublicDuel(Duelo2)).flags.DueloIndex = MySlot
    UserList(PublicDuel(Duelo2)).flags.DueloTeam = 2
    
    Call AssignTeamMember(MySlot, UserList(PublicDuel(Duelo2)).flags.DueloTeam, PublicDuel(Duelo2))
    Call WriteOkDueloPublico(PublicDuel(Duelo1))
    Call WriteParalizeOK(PublicDuel(Duelo1))
    Call WriteOkDueloPublico(PublicDuel(Duelo2))
    Call WriteParalizeOK(PublicDuel(Duelo2))
    
    PublicDuel(Duelo1) = 0
    PublicDuel(Duelo2) = 0
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub EmpezarDueloPublico de modDuelos.bas")
End Sub

Public Sub BuscarContrincante(ByVal DuelIndex As Byte)
On Error GoTo ErrHandler
  
    Dim I As Byte
    Dim a As Integer
    Dim b As Integer
    
    If DuelIndex < 1 Or DuelIndex > UBound(PublicDuel) Then Exit Sub
    
    For I = 1 To UBound(PublicDuel)
        If PublicDuel(I) > 0 Then
            If I <> DuelIndex Then
                If (UserList(PublicDuel(DuelIndex)).Stats.MaxMan = 0 And UserList(PublicDuel(I)).Stats.MaxMan = 0) _
                    Or (UserList(PublicDuel(DuelIndex)).Stats.MaxMan > 0 And UserList(PublicDuel(I)).Stats.MaxMan > 0) Then
                    a = UserList(PublicDuel(DuelIndex)).Stats.ELV
                    b = UserList(PublicDuel(I)).Stats.ELV
                    If Abs(a - b) <= 3 Then
                        UserList(PublicDuel(DuelIndex)).flags.DueloPublico = 0
                        UserList(PublicDuel(I)).flags.DueloPublico = 0
                        If PuedeDuelo(PublicDuel(DuelIndex), DuelData.MinGoldPublic) Then
                            If PuedeParticiparDuelo(PublicDuel(DuelIndex), PublicDuel(I), DuelData.MinGoldPublic) Then
                                Call EmpezarDueloPublico(DuelIndex, I)
                                Exit Sub
                            End If
                            UserList(PublicDuel(DuelIndex)).flags.DueloPublico = 1
                        End If
                        UserList(PublicDuel(I)).flags.DueloPublico = 1
                    End If
                End If
            End If
        End If
    Next I
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub BuscarContrincante de modDuelos.bas")
End Sub

Public Sub IngresarDueloPublico(ByVal UserIndex As Integer)
On Error GoTo ErrHandler
  
    If PuedeDuelo(UserIndex, DuelData.MinGoldPublic) Then
        Call WriteOkDueloPublico(UserIndex)
        Call WriteParalizeOK(UserIndex)
        UserList(UserIndex).flags.DueloPublico = NextOpenDuel(UserIndex)
        Call BuscarContrincante(UserList(UserIndex).flags.DueloPublico)
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub IngresarDueloPublico de modDuelos.bas")
End Sub

Public Function NextOpenDuel(ByVal UserIndex As Integer) As Integer
On Error GoTo ErrHandler
  
    Dim I As Byte

    For I = 1 To UBound(PublicDuel)
        If PublicDuel(I) = 0 Then
            PublicDuel(I) = UserIndex
            NextOpenDuel = I
            Exit Function
        End If
    Next I
    
    ReDim Preserve PublicDuel(1 To UBound(PublicDuel) + 1) As Integer
    PublicDuel(UBound(PublicDuel)) = UserIndex
    NextOpenDuel = UBound(PublicDuel)
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function NextOpenDuel de modDuelos.bas")
End Function

Function PuedeDuelo(ByVal Retador As Integer, ByVal Oro As Long) As Boolean
On Error GoTo ErrHandler
  
    If Not DuelData.On Then
        PuedeDuelo = False
        Call WriteConsoleMsg(Retador, "El sistema de duelos se encuentra desactivado.", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    If Not TieneOro(Retador, Oro) Then
        PuedeDuelo = False
        Call WriteConsoleMsg(Retador, "No posees el oro por el que intentas competir.", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    If UserList(Retador).flags.Muerto = 1 Then
        PuedeDuelo = False
        Call WriteConsoleMsg(Retador, "¡¡Estas Muerto!!", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    If UserList(Retador).flags.invisible = 1 Then
        PuedeDuelo = False
        Call WriteConsoleMsg(Retador, "No puedes retar a alguien a duelo estando invisible.", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    If UserList(Retador).flags.Mimetizado = 1 Then
        PuedeDuelo = False
        Call WriteConsoleMsg(Retador, "No puedes retar a alguien a duelo estando mimetizado.", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    If UserList(Retador).Pos.Map <> 1 And UserList(Retador).Pos.Map <> 171 Then
        PuedeDuelo = False
        Call WriteConsoleMsg(Retador, "Solo puedes retar a alguien a duelo desde Ullathorpe o Arena.", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    If UserList(Retador).flags.DueloIndex > 0 Then
        PuedeDuelo = False
        Call WriteConsoleMsg(Retador, "¡Ya estás en un duelo!", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    If UserList(Retador).flags.DueloPublico > 0 Then
        PuedeDuelo = False
        Call WriteConsoleMsg(Retador, "¡Ya estás en la lista de espera!", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    If Not UserList(Retador).Stats.ELV >= DuelData.MinLevel Then
        PuedeDuelo = False
        Call WriteConsoleMsg(Retador, "Necesitas ser al menos nivel " & DuelData.MinLevel & " para retar a un duelo.", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    PuedeDuelo = True
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function PuedeDuelo de modDuelos.bas")
End Function


Function PuedeParticiparDuelo(ByVal Retador As Integer, ByVal Retado As Integer, ByVal Oro As Long) As Boolean
On Error GoTo ErrHandler
  
    If Not TieneOro(Retado, Oro) Then
        PuedeParticiparDuelo = False
        Call WriteConsoleMsg(Retador, "Tu oponente no tiene el oro que deseas apostar.", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    If UserList(Retado).flags.Muerto = 1 Then
        PuedeParticiparDuelo = False
        Call WriteConsoleMsg(Retador, "¡¡Tu oponente esta muerto!!", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    If UserList(Retado).flags.invisible = 1 Then
        PuedeParticiparDuelo = False
        Call WriteConsoleMsg(Retador, "No puedes retar a alguien a duelo invisible.", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    If UserList(Retado).flags.Mimetizado = 1 Then
        PuedeParticiparDuelo = False
        Call WriteConsoleMsg(Retador, "No puedes retar a alguien a duelo que esta mimetizado.", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    If UserList(Retado).Pos.Map <> 1 And UserList(Retado).Pos.Map <> 171 Then
        PuedeParticiparDuelo = False
        Call WriteConsoleMsg(Retador, "No puedes retar a alguien a duelo fuera de Ullathorpe o Arena.", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    If UserList(Retado).flags.DueloIndex > 0 And Not UserList(Retado).flags.DueloIndex = UserList(Retador).flags.DueloIndex Then
        PuedeParticiparDuelo = False
        Call WriteConsoleMsg(Retador, "¡Tu oponente ya está en un duelo!", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    If UserList(Retado).flags.DueloPublico > 0 Then
        PuedeParticiparDuelo = False
        Call WriteConsoleMsg(Retador, "¡Tu oponente está en la lista de espera!", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    If Not UserList(Retado).Stats.ELV >= DuelData.MinLevel Then
        PuedeParticiparDuelo = False
        Call WriteConsoleMsg(Retador, "Tu oponente debe ser al menos nivel " & DuelData.MinLevel & " para retar a duelo.", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    PuedeParticiparDuelo = True
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function PuedeParticiparDuelo de modDuelos.bas")
End Function

Function PuedeAceptarDuelo(ByVal Retado As Integer, ByVal Slot As Byte) As Boolean
On Error GoTo ErrHandler
  
    If Not TieneOro(Retado, DuelData.Duelo(Slot).Oro) Then
        PuedeAceptarDuelo = False
        Call WriteConsoleMsg(Retado, "No tienes el oro que deseas apostar.", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    If UserList(Retado).flags.Muerto = 1 Then
        PuedeAceptarDuelo = False
        Call WriteConsoleMsg(Retado, "¡¡Estás Muerto!!", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    If UserList(Retado).flags.invisible = 1 Then
        PuedeAceptarDuelo = False
        Call WriteConsoleMsg(Retado, "No puedes aceptar un duelo estando invisible.", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    If UserList(Retado).flags.Comerciando Then
        PuedeAceptarDuelo = False
        Call WriteConsoleMsg(Retado, "No puedes aceptar un duelo mientras comercias.", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    If UserList(Retado).flags.HelpMode Then
        PuedeAceptarDuelo = False
        Call WriteConsoleMsg(Retado, "No puedes aceptar un duelo mientras estas en consulta.", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    If UserList(Retado).Pos.Map <> 1 And UserList(Retado).Pos.Map <> 171 Then
        PuedeAceptarDuelo = False
        Call WriteConsoleMsg(Retado, "No puedes aceptar un duelo fuera de Ullathorpe o Arena.", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    If UserList(Retado).flags.DueloIndex > 0 And Not UserList(Retado).flags.DueloIndex = Slot Then
        PuedeAceptarDuelo = False
        Call WriteConsoleMsg(Retado, "¡Ya estás en un duelo!", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    If UserList(Retado).flags.DueloPublico > 0 Then
        PuedeAceptarDuelo = False
        Call WriteConsoleMsg(Retado, "¡Ya estás en la lista de espera!", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    If Not UserList(Retado).Stats.ELV >= DuelData.MinLevel Then
        PuedeAceptarDuelo = False
        Call WriteConsoleMsg(Retado, "Necesitas ser al menos Nivel " & DuelData.MinLevel & " para aceptar un Duelo.", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    
    PuedeAceptarDuelo = True
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function PuedeAceptarDuelo de modDuelos.bas")
End Function

Function TieneOro(ByVal UserIndex As Integer, ByVal Oro As Long) As Boolean
On Error GoTo ErrHandler
  
    If Not UserList(UserIndex).Stats.GLD >= Oro Then
        TieneOro = False
        Exit Function
    End If
    
    TieneOro = True
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function TieneOro de modDuelos.bas")
End Function

Function GetNewDueloSlot(ByVal TipoDuelo As eDuelType) As Byte
On Error GoTo ErrHandler
  
    Dim I As Byte
    If TipoDuelo = eDuelType.vs1 Then
        For I = 1 To UBound(Arena1v1)
            If DuelData.Duelo(I).estado = eDuelState.Vacio Then
                GetNewDueloSlot = I
                Exit Function
            End If
        Next I
    ElseIf TipoDuelo = eDuelType.vs2 Then
        For I = UBound(Arena1v1) + 1 To UBound(Arena1v1) + UBound(Arena2v2)
            If DuelData.Duelo(I).estado = eDuelState.Vacio Then
                GetNewDueloSlot = I
                Exit Function
            End If
        Next I
    ElseIf TipoDuelo = eDuelType.vs3 Then
        For I = UBound(Arena1v1) + UBound(Arena2v2) + 1 To UBound(Arena1v1) + UBound(Arena2v2) + UBound(Arena3v3)
            If DuelData.Duelo(I).estado = eDuelState.Vacio Then
                GetNewDueloSlot = I
                Exit Function
            End If
        Next I
    ElseIf TipoDuelo = eDuelType.vs4 Then
        For I = UBound(Arena1v1) + UBound(Arena2v2) + UBound(Arena3v3) + 1 To UBound(Arena1v1) + UBound(Arena2v2) + UBound(Arena3v3) + UBound(Arena4v4)
            If DuelData.Duelo(I).estado = eDuelState.Vacio Then
                GetNewDueloSlot = I
                Exit Function
            End If
        Next I
    End If
    
    GetNewDueloSlot = 0
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetNewDueloSlot de modDuelos.bas")
End Function

Sub GiveDueloSlot(ByVal TipoDuelo As eDuelType, ByVal Slot As Byte, ByVal Retador As Integer)
On Error GoTo ErrHandler
  
    If Not Slot > 0 Then Exit Sub
    DuelData.Duelo(Slot).TipoDuelo = TipoDuelo
    DuelData.Duelo(Slot).Counter = 300
    DuelData.Duelo(Slot).estado = eDuelState.Esperando_Jugadores
    DuelData.Duelo(Slot).Arena = NextOpenArena(TipoDuelo)
    If TipoDuelo = eDuelType.vs1 Then
        Arena1v1(DuelData.Duelo(Slot).Arena).EnUso = 1
    ElseIf TipoDuelo = eDuelType.vs2 Then
        Arena2v2(DuelData.Duelo(Slot).Arena).EnUso = 1
    ElseIf TipoDuelo = eDuelType.vs3 Then
        Arena3v3(DuelData.Duelo(Slot).Arena).EnUso = 1
    ElseIf TipoDuelo = eDuelType.vs4 Then
        Arena4v4(DuelData.Duelo(Slot).Arena).EnUso = 1
    End If
    UserList(Retador).flags.DueloIndex = Slot
    UserList(Retador).flags.DueloTeam = 1
    Call AssignTeamMember(Slot, UserList(Retador).flags.DueloTeam, Retador)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub GiveDueloSlot de modDuelos.bas")
End Sub

Sub SetDueloConfig(ByVal Slot As Byte, ByVal TipoDuelo As eDuelType, ByVal Drop As Boolean, ByVal Resucitar As Boolean, ByVal Oro As Long)
On Error GoTo ErrHandler
  
    If Not Slot > 0 Then Exit Sub
    
    DuelData.Duelo(Slot).Drop = Drop
    DuelData.Duelo(Slot).Oro = Oro
    DuelData.Duelo(Slot).Resucitar = Resucitar
    DuelData.Duelo(Slot).TipoDuelo = TipoDuelo
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SetDueloConfig de modDuelos.bas")
End Sub

Function GetTipoDuelo(ByVal Slot As Byte) As eDuelType
On Error GoTo ErrHandler

    If Not Slot > 0 Then Exit Function
    
    GetTipoDuelo = DuelData.Duelo(Slot).TipoDuelo
    
    Exit Function
ErrHandler:
    Call LogError("Error " & Err.Number & " (" & Err.Description & ") en Function GetTipoDuelo del Módulo modDuelos")
End Function

Sub AssignTeamMember(ByVal Slot As Byte, ByVal Team As Byte, ByVal Player As Integer)
On Error GoTo ErrHandler
  
    If Not Slot > 0 Then Exit Sub
    
    If GetTipoDuelo(Slot) = eDuelType.vs1 Then
        If Not DuelData.Duelo(Slot).Team(Team).Player(1) > 0 Then
            DuelData.Duelo(Slot).Team(Team).Player(1) = Player
            DuelData.Duelo(Slot).Team(Team).Muerto(1) = False
            If Not Team = 1 Then Call SendData(SendTarget.ToDuelo, Slot, PrepareMessageConsoleMsg(UserList(Player).Name & " ha aceptado el Duelo.", FontTypeNames.FONTTYPE_INFO))
        End If
    ElseIf GetTipoDuelo(Slot) = eDuelType.vs2 Then
        If DuelData.Duelo(Slot).Team(Team).Player(1) > 0 Then
            If Not DuelData.Duelo(Slot).Team(Team).Player(2) > 0 Then
                DuelData.Duelo(Slot).Team(Team).Player(2) = Player
                DuelData.Duelo(Slot).Team(Team).Muerto(2) = False
                Call SendData(SendTarget.ToDuelo, Slot, PrepareMessageConsoleMsg(UserList(Player).Name & " ha aceptado el Duelo.", FontTypeNames.FONTTYPE_INFO))
            End If
        Else
            DuelData.Duelo(Slot).Team(Team).Player(1) = Player
            DuelData.Duelo(Slot).Team(Team).Muerto(1) = False
            If Not Team = 1 Then Call SendData(SendTarget.ToDuelo, Slot, PrepareMessageConsoleMsg(UserList(Player).Name & " ha aceptado el Duelo.", FontTypeNames.FONTTYPE_INFO))
        End If
    ElseIf GetTipoDuelo(Slot) = eDuelType.vs3 Then
        If DuelData.Duelo(Slot).Team(Team).Player(1) > 0 Then
            If DuelData.Duelo(Slot).Team(Team).Player(2) > 0 Then
                If Not DuelData.Duelo(Slot).Team(Team).Player(3) > 0 Then
                    DuelData.Duelo(Slot).Team(Team).Player(3) = Player
                    DuelData.Duelo(Slot).Team(Team).Muerto(3) = False
                    Call SendData(SendTarget.ToDuelo, Slot, PrepareMessageConsoleMsg(UserList(Player).Name & " ha aceptado el Duelo.", FontTypeNames.FONTTYPE_INFO))
                End If
            Else
                DuelData.Duelo(Slot).Team(Team).Player(2) = Player
                DuelData.Duelo(Slot).Team(Team).Muerto(2) = False
                Call SendData(SendTarget.ToDuelo, Slot, PrepareMessageConsoleMsg(UserList(Player).Name & " ha aceptado el Duelo.", FontTypeNames.FONTTYPE_INFO))
            End If
        Else
            DuelData.Duelo(Slot).Team(Team).Player(1) = Player
            DuelData.Duelo(Slot).Team(Team).Muerto(1) = False
            If Not Team = 1 Then Call SendData(SendTarget.ToDuelo, Slot, PrepareMessageConsoleMsg(UserList(Player).Name & " ha aceptado el Duelo.", FontTypeNames.FONTTYPE_INFO))
        End If
    ElseIf GetTipoDuelo(Slot) = eDuelType.vs4 Then
        If DuelData.Duelo(Slot).Team(Team).Player(1) > 0 Then
            If DuelData.Duelo(Slot).Team(Team).Player(2) > 0 Then
                If DuelData.Duelo(Slot).Team(Team).Player(3) > 0 Then
                    If Not DuelData.Duelo(Slot).Team(Team).Player(4) > 0 Then
                        DuelData.Duelo(Slot).Team(Team).Player(4) = Player
                        DuelData.Duelo(Slot).Team(Team).Muerto(4) = False
                        Call SendData(SendTarget.ToDuelo, Slot, PrepareMessageConsoleMsg(UserList(Player).Name & " ha aceptado el Duelo.", FontTypeNames.FONTTYPE_INFO))
                    End If
                Else
                    DuelData.Duelo(Slot).Team(Team).Player(3) = Player
                    DuelData.Duelo(Slot).Team(Team).Muerto(3) = False
                    Call SendData(SendTarget.ToDuelo, Slot, PrepareMessageConsoleMsg(UserList(Player).Name & " ha aceptado el Duelo.", FontTypeNames.FONTTYPE_INFO))
                End If
            Else
                DuelData.Duelo(Slot).Team(Team).Player(2) = Player
                DuelData.Duelo(Slot).Team(Team).Muerto(2) = False
                Call SendData(SendTarget.ToDuelo, Slot, PrepareMessageConsoleMsg(UserList(Player).Name & " ha aceptado el Duelo.", FontTypeNames.FONTTYPE_INFO))
            End If
        Else
            DuelData.Duelo(Slot).Team(Team).Player(1) = Player
            DuelData.Duelo(Slot).Team(Team).Muerto(1) = False
            If Not Team = 1 Then Call SendData(SendTarget.ToDuelo, Slot, PrepareMessageConsoleMsg(UserList(Player).Name & " ha aceptado el Duelo.", FontTypeNames.FONTTYPE_INFO))
        End If
    End If
    Call CheckDueloPlayers(Slot)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub AssignTeamMember de modDuelos.bas")
End Sub

Sub CheckDueloPlayers(ByVal Slot As Byte)
On Error GoTo ErrHandler
  
    If Not Slot > 0 Then Exit Sub
    
    If GetTipoDuelo(Slot) = eDuelType.vs1 Then
        If DuelData.Duelo(Slot).Team(1).Player(1) > 0 Then
            If DuelData.Duelo(Slot).Team(2).Player(1) > 0 Then
                Call SendData(SendTarget.ToDuelo, Slot, PrepareMessageConsoleMsg("Todos los participantes han aceptado el duelo.", FontTypeNames.FONTTYPE_INFO))
                Call InvocarDueloPlayers(Slot)
            End If
        End If
    ElseIf GetTipoDuelo(Slot) = eDuelType.vs2 Then
        If DuelData.Duelo(Slot).Team(1).Player(1) > 0 Then
            If DuelData.Duelo(Slot).Team(1).Player(2) > 0 Then
                If DuelData.Duelo(Slot).Team(2).Player(1) > 0 Then
                    If DuelData.Duelo(Slot).Team(2).Player(2) > 0 Then
                        Call SendData(SendTarget.ToDuelo, Slot, PrepareMessageConsoleMsg("Todos los participantes han aceptado el duelo.", FontTypeNames.FONTTYPE_INFO))
                        Call InvocarDueloPlayers(Slot)
                    End If
                End If
            End If
        End If
    ElseIf GetTipoDuelo(Slot) = eDuelType.vs3 Then
        If DuelData.Duelo(Slot).Team(1).Player(1) > 0 Then
            If DuelData.Duelo(Slot).Team(1).Player(2) > 0 Then
                If DuelData.Duelo(Slot).Team(1).Player(3) > 0 Then
                    If DuelData.Duelo(Slot).Team(2).Player(1) > 0 Then
                        If DuelData.Duelo(Slot).Team(2).Player(2) > 0 Then
                            If DuelData.Duelo(Slot).Team(2).Player(3) > 0 Then
                                Call SendData(SendTarget.ToDuelo, Slot, PrepareMessageConsoleMsg("Todos los participantes han aceptado el duelo.", FontTypeNames.FONTTYPE_INFO))
                                Call InvocarDueloPlayers(Slot)
                            End If
                        End If
                    End If
                End If
            End If
        End If
    ElseIf GetTipoDuelo(Slot) = eDuelType.vs4 Then
        If DuelData.Duelo(Slot).Team(1).Player(1) > 0 Then
            If DuelData.Duelo(Slot).Team(1).Player(2) > 0 Then
                If DuelData.Duelo(Slot).Team(1).Player(3) > 0 Then
                    If DuelData.Duelo(Slot).Team(1).Player(4) > 0 Then
                        If DuelData.Duelo(Slot).Team(2).Player(1) > 0 Then
                            If DuelData.Duelo(Slot).Team(2).Player(2) > 0 Then
                                If DuelData.Duelo(Slot).Team(2).Player(3) > 0 Then
                                    If DuelData.Duelo(Slot).Team(2).Player(4) > 0 Then
                                        Call SendData(SendTarget.ToDuelo, Slot, PrepareMessageConsoleMsg("Todos los participantes han aceptado el duelo.", FontTypeNames.FONTTYPE_INFO))
                                        Call InvocarDueloPlayers(Slot)
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CheckDueloPlayers de modDuelos.bas")
End Sub

Function GetUserTeam(ByVal Slot As Byte, ByVal UserIndex As Integer) As Byte
On Error GoTo ErrHandler

    If Not Slot > 0 Then Exit Function
    
    Dim I As Byte
    If GetTipoDuelo(Slot) = eDuelType.vs1 Then
        For I = 1 To 2
            If DuelData.Duelo(Slot).Team(I).Player(1) = UserIndex Then
                GetUserTeam = I
                Exit Function
            End If
        Next I
    ElseIf GetTipoDuelo(Slot) = eDuelType.vs2 Then
        For I = 1 To 2
            If DuelData.Duelo(Slot).Team(I).Player(1) = UserIndex Then
                GetUserTeam = I
                Exit Function
            ElseIf DuelData.Duelo(Slot).Team(I).Player(2) = UserIndex Then
                GetUserTeam = I
                Exit Function
            End If
        Next I
    ElseIf GetTipoDuelo(Slot) = eDuelType.vs3 Then
        For I = 1 To 2
            If DuelData.Duelo(Slot).Team(I).Player(1) = UserIndex Then
                GetUserTeam = I
                Exit Function
            ElseIf DuelData.Duelo(Slot).Team(I).Player(2) = UserIndex Then
                GetUserTeam = I
                Exit Function
            ElseIf DuelData.Duelo(Slot).Team(I).Player(3) = UserIndex Then
                GetUserTeam = I
                Exit Function
            End If
        Next I
    ElseIf GetTipoDuelo(Slot) = eDuelType.vs4 Then
        For I = 1 To 2
            If DuelData.Duelo(Slot).Team(I).Player(1) = UserIndex Then
                GetUserTeam = I
                Exit Function
            ElseIf DuelData.Duelo(Slot).Team(I).Player(2) = UserIndex Then
                GetUserTeam = I
                Exit Function
            ElseIf DuelData.Duelo(Slot).Team(I).Player(3) = UserIndex Then
                GetUserTeam = I
                Exit Function
            ElseIf DuelData.Duelo(Slot).Team(I).Player(4) = UserIndex Then
                GetUserTeam = I
                Exit Function
            End If
        Next I
    End If
    
    GetUserTeam = 0

    Exit Function
ErrHandler:
    Call LogError("Error " & Err.Number & " (" & Err.Description & ") en Function GetUserTeam del Módulo modDuelos")
End Function

Sub TerminoDuelo(ByVal Slot As Byte, ByVal LoserTeam As Byte)
On Error GoTo ErrHandler
  
    If Not Slot > 0 Then Exit Sub
    
    If DuelData.Duelo(Slot).estado = eDuelState.Esperando_Inicio Or DuelData.Duelo(Slot).estado = Esperando_Final Then
        If GetTipoDuelo(Slot) = eDuelType.vs1 Then
            If DuelData.Duelo(Slot).Team(1).Player(1) > 0 Then Call WriteParalizeOK(DuelData.Duelo(Slot).Team(1).Player(1), False)
            If DuelData.Duelo(Slot).Team(2).Player(1) > 0 Then Call WriteParalizeOK(DuelData.Duelo(Slot).Team(2).Player(1), False)
        ElseIf GetTipoDuelo(Slot) = eDuelType.vs2 Then
            If DuelData.Duelo(Slot).Team(1).Player(1) > 0 Then Call WriteParalizeOK(DuelData.Duelo(Slot).Team(1).Player(1), False)
            If DuelData.Duelo(Slot).Team(2).Player(1) > 0 Then Call WriteParalizeOK(DuelData.Duelo(Slot).Team(2).Player(1), False)
            If DuelData.Duelo(Slot).Team(1).Player(2) > 0 Then Call WriteParalizeOK(DuelData.Duelo(Slot).Team(1).Player(2), False)
            If DuelData.Duelo(Slot).Team(2).Player(2) > 0 Then Call WriteParalizeOK(DuelData.Duelo(Slot).Team(2).Player(2), False)
        ElseIf GetTipoDuelo(Slot) = eDuelType.vs3 Then
            If DuelData.Duelo(Slot).Team(1).Player(1) > 0 Then Call WriteParalizeOK(DuelData.Duelo(Slot).Team(1).Player(1), False)
            If DuelData.Duelo(Slot).Team(2).Player(1) > 0 Then Call WriteParalizeOK(DuelData.Duelo(Slot).Team(2).Player(1), False)
            If DuelData.Duelo(Slot).Team(1).Player(2) > 0 Then Call WriteParalizeOK(DuelData.Duelo(Slot).Team(1).Player(2), False)
            If DuelData.Duelo(Slot).Team(2).Player(2) > 0 Then Call WriteParalizeOK(DuelData.Duelo(Slot).Team(2).Player(2), False)
            If DuelData.Duelo(Slot).Team(1).Player(3) > 0 Then Call WriteParalizeOK(DuelData.Duelo(Slot).Team(1).Player(3), False)
            If DuelData.Duelo(Slot).Team(2).Player(3) > 0 Then Call WriteParalizeOK(DuelData.Duelo(Slot).Team(2).Player(3), False)
        ElseIf GetTipoDuelo(Slot) = eDuelType.vs4 Then
            If DuelData.Duelo(Slot).Team(1).Player(1) > 0 Then Call WriteParalizeOK(DuelData.Duelo(Slot).Team(1).Player(1), False)
            If DuelData.Duelo(Slot).Team(2).Player(1) > 0 Then Call WriteParalizeOK(DuelData.Duelo(Slot).Team(2).Player(1), False)
            If DuelData.Duelo(Slot).Team(1).Player(2) > 0 Then Call WriteParalizeOK(DuelData.Duelo(Slot).Team(1).Player(2), False)
            If DuelData.Duelo(Slot).Team(2).Player(2) > 0 Then Call WriteParalizeOK(DuelData.Duelo(Slot).Team(2).Player(2), False)
            If DuelData.Duelo(Slot).Team(1).Player(3) > 0 Then Call WriteParalizeOK(DuelData.Duelo(Slot).Team(1).Player(3), False)
            If DuelData.Duelo(Slot).Team(2).Player(3) > 0 Then Call WriteParalizeOK(DuelData.Duelo(Slot).Team(2).Player(3), False)
            If DuelData.Duelo(Slot).Team(1).Player(4) > 0 Then Call WriteParalizeOK(DuelData.Duelo(Slot).Team(1).Player(4), False)
            If DuelData.Duelo(Slot).Team(2).Player(4) > 0 Then Call WriteParalizeOK(DuelData.Duelo(Slot).Team(2).Player(4), False)
        End If
    End If
    
    If DuelData.Duelo(Slot).Drop And LoserTeam <> 0 Then
        DuelData.Duelo(Slot).Counter = 30
    Else
        DuelData.Duelo(Slot).Counter = 3
    End If
    
    ' Empate.
    If LoserTeam = 0 Then
        
    ElseIf LoserTeam = 1 Then
        Call PerdioDuelo(Slot, 1)
        Call GanoDuelo(Slot, 2)
    Else
        Call PerdioDuelo(Slot, 2)
        Call GanoDuelo(Slot, 1)
    End If
    
    DuelData.Duelo(Slot).estado = eDuelState.Esperando_Final
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub TerminoDuelo de modDuelos.bas")
End Sub

Public Sub TerminarDueloTimeout(ByVal Slot As Byte)
On Error GoTo ErrHandler

    If Not Slot > 0 Then Exit Sub
    
    DuelData.Duelo(Slot).estado = eDuelState.Esperando_Final
    DuelData.Duelo(Slot).Counter = 2
        
    Call SendData(SendTarget.ToDuelo, Slot, PrepareMessageConsoleMsg("Se ha acabado el tiempo máximo para completar el duelo.", FontTypeNames.FONTTYPE_TALK))
    'Call TerminoDuelo(Slot, 0)
    
    Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub TerminarDueloTimeout de modDuelos.bas")
End Sub

Sub AbandonarDuelo(ByVal Slot As Byte, ByVal UserIndex As Integer)
On Error GoTo ErrHandler

    If Not Slot > 0 Then Exit Sub
    
    Dim TeamSlot As Byte
    Dim Team As Byte
    Team = GetUserTeam(Slot, UserIndex)
    If Not GetTeamPlayers(Slot, Team) > 0 Then Exit Sub
    
    If GetTipoDuelo(Slot) = eDuelType.vs1 Then
        Call SendData(SendTarget.ToDuelo, Slot, PrepareMessageConsoleMsg(UserList(UserIndex).Name & " ha abandonado el duelo.", FontTypeNames.FONTTYPE_INFO))
        If Team = 1 Then
            Call TerminoDuelo(Slot, 1)
        Else
            Call TerminoDuelo(Slot, 2)
        End If
    ElseIf GetTipoDuelo(Slot) = eDuelType.vs2 Then
        If GetTeamPlayers(Slot, Team) = 2 Then
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - DuelData.Duelo(Slot).Oro
            UserList(UserIndex).flags.DueloIndex = 0
            UserList(UserIndex).flags.DueloTeam = 0
            If DuelData.Duelo(Slot).Drop Then
                Call TirarTodosLosItems(UserIndex)
            End If
            Call WriteUpdateGold(UserIndex)
            Call WarpReturnDuelo(UserIndex)
            Call SendData(SendTarget.ToDuelo, Slot, PrepareMessageConsoleMsg(UserList(UserIndex).Name & " ha abandonado el duelo.", FontTypeNames.FONTTYPE_INFO))
            TeamSlot = GetTeamSlot(Slot, Team, UserIndex)
            If TeamSlot = 1 Then
                DuelData.Duelo(Slot).Team(Team).Player(1) = 0
                DuelData.Duelo(Slot).Team(Team).Muerto(1) = True
            ElseIf TeamSlot = 2 Then
                DuelData.Duelo(Slot).Team(Team).Player(2) = 0
                DuelData.Duelo(Slot).Team(Team).Muerto(2) = True
            End If
        Else
            Call SendData(SendTarget.ToDuelo, Slot, PrepareMessageConsoleMsg(UserList(UserIndex).Name & " ha abandonado el duelo.", FontTypeNames.FONTTYPE_INFO))
            If Team = 1 Then
                Call TerminoDuelo(Slot, 1)
            Else
                Call TerminoDuelo(Slot, 2)
            End If
        End If
    ElseIf GetTipoDuelo(Slot) = eDuelType.vs3 Then
        If GetTeamPlayers(Slot, Team) >= 2 Then
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - DuelData.Duelo(Slot).Oro
            UserList(UserIndex).flags.DueloIndex = 0
            UserList(UserIndex).flags.DueloTeam = 0
            If DuelData.Duelo(Slot).Drop Then
                Call TirarTodosLosItems(UserIndex)
            End If
            Call WriteUpdateGold(UserIndex)
            Call WarpReturnDuelo(UserIndex)
            Call SendData(SendTarget.ToDuelo, Slot, PrepareMessageConsoleMsg(UserList(UserIndex).Name & " ha abandonado el duelo.", FontTypeNames.FONTTYPE_INFO))
            TeamSlot = GetTeamSlot(Slot, Team, UserIndex)
            If TeamSlot = 1 Then
                DuelData.Duelo(Slot).Team(Team).Player(1) = 0
                DuelData.Duelo(Slot).Team(Team).Muerto(1) = True
            ElseIf TeamSlot = 2 Then
                DuelData.Duelo(Slot).Team(Team).Player(2) = 0
                DuelData.Duelo(Slot).Team(Team).Muerto(2) = True
            ElseIf TeamSlot = 3 Then
                DuelData.Duelo(Slot).Team(Team).Player(3) = 0
                DuelData.Duelo(Slot).Team(Team).Muerto(3) = 0
            End If
        Else
            Call SendData(SendTarget.ToDuelo, Slot, PrepareMessageConsoleMsg(UserList(UserIndex).Name & " ha abandonado el duelo.", FontTypeNames.FONTTYPE_INFO))
            If Team = 1 Then
                Call TerminoDuelo(Slot, 1)
            Else
                Call TerminoDuelo(Slot, 2)
            End If
        End If
    ElseIf GetTipoDuelo(Slot) = eDuelType.vs4 Then
        If GetTeamPlayers(Slot, Team) >= 2 Then
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - DuelData.Duelo(Slot).Oro
            UserList(UserIndex).flags.DueloIndex = 0
            UserList(UserIndex).flags.DueloTeam = 0
            If DuelData.Duelo(Slot).Drop Then
                Call TirarTodosLosItems(UserIndex)
            End If
            Call WriteUpdateGold(UserIndex)
            Call WarpReturnDuelo(UserIndex)
            Call SendData(SendTarget.ToDuelo, Slot, PrepareMessageConsoleMsg(UserList(UserIndex).Name & " ha abandonado el duelo.", FontTypeNames.FONTTYPE_INFO))
            TeamSlot = GetTeamSlot(Slot, Team, UserIndex)
            If TeamSlot = 1 Then
                DuelData.Duelo(Slot).Team(Team).Player(1) = 0
                DuelData.Duelo(Slot).Team(Team).Muerto(1) = True
            ElseIf TeamSlot = 2 Then
                DuelData.Duelo(Slot).Team(Team).Player(2) = 0
                DuelData.Duelo(Slot).Team(Team).Muerto(2) = True
            ElseIf TeamSlot = 3 Then
                DuelData.Duelo(Slot).Team(Team).Player(3) = 0
                DuelData.Duelo(Slot).Team(Team).Muerto(3) = True
            ElseIf TeamSlot = 4 Then
                DuelData.Duelo(Slot).Team(Team).Player(4) = 0
                DuelData.Duelo(Slot).Team(Team).Muerto(4) = True
            End If
        Else
            Call SendData(SendTarget.ToDuelo, Slot, PrepareMessageConsoleMsg(UserList(UserIndex).Name & " ha abandonado el duelo.", FontTypeNames.FONTTYPE_INFO))
            If Team = 1 Then
                Call TerminoDuelo(Slot, 1)
            Else
                Call TerminoDuelo(Slot, 2)
            End If
        End If
    End If
    Exit Sub

ErrHandler:
        Call LogError("Error en Sub AbandonarDuelo. Slot: " & Slot & " " & Err.Description)
End Sub

Function GetTeamPlayers(ByVal Slot As Byte, ByVal Team As Byte) As Byte
On Error GoTo ErrHandler
  
    If Not Slot > 0 Then Exit Function
    GetTeamPlayers = 0
    
    If GetTipoDuelo(Slot) = eDuelType.vs1 Then
        If DuelData.Duelo(Slot).Team(Team).Player(1) > 0 Then
            GetTeamPlayers = GetTeamPlayers + 1
        End If
    ElseIf GetTipoDuelo(Slot) = eDuelType.vs2 Then
        If DuelData.Duelo(Slot).Team(Team).Player(1) > 0 Then
            GetTeamPlayers = GetTeamPlayers + 1
        End If
        If DuelData.Duelo(Slot).Team(Team).Player(2) > 0 Then
            GetTeamPlayers = GetTeamPlayers + 1
        End If
    ElseIf GetTipoDuelo(Slot) = eDuelType.vs3 Then
        If DuelData.Duelo(Slot).Team(Team).Player(1) > 0 Then
            GetTeamPlayers = GetTeamPlayers + 1
        End If
        If DuelData.Duelo(Slot).Team(Team).Player(2) > 0 Then
            GetTeamPlayers = GetTeamPlayers + 1
        End If
        If DuelData.Duelo(Slot).Team(Team).Player(3) > 0 Then
            GetTeamPlayers = GetTeamPlayers + 1
        End If
    ElseIf GetTipoDuelo(Slot) = eDuelType.vs4 Then
        If DuelData.Duelo(Slot).Team(Team).Player(1) > 0 Then
            GetTeamPlayers = GetTeamPlayers + 1
        End If
        If DuelData.Duelo(Slot).Team(Team).Player(2) > 0 Then
            GetTeamPlayers = GetTeamPlayers + 1
        End If
        If DuelData.Duelo(Slot).Team(Team).Player(3) > 0 Then
            GetTeamPlayers = GetTeamPlayers + 1
        End If
        If DuelData.Duelo(Slot).Team(Team).Player(4) > 0 Then
            GetTeamPlayers = GetTeamPlayers + 1
        End If
    End If
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetTeamPlayers de modDuelos.bas")
End Function

Function GetTeamSlot(ByVal Slot As Byte, ByVal Team As Byte, ByVal UserIndex As Integer) As Byte
On Error GoTo ErrHandler

    If Not Slot > 0 Then Exit Function
    'If Not Team > 0 Then Exit Function
    
    If GetTipoDuelo(Slot) = eDuelType.vs1 Then
        If DuelData.Duelo(Slot).Team(Team).Player(1) = UserIndex Then
            GetTeamSlot = 1
            Exit Function
        End If
    ElseIf GetTipoDuelo(Slot) = eDuelType.vs2 Then
        If DuelData.Duelo(Slot).Team(Team).Player(1) = UserIndex Then
            GetTeamSlot = 1
            Exit Function
        ElseIf DuelData.Duelo(Slot).Team(Team).Player(2) = UserIndex Then
            GetTeamSlot = 2
            Exit Function
        End If
    ElseIf GetTipoDuelo(Slot) = eDuelType.vs3 Then
        If DuelData.Duelo(Slot).Team(Team).Player(1) = UserIndex Then
            GetTeamSlot = 1
            Exit Function
        ElseIf DuelData.Duelo(Slot).Team(Team).Player(2) = UserIndex Then
            GetTeamSlot = 2
            Exit Function
        ElseIf DuelData.Duelo(Slot).Team(Team).Player(3) = UserIndex Then
            GetTeamSlot = 3
            Exit Function
        End If
    ElseIf GetTipoDuelo(Slot) = eDuelType.vs4 Then
        If DuelData.Duelo(Slot).Team(Team).Player(1) = UserIndex Then
            GetTeamSlot = 1
            Exit Function
        ElseIf DuelData.Duelo(Slot).Team(Team).Player(2) = UserIndex Then
            GetTeamSlot = 2
            Exit Function
        ElseIf DuelData.Duelo(Slot).Team(Team).Player(3) = UserIndex Then
            GetTeamSlot = 3
            Exit Function
        ElseIf DuelData.Duelo(Slot).Team(Team).Player(4) = UserIndex Then
            GetTeamSlot = 4
            Exit Function
        End If
    End If
    
    GetTeamSlot = 0
    
    Exit Function
ErrHandler:
    Call LogError("Error " & Err.Number & " (" & Err.Description & ") en Function GetTeamSlot del Módulo modDuelos. Slot de Duelo: " & Slot & " - Slot de Equipo: " & Team & " - Estado del Duelo: " & DuelData.Duelo(Slot).estado)
End Function

Sub PerdioDuelo(ByVal Slot As Byte, ByVal Team As Byte)
On Error GoTo ErrHandler

    If Not Slot > 0 Then Exit Sub
    
    If GetTipoDuelo(Slot) = eDuelType.vs1 Then
        If DuelData.Duelo(Slot).Team(Team).Player(1) > 0 Then
            UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.GLD = UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.GLD - DuelData.Duelo(Slot).Oro
            UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.OroDuelos = UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.OroDuelos - DuelData.Duelo(Slot).Oro
            UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.DuelosPerdidos = UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.DuelosPerdidos + 1
            If UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.GLD < 0 Then UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.GLD = 0
            Call WriteUpdateGold(DuelData.Duelo(Slot).Team(Team).Player(1))
            If DuelData.Duelo(Slot).Drop Then
                Call TirarTodosLosItems(DuelData.Duelo(Slot).Team(Team).Player(1))
            End If
            Call WarpReturnDuelo(DuelData.Duelo(Slot).Team(Team).Player(1))
            UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).flags.DueloIndex = 0
            UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).flags.DueloTeam = 0
            Call WriteConsoleMsg(DuelData.Duelo(Slot).Team(Team).Player(1), "Has perdido el duelo.", FontTypeNames.FONTTYPE_TALK)
            DuelData.Duelo(Slot).Team(Team).Player(1) = 0
        End If
    ElseIf GetTipoDuelo(Slot) = eDuelType.vs2 Then
        If DuelData.Duelo(Slot).Team(Team).Player(1) > 0 Then
            UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.GLD = UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.GLD - DuelData.Duelo(Slot).Oro
            UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.OroDuelos = UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.OroDuelos - DuelData.Duelo(Slot).Oro
            UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.DuelosPerdidos = UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.DuelosPerdidos + 1
            If UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.GLD < 0 Then UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.GLD = 0
            Call WriteUpdateGold(DuelData.Duelo(Slot).Team(Team).Player(1))
            If DuelData.Duelo(Slot).Drop Then
                Call TirarTodosLosItems(DuelData.Duelo(Slot).Team(Team).Player(1))
            End If
            Call WarpReturnDuelo(DuelData.Duelo(Slot).Team(Team).Player(1))
            UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).flags.DueloIndex = 0
            UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).flags.DueloTeam = 0
            Call WriteConsoleMsg(DuelData.Duelo(Slot).Team(Team).Player(1), "Has perdido el duelo.", FontTypeNames.FONTTYPE_TALK)
            DuelData.Duelo(Slot).Team(Team).Player(1) = 0
        End If
        If DuelData.Duelo(Slot).Team(Team).Player(2) > 0 Then
            UserList(DuelData.Duelo(Slot).Team(Team).Player(2)).Stats.GLD = UserList(DuelData.Duelo(Slot).Team(Team).Player(2)).Stats.GLD - DuelData.Duelo(Slot).Oro
            UserList(DuelData.Duelo(Slot).Team(Team).Player(2)).Stats.OroDuelos = UserList(DuelData.Duelo(Slot).Team(Team).Player(2)).Stats.OroDuelos - DuelData.Duelo(Slot).Oro
            UserList(DuelData.Duelo(Slot).Team(Team).Player(2)).Stats.DuelosPerdidos = UserList(DuelData.Duelo(Slot).Team(Team).Player(2)).Stats.DuelosPerdidos + 1
            If UserList(DuelData.Duelo(Slot).Team(Team).Player(2)).Stats.GLD < 0 Then UserList(DuelData.Duelo(Slot).Team(Team).Player(2)).Stats.GLD = 0
            Call WriteUpdateGold(DuelData.Duelo(Slot).Team(Team).Player(2))
            If DuelData.Duelo(Slot).Drop Then
                Call TirarTodosLosItems(DuelData.Duelo(Slot).Team(Team).Player(2))
            End If
            Call WarpReturnDuelo(DuelData.Duelo(Slot).Team(Team).Player(2))
            UserList(DuelData.Duelo(Slot).Team(Team).Player(2)).flags.DueloIndex = 0
            UserList(DuelData.Duelo(Slot).Team(Team).Player(2)).flags.DueloTeam = 0
            Call WriteConsoleMsg(DuelData.Duelo(Slot).Team(Team).Player(2), "Has perdido el duelo.", FontTypeNames.FONTTYPE_TALK)
            DuelData.Duelo(Slot).Team(Team).Player(2) = 0
        End If
    ElseIf GetTipoDuelo(Slot) = eDuelType.vs3 Then
        If DuelData.Duelo(Slot).Team(Team).Player(1) > 0 Then
            UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.GLD = UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.GLD - DuelData.Duelo(Slot).Oro
            UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.OroDuelos = UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.OroDuelos - DuelData.Duelo(Slot).Oro
            UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.DuelosPerdidos = UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.DuelosPerdidos + 1
            If UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.GLD < 0 Then UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.GLD = 0
            Call WriteUpdateGold(DuelData.Duelo(Slot).Team(Team).Player(1))
            If DuelData.Duelo(Slot).Drop Then
                Call TirarTodosLosItems(DuelData.Duelo(Slot).Team(Team).Player(1))
            End If
            Call WarpReturnDuelo(DuelData.Duelo(Slot).Team(Team).Player(1))
            UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).flags.DueloIndex = 0
            UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).flags.DueloTeam = 0
            Call WriteConsoleMsg(DuelData.Duelo(Slot).Team(Team).Player(1), "Has perdido el duelo.", FontTypeNames.FONTTYPE_TALK)
            DuelData.Duelo(Slot).Team(Team).Player(1) = 0
        End If
        If DuelData.Duelo(Slot).Team(Team).Player(2) > 0 Then
            UserList(DuelData.Duelo(Slot).Team(Team).Player(2)).Stats.GLD = UserList(DuelData.Duelo(Slot).Team(Team).Player(2)).Stats.GLD - DuelData.Duelo(Slot).Oro
            UserList(DuelData.Duelo(Slot).Team(Team).Player(2)).Stats.OroDuelos = UserList(DuelData.Duelo(Slot).Team(Team).Player(2)).Stats.OroDuelos - DuelData.Duelo(Slot).Oro
            UserList(DuelData.Duelo(Slot).Team(Team).Player(2)).Stats.DuelosPerdidos = UserList(DuelData.Duelo(Slot).Team(Team).Player(2)).Stats.DuelosPerdidos + 1
            If UserList(DuelData.Duelo(Slot).Team(Team).Player(2)).Stats.GLD < 0 Then UserList(DuelData.Duelo(Slot).Team(Team).Player(2)).Stats.GLD = 0
            Call WriteUpdateGold(DuelData.Duelo(Slot).Team(Team).Player(2))
            If DuelData.Duelo(Slot).Drop Then
                Call TirarTodosLosItems(DuelData.Duelo(Slot).Team(Team).Player(2))
            End If
            Call WarpReturnDuelo(DuelData.Duelo(Slot).Team(Team).Player(2))
            UserList(DuelData.Duelo(Slot).Team(Team).Player(2)).flags.DueloIndex = 0
            UserList(DuelData.Duelo(Slot).Team(Team).Player(2)).flags.DueloTeam = 0
            Call WriteConsoleMsg(DuelData.Duelo(Slot).Team(Team).Player(2), "Has perdido el duelo.", FontTypeNames.FONTTYPE_TALK)
            DuelData.Duelo(Slot).Team(Team).Player(2) = 0
        End If
        If DuelData.Duelo(Slot).Team(Team).Player(3) > 0 Then
            UserList(DuelData.Duelo(Slot).Team(Team).Player(3)).Stats.GLD = UserList(DuelData.Duelo(Slot).Team(Team).Player(3)).Stats.GLD - DuelData.Duelo(Slot).Oro
            UserList(DuelData.Duelo(Slot).Team(Team).Player(3)).Stats.OroDuelos = UserList(DuelData.Duelo(Slot).Team(Team).Player(3)).Stats.OroDuelos - DuelData.Duelo(Slot).Oro
            UserList(DuelData.Duelo(Slot).Team(Team).Player(3)).Stats.DuelosPerdidos = UserList(DuelData.Duelo(Slot).Team(Team).Player(3)).Stats.DuelosPerdidos + 1
            If UserList(DuelData.Duelo(Slot).Team(Team).Player(3)).Stats.GLD < 0 Then UserList(DuelData.Duelo(Slot).Team(Team).Player(3)).Stats.GLD = 0
            Call WriteUpdateGold(DuelData.Duelo(Slot).Team(Team).Player(3))
            If DuelData.Duelo(Slot).Drop Then
                Call TirarTodosLosItems(DuelData.Duelo(Slot).Team(Team).Player(3))
            End If
            Call WarpReturnDuelo(DuelData.Duelo(Slot).Team(Team).Player(3))
            UserList(DuelData.Duelo(Slot).Team(Team).Player(3)).flags.DueloIndex = 0
            UserList(DuelData.Duelo(Slot).Team(Team).Player(3)).flags.DueloTeam = 0
            Call WriteConsoleMsg(DuelData.Duelo(Slot).Team(Team).Player(3), "Has perdido el duelo.", FontTypeNames.FONTTYPE_TALK)
            DuelData.Duelo(Slot).Team(Team).Player(3) = 0
        End If
    ElseIf GetTipoDuelo(Slot) = eDuelType.vs4 Then
        If DuelData.Duelo(Slot).Team(Team).Player(1) > 0 Then
            UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.GLD = UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.GLD - DuelData.Duelo(Slot).Oro
            UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.OroDuelos = UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.OroDuelos - DuelData.Duelo(Slot).Oro
            UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.DuelosPerdidos = UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.DuelosPerdidos + 1
            If UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.GLD < 0 Then UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.GLD = 0
            Call WriteUpdateGold(DuelData.Duelo(Slot).Team(Team).Player(1))
            If DuelData.Duelo(Slot).Drop Then
                Call TirarTodosLosItems(DuelData.Duelo(Slot).Team(Team).Player(1))
            End If
            Call WarpReturnDuelo(DuelData.Duelo(Slot).Team(Team).Player(1))
            UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).flags.DueloIndex = 0
            UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).flags.DueloTeam = 0
            Call WriteConsoleMsg(DuelData.Duelo(Slot).Team(Team).Player(1), "Has perdido el duelo.", FontTypeNames.FONTTYPE_TALK)
            DuelData.Duelo(Slot).Team(Team).Player(1) = 0
        End If
        If DuelData.Duelo(Slot).Team(Team).Player(2) > 0 Then
            UserList(DuelData.Duelo(Slot).Team(Team).Player(2)).Stats.GLD = UserList(DuelData.Duelo(Slot).Team(Team).Player(2)).Stats.GLD - DuelData.Duelo(Slot).Oro
            UserList(DuelData.Duelo(Slot).Team(Team).Player(2)).Stats.OroDuelos = UserList(DuelData.Duelo(Slot).Team(Team).Player(2)).Stats.OroDuelos - DuelData.Duelo(Slot).Oro
            UserList(DuelData.Duelo(Slot).Team(Team).Player(2)).Stats.DuelosPerdidos = UserList(DuelData.Duelo(Slot).Team(Team).Player(2)).Stats.DuelosPerdidos + 1
            If UserList(DuelData.Duelo(Slot).Team(Team).Player(2)).Stats.GLD < 0 Then UserList(DuelData.Duelo(Slot).Team(Team).Player(2)).Stats.GLD = 0
            Call WriteUpdateGold(DuelData.Duelo(Slot).Team(Team).Player(2))
            If DuelData.Duelo(Slot).Drop Then
                Call TirarTodosLosItems(DuelData.Duelo(Slot).Team(Team).Player(2))
            End If
            Call WarpReturnDuelo(DuelData.Duelo(Slot).Team(Team).Player(2))
            UserList(DuelData.Duelo(Slot).Team(Team).Player(2)).flags.DueloIndex = 0
            UserList(DuelData.Duelo(Slot).Team(Team).Player(2)).flags.DueloTeam = 0
            Call WriteConsoleMsg(DuelData.Duelo(Slot).Team(Team).Player(2), "Has perdido el duelo.", FontTypeNames.FONTTYPE_TALK)
            DuelData.Duelo(Slot).Team(Team).Player(2) = 0
        End If
        If DuelData.Duelo(Slot).Team(Team).Player(3) > 0 Then
            UserList(DuelData.Duelo(Slot).Team(Team).Player(3)).Stats.GLD = UserList(DuelData.Duelo(Slot).Team(Team).Player(3)).Stats.GLD - DuelData.Duelo(Slot).Oro
            UserList(DuelData.Duelo(Slot).Team(Team).Player(3)).Stats.OroDuelos = UserList(DuelData.Duelo(Slot).Team(Team).Player(3)).Stats.OroDuelos - DuelData.Duelo(Slot).Oro
            UserList(DuelData.Duelo(Slot).Team(Team).Player(3)).Stats.DuelosPerdidos = UserList(DuelData.Duelo(Slot).Team(Team).Player(3)).Stats.DuelosPerdidos + 1
            If UserList(DuelData.Duelo(Slot).Team(Team).Player(3)).Stats.GLD < 0 Then UserList(DuelData.Duelo(Slot).Team(Team).Player(3)).Stats.GLD = 0
            Call WriteUpdateGold(DuelData.Duelo(Slot).Team(Team).Player(3))
            If DuelData.Duelo(Slot).Drop Then
                Call TirarTodosLosItems(DuelData.Duelo(Slot).Team(Team).Player(3))
            End If
            Call WarpReturnDuelo(DuelData.Duelo(Slot).Team(Team).Player(3))
            UserList(DuelData.Duelo(Slot).Team(Team).Player(3)).flags.DueloIndex = 0
            UserList(DuelData.Duelo(Slot).Team(Team).Player(3)).flags.DueloTeam = 0
            Call WriteConsoleMsg(DuelData.Duelo(Slot).Team(Team).Player(3), "Has perdido el duelo.", FontTypeNames.FONTTYPE_TALK)
            DuelData.Duelo(Slot).Team(Team).Player(3) = 0
        End If
        If DuelData.Duelo(Slot).Team(Team).Player(4) > 0 Then
            UserList(DuelData.Duelo(Slot).Team(Team).Player(4)).Stats.GLD = UserList(DuelData.Duelo(Slot).Team(Team).Player(4)).Stats.GLD - DuelData.Duelo(Slot).Oro
            UserList(DuelData.Duelo(Slot).Team(Team).Player(4)).Stats.OroDuelos = UserList(DuelData.Duelo(Slot).Team(Team).Player(4)).Stats.OroDuelos - DuelData.Duelo(Slot).Oro
            UserList(DuelData.Duelo(Slot).Team(Team).Player(4)).Stats.DuelosPerdidos = UserList(DuelData.Duelo(Slot).Team(Team).Player(4)).Stats.DuelosPerdidos + 1
            If UserList(DuelData.Duelo(Slot).Team(Team).Player(4)).Stats.GLD < 0 Then UserList(DuelData.Duelo(Slot).Team(Team).Player(4)).Stats.GLD = 0
            Call WriteUpdateGold(DuelData.Duelo(Slot).Team(Team).Player(4))
            If DuelData.Duelo(Slot).Drop Then
                Call TirarTodosLosItems(DuelData.Duelo(Slot).Team(Team).Player(4))
            End If
            Call WarpReturnDuelo(DuelData.Duelo(Slot).Team(Team).Player(4))
            UserList(DuelData.Duelo(Slot).Team(Team).Player(4)).flags.DueloIndex = 0
            UserList(DuelData.Duelo(Slot).Team(Team).Player(4)).flags.DueloTeam = 0
            Call WriteConsoleMsg(DuelData.Duelo(Slot).Team(Team).Player(4), "Has perdido el duelo.", FontTypeNames.FONTTYPE_TALK)
            DuelData.Duelo(Slot).Team(Team).Player(4) = 0
        End If
    End If
    Exit Sub

ErrHandler:
        Call LogError("Error en Sub PerdioDuelo. Slot: " & Slot & ". Equipo: " & Team & ". " & Err.Description)
End Sub

Sub GanoDuelo(ByVal Slot As Byte, ByVal Team As Byte)
On Error GoTo ErrHandler

    If Not Slot > 0 Then Exit Sub
    
    DuelData.Duelo(Slot).Ganador = Team
    If GetTipoDuelo(Slot) = eDuelType.vs1 Then
        If DuelData.Duelo(Slot).Team(Team).Player(1) > 0 Then
            UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.GLD = UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.GLD + DuelData.Duelo(Slot).Oro
            UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.OroDuelos = UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.OroDuelos + DuelData.Duelo(Slot).Oro
            UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.DuelosGanados = UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.DuelosGanados + 1
            If UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.GLD > MaxOro Then UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.GLD = MaxOro
            Call WriteUpdateGold(DuelData.Duelo(Slot).Team(Team).Player(1))
        End If
        Call SendData(SendTarget.ToDuelo, Slot, PrepareMessageConsoleMsg("¡Hás ganado el duelo!. Prepárate para regresar en " & DuelData.Duelo(Slot).Counter & " segundos.", FontTypeNames.FONTTYPE_TALK))
    ElseIf GetTipoDuelo(Slot) = eDuelType.vs2 Then
        If DuelData.Duelo(Slot).Team(Team).Player(1) > 0 Then
            UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.GLD = UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.GLD + DuelData.Duelo(Slot).Oro
            UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.OroDuelos = UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.OroDuelos + DuelData.Duelo(Slot).Oro
            UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.DuelosGanados = UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.DuelosGanados + 1
            If UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.GLD > MaxOro Then UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.GLD = MaxOro
            Call WriteUpdateGold(DuelData.Duelo(Slot).Team(Team).Player(1))
        End If
        If DuelData.Duelo(Slot).Team(Team).Player(2) > 0 Then
            UserList(DuelData.Duelo(Slot).Team(Team).Player(2)).Stats.GLD = UserList(DuelData.Duelo(Slot).Team(Team).Player(2)).Stats.GLD + DuelData.Duelo(Slot).Oro
            UserList(DuelData.Duelo(Slot).Team(Team).Player(2)).Stats.OroDuelos = UserList(DuelData.Duelo(Slot).Team(Team).Player(2)).Stats.OroDuelos + DuelData.Duelo(Slot).Oro
            UserList(DuelData.Duelo(Slot).Team(Team).Player(2)).Stats.DuelosGanados = UserList(DuelData.Duelo(Slot).Team(Team).Player(2)).Stats.DuelosGanados + 1
            If UserList(DuelData.Duelo(Slot).Team(Team).Player(2)).Stats.GLD > MaxOro Then UserList(DuelData.Duelo(Slot).Team(Team).Player(2)).Stats.GLD = MaxOro
            Call WriteUpdateGold(DuelData.Duelo(Slot).Team(Team).Player(2))
        End If
        Call SendData(SendTarget.ToDuelo, Slot, PrepareMessageConsoleMsg("¡Hás ganado el duelo!. Prepárate para regresar en " & DuelData.Duelo(Slot).Counter & " segundos.", FontTypeNames.FONTTYPE_TALK))
    ElseIf GetTipoDuelo(Slot) = eDuelType.vs3 Then
        If DuelData.Duelo(Slot).Team(Team).Player(1) > 0 Then
            UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.GLD = UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.GLD + DuelData.Duelo(Slot).Oro
            UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.OroDuelos = UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.OroDuelos + DuelData.Duelo(Slot).Oro
            UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.DuelosGanados = UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.DuelosGanados + 1
            If UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.GLD > MaxOro Then UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.GLD = MaxOro
            Call WriteUpdateGold(DuelData.Duelo(Slot).Team(Team).Player(1))
        End If
        If DuelData.Duelo(Slot).Team(Team).Player(2) > 0 Then
            UserList(DuelData.Duelo(Slot).Team(Team).Player(2)).Stats.GLD = UserList(DuelData.Duelo(Slot).Team(Team).Player(2)).Stats.GLD + DuelData.Duelo(Slot).Oro
            UserList(DuelData.Duelo(Slot).Team(Team).Player(2)).Stats.OroDuelos = UserList(DuelData.Duelo(Slot).Team(Team).Player(2)).Stats.OroDuelos + DuelData.Duelo(Slot).Oro
            UserList(DuelData.Duelo(Slot).Team(Team).Player(2)).Stats.DuelosGanados = UserList(DuelData.Duelo(Slot).Team(Team).Player(2)).Stats.DuelosGanados + 1
            If UserList(DuelData.Duelo(Slot).Team(Team).Player(2)).Stats.GLD > MaxOro Then UserList(DuelData.Duelo(Slot).Team(Team).Player(2)).Stats.GLD = MaxOro
            Call WriteUpdateGold(DuelData.Duelo(Slot).Team(Team).Player(2))
        End If
        If DuelData.Duelo(Slot).Team(Team).Player(3) > 0 Then
            UserList(DuelData.Duelo(Slot).Team(Team).Player(3)).Stats.GLD = UserList(DuelData.Duelo(Slot).Team(Team).Player(3)).Stats.GLD + DuelData.Duelo(Slot).Oro
            UserList(DuelData.Duelo(Slot).Team(Team).Player(3)).Stats.OroDuelos = UserList(DuelData.Duelo(Slot).Team(Team).Player(3)).Stats.OroDuelos + DuelData.Duelo(Slot).Oro
            UserList(DuelData.Duelo(Slot).Team(Team).Player(3)).Stats.DuelosGanados = UserList(DuelData.Duelo(Slot).Team(Team).Player(3)).Stats.DuelosGanados + 1
            If UserList(DuelData.Duelo(Slot).Team(Team).Player(3)).Stats.GLD > MaxOro Then UserList(DuelData.Duelo(Slot).Team(Team).Player(3)).Stats.GLD = MaxOro
            Call WriteUpdateGold(DuelData.Duelo(Slot).Team(Team).Player(3))
        End If
        Call SendData(SendTarget.ToDuelo, Slot, PrepareMessageConsoleMsg("¡Hás ganado el duelo!. Prepárate para regresar en " & DuelData.Duelo(Slot).Counter & " segundos.", FontTypeNames.FONTTYPE_TALK))
    ElseIf GetTipoDuelo(Slot) = eDuelType.vs4 Then
        If DuelData.Duelo(Slot).Team(Team).Player(1) > 0 Then
            UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.GLD = UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.GLD + DuelData.Duelo(Slot).Oro
            UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.OroDuelos = UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.OroDuelos + DuelData.Duelo(Slot).Oro
            UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.DuelosGanados = UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.DuelosGanados + 1
            If UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.GLD > MaxOro Then UserList(DuelData.Duelo(Slot).Team(Team).Player(1)).Stats.GLD = MaxOro
            Call WriteUpdateGold(DuelData.Duelo(Slot).Team(Team).Player(1))
        End If
        If DuelData.Duelo(Slot).Team(Team).Player(2) > 0 Then
            UserList(DuelData.Duelo(Slot).Team(Team).Player(2)).Stats.GLD = UserList(DuelData.Duelo(Slot).Team(Team).Player(2)).Stats.GLD + DuelData.Duelo(Slot).Oro
            UserList(DuelData.Duelo(Slot).Team(Team).Player(2)).Stats.OroDuelos = UserList(DuelData.Duelo(Slot).Team(Team).Player(2)).Stats.OroDuelos + DuelData.Duelo(Slot).Oro
            UserList(DuelData.Duelo(Slot).Team(Team).Player(2)).Stats.DuelosGanados = UserList(DuelData.Duelo(Slot).Team(Team).Player(2)).Stats.DuelosGanados + 1
            If UserList(DuelData.Duelo(Slot).Team(Team).Player(2)).Stats.GLD > MaxOro Then UserList(DuelData.Duelo(Slot).Team(Team).Player(2)).Stats.GLD = MaxOro
            Call WriteUpdateGold(DuelData.Duelo(Slot).Team(Team).Player(2))
        End If
        If DuelData.Duelo(Slot).Team(Team).Player(3) > 0 Then
            UserList(DuelData.Duelo(Slot).Team(Team).Player(3)).Stats.GLD = UserList(DuelData.Duelo(Slot).Team(Team).Player(3)).Stats.GLD + DuelData.Duelo(Slot).Oro
            UserList(DuelData.Duelo(Slot).Team(Team).Player(3)).Stats.OroDuelos = UserList(DuelData.Duelo(Slot).Team(Team).Player(3)).Stats.OroDuelos + DuelData.Duelo(Slot).Oro
            UserList(DuelData.Duelo(Slot).Team(Team).Player(3)).Stats.DuelosGanados = UserList(DuelData.Duelo(Slot).Team(Team).Player(3)).Stats.DuelosGanados + 1
            If UserList(DuelData.Duelo(Slot).Team(Team).Player(3)).Stats.GLD > MaxOro Then UserList(DuelData.Duelo(Slot).Team(Team).Player(3)).Stats.GLD = MaxOro
            Call WriteUpdateGold(DuelData.Duelo(Slot).Team(Team).Player(3))
        End If
        If DuelData.Duelo(Slot).Team(Team).Player(4) > 0 Then
            UserList(DuelData.Duelo(Slot).Team(Team).Player(4)).Stats.GLD = UserList(DuelData.Duelo(Slot).Team(Team).Player(4)).Stats.GLD + DuelData.Duelo(Slot).Oro
            UserList(DuelData.Duelo(Slot).Team(Team).Player(4)).Stats.OroDuelos = UserList(DuelData.Duelo(Slot).Team(Team).Player(4)).Stats.OroDuelos + DuelData.Duelo(Slot).Oro
            UserList(DuelData.Duelo(Slot).Team(Team).Player(4)).Stats.DuelosGanados = UserList(DuelData.Duelo(Slot).Team(Team).Player(4)).Stats.DuelosGanados + 1
            If UserList(DuelData.Duelo(Slot).Team(Team).Player(4)).Stats.GLD > MaxOro Then UserList(DuelData.Duelo(Slot).Team(Team).Player(4)).Stats.GLD = MaxOro
            Call WriteUpdateGold(DuelData.Duelo(Slot).Team(Team).Player(4))
        End If
        Call SendData(SendTarget.ToDuelo, Slot, PrepareMessageConsoleMsg("¡Hás ganado el duelo!. Prepárate para regresar en " & DuelData.Duelo(Slot).Counter & " segundos.", FontTypeNames.FONTTYPE_TALK))
    End If
    Exit Sub

ErrHandler:
        Call LogError("Error en Sub GanoDuelo. Slot: " & Slot & " " & Err.Description)
End Sub

Sub ClearDueloSlot(ByVal Slot As Byte)
On Error GoTo ErrHandler

    If Not Slot > 0 Then Exit Sub
    
    DuelData.Duelo(Slot).estado = eDuelState.Vacio
    DuelData.Duelo(Slot).Counter = 0
    DuelData.Duelo(Slot).Drop = False
    DuelData.Duelo(Slot).Oro = 0
    DuelData.Duelo(Slot).Resucitar = False
    If DuelData.Duelo(Slot).Team(1).Player(1) > 0 Then
        UserList(DuelData.Duelo(Slot).Team(1).Player(1)).flags.DueloIndex = 0
        UserList(DuelData.Duelo(Slot).Team(1).Player(1)).flags.DueloTeam = 0
    End If
    If DuelData.Duelo(Slot).Team(2).Player(1) > 0 Then
        UserList(DuelData.Duelo(Slot).Team(2).Player(1)).flags.DueloIndex = 0
        UserList(DuelData.Duelo(Slot).Team(2).Player(1)).flags.DueloTeam = 0
    End If
    If DuelData.Duelo(Slot).Team(1).Player(2) > 0 Then
        UserList(DuelData.Duelo(Slot).Team(1).Player(2)).flags.DueloIndex = 0
        UserList(DuelData.Duelo(Slot).Team(1).Player(2)).flags.DueloTeam = 0
    End If
    If DuelData.Duelo(Slot).Team(2).Player(2) > 0 Then
        UserList(DuelData.Duelo(Slot).Team(2).Player(2)).flags.DueloIndex = 0
        UserList(DuelData.Duelo(Slot).Team(2).Player(2)).flags.DueloTeam = 0
    End If
    If DuelData.Duelo(Slot).Team(1).Player(3) > 0 Then
        UserList(DuelData.Duelo(Slot).Team(1).Player(3)).flags.DueloIndex = 0
        UserList(DuelData.Duelo(Slot).Team(1).Player(3)).flags.DueloTeam = 0
    End If
    If DuelData.Duelo(Slot).Team(2).Player(3) > 0 Then
        UserList(DuelData.Duelo(Slot).Team(2).Player(3)).flags.DueloIndex = 0
        UserList(DuelData.Duelo(Slot).Team(2).Player(3)).flags.DueloTeam = 0
    End If
    If DuelData.Duelo(Slot).Team(1).Player(4) > 0 Then
        UserList(DuelData.Duelo(Slot).Team(1).Player(4)).flags.DueloIndex = 0
        UserList(DuelData.Duelo(Slot).Team(1).Player(4)).flags.DueloTeam = 0
    End If
    If DuelData.Duelo(Slot).Team(2).Player(4) > 0 Then
        UserList(DuelData.Duelo(Slot).Team(2).Player(4)).flags.DueloIndex = 0
        UserList(DuelData.Duelo(Slot).Team(2).Player(4)).flags.DueloTeam = 0
    End If
    
    If GetTipoDuelo(Slot) = eDuelType.vs1 Then
        Arena1v1(DuelData.Duelo(Slot).Arena).EnUso = 0
    ElseIf GetTipoDuelo(Slot) = eDuelType.vs2 Then
        Arena2v2(DuelData.Duelo(Slot).Arena).EnUso = 0
    ElseIf GetTipoDuelo(Slot) = eDuelType.vs3 Then
        Arena3v3(DuelData.Duelo(Slot).Arena).EnUso = 0
    ElseIf GetTipoDuelo(Slot) = eDuelType.vs4 Then
        Arena4v4(DuelData.Duelo(Slot).Arena).EnUso = 0
    End If
    
    DuelData.Duelo(Slot).Team(1).Player(1) = 0
    DuelData.Duelo(Slot).Team(2).Player(1) = 0
    DuelData.Duelo(Slot).Team(1).Player(2) = 0
    DuelData.Duelo(Slot).Team(2).Player(2) = 0
    DuelData.Duelo(Slot).Team(1).Player(3) = 0
    DuelData.Duelo(Slot).Team(2).Player(3) = 0
    DuelData.Duelo(Slot).Team(1).Player(4) = 0
    DuelData.Duelo(Slot).Team(2).Player(4) = 0
    DuelData.Duelo(Slot).Team(1).Muerto(1) = False
    DuelData.Duelo(Slot).Team(2).Muerto(1) = False
    DuelData.Duelo(Slot).Team(1).Muerto(2) = False
    DuelData.Duelo(Slot).Team(2).Muerto(2) = False
    DuelData.Duelo(Slot).Team(1).Muerto(3) = False
    DuelData.Duelo(Slot).Team(2).Muerto(3) = False
    DuelData.Duelo(Slot).Team(1).Muerto(4) = False
    DuelData.Duelo(Slot).Team(2).Muerto(4) = False
    DuelData.Duelo(Slot).Ganador = 0
    Call LimpiarArenaItems(DuelData.Duelo(Slot).TipoDuelo, Slot)
    DuelData.Duelo(Slot).TipoDuelo = 0
    DuelData.Duelo(Slot).Arena = 0
    Exit Sub

ErrHandler:
        Call LogError("Error en Sub CleanDueloSlot. Slot: " & Slot & " " & Err.Description)
End Sub

Sub LimpiarArenaItems(ByVal TipoArena As Byte, ByVal Slot As Byte)
On Error GoTo ErrHandler
    Dim X As Integer, Y As Integer
    
    If TipoArena = eDuelType.vs1 Then
        With Arena1v1(DuelData.Duelo(Slot).Arena)
            For X = .X1 To .X2
                For Y = .Y1 To .Y2
                    If MapData(.Map, X, Y).ObjInfo.ObjIndex > 0 Then
                        Call EraseObj(MapData(.Map, X, Y).ObjInfo.Amount, .Map, X, Y)
                    End If
                Next Y
            Next X
        End With
    ElseIf TipoArena = eDuelType.vs2 Then
        With Arena2v2(DuelData.Duelo(Slot).Arena)
            For X = .X1 To .X2
                For Y = .Y1 To .Y2
                    If MapData(.Map, X, Y).ObjInfo.ObjIndex > 0 Then
                        Call EraseObj(MapData(.Map, X, Y).ObjInfo.Amount, .Map, X, Y)
                    End If
                Next Y
            Next X
        End With
    ElseIf TipoArena = eDuelType.vs3 Then
        With Arena3v3(DuelData.Duelo(Slot).Arena)
            For X = .X1 To .X2
                For Y = .Y1 To .Y2
                    If MapData(.Map, X, Y).ObjInfo.ObjIndex > 0 Then
                        Call EraseObj(MapData(.Map, X, Y).ObjInfo.Amount, .Map, X, Y)
                    End If
                Next Y
            Next X
        End With
    ElseIf TipoArena = eDuelType.vs4 Then
        With Arena4v4(DuelData.Duelo(Slot).Arena)
            For X = .X1 To .X2
                For Y = .Y1 To .Y2
                    If MapData(.Map, X, Y).ObjInfo.ObjIndex > 0 Then
                        Call EraseObj(MapData(.Map, X, Y).ObjInfo.Amount, .Map, X, Y)
                    End If
                Next Y
            Next X
        End With
    End If
    Exit Sub

ErrHandler:
        Call LogError("Error en Sub LimpiarArenaItems. Slot: " & Slot & " " & Err.Description)
End Sub

Public Function EnMapaDuelos(ByVal UserIndex As Integer) As Boolean
On Error GoTo ErrHandler
  
    Dim I As Byte
    
    For I = LBound(MapArenas) To UBound(MapArenas)
        If UserList(UserIndex).Pos.Map = MapArenas(I) Then
            EnMapaDuelos = True
            Exit Function
        End If
    Next I
    
    EnMapaDuelos = False
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function EnMapaDuelos de modDuelos.bas")
End Function

Sub CheckDueloPlayersState(ByVal Slot As Byte)
On Error GoTo ErrHandler
  
    If Not Slot > 0 Then Exit Sub
    
    Dim I As Byte
    If GetTipoDuelo(Slot) = eDuelType.vs1 Then
        For I = 1 To 2
            If DuelData.Duelo(Slot).Team(I).Muerto(1) = True Then
                Call TerminoDuelo(Slot, I)
                Exit Sub
            End If
        Next I
    ElseIf GetTipoDuelo(Slot) = eDuelType.vs2 Then
        For I = 1 To 2
            If DuelData.Duelo(Slot).Team(I).Muerto(1) = True Then
                If DuelData.Duelo(Slot).Team(I).Muerto(2) = True Then
                    Call TerminoDuelo(Slot, I)
                    Exit Sub
                End If
            End If
        Next I
    ElseIf GetTipoDuelo(Slot) = eDuelType.vs3 Then
        For I = 1 To 2
            If DuelData.Duelo(Slot).Team(I).Muerto(1) = True Then
                If DuelData.Duelo(Slot).Team(I).Muerto(2) = True Then
                    If DuelData.Duelo(Slot).Team(I).Muerto(3) = True Then
                        Call TerminoDuelo(Slot, I)
                        Exit Sub
                    End If
                End If
            End If
        Next I
    ElseIf GetTipoDuelo(Slot) = eDuelType.vs4 Then
        For I = 1 To 2
            If DuelData.Duelo(Slot).Team(I).Muerto(1) = True Then
                If DuelData.Duelo(Slot).Team(I).Muerto(2) = True Then
                    If DuelData.Duelo(Slot).Team(I).Muerto(3) = True Then
                        If DuelData.Duelo(Slot).Team(I).Muerto(4) = True Then
                            Call TerminoDuelo(Slot, I)
                            Exit Sub
                        End If
                    End If
                End If
            End If
        Next I
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CheckDueloPlayersState de modDuelos.bas")
End Sub

Function CondicionPlayerOK(ByVal Slot As Byte) As Boolean
On Error GoTo ErrHandler
  
    If Not Slot > 0 Then Exit Function
    
    If GetTipoDuelo(Slot) = eDuelType.vs1 Then
        If Not PuedeAceptarDuelo(DuelData.Duelo(Slot).Team(1).Player(1), Slot) Or _
            Not PuedeAceptarDuelo(DuelData.Duelo(Slot).Team(2).Player(1), Slot) Then
            CondicionPlayerOK = False
            Exit Function
        End If
    ElseIf GetTipoDuelo(Slot) = eDuelType.vs2 Then
        If Not PuedeAceptarDuelo(DuelData.Duelo(Slot).Team(1).Player(1), Slot) Or _
            Not PuedeAceptarDuelo(DuelData.Duelo(Slot).Team(1).Player(2), Slot) Or _
            Not PuedeAceptarDuelo(DuelData.Duelo(Slot).Team(2).Player(1), Slot) Or _
            Not PuedeAceptarDuelo(DuelData.Duelo(Slot).Team(2).Player(2), Slot) Then
            CondicionPlayerOK = False
            Exit Function
        End If
    ElseIf GetTipoDuelo(Slot) = eDuelType.vs3 Then
        If Not PuedeAceptarDuelo(DuelData.Duelo(Slot).Team(1).Player(1), Slot) Or _
            Not PuedeAceptarDuelo(DuelData.Duelo(Slot).Team(1).Player(2), Slot) Or _
            Not PuedeAceptarDuelo(DuelData.Duelo(Slot).Team(1).Player(3), Slot) Or _
            Not PuedeAceptarDuelo(DuelData.Duelo(Slot).Team(2).Player(1), Slot) Or _
            Not PuedeAceptarDuelo(DuelData.Duelo(Slot).Team(2).Player(2), Slot) Or _
            Not PuedeAceptarDuelo(DuelData.Duelo(Slot).Team(2).Player(3), Slot) Then
            CondicionPlayerOK = False
            Exit Function
        End If
    ElseIf GetTipoDuelo(Slot) = eDuelType.vs4 Then
        If Not PuedeAceptarDuelo(DuelData.Duelo(Slot).Team(1).Player(1), Slot) Or _
            Not PuedeAceptarDuelo(DuelData.Duelo(Slot).Team(1).Player(2), Slot) Or _
            Not PuedeAceptarDuelo(DuelData.Duelo(Slot).Team(1).Player(3), Slot) Or _
            Not PuedeAceptarDuelo(DuelData.Duelo(Slot).Team(1).Player(4), Slot) Or _
            Not PuedeAceptarDuelo(DuelData.Duelo(Slot).Team(2).Player(1), Slot) Or _
            Not PuedeAceptarDuelo(DuelData.Duelo(Slot).Team(2).Player(2), Slot) Or _
            Not PuedeAceptarDuelo(DuelData.Duelo(Slot).Team(2).Player(3), Slot) Or _
            Not PuedeAceptarDuelo(DuelData.Duelo(Slot).Team(2).Player(4), Slot) Then
            CondicionPlayerOK = False
            Exit Function
        End If
    End If
    
    CondicionPlayerOK = True
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function CondicionPlayerOK de modDuelos.bas")
End Function

Sub InvocarDueloPlayers(ByVal Slot As Byte)
On Error GoTo ErrHandler
  
    If Not CondicionPlayerOK(Slot) Then
        Call SendData(SendTarget.ToDuelo, Slot, PrepareMessageConsoleMsg("El duelo ha sido cancelado, pues no todos los jugadores cumplían los requisitos necesarios.", FontTypeNames.FONTTYPE_INFO))
        Call ClearDueloSlot(Slot)
        Exit Sub
    End If
    
    Call WarpArena(Slot)
    DuelData.Duelo(Slot).Counter = 5
    DuelData.Duelo(Slot).estado = eDuelState.Esperando_Inicio
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub InvocarDueloPlayers de modDuelos.bas")
End Sub

Sub IniciarDuelo(ByVal Slot As Byte)
On Error GoTo ErrHandler
  
    If Not Slot > 0 Then Exit Sub
    
    DuelData.Duelo(Slot).Counter = DuelData.DuelDuration
    
    DuelData.Duelo(Slot).estado = eDuelState.Iniciado
    Call SendData(SendTarget.ToDuelo, Slot, PrepareMessageConsoleMsg("¡El duelo ha iniciado!", FontTypeNames.FONTTYPE_TALK))
    If GetTipoDuelo(Slot) = eDuelType.vs1 Then
        If DuelData.Duelo(Slot).Team(1).Player(1) > 0 Then Call WriteParalizeOK(DuelData.Duelo(Slot).Team(1).Player(1), False)
        If DuelData.Duelo(Slot).Team(2).Player(1) > 0 Then Call WriteParalizeOK(DuelData.Duelo(Slot).Team(2).Player(1), False)
    ElseIf GetTipoDuelo(Slot) = eDuelType.vs2 Then
        If DuelData.Duelo(Slot).Team(1).Player(1) > 0 Then Call WriteParalizeOK(DuelData.Duelo(Slot).Team(1).Player(1), False)
        If DuelData.Duelo(Slot).Team(2).Player(1) > 0 Then Call WriteParalizeOK(DuelData.Duelo(Slot).Team(2).Player(1), False)
        If DuelData.Duelo(Slot).Team(1).Player(2) > 0 Then Call WriteParalizeOK(DuelData.Duelo(Slot).Team(1).Player(2), False)
        If DuelData.Duelo(Slot).Team(2).Player(2) > 0 Then Call WriteParalizeOK(DuelData.Duelo(Slot).Team(2).Player(2), False)
    ElseIf GetTipoDuelo(Slot) = eDuelType.vs3 Then
        If DuelData.Duelo(Slot).Team(1).Player(1) > 0 Then Call WriteParalizeOK(DuelData.Duelo(Slot).Team(1).Player(1), False)
        If DuelData.Duelo(Slot).Team(2).Player(1) > 0 Then Call WriteParalizeOK(DuelData.Duelo(Slot).Team(2).Player(1), False)
        If DuelData.Duelo(Slot).Team(1).Player(2) > 0 Then Call WriteParalizeOK(DuelData.Duelo(Slot).Team(1).Player(2), False)
        If DuelData.Duelo(Slot).Team(2).Player(2) > 0 Then Call WriteParalizeOK(DuelData.Duelo(Slot).Team(2).Player(2), False)
        If DuelData.Duelo(Slot).Team(1).Player(3) > 0 Then Call WriteParalizeOK(DuelData.Duelo(Slot).Team(1).Player(3), False)
        If DuelData.Duelo(Slot).Team(2).Player(3) > 0 Then Call WriteParalizeOK(DuelData.Duelo(Slot).Team(2).Player(3), False)
    ElseIf GetTipoDuelo(Slot) = eDuelType.vs4 Then
        If DuelData.Duelo(Slot).Team(1).Player(1) > 0 Then Call WriteParalizeOK(DuelData.Duelo(Slot).Team(1).Player(1), False)
        If DuelData.Duelo(Slot).Team(2).Player(1) > 0 Then Call WriteParalizeOK(DuelData.Duelo(Slot).Team(2).Player(1), False)
        If DuelData.Duelo(Slot).Team(1).Player(2) > 0 Then Call WriteParalizeOK(DuelData.Duelo(Slot).Team(1).Player(2), False)
        If DuelData.Duelo(Slot).Team(2).Player(2) > 0 Then Call WriteParalizeOK(DuelData.Duelo(Slot).Team(2).Player(2), False)
        If DuelData.Duelo(Slot).Team(1).Player(3) > 0 Then Call WriteParalizeOK(DuelData.Duelo(Slot).Team(1).Player(3), False)
        If DuelData.Duelo(Slot).Team(2).Player(3) > 0 Then Call WriteParalizeOK(DuelData.Duelo(Slot).Team(2).Player(3), False)
        If DuelData.Duelo(Slot).Team(1).Player(4) > 0 Then Call WriteParalizeOK(DuelData.Duelo(Slot).Team(1).Player(4), False)
        If DuelData.Duelo(Slot).Team(2).Player(4) > 0 Then Call WriteParalizeOK(DuelData.Duelo(Slot).Team(2).Player(4), False)
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub IniciarDuelo de modDuelos.bas")
End Sub

Sub CancelarDuelo(ByVal Slot As Byte, Optional ByVal Normal As Boolean = True)
On Error GoTo ErrHandler
  
    Dim I As Integer

    For I = 1 To LastUser
        If UserList(I).flags.UserLogged Then
            If UserList(I).flags.DueloIndex = Slot Then
                UserList(I).flags.DueloIndex = 0
                UserList(I).flags.DueloTeam = 0
                If Normal Then
                    Call WriteConsoleMsg(I, "El duelo ha sido rechazado.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(I, "El duelo no ha sido aceptado a tiempo.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
    Next I
    
    Call ClearDueloSlot(Slot)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CancelarDuelo de modDuelos.bas")
End Sub

Sub CerrarDuelo(ByVal Slot As Byte)
On Error GoTo ErrHandler

    If Not Slot > 0 Then Exit Sub
        If GetTipoDuelo(Slot) = eDuelType.vs1 Then
            If DuelData.Duelo(Slot).Team(1).Player(1) > 0 Then _
                Call WarpReturnDuelo(DuelData.Duelo(Slot).Team(1).Player(1))
            If DuelData.Duelo(Slot).Team(2).Player(1) > 0 Then _
                Call WarpReturnDuelo(DuelData.Duelo(Slot).Team(2).Player(1))
        ElseIf GetTipoDuelo(Slot) = eDuelType.vs2 Then
            If DuelData.Duelo(Slot).Team(1).Player(1) > 0 Then _
                Call WarpReturnDuelo(DuelData.Duelo(Slot).Team(1).Player(1))
            If DuelData.Duelo(Slot).Team(1).Player(2) > 0 Then _
                Call WarpReturnDuelo(DuelData.Duelo(Slot).Team(1).Player(2))
            If DuelData.Duelo(Slot).Team(2).Player(1) > 0 Then _
                Call WarpReturnDuelo(DuelData.Duelo(Slot).Team(2).Player(1))
            If DuelData.Duelo(Slot).Team(2).Player(2) > 0 Then _
                Call WarpReturnDuelo(DuelData.Duelo(Slot).Team(2).Player(2))
        ElseIf GetTipoDuelo(Slot) = eDuelType.vs3 Then
            If DuelData.Duelo(Slot).Team(1).Player(1) > 0 Then _
                Call WarpReturnDuelo(DuelData.Duelo(Slot).Team(1).Player(1))
            If DuelData.Duelo(Slot).Team(1).Player(2) > 0 Then _
                Call WarpReturnDuelo(DuelData.Duelo(Slot).Team(1).Player(2))
            If DuelData.Duelo(Slot).Team(1).Player(3) > 0 Then _
                Call WarpReturnDuelo(DuelData.Duelo(Slot).Team(1).Player(3))
            If DuelData.Duelo(Slot).Team(2).Player(1) > 0 Then _
                Call WarpReturnDuelo(DuelData.Duelo(Slot).Team(2).Player(1))
            If DuelData.Duelo(Slot).Team(2).Player(2) > 0 Then _
                Call WarpReturnDuelo(DuelData.Duelo(Slot).Team(2).Player(2))
            If DuelData.Duelo(Slot).Team(2).Player(3) > 0 Then _
                Call WarpReturnDuelo(DuelData.Duelo(Slot).Team(2).Player(3))
        ElseIf GetTipoDuelo(Slot) = eDuelType.vs4 Then
            If DuelData.Duelo(Slot).Team(1).Player(1) > 0 Then _
                Call WarpReturnDuelo(DuelData.Duelo(Slot).Team(1).Player(1))
            If DuelData.Duelo(Slot).Team(1).Player(2) > 0 Then _
                Call WarpReturnDuelo(DuelData.Duelo(Slot).Team(1).Player(2))
            If DuelData.Duelo(Slot).Team(1).Player(3) > 0 Then _
                Call WarpReturnDuelo(DuelData.Duelo(Slot).Team(1).Player(3))
            If DuelData.Duelo(Slot).Team(1).Player(4) > 0 Then _
                Call WarpReturnDuelo(DuelData.Duelo(Slot).Team(1).Player(4))
            If DuelData.Duelo(Slot).Team(2).Player(1) > 0 Then _
                Call WarpReturnDuelo(DuelData.Duelo(Slot).Team(2).Player(1))
            If DuelData.Duelo(Slot).Team(2).Player(2) > 0 Then _
                Call WarpReturnDuelo(DuelData.Duelo(Slot).Team(2).Player(2))
            If DuelData.Duelo(Slot).Team(2).Player(3) > 0 Then _
                Call WarpReturnDuelo(DuelData.Duelo(Slot).Team(2).Player(3))
            If DuelData.Duelo(Slot).Team(2).Player(4) > 0 Then _
                Call WarpReturnDuelo(DuelData.Duelo(Slot).Team(2).Player(4))
        End If
        Call ClearDueloSlot(Slot)
    Exit Sub

ErrHandler:
        Call LogError("Error en Sub CerrarDuelo. Slot: " & Slot & " " & Err.Description)
End Sub

Sub WarpArena(ByVal Slot As Byte)
On Error GoTo ErrHandler

    If Not Slot > 0 Then Exit Sub
    Dim Pos1 As WorldPos
    Dim Pos2 As WorldPos
    Dim Pos3 As WorldPos
    Dim Pos4 As WorldPos
    Dim Pos5 As WorldPos
    Dim Pos6 As WorldPos
    Dim Pos7 As WorldPos
    Dim Pos8 As WorldPos
    
    If GetTipoDuelo(Slot) = eDuelType.vs1 Then
        Pos1.Map = Arena1v1(DuelData.Duelo(Slot).Arena).Map
        Pos1.X = Arena1v1(DuelData.Duelo(Slot).Arena).X1
        Pos1.Y = Arena1v1(DuelData.Duelo(Slot).Arena).Y1
        Pos2.Map = Arena1v1(DuelData.Duelo(Slot).Arena).Map
        Pos2.X = Arena1v1(DuelData.Duelo(Slot).Arena).X2
        Pos2.Y = Arena1v1(DuelData.Duelo(Slot).Arena).Y2
        UserList(DuelData.Duelo(Slot).Team(1).Player(1)).VolverDueloPos = UserList(DuelData.Duelo(Slot).Team(1).Player(1)).Pos
        UserList(DuelData.Duelo(Slot).Team(2).Player(1)).VolverDueloPos = UserList(DuelData.Duelo(Slot).Team(2).Player(1)).Pos
        Call UsUaRiOs.SetInvisible(DuelData.Duelo(Slot).Team(1).Player(1), UserList(DuelData.Duelo(Slot).Team(1).Player(1)).Char.CharIndex, False)
        Call UsUaRiOs.SetInvisible(DuelData.Duelo(Slot).Team(2).Player(1), UserList(DuelData.Duelo(Slot).Team(2).Player(1)).Char.CharIndex, False)
        Call WarpUserChar(DuelData.Duelo(Slot).Team(1).Player(1), Pos1.Map, Pos1.X, Pos1.Y, False)
        Call WarpUserChar(DuelData.Duelo(Slot).Team(2).Player(1), Pos2.Map, Pos2.X, Pos2.Y, False)
        Call WriteParalizeOK(DuelData.Duelo(Slot).Team(1).Player(1), False)
        Call WriteParalizeOK(DuelData.Duelo(Slot).Team(2).Player(1), False)
    ElseIf GetTipoDuelo(Slot) = eDuelType.vs2 Then
        Pos1.Map = Arena2v2(DuelData.Duelo(Slot).Arena).Map
        Pos1.X = Arena2v2(DuelData.Duelo(Slot).Arena).X1
        Pos1.Y = Arena2v2(DuelData.Duelo(Slot).Arena).Y1
        Pos2.Map = Arena2v2(DuelData.Duelo(Slot).Arena).Map
        Pos2.X = Arena2v2(DuelData.Duelo(Slot).Arena).X1 + 1
        Pos2.Y = Arena2v2(DuelData.Duelo(Slot).Arena).Y1
        Pos3.Map = Arena2v2(DuelData.Duelo(Slot).Arena).Map
        Pos3.X = Arena2v2(DuelData.Duelo(Slot).Arena).X2
        Pos3.Y = Arena2v2(DuelData.Duelo(Slot).Arena).Y2
        Pos4.Map = Arena2v2(DuelData.Duelo(Slot).Arena).Map
        Pos4.X = Arena2v2(DuelData.Duelo(Slot).Arena).X2 - 1
        Pos4.Y = Arena2v2(DuelData.Duelo(Slot).Arena).Y2
        UserList(DuelData.Duelo(Slot).Team(1).Player(1)).VolverDueloPos = UserList(DuelData.Duelo(Slot).Team(1).Player(1)).Pos
        UserList(DuelData.Duelo(Slot).Team(1).Player(2)).VolverDueloPos = UserList(DuelData.Duelo(Slot).Team(1).Player(2)).Pos
        UserList(DuelData.Duelo(Slot).Team(2).Player(1)).VolverDueloPos = UserList(DuelData.Duelo(Slot).Team(2).Player(1)).Pos
        UserList(DuelData.Duelo(Slot).Team(2).Player(2)).VolverDueloPos = UserList(DuelData.Duelo(Slot).Team(2).Player(2)).Pos
        Call UsUaRiOs.SetInvisible(DuelData.Duelo(Slot).Team(1).Player(1), UserList(DuelData.Duelo(Slot).Team(1).Player(1)).Char.CharIndex, False)
        Call UsUaRiOs.SetInvisible(DuelData.Duelo(Slot).Team(1).Player(2), UserList(DuelData.Duelo(Slot).Team(1).Player(2)).Char.CharIndex, False)
        Call UsUaRiOs.SetInvisible(DuelData.Duelo(Slot).Team(2).Player(1), UserList(DuelData.Duelo(Slot).Team(2).Player(1)).Char.CharIndex, False)
        Call UsUaRiOs.SetInvisible(DuelData.Duelo(Slot).Team(2).Player(2), UserList(DuelData.Duelo(Slot).Team(2).Player(2)).Char.CharIndex, False)
        Call WarpUserChar(DuelData.Duelo(Slot).Team(1).Player(1), Pos1.Map, Pos1.X, Pos1.Y, False)
        Call WarpUserChar(DuelData.Duelo(Slot).Team(1).Player(2), Pos2.Map, Pos2.X, Pos2.Y, False)
        Call WarpUserChar(DuelData.Duelo(Slot).Team(2).Player(1), Pos3.Map, Pos3.X, Pos3.Y, False)
        Call WarpUserChar(DuelData.Duelo(Slot).Team(2).Player(2), Pos4.Map, Pos4.X, Pos4.Y, False)
        Call WriteParalizeOK(DuelData.Duelo(Slot).Team(1).Player(1), False)
        Call WriteParalizeOK(DuelData.Duelo(Slot).Team(2).Player(1), False)
        Call WriteParalizeOK(DuelData.Duelo(Slot).Team(1).Player(2), False)
        Call WriteParalizeOK(DuelData.Duelo(Slot).Team(2).Player(2), False)
    ElseIf GetTipoDuelo(Slot) = eDuelType.vs3 Then
        Pos1.Map = Arena2v2(DuelData.Duelo(Slot).Arena).Map
        Pos1.X = Arena2v2(DuelData.Duelo(Slot).Arena).X1
        Pos1.Y = Arena2v2(DuelData.Duelo(Slot).Arena).Y1
        Pos2.Map = Arena2v2(DuelData.Duelo(Slot).Arena).Map
        Pos2.X = Arena2v2(DuelData.Duelo(Slot).Arena).X1 + 1
        Pos2.Y = Arena2v2(DuelData.Duelo(Slot).Arena).Y1
        Pos3.Map = Arena2v2(DuelData.Duelo(Slot).Arena).Map
        Pos3.X = Arena2v2(DuelData.Duelo(Slot).Arena).X1
        Pos3.Y = Arena2v2(DuelData.Duelo(Slot).Arena).Y1 + 1
        Pos4.Map = Arena2v2(DuelData.Duelo(Slot).Arena).Map
        Pos4.X = Arena2v2(DuelData.Duelo(Slot).Arena).X2
        Pos4.Y = Arena2v2(DuelData.Duelo(Slot).Arena).Y2
        Pos5.Map = Arena2v2(DuelData.Duelo(Slot).Arena).Map
        Pos5.X = Arena2v2(DuelData.Duelo(Slot).Arena).X2 - 1
        Pos5.Y = Arena2v2(DuelData.Duelo(Slot).Arena).Y2
        Pos6.Map = Arena2v2(DuelData.Duelo(Slot).Arena).Map
        Pos6.X = Arena2v2(DuelData.Duelo(Slot).Arena).X2
        Pos6.Y = Arena2v2(DuelData.Duelo(Slot).Arena).Y2 - 1
        UserList(DuelData.Duelo(Slot).Team(1).Player(1)).VolverDueloPos = UserList(DuelData.Duelo(Slot).Team(1).Player(1)).Pos
        UserList(DuelData.Duelo(Slot).Team(1).Player(2)).VolverDueloPos = UserList(DuelData.Duelo(Slot).Team(1).Player(2)).Pos
        UserList(DuelData.Duelo(Slot).Team(1).Player(3)).VolverDueloPos = UserList(DuelData.Duelo(Slot).Team(1).Player(3)).Pos
        UserList(DuelData.Duelo(Slot).Team(2).Player(1)).VolverDueloPos = UserList(DuelData.Duelo(Slot).Team(2).Player(1)).Pos
        UserList(DuelData.Duelo(Slot).Team(2).Player(2)).VolverDueloPos = UserList(DuelData.Duelo(Slot).Team(2).Player(2)).Pos
        UserList(DuelData.Duelo(Slot).Team(2).Player(3)).VolverDueloPos = UserList(DuelData.Duelo(Slot).Team(2).Player(3)).Pos
        Call UsUaRiOs.SetInvisible(DuelData.Duelo(Slot).Team(1).Player(1), UserList(DuelData.Duelo(Slot).Team(1).Player(1)).Char.CharIndex, False)
        Call UsUaRiOs.SetInvisible(DuelData.Duelo(Slot).Team(1).Player(2), UserList(DuelData.Duelo(Slot).Team(1).Player(2)).Char.CharIndex, False)
        Call UsUaRiOs.SetInvisible(DuelData.Duelo(Slot).Team(1).Player(3), UserList(DuelData.Duelo(Slot).Team(1).Player(3)).Char.CharIndex, False)
        Call UsUaRiOs.SetInvisible(DuelData.Duelo(Slot).Team(2).Player(1), UserList(DuelData.Duelo(Slot).Team(2).Player(1)).Char.CharIndex, False)
        Call UsUaRiOs.SetInvisible(DuelData.Duelo(Slot).Team(2).Player(2), UserList(DuelData.Duelo(Slot).Team(2).Player(2)).Char.CharIndex, False)
        Call UsUaRiOs.SetInvisible(DuelData.Duelo(Slot).Team(2).Player(3), UserList(DuelData.Duelo(Slot).Team(2).Player(3)).Char.CharIndex, False)
        Call WarpUserChar(DuelData.Duelo(Slot).Team(1).Player(1), Pos1.Map, Pos1.X, Pos1.Y, False)
        Call WarpUserChar(DuelData.Duelo(Slot).Team(1).Player(2), Pos2.Map, Pos2.X, Pos2.Y, False)
        Call WarpUserChar(DuelData.Duelo(Slot).Team(1).Player(3), Pos3.Map, Pos3.X, Pos3.Y, False)
        Call WarpUserChar(DuelData.Duelo(Slot).Team(2).Player(1), Pos4.Map, Pos4.X, Pos4.Y, False)
        Call WarpUserChar(DuelData.Duelo(Slot).Team(2).Player(2), Pos5.Map, Pos5.X, Pos5.Y, False)
        Call WarpUserChar(DuelData.Duelo(Slot).Team(2).Player(3), Pos6.Map, Pos6.X, Pos6.Y, False)
        Call WriteParalizeOK(DuelData.Duelo(Slot).Team(1).Player(1), False)
        Call WriteParalizeOK(DuelData.Duelo(Slot).Team(2).Player(1), False)
        Call WriteParalizeOK(DuelData.Duelo(Slot).Team(1).Player(2), False)
        Call WriteParalizeOK(DuelData.Duelo(Slot).Team(2).Player(2), False)
        Call WriteParalizeOK(DuelData.Duelo(Slot).Team(1).Player(3), False)
        Call WriteParalizeOK(DuelData.Duelo(Slot).Team(2).Player(3), False)
    ElseIf GetTipoDuelo(Slot) = eDuelType.vs4 Then
        Pos1.Map = Arena2v2(DuelData.Duelo(Slot).Arena).Map
        Pos1.X = Arena2v2(DuelData.Duelo(Slot).Arena).X1
        Pos1.Y = Arena2v2(DuelData.Duelo(Slot).Arena).Y1
        Pos2.Map = Arena2v2(DuelData.Duelo(Slot).Arena).Map
        Pos2.X = Arena2v2(DuelData.Duelo(Slot).Arena).X1 + 1
        Pos2.Y = Arena2v2(DuelData.Duelo(Slot).Arena).Y1
        Pos3.Map = Arena2v2(DuelData.Duelo(Slot).Arena).Map
        Pos3.X = Arena2v2(DuelData.Duelo(Slot).Arena).X1
        Pos3.Y = Arena2v2(DuelData.Duelo(Slot).Arena).Y1 + 1
        Pos4.Map = Arena2v2(DuelData.Duelo(Slot).Arena).Map
        Pos4.X = Arena2v2(DuelData.Duelo(Slot).Arena).X1 + 1
        Pos4.Y = Arena2v2(DuelData.Duelo(Slot).Arena).Y1 + 1
        Pos5.Map = Arena2v2(DuelData.Duelo(Slot).Arena).Map
        Pos5.X = Arena2v2(DuelData.Duelo(Slot).Arena).X2
        Pos5.Y = Arena2v2(DuelData.Duelo(Slot).Arena).Y2
        Pos6.Map = Arena2v2(DuelData.Duelo(Slot).Arena).Map
        Pos6.X = Arena2v2(DuelData.Duelo(Slot).Arena).X2 - 1
        Pos6.Y = Arena2v2(DuelData.Duelo(Slot).Arena).Y2
        Pos7.Map = Arena2v2(DuelData.Duelo(Slot).Arena).Map
        Pos7.X = Arena2v2(DuelData.Duelo(Slot).Arena).X2
        Pos7.Y = Arena2v2(DuelData.Duelo(Slot).Arena).Y2 - 1
        Pos8.Map = Arena2v2(DuelData.Duelo(Slot).Arena).Map
        Pos8.X = Arena2v2(DuelData.Duelo(Slot).Arena).X2 - 1
        Pos8.Y = Arena2v2(DuelData.Duelo(Slot).Arena).Y2 - 1
        UserList(DuelData.Duelo(Slot).Team(1).Player(1)).VolverDueloPos = UserList(DuelData.Duelo(Slot).Team(1).Player(1)).Pos
        UserList(DuelData.Duelo(Slot).Team(1).Player(2)).VolverDueloPos = UserList(DuelData.Duelo(Slot).Team(1).Player(2)).Pos
        UserList(DuelData.Duelo(Slot).Team(1).Player(3)).VolverDueloPos = UserList(DuelData.Duelo(Slot).Team(1).Player(3)).Pos
        UserList(DuelData.Duelo(Slot).Team(1).Player(4)).VolverDueloPos = UserList(DuelData.Duelo(Slot).Team(1).Player(4)).Pos
        UserList(DuelData.Duelo(Slot).Team(2).Player(1)).VolverDueloPos = UserList(DuelData.Duelo(Slot).Team(2).Player(1)).Pos
        UserList(DuelData.Duelo(Slot).Team(2).Player(2)).VolverDueloPos = UserList(DuelData.Duelo(Slot).Team(2).Player(2)).Pos
        UserList(DuelData.Duelo(Slot).Team(2).Player(3)).VolverDueloPos = UserList(DuelData.Duelo(Slot).Team(2).Player(3)).Pos
        UserList(DuelData.Duelo(Slot).Team(2).Player(4)).VolverDueloPos = UserList(DuelData.Duelo(Slot).Team(2).Player(4)).Pos
        Call UsUaRiOs.SetInvisible(DuelData.Duelo(Slot).Team(1).Player(1), UserList(DuelData.Duelo(Slot).Team(1).Player(1)).Char.CharIndex, False)
        Call UsUaRiOs.SetInvisible(DuelData.Duelo(Slot).Team(1).Player(2), UserList(DuelData.Duelo(Slot).Team(1).Player(2)).Char.CharIndex, False)
        Call UsUaRiOs.SetInvisible(DuelData.Duelo(Slot).Team(1).Player(3), UserList(DuelData.Duelo(Slot).Team(1).Player(3)).Char.CharIndex, False)
        Call UsUaRiOs.SetInvisible(DuelData.Duelo(Slot).Team(1).Player(4), UserList(DuelData.Duelo(Slot).Team(1).Player(4)).Char.CharIndex, False)
        Call UsUaRiOs.SetInvisible(DuelData.Duelo(Slot).Team(2).Player(1), UserList(DuelData.Duelo(Slot).Team(2).Player(1)).Char.CharIndex, False)
        Call UsUaRiOs.SetInvisible(DuelData.Duelo(Slot).Team(2).Player(2), UserList(DuelData.Duelo(Slot).Team(2).Player(2)).Char.CharIndex, False)
        Call UsUaRiOs.SetInvisible(DuelData.Duelo(Slot).Team(2).Player(3), UserList(DuelData.Duelo(Slot).Team(2).Player(3)).Char.CharIndex, False)
        Call UsUaRiOs.SetInvisible(DuelData.Duelo(Slot).Team(2).Player(4), UserList(DuelData.Duelo(Slot).Team(2).Player(4)).Char.CharIndex, False)
        Call WarpUserChar(DuelData.Duelo(Slot).Team(1).Player(1), Pos1.Map, Pos1.X, Pos1.Y, False)
        Call WarpUserChar(DuelData.Duelo(Slot).Team(1).Player(2), Pos2.Map, Pos2.X, Pos2.Y, False)
        Call WarpUserChar(DuelData.Duelo(Slot).Team(1).Player(3), Pos3.Map, Pos3.X, Pos3.Y, False)
        Call WarpUserChar(DuelData.Duelo(Slot).Team(1).Player(4), Pos4.Map, Pos4.X, Pos4.Y, False)
        Call WarpUserChar(DuelData.Duelo(Slot).Team(2).Player(1), Pos5.Map, Pos5.X, Pos5.Y, False)
        Call WarpUserChar(DuelData.Duelo(Slot).Team(2).Player(2), Pos6.Map, Pos6.X, Pos6.Y, False)
        Call WarpUserChar(DuelData.Duelo(Slot).Team(2).Player(3), Pos7.Map, Pos7.X, Pos7.Y, False)
        Call WarpUserChar(DuelData.Duelo(Slot).Team(2).Player(4), Pos8.Map, Pos8.X, Pos8.Y, False)
        Call WriteParalizeOK(DuelData.Duelo(Slot).Team(1).Player(1), False)
        Call WriteParalizeOK(DuelData.Duelo(Slot).Team(2).Player(1), False)
        Call WriteParalizeOK(DuelData.Duelo(Slot).Team(1).Player(2), False)
        Call WriteParalizeOK(DuelData.Duelo(Slot).Team(2).Player(2), False)
        Call WriteParalizeOK(DuelData.Duelo(Slot).Team(1).Player(3), False)
        Call WriteParalizeOK(DuelData.Duelo(Slot).Team(2).Player(3), False)
        Call WriteParalizeOK(DuelData.Duelo(Slot).Team(1).Player(4), False)
        Call WriteParalizeOK(DuelData.Duelo(Slot).Team(2).Player(4), False)
    End If
    Exit Sub

ErrHandler:
        Call LogError("Error en Sub WarpArena. Slot: " & Slot & " " & Err.Description)
End Sub

Sub PeticionDuelo(ByVal UserIndex As Integer, ByVal vs As Byte, ByVal Oro As Long, ByVal Drop As Boolean, _
                    ByVal Nick1 As String, Optional ByVal Nick2 As String, Optional ByVal Nick3 As String, _
                    Optional ByVal Resucitar As Boolean = False, Optional ByVal Nick4 As String, _
                    Optional ByVal Nick5 As String, Optional ByVal Nick6 As String, Optional ByVal Nick7 As String)
On Error GoTo ErrHandler

    If Not PuedeDuelo(UserIndex, Oro) Then Exit Sub
    Dim n1 As Integer
    n1 = NameIndex(Nick1)
    
    If Not vs = 1 Then
        Dim n2 As Integer
        Dim n3 As Integer
        n2 = NameIndex(Nick2)
        n3 = NameIndex(Nick3)
        If Not vs = 2 Then
            Dim n4 As Integer
            Dim n5 As Integer
            n4 = NameIndex(Nick4)
            n5 = NameIndex(Nick5)
            If Not vs = 3 Then
                Dim n6 As Integer
                Dim n7 As Integer
                n6 = NameIndex(Nick6)
                n7 = NameIndex(Nick7)
            End If
        End If
    End If
    
    If Not Drop Then
        If Not Oro >= DuelData.MinGold Then
            Call WriteConsoleMsg(UserIndex, "La apuesta mínima para duelos sin drop es de " & DuelData.MinGold & " Monedas de Oro.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    End If
    
    If Not Oro >= 0 Then
        Call WriteConsoleMsg(UserIndex, "¡No puedes apostar cantidades negativas!", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If Not NameIndex(Nick1) > 0 Then
        Call WriteConsoleMsg(UserIndex, "El usuario " & Nick1 & " no se encuentra.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If Not vs = 1 Then
        If Not NameIndex(Nick2) > 0 Then
            Call WriteConsoleMsg(UserIndex, "El usuario " & Nick2 & " no se encuentra.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If Not NameIndex(Nick3) > 0 Then
            Call WriteConsoleMsg(UserIndex, "El usuario " & Nick3 & " no se encuentra.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If Not vs = 2 Then
            If Not NameIndex(Nick4) > 0 Then
                Call WriteConsoleMsg(UserIndex, "El usuario " & Nick4 & " no se encuentra.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            If Not NameIndex(Nick5) > 0 Then
                Call WriteConsoleMsg(UserIndex, "El usuario " & Nick5 & " no se encuentra.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            If Not vs = 3 Then
                If Not NameIndex(Nick6) > 0 Then
                    Call WriteConsoleMsg(UserIndex, "El usuario " & Nick6 & " no se encuentra.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                If Not NameIndex(Nick7) > 0 Then
                    Call WriteConsoleMsg(UserIndex, "El usuario " & Nick7 & " no se encuentra.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            End If
        End If
    End If
    
    
    If UserIndex = n1 Then
        Call WriteConsoleMsg(UserIndex, "No puedes anotarte a ti mismo.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If Not vs = 1 Then
        If UserIndex = n2 Or UserIndex = n3 Then
            Call WriteConsoleMsg(UserIndex, "No puedes anotarte a ti mismo.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If n1 = n2 Or n1 = n3 Or n2 = n3 Then
            Call WriteConsoleMsg(UserIndex, "No puedes invitar varias veces a la misma persona.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If Not vs = 2 Then
            If UserIndex = n4 Or UserIndex = n5 Then
                Call WriteConsoleMsg(UserIndex, "No puedes anotarte a ti mismo.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            If n1 = n4 Or n1 = n5 Or n2 = n4 Or n2 = n5 Or n3 = n4 Or n3 = n5 Or n4 = n5 Then
                Call WriteConsoleMsg(UserIndex, "No puedes invitar varias veces a la misma persona.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            If Not vs = 3 Then
                If UserIndex = n6 Or UserIndex = n7 Then
                    Call WriteConsoleMsg(UserIndex, "No puedes anotarte a ti mismo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                If n1 = n6 Or n1 = n7 Or n2 = n6 Or n2 = n7 Or n3 = n6 Or n3 = n5 Or n4 = n6 Or n4 = n7 Or n5 = n6 Or n5 = n7 Or n6 = n7 Then
                    Call WriteConsoleMsg(UserIndex, "No puedes invitar varias veces a la misma persona.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            End If
        End If
    End If
    
    If UserList(UserIndex).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then
        Call WriteConsoleMsg(UserIndex, "Eeeh GM no te hagas el loco con los pibes.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    If n1 > 0 Then
        If UserList(n1).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then
            Call WriteConsoleMsg(UserIndex, "No puedes retar a duelo a un GameMaster.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    End If
    If n2 > 0 Then
        If UserList(n2).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then
            Call WriteConsoleMsg(UserIndex, "No puedes retar a duelo a un GameMaster.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    End If
    If n3 > 0 Then
        If UserList(n3).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then
            Call WriteConsoleMsg(UserIndex, "No puedes retar a duelo a un GameMaster.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    End If
    If n4 > 0 Then
        If UserList(n4).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then
            Call WriteConsoleMsg(UserIndex, "No puedes retar a duelo a un GameMaster.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    End If
    If n5 > 0 Then
        If UserList(n5).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then
            Call WriteConsoleMsg(UserIndex, "No puedes retar a duelo a un GameMaster.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    End If
    If n6 > 0 Then
        If UserList(n6).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then
            Call WriteConsoleMsg(UserIndex, "No puedes retar a duelo a un GameMaster.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    End If
    If n7 > 0 Then
        If UserList(n7).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then
            Call WriteConsoleMsg(UserIndex, "No puedes retar a duelo a un GameMaster.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    End If
    If Not PuedeParticiparDuelo(UserIndex, n1, Oro) Then Exit Sub
    
    If Not vs = 1 Then
        If Not PuedeParticiparDuelo(UserIndex, n2, Oro) Or Not PuedeParticiparDuelo(UserIndex, n3, Oro) Then Exit Sub
        If Not vs = 2 Then
            If Not PuedeParticiparDuelo(UserIndex, n4, Oro) Or Not PuedeParticiparDuelo(UserIndex, n5, Oro) Then Exit Sub
            If Not vs = 3 Then
                If Not PuedeParticiparDuelo(UserIndex, n6, Oro) Or Not PuedeParticiparDuelo(UserIndex, n7, Oro) Then Exit Sub
            End If
        End If
    End If
        
    Dim MySlot As Byte
    If vs = 1 Then
        MySlot = GetNewDueloSlot(eDuelType.vs1)
        If Not MySlot > 0 Then
            Call WriteConsoleMsg(UserIndex, "No se ha podido enviar las invitación, todas las arenas para duelos 1vs1 están ocupadas.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Call GiveDueloSlot(eDuelType.vs1, MySlot, UserIndex)
        Call SetDueloConfig(MySlot, eDuelType.vs1, Drop, Resucitar, Oro)
    ElseIf vs = 2 Then
        MySlot = GetNewDueloSlot(eDuelType.vs2)
        If Not MySlot > 0 Then
            Call WriteConsoleMsg(UserIndex, "No se ha podido enviar las invitación, todas las arenas para duelos 2vs2 están ocupadas.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Call GiveDueloSlot(eDuelType.vs2, MySlot, UserIndex)
        Call SetDueloConfig(MySlot, eDuelType.vs2, Drop, Resucitar, Oro)
    ElseIf vs = 3 Then
        MySlot = GetNewDueloSlot(eDuelType.vs3)
        If Not MySlot > 0 Then
            Call WriteConsoleMsg(UserIndex, "No se ha podido enviar las invitación, todas las arenas para duelos 3vs3 están ocupadas.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Call GiveDueloSlot(eDuelType.vs3, MySlot, UserIndex)
        Call SetDueloConfig(MySlot, eDuelType.vs3, Drop, Resucitar, Oro)
    ElseIf vs = 4 Then
        MySlot = GetNewDueloSlot(eDuelType.vs4)
        If Not MySlot > 0 Then
            Call WriteConsoleMsg(UserIndex, "No se ha podido enviar las invitación, todas las arenas para duelos 4vs4 están ocupadas.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Call GiveDueloSlot(eDuelType.vs4, MySlot, UserIndex)
        Call SetDueloConfig(MySlot, eDuelType.vs4, Drop, Resucitar, Oro)
    End If
    

    If vs = 1 Then
        Call WriteMensajeDuelo(n1, MySlot, False, UserList(UserIndex).Name)
        
        UserList(n1).flags.DueloIndex = MySlot
        UserList(n1).flags.DueloTeam = 2
        Call WriteConsoleMsg(UserIndex, "La invitación ha sido enviada.", FontTypeNames.FONTTYPE_INFO)
    ElseIf vs = 2 Then
        Call WriteMensajeDuelo(n1, MySlot, True, UserList(UserIndex).Name, UserList(n2).Name, UserList(n3).Name)
        Call WriteMensajeDuelo(n2, MySlot, False, UserList(n3).Name, UserList(UserIndex).Name, UserList(n1).Name)
        Call WriteMensajeDuelo(n3, MySlot, False, UserList(n2).Name, UserList(UserIndex).Name, UserList(n1).Name)
        
        UserList(n1).flags.DueloIndex = MySlot
        UserList(n2).flags.DueloIndex = MySlot
        UserList(n3).flags.DueloIndex = MySlot
        UserList(n1).flags.DueloTeam = 1
        UserList(n2).flags.DueloTeam = 2
        UserList(n3).flags.DueloTeam = 2
        Call WriteConsoleMsg(UserIndex, "Las invitaciones han sido enviadas.", FontTypeNames.FONTTYPE_INFO)
    ElseIf vs = 3 Then
        Call WriteMensajeDuelo(n1, MySlot, True, UserList(UserIndex).Name, UserList(n2).Name, UserList(n3).Name, UserList(n4).Name, UserList(n5).Name)
        Call WriteMensajeDuelo(n2, MySlot, True, UserList(UserIndex).Name, UserList(n1).Name, UserList(n3).Name, UserList(n4).Name, UserList(n5).Name)
        Call WriteMensajeDuelo(n3, MySlot, False, UserList(n4).Name, UserList(n5).Name, UserList(UserIndex).Name, UserList(n1).Name, UserList(n2).Name)
        Call WriteMensajeDuelo(n4, MySlot, False, UserList(n3).Name, UserList(n5).Name, UserList(UserIndex).Name, UserList(n1).Name, UserList(n2).Name)
        Call WriteMensajeDuelo(n5, MySlot, False, UserList(n3).Name, UserList(n4).Name, UserList(UserIndex).Name, UserList(n1).Name, UserList(n2).Name)

        UserList(n1).flags.DueloIndex = MySlot
        UserList(n2).flags.DueloIndex = MySlot
        UserList(n3).flags.DueloIndex = MySlot
        UserList(n4).flags.DueloIndex = MySlot
        UserList(n5).flags.DueloIndex = MySlot
        UserList(n1).flags.DueloTeam = 1
        UserList(n2).flags.DueloTeam = 1
        UserList(n3).flags.DueloTeam = 2
        UserList(n4).flags.DueloTeam = 2
        UserList(n5).flags.DueloTeam = 2
        Call WriteConsoleMsg(UserIndex, "Las invitaciones han sido enviadas.", FontTypeNames.FONTTYPE_INFO)
    ElseIf vs = 4 Then
        Call WriteMensajeDuelo(n1, MySlot, True, UserList(UserIndex).Name, UserList(n2).Name, UserList(n3).Name, UserList(n4).Name, UserList(n5).Name, UserList(n6).Name, UserList(n7).Name)
        Call WriteMensajeDuelo(n2, MySlot, True, UserList(UserIndex).Name, UserList(n1).Name, UserList(n3).Name, UserList(n4).Name, UserList(n5).Name, UserList(n6).Name, UserList(n7).Name)
        Call WriteMensajeDuelo(n3, MySlot, True, UserList(UserIndex).Name, UserList(n1).Name, UserList(n2).Name, UserList(n4).Name, UserList(n5).Name, UserList(n6).Name, UserList(n7).Name)
        Call WriteMensajeDuelo(n4, MySlot, False, UserList(n5).Name, UserList(n6).Name, UserList(n7).Name, UserList(UserIndex).Name, UserList(n1).Name, UserList(n2).Name, UserList(n3).Name)
        Call WriteMensajeDuelo(n5, MySlot, False, UserList(n4).Name, UserList(n6).Name, UserList(n7).Name, UserList(UserIndex).Name, UserList(n1).Name, UserList(n2).Name, UserList(n3).Name)
        Call WriteMensajeDuelo(n6, MySlot, False, UserList(n4).Name, UserList(n5).Name, UserList(n7).Name, UserList(UserIndex).Name, UserList(n1).Name, UserList(n2).Name, UserList(n3).Name)
        Call WriteMensajeDuelo(n7, MySlot, False, UserList(n4).Name, UserList(n5).Name, UserList(n6).Name, UserList(UserIndex).Name, UserList(n1).Name, UserList(n2).Name, UserList(n3).Name)

        UserList(n1).flags.DueloIndex = MySlot
        UserList(n2).flags.DueloIndex = MySlot
        UserList(n3).flags.DueloIndex = MySlot
        UserList(n4).flags.DueloIndex = MySlot
        UserList(n5).flags.DueloIndex = MySlot
        UserList(n6).flags.DueloIndex = MySlot
        UserList(n7).flags.DueloIndex = MySlot
        UserList(n1).flags.DueloTeam = 1
        UserList(n2).flags.DueloTeam = 1
        UserList(n3).flags.DueloTeam = 1
        UserList(n4).flags.DueloTeam = 2
        UserList(n5).flags.DueloTeam = 2
        UserList(n6).flags.DueloTeam = 2
        UserList(n7).flags.DueloTeam = 2
        Call WriteConsoleMsg(UserIndex, "Las invitaciones han sido enviadas.", FontTypeNames.FONTTYPE_INFO)
    End If
Exit Sub

ErrHandler:
        Call LogError("Error en Sub PeticionDuelo. " & Err.Description)
End Sub

Sub ApresurarFinalDuelo(ByVal Slot As Byte)
On Error GoTo ErrHandler
  
    If Not Slot > 0 Then Exit Sub
    DuelData.Duelo(Slot).Counter = 0
    Call CerrarDuelo(Slot)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ApresurarFinalDuelo de modDuelos.bas")
End Sub

Public Function DueloIniciado(ByVal Slot As Byte) As Boolean
On Error GoTo ErrHandler
  
    If Not Slot > 0 Then
        DueloIniciado = False
        Exit Function
    End If
    
    DueloIniciado = (DuelData.Duelo(Slot).estado = eDuelState.Iniciado) Or (DuelData.Duelo(Slot).estado = eDuelState.Esperando_Final) Or (DuelData.Duelo(Slot).estado = eDuelState.Esperando_Inicio)
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function DueloIniciado de modDuelos.bas")
End Function

Public Function GetDueloFontColor(ByVal UserIndex As Integer) As FontTypeNames
On Error GoTo ErrHandler
  
    If UserList(UserIndex).flags.DueloIndex > 0 Then
        If UserList(UserIndex).flags.DueloTeam = 1 Then
            GetDueloFontColor = FontTypeNames.FONTTYPE_INFO
            Exit Function
        Else
            GetDueloFontColor = FontTypeNames.FONTTYPE_INFO
            Exit Function
        End If
    End If
    
    GetDueloFontColor = FontTypeNames.FONTTYPE_INFO
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetDueloFontColor de modDuelos.bas")
End Function

Sub CerrarTodosLosDuelos()
On Error GoTo ErrHandler

    Dim I As Byte
    For I = LBound(DuelData.Duelo) To UBound(DuelData.Duelo)
        Select Case DuelData.Duelo(I).estado
            Case eDuelState.Esperando_Inicio, eDuelState.Iniciado, eDuelState.Esperando_Final
                Call CerrarDuelo(I)
            Case eDuelState.Esperando_Jugadores
                Call CancelarDuelo(I)
        End Select
    Next I
    
    Exit Sub

ErrHandler:
    Call LogError("Error " & Err.Number & " (" & Err.Description & ") en Sub CerrarTodosLosDuelos del Módulo modDuelos")
End Sub

Public Sub WarpReturnDuelo(ByVal UserIndex As Integer)
On Error GoTo ErrHandler
  
    Dim MiPos As WorldPos
    
    MiPos = UserList(UserIndex).VolverDueloPos
    
    Call WarpUserChar(UserIndex, MiPos.Map, MiPos.X, MiPos.Y, True)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WarpReturnDuelo de modDuelos.bas")
End Sub

Public Function NextOpenArena(ByVal TipoDuelo As Byte) As Byte
On Error GoTo ErrHandler
    Dim I As Byte
    
    Select Case TipoDuelo
        Case eDuelType.vs1
            For I = 1 To UBound(Arena1v1)
                If Arena1v1(I).EnUso = 0 Then
                    NextOpenArena = I
                    Exit Function
                End If
            Next I
        Case eDuelType.vs2
            For I = 1 To UBound(Arena2v2)
                If Arena2v2(I).EnUso = 0 Then
                    NextOpenArena = I
                    Exit Function
                End If
            Next I
        Case eDuelType.vs3
            For I = 1 To UBound(Arena3v3)
                If Arena3v3(I).EnUso = 0 Then
                    NextOpenArena = I
                    Exit Function
                End If
            Next I
        Case eDuelType.vs4
            For I = 1 To UBound(Arena4v4)
                If Arena4v4(I).EnUso = 0 Then
                    NextOpenArena = I
                    Exit Function
                End If
            Next I
    End Select
    
    Exit Function
ErrHandler:
    Call LogError("Error " & Err.Number & " (" & Err.Description & ") en Function NextOpenArena del Módulo modDuelos")

End Function
