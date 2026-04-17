Attribute VB_Name = "Extra"
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

Public Function EsNewbie(ByVal UserIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    EsNewbie = UserList(UserIndex).Stats.ELV <= ConstantesBalance.LimiteNewbie
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function EsNewbie de GameLogic.bas")
End Function

Public Function esArmada(ByVal UserIndex As Integer) As Boolean
'***************************************************
'Autor: Pablo (ToxicWaste)
'Last Modification: 23/01/2007
'***************************************************
On Error GoTo ErrHandler
  

    esArmada = (UserList(UserIndex).Faccion.ArmadaReal = 1)
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function esArmada de GameLogic.bas")
End Function

Public Function esCaos(ByVal UserIndex As Integer) As Boolean
'***************************************************
'Autor: Pablo (ToxicWaste)
'Last Modification: 23/01/2007
'***************************************************
On Error GoTo ErrHandler
  

    esCaos = (UserList(UserIndex).Faccion.FuerzasCaos = 1)
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function esCaos de GameLogic.bas")
End Function

Public Function EsGm(ByVal UserIndex As Integer) As Boolean
'***************************************************
'Autor: Pablo (ToxicWaste)
'Last Modification: 23/01/2007
'***************************************************
On Error GoTo ErrHandler
  

    EsGm = (UserList(UserIndex).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero))
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function EsGm de GameLogic.bas")
End Function

Public Sub DoTileEvents(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
'***************************************************
'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
'Last Modification: 06/03/2010
'Handles the Map passage of Users. Allows the existance
'of exclusive maps for Newbies, Royal Army and Caos Legion members
'and enables GMs to enter every map without restriction.
'Uses: Mapinfo(map).Restringir = "NEWBIE" (newbies), "ARMADA", "CAOS", "FACCION" or "NO".
' 06/03/2010 : Now we have 5 attemps to not fall into a map change or another teleport while going into a teleport. (Marco)
'***************************************************

    Dim nPos As WorldPos
    Dim FxFlag As Boolean
    Dim TelepRadio As Integer
    Dim DestPos As WorldPos
    
On Error GoTo ErrHandler
    'Controla las salidas
    If InMapBounds(Map, X, Y) Then
        With MapData(Map, X, Y)
            If .ObjInfo.ObjIndex > 0 Then
                FxFlag = ObjData(.ObjInfo.ObjIndex).ObjType = eOBJType.otTeleport
                TelepRadio = ObjData(.ObjInfo.ObjIndex).Radio
            End If
            
            If .TileExit.Map > 0 And .TileExit.Map <= NumMaps Then
                
                ' Es un teleport, entra en una posicion random, acorde al radio (si es 0, es pos fija)
                ' We have 5 attempts to not falling into another teleport or a map exit.. If we get to the fifth attemp,
                ' the teleport will act as if its radius = 0.
                If FxFlag And TelepRadio > 0 Then
                    Dim attemps As Long
                    Dim exitMap As Boolean
                    Do
                        DestPos.X = .TileExit.X + RandomNumber(TelepRadio * (-1), TelepRadio)
                        DestPos.Y = .TileExit.Y + RandomNumber(TelepRadio * (-1), TelepRadio)
                        
                        attemps = attemps + 1
                        
                        exitMap = MapData(.TileExit.Map, DestPos.X, DestPos.Y).TileExit.Map > 0 And _
                                MapData(.TileExit.Map, DestPos.X, DestPos.Y).TileExit.Map <= NumMaps
                    Loop Until (attemps >= 5 Or exitMap = False)
                    
                    If attemps >= 5 Then
                        DestPos.X = .TileExit.X
                        DestPos.Y = .TileExit.Y
                    End If
                ' Posicion fija
                Else
                    DestPos.X = .TileExit.X
                    DestPos.Y = .TileExit.Y
                End If
                
                DestPos.Map = .TileExit.Map
                
                If EsGm(UserIndex) Then
                    Call LogGM(UserList(UserIndex).Name, "Utilizó un teleport hacia el mapa " & _
                        DestPos.Map & " (" & DestPos.X & "," & DestPos.Y & ")")
                End If
                
                ' Si es un mapa que no admite muertos
                If MapInfo(DestPos.Map).OnDeathGoTo.Map <> 0 Then
                    ' Si esta muerto no puede entrar
                    If UserList(UserIndex).flags.Muerto = 1 Then
                        Call WriteConsoleMsg(UserIndex, "Sólo se permite entrar al mapa a los personajes vivos.", FontTypeNames.FONTTYPE_INFO)
                        Call ClosestStablePos(UserList(UserIndex).Pos, nPos)
                        
                        If nPos.X <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, FxFlag)
                        End If
                        
                        Exit Sub
                    End If
                End If
                
                
                '¿Es mapa de newbies?
                If MapInfo(DestPos.Map).Restringir = eRestrict.restrict_newbie Then
                    '¿El usuario es un newbie?
                    If EsNewbie(UserIndex) Or EsGm(UserIndex) Then
                        If LegalPos(DestPos.Map, DestPos.X, DestPos.Y, PuedeAtravesarAgua(UserIndex)) Then
                            Call WarpUserChar(UserIndex, DestPos.Map, DestPos.X, DestPos.Y, FxFlag)
                        Else
                            Call ClosestLegalPos(DestPos, nPos, , , True)
                            If nPos.X <> 0 And nPos.Y <> 0 Then
                                Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, FxFlag)
                            End If
                        End If
                    Else 'No es newbie
                        Call WriteConsoleMsg(UserIndex, "Mapa exclusivo para newbies.", FontTypeNames.FONTTYPE_INFO)
                        Call ClosestStablePos(UserList(UserIndex).Pos, nPos)
        
                        If nPos.X <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, False)
                        End If
                    End If
                ElseIf MapInfo(DestPos.Map).Restringir = eRestrict.restrict_armada Then '¿Es mapa de Armadas?
                    '¿El usuario es Armada?
                    If esArmada(UserIndex) Or EsGm(UserIndex) Then
                        If LegalPos(DestPos.Map, DestPos.X, DestPos.Y, PuedeAtravesarAgua(UserIndex)) Then
                            Call WarpUserChar(UserIndex, DestPos.Map, DestPos.X, DestPos.Y, FxFlag)
                        Else
                            Call ClosestLegalPos(DestPos, nPos, , , True)
                            If nPos.X <> 0 And nPos.Y <> 0 Then
                                Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, FxFlag)
                            End If
                        End If
                    Else 'No es armada
                        Call WriteConsoleMsg(UserIndex, "Mapa exclusivo para miembros del ejército real.", FontTypeNames.FONTTYPE_INFO)
                        Call ClosestStablePos(UserList(UserIndex).Pos, nPos)
                        
                        If nPos.X <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, FxFlag)
                        End If
                    End If
                ElseIf MapInfo(DestPos.Map).Restringir = eRestrict.restrict_caos Then '¿Es mapa de Caos?
                    '¿El usuario es Caos?
                    If esCaos(UserIndex) Or EsGm(UserIndex) Then
                        If LegalPos(DestPos.Map, DestPos.X, DestPos.Y, PuedeAtravesarAgua(UserIndex)) Then
                            Call WarpUserChar(UserIndex, DestPos.Map, DestPos.X, DestPos.Y, FxFlag)
                        Else
                            Call ClosestLegalPos(DestPos, nPos, , , True)
                            If nPos.X <> 0 And nPos.Y <> 0 Then
                                Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, FxFlag)
                            End If
                        End If
                    Else 'No es caos
                        Call WriteConsoleMsg(UserIndex, "Mapa exclusivo para miembros de la legión oscura.", FontTypeNames.FONTTYPE_INFO)
                        Call ClosestStablePos(UserList(UserIndex).Pos, nPos)
                        
                        If nPos.X <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, FxFlag)
                        End If
                    End If
                ElseIf MapInfo(DestPos.Map).Restringir = eRestrict.restrict_faccion Then '¿Es mapa de faccionarios?
                    '¿El usuario es Armada o Caos?
                    If esArmada(UserIndex) Or esCaos(UserIndex) Or EsGm(UserIndex) Then
                        If LegalPos(DestPos.Map, DestPos.X, DestPos.Y, PuedeAtravesarAgua(UserIndex)) Then
                            Call WarpUserChar(UserIndex, DestPos.Map, DestPos.X, DestPos.Y, FxFlag)
                        Else
                            Call ClosestLegalPos(DestPos, nPos, , , True)
                            If nPos.X <> 0 And nPos.Y <> 0 Then
                                Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, FxFlag)
                            End If
                        End If
                    Else 'No es Faccionario
                        Call WriteConsoleMsg(UserIndex, "Solo se permite entrar al mapa si eres miembro de alguna facción.", FontTypeNames.FONTTYPE_INFO)
                        Call ClosestStablePos(UserList(UserIndex).Pos, nPos)
                        
                        If nPos.X <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, FxFlag)
                        End If
                    End If
                Else 'No es un mapa de newbies, ni Armadas, ni Caos, ni faccionario.
                    If LegalPos(DestPos.Map, DestPos.X, DestPos.Y, PuedeAtravesarAgua(UserIndex)) Then
                        Call WarpUserChar(UserIndex, DestPos.Map, DestPos.X, DestPos.Y, FxFlag)
                    Else
                        Call ClosestLegalPos(DestPos, nPos, , , True)
                        If nPos.X <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, FxFlag)
                        End If
                    End If
                End If

                'Te fusite del mapa. La criatura ya no es más tuya ni te reconoce como que vos la atacaste.
                Dim aN As Integer
                
                aN = UserList(UserIndex).flags.AtacadoPorNpc
                If aN > 0 Then
                    If Npclist(aN).MaestroUser > 0 Then
                        Call AllFollowAmo(Npclist(aN).MaestroUser)
                    Else
                        Npclist(aN).Movement = Npclist(aN).flags.OldMovement
                        Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
                        Npclist(aN).flags.AttackedBy = vbNullString
                    End If
                    
                    
                End If
            
                aN = UserList(UserIndex).flags.NPCAtacado
                If aN > 0 Then
                    If Npclist(aN).flags.AttackedFirstBy = UserList(UserIndex).Name Then
                        Npclist(aN).flags.AttackedFirstBy = vbNullString
                    End If
                End If
                
                UserList(UserIndex).flags.AtacadoPorNpc = 0
                UserList(UserIndex).flags.NPCAtacado = 0
            End If
        End With
    End If
Exit Sub

ErrHandler:
    Call LogError("Error en DotileEvents. Error: " & Err.Number & " - Desc: " & Err.Description)
End Sub

Function InRangoVision(ByVal UserIndex As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    If X > UserList(UserIndex).Pos.X - MinXBorder And X < UserList(UserIndex).Pos.X + MinXBorder Then
        If Y > UserList(UserIndex).Pos.Y - MinYBorder And Y < UserList(UserIndex).Pos.Y + MinYBorder Then
            InRangoVision = True
            Exit Function
        End If
    End If
    InRangoVision = False

  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function InRangoVision de GameLogic.bas")
End Function

Public Function InVisionRangeAndMap(ByVal UserIndex As Integer, ByRef OtherUserPos As WorldPos) As Boolean
'***************************************************
'Author: ZaMa
'Last Modification: 20/11/2010
'
'***************************************************
On Error GoTo ErrHandler
  
    
    With UserList(UserIndex)
        
        ' Same map?
        If .Pos.Map <> OtherUserPos.Map Then Exit Function
    
        ' In x range?
        If OtherUserPos.X < .Pos.X - MinXBorder Or OtherUserPos.X > .Pos.X + MinXBorder Then Exit Function
        
        ' In y range?
        If OtherUserPos.Y < .Pos.Y - MinYBorder Or OtherUserPos.Y > .Pos.Y + MinYBorder Then Exit Function
    End With

    InVisionRangeAndMap = True
    
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function InVisionRangeAndMap de GameLogic.bas")
End Function

Function InMapBounds(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    If (Map <= 0 Or Map > NumMaps) Or X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
        InMapBounds = False
    Else
        InMapBounds = True
    End If
    
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function InMapBounds de GameLogic.bas")
    End Function


Function RhombLegalPos(ByVal Map As Integer, ByRef X As Integer, ByRef Y As Integer, _
                               ByVal Distance As Long, Optional PuedeAgua As Boolean = False, _
                               Optional PuedeTierra As Boolean = True, _
                               Optional ByVal CheckExitTile As Boolean = False, Optional ByVal UpdatePos As Boolean = True) As Boolean
'***************************************************
'Author: Marco Vanotti (Marco)
'Last Modification: -
' walks all the perimeter of a rhomb of side  "distance + 1",
' which starts at Pos.x - Distance and Pos.y
'***************************************************
On Error GoTo ErrHandler
  

    Dim I As Long
    Dim vX As Long
    Dim vY As Long
    
    vX = X - Distance
    vY = Y
    
    For I = 0 To Distance - 1
        If (LegalPos(Map, vX + I, vY - I, PuedeAgua, PuedeTierra, CheckExitTile)) Then
            If UpdatePos Then
                X = vX + I
                Y = vY - I
            End If
            RhombLegalPos = True
            Exit Function
        End If
    Next
    
    vX = X
    vY = Y - Distance
    
    For I = 0 To Distance - 1
        If (LegalPos(Map, vX + I, vY + I, PuedeAgua, PuedeTierra, CheckExitTile)) Then
            If UpdatePos Then
                X = vX + I
                Y = vY + I
            End If
            RhombLegalPos = True
            Exit Function
        End If
    Next
    
    vX = X + Distance
    vY = Y
    
    For I = 0 To Distance - 1
        If (LegalPos(Map, vX - I, vY + I, PuedeAgua, PuedeTierra, CheckExitTile)) Then
            If UpdatePos Then
                X = vX - I
                Y = vY + I
            End If
            RhombLegalPos = True
            Exit Function
        End If
    Next
    
    vX = X
    vY = Y + Distance
    
    For I = 0 To Distance - 1
        If (LegalPos(Map, vX - I, vY - I, PuedeAgua, PuedeTierra, CheckExitTile)) Then
            If UpdatePos Then
                X = vX - I
                Y = vY - I
            End If
            RhombLegalPos = True
            Exit Function
        End If
    Next
    
    RhombLegalPos = False
    
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function RhombLegalPos de GameLogic.bas")
End Function

Public Function RhombLegalTilePos(ByRef Pos As WorldPos, ByRef vX As Long, ByRef vY As Long, _
                                  ByVal Distance As Long, ByVal ObjIndex As Integer, ByVal ObjAmount As Long, _
                                  ByVal PuedeAgua As Boolean, ByVal PuedeTierra As Boolean) As Boolean
'***************************************************
'Author: ZaMa
'Last Modification: -
' walks all the perimeter of a rhomb of side  "distance + 1",
' which starts at Pos.x - Distance and Pos.y
' and searchs for a valid position to drop items
'***************************************************
On Error GoTo ErrHandler

    Dim I As Long
    
    Dim X As Integer
    Dim Y As Integer
    
    vX = Pos.X - Distance
    vY = Pos.Y
    
    For I = 0 To Distance - 1
        
        X = vX + I
        Y = vY - I
        
        If (LegalPos(Pos.Map, X, Y, PuedeAgua, PuedeTierra, True)) Then
            
            ' No hay obj tirado o la suma de lo que hay + lo nuevo <= 10k
            If Not HayObjeto(Pos.Map, X, Y, ObjIndex, ObjAmount) Then
                vX = X
                vY = Y
                
                RhombLegalTilePos = True
                Exit Function
            End If
            
        End If
    Next
    
    vX = Pos.X
    vY = Pos.Y - Distance
    
    For I = 0 To Distance - 1
        
        X = vX + I
        Y = vY + I
        
        If (LegalPos(Pos.Map, X, Y, PuedeAgua, PuedeTierra, True)) Then
            
            ' No hay obj tirado o la suma de lo que hay + lo nuevo <= 10k
            If Not HayObjeto(Pos.Map, X, Y, ObjIndex, ObjAmount) Then
                vX = X
                vY = Y
                
                RhombLegalTilePos = True
                Exit Function
            End If
        End If
    Next
    
    vX = Pos.X + Distance
    vY = Pos.Y
    
    For I = 0 To Distance - 1
        
        X = vX - I
        Y = vY + I
    
        If (LegalPos(Pos.Map, X, Y, PuedeAgua, PuedeTierra, True)) Then
        
            ' No hay obj tirado o la suma de lo que hay + lo nuevo <= 10k
            If Not HayObjeto(Pos.Map, X, Y, ObjIndex, ObjAmount) Then
                vX = X
                vY = Y
                
                RhombLegalTilePos = True
                Exit Function
            End If
        End If
    Next
    
    vX = Pos.X
    vY = Pos.Y + Distance
    
    For I = 0 To Distance - 1
        
        X = vX - I
        Y = vY - I
    
        If (LegalPos(Pos.Map, X, Y, PuedeAgua, PuedeTierra, True)) Then
            ' No hay obj tirado o la suma de lo que hay + lo nuevo <= 10k
            If Not HayObjeto(Pos.Map, X, Y, ObjIndex, ObjAmount) Then
                vX = X
                vY = Y
                
                RhombLegalTilePos = True
                Exit Function
            End If
        End If
    Next
    
    RhombLegalTilePos = False
    
    Exit Function
    
ErrHandler:
    Call LogError("Error en RhombLegalTilePos. Error: " & Err.Number & " - " & Err.Description)
End Function

Public Function RhombObjExists(ByRef Pos As WorldPos, ByRef vX As Long, ByRef vY As Long, _
                                  ByVal Distance As Long, ByVal ObjIndex As Integer) As Boolean
'***************************************************
'Author: ZaMa
'Last Modification: -
' walks all the perimeter of a rhomb of side  "distance + 1",
' which starts at Pos.x - Distance and Pos.y
' and searchs for a valid position to drop items
'***************************************************
On Error GoTo ErrHandler

    Dim I As Long
    
    Dim X As Integer
    Dim Y As Integer
    
    vX = Pos.X - Distance
    vY = Pos.Y
    
    
    X = vX
    Y = vY
        
    For I = 0 To Distance
        ' No hay obj tirado o la suma de lo que hay + lo nuevo <= 10k
        If ObjectExistsInTile(Pos.Map, X, Y, ObjIndex) Then
            vX = X
            vY = Y
            
            RhombObjExists = True
            Exit Function
        End If
        
        X = vX + I
        Y = vY - I
    Next
    
    vX = Pos.X
    vY = Pos.Y - Distance
    
    X = vX
    Y = vY
    
    For I = 0 To Distance

        ' No hay obj tirado o la suma de lo que hay + lo nuevo <= 10k
        If ObjectExistsInTile(Pos.Map, X, Y, ObjIndex) Then
            vX = X
            vY = Y
            
            RhombObjExists = True
            Exit Function
        End If
        
        X = vX + I
        Y = vY + I
    Next
    
    vX = Pos.X + Distance
    vY = Pos.Y
    
    X = vX
    Y = vY
    
    For I = 0 To Distance
        ' No hay obj tirado o la suma de lo que hay + lo nuevo <= 10k
        If ObjectExistsInTile(Pos.Map, X, Y, ObjIndex) Then
            vX = X
            vY = Y
            
            RhombObjExists = True
            Exit Function
        End If
        
        X = vX - I
        Y = vY + I
    Next
    
    vX = Pos.X
    vY = Pos.Y + Distance
    
    X = vX
    Y = vY
    
    For I = 0 To Distance - 1
            
        ' No hay obj tirado o la suma de lo que hay + lo nuevo <= 10k
        If ObjectExistsInTile(Pos.Map, X, Y, ObjIndex) Then
            vX = X
            vY = Y
            
            RhombObjExists = True
            Exit Function
        End If
        
        X = vX - I
        Y = vY - I
    Next
    
    RhombObjExists = False
    
    Exit Function
    
ErrHandler:
    Call LogError("Error en RhombObjExists. Error: " & Err.Number & " - " & Err.Description)
End Function

Public Function HayObjeto(ByVal mapa As Integer, ByVal X As Long, ByVal Y As Long, _
                          ByVal ObjIndex As Integer, ByVal ObjAmount As Long) As Boolean
'***************************************************
'Author: ZaMa
'Last Modification: -
'Checks if there's space in a tile to add an itemAmount
'***************************************************
On Error GoTo ErrHandler
  
    Dim MapObjIndex As Integer
    
    MapObjIndex = MapData(mapa, X, Y).ObjInfo.ObjIndex
            
    ' Hay un objeto tirado?
    If MapObjIndex <> 0 Then
        ' Es el mismo objeto?
        If MapObjIndex = ObjIndex Then
            ' La suma es menor a 10k?
            HayObjeto = (MapData(mapa, X, Y).ObjInfo.Amount + ObjAmount > MAX_INVENTORY_OBJS)
        Else
            HayObjeto = True
        End If
    Else
        HayObjeto = False
    End If

  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function HayObjeto de GameLogic.bas")
End Function


Public Function ObjectExistsInTile(ByVal mapa As Integer, ByVal X As Long, ByVal Y As Long, _
                          ByVal ObjIndex As Integer) As Boolean
'***************************************************
'Author: Nightw
'Last Modification: -
'Checks if a given object exists in a given tile.
'***************************************************
On Error GoTo ErrHandler
  
    Dim MapObjIndex As Integer
    Dim ObjExists As Boolean
    ObjExists = False
    
    MapObjIndex = MapData(mapa, X, Y).ObjInfo.ObjIndex
            
    If MapObjIndex <> 0 Then
        ' Es el mismo objeto?
        If MapObjIndex = ObjIndex Then
            ' La suma es menor a 10k?
            ObjectExistsInTile = True
        End If
    End If

  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function ObjectExistsInTile de GameLogic.bas")
End Function


Sub ClosestLegalPos(Pos As WorldPos, ByRef nPos As WorldPos, Optional PuedeAgua As Boolean = False, _
                    Optional PuedeTierra As Boolean = True, Optional ByVal CheckExitTile As Boolean = False, Optional ByVal CheckPortals As Boolean = True)
'*****************************************************************
'Author: Unknown (original version)
'Last Modification: 09/14/2010 (Marco)
'History:
' - 01/24/2007 (ToxicWaste)
'Encuentra la posicion legal mas cercana y la guarda en nPos
'*****************************************************************
On Error GoTo ErrHandler
  

    Dim Found As Boolean
    Dim LoopC As Integer
    Dim tX As Integer
    Dim tY As Integer
    
    nPos = Pos
    tX = Pos.X
    tY = Pos.Y
    
    LoopC = 1
    
    ' La primera posicion es valida?
    If LegalPos(Pos.Map, nPos.X, nPos.Y, PuedeAgua, PuedeTierra, CheckExitTile, CheckPortals) Then
        Found = True
    
    ' Busca en las demas posiciones, en forma de "rombo"
    Else
        While (Not Found) And LoopC <= 12
            If RhombLegalPos(Pos.Map, tX, tY, LoopC, PuedeAgua, PuedeTierra, CheckExitTile) Then
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
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ClosestLegalPos de GameLogic.bas")
End Sub

Private Sub ClosestStablePos(Pos As WorldPos, ByRef nPos As WorldPos)
'***************************************************
'Author: Unknown
'Last Modification: 09/14/2010
'Encuentra la posicion legal mas cercana que no sea un portal y la guarda en nPos
'*****************************************************************
On Error GoTo ErrHandler
  

Call ClosestLegalPos(Pos, nPos, , , True)

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ClosestStablePos de GameLogic.bas")
End Sub

Function NameIndex(ByVal Name As String) As Integer
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    Dim UserIndex As Long
    
    '¿Nombre valido?
    If LenB(Name) = 0 Then
        NameIndex = 0
        Exit Function
    End If
    
    If InStrB(Name, "+") <> 0 Then
        Name = UCase$(Replace(Name, "+", " "))
    End If
    
    UserIndex = 1
    Do Until UCase$(UserList(UserIndex).Name) = UCase$(Name)
        
        UserIndex = UserIndex + 1
        
        If UserIndex > MaxUsers Then
            NameIndex = 0
            Exit Function
        End If
    Loop
     
    NameIndex = UserIndex
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function NameIndex de GameLogic.bas")
End Function

Function CheckForSameIP(ByVal UserIndex As Integer, ByVal UserIP As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    Dim LoopC As Long
    
    For LoopC = 1 To LastUser
        If UserList(LoopC).flags.UserLogged = True Then
            If UserList(LoopC).IP = UserIP And UserIndex <> LoopC Then
                CheckForSameIP = True
                Exit Function
            End If
        End If
    Next LoopC
    
    CheckForSameIP = False
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function CheckForSameIP de GameLogic.bas")
End Function

Function CheckForSameName(ByVal Name As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

'Controlo que no existan usuarios con el mismo nombre
    Dim LoopC As Long
    
    For LoopC = 1 To LastUser
        If UserList(LoopC).flags.UserLogged Then
            
            'If UCase$(UserList(LoopC).Name) = UCase$(Name) And UserList(LoopC).ConnID <> -1 Then
            'OJO PREGUNTAR POR EL CONNID <> -1 PRODUCE QUE UN PJ EN DETERMINADO
            'MOMENTO PUEDA ESTAR LOGUEADO 2 VECES (IE: CIERRA EL SOCKET DESDE ALLA)
            'ESE EVENTO NO DISPARA UN SAVE USER, LO QUE PUEDE SER UTILIZADO PARA DUPLICAR ITEMS
            'ESTE BUG EN ALKON PRODUJO QUE EL SERVIDOR ESTE CAIDO DURANTE 3 DIAS. ATENTOS.
            
            If UCase$(UserList(LoopC).Name) = UCase$(Name) Then
                CheckForSameName = True
                Exit Function
            End If
        End If
    Next LoopC
    
    CheckForSameName = False
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function CheckForSameName de GameLogic.bas")
End Function

Sub HeadtoPos(ByVal head As eHeading, ByRef Pos As WorldPos)
'***************************************************
'Author: Unknown
'Last Modification: -
'Toma una posicion y se mueve hacia donde esta perfilado
'*****************************************************************
On Error GoTo ErrHandler
  

    Select Case head
        Case eHeading.NORTH
            Pos.Y = Pos.Y - 1
        
        Case eHeading.SOUTH
            Pos.Y = Pos.Y + 1
        
        Case eHeading.EAST
            Pos.X = Pos.X + 1
        
        Case eHeading.WEST
            Pos.X = Pos.X - 1
    End Select
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HeadtoPos de GameLogic.bas")
End Sub


Public Function HeadToPosLateral(ByVal heading As Integer, ByVal MovementAmount As Integer, ByRef Pos As WorldPos) As WorldPos
    
On Error GoTo ErrHandler
    HeadToPosLateral = Pos
    Select Case heading
        Case eHeading.NORTH, eHeading.SOUTH
            HeadToPosLateral.X = Pos.X + MovementAmount
        Case eHeading.EAST, eHeading.WEST
            HeadToPosLateral.Y = Pos.Y + MovementAmount
    End Select
    
    Exit Function
    
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HeadToPosLateral de GameLogic.bas")
End Function

Function LegalPos(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal PuedeAgua As Boolean = False, Optional ByVal PuedeTierra As Boolean = True, Optional ByVal CheckExitTile As Boolean = False, Optional ByVal CheckPortals As Boolean = True) As Boolean
'***************************************************
'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
'Last Modification: 23/01/2007
'Checks if the position is Legal.
'***************************************************
On Error GoTo ErrHandler
  

    '¿Es un mapa valido?
    If (Map <= 0 Or Map > NumMaps) Or _
       (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
                LegalPos = False
    Else
        With MapData(Map, X, Y)
            If PuedeAgua And PuedeTierra Then
                LegalPos = (.Blocked <> 1) And _
                           (.UserIndex = 0) And _
                           (.NpcIndex = 0)
            ElseIf PuedeTierra And Not PuedeAgua Then
                LegalPos = (.Blocked <> 1) And _
                           (.UserIndex = 0) And _
                           (.NpcIndex = 0) And _
                           (Not HayAgua(Map, X, Y))
            ElseIf PuedeAgua And Not PuedeTierra Then
                LegalPos = (.Blocked <> 1) And _
                           (.UserIndex = 0) And _
                           (.NpcIndex = 0) And _
                           (HayAgua(Map, X, Y))
            Else
                LegalPos = False
            End If
        End With
        
        If CheckExitTile Then
            LegalPos = LegalPos And (MapData(Map, X, Y).TileExit.Map = 0)
        End If
        
        If CheckPortals Then
            LegalPos = LegalPos And Not IsPortal(Map, X, Y)
        End If
        
    End If

  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function LegalPos de GameLogic.bas")
End Function


Function MoveToLegalPos(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal PuedeAgua As Boolean = False, Optional ByVal PuedeTierra As Boolean = True) As Boolean
'***************************************************
'Autor: ZaMa
'Last Modification: 13/07/2009
'Checks if the position is Legal, but considers that if there's a casper, it's a legal movement.
'13/07/2009: ZaMa - Now it's also legal move where an invisible admin is.
'***************************************************
On Error GoTo ErrHandler
  

Dim UserIndex As Integer
Dim IsDeadChar As Boolean
Dim IsAdminInvisible As Boolean

'¿Es un mapa valido?
If (Map <= 0 Or Map > NumMaps) Or _
   (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
        MoveToLegalPos = False
Else
    With MapData(Map, X, Y)
        UserIndex = .UserIndex
        
        If UserIndex > 0 Then
            IsDeadChar = (UserList(UserIndex).flags.Muerto = 1)
            IsAdminInvisible = (UserList(UserIndex).flags.AdminInvisible = 1)
        Else
            IsDeadChar = False
            IsAdminInvisible = False
        End If
        
        If PuedeAgua And PuedeTierra Then
            MoveToLegalPos = (.Blocked <> 1) And _
                       (UserIndex = 0 Or IsDeadChar Or IsAdminInvisible) And _
                       (.NpcIndex = 0)
        ElseIf PuedeTierra And Not PuedeAgua Then
            MoveToLegalPos = (.Blocked <> 1) And _
                       (UserIndex = 0 Or IsDeadChar Or IsAdminInvisible) And _
                       (.NpcIndex = 0) And _
                       (Not HayAgua(Map, X, Y))
        ElseIf PuedeAgua And Not PuedeTierra Then
            MoveToLegalPos = (.Blocked <> 1) And _
                       (UserIndex = 0 Or IsDeadChar Or IsAdminInvisible) And _
                       (.NpcIndex = 0) And _
                       (HayAgua(Map, X, Y))
        Else
            MoveToLegalPos = False
        End If
    End With
End If

  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function MoveToLegalPos de GameLogic.bas")
End Function

Public Sub FindLegalPos(ByVal UserIndex As Integer, ByVal Map As Integer, ByRef X As Integer, ByRef Y As Integer)
'***************************************************
'Autor: ZaMa
'Last Modification: 26/03/2009
'Search for a Legal pos for the user who is being teleported.
'***************************************************
On Error GoTo ErrHandler
  

    If MapData(Map, X, Y).UserIndex <> 0 Or _
        MapData(Map, X, Y).NpcIndex <> 0 Then
                    
        ' Se teletransporta a la misma pos a la que estaba
        If MapData(Map, X, Y).UserIndex = UserIndex Then Exit Sub
                            
        Dim FoundPlace As Boolean
        Dim tX As Long
        Dim tY As Long
        Dim Rango As Long
        Dim OtherUserIndex As Integer
    
        For Rango = 1 To 5
            For tY = Y - Rango To Y + Rango
                For tX = X - Rango To X + Rango
                    'Reviso que no haya User ni NPC
                    If MapData(Map, tX, tY).UserIndex = 0 And _
                        MapData(Map, tX, tY).NpcIndex = 0 Then
                        
                        If InMapBounds(Map, tX, tY) Then FoundPlace = True
                        
                        Exit For
                    End If

                Next tX
        
                If FoundPlace Then _
                    Exit For
            Next tY
            
            If FoundPlace Then _
                    Exit For
        Next Rango

    
        If FoundPlace Then 'Si encontramos un lugar, listo, nos quedamos ahi
            X = tX
            Y = tY
        Else
            'Muy poco probable, pero..
            'Si no encontramos un lugar, sacamos al usuario que tenemos abajo, y si es un NPC, lo pisamos.
            OtherUserIndex = MapData(Map, X, Y).UserIndex
            If OtherUserIndex <> 0 Then
                'Si no encontramos lugar, y abajo teniamos a un usuario, lo pisamos y cerramos su comercio seguro
                If isTradingWithUser(OtherUserIndex) Then
                    Dim tempUsu As Integer
                    
                    tempUsu = getTradingUser(OtherUserIndex)
                    
                    'Le avisamos al que estaba comerciando que se tuvo que ir.
                    If UserList(tempUsu).flags.UserLogged Then
                        Call FinComerciarUsu(tempUsu)
                        Call WriteConsoleMsg(tempUsu, "Comercio cancelado. El otro usuario se ha desconectado.", FontTypeNames.FONTTYPE_TALK)
                    End If
                    
                    'Lo sacamos.
                    If UserList(OtherUserIndex).flags.UserLogged Then
                        Call FinComerciarUsu(OtherUserIndex)
                        Call WriteErrorMsg(OtherUserIndex, "Alguien se ha conectado donde te encontrabas, por favor reconéctate...")
                    End If
                End If
            
                Call CloseSocket(OtherUserIndex)
            End If
        End If
    End If

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub FindLegalPos de GameLogic.bas")
End Sub

Function LegalPosNPC(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal AguaValida As Byte, _
                        Optional ByVal IsPet As Boolean = False, Optional ByVal TierraInvalida As Boolean = False) As Boolean
'***************************************************
'Autor: Unkwnown
'Last Modification: 09/23/2009
'Checks if it's a Legal pos for the npc to move to.
'09/23/2009: Pato - If UserIndex is a AdminInvisible, then is a legal pos.
'***************************************************
On Error GoTo ErrHandler

    Dim IsDeadChar As Boolean
    Dim UserIndex As Integer
    Dim IsAdminInvisible As Boolean
    
    
    If (Map <= 0 Or Map > NumMaps) Or _
        (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
        LegalPosNPC = False
        Exit Function
    End If

    With MapData(Map, X, Y)
        UserIndex = .UserIndex
        If UserIndex > 0 Then
            IsDeadChar = UserList(UserIndex).flags.Muerto = 1
            IsAdminInvisible = (UserList(UserIndex).flags.AdminInvisible = 1)
        Else
            IsDeadChar = False
            IsAdminInvisible = False
        End If
        
        ' if it's a pet, check if is going to walk on a tp
        If IsPet And .ObjInfo.ObjIndex <> 0 Then
            If ObjData(.ObjInfo.ObjIndex).ObjType = eOBJType.otTeleport Then
                LegalPosNPC = False
                Exit Function
            End If
        End If
    
        If AguaValida = 0 Then
            LegalPosNPC = (.Blocked <> 1) And _
            (.UserIndex = 0 Or IsDeadChar Or IsAdminInvisible) And _
            (.NpcIndex = 0) And _
            (.Trigger <> eTrigger.POSINVALIDA Or IsPet) _
            And Not HayAgua(Map, X, Y)
        ElseIf TierraInvalida = False Then
            LegalPosNPC = (.Blocked <> 1) And _
            (.UserIndex = 0 Or IsDeadChar Or IsAdminInvisible) And _
            (.NpcIndex = 0) And _
            (.Trigger <> eTrigger.POSINVALIDA Or IsPet)
        Else
            LegalPosNPC = (.Blocked <> 1) And _
            (.UserIndex = 0 Or IsDeadChar Or IsAdminInvisible) And _
            (.NpcIndex = 0) And _
            (.Trigger <> eTrigger.POSINVALIDA And HayAgua(Map, X, Y) Or IsPet)
        End If
    End With
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function LegalPosNPC de GameLogic.bas")
End Function

Sub SendHelp(ByVal Index As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

Dim NumHelpLines As Integer
Dim LoopC As Integer

NumHelpLines = Val(GetVar(DatPath & "Help.dat", "INIT", "NumLines"))

For LoopC = 1 To NumHelpLines
    Call WriteConsoleMsg(Index, GetVar(DatPath & "Help.dat", "Help", "Line" & LoopC), FontTypeNames.FONTTYPE_INFO)
Next LoopC

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SendHelp de GameLogic.bas")
End Sub

Public Sub Expresar(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    If Npclist(NpcIndex).NroExpresiones > 0 Then
        Dim randomi
        randomi = RandomNumber(1, Npclist(NpcIndex).NroExpresiones)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(Npclist(NpcIndex).Expresiones(randomi), Npclist(NpcIndex).Char.CharIndex, vbWhite))
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Expresar de GameLogic.bas")
End Sub

Sub LookatTile(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 30/03/2017
'13/02/2009: ZaMa - El nombre del gm que aparece por consola al clickearlo, tiene el color correspondiente a su rango
'16/09/2014: D'Artagnan - Citizen or criminal tags unvisible if enlisted.
'27/07/2016: Anagrama - Ahora muestra en consola el nombre de los npc no hostiles.
'30/03/2017: G Toyz - Ahora muestra en consola el nombre y el tag de los npcs no hostiles.
'***************************************************

On Error GoTo ErrHandler

'Responde al click del usuario sobre el mapa
Dim FoundChar As Byte
Dim FoundSomething As Byte
Dim TempCharIndex As Integer
Dim Stat As String
Dim ft As FontTypeNames

With UserList(UserIndex)
    '¿Rango Visión? (ToxicWaste)
    If (Abs(.Pos.Y - Y) > RANGO_VISION_Y) Or (Abs(.Pos.X - X) > RANGO_VISION_X) Then
        Exit Sub
    End If
    
    '¿Posicion valida?
    If InMapBounds(Map, X, Y) Then
        With .flags
            .TargetMap = Map
            .TargetX = X
            .TargetY = Y
            '¿Es un obj?
            If MapData(Map, X, Y).ObjInfo.ObjIndex > 0 Then
                'Informa el nombre
                .TargetObjMap = Map
                .TargetObjX = X
                .TargetObjY = Y
                FoundSomething = 1
            ElseIf MapData(Map, X + 1, Y).ObjInfo.ObjIndex > 0 Then
                'Informa el nombre
                If ObjData(MapData(Map, X + 1, Y).ObjInfo.ObjIndex).ObjType = eOBJType.otPuertas Then
                    .TargetObjMap = Map
                    .TargetObjX = X + 1
                    .TargetObjY = Y
                    FoundSomething = 1
                End If
            ElseIf MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex > 0 Then
                If ObjData(MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex).ObjType = eOBJType.otPuertas Then
                    'Informa el nombre
                    .TargetObjMap = Map
                    .TargetObjX = X + 1
                    .TargetObjY = Y + 1
                    FoundSomething = 1
                End If
            ElseIf MapData(Map, X, Y + 1).ObjInfo.ObjIndex > 0 Then
                If ObjData(MapData(Map, X, Y + 1).ObjInfo.ObjIndex).ObjType = eOBJType.otPuertas Then
                    'Informa el nombre
                    .TargetObjMap = Map
                    .TargetObjX = X
                    .TargetObjY = Y + 1
                    FoundSomething = 1
                End If
            End If
            
            If FoundSomething = 1 Then
                .TargetObj = MapData(Map, .TargetObjX, .TargetObjY).ObjInfo.ObjIndex
                
                Dim AdditionalData As String
                Dim IgnoreMessage As Boolean
                
                If ObjData(.TargetObj).ObjType = eOBJType.otTrigger Or ObjData(.TargetObj).ObjType = eOBJType.otTrampa Then
                    IgnoreMessage = (ObjData(.TargetObj).Trigger.CanDetect = 0)
                End If
                
                If ObjData(.TargetObj).ObjType = otResource And MapData(.TargetObjMap, .TargetObjX, .TargetObjY).ObjInfo.PendingQty = 0 Then
                    AdditionalData = " - Agotado"
                End If
                
                If Not IgnoreMessage Then
                    If MostrarCantidad(.TargetObj) Then
                        Call WriteConsoleMsg(UserIndex, ObjData(.TargetObj).Name & " - " & MapData(.TargetObjMap, .TargetObjX, .TargetObjY).ObjInfo.Amount & AdditionalData, FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call WriteConsoleMsg(UserIndex, ObjData(.TargetObj).Name & AdditionalData, FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
            End If
            '¿Es un personaje?
            If Y + 1 <= YMaxMapSize Then
                If MapData(Map, X, Y + 1).UserIndex > 0 Then
                    TempCharIndex = MapData(Map, X, Y + 1).UserIndex
                    FoundChar = 1
                End If
                If MapData(Map, X, Y + 1).NpcIndex > 0 Then
                    TempCharIndex = MapData(Map, X, Y + 1).NpcIndex
                    FoundChar = 2
                End If
            End If
            '¿Es un personaje?
            If FoundChar = 0 Then
                If MapData(Map, X, Y).UserIndex > 0 Then
                    TempCharIndex = MapData(Map, X, Y).UserIndex
                    FoundChar = 1
                    
                    If EsGm(UserIndex) Or (Not EsGm(UserIndex) And (EsGm(TempCharIndex) And UserList(TempCharIndex).flags.AdminInvisible = True)) Then
                        FoundChar = 0
                        FoundSomething = False
                    End If
               
                End If
                If MapData(Map, X, Y).NpcIndex > 0 Then
                    TempCharIndex = MapData(Map, X, Y).NpcIndex
                    FoundChar = 2
                End If
            End If
        End With
    
    
        'Reaccion al personaje
        If FoundChar = 1 Then '  ¿Encontro un Usuario?
           If UserList(TempCharIndex).flags.AdminInvisible = 0 Or .flags.Privilegios And PlayerType.Dios Then
                With UserList(TempCharIndex)
                    If LenB(.DescRM) = 0 And (.ShowName Or .flags.Mimetizado) Then 'No tiene descRM y quiere que se vea su nombre o está mimetizado.
                        If EsNewbie(TempCharIndex) Then
                            Stat = " <NEWBIE>"
                            ft = FontTypeNames.FONTTYPE_NEWBIE
                        Else
                            ft = FontTypeNames.FONTTYPE_NEUTRAL
                        End If
                        
                        If .Faccion.ArmadaReal = 1 Or .Faccion.Alignment = eCharacterAlignment.FactionRoyal Then
                            Stat = Stat & " <Ejército Real> " & "<" & TituloReal(TempCharIndex) & ">"
                        ElseIf .Faccion.FuerzasCaos = 1 Or .Faccion.Alignment = eCharacterAlignment.FactionLegion Then
                            Stat = Stat & " <Legión Oscura> " & "<" & TituloCaos(TempCharIndex) & ">"
                        End If
                        
                        If .Guild.IdGuild > 0 Then
                            Stat = Stat & " <" & GuildList(.Guild.GuildIndex).Name & ">"
                        End If
                        
                        If Len(.desc) > 0 Then
                            Stat = "Ves a " & .secName & Stat & " - " & .desc
                        Else
                            Stat = "Ves a " & .secName & Stat
                        End If
                                           
                        If .flags.Privilegios And PlayerType.RoyalCouncil Then
                            Stat = Stat & " [CONSEJO DE BANDERBILL]"
                            ft = FontTypeNames.FONTTYPE_CONSEJOVesA
                        ElseIf .flags.Privilegios And PlayerType.ChaosCouncil Then
                            Stat = Stat & " [CONCILIO DE LAS SOMBRAS]"
                            ft = FontTypeNames.FONTTYPE_CONSEJOCAOSVesA
                        Else
                            If Not .flags.Privilegios And PlayerType.User Then
                                Stat = Stat & " <GAME MASTER>"
                                
                                ' Elijo el color segun el rango del GM:
                                ' Dios
                                If .flags.Privilegios = PlayerType.Dios Then
                                    ft = FontTypeNames.FONTTYPE_DIOS
                                ' Gm
                                ElseIf .flags.Privilegios = PlayerType.SemiDios Then
                                    ft = FontTypeNames.FONTTYPE_GM
                                ' Conse
                                ElseIf .flags.Privilegios = PlayerType.Consejero Then
                                    ft = FontTypeNames.FONTTYPE_CONSE
                                ' Rm o Dsrm
                                ElseIf .flags.Privilegios = (PlayerType.RoleMaster Or PlayerType.Consejero) Or .flags.Privilegios = (PlayerType.RoleMaster Or PlayerType.Dios) Then
                                    ft = FontTypeNames.FONTTYPE_EJECUCION
                                End If
                                
                            ElseIf .Faccion.Alignment = FactionLegion Then
                                'If .Faccion.FuerzasCaos = 0 Then _
                                '    Stat = Stat & " <CRIMINAL>"
                                ft = FontTypeNames.FONTTYPE_FIGHT
                            ElseIf .Faccion.Alignment = FactionRoyal Then
                                'If .Faccion.ArmadaReal = 0 Then _
                                '    Stat = Stat & " <CIUDADANO>"
                                ft = FontTypeNames.FONTTYPE_CITIZEN
                            End If
                        End If
                    Else  'Si tiene descRM la muestro siempre.
                        Stat = .DescRM
                        ft = FontTypeNames.FONTTYPE_INFOBOLD
                    End If
                End With
                
                If LenB(Stat) > 0 Then
                    Call WriteConsoleMsg(UserIndex, Stat, ft)
                End If
                
                FoundSomething = 1
                .flags.TargetUser = TempCharIndex
                .flags.TargetNPC = 0
                .flags.TargetNpcTipo = eNPCType.Comun
           End If
        End If
    
        With .flags
            If FoundChar = 2 Then '¿Encontro un NPC?
                Dim estatus As String
                Dim MinHp As Long
                Dim MaxHp As Long
                Dim SupervivenciaSkill As Byte
                Dim sDesc As String
                
                MinHp = Npclist(TempCharIndex).Stats.MinHp
                MaxHp = Npclist(TempCharIndex).Stats.MaxHp
                SupervivenciaSkill = GetSkills(UserIndex, eSkill.Supervivencia)
                
                If .Privilegios And (PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin) Then
                    estatus = "(" & MinHp & "/" & MaxHp & ") "
                Else
                    If .Muerto = 0 Then
                    
                        If SupervivenciaSkill <= 20 Then
                            estatus = "(Dudoso) "
                            
                        ElseIf SupervivenciaSkill <= 40 Then
                            If MinHp < (MaxHp / 2) Then
                                estatus = "(Herido) "
                            Else
                                estatus = "(Sano) "
                            End If
                            
                        ElseIf SupervivenciaSkill <= 60 Then
                            If MinHp < (MaxHp * 0.5) Then
                                estatus = "(Malherido) "
                            ElseIf MinHp < (MaxHp * 0.75) Then
                                estatus = "(Herido) "
                            Else
                                estatus = "(Sano) "
                            End If
                            
                        ElseIf SupervivenciaSkill <= 80 Then
                            If MinHp < (MaxHp * 0.25) Then
                                estatus = "(Muy malherido) "
                            ElseIf MinHp < (MaxHp * 0.5) Then
                                estatus = "(Herido) "
                            ElseIf MinHp < (MaxHp * 0.75) Then
                                estatus = "(Levemente herido) "
                            Else
                                estatus = "(Sano) "
                            End If
                            
                        ElseIf SupervivenciaSkill < 100 Then
                            If MinHp < (MaxHp * 0.05) Then
                                estatus = "(Agonizando) "
                            ElseIf MinHp < (MaxHp * 0.1) Then
                                estatus = "(Casi muerto) "
                            ElseIf MinHp < (MaxHp * 0.25) Then
                                estatus = "(Muy Malherido) "
                            ElseIf MinHp < (MaxHp * 0.5) Then
                                estatus = "(Herido) "
                            ElseIf MinHp < (MaxHp * 0.75) Then
                                estatus = "(Levemente herido) "
                            ElseIf MinHp < (MaxHp) Then
                                estatus = "(Sano) "
                            Else
                                estatus = "(Intacto) "
                            End If
                        Else
                            estatus = "(" & MinHp & "/" & MaxHp & ") "
                        End If
                    End If
                End If
                
                If Len(Npclist(TempCharIndex).desc) > 1 Then
                    Stat = Npclist(TempCharIndex).desc
                    
                    '¿Es el rey o el demonio?
                    If Npclist(TempCharIndex).NPCtype = eNPCType.Noble Then
                        If Npclist(TempCharIndex).flags.Faccion = 0 Then 'Es el Rey.
                            'Si es de la Legión Oscura mostramos el mensaje correspondiente y lo ejecutamos, sólo si es un usuario común:
                            If UserList(UserIndex).Faccion.FuerzasCaos = 1 Or UserList(UserIndex).Faccion.Alignment = eCharacterAlignment.FactionLegion Then
                                Stat = MENSAJE_REY_CAOS
                                If .Privilegios And PlayerType.User Then
                                    If .Muerto = 0 Then Call UserDie(UserIndex)
                                End If
                            End If
                        Else 'Es el demonio
                            'Si es de la Armada Real mostramos el mensaje correspondiente y lo ejecutamos. sólo si es un usuario común:
                            
                            If UserList(UserIndex).Faccion.ArmadaReal = 1 Or UserList(UserIndex).Faccion.Alignment = eCharacterAlignment.FactionRoyal Then
                                Stat = MENSAJE_DEMONIO_REAL
                                If .Privilegios And PlayerType.User Then
                                    If .Muerto = 0 Then Call UserDie(UserIndex)
                                End If
                           
                            End If
                        End If
                    End If
                    
                    'Enviamos el mensaje propiamente dicho:
                    Call WriteChatOverHead(UserIndex, Stat, Npclist(TempCharIndex).Char.CharIndex, vbWhite)
                End If
                
                If Npclist(TempCharIndex).Attackable Then

                    If Npclist(TempCharIndex).MaestroUser > 0 Then
                        Call WriteConsoleMsg(UserIndex, estatus & Npclist(TempCharIndex).Name & " es mascota de " & UserList(Npclist(TempCharIndex).MaestroUser).Name & ".", FontTypeNames.FONTTYPE_INFO)
                    Else
                        sDesc = estatus & Npclist(TempCharIndex).Name
                        If Npclist(TempCharIndex).Owner > 0 Then sDesc = sDesc & " le pertenece a " & UserList(Npclist(TempCharIndex).Owner).Name
                        sDesc = sDesc & "."
                        
                        Call WriteConsoleMsg(UserIndex, sDesc, FontTypeNames.FONTTYPE_INFO)
                        
                        If .Privilegios And (PlayerType.Dios Or PlayerType.Admin) And Npclist(TempCharIndex).flags.AttackedFirstBy <> "" Then
                            Call WriteConsoleMsg(UserIndex, "Le pegó primero: " & Npclist(TempCharIndex).flags.AttackedFirstBy & ".", FontTypeNames.FONTTYPE_INFO)
                        End If
                    End If
                    
                Else
                    Call WriteConsoleMsg(UserIndex, "Ves a " & Npclist(TempCharIndex).Name & IIf(Len(Npclist(TempCharIndex).Tag) > 0, " - <" & Npclist(TempCharIndex).Tag & ">", vbNullString) & ".", FontTypeNames.FONTTYPE_NPCNAME)
                End If
                
                FoundSomething = 1
                .TargetNpcTipo = Npclist(TempCharIndex).NPCtype
                .TargetNPC = TempCharIndex
                .TargetUser = 0
                .TargetObj = 0
            End If
            
            If FoundChar = 0 Then
                .TargetNPC = 0
                .TargetNpcTipo = eNPCType.Comun
                .TargetUser = 0
            End If
            
            '*** NO ENCOTRO NADA ***
            If FoundSomething = 0 Then
                .TargetNPC = 0
                .TargetNpcTipo = eNPCType.Comun
                .TargetUser = 0
                .TargetObj = 0
                .TargetObjMap = 0
                .TargetObjX = 0
                .TargetObjY = 0
                Call WriteMultiMessage(UserIndex, eMessages.DontSeeAnything)
            End If
        End With
    Else
        If FoundSomething = 0 Then
            With .flags
                .TargetNPC = 0
                .TargetNpcTipo = eNPCType.Comun
                .TargetUser = 0
                .TargetObj = 0
                .TargetObjMap = 0
                .TargetObjX = 0
                .TargetObjY = 0
            End With
            
            Call WriteMultiMessage(UserIndex, eMessages.DontSeeAnything)
        End If
    End If
End With

Exit Sub

ErrHandler:
    Call LogError("Error en LookAtTile. Error " & Err.Number & " : " & Err.Description)

End Sub

Public Sub ShowMenu(ByVal UserIndex As Integer, ByVal Map As Integer, _
    ByVal X As Integer, ByVal Y As Integer)
'***************************************************
'Autor: ZaMa
'Last Modification: 30/03/2017: [G Toyz]
'Shows menu according to user, npc or object right clicked.
'27/11/2014: D'Artagnan - Bug fixes.
'30/03/2017: G Toyz - Fix: No abrían los comercios.
'***************************************************

On Error GoTo ErrHandler

    With UserList(UserIndex)

        ' In Vision Range
        If (Abs(.Pos.Y - Y) > RANGO_VISION_Y) Or (Abs(.Pos.X - X) > RANGO_VISION_X) Then Exit Sub
        ' Valid position?
        If Not InMapBounds(Map, X, Y) Then Exit Sub
        With .flags
            ' Trading?
            If .Comerciando Then Exit Sub
            ' Reset flags
            .TargetNPC = 0
            .TargetNpcTipo = eNPCType.Comun
            .TargetUser = 0
            .TargetObj = 0
            .TargetObjMap = 0
            .TargetObjX = 0
            .TargetObjY = 0
            
            .TargetMap = Map
            .TargetX = X
            .TargetY = Y
            
            Dim tmpIndex As Integer
            Dim FoundChar As Byte
            Dim MenuIndex As Integer
            
            ' Npc or user? (lower position)
            If Y + 1 <= YMaxMapSize Then
                ' User?
                tmpIndex = MapData(Map, X, Y + 1).UserIndex
               
                If tmpIndex > 0 Then
                    ' Invalid: Admin invisible, user invisible or hidden (if not is userindex)
                    If (UserList(tmpIndex).flags.AdminInvisible Or _
                        UserList(tmpIndex).flags.invisible Or _
                        UserList(tmpIndex).flags.Oculto) = 0 Or tmpIndex = UserIndex Then
                        
                        FoundChar = 1
                    End If
                End If
                
                ' Npc?
                If MapData(Map, X, Y + 1).NpcIndex > 0 Then
                    tmpIndex = MapData(Map, X, Y + 1).NpcIndex
                    FoundChar = 2
                End If
            End If
            ' Npc or user? (upper position)
            If FoundChar = 0 Then
                
                ' User?
                tmpIndex = MapData(Map, X, Y).UserIndex

                If tmpIndex > 0 Then
                    ' Invalid: Admin invisible, user invisible or hidden (if not is userindex)
                    If (UserList(tmpIndex).flags.AdminInvisible Or _
                        UserList(tmpIndex).flags.invisible Or _
                        UserList(tmpIndex).flags.Oculto) = 0 Or tmpIndex = UserIndex Then

                        FoundChar = 1
                    End If
                End If
        
                ' Npc?
                If MapData(Map, X, Y).NpcIndex > 0 Then

                    tmpIndex = MapData(Map, X, Y).NpcIndex

                    FoundChar = 2
                End If
            End If
            
            ' User
            If FoundChar = 1 Then
                ' Self clicked => pick item
                If tmpIndex = UserIndex Then
                    'Lower rank administrators can't pick up items
                    If .Privilegios And PlayerType.Consejero Then
                        If Not .Privilegios And PlayerType.RoleMaster Then Exit Sub
                    End If
                    ' Pick item
                    Call GetObj(UserIndex)

                Else
                    ' Sharing npc?
                    If .ShareNpcWith = tmpIndex Then
                        MenuIndex = eMenues.ieOtroUserCompartiendoNpc
                    Else
                        MenuIndex = eMenues.ieOtroUser
                    End If
                    
                    .TargetUser = tmpIndex
                End If

            ' Npc
            ElseIf FoundChar = 2 Then
                ' Has menu attached?

                If Npclist(tmpIndex).MenuIndex <> 0 Then
                    MenuIndex = Npclist(tmpIndex).MenuIndex
                End If
                
                If Npclist(tmpIndex).flags.Domable <> 0 Then MenuIndex = eMenues.ieNpcDomable
                    
                .TargetNpcTipo = Npclist(tmpIndex).NPCtype
                .TargetNPC = tmpIndex

                ' Alive or priest target?
                If MapData(Map, X, Y).NpcIndex > 0 Then
                    If .Muerto = 1 And Npclist(MapData(Map, X, Y).NpcIndex).NPCtype <> 1 Then Exit Sub
                End If

            End If

            ' No user or npc found
            If FoundChar = 0 Then
                
                ' Is there any object?
                tmpIndex = MapData(Map, X, Y).ObjInfo.ObjIndex

                If tmpIndex > 0 Then
                    ' Has menu attached?

                    MenuIndex = ObjData(tmpIndex).MenuIndex

                    If MenuIndex = eMenues.ieFogata Then
                        If .Descansar = 1 Then MenuIndex = eMenues.ieFogataDescansando
                    End If
    
                    .TargetObj = tmpIndex
                    .TargetObjMap = Map
                    .TargetObjX = X
                    .TargetObjY = Y
                End If
            End If
        End With
    End With

    ' Show it
    If MenuIndex <> 0 Then _
        Call WriteShowMenu(UserIndex, MenuIndex)

    Exit Sub

ErrHandler:
    Call LogError("Error en ShowMenu. Error " & Err.Number & " : " & Err.Description)
End Sub

Function FindDirection(Pos As WorldPos, Target As WorldPos) As eHeading
'***************************************************
'Author: Unknown
'Last Modification: -
'Devuelve la direccion en la cual el target se encuentra
'desde pos, 0 si la direc es igual
'*****************************************************************
On Error GoTo ErrHandler
  

    Dim X As Integer
    Dim Y As Integer
    
    X = Pos.X - Target.X
    Y = Pos.Y - Target.Y
    
    'NE
    If Sgn(X) = -1 And Sgn(Y) = 1 Then
        FindDirection = IIf(RandomNumber(0, 1), eHeading.NORTH, eHeading.EAST)
        Exit Function
    End If
    
    'NW
    If Sgn(X) = 1 And Sgn(Y) = 1 Then
        FindDirection = IIf(RandomNumber(0, 1), eHeading.WEST, eHeading.NORTH)
        Exit Function
    End If
    
    'SW
    If Sgn(X) = 1 And Sgn(Y) = -1 Then
        FindDirection = IIf(RandomNumber(0, 1), eHeading.WEST, eHeading.SOUTH)
        Exit Function
    End If
    
    'SE
    If Sgn(X) = -1 And Sgn(Y) = -1 Then
        FindDirection = IIf(RandomNumber(0, 1), eHeading.SOUTH, eHeading.EAST)
        Exit Function
    End If
    
    'Sur
    If Sgn(X) = 0 And Sgn(Y) = -1 Then
        FindDirection = eHeading.SOUTH
        Exit Function
    End If
    
    'norte
    If Sgn(X) = 0 And Sgn(Y) = 1 Then
        FindDirection = eHeading.NORTH
        Exit Function
    End If
    
    'oeste
    If Sgn(X) = 1 And Sgn(Y) = 0 Then
        FindDirection = eHeading.WEST
        Exit Function
    End If
    
    'este
    If Sgn(X) = -1 And Sgn(Y) = 0 Then
        FindDirection = eHeading.EAST
        Exit Function
    End If
    
    'misma
    If Sgn(X) = 0 And Sgn(Y) = 0 Then
        FindDirection = 0
        Exit Function
    End If

  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function FindDirection de GameLogic.bas")
End Function

Public Function ItemNoEsDeMapa(ByVal Index As Integer, ByVal bIsExit As Boolean) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    With ObjData(Index)
        ItemNoEsDeMapa = .ObjType <> eOBJType.otPuertas And _
                    .ObjType <> eOBJType.otForos And _
                    .ObjType <> eOBJType.otCarteles And _
                    .ObjType <> eOBJType.otArboles And _
                    .ObjType <> eOBJType.otYacimiento And _
                    Not (.ObjType = eOBJType.otTeleport And bIsExit)
    
    End With

  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function ItemNoEsDeMapa de GameLogic.bas")
End Function

Public Function MostrarCantidad(ByVal Index As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    With ObjData(Index)
        MostrarCantidad = .ObjType <> eOBJType.otPuertas And _
                    .ObjType <> eOBJType.otForos And _
                    .ObjType <> eOBJType.otCarteles And _
                    .ObjType <> eOBJType.otArboles And _
                    .ObjType <> eOBJType.otYacimiento And _
                    .ObjType <> eOBJType.otTeleport And _
                    .ObjType <> eOBJType.otTrigger And _
                    .ObjType <> eOBJType.otResource
                    
    End With

  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function MostrarCantidad de GameLogic.bas")
End Function

Public Function RestrictStringToByte(ByRef restrict As String) As Byte
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 04/18/2011
'
'***************************************************
On Error GoTo ErrHandler
  
restrict = UCase$(restrict)

Select Case restrict
    Case "NEWBIE"
        RestrictStringToByte = eRestrict.restrict_newbie
        
    Case "ARMADA"
        RestrictStringToByte = eRestrict.restrict_armada
        
    Case "CAOS"
        RestrictStringToByte = eRestrict.restrict_caos
        
    Case "FACCION"
        RestrictStringToByte = eRestrict.restrict_faccion
        
    Case Else
        RestrictStringToByte = eRestrict.restrict_no
End Select
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function RestrictStringToByte de GameLogic.bas")
End Function

Public Function RestrictByteToString(ByVal restrict As Byte) As String
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 04/18/2011
'
'***************************************************
On Error GoTo ErrHandler
  
Select Case restrict
    Case eRestrict.restrict_newbie
        RestrictByteToString = "NEWBIE"
        
    Case eRestrict.restrict_armada
        RestrictByteToString = "ARMADA"
        
    Case eRestrict.restrict_caos
        RestrictByteToString = "CAOS"
        
    Case eRestrict.restrict_faccion
        RestrictByteToString = "FACCION"
        
    Case eRestrict.restrict_no
        RestrictByteToString = "NO"
End Select
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function RestrictByteToString de GameLogic.bas")
End Function

Public Function TerrainZoneStringToByte(ByRef restrict As String) As Byte
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 04/18/2011
'
'***************************************************
On Error GoTo ErrHandler
  
restrict = UCase$(restrict)

Select Case restrict
    Case "NIEVE"
        TerrainZoneStringToByte = eTerrainZone.terrain_nieve
        
    Case "DESIERTO"
        TerrainZoneStringToByte = eTerrainZone.terrain_desierto
        
    Case "CIUDAD"
        TerrainZoneStringToByte = eTerrainZone.zone_ciudad
        
    Case "CAMPO"
        TerrainZoneStringToByte = eTerrainZone.zone_campo
        
    Case "DUNGEON"
        TerrainZoneStringToByte = eTerrainZone.zone_dungeon
        
    Case Else
        TerrainZoneStringToByte = eTerrainZone.terrain_bosque
End Select
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function TerrainZoneStringToByte de GameLogic.bas")
End Function

Public Function TerrainZoneByteToString(ByVal restrict As Byte) As String
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 04/18/2011
'
'***************************************************
On Error GoTo ErrHandler
  
Select Case restrict
    Case eTerrainZone.terrain_nieve
        TerrainZoneByteToString = "NIEVE"
        
    Case eTerrainZone.terrain_desierto
        TerrainZoneByteToString = "DESIERTO"
        
    Case eTerrainZone.zone_ciudad
        TerrainZoneByteToString = "CIUDAD"
        
    Case eTerrainZone.zone_campo
        TerrainZoneByteToString = "CAMPO"
        
    Case eTerrainZone.zone_dungeon
        TerrainZoneByteToString = "DUNGEON"
        
    Case eTerrainZone.terrain_bosque
        TerrainZoneByteToString = "BOSQUE"
End Select
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function TerrainZoneByteToString de GameLogic.bas")
End Function

Sub AddTrapToList(ByVal UserIndex As Integer, ByVal MapTrap As Integer, ByVal XTrap As Integer, ByVal YTrap As Integer)
'***************************************************
'Last Modification: 24/08/2020
'
'***************************************************
On Error GoTo ErrHandler
    
    Dim Pos As Integer
    Dim I As Integer
    Dim MaxTrapQty As Integer
    
    Pos = 0
    MaxTrapQty = ConstantesBalance.MaxActiveTrapQty
    
    With UserList(UserIndex).flags
        ' check for empty space for trap
        For I = 1 To MaxTrapQty
            If .ActiveTraps(I).Map = 0 And .ActiveTraps(I).X = 0 And .ActiveTraps(I).Y = 0 Then
                Pos = I
                Exit For
            End If
        Next I
         
        If Pos = 0 Then
            ' Disable the first trap
            Call DisableTrap(UserIndex, .ActiveTraps(1).Map, .ActiveTraps(1).X, .ActiveTraps(1).Y)

            For I = 1 To MaxTrapQty
                If I <> MaxTrapQty Then
                    ' Move trap position in array
                    .ActiveTraps(I) = .ActiveTraps(I + 1)
                Else
                    ' Add trap in last position
                    .ActiveTraps(I).Map = MapTrap
                    .ActiveTraps(I).X = XTrap
                    .ActiveTraps(I).Y = YTrap
                End If
            Next I
        
        Else
            ' Add trap in last position avalaible
            .ActiveTraps(Pos).Map = MapTrap
            .ActiveTraps(Pos).X = XTrap
            .ActiveTraps(Pos).Y = YTrap
        End If
    End With
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub AddTrapToList de GameLogic.bas")
End Sub

Sub DelTrapFromList(ByVal UserIndex As Integer, ByVal MapTrap As Integer, ByVal XTrap As Integer, ByVal YTrap As Integer)
'***************************************************
'Last Modification: 24/08/2020
'
'***************************************************
On Error GoTo ErrHandler
    
    Dim Flag As Boolean
    Dim I As Integer
    Dim MaxTrapQty As Integer

    Flag = False
    MaxTrapQty = ConstantesBalance.MaxActiveTrapQty
    With UserList(UserIndex).flags
        'look for the first match and change from there
        For I = 1 To MaxTrapQty
            If .ActiveTraps(I).Map = MapTrap And .ActiveTraps(I).X = XTrap And .ActiveTraps(I).Y = YTrap Then
                Flag = True
                .ActiveTraps(I).Map = 0
                .ActiveTraps(I).X = 0
                .ActiveTraps(I).Y = 0
            End If
            If I <> MaxTrapQty And Flag Then
                'move everything one position back
                .ActiveTraps(I) = .ActiveTraps(I + 1)
            Else
                If I = MaxTrapQty And Flag Then
                    'blank trap in last position
                    .ActiveTraps(I).Map = 0
                    .ActiveTraps(I).X = 0
                    .ActiveTraps(I).Y = 0
                End If
            End If
    
        Next I
     End With
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub DelTrapFromList de GameLogic.bas")
End Sub


Sub DisableTrap(ByVal UserIndex As Integer, ByVal MapTrap As Integer, ByVal XTrap As Integer, ByVal YTrap As Integer)
On Error GoTo ErrHandler
    Dim ObjIndex As Integer
    ObjIndex = MapData(MapTrap, XTrap, YTrap).ObjInfo.ObjIndex
    
    ' If the object is not defined then do nothing
    If ObjIndex <= 0 Then
        Exit Sub
    End If
    
    ' If the object is not defined then do nothing
    If ObjData(ObjIndex).TrapActivatedObject <= 0 Then
        Exit Sub
    End If
    
    If ObjData(ObjIndex).ObjType <> eOBJType.otTrampa Then
        Exit Sub
    End If
    
    ' Get the index of the new object
    ObjIndex = ObjData(ObjIndex).TrapActivatedObject
    
    ' if it is trap type and has spellindex
    MapData(MapTrap, XTrap, YTrap).ObjInfo.ObjIndex = ObjIndex
    MapData(MapTrap, XTrap, YTrap).ObjInfo.CurrentGrhIndex = ObjData(ObjIndex).GrhIndex
        
    With ObjData(ObjIndex)
        Call SendToItemArea(MapTrap, XTrap, YTrap, PrepareMessageObjectUpdate(XTrap, YTrap, .GrhIndex, .ObjType, GetCreateObjectMetadata(ObjIndex, MapTrap, XTrap, YTrap)))
    End With
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub DisableTrap de GameLogic.bas")
End Sub

Sub DisableAllTrapsForUser(ByVal UserIndex As Integer)
'***************************************************
'Last Modification: 24/08/2020
'
'***************************************************
On Error GoTo ErrHandler
    Dim TrapNumber As Integer
    With UserList(UserIndex).flags
        For TrapNumber = 1 To ConstantesBalance.MaxActiveTrapQty
            If .ActiveTraps(TrapNumber).Map <> 0 And .ActiveTraps(TrapNumber).X <> 0 And .ActiveTraps(TrapNumber).Y <> 0 Then
                Call DisableTrap(UserIndex, UserList(UserIndex).flags.ActiveTraps(TrapNumber).Map, UserList(UserIndex).flags.ActiveTraps(TrapNumber).X, UserList(UserIndex).flags.ActiveTraps(TrapNumber).Y)
                UserList(UserIndex).flags.ActiveTraps(TrapNumber).Map = 0
                UserList(UserIndex).flags.ActiveTraps(TrapNumber).X = 0
                UserList(UserIndex).flags.ActiveTraps(TrapNumber).Y = 0
            End If
        Next TrapNumber
    End With
    Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub DisableAllTrapsForUser de GameLogic.bas")
End Sub

Public Sub ObtainListObjectNearPlayer(ByRef Pos As WorldPos, ByVal MaxDistance As Integer, ByRef ListPosObject() As tSpellPosition, ByVal UserIndex As Integer)

    Dim J As Integer
    Dim MaxArrayDimension As Integer
    Dim Index As Integer
    Dim Query() As Collision.UUID
    Dim PosObj As WorldPos
    Dim I       As Long
    Dim DistanceFromPos As Double
    
    MaxArrayDimension = ModAreas.QueryEntities(UserIndex, ENTITY_TYPE_PLAYER, Query, ENTITY_TYPE_OBJECT)
    ReDim ListPosObject(1 To MaxArrayDimension + 1) As tSpellPosition
    Index = 0
    
    For I = 0 To MaxArrayDimension
      PosObj = ModAreas.Unpack(Query(I).Name)
      DistanceFromPos = Distance(Pos.X, Pos.Y, PosObj.X, PosObj.Y)
      If Pos.Map = PosObj.Map And MaxDistance >= DistanceFromPos Then
         Call LoadPosObject(ListPosObject, PosObj.Map, PosObj.X, PosObj.Y, DistanceFromPos, Index)
      End If
    Next I
    
    If Index = 0 Then
        ReDim ListPosObject(0) As tSpellPosition
        Exit Sub
    End If
    
    ReDim Preserve ListPosObject(1 To Index) As tSpellPosition
    Call VectorDistancePosOrder(ListPosObject)
    
    Exit Sub
End Sub

Private Sub LoadPosObject(ByRef Vector() As tSpellPosition, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Distance As Integer, ByRef Indice As Integer)
    Dim Ret As WorldPos
    
    If MapData(Map, X, Y).ObjInfo.ObjIndex = 0 Then
        Exit Sub
    End If
    
    Indice = Indice + 1
    
    Ret.Map = Map
    Ret.X = X
    Ret.Y = Y
    
    Vector(Indice).Pos = Ret
    Vector(Indice).DistanceFromTarget = Distance
    
    Exit Sub
End Sub

Public Function VectorWorldPosSize(ByRef Vector() As tSpellPosition) As Integer

On Error GoTo ErrHandler

    If ((Not Vector) = -1) Then
        VectorWorldPosSize = 0
    Else
        VectorWorldPosSize = UBound(Vector)
    End If

  Exit Function

ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function VectorWorldPosSize de GameLogic.bas")
End Function

Public Sub VectorDistancePosOrder(ByRef Vector() As tSpellPosition)

On Error GoTo ErrHandler

    Dim I As Integer, J As Integer
    Dim Size As Integer
    Dim PosAux As tSpellPosition
    
    Size = VectorWorldPosSize(Vector)
    
    For I = 1 To Size
        For J = (I + 1) To Size
            If Vector(I).DistanceFromTarget > Vector(J).DistanceFromTarget Then
                PosAux = Vector(I)
                Vector(I) = Vector(J)
                Vector(J) = PosAux
            End If
        Next J
    Next I

  Exit Sub

ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub VectorDistancePosOrder de GameLogic.bas")
End Sub

