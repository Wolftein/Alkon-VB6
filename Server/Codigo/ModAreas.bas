Attribute VB_Name = "ModAreas"
Option Explicit

Public Const ENTITY_TYPE_PLAYER As Long = 0
Public Const ENTITY_TYPE_NPC    As Long = 1
Public Const ENTITY_TYPE_OBJECT As Long = 2

Public Const DEFAULT_ENTITY_WIDTH As Byte = 2
Public Const DEFAULT_ENTITY_HEIGHT As Byte = 2

Public Const DEFAULT_PLAYER_WIDTH As Byte = 1
Public Const DEFAULT_PLAYER_HEIGHT As Byte = 2

Private World As Collision.Grid

Public Sub Initialise(ByVal Zones As Long)
On Error GoTo ErrHandler

    Set World = New Collision.Grid
    
    Call World.Initialise(Zones + 1)
    Call World.Attach(AddressOf OnCreateEntity, AddressOf OnDeleteEntity, AddressOf OnUpdateEntity)

    Exit Sub
    
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Initialise de modAreas.bas")
End Sub

Public Sub CreateEntity(ByVal Name As Long, ByVal Tag As Long, ByRef Coordinates As WorldPos, ByVal Width As Byte, ByVal Height As Byte)
On Error GoTo ErrHandler

    Dim UUID As Collision.UUID
    UUID.Name = Name
    UUID.Type = Tag
    
    Select Case Tag
        Case ENTITY_TYPE_PLAYER
            Call World.Create(UUID, 9, 7, Coordinates.Map, Coordinates.X, Coordinates.Y, Width - 1, Height - 1) ' TODO: Width / Height / Radius
        Case ENTITY_TYPE_NPC
            Call World.Create(UUID, 8, 6, Coordinates.Map, Coordinates.X, Coordinates.Y, Width - 1, Height - 1)    ' TODO: Width / Height / Radius
        Case ENTITY_TYPE_OBJECT
            Call World.Create(UUID, 0, 0, Coordinates.Map, Coordinates.X, Coordinates.Y, Width, Height) ' TODO: Width / Height
    End Select

    Exit Sub
    
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CreateEntity de modAreas.bas")
End Sub

Public Sub DeleteEntity(ByVal Name As Long, ByVal Tag As Long)
On Error GoTo ErrHandler

    Dim UUID As Collision.UUID
    UUID.Name = Name
    UUID.Type = Tag
    
    Call World.Delete(UUID)
    
    Exit Sub
    
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub DeleteEntity de modAreas.bas")
End Sub

Public Sub UpdateEntity(ByVal Name As Long, ByVal Tag As Long, ByRef Coordinates As WorldPos, ByVal Warp As Boolean)
On Error GoTo ErrHandler

    Dim UUID As Collision.UUID
    UUID.Name = Name
    UUID.Type = Tag
    
    Call World.Update(UUID, Coordinates.Map, Coordinates.X, Coordinates.Y, Warp)
    
    Exit Sub
    
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub UpdateEntity de modAreas.bas")
End Sub

Public Function QueryEntities(ByVal Name As Long, ByVal Tag As Long, ByRef Result() As Collision.UUID, Optional ByVal Selection As Long = 255, Optional ByVal Sort As Grid_Sort = eGridSortNone) As Long
On Error GoTo ErrHandler

    Dim UUID As Collision.UUID
    UUID.Name = Name
    UUID.Type = Tag
    
    Call World.Search(UUID, Selection, Sort, Result)
    
    QueryEntities = UBound(Result)
    
    Exit Function
    
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub QueryEntities de modAreas.bas")
End Function

Public Function QueryObservers(ByVal Name As Long, ByVal Tag As Long, ByRef Result() As Collision.UUID, Optional ByVal Selection As Long = 255, Optional ByVal Sort As Grid_Sort = eGridSortNone) As Long
On Error GoTo ErrHandler

    Dim UUID As Collision.UUID
    UUID.Name = Name
    UUID.Type = Tag
    
    Call World.query(UUID, Selection, Sort, Result)
    
    QueryObservers = UBound(Result)
    
    Exit Function
    
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub QueryObservers de modAreas.bas")
End Function

Public Function QueryAt(ByVal Map As Long, ByVal X As Long, ByVal Y As Long, ByVal Distance As Long, ByRef Result() As Collision.UUID, Optional ByVal Selection As Long = 255, Optional ByVal Sort As Grid_Sort = eGridSortNone) As Long
On Error GoTo ErrHandler

    Call World.QueryAt(Map, X, Y, Distance, Selection, Sort, Result)
    
    QueryAt = UBound(Result)
    
    Exit Function
    
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub QueryAt de modAreas.bas")
End Function

Public Function Pack(ByVal Map As Long, ByVal X As Long, ByVal Y As Long) As Long

    Pack = ((Map And &H3FF) * &H4000) Or ((X And &H7F) * &H80) Or (Y And &H7F) ' 10 + 7 + 7 = 24B UniqueID
       
End Function

Public Function Unpack(ByVal ID As Long) As WorldPos
    Unpack.Map = (ID \ &H4000) And &H3FF
    Unpack.X = (ID \ &H80) And &H7F
    Unpack.Y = (ID And &H7F)
End Function

Private Sub OnCreateEntity(ByRef Instigator As Collision.UUID, ByRef Observer As Collision.UUID)
On Error GoTo ErrHandler

    Dim Coordinates As WorldPos
          
    If (Not Observer.Type = ENTITY_TYPE_PLAYER) Then
        Exit Sub
    End If

    'Debug.Print "OnCreateEntity (On Player)", Instigator.Name, Observer.Name

    Select Case Instigator.Type
        Case ENTITY_TYPE_PLAYER
            With UserList(Instigator.Name)
                If Not (.flags.AdminInvisible = 1) Then
                    Call MakeUserChar(False, Instigator.Name, Observer.Name)
                    
                    If .flags.Navegando = 0 Then
                        If UserList(Observer.Name).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then
                            If (MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger = eTrigger.zonaOscura) Or (.flags.invisible Or .flags.Oculto) Then
                                Call WriteSetInvisible(Observer.Name, .Char.CharIndex, True)
                            End If
                        End If
                    End If
                End If
            End With
        Case ENTITY_TYPE_NPC
            With Npclist(Instigator.Name)
                If (MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger <> eTrigger.zonaOscura) Or ((UserList(Observer.Name).flags.Privilegios And (PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin)) <> 0) Then
                    Call MakeNPCChar(False, Observer.Name, Instigator.Name, .Pos.Map, .Pos.X, .Pos.Y)
                End If
            End With
        Case ENTITY_TYPE_OBJECT
            Dim ObjIndex As Integer
            Dim GrhIndex As Integer
            
            Coordinates = Unpack(Instigator.Name)
            
            ObjIndex = MapData(Coordinates.Map, Coordinates.X, Coordinates.Y).ObjInfo.ObjIndex
            GrhIndex = MapData(Coordinates.Map, Coordinates.X, Coordinates.Y).ObjInfo.CurrentGrhIndex
            
            With ObjData(ObjIndex)
                Call WriteObjectCreate(Observer.Name, GrhIndex, Coordinates.X, Coordinates.Y, .Luminous, .LightOffsetX, .LightOffsetY, .LightSize, .CanBeTransparent, .ObjType, GetCreateObjectMetadata(ObjIndex, Coordinates.Map, Coordinates.X, Coordinates.Y))
            End With
            
    End Select

    Exit Sub
    
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub OnCreateEntity de modAreas.bas")
End Sub

Private Sub OnDeleteEntity(ByRef Instigator As Collision.UUID, ByRef Observer As Collision.UUID)
On Error GoTo ErrHandler
    Dim Coordinates As WorldPos
            
    Select Case Observer.Type
        Case ENTITY_TYPE_NPC
            With Npclist(Observer.Name)
                
                Select Case Instigator.Type
                    Case ENTITY_TYPE_PLAYER
                        If (.Target = Instigator.Name) Then
                            .Target = 0
                        End If
                    Case ENTITY_TYPE_NPC
                        If (.TargetNpc = Instigator.Name) Then
                            .TargetNpc = 0
                        End If
                End Select
            End With
            
            Exit Sub
        Case ENTITY_TYPE_OBJECT
            Exit Sub
    End Select

    'Debug.Print "OnDeleteEntity (On Player)", Instigator.Name, Observer.Name

    Select Case Instigator.Type
        Case ENTITY_TYPE_PLAYER
            With UserList(Instigator.Name)
                If .flags.AdminInvisible <> 1 Then
                    Call SendData(ToUser, Observer.Name, PrepareMessageCharacterRemove(.Char.CharIndex))
                End If
            End With
        Case ENTITY_TYPE_NPC
            With Npclist(Instigator.Name)
                Call SendData(ToUser, Observer.Name, PrepareMessageCharacterRemove(.Char.CharIndex))
            End With
        Case ENTITY_TYPE_OBJECT
            Coordinates = Unpack(Instigator.Name)
            Call WriteObjectDelete(Observer.Name, Coordinates.X, Coordinates.Y)
    End Select

    Exit Sub
    
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub OnDeleteEntity de modAreas.bas")
End Sub

Private Sub OnUpdateEntity(ByRef Instigator As Collision.UUID, ByRef Observer As Collision.UUID, ByVal Warped As Boolean)
On Error GoTo ErrHandler

    If (Not Observer.Type = ENTITY_TYPE_PLAYER) Then
        Exit Sub
    End If

    'Debug.Print "OnUpdateEntity (On Player)", Instigator.Name, Observer.Name

    Select Case Instigator.Type
        Case ENTITY_TYPE_PLAYER
            With UserList(Instigator.Name)
                If .flags.AdminInvisible <> 1 Then
                    Call SendData(ToUser, Observer.Name, PrepareMessageCharacterMove(.Char.CharIndex, .Pos.X, .Pos.Y, Warped))
                End If
            End With
        Case ENTITY_TYPE_NPC
            With Npclist(Instigator.Name)
                Call SendData(ToUser, Observer.Name, PrepareMessageCharacterMove(.Char.CharIndex, .Pos.X, .Pos.Y, Warped))
            End With
    End Select
    
    Exit Sub
    
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub OnUpdateEntity de modAreas.bas")
End Sub


Public Function GetCreateObjectMetadata(ByVal ObjIndex As Long, ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte) As Integer
On Error GoTo ErrHandler
    
    If ObjIndex = 0 Then Exit Function
    
    With ObjData(ObjIndex)
        Select Case .ObjType
            Case eOBJType.otResource    ' Is the tile blocked?
                GetCreateObjectMetadata = MapData(Map, X, Y).Blocked
            Case eOBJType.otPuertas
                GetCreateObjectMetadata = ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Cerrada
            Case Else
                GetCreateObjectMetadata = 0
        End Select
    End With
    
Exit Function

ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetObjectMetadata de modAreas.bas")
End Function
