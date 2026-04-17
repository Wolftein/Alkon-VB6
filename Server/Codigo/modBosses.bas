Attribute VB_Name = "modBosses"

Option Explicit

Public Type tBossData
    NpcIndex As Integer
    SpawnPosQty As Byte
    SpawnPos() As WorldPos
    Maps() As Integer
    Minions() As Integer
    MinAmount As Integer
    MaxAmount As Integer
    Amount As Integer
    Alive As Boolean
    CurAmount As Integer
    SpawnOnStartup As Boolean
End Type

Public BossData() As tBossData

Public Sub LoadBossData()
On Error GoTo ErrHandler

    Dim I As Integer
    Dim a As Byte
    Dim NumBosses As Integer
    Dim Tmp As String
    Dim TmpArray() As String
    
    NumBosses = Val(GetVar(DatPath & "Bosses.dat", "INIT", "NumBosses"))
    ReDim BossData(1 To NumBosses) As tBossData
    
    For I = 1 To NumBosses
        With BossData(I)
            .NpcIndex = Val(GetVar(DatPath & "Bosses.dat", "Boss" & I, "NpcIndex"))
            
            .SpawnPosQty = Val(GetVar(DatPath & "Bosses.dat", "Boss" & I, "SpawnPosQty"))
            
            If .SpawnPosQty > 0 Then
                ReDim .SpawnPos(1 To .SpawnPosQty)
                                
                For A = 1 To .SpawnPosQty
                    Tmp = GetVar(DatPath & "Bosses.dat", "Boss" & I, "SpawnPos" & A)
                    .SpawnPos(A).Map = Val(ReadField(1, Tmp, Asc("-")))
                    .SpawnPos(A).X = Val(ReadField(2, Tmp, Asc("-")))
                    .SpawnPos(A).Y = Val(ReadField(3, Tmp, Asc("-")))
                Next A
                
            Else
                Erase .SpawnPos
            End If
            
            
            TmpArray = Split(GetVar(DatPath & "Bosses.dat", "Boss" & I, "Maps"), "-")
            ReDim .Maps(0 To UBound(TmpArray)) As Integer
            For a = 0 To UBound(TmpArray)
                .Maps(a) = Val(TmpArray(a))
            Next a
            TmpArray = Split(GetVar(DatPath & "Bosses.dat", "Boss" & I, "Minions"), "-")
            ReDim .Minions(0 To UBound(TmpArray)) As Integer
            For a = 0 To UBound(TmpArray)
                .Minions(a) = Val(TmpArray(a))
            Next a
            Tmp = GetVar(DatPath & "Bosses.dat", "Boss" & I, "Amount")
            .MinAmount = Val(ReadField(1, Tmp, Asc("-")))
            .MaxAmount = Val(ReadField(2, Tmp, Asc("-")))
            .Amount = RandomNumber(.MinAmount, .MaxAmount)
            .SpawnOnStartup = CBool(Val(GetVar(DatPath & "Bosses.dat", "Boss" & I, "SpawnOnStartup")))
            
            If .SpawnOnStartup Then
                Call modBosses.SpawnBoss(I)
            End If
            
        End With
    Next I
    
    Exit Sub
    
ErrHandler:
    Call LogError("Error " & Err.Number & " (" & Err.Description & ") en Sub LoadBossData del Módulo modBosses")
End Sub

Public Sub RestartBossSpawn(ByVal BossIndex As Byte)
On Error GoTo ErrHandler
    Dim Amount As Integer
    
    With BossData(BossIndex)
        .Amount = RandomNumber(.MinAmount, .MaxAmount)
        .Alive = False
    End With
    
    Exit Sub
    
ErrHandler:
    Call LogError("Error " & Err.Number & " (" & Err.Description & ") en Sub RestartBossSpawn del Módulo modBosses")
End Sub

Public Sub CheckBossSpawn(ByVal NpcIndex As Integer)
On Error GoTo ErrHandler
    Dim I As Byte
    Dim BossIndex As Byte
    
    If NpcIndex <= 0 Then Exit Sub
    
    BossIndex = GetBossIndexFromMap(Npclist(NpcIndex).Pos.Map)
    If BossIndex <= 0 Then Exit Sub
    
    With BossData(BossIndex)
        For I = 0 To UBound(.Minions)
            If .Minions(I) = Npclist(NpcIndex).Numero Then
                If .CurAmount >= .Amount Then
                    Call SpawnBoss(BossIndex)
                    Exit Sub
                Else
                    If .Alive = False Then _
                        .CurAmount = .CurAmount + 1
                    Exit Sub
                End If
            End If
        Next I
    End With
    
    Exit Sub
    
ErrHandler:
    Call LogError("Error " & Err.Number & " (" & Err.Description & ") para Npc: " & NpcIndex & " en Sub CheckBossSpawn del Módulo modBosses")
End Sub

Public Sub SpawnBoss(ByVal BossIndex As Integer)
    
    If BossIndex > UBound(BossData) Or BossIndex < 1 Then
        Exit Sub
    End If
    
    Dim BossNpcIndex As Integer
    Dim BossText As String
    Dim I As Integer
    Dim SpawnPosIndex As Byte
    
    With BossData(BossIndex)
    
        ' We can't spawn the same boss again.
        If BossData(BossIndex).Alive = True Then Exit Sub
        
        ' If therre's no spawn point configured, then exit
        If .SpawnPosQty < 1 Then Exit Sub
        
        SpawnPosIndex = RandomNumber(1, .SpawnPosQty)
        
        BossNpcIndex = SpawnNpc(.NpcIndex, .SpawnPos(SpawnPosIndex), False, False)
        
        BossText = "El jefe " & Npclist(BossNpcIndex).Name & " ha sido invocado"
        
        ' Send the Shake effect to the users located in the map where the boss should spawn.
        Call SendData(toMap, .SpawnPos(SpawnPosIndex).Map, PrepareStartEffect(ClientPresentEffects.BossAppears, BossText))
        
        ' Send the Shake effect to the users in the maps where the NPCs should be killed
        For I = 0 To UBound(.Maps)
            Call SendData(toMap, .Maps(I), PrepareStartEffect(ePresentEffect.SpawnBoss, BossText))
        Next I

        Npclist(BossNpcIndex).flags.Boss = BossIndex
        .Alive = True
        
        .CurAmount = 0
    End With
    
End Sub

Public Function GetBossIndexFromMap(ByVal Map As Integer) As Byte
On Error GoTo ErrHandler
    Dim I As Byte
    Dim MapI As Byte
    
    For I = 1 To UBound(BossData)
        With BossData(I)
            For MapI = 0 To UBound(.Maps)
                If .Maps(MapI) = Map Then
                    GetBossIndexFromMap = I
                    Exit Function
                End If
            Next MapI
        End With
    Next I
    
    Exit Function
    
ErrHandler:
    Call LogError("Error " & Err.Number & " (" & Err.Description & ") en Function GetBossIndexFromMap del Módulo modBosses")
End Function
