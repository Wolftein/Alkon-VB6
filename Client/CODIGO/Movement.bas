Attribute VB_Name = "Movement"
Private Const DIRECTIONS As Byte = 4

Private LastDirection(DIRECTIONS) As Byte


Private Function Exists(ByVal Direction As E_Heading) As Boolean
On Error GoTo ErrHandler

    Dim I As Long
    
    For I = LBound(LastDirection) To UBound(LastDirection)
        If LastDirection(I) = Direction Then
            Exists = True
            Exit Function
        End If
    Next I
    
    Exists = False
    
    Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function Exists de Movement.bas")
End Function

Private Sub Add(ByVal Direction As E_Heading)
On Error GoTo ErrHandler

    Dim I As Long
    
    If Exists(Direction) Then Exit Sub

    For I = LBound(LastDirection) To UBound(LastDirection)
        If LastDirection(I) = 0 Then
            LastDirection(I) = Direction
            Exit Sub
        End If
    Next I
    
    Exit Sub
    
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function Add de Movement.bas")
End Sub

Private Sub Remove(ByVal Direction As E_Heading)
On Error GoTo ErrHandler

    Dim I As Long
    Dim a As Long
    
    For I = LBound(LastDirection) To (UBound(LastDirection) - 1)
        If LastDirection(I) = Direction Then
            For a = I To (UBound(LastDirection) - 1)
                LastDirection(a) = LastDirection(a + 1)
            Next a
            LastDirection(UBound(LastDirection)) = 0
            Exit Sub
        End If
    Next I
    
    If LastDirection(UBound(LastDirection)) = Direction Then LastDirection(UBound(LastDirection)) = 0
    
    Exit Sub

ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function Remove de Movement.bas")
End Sub

Public Function GetDirection() As E_Heading
On Error GoTo ErrHandler

    Dim I As Long

    For I = UBound(LastDirection) To LBound(LastDirection) Step -1
        If LastDirection(I) <> 0 Then
            GetDirection = LastDirection(I)
            Exit Function
        End If
    Next I
    GetDirection = 0
    
    Exit Function
    
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetDirection de Movement.bas")
End Function

Public Sub DirectionKeyDown(ByVal Key As Byte, ByVal Direction As E_Heading)
On Error GoTo ErrHandler

    If GetKeyState(CustomKeys.BindedKey(Key)) < 0 Then
        Call Add(Direction)
    Else
        Call Remove(Direction)
    End If
    
    Exit Sub
    
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function DirectionKeyDown de Movement.bas")
End Sub
