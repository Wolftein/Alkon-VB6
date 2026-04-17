Attribute VB_Name = "modIntervals"
Public TickCount As Long

Public Function SetIntervalEnd(ByVal Interval As Long) As Long
On Error GoTo ErrHandler

    SetIntervalEnd = TickCount + Interval
    Exit Function
    
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function SetIntervalEnd de modIntervals.bas. IntervalEnd: " & (TickCount + Interval))
End Function

Public Function GetIntervalRemainingTime(ByVal EndTick As Long) As Long
On Error GoTo ErrHandler

    GetIntervalRemainingTime = EndTick - TickCount
    Exit Function

ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetIntervalRemainingTime de modIntervals.bas")
End Function

Public Function IsIntervalReached(ByVal EndTick As Long) As Boolean
On Error GoTo ErrHandler

    IsIntervalReached = GetIntervalRemainingTime(EndTick) <= 0
    Exit Function
    
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function IsIntervalReached de modIntervals.bas")
End Function

Public Function FromSeconds(ByVal Seconds As Long) As Long
    FromSeconds = Seconds * 1000
End Function

Public Function FromMinutes(ByVal Minutes As Long) As Long
    FromMinutes = FromSeconds(Minutes * 60)
End Function
