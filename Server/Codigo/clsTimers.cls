VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTimers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CANTIDADTIMERS As Byte = 5

Private Type Timer
    Interval As Integer
    CurrentTick As Long
    startTick As Long
    Run As Boolean
End Type

Dim Timer(1 To CANTIDADTIMERS) As Timer

Public Enum TimersIndex
    Walk = 1
    Hit = 2
    MoveAttack = 3
End Enum

Public Sub Initialize(ByVal NpcIndex As Integer)
On Error GoTo ErrHandler
  
    Dim I As Byte
    
    Call SetInterval(1, Npclist(NpcIndex).Intervalos.Walk)
    Call SetInterval(2, Npclist(NpcIndex).Intervalos.Hit)
    Call SetInterval(3, Npclist(NpcIndex).Intervalos.MoveAttack)
    
    For I = 1 To CANTIDADTIMERS
        If Timer(I).Interval > 0 Then Call Start(I)
    Next I
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Initialize de clsTimers.cls")
End Sub

''
' Sets a new interval for a timer.
'
' @param TimerIndex Timer´s Index
' @param Interval New lenght for the Timer´s cicle in miliseconds.
' @remarks  Must be done after creating the timer and before using it, otherwise, Interval will be 0

Public Sub SetInterval(ByVal TimerIndex As TimersIndex, ByVal Interval As Integer)
'*************************************************
'Author: Nacho (Integer)
'Last modified:
'Desc: Sets a new interval for a timer.
'*************************************************
On Error GoTo ErrHandler
  
    If TimerIndex < 1 Or TimerIndex > CANTIDADTIMERS Then Exit Sub
    
    Timer(TimerIndex).Interval = Interval
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SetInterval de clsTimers.cls")
End Sub

''
' Retrieves the interval of a timer.
'
' @param TimerIndex Timer´s Index
'
' @return The requested timer's interval.

Public Function GetInterval(ByVal TimerIndex As TimersIndex) As Integer
'*************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last modified: 13/10/2008
'Desc: Retrieves the interval of a timer.
'*************************************************
On Error GoTo ErrHandler
  
    If TimerIndex < 1 Or TimerIndex > CANTIDADTIMERS Then Exit Function
    
    GetInterval = Timer(TimerIndex).Interval
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetInterval de clsTimers.cls")
End Function

''
' Starts the timer.
'
' @param TimerIndex Timer´s Index
' @remarks  Do this for starting the timer.

Public Sub Start(ByVal TimerIndex As TimersIndex)
'*************************************************
'Author: Nacho (Integer)
'Last modified:
'Desc: Starts running the Timer.
'*************************************************
On Error GoTo ErrHandler
  
    If TimerIndex < 1 Or TimerIndex > CANTIDADTIMERS Then Exit Sub
    
    Timer(TimerIndex).Run = True 'We allow the Timer to run.
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Start de clsTimers.cls")
End Sub

''
' Resume the timer form where we had stoped.
'
' @param TimerIndex Timer´s Index
' @remarks  Makes the Timer continue from where it had stoped.

Public Sub Continue(ByVal TimerIndex As TimersIndex)
'*************************************************
'Author: Nacho (Integer)
'Last modified:
'Desc: Continue our timer from where we had stopped it.
'*************************************************
On Error GoTo ErrHandler
  
    If TimerIndex < 1 Or TimerIndex > CANTIDADTIMERS Then Exit Sub
    
    Timer(TimerIndex).Run = True 'We allow the Timer to run.
    
    'We´ll set the StartTick again
    Timer(TimerIndex).startTick = Timer(TimerIndex).startTick + Timer(TimerIndex).CurrentTick
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Continue de clsTimers.cls")
End Sub

''
' Pauses the Timer
'
' @param TimerIndex Timer´s Index
' @remarks  Can be resumed with Continue()

Public Sub Pause(ByVal TimerIndex As TimersIndex)
'*************************************************
'Author: Nacho (Integer)
'Last modified:
'Desc: Pauses the timer.
'*************************************************
On Error GoTo ErrHandler
  
    If TimerIndex < 1 Or TimerIndex > CANTIDADTIMERS Then Exit Sub
    
    Timer(TimerIndex).Run = False 'Don´t allow the timer to run.
    
    Timer(TimerIndex).CurrentTick = GetTickCount() - Timer(TimerIndex).startTick 'Calculates acutual tick
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Pause de clsTimers.cls")
End Sub

''
' Check if the timer has already completed it´s cicle.
'
' @param TimerIndex Timer´s Index
' @param Restart If true, restart if we timer has ticked
' @return True if the interval has already been elapsed
' @remarks  Can´t be done if the timer is stoped or if it had never been started.

Public Function Check(ByVal TimerIndex As TimersIndex, Optional Restart As Boolean = True) As Boolean
'*************************************************
'Author: Nacho Agustin (Integer)
'Last modified: 08/26/06
'Modification: NIGO: Added Restart as boolean
'Desc: Checks if the Timer has alredy "ticked"
'Returns: True if it has ticked, False if not.
'*************************************************
On Error GoTo ErrHandler
  

    If TimerIndex < 1 Or TimerIndex > CANTIDADTIMERS Then Exit Function
    
    If Not Timer(TimerIndex).Run Then Exit Function 'If we+ weren´t allowed to run, exit function
    
    Timer(TimerIndex).CurrentTick = GetTickCount() - Timer(TimerIndex).startTick 'Calcutates CurrentTick
    
    If Timer(TimerIndex).CurrentTick >= Timer(TimerIndex).Interval Then
        Check = True 'We have Ticked!
        If Restart Then
            Timer(TimerIndex).startTick = GetTickCount() 'Restart Timer (Nicer than calling Restart() )
        End If
    End If
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function Check de clsTimers.cls")
End Function

''
' Restarts timer.
'
' @param TimerIndex Timer´s Index

Public Sub Restart(ByVal TimerIndex As TimersIndex)
'*************************************************
'Author: Nacho Agustin (Integer)
'Last modified:
'Desc: Restarts timer
'*************************************************
On Error GoTo ErrHandler
  
    If TimerIndex < 1 Or TimerIndex > CANTIDADTIMERS Then Exit Sub
    
    Timer(TimerIndex).startTick = GetTickCount()
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Restart de clsTimers.cls")
End Sub

