VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMessageManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private CurrentParameters(30) As Variant
Private CurrentParametersCount As Integer
Private CurrentUserIndex As Integer
Private CurrentMessageId As Integer

Public Function Prepare(ByVal UserIndex As Integer, ByVal Id As Integer) As clsMessageManager
    CurrentMessageId = Id
    CurrentParametersCount = 0
    CurrentUserIndex = UserIndex
    Set Prepare = Me
End Function

Public Function AddParameterAsNumber(ByVal Value As Long) As clsMessageManager

    CurrentParameters(CurrentParametersCount) = Value
    
    CurrentParametersCount = CurrentParametersCount + 1
    
    Set AddParameterAsNumber = Me
End Function

Public Function AddParameterAsText(ByRef Value As String) As clsMessageManager

    CurrentParameters(CurrentParametersCount) = Value
    
    CurrentParametersCount = CurrentParametersCount + 1
    
    Set AddParameterAsText = Me
End Function

Public Sub SendToConsole()
    If CurrentMessageId = 0 Then
        Exit Sub
    End If
    
    Call Protocol.WriteConsoleFormattedMessage(CurrentUserIndex, CurrentMessageId, CurrentParameters, CurrentParametersCount)
    
    CurrentMessageId = 0
    CurrentParametersCount = 0
    CurrentUserIndex = 0
End Sub

Public Sub SendToMessageBox()
    If CurrentMessageId = 0 Then
        Exit Sub
    End If
    
    Call Protocol.WriteShowFormattedMessageBox(CurrentUserIndex, CurrentMessageId, CurrentParameters, CurrentParametersCount)
    
    CurrentMessageId = 0
    CurrentParametersCount = 0
    CurrentUserIndex = 0
End Sub
