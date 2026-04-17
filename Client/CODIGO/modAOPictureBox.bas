Attribute VB_Name = "modAOPictureBox"
Option Explicit

Public s_Press_Key_1 As Boolean

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function AOPictureBox_GetOrSetSingleton(ByVal IsAction As Boolean, ByVal Control As AOPictureBox) As AOPictureBox
    Static s_Control As AOPictureBox
    
    Set AOPictureBox_GetOrSetSingleton = IIf(IsAction, Control, s_Control)
    
    If (IsAction) Then
        Set s_Control = Control
    End If
End Function
