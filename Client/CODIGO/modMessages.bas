Attribute VB_Name = "modMessages"
Option Explicit

Public Enum eMessageId
    None = 0
    Spell_Hits_Npc = 1
    Char_Killed_By_User = 2
    Char_Hit = 3
    Cant_Quit_Paralized = 4
    
    Guild_No_Permission = 5
    
    Guild_Invitation_Limit_Reached = 6
    Guild_Invitation_User_Offline = 7
    Guild_Invitation_User_Already_Has_Guild = 8
    Guild_Invitation_User_In_Other_Faction = 9
    Guild_Invitation_User_Already_Invitated = 10
    Guild_Invitation_Sent = 11
    Guild_Invitation_User_Invitated = 12
End Enum

Public Enum eMessageParameterType
    IsNumber = 0
    IsText = 1
End Enum

Public Type FormattedMessage
    Message As String
    FontType As Integer
    MessageType As Integer
End Type

Public MessageManager As New clsMessageManager

Public Sub ShowConsoleMessage(ByVal MessageId As Integer, ParamArray values() As Variant)
    Call MessageManager.Prepare(MessageId)
    
    Dim I As Integer
    Dim TempString As String
    Dim TempInt As Integer
    
    For I = 0 To UBound(values)
        If VarType(values(I)) = vbString Then
            TempString = values(I)
            Call MessageManager.AddParameterAsText(TempString)
        Else
            TempInt = values(I)
            Call MessageManager.AddParameterAsNumber(TempInt)
        End If
    Next I

    Dim Result As FormattedMessage
    Result = MessageManager.Format()
    
    With FontTypes(Result.FontType)
        Call ShowConsoleMsg(Result.Message, .red, .green, .blue, .bold, .italic, Result.MessageType)
    End With
End Sub
