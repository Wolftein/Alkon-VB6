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

Private Type ReplaceInfo
    Literal As String
    ParameterNumber As Integer
    IsConditional As Boolean
    Singular As String
    Plural As String
End Type

Private Type MessageFormatInfo
    Message As String
    
    ReplaceInfos() As ReplaceInfo
    ReplaceInfosCount As Integer

    NeedToBeFormatted As Boolean
    NeedToBeEscaped As Boolean
    
    MessageType As Integer
    FontType As Integer
End Type

Private Type MessageParameterInfo
    TextValue As String
    IsSingular As Boolean
End Type

Private MaxMessages As Integer

Private MessageFormatInfos() As MessageFormatInfo

Private CurrentParameters(30) As MessageParameterInfo
Private CurrentParametersCount As Integer
Private CurrentMessageId As Integer

Public Sub LoadFromIniFile(ByRef IniFile As String)
On Error GoTo ErrHandler
  
    Dim I As Integer
    Dim Message As String
    
    Dim FontType As String
    Dim MessageType As String
    
    Dim IniManager As New clsIniManager

    Call IniManager.Initialize(IniFile)
    
    MaxMessages = CInt(Val(IniManager.GetValue("INIT", "MaxMessages")))
    ReDim MessageFormatInfos(MaxMessages + 1)
    
    For I = 1 To MaxMessages
        Message = IniManager.GetValue("MESSAGE" & I, "Message")
        
        If Message <> vbNullString Then
            FontType = IniManager.GetValue("MESSAGE" & I, "FontType")
            MessageType = IniManager.GetValue("MESSAGE" & I, "MessageType")
            
            Call Add(I, Message, FontType, MessageType)
        End If
        
    Next I
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadFromIniFile de clsMessageManager.bas")
End Sub

Private Function GetFontTypeFromText(ByRef FontTypeText As String) As FontTypeNames
    Select Case UCase$(FontTypeText)
        Case "TALK"
            GetFontTypeFromText = FontTypeNames.FONTTYPE_TALK
        Case "FIGHT"
            GetFontTypeFromText = FontTypeNames.FONTTYPE_FIGHT
        Case "WARNING"
            GetFontTypeFromText = FontTypeNames.FONTTYPE_WARNING
        Case "INFO"
            GetFontTypeFromText = FontTypeNames.FONTTYPE_INFO
        Case "INFOBOLD"
            GetFontTypeFromText = FontTypeNames.FONTTYPE_INFOBOLD
        Case "EJECUCION"
            GetFontTypeFromText = FontTypeNames.FONTTYPE_EJECUCION
        Case "PARTY"
            GetFontTypeFromText = FontTypeNames.FONTTYPE_PARTY
        Case "VENENO"
            GetFontTypeFromText = FontTypeNames.FONTTYPE_VENENO
        Case "GUILD"
            GetFontTypeFromText = FontTypeNames.FONTTYPE_GUILD
        Case "SERVER"
            GetFontTypeFromText = FontTypeNames.FONTTYPE_SERVER
        Case "GUILDMSG"
            GetFontTypeFromText = FontTypeNames.FONTTYPE_GUILDMSG
        Case "CONSEJO"
            GetFontTypeFromText = FontTypeNames.FONTTYPE_CONSEJO
        Case "CONSEJOCAOS"
            GetFontTypeFromText = FontTypeNames.FONTTYPE_CONSEJOCAOS
        Case "CONSEJOVESA"
            GetFontTypeFromText = FontTypeNames.FONTTYPE_CONSEJOVesA
        Case "CONSEJOCAOSVESA"
            GetFontTypeFromText = FontTypeNames.FONTTYPE_CONSEJOCAOSVesA
        Case "CENTINELA"
            GetFontTypeFromText = FontTypeNames.FONTTYPE_CENTINELA
        Case "GMMSG"
            GetFontTypeFromText = FontTypeNames.FONTTYPE_GMMSG
        Case "GM"
            GetFontTypeFromText = FontTypeNames.FONTTYPE_GM
        Case "CITIZEN"
            GetFontTypeFromText = FontTypeNames.FONTTYPE_CITIZEN
        Case "CONSE"
            GetFontTypeFromText = FontTypeNames.FONTTYPE_CONSE
        Case "DIOS"
            GetFontTypeFromText = FontTypeNames.FONTTYPE_DIOS
        Case "NPCNAME"
            GetFontTypeFromText = FontTypeNames.FONTTYPE_NPCNAME
        Case Else
            GetFontTypeFromText = FontTypeNames.FONTTYPE_INFO
            Call LogError("Warning: el FontType " + FontTypeText + " no existe .Function GetFontTypeFromText de clsMessageManager.bas")
    End Select
    
End Function

Private Function GetMessageTypeFromText(ByRef MessageTypeText As String) As eMessageType
    Select Case UCase$(MessageTypeText)
        Case "NONE"
            GetMessageTypeFromText = eMessageType.None
        Case "INFO"
            GetMessageTypeFromText = eMessageType.Info
        Case "ADMIN"
            GetMessageTypeFromText = eMessageType.Admin
        Case "GUILD"
            GetMessageTypeFromText = eMessageType.Guild
        Case "PARTY"
            GetMessageTypeFromText = eMessageType.Party
        Case "COMBATE"
            GetMessageTypeFromText = eMessageType.combate
        Case "TRABAJO"
            GetMessageTypeFromText = eMessageType.Trabajo
        Case "M_MOTD"
            GetMessageTypeFromText = eMessageType.m_MOTD
        Case Else
            GetMessageTypeFromText = eMessageType.None
            Call LogError("Warning: el MessageType " + MessageTypeText + " no existe .Function GetMessageTypeFromText de clsMessageManager.bas")
    End Select
End Function

Private Sub Add(ByVal Id As Integer, ByRef Message As String, ByRef FontTypeText As String, ByRef MessageTypeText As String)
    If Id > MaxMessages Then
        Call Err.Raise(-1, "clsMessangerManager", "the id(" & Id & ") can't be more than " & (MaxMessages - 1))
        Exit Sub
    End If
    If Message = vbNullString Then
        Call Err.Raise(-1, "clsMessangerManager", "the message can't be empty")
        Exit Sub
    End If
    
    MessageFormatInfos(Id) = Parse(Message)
    
    MessageFormatInfos(Id).FontType = GetFontTypeFromText(FontTypeText)
    MessageFormatInfos(Id).MessageType = GetMessageTypeFromText(MessageTypeText)
    
End Sub

Public Function Prepare(ByVal Id As Integer) As clsMessageManager
    CurrentMessageId = Id
    CurrentParametersCount = 0
    Set Prepare = Me
End Function

Public Function AddParameterAsNumber(ByVal ParameterValue As Long) As clsMessageManager

    With CurrentParameters(CurrentParametersCount)
        .IsSingular = ParameterValue = 1
        .TextValue = ParameterValue
    End With
    
    CurrentParametersCount = CurrentParametersCount + 1
    Set AddParameterAsNumber = Me
End Function

Public Function AddParameterAsText(ByRef ParameterValue As String) As clsMessageManager

    With CurrentParameters(CurrentParametersCount)
        .IsSingular = 1
        .TextValue = ParameterValue
    End With
    
    CurrentParametersCount = CurrentParametersCount + 1
    Set AddParameterAsText = Me
End Function

Friend Function Format() As FormattedMessage
    If Not IsMessageIdValid(CurrentMessageId) Then
        Format.Message = "MESSAGE_NOT_FOUND " & CurrentMessageId
        
        Format.FontType = FontTypeNames.FONTTYPE_INFO
        Format.MessageType = eMessageType.Info
    Else
        Format.Message = ApplyFormat(MessageFormatInfos(CurrentMessageId), CurrentParameters, CurrentParametersCount)
        
        Format.FontType = MessageFormatInfos(CurrentMessageId).FontType
        Format.MessageType = MessageFormatInfos(CurrentMessageId).MessageType
    End If
End Function
Private Function IsMessageIdValid(ByVal MessageId As Integer) As Boolean
    If MessageId > MaxMessages Or MessageId <= 0 Then
        IsMessageIdValid = False
        Exit Function
    End If
    If MessageFormatInfos(MessageId).Message = vbNullString Then
        IsMessageIdValid = False
        Exit Function
    End If
    
    IsMessageIdValid = True
End Function
Private Function ApplyFormat(ByRef MFI As MessageFormatInfo, ByRef Parameters() As MessageParameterInfo, ByVal ParametersCount As Integer) As String
    Dim Output As String
    
    Output = MFI.Message
    If MFI.NeedToBeFormatted Then
        Dim Index As Integer
        
        For Index = 0 To MFI.ReplaceInfosCount - 1
            With MFI.ReplaceInfos(Index)
                If ParametersCount < .ParameterNumber + 1 Then
                    Output = Replace$(Output, .Literal, "{NOT_FOUND " & .ParameterNumber & "}")
                Else
                    If .IsConditional Then
                        If Parameters(.ParameterNumber).IsSingular Then
                            Output = Replace$(Output, .Literal, .Singular)
                        Else
                            Output = Replace$(Output, .Literal, .Plural)
                        End If
                    Else
                        Output = Replace$(Output, .Literal, Parameters(.ParameterNumber).TextValue)
                    End If
                End If
            End With
        Next
    End If
    
    If MFI.NeedToBeEscaped Then
        Output = Replace$(Output, "\{", "{")
        Output = Replace$(Output, "\}", "}")
        Output = Replace$(Output, "\\", "\")
        Output = Replace$(Output, "\n", vbCrLf)
    End If
    
    ApplyFormat = Output
End Function

Private Function Parse(ByRef MessageToFormat As String) As MessageFormatInfo
    Dim MFI As MessageFormatInfo
    
    MFI.Message = MessageToFormat

    Dim Index As Integer
    
    Dim Temp As String
    Dim TempLiteral As String
    Dim ParameterNumber As Integer
    Dim Singular As String
    Dim Plural As String
            
    Dim StartIndex As Integer
    Dim EndIndex As Integer
    
    Dim MessageChar As String
    
    Index = 1
    
    Do
    
    MessageChar = mid$(MessageToFormat, Index, 1)
    
    Select Case MessageChar
        Case "\"
            Index = Index + 1
            MFI.NeedToBeEscaped = True
        Case "{"
            StartIndex = Index
            Temp = GetSlice(MessageToFormat, StartIndex + 1, "}")
            Index = Index + Len(Temp) + 1
            ParameterNumber = CInt(Temp)
            EndIndex = Index + 1

            TempLiteral = mid$(MessageToFormat, StartIndex, EndIndex - StartIndex)
            If ExistOnAll(MFI, TempLiteral) = False Then
                Call AddParameterToReplace(MFI, ParameterNumber, TempLiteral)
            End If
            
        Case "["
            StartIndex = Index
            Temp = GetSlice(MessageToFormat, Index + 1, "]")
            Index = Index + Len(Temp) + 1
            ParameterNumber = CInt(Temp)
            
            Temp = GetSlice(MessageToFormat, Index + 2, "]")
            Index = Index + Len(Temp) + 2
            Call SplitBy(Temp, "|", Singular, Plural)
            EndIndex = Index + 1
            
            TempLiteral = mid$(MessageToFormat, StartIndex, EndIndex - StartIndex)
            If ExistOnAll(MFI, TempLiteral) = False Then
                Call AddConditionalToReplace(MFI, ParameterNumber, TempLiteral, Singular, Plural)
            End If
            
    End Select
    
    Index = Index + 1
    Loop While (Index <= Len(MessageToFormat))
    
    MFI.NeedToBeFormatted = MFI.ReplaceInfosCount <> 0
    
    If MFI.ReplaceInfosCount Then
        ReDim Preserve MFI.ReplaceInfos(MFI.ReplaceInfosCount - 1)
    End If
    
    Parse = MFI
End Function
Private Function ExistOnAll(ByRef MFI As MessageFormatInfo, ByRef Literal As String) As Boolean
    Dim Index As Integer
    
    For Index = 0 To MFI.ReplaceInfosCount - 1
        With MFI.ReplaceInfos(Index)
            If StrComp(.Literal, Literal, vbBinaryCompare) = 0 Then
                ExistOnAll = True
                Exit Function
            End If
        End With
     Next
     
     ExistOnAll = False
End Function
Private Sub AddParameterToReplace(ByRef MFI As MessageFormatInfo, ByVal ParameterNumber As Integer, ByRef Literal As String)
     Dim Data As ReplaceInfo
     
     With Data
        .ParameterNumber = ParameterNumber
        .Literal = Literal
        .IsConditional = False
     End With
     
    Call AddReplace(MFI, Data)
End Sub
Private Sub AddConditionalToReplace(ByRef MFI As MessageFormatInfo, ByVal ParameterNumber As Integer, ByRef Literal As String, ByVal Singular As String, ByVal Plural As String)
     Dim Data As ReplaceInfo
     
     With Data
        .ParameterNumber = ParameterNumber
        .Literal = Literal
        .IsConditional = True
        .Singular = Singular
        .Plural = Plural
     End With
     
    Call AddReplace(MFI, Data)
End Sub

Private Sub AddReplace(ByRef MFI As MessageFormatInfo, ByRef Data As ReplaceInfo)
    MFI.ReplaceInfosCount = MFI.ReplaceInfosCount + 1
    ReDim Preserve MFI.ReplaceInfos(MFI.ReplaceInfosCount)
    MFI.ReplaceInfos(MFI.ReplaceInfosCount - 1) = Data
End Sub

Private Function GetSlice(ByRef Source As String, ByVal Index As Integer, ByRef CharEnd As String) As String
    Dim CharIndex As Integer
    
    CharIndex = InStr(Index, Source, CharEnd)
    
    GetSlice = mid(Source, Index, CharIndex - Index)
    
End Function

Private Sub SplitBy(ByRef Source As String, ByRef SplitChar As String, ByRef Left As String, ByRef Right As String)
    Dim CharIndex As Integer
    
    CharIndex = InStr(1, Source, SplitChar)
    
    Left = mid$(Source, 1, CharIndex - 1)
    Right = mid$(Source, CharIndex + 1, Len(Source) - (CharIndex - 1))
    
End Sub
