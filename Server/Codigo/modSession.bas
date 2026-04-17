Attribute VB_Name = "modSession"
'**************************************************************************
'Argentum Online
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
'**************************************************************************

'
' Module author: D'Artagnan (18/06/2014)
'

Option Explicit

Private Const SESSION_LIFE_TIME As Long = 60000  ' 5 minutes
Private Const ACTIVE_SESSIONS_LIMIT As Integer = 1000
Private Const TOKEN_CODE_SIZE As Byte = 6


Public Type SessionCharacter
    charName As String
    CharId As Long
End Type

Private Type Session
    ' General session data.
    ServerTempCode As String
    TokenGeneratedAtTick As Long
    ClientIP As String
    ClientTempCode As String
    ' Account info.
    nAccountID As Long
    sAccountName As String
    AccountEmail As String
    asAccountCharNames(1 To MAX_ACCOUNT_CHARS) As SessionCharacter
    nCharCount As Integer
    Token As String
End Type

Public aActiveSessions() As Session

Public Function GetMaxAllowedSessions() As Integer
    GetMaxAllowedSessions = ServerConfiguration.Session.MaxQuantity
End Function

Public Sub InitializeSessionSystem()
    ReDim aActiveSessions(0 To ServerConfiguration.Session.MaxQuantity) As Session
End Sub

Public Function SessionExpired(ByVal nSessionId As Integer) As Boolean
'******************************************
'Author: D'Artagnan
'Date: 18/06/2014
'Return True if the specified session has expired.
'False otherwise.
'******************************************
On Error GoTo ErrHandler
  
    SessionExpired = (GetTickCount() - aActiveSessions(nSessionId).TokenGeneratedAtTick) > ServerConfiguration.Session.Lifetime
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function SessionExpired de modSession.bas")
End Function

Public Function GetNewSession(ByVal UserIndex As Integer, ByRef ClientTempCode As String, ByRef AccountName As String, ByRef AccountEmail As String) As Integer
On Error GoTo ErrHandler

    Dim SessionPreviousID As Integer
    Dim I As Integer
    
    SessionPreviousID = -1
    
    For I = 0 To GetMaxAllowedSessions()
        With aActiveSessions(I)
            ' Only get expired tokens
            If modSession.SessionExpired(I) Or .TokenGeneratedAtTick = 0 Then
            
                Call modSession.CleanSessionSlot(I)
            
                .ClientIP = UserList(UserIndex).IP
                .ClientTempCode = ClientTempCode
                .ServerTempCode = RandomString(ServerConfiguration.Session.TokenSize)
                .sAccountName = AccountName
                .AccountEmail = AccountEmail
                .TokenGeneratedAtTick = GetTickCount()
                
                SessionPreviousID = I
                Exit For
            End If
        End With
    Next I
    
    GetNewSession = SessionPreviousID
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function SessionGetToken de modSession.bas")
End Function

Public Function ExtendSessionLifetime(ByVal SessionIndex As Integer)
    aActiveSessions(SessionIndex).TokenGeneratedAtTick = GetTickCount()
End Function

Public Function IsTokenValid(ByRef sToken As String, ByVal UserIndex As Integer, Optional ByRef SessionIndex As Integer = -1) As Boolean
'******************************************
'Author: -
'Details: Used to determine if the client is the owner of the session, in order to re-use the slot.
'         The AccountName could be different, but we know it's coming from the same client.
'Date: 18/06/2014
'Return a not-used session ID. Otherwise, the return value is -1.
'******************************************
On Error GoTo ErrHandler:

    Dim SessionId As Integer
    Dim IP As String
    Dim Token As String
    Dim ClientTempCode As String
    Dim ServerTempCode As String
    
    IsTokenValid = False
    
    Dim data() As String
    data = Split(sToken, "-")
    
    ' Not the right amount of parameters in the token
    If UBound(data) <> 3 Then Exit Function
    
    ' Get the data from inside the token.
    SessionId = Val(data(0))
    IP = data(1)
    ClientTempCode = data(2)
    ServerTempCode = data(3)
        
    ' If the token is not a valid index, exit.
    If SessionId < 0 Then Exit Function
    
    SessionIndex = SessionId
    
    With aActiveSessions(SessionId)
        ' The token is valid only if:
        ' * the IP address of the received token matches with the one stored in the session
        ' * The ClientTempCode and ServerTempCode received matches the one stored in the session.
        ' * The IP address in the session is the same as the one from the connected user.
        
        IsTokenValid = (.ClientIP = IP) And (.ClientIP = UserList(UserIndex).IP) And (.ServerTempCode = ServerTempCode) And (.ClientTempCode = ClientTempCode)
                
    End With
            
    Exit Function
    
ErrHandler:
    Call LogError("Error al chequear si el token es válido: " & sToken)
End Function


Public Function RegenerateSessionServerCode(ByVal SessionIndex As Integer) As String
On Error GoTo ErrHandler:

    ' Assign the new token to the session.
    aActiveSessions(SessionIndex).ServerTempCode = RandomString(ServerConfiguration.Session.TokenSize)
    
    ' Return the new token
    RegenerateSessionServerCode = aActiveSessions(SessionIndex).ServerTempCode
            
    Exit Function
    
ErrHandler:
    Call LogError("Error al limpiar el slot de sesión: " & SessionIndex)
End Function

Public Sub SetAccountDataToSession(ByVal SessionIndex As Integer, ByVal AccountId As Integer, ByRef AccountName As String, ByRef AccountEmail As String)
    
    With aActiveSessions(SessionIndex)
        .nAccountID = AccountId
        .sAccountName = AccountName
        .AccountEmail = AccountEmail
    End With
    
End Sub

Public Sub CleanCharactersFromSession(ByVal SessionIndex As Integer)
    Dim I As Integer
    
    With aActiveSessions(SessionIndex)
        For I = 1 To MAX_ACCOUNT_CHARS
            .asAccountCharNames(I).CharId = 0
            .asAccountCharNames(I).charName = vbNullString
        Next I
    End With
End Sub

Public Sub CleanSessionSlot(ByVal SessionIndex As Integer)
    
On Error GoTo ErrHandler:
    Dim I As Integer
    
    ' Validate if the index is between the bounds of session list array
    If SessionIndex < 0 Or SessionIndex > GetMaxAllowedSessions() Then Exit Sub
    
    With aActiveSessions(SessionIndex)
        
        For I = 1 To MAX_ACCOUNT_CHARS
            .asAccountCharNames(I).CharId = 0
            .asAccountCharNames(I).charName = vbNullString
        Next I
        
        .ClientIP = vbNullString
        .ClientTempCode = vbNullString
        .TokenGeneratedAtTick = 0
        .nAccountID = 0
        .nCharCount = 0
        .sAccountName = vbNullString
        .AccountEmail = vbNullString
        .ServerTempCode = vbNullString
        .Token = vbNullString
    End With
        
    Exit Sub
    
ErrHandler:
    Call LogError("Error al limpiar el slot de sesión: " & SessionIndex)
End Sub

Public Sub RecreateToken(ByVal SessionIndex As Integer)
On Error GoTo ErrHandler:

    With aActiveSessions(SessionIndex)
        .Token = SessionIndex & "-" & .ClientIP & "-" & .ClientTempCode & "-" & .ServerTempCode
    End With
       
    Exit Sub
ErrHandler:
    Call LogError("Error al generar el token string: " & SessionIndex)
End Sub

Public Sub SessionRemove(ByVal nSessionId As Integer, Optional ByVal bExpired As Boolean = False)
'******************************************
'Author: D'Artagnan
'Date: 18/06/2014
'Remove the session and free the specified slot.
'******************************************
On Error GoTo ErrHandler
  
    Dim emptySession As Session
    Dim sToken As String
    
    sToken = aActiveSessions(nSessionId).ServerTempCode
    aActiveSessions(nSessionId) = emptySession
    With aActiveSessions(nSessionId)
        .ClientIP = vbNullString
        .ClientTempCode = vbNullString
        .TokenGeneratedAtTick = -1
        .nAccountID = -1
        .nCharCount = 0
        .sAccountName = vbNullString
        .AccountEmail = vbNullString
        .ServerTempCode = vbNullString
        'Erase .asAccountCharNames
        .Token = vbNullString
    End With
    
    ' Keep token.
    If Not bExpired Then
        aActiveSessions(nSessionId).ServerTempCode = sToken
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SessionRemove de modSession.bas")
End Sub

Public Sub SessionRemoveAccountCharName(ByVal nSessionId, ByVal nCharSlot As Integer)
'******************************************
'Author: D'Artagnan
'Date: 18/06/2014
'Remove the character data at the specified slot.
'Remaining charaters will be reordered to fill the free slot.
'******************************************
On Error GoTo ErrHandler
  
    Dim I As Integer
    Dim SessionCharacter As SessionCharacter
    
    ' Erase data at specified slot.
    Call SessionSetAccountCharacterData(nSessionId, nCharSlot, vbNullString, -1)
    
    ' Reorder array if necessary.
    If nCharSlot < aActiveSessions(nSessionId).nCharCount Then
        For I = nCharSlot + 1 To aActiveSessions(nSessionId).nCharCount
            ' Get character data at the current slot.
            SessionCharacter = aActiveSessions(nSessionId).asAccountCharNames(I)
            
            If SessionCharacter.charName <> vbNullString Then
                ' Store it in the previous slot.
                Call SessionSetAccountCharacterData(nSessionId, I - 1, SessionCharacter.charName, SessionCharacter.CharId)
                ' Delete remaining slot.
                Call SessionSetAccountCharacterData(nSessionId, I, vbNullString, -1)
            End If
        Next I
    End If
    
    aActiveSessions(nSessionId).nCharCount = aActiveSessions(nSessionId).nCharCount - 1
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SessionRemoveAccountCharName de modSession.bas")
End Sub

Public Sub RemoveExpiredSessions()
'******************************************
'Author: D'Artagnan
'Date: 18/06/2014
'Remove the expired sessions from the list.
'******************************************
On Error GoTo ErrHandler
  
    Dim I As Integer
    
    For I = 0 To GetMaxAllowedSessions()
        If aActiveSessions(I).TokenGeneratedAtTick > 0 And modSession.SessionExpired(I) Then
            Call modSession.CleanSessionSlot(I)
            'Call SessionRemove(I, True)
        End If
    Next I
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SessionRemoveExpiredSessions de modSession.bas")
End Sub

Public Sub SessionSetAccountCharacterData(ByVal nSessionId As Integer, ByVal nCharSlot As Integer, _
                                     ByRef sCharName As String, ByVal CharId As Long)
'******************************************
'Author: D'Artagnan
'Date: 18/06/2014
'Set (or replace) a character name at the specified slot.
'******************************************
On Error GoTo ErrHandler
  
    With aActiveSessions(nSessionId)
        If LenB(sCharName) > 0 And nCharSlot > .nCharCount Then
            .nCharCount = nCharSlot
        End If
        .asAccountCharNames(nCharSlot).charName = sCharName
        .asAccountCharNames(nCharSlot).CharId = CharId
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SessionSetAccountCharacterData de modSession.bas")
End Sub
