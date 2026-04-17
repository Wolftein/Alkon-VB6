Attribute VB_Name = "modMessageQueueProxy"
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
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
Option Explicit

Public MQReceivedDataBuffer As New clsByteQueue
Public MQOutboundByteQueue As New clsByteQueue

Public Enum eProxyServerResponses
    KICK_CHAR = 1
    CHANGE_HEAD = 2
    CHANGE_GENDER = 3
End Enum

Public Enum eProxyServerMessages
    ACCOUNT_CREATE = 1
    ACCOUNT_RECOVER = 2
    ACCOUNT_PASSWORD_CHANGED = 3
    DEATH_EVENT_USER_KILLED_USER = 4
End Enum

Public Function HandleMQReceiverMessage(ByRef buffer As clsByteQueue) As Boolean


    Select Case buffer.PeekByte
        Case eProxyServerResponses.KICK_CHAR ' Kick character from the game
            Call buffer.ReadByte
            
            ' TODO: Place the logic here
            'Call HandleKick(CInt(msgSpltd(1)))
       Case eProxyServerResponses.CHANGE_HEAD  ' Change character Head
            
            Call buffer.ReadByte
            
            ' TODO: Place the logic here
            'Call HandleChangeHead(msgSpltd(1), msgSpltd(2))
        Case eProxyServerResponses.CHANGE_GENDER 'Change Character Gender
            Call buffer.ReadByte
            
            ' TODO: Place the logic here
            'Call HandleChangeGender(msgSpltd(1), msgSpltd(2))
    End Select
    
    HandleMQReceiverMessage = buffer.length <> 0
    
        'Case "KU" ' Kick a char
        
        'Case "CH" ' Change Head
        
        'Case "CG" ' Change Gender

        'Case "CS" ' Closes the socket connection by request
        '    Call CloseConnection(socketIndex)

End Function


Private Sub CloseConnection(ByVal SocketIndex As Integer)
On Error GoTo ErrHandler
  
    frmMain.sckMQReceiver(SocketIndex).Close
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CloseConnection de modMQ.bas")
End Sub

Private Sub HandleKick(ByVal CharId As Long)
'***************************************************
'Author: Sergio Alejandro Masolini (Nightw)
'Last Modification: -
'If user is invisible, it automatically becomes
'visible before doing the countdown to exit
'***************************************************
On Error GoTo ErrHandler
  
    Dim I As Integer
    Dim UserIndex As Integer
    Dim userOnline As Boolean
    userOnline = False
    For I = 1 To UBound(UserList)
        If UserList(I).ID = CharId Then
            UserIndex = I
            userOnline = True
            Exit For
        End If
    Next I
    
    If UserIndex = 0 Then
        Exit Sub
    End If
    
    With UserList(UserIndex)
        
        Call WriteConsoleMsg(UserIndex, "Estás siendo expulsado por una solicitud desde la web.", FontTypeNames.FONTTYPE_WARNING)
        
        If .flags.Paralizado <> 1 Then
            ' Do shit here.
        End If

        Call ExitSecureCommerce(UserIndex)
        
        ' Remove token once used.
        ' I think this shouldn't be used because there's no session active
        ' for the user when he's online. The session is generated after he logs out.
        ' Kicking the user from the website should be also kicking him out from the account
        'Call SessionRemove(.nSessionId)
        
        .bForceCloseAccount = True
        .bShowAccountForm = False
        
        Call Cerrar_Usuario(UserIndex)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleKick de modMQ.bas")
End Sub

Private Sub HandleChangeHead(ByRef CharId As String, ByRef headNum As String)
'***************************************************
'Author: Lucas Ezequiel Figelj (Luke)
'Last Modification: -
'Changes user head from Web Request
'***************************************************
On Error GoTo ErrHandler
  
    Dim UserID As String
    Dim head As String
    Dim UserIndex As Integer
    Dim userOnline As Boolean
    Dim I As Integer
    
    userOnline = False
    
    For I = 1 To UBound(UserList)
        If UserList(I).ID = CharId Then
            UserIndex = I
            userOnline = True
            Exit For
        End If
    Next I
    
    UserID = CharId
    head = headNum
    
    If userOnline = True Then
        With UserList(UserIndex)
            .OrigChar.head = head

            Call WriteConsoleMsg(UserIndex, "La apariencia de tu personaje fue modificada desde la web. Reloguea para ver el cambio.", FontTypeNames.FONTTYPE_WARNING)
            Call ChangeUserChar(UserIndex, .Char.body, .OrigChar.head, .Char.heading, ConstantesGRH.NingunArma, ConstantesGRH.NingunEscudo, ConstantesGRH.NingunCasco)
        End With
    End If
    Call UpdateCharData("USER_INFO", "HEAD", UserID, head)

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleChangeHead de modMQ.bas")
End Sub

Private Sub HandleChangeGender(ByRef CharId As String, ByRef genderNum As String)
'***************************************************
'Author: Lucas Ezequiel Figelj (Luke)
'Last Modification: -
'Changes user gender from Web Request
'***************************************************
On Error GoTo ErrHandler
  
    Dim UserID As String
    Dim gender As String
    Dim UserIndex As Integer
    Dim userOnline As Boolean
    Dim I As Integer
    Dim defaultHead As String
    Dim defaultBody As String
    
    defaultBody = "1"
    defaultHead = "70"
    userOnline = False
    
    For I = 1 To UBound(UserList)
        If UserList(I).ID = CharId Then
            UserIndex = I
            userOnline = True
            Exit For
        End If
    Next I
    
    UserID = CharId
    gender = genderNum
    
    If userOnline = True Then
    
        With UserList(UserIndex)
    
            Call WriteConsoleMsg(UserIndex, "La apariencia de tu personaje fue modificada desde la web. Reloguea para ver el cambio.", FontTypeNames.FONTTYPE_WARNING)
        
        End With
    
    End If
    
    Call UpdateCharData("USER_INFO", "GENDER", UserID, gender)
    Call UpdateCharData("USER_INFO", "HEAD", UserID, defaultHead)
    Call UpdateCharData("USER_INFO", "BODY", UserID, defaultBody)
    

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HandleChangeGender de modMQ.bas")
End Sub

Public Function SendAccountCreateMessage(ByRef Email As String, ByRef AccountName As String, ByVal AccountId As Integer, ByRef Password As String, ByVal UserIP As String) As String
    
    On Error GoTo ErrHandler:
    
    Call MQOutboundByteQueue.WriteByte(eProxyServerMessages.ACCOUNT_CREATE)
    
    Call MQOutboundByteQueue.WriteInteger(AccountId)
    Call MQOutboundByteQueue.WriteASCIIString(AccountName)
    Call MQOutboundByteQueue.WriteASCIIString(Email)
    Call MQOutboundByteQueue.WriteASCIIString(Password)
    Call MQOutboundByteQueue.WriteASCIIString(UserIP)
    
    Call SendProxyServerData(MQOutboundByteQueue)
    
    Exit Function
    
ErrHandler:
    Call LogError("Error en SendAccountCreateMessage. Error: : " & Err.Number & ": " & Err.Description)
End Function

Public Function SendAccountPasswordChangedMessage(ByVal AccountId As Integer, ByRef AccountName As String, ByRef Email As String, ByRef NewPassword As String, ByVal UserIP As String) As String
    
    On Error GoTo ErrHandler:
    
    Call MQOutboundByteQueue.WriteByte(eProxyServerMessages.ACCOUNT_PASSWORD_CHANGED)
    
    Call MQOutboundByteQueue.WriteInteger(AccountId)
    Call MQOutboundByteQueue.WriteASCIIString(AccountName)
    Call MQOutboundByteQueue.WriteASCIIString(Email)
    Call MQOutboundByteQueue.WriteASCIIString(NewPassword)
    Call MQOutboundByteQueue.WriteASCIIString(UserIP)
    
    Call SendProxyServerData(MQOutboundByteQueue)
    
    Exit Function
    
ErrHandler:
    Call LogError("Error en SendAccountPasswordChangedMessage. Error: : " & Err.Number & ": " & Err.Description)
End Function

Public Sub SendProxyServerData(ByRef buffer As clsByteQueue)
On Error GoTo ErrHandler:

    Dim blockToSend() As Byte
    ReDim blockToSend(0 To buffer.length - 1)
    Call buffer.ReadBlock(blockToSend, buffer.length)

    If IsProxyServerOnline() Then
        Call frmMain.sckProxySender.SendData(blockToSend)
    End If
    Exit Sub
    
ErrHandler:
    Call LogError("Error en SendProxyServerData sending " & buffer.length & " bytes. Error: : " & Err.Number & ": " & Err.Description)
End Sub

Public Function SendDeathEventUserKilledUser(ByVal AttackerId As Long, ByRef AttackerName As String, ByRef AttackerPosition As String, _
                                            ByVal AttackerLevel As Integer, ByVal AttackerEloOld As Long, ByVal AttackerEloNew As Long, _
                                            ByVal VictimId As Long, ByRef VictimName As String, ByRef VictimPosition As String, _
                                            ByVal VictimLevel As Integer, ByVal VictimEloOld As Long, ByVal VictimEloNew As Long, _
                                            ByVal DamageType As Byte, ByVal DamageValue As Long, ByVal DamageWeaponIndex As Integer, _
                                            ByVal AttackerGuildId As Long, ByRef AttackerGuildName As String, _
                                            ByVal VictimGuildId As Long, ByRef VictimGuildName As String, _
                                            ByVal AttackerGuildEloOld As Long, ByVal AttackerGuildEloNew As Long, _
                                            ByVal VictimGuildEloOld As Long, ByVal VictimGuildEloNew As Long) As String
    
On Error GoTo ErrHandler:
    
    Call MQOutboundByteQueue.WriteByte(eProxyServerMessages.DEATH_EVENT_USER_KILLED_USER)
    
    Call MQOutboundByteQueue.WriteLong(AttackerId)
    Call MQOutboundByteQueue.WriteASCIIString(AttackerName)
    Call MQOutboundByteQueue.WriteASCIIString(AttackerPosition)
    Call MQOutboundByteQueue.WriteByte(AttackerLevel)
    Call MQOutboundByteQueue.WriteLong(AttackerEloOld)
    Call MQOutboundByteQueue.WriteLong(AttackerEloNew)
    
    Call MQOutboundByteQueue.WriteLong(VictimId)
    Call MQOutboundByteQueue.WriteASCIIString(VictimName)
    Call MQOutboundByteQueue.WriteASCIIString(VictimPosition)
    Call MQOutboundByteQueue.WriteByte(VictimLevel)
    Call MQOutboundByteQueue.WriteLong(VictimEloOld)
    Call MQOutboundByteQueue.WriteLong(VictimEloNew)
    
    Call MQOutboundByteQueue.WriteByte(DamageType)
    Call MQOutboundByteQueue.WriteLong(DamageValue)
    Call MQOutboundByteQueue.WriteInteger(DamageWeaponIndex)
    
    Call MQOutboundByteQueue.WriteLong(AttackerGuildId)
    Call MQOutboundByteQueue.WriteASCIIString(AttackerGuildName)
    
    Call MQOutboundByteQueue.WriteLong(VictimGuildId)
    Call MQOutboundByteQueue.WriteASCIIString(VictimGuildName)
    
    Call MQOutboundByteQueue.WriteLong(AttackerGuildEloOld)
    Call MQOutboundByteQueue.WriteLong(AttackerGuildEloNew)
    
    Call MQOutboundByteQueue.WriteLong(VictimGuildEloOld)
    Call MQOutboundByteQueue.WriteLong(VictimGuildEloNew)
    
    Call MQOutboundByteQueue.WriteASCIIString(Now())
        
    Call SendProxyServerData(MQOutboundByteQueue)
    
    Exit Function
    
ErrHandler:
    Call LogError("Error en SendDeathEventUserKilledUser. Error: : " & Err.Number & ": " & Err.Description)
End Function

Public Function IsProxyServerOnline() As Boolean
    On Error GoTo ErrHandler:

    IsProxyServerOnline = frmMain.sckProxySender.State = sckConnected
    Exit Function
    
ErrHandler:
    Call LogError("Error en IsStateServerOnline. Error: : " & Err.Number & ": " & Err.Description)
End Function
