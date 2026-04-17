Attribute VB_Name = "modAccount"
'**************************************************************************
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
Option Explicit

' Creado por Zama (Basado en código de maTih.-)

Public Enum eAccountStatus
    ActivationPending = 0
    Activated = 1
    Banned = 2
End Enum

'Cuenta multi-logeable , 1 activado 0 desactivado.
Private Const MULTI_LOG As Byte = 0


Public Function AccountIsPremium(ByVal nAccountID As Integer) As Boolean
'******************************************
'Author: D'Artagnan
'Date: 29/04/2015
'
'******************************************
On Error GoTo ErrHandler
  
    Dim rsQuery As ADODB.Recordset
    
    Set rsQuery = ExecuteSql("SELECT COUNT(1) FROM PREMIUM_EMAILS AS PE " & _
                             "INNER JOIN ACCOUNT_INFO AS AI ON AI.ID_ACCOUNT=" & CStr(nAccountID) & _
                             " WHERE PE.EMAIL=AI.EMAIL")
    
    AccountIsPremium = CInt(rsQuery.Fields(0)) > 0
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function AccountIsPremium de modAccount.bas")
End Function
 
Public Function CreateAccount(ByVal UserIndex As Integer, ByRef sName As String, _
    ByRef sPassword As String, ByRef sPregunta As String, _
    ByRef sRespuesta As String, ByRef sEmail As String) As Boolean
'***************************************************
'Author: Zama
'Creation Date: 23/01/2014
'Last Modification: 23/01/2014
'Validates and creates Account
'***************************************************
On Error GoTo ErrHandler
  
    ' Valid info?
    Dim sError As String
    
    If Not AccountInfo_Validate(sName, sPassword, sPregunta, sRespuesta, sEmail, sError) Then
        Call DisconnectWithMessage(UserIndex, sError)
        Exit Function
    End If
    
    Dim AccountId As Long
    Dim defaultActivationStatus As eAccountStatus
    
        ' Set the default activation status: if we're not using the Proxy Server, then set to Active. Otherwhise use ActivationPending
    If ServerConfiguration.UseExternalAccountValidation Then
    
        ' Validate if the ProxySender is connected, as this service is required for the external account validation process.
        If frmMain.sckProxySender.State <> sckConnected Then
            Call DisconnectWithMessage(UserIndex, "No se puede crear una cuenta en este momento, intentelo nuevamente más tarde.")
            Exit Function
        End If
    
        defaultActivationStatus = eAccountStatus.ActivationPending
    Else
        defaultActivationStatus = eAccountStatus.Activated
    End If

    CreateAccount = SaveAccountInfoDB(AccountId, sName, sEmail, sPassword, sPregunta, sRespuesta, defaultActivationStatus)
    
    If Not CreateAccount Then
        Call DisconnectWithMessage(UserIndex, "No se pudo crear la cuenta. Intente nuevamente más tarde.")
        Exit Function
    End If
    
    If ServerConfiguration.UseExternalAccountValidation Then
        ' Send the confirmation email to start the activation/validation process using the external tools
        Call modMessageQueueProxy.SendAccountCreateMessage(sEmail, sName, AccountId, sPassword, UserList(UserIndex).IP)
    Else
        UserList(UserIndex).AccountId = AccountId
    End If

    CreateAccount = True
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function CreateAccount de modAccount.bas")
End Function

Public Function ConectarNuevoPersonaje(ByVal UserIndex As Integer, ByVal sName As String, ByVal nGender As eGenero, _
                                  ByVal nRace As eRaza, ByVal nClass As eClass, ByVal nHead As Integer, ByVal nHome As Byte, ByVal SessionIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 17/06/2014 (D'Artagnan)
'Conecta un nuevo personaje.
'02/06/2014: D'Artagnan - New arguments.
'***************************************************
On Error GoTo ErrHandler
    Dim CharacterCreated As Boolean
   
    With UserList(UserIndex)
        'Agrega el personaje a la cuenta.
        CharacterCreated = Account_AddChar(UserIndex, SessionIndex, sName, nGender, nRace, nClass, nHead, nHome)
            
        If CharacterCreated Then
            ' Now we are passing the .ID to the ConnectUser because that property is already filled with the right
            ' id after the creation of the character.
            Call ConnectUser(UserIndex, sName, .ID, sName, False)
        End If
    End With
    
    ConectarNuevoPersonaje = CharacterCreated

  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ConectarNuevoPersonaje de modAccount.bas")
  Err.Raise Err.Number
End Function
 
Public Function ConnectChar(ByVal UserIndex As Integer, ByVal CharSlot As Byte) As Boolean
'***************************************************
'Author: ZaMa
'Creation Date: 23/01/2014
'Last Modification: 28/10/2014
'Validates and connects Account Char.
'Return True if the character has been successfully connected.
'False otherwise.
'06/07/2014: D'Artagnan - Solve special name.
'28/10/2014: D'Artagnan - Return value.
'***************************************************
On Error GoTo ErrHandler
  
    Dim sUserName As String
    Dim sSpecialName As String

    Dim SessionCharacter As SessionCharacter
    SessionCharacter = aActiveSessions(UserList(UserIndex).nSessionId).asAccountCharNames(CharSlot)
    
    sUserName = SessionCharacter.charName

    If InStr(1, sUserName, vbNullChar) > 0 Then
        sSpecialName = ReadField(1, sUserName, Asc(vbNullChar))
        sUserName = ReadField(2, sUserName, Asc(vbNullChar))
    Else
        sSpecialName = sUserName
    End If

    'No hay personaje.
    If (LenB(sUserName) = 0) Then Exit Function
    
    ConnectChar = ConnectUser(UserIndex, sUserName, SessionCharacter.CharId, sSpecialName, False)
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function ConnectChar de modAccount.bas")
End Function
 
Public Sub AccountChar_Delete(ByVal UserIndex As Integer, ByVal SessionIndex As Long, ByVal CharSlot As Byte, ByRef AccountAnswer As String)
'***************************************************
'Author: ZaMa
'Creation Date: 23/01/2014
'Last Modification: 02/07/2014 (D'Artagnan)
'Deletes char.
'02/07/2014: D'Artagnan - Remove user from database.
'***************************************************
On Error GoTo ErrHandler
  
    
    Dim sError As String
    Dim I As Integer
    Dim nUserID As Long
    Dim asUserInfoTables() As String
    Dim sUserInfoTables As String
    Dim AccountName As String
    

    If Not AccountCharDelete_Validate(UserIndex, SessionIndex, CharSlot, AccountAnswer, sError) Then
        Call DisconnectWithMessage(UserIndex, sError)
        Exit Sub
    End If
    
    ' Get the char list
    Dim charList() As Char_Acc_Data
    ReDim charList(1 To MAX_ACCOUNT_CHARS)
    Call GetAccountCharacters(UserIndex, aActiveSessions(SessionIndex).nAccountID, charList)
       
    nUserID = charList(CharSlot).ID
        
    If nUserID < 0 Then
        Call DisconnectWithMessage(UserIndex, "Error al eliminar el personaje. Contacte a un administrador.")
        Exit Sub
    End If
    
    ' Call the stored procedure to remove the character info from all the tables
    Dim Cmd As ADODB.Command
    Set Cmd = New ADODB.Command
                      
    
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "sp_DeleteChar"
    
    Cmd.Parameters.Append Cmd.CreateParameter("userID", adInteger, adParamInput, 1, nUserID)
    
    Call ExecuteSqlCommand(Cmd)
    Set Cmd = Nothing
    
    ' Remove character from session.
    Call SessionRemoveAccountCharName(SessionIndex, CharSlot)
    
    Call WriteAccountRemoveChar(UserIndex, CharSlot)

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub AccountChar_Delete de modAccount.bas")
End Sub
  
Public Function connect(ByVal UserIndex As Integer, ByRef AccountName As String, ByRef sPassword As String, _
                        Optional ByRef sPreviousToken As String = vbNullString, _
                        Optional ByVal bCheckPassword As Boolean = True, Optional ByRef ClientTempCode As String = "") As Boolean
'***************************************************
'Author: ZaMa
'Creation Date: 23/01/2014
'Last Modification: 20/06/2014 (D'Artagnan)
'Conecta la cuenta.
'20/06/2014: D'Artagnan - Create new session.
'***************************************************
On Error GoTo ErrHandler
  
    Dim sError As String
    Dim sToken As String
    Dim AccountId As Long
    Dim AccountSessionId As Integer
    Dim AccountEmail As String
    
    AccountSessionId = -1
    
    ' valid?
    If Not AccountLogin_Validate(AccountId, AccountName, sPassword, AccountEmail, sError, bCheckPassword) Then
        Call DisconnectWithMessage(UserIndex, sError)
        Exit Function
    End If

    If sPreviousToken = "" Then
        AccountSessionId = modSession.GetNewSession(UserIndex, ClientTempCode, AccountName, AccountEmail)
    Else
        ' Check if the previous token corresponds to an active session. If not, create a new one.
        ' This is to reuse the slot and avoid going the list of sessions to find a new one.
        If Not modSession.IsTokenValid(sPreviousToken, UserIndex, AccountSessionId) Then
            AccountSessionId = modSession.GetNewSession(UserIndex, ClientTempCode, AccountName, AccountEmail)
            Call LogError(AccountName & ": Token inválido. Creando token nuevo: " & AccountSessionId & " (" & aActiveSessions(AccountSessionId).ServerTempCode & ")")
        End If
    End If
    
    If AccountSessionId = -1 Then
        ' ERROR - Close connection, let the user know that we couldn't create a new session to connect
        ' The server is busy or full right now.
        Call DisconnectWithMessage(UserIndex, "No se pudo encontrar ni crear una sesión en este momento. Intente nuevamente más tarde.")
        Exit Function
    End If
    
    ' Clean all the characters from the session.
    Call modSession.CleanCharactersFromSession(AccountSessionId)
    
    ' Refresh the expiration counter
    Call modSession.ExtendSessionLifetime(AccountSessionId)
          
    ' NIGHTW TODO END: Sessions

    ' Generate a new ServerCode and send it to the client
    Call modSession.SetAccountDataToSession(AccountSessionId, AccountId, AccountName, AccountEmail)
    Call modSession.RegenerateSessionServerCode(AccountSessionId)
    Call modSession.RecreateToken(AccountSessionId)
    Call WriteSendSessionToken(UserIndex, AccountSessionId)
    
    ' Send char List
    Call SendCharList(UserIndex, AccountSessionId, AccountId)
    
    
    connect = True
    
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function connect de modAccount.bas")
End Function
 
Public Function Account_AddChar(ByVal UserIndex As Integer, ByVal SessionId As Integer, ByVal sName As String, _
                                ByVal nGender As eGenero, ByVal nRace As eRaza, ByVal nClass As eClass, _
                                ByVal nHead As Integer, ByVal nHome As Byte) As Boolean
'***************************************************
'Author: ZaMa
'Creation Date: 23/01/2014
'Last Modification: 02/06/2014 (D'Artagnan)
'Agrega un personaje a la cuenta.
'02/06/2014: D'Artagnan - Error handling and set user info into the db.
'***************************************************
 
 On Error GoTo ErrHandler
 
    ' Next free slot
    Dim CharSlot  As Byte
    CharSlot = GetFreeSlot(UserIndex, SessionId)
     
    'Encuentra slot.
    If (CharSlot = 0) Then
        Call DisconnectWithMessage(UserIndex, "No tienes espacio para más personajes.")
        Exit Function
    End If
    
    ' Look for valid input values.
    If Not ValidCharacterData(UserIndex, sName, nRace, nClass, nGender, nHead) Then
        Exit Function
    End If

    ' Append user info into the database.
    Call CreateCharacter(UserIndex, sName, "", nRace, nGender, nClass, nHead, aActiveSessions(SessionId).nAccountID)
    
    UserList(UserIndex).AccountCharNames(CharSlot) = sName
    
    Account_AddChar = True
    Exit Function

ErrHandler:
    Call LogError("Account_AddChar failed. Error " & Err.Number & ": " & Err.Description & ".")
    Err.Raise Err.Number
End Function
 
Public Sub SendCharList(ByVal UserIndex As Integer, ByVal SessionIndex As Integer, ByVal AccountId As Long)
'***************************************************
'Author: ZaMa
'Creation Date: 23/01/2014
'Last Modification: 20/06/2014 (D'Artagnan)
'Sends account char list.
'04/06/2014: D'Artagnan - Look for the user data in the database (if enabled).
'20/06/2014: D'Artagnan - Send the session token.
'***************************************************
On Error GoTo ErrHandler
  

    Dim Slot As Byte
    Dim CharDetail As Char_Acc_Data
    
    Dim aCharactersData(1 To MAX_ACCOUNT_CHARS) As Char_Acc_Data
    
    Call GetAccountCharacters(UserIndex, AccountId, aCharactersData)
    
    ' Send information of each character.
    For Slot = 1 To MAX_ACCOUNT_CHARS
        ' Make sure it's not an empty item.
        If (aCharactersData(Slot).Nick_Name <> vbNullString) Then
            ' Store character in session.
            Call SessionSetAccountCharacterData(SessionIndex, Slot, _
                                           aCharactersData(Slot).Nick_Name, aCharactersData(Slot).ID)
                                           
            ' Send character to client.
            Call Protocol.WriteAccountPersonaje(UserIndex, Slot, aCharactersData(Slot))
        End If
    Next Slot
    
    ' Shows account form.
    Call Protocol.WriteAccountShow(UserIndex)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SendCharList de modAccount.bas")
End Sub
 
Public Function AccountCharDelete_Validate(ByVal UserIndex As Integer, ByVal SessionIndex As Integer, _
                                           ByVal CharSlot As Byte, ByRef AccountAnswer As String, ByRef sError As String) As Boolean
'***************************************************
'Author: ZaMa
'Creation Date: 23/01/2014
'Last Modification: 29/10/2014 (D'Artagnan)
'Validates account char delete
'02/07/2014: D'Artagnan - Get username from database.
'29/10/2014: D'Artagnan - Ban check.
'***************************************************
On Error GoTo ErrHandler
    Dim AccountName As String
    Dim AccountId As Integer
    Dim charName As String
    Dim CharId As Long
    
    With aActiveSessions(SessionIndex).asAccountCharNames(CharSlot)
        CharId = .CharId
        charName = .charName
    End With
    
    With aActiveSessions(SessionIndex)
        AccountName = .sAccountName
        AccountId = .nAccountID
    End With
    
    ' Valid account?
    If (AccountName = vbNullString) Then
         sError = "Cuenta inválida."
         Exit Function
    End If
    
    ' Valid slot?
    If (CharSlot = 0) Or (CharSlot > MAX_ACCOUNT_CHARS) Then
        sError = "Personaje inválido!"
        Exit Function
    End If
    
    ' Valid char?
    If LenB(charName) = 0 Then
        sError = "Personaje inválido!"
        Exit Function
    End If
    
    'Dim UserID As Long
    Dim isBanned As Boolean
    
    'Call GetCharInfo(charName, UserID, sPassword, isBanned)
    Call GetCharInfoByCharId(CharId, isBanned)
    
    ' Valid account answer?
    'Dim nUserID As Long
    Dim StoredAccAnswer As String
    'Dim nAccountID As Long
    'nUserID = GetUserID(SessionGetAccountCharName(UserList(UserIndex).nSessionId, CharSlot))
    'nAccountID = GetAccountIDByUserID(CLng(nUserID))
    StoredAccAnswer = GetAccountAnswer(AccountId)
    If StoredAccAnswer <> AccountAnswer Then
        sError = "El token ingresado es inválido."
        Exit Function
    End If
    
    ' Ban check
    If isBanned Then
        sError = "Este personaje no puede ser eliminado."
        Exit Function
    End If
    
    ' Online?
    Dim CharIndex As Integer
    CharIndex = NameIndex(charName)
    If (CharIndex <> 0) Then
       sError = "El personaje está conectado, no puede ser eliminado."
       Exit Function
    End If
    
    ' Has guild?
    Dim GuildIndex As Integer
    GuildIndex = Val(GetCharData("USER_INFO", "GUILD_ID", CharId))
    
    If GuildIndex <> 0 Then
        sError = "El personaje se encuentra en un clan, no puede ser eliminado."
        Exit Function
    End If
                   
    ' Is Level 30 or more
    Dim CharLevel As Integer
    CharLevel = Val(GetCharData("USER_STATS", "NIVEL", CharId))
    If CharLevel > 30 Then
        sError = "No puedes borrar personajes nivel 30 a superiores."
        Exit Function
    End If
    
    AccountCharDelete_Validate = True
    
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function AccountCharDelete_Validate de modAccount.bas")
End Function
 
Public Function AccountInfo_Validate(ByRef sName As String, ByRef sPassword As String, _
    ByRef sPregunta As String, ByRef sRespuesta As String, _
    ByRef sEmail As String, ByRef sError As String) As Boolean
'***************************************************
'Author: Zama
'Creation Date: 23/01/2014
'Last Modification: 07/11/2014 (D'Artagnan)
'Checkea la creación de una nueva cuenta.
'07/11/2014: D'Artagnan - Check maximum fields length.
'***************************************************
On Error GoTo ErrHandler
  
    AccountInfo_Validate = False
    
    ' Valid name?
    sName = Trim$(sName)
    If (sName = vbNullString) Then
        sError = "El nombre de la cuenta es inválido."
        Exit Function
    End If
    
    ' Maximum account name length.
    If Len(sName) > 30 Then
        sError = "El nombre de la cuenta es demasiado extenso."
        Exit Function
    End If
    
    ' Valid Password?
    sPassword = Trim$(sPassword)
    If (sPassword = vbNullString) Then
        sError = "La contraseña es inválida."
        Exit Function
    End If
    
    ' Maximum password length.
    If Len(sPassword) > 32 Then
        sError = "La contraseña es demasiado extensa."
        Exit Function
    End If
    
    ' Valid Question?
    sPregunta = Trim$(sPregunta)
    If (sPregunta = vbNullString) Then
        sError = "Falta ingresar la pregunta."
        Exit Function
    End If
    
    ' Maximum question length.
    If Len(sPregunta) > 50 Then
        sError = "La pregunta es demasiado extensa."
        Exit Function
    End If

    ' Maximum answer length.
    If Len(sEmail) > 50 Then
        sError = "La respuesta es demasiado extensa."
        Exit Function
    End If
    
    ' Valid email?
    sEmail = Trim$(sEmail)
    If (sEmail = vbNullString) Then
        sError = "Falta ingresar el email."
        Exit Function
    End If
    
    ' Maximum email length.
    If Len(sEmail) > 50 Then
        sError = "La dirección de correo es demasiado extensa."
        Exit Function
    End If
    
    If AccountEmailExists(sEmail) Then
        sError = "Ya existe una cuenta asociada al email."
        Exit Function
    End If
    
    ' Valid name chars?
    If Not AsciiValidos(sName, False, True, False) Then
        sError = "El nombre posee carácteres inválidos."
        Exit Function
    End If
    

    ' Already exists?
    Dim AccountId As Long
    AccountId = GetAccountID(sName)
    If AccountId <> 0 Then
        sError = "Ya existe una cuenta con ese nombre."
        Exit Function
    End If
    
    AccountInfo_Validate = True
   
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function AccountInfo_Validate de modAccount.bas")
End Function
 
Public Function AccountLogin_Validate(ByRef AccountId As Long, ByRef AccountName As String, ByRef sPassword As String, ByRef AccountEmail As String, _
                                      ByRef sError As String, Optional ByVal bCheckPassword As Boolean = True) As Boolean
'***************************************************
'Author: ZaMa
'Creation Date: 23/01/2014
'Last Modification: 18/06/2014
'Validates account login.
'18/06/2014: D'Artagnan - Optional password checking.
'***************************************************
On Error GoTo ErrHandler

    Dim CurrentPassword As String
    Dim AccountData As tAccountData

    ' Valid name?
    AccountName = Trim$(AccountName)
    If (AccountName = vbNullString) Then
        sError = "El nombre de la cuenta es inválido."
        Exit Function
    End If
    
    ' Valid name chars?
    If Not AsciiValidos(AccountName, False, True) Then
        sError = "El nombre posee carácteres inválidos."
        Exit Function
    End If

    ' Exists?
    AccountData = GetAccount(AccountName)

    If bCheckPassword Then
        If AccountData.ID <= 0 Or AccountData.Password <> sPassword Then
            sError = "Los datos ingresados son incorrectos. Por favor, verifique el nombre de cuenta y " & _
                     "su clave de acceso e intente nuevamente."
            Exit Function
        End If
    End If
    
    ' Banned?
    If AccountData.Status = eAccountStatus.Banned Then
        sError = "La cuenta está baneada."
        Exit Function
    End If
    
    If AccountData.Status = eAccountStatus.ActivationPending Then
        sError = "Tu cuenta no se encuentra activada. Por favor, revisa tu correo electrónico."
        Exit Function
    End If

    AccountId = AccountData.ID
    AccountEmail = AccountData.Email
    AccountLogin_Validate = True
               
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function AccountLogin_Validate de modAccount.bas")
End Function
 
Public Function GetAccountCharDetail(ByVal AccountName As String, ByVal CharSlot As Byte) As Char_Acc_Data
'***************************************************
'Author: ZaMa
'Creation Date: 23/01/2014
'Last Modification: 23/01/2014
'Returns account char details.
'***************************************************
On Error GoTo ErrHandler
  
    'FALTA_DB: mandar todo desde db?
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetAccountCharDetail de modAccount.bas")
End Function

Public Function IsAccountLogged(ByRef AccountName As String, ByVal AccountId As Long) As Boolean
'***************************************************
'Author: ZaMa
'Creation Date: 23/01/2014
'Last Modification: 23/01/2014
'Checkea si la cuenta está logeada.
'***************************************************
On Error GoTo ErrHandler
  
   
    ' Allow always
    If (MULTI_LOG <> 0) Then
        Exit Function
    End If
     
    Dim Counter As Long
    For Counter = 1 To LastUser
        If UserList(Counter).AccountId = AccountId Then
            IsAccountLogged = True
            Exit Function
        End If
    Next Counter
 
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function IsAccountLogged de modAccount.bas")
End Function
 
Public Function GetAccountPath() As String
'***************************************************
'Author: ZaMa
'Creation Date: 23/01/2014
'Last Modification: 23/01/2014
'Dir de las cuentas.
'***************************************************
On Error GoTo ErrHandler
  
 
    GetAccountPath = App.Path & "\Accounts\"
     
    'No existe el directorio?
    If Not FileExist(GetAccountPath, vbDirectory) Then MkDir GetAccountPath
 
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetAccountPath de modAccount.bas")
End Function
 
Private Function GetFreeSlot(ByVal UserIndex As Integer, ByVal SessionIndex As Integer) As Byte
'***************************************************
'Author: ZaMa
'Creation Date: 23/01/2014
'Last Modification: 23/01/2014
'Returns first free slot.
'***************************************************
On Error GoTo ErrHandler

    Dim charName As String
    Dim Slot As Long
    
    For Slot = 1 To MAX_ACCOUNT_CHARS
    
        charName = aActiveSessions(SessionIndex).asAccountCharNames(Slot).charName
        
        If LenB(charName) = 0 Or charName = vbNullString Then
           GetFreeSlot = CByte(Slot)
           Exit Function
        End If
    Next Slot
    
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetFreeSlot de modAccount.bas")
End Function
 
Private Function GetAccountBanned(ByVal AccountName As String, ByVal AccountId As Long) As Byte
'***************************************************
'Author: Zama
'Creation Date: 23/01/2014
'Last Modification: 23/01/2014
'Checkea si la cuenta está baneada.
'***************************************************
On Error GoTo ErrHandler
  
    GetAccountBanned = CByte(IIf(Val(GetAccountData("STATUS", AccountId)) = eAccountStatus.Banned, 1, 0))
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetAccountBanned de modAccount.bas")
End Function

Public Function GetAccount(ByRef AccountName As String) As tAccountData
On Error GoTo ErrHandler:
    
    Dim strQuery As String
    Dim rsQuery As Recordset
    Dim AccountData As tAccountData
    
    strQuery = "SELECT " & _
        "ID_ACCOUNT, " & _
        "NAME, " & _
        "EMAIL, " & _
        "PASSWORD, " & _
        "ANSWER AS TOKEN, " & _
        "STATUS, " & _
        "BAN_DETAIL, " & _
        "CREATION_DATE, " & _
        "BANK_GOLD, " & _
        "BANK_PASSWORD " & _
    "FROM account_info " & _
    "WHERE NAME = '" & AccountName & "'"
    
    
    Set rsQuery = ExecuteSql(strQuery)
    
    If rsQuery.RecordCount <= 0 Then
        Exit Function
    End If
        
    With AccountData
        .ID = rsQuery.Fields("ID_ACCOUNT")
        .Name = rsQuery.Fields("NAME")
        .Email = rsQuery.Fields("EMAIL")
        .Password = rsQuery.Fields("PASSWORD")
        .Token = rsQuery.Fields("TOKEN")
        .Status = CByte(rsQuery.Fields("STATUS"))
        .BanDetail = rsQuery.Fields("BAN_DETAIL")
        .CreationDate = CDate(rsQuery.Fields("CREATION_DATE"))

        .BankGold = CLng(rsQuery.Fields("BANK_GOLD"))
        .BankPassword = rsQuery.Fields("BANK_PASSWORD")
    End With
    
    GetAccount = AccountData
        
    Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetAccount de modAccount.bas")
    
End Function

Public Function GetAccountById(ByRef AccountId As Long) As tAccountData
On Error GoTo ErrHandler:
    
    Dim strQuery As String
    Dim rsQuery As Recordset
    Dim AccountData As tAccountData
    
    strQuery = "SELECT " & _
        "ID_ACCOUNT, " & _
        "NAME, " & _
        "EMAIL, " & _
        "PASSWORD, " & _
        "ANSWER AS TOKEN, " & _
        "STATUS, " & _
        "BAN_DETAIL, " & _
        "CREATION_DATE, " & _
        "BANK_GOLD, " & _
        "BANK_PASSWORD " & _
    "FROM ACCOUNT_INFO " & _
    "WHERE ID_ACCOUNT = '" & AccountId & "'"
    
    
    Set rsQuery = ExecuteSql(strQuery)
    
    If rsQuery.RecordCount <= 0 Then
        Exit Function
    End If
        
    With AccountData
        .ID = rsQuery.Fields("ID_ACCOUNT")
        .Name = rsQuery.Fields("NAME")
        .Email = rsQuery.Fields("EMAIL")
        .Password = rsQuery.Fields("PASSWORD")
        .Token = rsQuery.Fields("TOKEN")
        .Status = CByte(rsQuery.Fields("STATUS"))
        .BanDetail = rsQuery.Fields("BAN_DETAIL")
        .CreationDate = CDate(rsQuery.Fields("CREATION_DATE"))

        .BankGold = CLng(rsQuery.Fields("BANK_GOLD"))
        .BankPassword = rsQuery.Fields("BANK_PASSWORD")
    End With
    
    GetAccountById = AccountData
        
    Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetAccountById de modAccount.bas")
    
End Function


Private Function GetNormalizedDate(ByRef datePart As String, ByRef timePart As String) As String
    
    timePart = Replace$(timePart, "a.m.", "AM")
    timePart = Replace$(timePart, "p.m.", "PM")
    
    GetNormalizedDate = datePart & " " & timePart
    
End Function



Public Function GetAccountIDByUserID(ByVal nUserID As Long) As Long
'***************************************************
'Author: D'Artagnan
'Date: 05/01/2015
'Return the account ID associated with the specified user ID.
'***************************************************
On Error GoTo ErrHandler
  
    GetAccountIDByUserID = CLng(GetCharData("USER_INFO", "ID_ACCOUNT", nUserID))
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetAccountIDByUserID de modAccount.bas")
End Function
 
Private Function GetAccountPassword(ByRef AccountName As String, ByVal AccountId As Long) As String
'***************************************************
'Author: Zama
'Creation Date: 23/01/2014
'Last Modification: 23/01/2014
'Devuelve la pass de una cuenta.
'***************************************************
On Error GoTo ErrHandler

    GetAccountPassword = GetAccountData("PASSWORD", AccountId)
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetAccountPassword de modAccount.bas")
End Function
 
Public Function GetAccountQuestion(ByRef AccountName As String, ByVal AccountId As Long) As String
'***************************************************
'Author: Zama
'Creation Date: 23/01/2014
'Last Modification: 23/01/2014
'Returns secret question.
'***************************************************
On Error GoTo ErrHandler
  
    GetAccountQuestion = GetAccountData("SECRET_QUESTION", AccountId)
 
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetAccountQuestion de modAccount.bas")
End Function
 
Public Function GetAccountAnswer(ByVal AccountId As Long) As String
'***************************************************
'Author: Zama
'Creation Date: 23/01/2014
'Last Modification: 23/01/2014
'Returns secret question's answer.
'***************************************************
On Error GoTo ErrHandler

    GetAccountAnswer = GetAccountData("ANSWER", AccountId)
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetAccountAnswer de modAccount.bas")
End Function

Private Sub GetAccountCharacters(ByVal nUserIndex As Integer, ByVal AccountId As Long, _
                                 ByRef aBuffer() As Char_Acc_Data)
'***************************************************
'Author: D'Artagnan
'Creation Date: 04/06/2014
'Last Modification: -
'Fill the buffer with the user(s) info stored in the specified account.
'***************************************************
On Error GoTo ErrHandler
  



    Dim nSlot As Integer
    Dim nAccountID As Long
    Dim nCurrentCharacterID As Long
    Dim currentCharacter As Char_Acc_Data
    Dim sQuery As String
    Dim rsQuery As Recordset
    
    'nAccountID = GetAccountID(sAccountName)
    
    ' Get those characters which match with the account ID
    sQuery = "SELECT UI.ID_USER, UI.NAME, UI.LAST_POS, UI.BODY, UI.HEAD, UI.WEAPON_ANIM, " & _
             "UI.SHIELD_ANIM, UI.HELMET_ANIM, US.NIVEL, UF.Muerto, UF.Navegando, UI.PUNISHMENT, " & _
             "UFA.ALIGNMENT, UI.BANNED, GI.ID_GUILD, GI.NAME AS GUILD_NAME FROM USER_INFO AS UI " & _
             "INNER JOIN USER_STATS AS US ON UI.ID_USER = US.ID_USER " & _
             "INNER JOIN USER_FLAGS AS UF ON UI.ID_USER = UF.ID_USER " & _
             "LEFT JOIN GUILD_INFO AS GI ON UI.GUILD_ID = GI.ID_GUILD " & _
             "INNER JOIN USER_FACTION AS UFA ON UI.ID_USER = UFA.ID_USER " & _
             "WHERE ID_ACCOUNT = " & CStr(AccountId)
    
    Set rsQuery = ExecuteSql(sQuery)
    
    ' Iterate until reach the maximum slots
    For nSlot = 1 To MAX_ACCOUNT_CHARS
        If rsQuery.EOF Then Exit Sub
        
        With currentCharacter
            ' User ID
            nCurrentCharacterID = CLng(rsQuery.Fields("ID_USER"))
    
            ' General
            .ID = nCurrentCharacterID
            .Nick_Name = CStr(rsQuery.Fields("NAME"))
            .Pos_Map = CStr(rsQuery.Fields("LAST_POS"))
    
            ' Level
            .Nivel = CInt(rsQuery.Fields("NIVEL"))
    
            ' Flags
            .Muerto = CInt(rsQuery.Fields("MUERTO"))
            .bSailing = CInt(rsQuery.Fields("NAVEGANDO"))
            
            ' Guild
            If Not IsNull(rsQuery.Fields("ID_GUILD")) Then
                .IdGuild = CLng(rsQuery.Fields("ID_GUILD"))
            Else
                .IdGuild = 0
            End If
           
            If Not IsNull(rsQuery.Fields("GUILD_NAME")) Then
                .GuildName = rsQuery.Fields("GUILD_NAME")
            Else
                .GuildName = vbNullString
            End If

            ' Faction/Alignment
            .Alignment = CByte(rsQuery.Fields("ALIGNMENT"))
            
            ' Punishments
            .JailRemainingTime = CLng(rsQuery.Fields("PUNISHMENT"))
            .Banned = CBool(rsQuery.Fields("BANNED"))
    
            ' Graphics
            .Character.body = CInt(rsQuery.Fields("BODY"))
            .Character.head = CInt(rsQuery.Fields("HEAD"))
            .Character.WeaponAnim = CInt(rsQuery.Fields("WEAPON_ANIM"))
            .Character.ShieldAnim = CInt(rsQuery.Fields("SHIELD_ANIM"))
            .Character.CascoAnim = CInt(rsQuery.Fields("HELMET_ANIM"))
    
        End With
    
        ' Store data in the buffer
        aBuffer(nSlot) = currentCharacter
        
        rsQuery.MoveNext
    Next nSlot
    
    Exit Sub

ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub GetAccountCharacters de modAccount.bas")
End Sub

Public Function IsCharFromAccount(ByVal AccountId As Integer, ByRef charName As String) As Boolean
    
On Error GoTo ErrHandler
    Dim sQuery As String
    Dim rsQuery As Recordset
    
    sQuery = "SELECT COUNT(1) as CharMatches FROM USER_INFO "
    sQuery = sQuery & " INNER JOIN ACCOUNT_INFO ON USER_INFO.ID_ACCOUNT = ACCOUNT_INFO.ID_ACCOUNT "
    sQuery = sQuery & " WHERE USER_INFO.NAME = '" & charName & "' and USER_INFO.ID_ACCOUNT = " & AccountId
    
    Set rsQuery = ExecuteSql(sQuery)
    
    If rsQuery.EOF Then Exit Function

    IsCharFromAccount = CInt(rsQuery.Fields("CharMatches")) = 1
    
    Exit Function

ErrHandler:
    Call LogError("Error " & Err.Number & " in IsCharFromId(). " & Err.Description)
End Function

Private Function GetAccountCharName(ByRef AccountName As String, ByVal CharSlot As Byte) As String
'***************************************************
'Author: Zama
'Creation Date: 23/01/2014
'Last Modification: 23/01/2014
'Returns Account char name (if exists)
'***************************************************
On Error GoTo ErrHandler
  
 
    Dim sName As String
    sName = GetVar(GetAccountPath & AccountName & ".cuenta", "PERSONAJES", "PERSONAJE" & CStr(CharSlot))
    
    If (sName = "NoUsado") Then sName = vbNullString
    
    GetAccountCharName = sName
    
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetAccountCharName de modAccount.bas")
End Function
