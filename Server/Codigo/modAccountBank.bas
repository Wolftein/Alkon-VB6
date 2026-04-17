Attribute VB_Name = "modAccountBank"
Option Explicit

Public Type tAccBank
    Id As Long
    Password As String
    Oro As Long
    Object(1 To MAX_BANCOINVENTORY_SLOTS) As UserOBJ
End Type

Public BovedaCuenta() As tAccBank

Public Sub InitAccBank()
'---------------------------------------------------------------------------------------
' Module    : modAccountBank
' Author    : Anagrama
' Date      : 21/08/2016
' Purpose   : Inicia el sistema. Por ahora hace una negrada nada mas para darle un tamaño inicial al array dinamico.
'---------------------------------------------------------------------------------------
On Error GoTo ErrHandler
  
    ReDim BovedaCuenta(1 To 1) As tAccBank
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub InitAccBank de modAccountBank.bas")
End Sub

Public Function GetAccBankIndex(ByVal Id As Long) As Integer
'---------------------------------------------------------------------------------------
' Module    : modAccountBank
' Author    : Anagrama
' Date      : 21/08/2016
' Purpose   : Devuelve el indice de la AccountID en el array BovedaCuenta.
'---------------------------------------------------------------------------------------
On Error GoTo ErrHandler
  
    Dim I As Integer
    
    If Id = 0 Then Exit Function
    
    For I = 1 To UBound(BovedaCuenta)
        If BovedaCuenta(I).Id = Id Then
            GetAccBankIndex = I
            Exit Function
        End If
    Next I
    
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetAccBankIndex de modAccountBank.bas")
End Function

Private Function CanCloseAccBank(ByVal BovAccIndex As Integer) As Boolean
'---------------------------------------------------------------------------------------
' Module    : modAccountBank
' Author    : Anagrama
' Date      : 21/08/2016
' Purpose   : Revisa si algun otro personaje esta vinculado a esa cuenta y no puede cerrarla.
'---------------------------------------------------------------------------------------
On Error GoTo ErrHandler
  
    Dim I As Integer

    For I = 1 To LastUser
        If UserList(I).flags.UserLogged Then
            If UserList(I).flags.AccountBank = BovAccIndex Then
                Exit Function
            End If
        End If
    Next I
    
    CanCloseAccBank = True
    
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function CanCloseAccBank de modAccountBank.bas")
End Function

Public Sub CloseAccBank(ByVal BovAccIndex As Integer)
'---------------------------------------------------------------------------------------
' Module    : modAccountBank
' Author    : Anagrama
' Date      : 21/08/2016
' Purpose   : Limpia el slot.
'---------------------------------------------------------------------------------------
On Error GoTo ErrHandler
  
    Dim Slot As Byte
    
    If BovAccIndex = 0 Then Exit Sub
    
    If Not CanCloseAccBank(BovAccIndex) Then Exit Sub

    BovedaCuenta(BovAccIndex).Id = 0
    BovedaCuenta(BovAccIndex).Oro = 0
    BovedaCuenta(BovAccIndex).Password = vbNullString
    
    For Slot = 1 To MAX_BANCOINVENTORY_SLOTS
        BovedaCuenta(BovAccIndex).Object(Slot).ObjIndex = 0
        BovedaCuenta(BovAccIndex).Object(Slot).Amount = 0
    Next Slot
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CloseAccBank de modAccountBank.bas")
End Sub

Public Sub SaveAccountBankDB(ByVal UserIndex As Integer, ByVal NewChar As Boolean)
'---------------------------------------------------------------------------------------
' Module    : modAccountBank
' Author    : Anagrama
' Date      : 21/08/2016
' Purpose   : Guarda la informacion de la boveda en la BD.
'---------------------------------------------------------------------------------------

On Error GoTo ErrHandler:

    Dim query As String
    Dim Slot As Long
    Dim AccountId As Long
    Dim BovAccIndex As Integer

    BovAccIndex = UserList(UserIndex).flags.AccountBank
    
    If BovAccIndex = 0 Then Exit Sub
    
    AccountId = BovedaCuenta(BovAccIndex).Id

    With BovedaCuenta(BovAccIndex)
        If Not NewChar Then
            query = _
                "DELETE FROM ACCOUNT_BANK " & _
                "WHERE ID_ACCOUNT = '" & CStr(AccountId) & "' "
                
            Call ExecuteSql(query)
        Else
            Exit Sub
        End If
        
        ' Save the Account Gold and password
        query = "UPDATE ACCOUNT_INFO SET BANK_GOLD = " & CStr(BovedaCuenta(BovAccIndex).Oro) & _
                ", BANK_PASSWORD = '" & CStr(BovedaCuenta(BovAccIndex).Password) & "' " & _
                " WHERE ID_ACCOUNT = " & CStr(AccountId)

        Call ExecuteSql(query)

        ' Save the account items
        For Slot = 1 To MAX_BANCOINVENTORY_SLOTS
            With .Object(Slot)
                If .ObjIndex <> 0 Then
                    query = _
                        "INSERT INTO ACCOUNT_BANK (ID_ACCOUNT, SLOT, OBJ_INDEX, AMOUNT) VALUES ('" & _
                            CStr(AccountId) & "','" & _
                            CStr(Slot) & "','" & _
                            CStr(.ObjIndex) & "','" & _
                            CStr(.Amount) & "' " & _
                        ")"
                    
                    Call ExecuteSql(query)
                End If
            End With
        Next Slot
    End With
    
    Exit Sub
    
ErrHandler:
    LogError ("Error en SaveAccountBankDB: " & Err.Description)
End Sub

Public Sub LoadAccountBankDB(ByVal UserIndex As Integer)
'---------------------------------------------------------------------------------------
' Module    : modAccountBank
' Author    : Anagrama
' Date      : 21/08/2016
' Purpose   : Carga la informacion de la cuenta de la BD si existe.
'---------------------------------------------------------------------------------------
On Error GoTo ErrHandler
    Dim Slot As Byte
    Dim AccountId As Long
    Dim BovAccIndex As Integer
    Dim query As String
    Dim Rs As Recordset
    
    AccountId = UserList(UserIndex).AccountId
    BovAccIndex = NextOpenAccBank(AccountId)
    UserList(UserIndex).flags.AccountBank = BovAccIndex
    
    If BovAccIndex = 0 Then Exit Sub
    
    BovedaCuenta(BovAccIndex).Id = AccountId
    
    ' Get bank password and gold from the ACCOUNT_INFO table
    query = "SELECT BANK_PASSWORD, BANK_GOLD " & _
            "FROM ACCOUNT_INFO " & _
            "WHERE ID_ACCOUNT = '" & CStr(AccountId) & "' "
        
    Set Rs = ExecuteSql(query)
    If Not Rs.EOF Then
        BovedaCuenta(BovAccIndex).Password = CStr(Rs.Fields("BANK_PASSWORD"))
        BovedaCuenta(BovAccIndex).Oro = CLng(Rs.Fields("BANK_GOLD"))
    Else
        Exit Sub
    End If
    
    ' Get the account items from the ACCOUNT_BANK table
    query = "SELECT SLOT, OBJ_INDEX, AMOUNT " & _
        "FROM ACCOUNT_BANK " & _
        "WHERE ID_ACCOUNT = '" & CStr(AccountId) & "' "
    
    Set Rs = ExecuteSql(query)
    
    With BovedaCuenta(BovAccIndex)
        Do While Not Rs.EOF
            Slot = CByte(Rs.Fields("SLOT"))
            .Object(Slot).ObjIndex = CInt(Rs.Fields("OBJ_INDEX"))
            .Object(Slot).Amount = CInt(Rs.Fields("AMOUNT"))
            
            Rs.MoveNext
        Loop
    End With
    
    Exit Sub

ErrHandler:
    Call LogError("Error " & Err.Number & " in LoadAccountBankDB(). " & Err.Description)
    
End Sub

Private Function NextOpenAccBank(ByVal AccountId As Long) As Integer
'---------------------------------------------------------------------------------------
' Module    : modAccountBank
' Author    : Anagrama
' Date      : 21/08/2016
' Purpose   : Devuelve el siguiente slot abierto en el array BovedaCuenta, si no hay suma uno.
'---------------------------------------------------------------------------------------
On Error GoTo ErrHandler
    Dim I As Integer
    
    If AccountId = 0 Then Exit Function
    
    For I = 1 To UBound(BovedaCuenta)
        If BovedaCuenta(I).Id = 0 Then
            NextOpenAccBank = I
            Exit Function
        End If
    Next I

    ReDim Preserve BovedaCuenta(1 To UBound(BovedaCuenta) + 1) As tAccBank
    NextOpenAccBank = UBound(BovedaCuenta)
    
    Exit Function
    
ErrHandler:
    Call LogError("Error " & Err.Number & " in NextOpenAccBank(). " & Err.Description)
End Function

Public Sub ChangeAccBankPass(ByVal UserIndex As Integer, ByVal Token As String, ByVal Pass As String)
'---------------------------------------------------------------------------------------
' Module    : modAccountBank
' Author    : Anagrama
' Date      : 21/08/2016
' Purpose   : Cambia la contraseña de la boveda si tiene el token correcto.
'---------------------------------------------------------------------------------------
On Error GoTo ErrHandler
  
    If GetAccountAnswer(UserList(UserIndex).AccountId) <> Token Then
        Call WriteConsoleMsg(UserIndex, "El Token es inválido.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If Not AsciiValidos(Pass, False, True) Then
        Call WriteConsoleMsg(UserIndex, "La contraseña no puede contener caracteres inválidos.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
        
    If Len(Pass) > 16 Then
        Call WriteConsoleMsg(UserIndex, "La contraseña no puede tener mas de 16 caracteres.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    BovedaCuenta(UserList(UserIndex).flags.AccountBank).Password = Pass
    Call WriteConsoleMsg(UserIndex, "La contraseña ha sido cambiada.", FontTypeNames.FONTTYPE_INFO)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ChangeAccBankPass de modAccountBank.bas")
End Sub
Sub IniciarDepositoAcc(ByVal UserIndex As Integer, ByVal Password As String)
'---------------------------------------------------------------------------------------
' Module    : modAccountBank
' Author    : Anagrama
' Date      : 21/08/2016
' Purpose   : Inicio del manejo de la boveda.
'---------------------------------------------------------------------------------------
On Error GoTo ErrHandler
    
    If BovedaCuenta(UserList(UserIndex).flags.AccountBank).Password <> vbNullString Then
        If Password = vbNullString Then
            Call WriteAccBankRequestPass(UserIndex)
        Else
            If BovedaCuenta(UserList(UserIndex).flags.AccountBank).Password = Password Then
                Call UpdateAccBankInv(True, UserIndex, 0)
                Call WriteUpdateUserStats(UserIndex)
                Call WriteAccBankInit(UserIndex)
                UserList(UserIndex).flags.Comerciando = TRADING_BANK
            Else
                Call WriteConsoleMsg(UserIndex, "La contraseña es incorrecta.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
    Else
        Call UpdateAccBankInv(True, UserIndex, 0)
        Call WriteUpdateUserStats(UserIndex)
        Call WriteAccBankInit(UserIndex)
        UserList(UserIndex).flags.Comerciando = TRADING_BANK
    End If
    
    Exit Sub
    
ErrHandler:
    Call LogError("Error " & Err.Number & " in IniciarDepositoAcc(). " & Err.Description)
End Sub

Sub UpdateAccBankInv(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal Slot As Byte)
'---------------------------------------------------------------------------------------
' Module    : modAccountBank
' Author    : Anagrama
' Date      : 21/08/2016
' Purpose   : Actualiza el o los slots de la boveda.
'---------------------------------------------------------------------------------------
On Error GoTo ErrHandler
  
    Dim LoopC As Long
    Dim ObjIndex As Integer
    Dim Amount As Integer
    Dim CanUse As Boolean
    
    With BovedaCuenta(UserList(UserIndex).flags.AccountBank)
        'Actualiza un solo slot
        If Not UpdateAll Then
            'Actualiza el inventario
            ObjIndex = .Object(Slot).ObjIndex
            If ObjIndex > 0 Then
                Amount = .Object(Slot).Amount
                CanUse = General.checkCanUseItem(UserIndex, ObjIndex)
            End If
            
            Call WriteChangeAccBankSlot(UserIndex, Slot, ObjIndex, Amount, CanUse)
        Else
            ' Limpio todos en el cliente
            Call WriteChangeAccBankSlot(UserIndex, 0, 0, 0, True)
            
            'Actualiza todos los slots
            For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
                
                ObjIndex = .Object(LoopC).ObjIndex
                If ObjIndex > 0 Then
                    CanUse = General.checkCanUseItem(UserIndex, ObjIndex)
                    Call WriteChangeAccBankSlot(UserIndex, CByte(LoopC), ObjIndex, .Object(LoopC).Amount, CanUse)
                End If
            Next LoopC
        End If
    End With

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub UpdateAccBankInv de modAccountBank.bas")
End Sub

Sub AccBankRetirarItem(ByVal UserIndex As Integer, ByVal SlotIndex As Integer, ByVal Cantidad As Integer)
'---------------------------------------------------------------------------------------
' Module    : modAccountBank
' Author    : Anagrama
' Date      : 21/08/2016
' Purpose   : Retira un item de la boveda.
'---------------------------------------------------------------------------------------
On Error GoTo ErrHandler

    Dim ObjIndex As Integer

    If Cantidad < 1 Then Exit Sub
    
    If UserList(UserIndex).flags.DueloIndex > 0 Then
        If DuelData.Duelo(UserList(UserIndex).flags.DueloIndex).Drop Then
            Call WriteConsoleMsg(UserIndex, "No puedes depositar objetos mientras tienes una petición de duelo por drop.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    End If
    
    Call WriteUpdateUserStats(UserIndex)
    
    With BovedaCuenta(UserList(UserIndex).flags.AccountBank)
        If .Object(SlotIndex).Amount > 0 Then
        
            If Cantidad > .Object(SlotIndex).Amount Then _
                Cantidad = .Object(SlotIndex).Amount
                
            ObjIndex = .Object(SlotIndex).ObjIndex
            
            'Agregamos el obj al inventario
            Call UserReciveAccBankObj(UserIndex, CInt(SlotIndex), Cantidad)
            
            If ObjData(ObjIndex).Log = 1 Then
                Call LogDesarrollo(UserList(UserIndex).Name & " retiró " & Cantidad & " " & _
                    ObjData(ObjIndex).Name & "[" & ObjIndex & "] de su boveda de cuenta.")
            End If
        End If
    End With
    
    Exit Sub
    
ErrHandler:

End Sub

Sub UserReciveAccBankObj(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, ByVal Cantidad As Integer)
'---------------------------------------------------------------------------------------
' Module    : modAccountBank
' Author    : Anagrama
' Date      : 21/08/2016
' Purpose   : Mete el item en el inventario del user.
'---------------------------------------------------------------------------------------
On Error GoTo ErrHandler
  
Dim Slot As Integer
Dim obji As Integer

With UserList(UserIndex)
    If BovedaCuenta(UserList(UserIndex).flags.AccountBank).Object(ObjIndex).Amount <= 0 Then Exit Sub
    
    obji = BovedaCuenta(UserList(UserIndex).flags.AccountBank).Object(ObjIndex).ObjIndex
    
    
    '¿Ya tiene un objeto de este tipo?
    Slot = 1
    Do Until .Invent.Object(Slot).ObjIndex = obji And _
       .Invent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS
        
        Slot = Slot + 1
        If Slot > .CurrentInventorySlots Then
            Exit Do
        End If
    Loop
    
    'Sino se fija por un slot vacio
    If Slot > .CurrentInventorySlots Then
        Slot = 1
        Do Until .Invent.Object(Slot).ObjIndex = 0
            Slot = Slot + 1

            If Slot > .CurrentInventorySlots Then
                Call WriteConsoleMsg(UserIndex, "No podés tener mas objetos.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        Loop
        .Invent.NroItems = .Invent.NroItems + 1
    End If
    
    'Mete el obj en el slot
    If .Invent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS Then
        'Menor que MAX_INV_OBJS
        .Invent.Object(Slot).ObjIndex = obji
        .Invent.Object(Slot).Amount = .Invent.Object(Slot).Amount + Cantidad
        
        Call QuitarAccBankInvItem(UserIndex, CByte(ObjIndex), Cantidad)
        Call UpdateUserInv(False, UserIndex, Slot)
    Else
        Call WriteConsoleMsg(UserIndex, "No podés tener mas objetos.", FontTypeNames.FONTTYPE_INFO)
    End If
End With

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub UserReciveAccBankObj de modAccountBank.bas")
End Sub

Sub QuitarAccBankInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Cantidad As Integer)
'---------------------------------------------------------------------------------------
' Module    : modAccountBank
' Author    : Anagrama
' Date      : 21/08/2016
' Purpose   : Retira un item de la boveda.
'---------------------------------------------------------------------------------------
On Error GoTo ErrHandler
  
Dim ObjIndex As Integer

With BovedaCuenta(UserList(UserIndex).flags.AccountBank)
    ObjIndex = .Object(Slot).ObjIndex

    'Quita un Obj
    .Object(Slot).Amount = .Object(Slot).Amount - Cantidad
    
    If .Object(Slot).Amount <= 0 Then
        .Object(Slot).ObjIndex = 0
        .Object(Slot).Amount = 0
    End If
    
    Call UpdateAccBankInv(False, UserIndex, Slot)
End With
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub QuitarAccBankInvItem de modAccountBank.bas")
End Sub

Sub UserDepositaAccBankItem(ByVal UserIndex As Integer, ByVal SlotIndex As Integer, ByVal Cantidad As Integer)
'---------------------------------------------------------------------------------------
' Module    : modAccountBank
' Author    : Anagrama
' Date      : 21/08/2016
' Purpose   : Deposita un item en la boveda.
' 2018-04-12: - IglorioN: Avoid newbie items from being deposited
'---------------------------------------------------------------------------------------
On Error GoTo ErrHandler

    Dim ObjIndex As Integer
    
    With UserList(UserIndex)
        If .Invent.Object(SlotIndex).Amount > 0 And Cantidad > 0 Then
        
            If Cantidad > .Invent.Object(SlotIndex).Amount Then _
                Cantidad = .Invent.Object(SlotIndex).Amount
            
            ObjIndex = .Invent.Object(SlotIndex).ObjIndex
            
            If ItemNewbie(ObjIndex) Then
                Call WriteConsoleMsg(UserIndex, "Los items newbie no pueden ser depositados en la bóveda de cuenta.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'Agregamos el obj que deposita al banco
            Call UserDejaAccBankObj(UserIndex, CInt(SlotIndex), Cantidad)
            
            If ObjData(ObjIndex).Log = 1 Then
                Call LogDesarrollo(UserList(UserIndex).Name & " depositó " & Cantidad & " " & _
                    ObjData(ObjIndex).Name & "[" & ObjIndex & "] en la boveda de cuenta")
            End If
        End If
    End With
    
    Exit Sub
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub UserDepositaAccBankItem de modAccountBank.bas")
End Sub

Sub UserDejaAccBankObj(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, ByVal Cantidad As Integer)
'---------------------------------------------------------------------------------------
' Module    : modAccountBank
' Author    : Anagrama
' Date      : 21/08/2016
' Purpose   : Le saca el item al user para dejarlo en la boveda.
'---------------------------------------------------------------------------------------
On Error GoTo ErrHandler
  
    Dim Slot As Integer
    Dim obji As Integer
    
    If Cantidad < 1 Then Exit Sub
    
    With BovedaCuenta(UserList(UserIndex).flags.AccountBank)
        obji = UserList(UserIndex).Invent.Object(ObjIndex).ObjIndex
        
        '¿Ya tiene un objeto de este tipo?
        Slot = 1
        Do Until .Object(Slot).ObjIndex = obji And _
            .Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS
            Slot = Slot + 1
            
            If Slot > MAX_BANCOINVENTORY_SLOTS Then
                Exit Do
            End If
        Loop
        
        'Sino se fija por un slot vacio antes del slot devuelto
        If Slot > MAX_BANCOINVENTORY_SLOTS Then
            Slot = 1
            Do Until .Object(Slot).ObjIndex = 0
                Slot = Slot + 1
                
                If Slot > MAX_BANCOINVENTORY_SLOTS_FIX Then
                    Call WriteConsoleMsg(UserIndex, "No tienes mas espacio en el banco!!", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            Loop

        End If
        
        If Slot > MAX_BANCOINVENTORY_SLOTS_FIX Then
            Call WriteConsoleMsg(UserIndex, "No tienes más espacio en el el banco de cuenta.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If .Object(Slot).Amount + Cantidad > MAX_INVENTORY_OBJS Then
            Call WriteConsoleMsg(UserIndex, "El banco de cuenta no puede cargar tantos objetos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
 
        .Object(Slot).ObjIndex = obji
        .Object(Slot).Amount = .Object(Slot).Amount + Cantidad
        
        Call QuitarUserInvItem(UserIndex, CByte(ObjIndex), Cantidad)
        Call UpdateAccBankInv(False, UserIndex, Slot)

    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub UserDejaAccBankObj de modAccountBank.bas")
End Sub

Public Sub MoveAccBankItem(ByVal UserIndex As Integer, ByVal nOriginalSlot As Integer, ByVal nNewSlot As Integer)
'---------------------------------------------------------------------------------------
' Module    : modAccountBank
' Author    : Anagrama
' Date      : 21/08/2016
' Purpose   : Mueve de lugar un item en la boveda.
'---------------------------------------------------------------------------------------
On Error GoTo ErrHandler
  
    Dim previousObject As UserOBJ
    
    With BovedaCuenta(UserList(UserIndex).flags.AccountBank)
        ' Save item in target slot.
        previousObject = .Object(nNewSlot)
        
        ' Store dragged item in the target slot.
        .Object(nNewSlot) = .Object(nOriginalSlot)
        
        ' Store replaced item in the original slot (if any).
        .Object(nOriginalSlot) = previousObject
    End With
    
    Call UpdateAccBankInv(False, UserIndex, nOriginalSlot)
    Call UpdateAccBankInv(False, UserIndex, nNewSlot)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub MoveAccBankItem de modAccountBank.bas")
End Sub


