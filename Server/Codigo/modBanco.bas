Attribute VB_Name = "modBanco"
'**************************************************************
' modBanco.bas - Handles the character's bank accounts.
'
' Implemented by Kevin Birmingham (NEB)
' kbneb@hotmail.com
'**************************************************************

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

Sub IniciarDeposito(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo ErrHandler

'Hacemos un Update del inventario del usuario
Call UpdateBanUserInv(True, UserIndex, 0)
'Actualizamos el dinero
Call WriteUpdateUserStats(UserIndex)
'Mostramos la ventana pa' comerciar y ver ladear la osamenta. jajaja
Call WriteBankInit(UserIndex)
UserList(UserIndex).flags.Comerciando = TRADING_BANK

ErrHandler:

End Sub

Sub UpdateBanUserInv(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal Slot As Byte)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    Dim LoopC As Long
    Dim ObjIndex As Integer
    Dim Amount As Integer
    Dim CanUse As Boolean
    
    With UserList(UserIndex)
        'Actualiza un solo slot
        If Not UpdateAll Then
            'Actualiza el inventario
            ObjIndex = .BancoInvent.Object(Slot).ObjIndex
            If ObjIndex > 0 Then
                Amount = .BancoInvent.Object(Slot).Amount
                CanUse = General.checkCanUseItem(UserIndex, ObjIndex)
            End If
            Call WriteChangeBankSlot(UserIndex, Slot, ObjIndex, Amount, CanUse)
        Else
            ' Limpio todos en el cliente
            Call WriteChangeBankSlot(UserIndex, 0, 0, 0, True)
            
            'Actualiza todos los slots
            For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
                ObjIndex = .BancoInvent.Object(LoopC).ObjIndex
                If ObjIndex > 0 Then
                    CanUse = General.checkCanUseItem(UserIndex, ObjIndex)
                    Call WriteChangeBankSlot(UserIndex, CByte(LoopC), ObjIndex, .BancoInvent.Object(LoopC).Amount, CanUse)
                End If
            Next LoopC
        End If
    End With

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub UpdateBanUserInv de modBanco.bas")
End Sub

Sub UserRetiraItem(ByVal UserIndex As Integer, ByVal SlotIndex As Integer, ByVal Cantidad As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo ErrHandler

    Dim ObjIndex As Integer
    
    Call WriteUpdateUserStats(UserIndex)

    If UserList(UserIndex).BancoInvent.Object(SlotIndex).Amount < 1 Then Exit Sub
    If Cantidad < 1 Then Exit Sub
    
    ' If the user is trying to whitdraw more items than the ones available, then it will whitdraw the amount it has in the bank
    If Cantidad > UserList(UserIndex).BancoInvent.Object(SlotIndex).Amount Then _
        Cantidad = UserList(UserIndex).BancoInvent.Object(SlotIndex).Amount
        
    ObjIndex = UserList(UserIndex).BancoInvent.Object(SlotIndex).ObjIndex
    
    'Agregamos el obj que compro al inventario
    Call UserReciveObj(UserIndex, CInt(SlotIndex), Cantidad)
    
    If ObjData(ObjIndex).Log = 1 Then
        Call LogDesarrollo(UserList(UserIndex).Name & " retiró " & Cantidad & " " & _
            ObjData(ObjIndex).Name & "[" & ObjIndex & "]")
    End If

ErrHandler:

End Sub

Sub UserReciveObj(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, ByVal Cantidad As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

Dim Slot As Integer
Dim obji As Integer

With UserList(UserIndex)
    If .BancoInvent.Object(ObjIndex).Amount <= 0 Then Exit Sub
    
    obji = .BancoInvent.Object(ObjIndex).ObjIndex
    
    
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
        
        Call QuitarBancoInvItem(UserIndex, CByte(ObjIndex), Cantidad)
        Call UpdateUserInv(False, UserIndex, Slot)
    Else
        Call WriteConsoleMsg(UserIndex, "No podés tener mas objetos.", FontTypeNames.FONTTYPE_INFO)
    End If
End With

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub UserReciveObj de modBanco.bas")
End Sub

Sub QuitarBancoInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Cantidad As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

Dim ObjIndex As Integer

With UserList(UserIndex)
    ObjIndex = .BancoInvent.Object(Slot).ObjIndex

    'Quita un Obj
    .BancoInvent.Object(Slot).Amount = .BancoInvent.Object(Slot).Amount - Cantidad
    
    If .BancoInvent.Object(Slot).Amount <= 0 Then
        .BancoInvent.NroItems = .BancoInvent.NroItems - 1
        .BancoInvent.Object(Slot).ObjIndex = 0
        .BancoInvent.Object(Slot).Amount = 0
    End If
    
    Call UpdateBanUserInv(False, UserIndex, Slot)
End With
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub QuitarBancoInvItem de modBanco.bas")
End Sub

Sub UserDepositaItem(ByVal UserIndex As Integer, ByVal SlotIndex As Integer, ByVal Cantidad As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo ErrHandler

    Dim ObjIndex As Integer
    
    If UserList(UserIndex).flags.DueloIndex > 0 Then
        If DuelData.Duelo(UserList(UserIndex).flags.DueloIndex).Drop Then
            Call WriteConsoleMsg(UserIndex, "No puedes depositar objetos mientras tienes una petición de duelo por drop.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    End If
    
    With UserList(UserIndex)
    
        If .Invent.Object(SlotIndex).Amount < 1 Then Exit Sub
        If Cantidad < 1 Then Exit Sub
        
        If .Invent.Object(SlotIndex).ObjIndex > 0 Then
            If ObjData(.Invent.Object(SlotIndex).ObjIndex).ObjType = otQuest Then
                Call WriteConsoleMsg(UserIndex, "No puedes depositar este objeto", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If

        If Cantidad > .Invent.Object(SlotIndex).Amount Then _
            Cantidad = .Invent.Object(SlotIndex).Amount
        
        ObjIndex = .Invent.Object(SlotIndex).ObjIndex
        
        'Agregamos el obj que deposita al banco
        Call UserDejaObj(UserIndex, CInt(SlotIndex), Cantidad)
        
        If ObjData(ObjIndex).Log = 1 Then
            Call LogDesarrollo(.Name & " depositó " & Cantidad & " " & _
                ObjData(ObjIndex).Name & "[" & ObjIndex & "]")
        End If
      
    End With
    
ErrHandler:
End Sub

Sub UserDejaObj(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, ByVal Cantidad As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    Dim Slot As Integer
    Dim obji As Integer
    
    If Cantidad < 1 Then Exit Sub
    
    With UserList(UserIndex)
        obji = .Invent.Object(ObjIndex).ObjIndex
        
        '¿Ya tiene un objeto de este tipo?
        Slot = 1
        Do Until .BancoInvent.Object(Slot).ObjIndex = obji And _
            .BancoInvent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS
            Slot = Slot + 1
            
            If Slot > MAX_BANCOINVENTORY_SLOTS_FIX Then
                Exit Do
            End If
        Loop
        
        'Sino se fija por un slot vacio antes del slot devuelto
        If Slot > MAX_BANCOINVENTORY_SLOTS_FIX Then
            Slot = 1
            Do Until .BancoInvent.Object(Slot).ObjIndex = 0
                Slot = Slot + 1
                
                If Slot > MAX_BANCOINVENTORY_SLOTS_FIX Then
                    Call WriteConsoleMsg(UserIndex, "No tienes mas espacio en el banco!!", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            Loop
            
            .BancoInvent.NroItems = .BancoInvent.NroItems + 1
        End If
        
        If Slot > MAX_BANCOINVENTORY_SLOTS_FIX Then
            Call WriteConsoleMsg(UserIndex, "No tienes más espacio en el el banco.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If .BancoInvent.Object(Slot).Amount + Cantidad > MAX_INVENTORY_OBJS Then
            Call WriteConsoleMsg(UserIndex, "El banco no puede cargar tantos objetos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
                
        .BancoInvent.Object(Slot).ObjIndex = obji
        .BancoInvent.Object(Slot).Amount = .BancoInvent.Object(Slot).Amount + Cantidad
        
        Call QuitarUserInvItem(UserIndex, CByte(ObjIndex), Cantidad)
        Call UpdateBanUserInv(False, UserIndex, Slot)

    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub UserDejaObj de modBanco.bas")
End Sub

Sub SendUserBovedaTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

On Error Resume Next
Dim J As Integer

Call WriteConsoleMsg(sendIndex, UserList(UserIndex).Name, FontTypeNames.FONTTYPE_INFO)
Call WriteConsoleMsg(sendIndex, "Tiene " & UserList(UserIndex).BancoInvent.NroItems & " objetos.", FontTypeNames.FONTTYPE_INFO)

For J = 1 To MAX_BANCOINVENTORY_SLOTS
    If UserList(UserIndex).BancoInvent.Object(J).ObjIndex > 0 Then
        Call WriteConsoleMsg(sendIndex, "Objeto " & J & " " & ObjData(UserList(UserIndex).BancoInvent.Object(J).ObjIndex).Name & " Cantidad:" & UserList(UserIndex).BancoInvent.Object(J).Amount, FontTypeNames.FONTTYPE_INFO)
    End If
Next

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SendUserBovedaTxt de modBanco.bas")
End Sub

Sub SendUserBovedaTxtFromChar(ByVal sendIndex As Integer, ByVal charName As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    Dim J As Integer
    Dim CharFile As String, Tmp As String
    Dim ObjInd As Long, ObjCant As Long
    Dim UserId As Long
    
    UserId = GetUserID(charName)
    
    If UserId <> 0 Then
        Call SendUserBovedaTxtFromDB(sendIndex, UserId, charName)
    Else
        Call WriteConsoleMsg(sendIndex, "Usuario inexistente: " & charName, FontTypeNames.FONTTYPE_INFO)
    End If

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SendUserBovedaTxtFromChar de modBanco.bas")
End Sub

Public Sub MoveBankItem(ByVal UserIndex As Integer, ByVal nOriginalSlot As Integer, ByVal nNewSlot As Integer)
'***************************************************
'Author: D'Artagnan
'Last Modification: 07/08/2014
'
'***************************************************
On Error GoTo ErrHandler
  
    Dim previousObject As UserOBJ
    
    With UserList(UserIndex)
        ' Save item in target slot.
        previousObject = .BancoInvent.Object(nNewSlot)
        
        ' Store dragged item in the target slot.
        .BancoInvent.Object(nNewSlot) = .BancoInvent.Object(nOriginalSlot)
        
        ' Store replaced item in the original slot (if any).
        .BancoInvent.Object(nOriginalSlot) = previousObject
    End With
    
    Call UpdateBanUserInv(False, UserIndex, nOriginalSlot)
    Call UpdateBanUserInv(False, UserIndex, nNewSlot)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub MoveBankItem de modBanco.bas")
End Sub
