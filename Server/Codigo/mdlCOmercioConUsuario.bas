Attribute VB_Name = "mdlCOmercioConUsuario"
'**************************************************************
' mdlComercioConUsuarios.bas - Allows players to commerce between themselves.
'
' Designed and implemented by Alejandro Santos (AlejoLP)
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

'[Alejo]
Option Explicit

Public Const MAX_OFFER_SLOTS As Integer = 30
Public Const GOLD_OFFER_SLOT As Integer = MAX_OFFER_SLOTS + 1

Public Type tCOmercioUsuario
    'DestUsu As Integer 'El otro Usuario
    DestNick As String
    Objeto(1 To MAX_OFFER_SLOTS) As Integer 'Indice de los objetos que se desea dar
    GoldAmount As Long
    
    cant(1 To MAX_OFFER_SLOTS) As Long 'Cuantos objetos desea dar
    Acepto As Boolean
    Confirmo As Boolean
End Type

Private Type tOfferItem
    ObjIndex As Integer
    Amount As Long
End Type

'origen: origen de la transaccion, originador del comando
'destino: receptor de la transaccion
Public Sub IniciarComercioConUsuario(ByVal Origen As Integer, ByVal Destino As Integer)
'***************************************************
'Autor: Unkown
'Last Modification: 25/12/2014
'25/12/2014: D'Artagnan - Store source user in a new flag.
'***************************************************
    On Error GoTo ErrHandler
    
    'Si ambos pusieron /comerciar entonces
    If getTradingUser(Origen) = Destino And _
       getTradingUser(Destino) = Origen Then
       
       'TODO: Hay que revisar si se puede dar esta situacion que no se puede expresar con el nuevo sistema, en caso de que se de habra que buscar solucion -.-
       'If isTradingWithUser(Origen) Or isTradingWithUser(Destino) Then
       '     Call WriteConsoleMsg(Origen, "No puedes comerciar en este momento", FontTypeNames.FONTTYPE_TALK)
       '     Call WriteConsoleMsg(Destino, "No puedes comerciar en este momento", FontTypeNames.FONTTYPE_TALK)
       '     Exit Sub
       ' End If
        
        'TODO: Es necesario esto O.o?
        'Actualiza el inventario del usuario
        'Call UpdateUserInv(True, Origen, 0)
        
        UserList(Destino).flags.nCommerceSourceUser = 0
        
        'Decirle al origen que abra la ventanita.
        Call WriteUserCommerceInit(Origen)
        UserList(Origen).flags.Comerciando = (Not Destino)
    
        'TODO: Es necesario esto O.o?
        'Actualiza el inventario del usuario
        'Call UpdateUserInv(True, Destino, 0)
        
        'Decirle al origen que abra la ventanita.
        Call WriteUserCommerceInit(Destino)
        UserList(Destino).flags.Comerciando = (Not Origen)
    
        'Call EnviarObjetoTransaccion(Origen)
    Else
        'Es el primero que comercia ?
        Call WriteConsoleMsg(Destino, UserList(Origen).Name & " desea comerciar. Si deseas aceptar, escribe /COMERCIAR.", FontTypeNames.FONTTYPE_TALK)
        UserList(Destino).flags.TargetUser = Origen
        UserList(Destino).flags.nCommerceSourceUser = Origen
        
    End If

    Exit Sub
ErrHandler:
        Call LogError("Error en IniciarComercioConUsuario: " & Err.Description)
End Sub

Public Sub EnviarOfertaOro(ByVal UserIndex As Integer)
'***************************************************
'Autor: Unkown
'Last Modification: 25/11/2009
'Sends the offer change to the other trading user
'25/11/2009: ZaMa - Implementado nuevo sistema de comercio con ofertas variables.
'***************************************************
On Error GoTo ErrHandler
    Dim OtherUserIndex As Integer
    
    OtherUserIndex = getTradingUser(UserIndex)
    
    Call WriteChangeUserTradeGold(UserIndex, UserList(OtherUserIndex).ComUsu.GoldAmount)

    Exit Sub

ErrHandler:
    LogError "Error en EnviarOfertaOro. Error: " & Err.Description & ". UserIndex: " & UserIndex & ". OtherIndex: " & OtherUserIndex
End Sub

Public Sub EnviarOferta(ByVal UserIndex As Integer, ByVal OfferSlot As Byte)
'***************************************************
'Autor: Unkown
'Last Modification: 25/11/2009
'Sends the offer change to the other trading user
'25/11/2009: ZaMa - Implementado nuevo sistema de comercio con ofertas variables.
'***************************************************
On Error GoTo ErrHandler

    Dim ObjIndex As Integer
    Dim ObjAmount As Long
    Dim OtherUserIndex As Integer
    Dim CanUse As Boolean
    
    OtherUserIndex = getTradingUser(UserIndex)
    
    With UserList(OtherUserIndex)
        If OfferSlot = GOLD_OFFER_SLOT Then
            ObjIndex = ConstantesItems.Oro
            ObjAmount = .ComUsu.GoldAmount
            CanUse = True
        Else
            ObjIndex = .ComUsu.Objeto(OfferSlot)
            ObjAmount = .ComUsu.cant(OfferSlot)
            CanUse = General.checkCanUseItem(UserIndex, ObjIndex)
        End If
    End With
   
    Call WriteChangeUserTradeSlot(UserIndex, OfferSlot, ObjIndex, ObjAmount, CanUse)

    Exit Sub

ErrHandler:
    LogError "Error en EnviarOferta. Error: " & Err.Description & ". UserIndex: " & UserIndex & ". OfferSlot: " & OfferSlot
End Sub

Public Sub FinComerciarUsu(ByVal UserIndex As Integer, Optional ByVal bResetTargetUser As Boolean = False)
'***************************************************
'Autor: Unkown
'Last Modification: 29/11/2014
'25/11/2009: ZaMa - Limpio los arrays (por el nuevo sistema)
'29/11/2014: D'Artagnan - New parameter: bResetTargetUser.
'***************************************************
On Error GoTo ErrHandler

    Dim I As Long
    
    With UserList(UserIndex)
        If getTradingUser(UserIndex) > 0 Then
            Call WriteUserCommerceEnd(UserIndex)
        End If
        
        .ComUsu.Acepto = False
        .ComUsu.Confirmo = False
        
        For I = 1 To MAX_OFFER_SLOTS
            .ComUsu.cant(I) = 0
            .ComUsu.Objeto(I) = 0
        Next I
        
        .ComUsu.GoldAmount = 0
        .ComUsu.DestNick = vbNullString
        .flags.Comerciando = 0
        .flags.nCommerceSourceUser = 0
        If bResetTargetUser Then .flags.TargetUser = 0
    End With
    
    Exit Sub

ErrHandler:
    LogError "Error en FinComerciarUsu. Error: " & Err.Description & ". UserIndex: " & UserIndex
End Sub

Public Sub AceptarComercioUsu(ByVal UserIndex As Integer)
'***************************************************
'Autor: Unkown
'Last Modification: 06/05/2010
'25/11/2009: ZaMa - Ahora se traspasan hasta 5 items + oro al comerciar
'06/05/2010: ZaMa - Ahora valida si los usuarios tienen los items que ofertan.
'***************************************************
On Error GoTo ErrHandler
  
    Dim TradingObj As Obj
    Dim OtroUserIndex As Integer
    Dim OfferSlot As Integer

    UserList(UserIndex).ComUsu.Acepto = True
    
    OtroUserIndex = getTradingUser(UserIndex)
    
    ' User valido?
    If OtroUserIndex <= 0 Or OtroUserIndex > MaxUsers Then
        Call FinComerciarUsu(UserIndex)
        Exit Sub
    End If
    
    ' Acepto el otro?
    If UserList(OtroUserIndex).ComUsu.Acepto = False Then
        Exit Sub
    End If
    
    ' Aceptaron ambos, chequeo que tengan los items que ofertaron
    If Not HasOfferedItems(UserIndex) Then
        
        Call WriteConsoleMsg(UserIndex, "¡¡¡El comercio se canceló porque no posees los ítems que ofertaste!!!", FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(OtroUserIndex, "¡¡¡El comercio se canceló porque " & UserList(UserIndex).Name & " no posee los ítems que ofertó!!!", FontTypeNames.FONTTYPE_FIGHT)
        
        Call FinComerciarUsu(UserIndex)
        Call FinComerciarUsu(OtroUserIndex)

        Exit Sub
        
    ElseIf Not HasOfferedItems(OtroUserIndex) Then
        
        Call WriteConsoleMsg(UserIndex, "¡¡¡El comercio se canceló porque " & UserList(OtroUserIndex).Name & " no posee los ítems que ofertó!!!", FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(OtroUserIndex, "¡¡¡El comercio se canceló porque no posees los ítems que ofertaste!!!", FontTypeNames.FONTTYPE_FIGHT)
        
        Call FinComerciarUsu(UserIndex)
        Call FinComerciarUsu(OtroUserIndex)

        Exit Sub
        
    End If
    
    ' Envio los items a quien corresponde
    For OfferSlot = 1 To MAX_OFFER_SLOTS + 1
        
        ' Items del 1er usuario
        With UserList(UserIndex)
            ' Le pasa el oro
            If OfferSlot = GOLD_OFFER_SLOT Then
                ' Quito la cantidad de oro ofrecida
                .Stats.GLD = .Stats.GLD - .ComUsu.GoldAmount
                ' Log
                If .ComUsu.GoldAmount > MIN_GOLD_AMOUNT_LOG Then Call LogDesarrollo(.Name & " soltó oro en comercio seguro con " & UserList(OtroUserIndex).Name & ". Cantidad: " & .ComUsu.GoldAmount)
                ' Update Usuario
                Call WriteUpdateGold(UserIndex)
                ' Se la doy al otro
                UserList(OtroUserIndex).Stats.GLD = UserList(OtroUserIndex).Stats.GLD + .ComUsu.GoldAmount
                ' Update Otro Usuario
                Call WriteUpdateGold(OtroUserIndex)
                
            ' Le pasa lo ofertado de los slots con items
            ElseIf .ComUsu.Objeto(OfferSlot) > 0 Then
                TradingObj.ObjIndex = .ComUsu.Objeto(OfferSlot)
                TradingObj.Amount = .ComUsu.cant(OfferSlot)
                                
                'Quita el objeto y se lo da al otro
                If Not MeterItemEnInventario(OtroUserIndex, TradingObj) Then
                    Call TirarItemAlPiso(UserList(OtroUserIndex).Pos, TradingObj)
                End If
            
                Call QuitarObjetos(TradingObj.ObjIndex, TradingObj.Amount, UserIndex)
                
                'Es un Objeto que tenemos que loguear? Pablo (ToxicWaste) 07/09/07
                If (ObjData(TradingObj.ObjIndex).Log = 1 Or (ObjData(TradingObj.ObjIndex).ObjType = eOBJType.otLlaves)) Then
                    Call LogDesarrollo(.Name & " le pasó en comercio seguro a " & UserList(OtroUserIndex).Name & " " & TradingObj.Amount & " " & ObjData(TradingObj.ObjIndex).Name)
            
                'Es mucha cantidad?
                ElseIf TradingObj.Amount > MIN_AMOUNT_LOG Then
                    'Si no es de los prohibidos de loguear, lo logueamos.
                    If ObjData(TradingObj.ObjIndex).NoLog <> 1 Then
                        Call LogDesarrollo(.Name & " le pasó en comercio seguro a " & UserList(OtroUserIndex).Name & " " & TradingObj.Amount & " " & ObjData(TradingObj.ObjIndex).Name)
                    End If
                
                ElseIf (TradingObj.Amount * ObjData(TradingObj.ObjIndex).Valor) >= MIN_VALUE_LOG Then
                    'Si no es de los prohibidos de loguear, lo logueamos.
                    If ObjData(TradingObj.ObjIndex).NoLog <> 1 Then
                        Call LogDesarrollo(.Name & " le pasó en comercio seguro a " & UserList(OtroUserIndex).Name & " " & TradingObj.Amount & " " & ObjData(TradingObj.ObjIndex).Name)
                    End If
                End If
            End If
        End With
        
        ' Items del 2do usuario
        With UserList(OtroUserIndex)
            ' Le pasa el oro
            If OfferSlot = GOLD_OFFER_SLOT Then
                ' Quito la cantidad de oro ofrecida
                .Stats.GLD = .Stats.GLD - .ComUsu.GoldAmount
                ' Log
                If .ComUsu.GoldAmount > MIN_GOLD_AMOUNT_LOG Then Call LogDesarrollo(.Name & " soltó oro en comercio seguro con " & UserList(UserIndex).Name & ". Cantidad: " & .ComUsu.GoldAmount)
                ' Update Usuario
                Call WriteUpdateGold(OtroUserIndex)
                'y se la doy al otro
                UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + .ComUsu.GoldAmount
                If .ComUsu.GoldAmount > MIN_GOLD_AMOUNT_LOG Then Call LogDesarrollo(UserList(UserIndex).Name & " recibió oro en comercio seguro con " & .Name & ". Cantidad: " & .ComUsu.GoldAmount)
                ' Update Otro Usuario
                Call WriteUpdateGold(UserIndex)
                
            ' Le pasa la oferta de los slots con items
            ElseIf .ComUsu.Objeto(OfferSlot) > 0 Then
                TradingObj.ObjIndex = .ComUsu.Objeto(OfferSlot)
                TradingObj.Amount = .ComUsu.cant(OfferSlot)
                                
                'Quita el objeto y se lo da al otro
                If Not MeterItemEnInventario(UserIndex, TradingObj) Then
                    Call TirarItemAlPiso(UserList(UserIndex).Pos, TradingObj)
                End If
            
                Call QuitarObjetos(TradingObj.ObjIndex, TradingObj.Amount, OtroUserIndex)
                
                'Es un Objeto que tenemos que loguear? Pablo (ToxicWaste) 07/09/07
                If ObjData(TradingObj.ObjIndex).Log = 1 Or (ObjData(TradingObj.ObjIndex).ObjType = eOBJType.otLlaves) Then
                    Call LogDesarrollo(.Name & " le pasó en comercio seguro a " & UserList(UserIndex).Name & " " & TradingObj.Amount & " " & ObjData(TradingObj.ObjIndex).Name)
                
            
                'Es mucha cantidad?
                ElseIf TradingObj.Amount > MIN_AMOUNT_LOG Then
                    'Si no es de los prohibidos de loguear, lo logueamos.
                    If ObjData(TradingObj.ObjIndex).NoLog <> 1 Then
                        Call LogDesarrollo(.Name & " le pasó en comercio seguro a " & UserList(UserIndex).Name & " " & TradingObj.Amount & " " & ObjData(TradingObj.ObjIndex).Name)
                    End If
                
                ElseIf (TradingObj.Amount * ObjData(TradingObj.ObjIndex).Valor) >= MIN_VALUE_LOG Then
                    'Si no es de los prohibidos de loguear, lo logueamos.
                    If ObjData(TradingObj.ObjIndex).NoLog <> 1 Then
                        Call LogDesarrollo(.Name & " le pasó en comercio seguro a " & UserList(UserIndex).Name & " " & TradingObj.Amount & " " & ObjData(TradingObj.ObjIndex).Name)
                    End If
                End If
            End If
        End With
        
    Next OfferSlot

    ' End Trade
    Call FinComerciarUsu(UserIndex)
    Call FinComerciarUsu(OtroUserIndex)

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub AceptarComercioUsu de mdlCOmercioConUsuario.bas")
End Sub

Public Sub AgregarOferta(ByVal UserIndex As Integer, ByVal OfferSlot As Byte, ByVal ObjIndex As Integer, ByVal Amount As Long, ByVal IsGold As Boolean)
'***************************************************
'Autor: ZaMa
'Last Modification: 24/11/2009
'Adds gold or items to the user's offer
'***************************************************
On Error GoTo ErrHandler

    With UserList(UserIndex).ComUsu
        ' Si ya confirmo su oferta, no puede cambiarla!
        If Not .Confirmo Then
            If IsGold Then
            ' Agregamos (o quitamos) mas oro a la oferta
                .GoldAmount = .GoldAmount + Amount
                
                ' Imposible que pase, pero por las dudas..
                If .GoldAmount < 0 Then .GoldAmount = 0
            Else
            ' Agreamos (o quitamos) el item y su cantidad en el slot correspondiente
                ' Si es 0 estoy modificando la cantidad, no agregando
                If ObjIndex > 0 Then .Objeto(OfferSlot) = ObjIndex
                .cant(OfferSlot) = .cant(OfferSlot) + Amount
                
                'Quitó todos los items de ese tipo
                If .cant(OfferSlot) <= 0 Then
                    ' Removemos el objeto para evitar conflictos
                    .Objeto(OfferSlot) = 0
                    .cant(OfferSlot) = 0
                End If
            End If
        End If
    End With
    
    Exit Sub
ErrHandler:
    LogError "Error en AgregarOferta. Error: " & Err.Description & ". UserIndex: " & UserIndex
End Sub

Public Function PuedeSeguirComerciando(ByVal UserIndex As Integer) As Boolean
'***************************************************
'Autor: ZaMa
'Last Modification: 24/11/2009
'Validates wether the conditions for the commerce to keep going are satisfied
'***************************************************
On Error GoTo ErrHandler

    Dim OtroUserIndex As Integer
    Dim ComercioInvalido As Boolean
    
    With UserList(UserIndex)
        OtroUserIndex = getTradingUser(UserIndex)
        
        ' Usuario valido?
        If (OtroUserIndex <= 0) Or (OtroUserIndex > MaxUsers) Then
            ComercioInvalido = True
        End If
        
        If Not ComercioInvalido Then
            ' Estan logueados?
            If (UserList(OtroUserIndex).flags.UserLogged = False) Or (.flags.UserLogged = False) Then
                ComercioInvalido = True
            End If
        End If
        
        If Not ComercioInvalido Then
            ' Se estan comerciando el uno al otro?
            If getTradingUser(OtroUserIndex) <> UserIndex Then
                ComercioInvalido = True
            End If
        End If
        
        If Not ComercioInvalido Then
            ' El nombre del otro es el mismo que al que le comercio?
            If UserList(OtroUserIndex).Name <> .ComUsu.DestNick Then
                ComercioInvalido = True
            End If
        End If
        
        If Not ComercioInvalido Then
            ' Mi nombre  es el mismo que al que el le comercia?
            If .Name <> UserList(OtroUserIndex).ComUsu.DestNick Then
                ComercioInvalido = True
            End If
        End If
        
        If Not ComercioInvalido Then
            ' Esta vivo?
            If UserList(OtroUserIndex).flags.Muerto = 1 Then
                ComercioInvalido = True
            End If
        End If
        
        ' Fin del comercio
        If ComercioInvalido = True Then
            Call FinComerciarUsu(UserIndex)
            
            If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
                Call FinComerciarUsu(OtroUserIndex)
            End If
            
            Exit Function
        End If
    End With

    PuedeSeguirComerciando = True
    
    Exit Function

ErrHandler:
    LogError "Error en PuedeSeguirComerciando. Error: " & Err.Description & ". UserIndex: " & UserIndex
End Function

Private Function HasOfferedItems(ByVal UserIndex As Integer) As Boolean
'***************************************************
'Autor: ZaMa
'Last Modification: 05/06/2010
'Checks whether the user has the offered items in his inventory or not.
'***************************************************
On Error GoTo ErrHandler
  

    Dim OfferedItems(MAX_OFFER_SLOTS - 1) As tOfferItem
    Dim Slot As Long
    Dim SlotAux As Long
    Dim SlotCount As Long
    
    Dim ObjIndex As Integer
    
    With UserList(UserIndex).ComUsu
        
        ' Agrupo los items que son iguales
        For Slot = 1 To MAX_OFFER_SLOTS
                    
            ObjIndex = .Objeto(Slot)
            
            If ObjIndex > 0 Then
            
                For SlotAux = 0 To SlotCount - 1
                    
                    If ObjIndex = OfferedItems(SlotAux).ObjIndex Then
                        ' Son iguales, aumento la cantidad
                        OfferedItems(SlotAux).Amount = OfferedItems(SlotAux).Amount + .cant(Slot)
                        Exit For
                    End If
                    
                Next SlotAux
                
                ' No encontro otro igual, lo agrego
                If SlotAux = SlotCount Then
                    OfferedItems(SlotCount).ObjIndex = ObjIndex
                    OfferedItems(SlotCount).Amount = .cant(Slot)
                    
                    SlotCount = SlotCount + 1
                End If
                
            End If
            
        Next Slot
        
        ' Chequeo que tengan la cantidad en el inventario
        For Slot = 0 To SlotCount - 1
            If Not HasEnoughItems(UserIndex, OfferedItems(Slot).ObjIndex, OfferedItems(Slot).Amount) Then Exit Function
        Next Slot
        
        ' Compruebo que tenga el oro que oferta
        If UserList(UserIndex).Stats.GLD < .GoldAmount Then Exit Function
        
    End With
    
    HasOfferedItems = True

  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function HasOfferedItems de mdlCOmercioConUsuario.bas")
End Function
