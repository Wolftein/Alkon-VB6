Attribute VB_Name = "modSistemaComercio"
'*****************************************************
'Sistema de Comercio para Argentum Online
'Programado por Nacho (Integer)
'integer-x@hotmail.com
'*****************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'(at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'**************************************************************************

Option Explicit

Enum eModoComercio
    Compra = 1
    Venta = 2
End Enum


''
' Makes a trade. (Buy or Sell)
'
' @param Modo The trade type (sell or buy)
' @param UserIndex Specifies the index of the user
' @param NpcIndex specifies the index of the npc
' @param Slot Specifies which slot are you trying to sell / buy
' @param Cantidad Specifies how many items in that slot are you trying to sell / buy
Public Sub Comercio(ByVal Modo As eModoComercio, ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal Slot As Integer, ByVal Cantidad As Integer)
'*************************************************
'Author: Nacho (Integer)
'Last modified: 10/04/2015
'27/07/08 (MarKoxX) | New changes in the way of trading (now when you buy it rounds to ceil and when you sell it rounds to floor)
'  - 06/13/08 (NicoNZ)
'07/06/2010: ZaMa - Los objetos se loguean si superan la cantidad de 1k (antes era solo si eran 1k).
'10/04/2012: ZaMa - Muevo una conidición para que los sastres reales compren cualquier objeto real.
'10/04/2015: D'Artagnan - The "commerce" skill affects no longer the item price.
'*************************************************
On Error GoTo ErrHandler
  
    Dim Precio As Long
    Dim Objeto As Obj
    Dim SalePrice As Long
    Dim UserSlot As Byte
    
    If Cantidad < 1 Or Slot < 1 Then Exit Sub
    
    If Modo = eModoComercio.Compra Then
        If Slot > MAX_INVENTORY_SLOTS Then
            Exit Sub
        ElseIf Cantidad > MAX_INVENTORY_OBJS Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserList(UserIndex).Name & " ha sido baneado por el sistema anti-cheats.", FontTypeNames.FONTTYPE_FIGHT))
            Call LogBan(UserList(UserIndex).Name, "Sistema Anti Cheats", "Intentar hackear el sistema de comercio. Quiso comprar demasiados ítems:" & Cantidad)
            UserList(UserIndex).flags.Ban = 1
            Call DisconnectWithMessage(UserIndex, "Has sido baneado por el Sistema AntiCheat.")
            Exit Sub
        ElseIf Not Npclist(NpcIndex).Invent.Object(Slot).Amount > 0 Then
            Exit Sub
        End If
        
        If Cantidad > Npclist(NpcIndex).Invent.Object(Slot).Amount Then Cantidad = Npclist(NpcIndex).Invent.Object(Slot).Amount
        
        Objeto.Amount = Cantidad
        Objeto.ObjIndex = Npclist(NpcIndex).Invent.Object(Slot).ObjIndex
        
        ' El precio ya no se redondea
        Precio = ObjData(Npclist(NpcIndex).Invent.Object(Slot).ObjIndex).Valor * Cantidad

        If UserList(UserIndex).Stats.GLD < Precio Then
            Call WriteConsoleMsg(UserIndex, "No tienes suficiente dinero.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If MeterItemEnInventario(UserIndex, Objeto) = False Then
            'Call WriteConsoleMsg(UserIndex, "No puedes cargar mas objetos.", FontTypeNames.FONTTYPE_INFO)
            'Call WriteTradeOK(UserIndex)
            Exit Sub
        End If
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Precio
        
        SalePrice = ObjData(Objeto.ObjIndex).SalePrice
        UserSlot = GetObjPosition(UserIndex, Objeto.ObjIndex)
        Call WriteSetUserSalePrice(UserIndex, UserSlot, SalePrice)
        
        Call QuitarNpcInvItem(NpcIndex, CByte(Slot), Cantidad)
        Call UpdateNpcInv(NpcIndex, CByte(Slot))
        
        'Bien, ahora logueo de ser necesario. Pablo (ToxicWaste) 07/09/07
        'Es un Objeto que tenemos que loguear?
        If ObjData(Objeto.ObjIndex).Log = 1 Then
            Call LogDesarrollo(UserList(UserIndex).Name & " compró del NPC " & Objeto.Amount & " " & ObjData(Objeto.ObjIndex).Name)
        ElseIf Objeto.Amount >= 1000 Then 'Es mucha cantidad?
            'Si no es de los prohibidos de loguear, lo logueamos.
            If ObjData(Objeto.ObjIndex).NoLog <> 1 Then
                Call LogDesarrollo(UserList(UserIndex).Name & " compró del NPC " & Objeto.Amount & " " & ObjData(Objeto.ObjIndex).Name)
            End If
        End If
        
        'Agregado para que no se vuelvan a vender las llaves si se recargan los .dat.
        If ObjData(Objeto.ObjIndex).ObjType = otLlaves Then
            Call WriteVar(DatPath & "NPCs.dat", "NPC" & Npclist(NpcIndex).Numero, "obj" & Slot, Objeto.ObjIndex & "-0")
            Call logVentaCasa(UserList(UserIndex).Name & " compró " & ObjData(Objeto.ObjIndex).Name)
        End If
        
    ElseIf Modo = eModoComercio.Venta Then
        
        If Cantidad > UserList(UserIndex).Invent.Object(Slot).Amount Then Cantidad = UserList(UserIndex).Invent.Object(Slot).Amount
        
        Objeto.Amount = Cantidad
        Objeto.ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
        Precio = ObjData(Objeto.ObjIndex).SalePrice * Cantidad
        
        If Objeto.ObjIndex = 0 Then
            Exit Sub
        
        ElseIf IsSecondaryArmour(Objeto.ObjIndex) Or ObjData(Objeto.ObjIndex).Intransferible = 1 Or _
               ObjData(Objeto.ObjIndex).NoComerciable = 1 Then
            Call WriteConsoleMsg(UserIndex, "No puedes vender este tipo de objeto.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        ElseIf (Npclist(NpcIndex).TipoItems <> ObjData(Objeto.ObjIndex).ObjType And Npclist(NpcIndex).TipoItems <> eOBJType.otCualquiera) Or Objeto.ObjIndex = ConstantesItems.Oro Then
            If ObjData(Objeto.ObjIndex).Real = 0 And ObjData(Objeto.ObjIndex).Caos = 0 Then
                Call WriteConsoleMsg(UserIndex, "Lo siento, no estoy interesado en este tipo de objetos.", FontTypeNames.FONTTYPE_INFO)
                'Call WriteTradeOK(UserIndex)
                Exit Sub
            End If
        End If
        
        If ObjData(Objeto.ObjIndex).Real <> 0 Then
            If Npclist(NpcIndex).Name <> "SR" Then
                Call WriteConsoleMsg(UserIndex, "Las armaduras y equipo del ejército real sólo pueden ser vendidos a los sastres reales.", FontTypeNames.FONTTYPE_INFO)
                'Call WriteTradeOK(UserIndex)
                Exit Sub
            End If
        ElseIf ObjData(Objeto.ObjIndex).Caos <> 0 Then
            If Npclist(NpcIndex).Name <> "SC" Then
                Call WriteConsoleMsg(UserIndex, "Las armaduras y equipo de la legión oscura sólo pueden ser vendidos a los sastres del demonio.", FontTypeNames.FONTTYPE_INFO)
                'Call WriteTradeOK(UserIndex)
                Exit Sub
            End If
        ElseIf UserList(UserIndex).Invent.Object(Slot).Amount < 0 Or Cantidad = 0 Then
            Exit Sub
        ElseIf Slot < LBound(UserList(UserIndex).Invent.Object()) Or Slot > UBound(UserList(UserIndex).Invent.Object()) Then
            Exit Sub
        ElseIf UserList(UserIndex).flags.Privilegios And PlayerType.Consejero Then
            Call WriteConsoleMsg(UserIndex, "No puedes vender ítems.", FontTypeNames.FONTTYPE_WARNING)
            'Call WriteTradeOK(UserIndex)
            Exit Sub
        ElseIf UserList(UserIndex).flags.DueloIndex > 0 Then
            If DuelData.Duelo(UserList(UserIndex).flags.DueloIndex).Drop Then
                Call WriteConsoleMsg(UserIndex, "No puedes vender objetos mientras tienes una petición de duelo por drop.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        ElseIf UserList(UserIndex).Stats.GLD >= ConstantesBalance.MaxOro Then
            Call WriteConsoleMsg(UserIndex, "No puedes vender mas objetos ya que has alcanzado la cantidad maxima de oro.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        ElseIf Precio + UserList(UserIndex).Stats.GLD >= ConstantesBalance.MaxOro Then
            Call WriteConsoleMsg(UserIndex, "No puedes vender tantos objetos, no podras llevar el oro.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Call QuitarUserInvItem(UserIndex, Slot, Cantidad)
        
        
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + Precio
        
        If UserList(UserIndex).Stats.GLD > ConstantesBalance.MaxOro Then _
            UserList(UserIndex).Stats.GLD = ConstantesBalance.MaxOro
        
        Dim NpcSlot As Integer
        NpcSlot = SlotEnNPCInv(NpcIndex, Objeto.ObjIndex, Objeto.Amount)
        
        If NpcSlot <= MAX_INVENTORY_SLOTS Then 'Slot valido
            'Mete el obj en el slot
            Npclist(NpcIndex).Invent.Object(NpcSlot).ObjIndex = Objeto.ObjIndex
            Npclist(NpcIndex).Invent.Object(NpcSlot).Amount = Npclist(NpcIndex).Invent.Object(NpcSlot).Amount + Objeto.Amount
            If Npclist(NpcIndex).Invent.Object(NpcSlot).Amount > MAX_INVENTORY_OBJS Then
                Npclist(NpcIndex).Invent.Object(NpcSlot).Amount = MAX_INVENTORY_OBJS
            End If
            
            Call UpdateNpcInv(NpcIndex, CByte(NpcSlot))
        End If
        
        'Bien, ahora logueo de ser necesario. Pablo (ToxicWaste) 07/09/07
        'Es un Objeto que tenemos que loguear?
        If ObjData(Objeto.ObjIndex).Log = 1 Then
            Call LogDesarrollo(UserList(UserIndex).Name & " vendió al NPC " & Objeto.Amount & " " & ObjData(Objeto.ObjIndex).Name)
        ElseIf Objeto.Amount >= 1000 Then 'Es mucha cantidad?
            'Si no es de los prohibidos de loguear, lo logueamos.
            If ObjData(Objeto.ObjIndex).NoLog <> 1 Then
                Call LogDesarrollo(UserList(UserIndex).Name & " vendió al NPC " & Objeto.Amount & " " & ObjData(Objeto.ObjIndex).Name)
            End If
        End If
    End If
    
    Call WriteUpdateUserStats(UserIndex)
    'Call WriteTradeOK(UserIndex)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Comercio de Comercio.bas")
End Sub

Public Sub IniciarComercioNPC(ByVal UserIndex As Integer)
'*************************************************
'Author: Nacho (Integer)
'Last modified: 2/8/06
'*************************************************
On Error GoTo ErrHandler
  
    Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC)
    UserList(UserIndex).flags.Comerciando = UserList(UserIndex).flags.TargetNPC
    Call EnviarUserInvPrices(UserIndex)
    Call WriteCommerceInit(UserIndex)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub IniciarComercioNPC de Comercio.bas")
End Sub

Private Function SlotEnNPCInv(ByVal NpcIndex As Integer, ByVal Objeto As Integer, ByVal Cantidad As Integer) As Integer
'*************************************************
'Author: Nacho (Integer)
'Last modified: 2/8/06
'*************************************************
On Error GoTo ErrHandler
  
    SlotEnNPCInv = 1
    Do Until Npclist(NpcIndex).Invent.Object(SlotEnNPCInv).ObjIndex = Objeto _
      And Npclist(NpcIndex).Invent.Object(SlotEnNPCInv).Amount + Cantidad <= MAX_INVENTORY_OBJS
        
        SlotEnNPCInv = SlotEnNPCInv + 1
        If SlotEnNPCInv > MAX_INVENTORY_SLOTS Then Exit Do
        
    Loop
    
    If SlotEnNPCInv > MAX_INVENTORY_SLOTS Then
    
        SlotEnNPCInv = 1
        
        Do Until Npclist(NpcIndex).Invent.Object(SlotEnNPCInv).ObjIndex = 0
            SlotEnNPCInv = SlotEnNPCInv + 1
            If SlotEnNPCInv > MAX_INVENTORY_SLOTS Then Exit Do
        Loop
        
        If SlotEnNPCInv <= MAX_INVENTORY_SLOTS Then Npclist(NpcIndex).Invent.NroItems = Npclist(NpcIndex).Invent.NroItems + 1
    End If
    
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function SlotEnNPCInv de Comercio.bas")
End Function

''
' Send the inventory of the Npc to the user
'
' @param userIndex The index of the User
' @param npcIndex The index of the NPC

Private Sub EnviarNpcInv(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
'*************************************************
'Author: Nacho (Integer)
'Last Modified: 10/04/2015
'Last Modified By: ZaMa
'03/05/2012: ZaMa - Solo envío los slots con items
'10/04/2015: D'Artagnan - Take off removed.
'*************************************************
On Error GoTo ErrHandler
  

    Dim Slot As Long
    Dim Price As Single
    Dim ObjIndex As Integer
    Dim CanUse As Boolean
    Dim Amount As Integer
    
    ' Reset
    Call WriteChangeNPCInventorySlot(UserIndex, 0, 0, 0, 0, True)
    
    For Slot = 1 To MAX_NORMAL_INVENTORY_SLOTS
        
        ObjIndex = Npclist(NpcIndex).Invent.Object(Slot).ObjIndex
        Amount = Npclist(NpcIndex).Invent.Object(Slot).Amount
        If ObjIndex > 0 And Amount > 0 Then
            Price = ObjData(ObjIndex).Valor
            CanUse = General.checkCanUseItem(UserIndex, ObjIndex)
            Call WriteChangeNPCInventorySlot(UserIndex, CByte(Slot), Amount, Price, ObjIndex, CanUse)
        End If
    Next Slot
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub EnviarNpcInv de Comercio.bas")
End Sub


Private Sub EnviarUserInvPrices(ByVal UserIndex As Integer)
On Error GoTo ErrHandler
    
    Dim Slot As Integer
    Dim ObjIndex As Integer
    Dim SalePrice As Long
    
    With UserList(UserIndex)
        For Slot = 1 To MAX_INVENTORY_SLOTS
            ObjIndex = .Invent.Object(Slot).ObjIndex
            If ObjIndex > 0 Then
                SalePrice = ObjData(ObjIndex).SalePrice
                Call WriteSetUserSalePrice(UserIndex, Slot, SalePrice)
            End If
        Next Slot
    End With
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub EnviarUserInvPrices de Comercio.bas")
End Sub

Private Sub UpdateNpcInv(ByVal NpcIndex As Integer, ByVal NpcSlot As Byte)
'*************************************************
'Author: Torres Patricio (Pato)
'Last Modified: 10/04/2015
'10/04/2015: D'Artagnan - Take off removed.
'*************************************************
On Error GoTo ErrHandler
  

    Dim Price As Single
    Dim ObjIndex As Integer
    Dim I As Long
    Dim Map As Integer
    Dim tempIndex As Integer
    Dim CanUse As Boolean
    
    Map = Npclist(NpcIndex).Pos.Map
    
    If Not MapaValido(Map) Then Exit Sub
    
    ObjIndex = Npclist(NpcIndex).Invent.Object(NpcSlot).ObjIndex
            
    Dim Query() As Collision.UUID
    Call ModAreas.QueryEntities(NpcIndex, ENTITY_TYPE_NPC, Query, ENTITY_TYPE_PLAYER)

    For I = 0 To UBound(Query)
        tempIndex = Query(I).Name
        
        If getTradingNPC(tempIndex) = NpcIndex Then
            If UserList(tempIndex).ConnIDValida Then
                
                If ObjIndex > 0 Then
                    Price = ObjData(ObjIndex).Valor
                    CanUse = General.checkCanUseItem(tempIndex, ObjIndex)
                End If
                
                Call WriteChangeNPCInventorySlot(tempIndex, NpcSlot, Npclist(NpcIndex).Invent.Object(NpcSlot).Amount, _
                                                                                Price, ObjIndex, CanUse)
                        
                ''Call WriteTradeOK(tempIndex)
            End If
        End If
    Next I
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub UpdateNpcInv de Comercio.bas")
End Sub
